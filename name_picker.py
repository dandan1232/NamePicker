#!/usr/bin/env python
# @File     : name_picker.py
# @Author   : 念安
# @Time     : 2025/8/29 10:53
# @Verison  : V1.0
# @Desctrion:

import sys, os, random, json
from datetime import datetime
import pandas as pd
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, QTimer, QEvent


from qfluentwidgets import (
    FluentWindow, setTheme, Theme, setFont,
    InfoBar, InfoBarPosition,
    NavigationItemPosition,
    FluentIcon as FI,
    PrimaryPushButton, PushButton,
    LineEdit, TableWidget,
    BodyLabel, StrongBodyLabel,
    Slider, SpinBox, CheckBox,
    ProgressBar, MessageBox, CardWidget,
    InfoBadge, InfoBadgePosition
)
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QTableWidgetItem, QAbstractItemView,
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QSizePolicy
)


CACHE_FILE = "roster_cache.xlsx"
STATE_FILE = "app_state.json"

COLUMN_ALIASES = {
    "学号": {"学号", "学员编号", "学生编号", "学籍号", "student_id", "id"},
    "姓名": {"姓名", "学生姓名", "name", "student_name"}
}

def resolve_columns(df: pd.DataFrame):
    lower_map = {c.lower(): c for c in df.columns}
    cols_map = {}
    for std_col, aliases in COLUMN_ALIASES.items():
        found = None
        for a in aliases:
            if a in df.columns: found = a; break
            if a.lower() in lower_map: found = lower_map[a.lower()]; break
        if not found: return None
        cols_map[std_col] = found
    return cols_map


class PandasModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self._df = df

    def rowCount(self, parent=QModelIndex()): return len(self._df)
    def columnCount(self, parent=QModelIndex()): return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid(): return None
        col_name = self._df.columns[index.column()]

        if role in (Qt.DisplayRole, Qt.EditRole):
            v = self._df.iat[index.row(), index.column()]
            return "" if pd.isna(v) else str(v)

        if role == Qt.TextAlignmentRole:
            if col_name in ("学号", "姓名"):
                return Qt.AlignCenter
            return Qt.AlignVCenter

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole: return None
        return self._df.columns[section] if orientation == Qt.Horizontal else section + 1

    def set_cell(self, row, col_name, value):
        if col_name in self._df.columns:
            # 写入前确保为文本列，避免 pandas FutureWarning
            if self._df[col_name].dtype.kind != "O":
                self._df[col_name] = self._df[col_name].astype(object)
            self._df.at[self._df.index[row], col_name] = value
            idx = self.index(row, self._df.columns.get_loc(col_name))
            self.dataChanged.emit(idx, idx, [Qt.DisplayRole])

    def df(self): return self._df


class MainWindow(FluentWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("课堂点名 · 章老师版")
        setTheme(Theme.LIGHT)   # 默认浅色
        setFont(self, 12)

        # 状态
        self.df = pd.DataFrame()
        self.model = None
        self.rolling = False
        self.last_show_text = ""
        self.no_repeat = True
        self.current_idx_pool = []

        # 定时器
        self.roll_timer = QTimer(self)
        self.roll_timer.setInterval(50)     # 滚动速度（毫秒）
        self.roll_timer.timeout.connect(self._roll_tick)

        self.auto_sign_timer = QTimer(self)
        self.auto_sign_timer.setSingleShot(True)
        self.auto_sign_timer.timeout.connect(self.sign_current_or_selected)

        self._load_state()
        self._build_ui()
        self._autoload_cache()

    # --------- 工具函数 ----------
    def _ensure_text(self, df):
        """把关键列都当作文本列，并把 NaN 变为空串"""
        for c in ["学号", "姓名", "签到状态", "签到时间"]:
            if c in df.columns:
                df[c] = df[c].astype(object)
                df[c] = df[c].where(df[c].notna(), "")
        return df

    def _toast(self, title: str, content: str, level: str = "info"):
        kw = dict(
            title=title, content=content, orient=Qt.Horizontal, isClosable=True,
            position=InfoBarPosition.TOP_RIGHT, duration=1800, parent=self
        )
        if level == "success": InfoBar.success(**kw)
        elif level == "warning": InfoBar.warning(**kw)
        elif level == "error": InfoBar.error(**kw)
        else: InfoBar.info(**kw)

    # --------- UI ----------
    def _build_ui(self):
        self.page_main = self._build_main_page()
        self.addSubInterface(self.page_main, FI.HOME, "点名", NavigationItemPosition.TOP)

        # —— 彩蛋页（占位，不显示内容）——
        self.page_egg = QWidget(self)
        self.page_egg.setObjectName("eggPage")  # 必须要有对象名
        self.addSubInterface(self.page_egg, FI.HEART, "彩蛋", NavigationItemPosition.BOTTOM)

        # 用事件过滤器来“拦截显示”，弹出彩蛋后再切回主页
        self.page_egg.installEventFilter(self)

        self.navigationInterface.setAcrylicEnabled(True)
        self.titleBar.raise_()

    def eventFilter(self, obj, event):
        # 当左下角“彩蛋”页被切换为可见时，立刻弹出彩蛋并回到主页
        if obj is getattr(self, "page_egg", None) and event.type() == QEvent.Show:
            self._show_easter_egg()
            try:
                self.stackedWidget.setCurrentWidget(self.page_main)
                self.navigationInterface.setCurrentItem("mainPage")
            except Exception:
                pass
            return True
        return super().eventFilter(obj, event)

    def _on_egg_clicked(self):
        self._show_easter_egg()  # 弹出彩蛋
        # 回到主页面（防止停在空白占位页）
        try:
            self.stackedWidget.setCurrentWidget(self.page_main)
            self.navigationInterface.setCurrentItem("mainPage")
        except Exception:
            pass

    def _show_easter_egg(self):
        MessageBox("🎁 彩蛋", "恭喜星星同学发现了隐藏的彩蛋！\n奖励你师德+1！", self).exec()

    def _build_main_page(self):
        page = QWidget(self)
        page.setObjectName("mainPage")
        page.setMinimumSize(1080, 640)

        # ===== 顶部工具条 =====
        self.btnImport = PrimaryPushButton(FI.FOLDER, "导入Excel", page)
        self.btnToggle = PrimaryPushButton(FI.PLAY, "开始", page)  # 开始/暂停在 toggle_roll 中切换
        self.btnSign = PrimaryPushButton(FI.CHECKBOX, "签到", page)
        self.btnClearAll = PushButton(FI.DELETE, "清空所有签到", page)
        self.btnClearSel = PushButton(FI.REMOVE, "清除选中行签到", page)
        self.chkNoRepeat = CheckBox("不重复抽取（默认）", page)
        self.btnTheme = PushButton(FI.BRUSH, "切换主题", page)
        self.searchBox = LineEdit(page);
        self.searchBox.setPlaceholderText("按学号/姓名搜索")

        self.btnImport.clicked.connect(self.load_excel)
        self.btnToggle.clicked.connect(self.toggle_roll)
        self.btnSign.clicked.connect(self.sign_current_or_selected)
        self.btnClearAll.clicked.connect(self.clear_all_sign)
        self.btnClearSel.clicked.connect(self.clear_selected_sign)
        self.chkNoRepeat.setChecked(self.no_repeat)
        self.chkNoRepeat.stateChanged.connect(self._toggle_no_repeat)
        self.btnTheme.clicked.connect(self._toggle_theme)
        self.searchBox.textChanged.connect(self._on_search)

        topBar = QHBoxLayout()
        topBar.setContentsMargins(0, 0, 0, 0)
        topBar.setSpacing(8)
        for w in [self.btnImport, self.btnToggle, self.btnSign, self.btnClearAll, self.btnClearSel,
                  self.chkNoRepeat, self.btnTheme, self.searchBox]:
            topBar.addWidget(w)
        topBar.addStretch(1)

        # ===== 大屏显示 =====
        self.bigText = StrongBodyLabel("——", page)
        self.bigText.setAlignment(Qt.AlignCenter)
        bigFrame = QFrame(page)
        bigFrame.setFrameShape(QFrame.StyledPanel)
        bigFrame.setStyleSheet("QFrame{background:rgba(0,0,0,0.05); border-radius:18px;}")
        bigLay = QVBoxLayout(bigFrame);
        bigLay.setContentsMargins(16, 16, 16, 16)
        self.bigText.setStyleSheet("QLabel{font-size:48px; font-weight:800;}")
        self.bigText.setMinimumHeight(120)
        bigLay.addWidget(self.bigText)

        # ===== 统计 & 控件 =====
        self.lblStats = StrongBodyLabel("总数：0 | 已签到：0 | 未签到：0", page)
        self.progress = ProgressBar(page);
        self.progress.setValue(0)

        statsLay = QVBoxLayout()
        statsLay.setSpacing(6)
        statsLay.addWidget(self.lblStats)
        statsLay.addWidget(self.progress)

        self.speedSlider = Slider(Qt.Horizontal, page)
        self.speedSlider.setRange(10, 200);
        self.speedSlider.setValue(50)
        self.speedSlider.valueChanged.connect(lambda v: self.roll_timer.setInterval(v))
        self.countdownSpin = SpinBox(page);
        self.countdownSpin.setRange(0, 10);
        self.countdownSpin.setValue(0)

        ctrlLay = QHBoxLayout()
        ctrlLay.setSpacing(12)
        ctrlLay.addLayout(statsLay, stretch=1)
        ctrlLay.addWidget(BodyLabel("滚动速度（毫秒）", page))
        ctrlLay.addWidget(self.speedSlider)
        ctrlLay.addWidget(BodyLabel("自动签到延迟（秒）", page))
        ctrlLay.addWidget(self.countdownSpin)

        # ===== 表格 =====
        self.table = TableWidget(page)  # 1.x 只接受 parent
        self.table.setRowCount(0)
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["学号", "姓名", "签到状态", "签到时间"])
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # ===== 页面总体布局 =====
        root = QVBoxLayout(page)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)
        root.addLayout(topBar)
        root.addWidget(bigFrame)
        root.addLayout(ctrlLay)
        root.addWidget(self.table, stretch=1)

        return page


    # --------- 数据/缓存 ----------
    def _toggle_theme(self):
        # 用实例变量记录当前主题，避免调用不存在的 self.theme()
        if not hasattr(self, "_is_dark"):
            self._is_dark = False
        self._is_dark = not self._is_dark
        setTheme(Theme.DARK if self._is_dark else Theme.LIGHT)

    def _toggle_no_repeat(self, _):
        self.no_repeat = self.chkNoRepeat.isChecked()
        self._save_state()
        self._rebuild_pool()

    def _save_state(self):
        try:
            json.dump({"no_repeat": self.no_repeat}, open(STATE_FILE, "w", encoding="utf-8"))
        except Exception:
            pass

    def _load_state(self):
        if os.path.exists(STATE_FILE):
            try:
                self.no_repeat = bool(json.load(open(STATE_FILE, "r", encoding="utf-8")).get("no_repeat", True))
            except Exception:
                self.no_repeat = True

    def _save_cache(self):
        """只缓存学号/姓名，保证下次打开签到列是空的"""
        if self.df.empty: return
        try:
            roster = self.df[["学号", "姓名"]].copy()
            roster["学号"] = roster["学号"].astype(str).str.strip()
            roster["姓名"] = roster["姓名"].astype(str).str.strip()
            roster.to_excel(CACHE_FILE, index=False)
        except Exception:
            pass

    def _autoload_cache(self):
        """命中缓存：载入学号/姓名，签到列重置为空"""
        if not os.path.exists(CACHE_FILE): return
        try:
            df = pd.read_excel(CACHE_FILE, engine="openpyxl" if CACHE_FILE.endswith("xlsx") else None)
            if df.empty or not {"学号", "姓名"}.issubset(df.columns): return
            df = df.dropna(how="all").copy()
            df["学号"] = df["学号"].astype(str).str.strip()
            df["姓名"] = df["姓名"].astype(str).str.strip()
            df["签到状态"] = ""
            df["签到时间"] = ""
            df = self._ensure_text(df)
            self._use_df(df.reset_index(drop=True))
            self._rebuild_pool()
            self._toast("已加载", "已加载上次的花名册缓存。", "success")
        except Exception:
            pass

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if not path: return
        try:
            df = pd.read_excel(path, engine="openpyxl" if path.endswith("xlsx") else None)
        except Exception as e:
            self._toast("读取失败", str(e), "error"); return
        if df.empty:
            self._toast("空文件", "Excel 内容为空。", "warning"); return

        cols = resolve_columns(df)
        if not cols:
            self._toast("缺少列", "需要包含“学号”和“姓名”列或其同义列。", "error"); return

        df = df.rename(columns={cols["学号"]: "学号", cols["姓名"]: "姓名"})
        df = df.dropna(how="all").copy()
        df["学号"] = df["学号"].astype(str).str.strip()
        df["姓名"] = df["姓名"].astype(str).str.strip()
        df["签到状态"] = ""
        df["签到时间"] = ""
        df = self._ensure_text(df)

        self._use_df(df.reset_index(drop=True))
        self._rebuild_pool()
        self._save_cache()
        self._toast("导入成功", f"已载入 {len(self.df)} 名学生。", "success")

    def _use_df(self, df: pd.DataFrame):
        self.df = df
        self.model = PandasModel(self.df)

        # 重建 TableWidget 内容
        self.table.clearContents()
        self.table.setRowCount(len(df))
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["学号", "姓名", "签到状态", "签到时间"])

        for r in range(len(df)):
            for c, col_name in enumerate(["学号", "姓名", "签到状态", "签到时间"]):
                v = "" if pd.isna(df.at[r, col_name]) else str(df.at[r, col_name])
                item = QTableWidgetItem(v)
                if col_name in ("学号", "姓名"):
                    item.setTextAlignment(Qt.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignVCenter)
                self.table.setItem(r, c, item)

        # 列宽设置
        idx = {"学号":0, "姓名":1, "签到状态":2, "签到时间":3}
        try:
            self.table.setColumnWidth(idx["学号"], 120)
            self.table.setColumnWidth(idx["姓名"], 180)
            self.table.setColumnWidth(idx["签到状态"], 120)
            self.table.setColumnWidth(idx["签到时间"], 180)
        except Exception:
            pass

        self._update_stats()

    # --------- 抽取/签到 ----------
    def _rebuild_pool(self):
        if self.df is None or self.df.empty:
            self.current_idx_pool = []; return
        if self.no_repeat:
            self.current_idx_pool = self.df.index[self.df["签到状态"] != "已签到"].tolist()
        else:
            self.current_idx_pool = list(range(len(self.df)))
        random.shuffle(self.current_idx_pool)

    def toggle_roll(self):
        if self.df is None or self.df.empty:
            self._toast("提示", "请先导入花名册。", "warning"); return
        if not self.rolling:
            if self.no_repeat and (self.df["签到状态"] != "已签到").sum() == 0:
                self._toast("完成", "全部学生已签到。", "success"); return
            if not self.current_idx_pool: self._rebuild_pool()
            self.rolling = True
            self.roll_timer.start()
            self.btnToggle.setText("暂停")
            self.btnToggle.setIcon(FI.PAUSE.icon())
            sec = self.countdownSpin.value()
            if sec > 0: self.auto_sign_timer.start(sec * 1000)
        else:
            self.roll_timer.stop()
            self.auto_sign_timer.stop()
            self.rolling = False
            self.btnToggle.setText("开始")
            self.btnToggle.setIcon(FI.PLAY.icon())

    def _roll_tick(self):
        if not self.current_idx_pool:
            self._rebuild_pool()
            if not self.current_idx_pool:
                self.toggle_roll(); return
        idx = random.choice(self.current_idx_pool)
        name = self.df.at[idx, "姓名"]
        sid = self.df.at[idx, "学号"]
        self.last_show_text = f"{sid}  {name}"
        self.bigText.setText(self.last_show_text)

    def _find_row_by_sid_or_name(self):
        # 签到前确保暂停
        if self.rolling: self.toggle_roll()

        sid = None; name_from_big = None
        text = (self.last_show_text or "").strip()
        if text:
            parts = text.split()
            if parts:
                sid = parts[0].strip()
                if len(parts) > 1: name_from_big = " ".join(parts[1:]).strip()

        row = None
        if sid:
            series_ids = self.df["学号"].astype(str).str.strip()
            m = series_ids[series_ids == str(sid)].index.tolist()
            if m: row = m[0]
        if row is None and name_from_big:
            series_names = self.df["姓名"].astype(str).str.strip()
            m = series_names[series_names == name_from_big].index.tolist()
            if len(m) == 1: row = m[0]
        if row is None:
            items = self.table.selectedIndexes()
            if items: row = items[0].row()
        return row

    def sign_current_or_selected(self):
        if self.df is None or self.df.empty:
            self._toast("提示", "请先导入花名册。", "warning"); return

        row = self._find_row_by_sid_or_name()
        if row is None:
            self._toast("提示", "没有可签到的对象：请先开始滚动或在表格中选中一行。", "warning"); return

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.model.set_cell(row, "签到状态", "已签到")
        self.model.set_cell(row, "签到时间", now)

        # 同步表格
        self.table.item(row, 2).setText("已签到")
        self.table.item(row, 3).setText(now)

        self._update_stats()
        if self.no_repeat and row in self.current_idx_pool:
            try: self.current_idx_pool.remove(row)
            except ValueError: pass

        self._save_cache()
        self._toast("已签到", f"{self.df.at[row,'学号']} {self.df.at[row,'姓名']} ✓", "success")

    # --------- 清除 ----------
    def clear_all_sign(self):
        if self.df is None or self.df.empty: return
        m = MessageBox("确认", "确定要清空所有签到状态吗？", self)
        if m.exec():
            self.df["签到状态"] = ""
            self.df["签到时间"] = ""
            for r in range(len(self.df)):
                self.table.item(r, 2).setText("")
                self.table.item(r, 3).setText("")
            self._update_stats(); self._rebuild_pool(); self._save_cache()
            self._toast("已清空", "已清空所有签到状态。", "success")

    def clear_selected_sign(self):
        if self.df is None or self.df.empty: return
        rows = sorted(set(i.row() for i in self.table.selectedIndexes()))
        if not rows:
            self._toast("提示", "请在表格中选中至少一行。", "warning"); return
        for r in rows:
            self.model.set_cell(r, "签到状态", "")
            self.model.set_cell(r, "签到时间", "")
            self.table.item(r, 2).setText("")
            self.table.item(r, 3).setText("")
        self._update_stats(); self._rebuild_pool(); self._save_cache()
        self._toast("已清除", f"已清除 {len(rows)} 行的签到。", "success")

    # --------- 其它 ----------
    def _on_search(self, kw: str):
        kw = kw.strip()
        for r in range(self.table.rowCount()):
            sid = self.table.item(r, 0).text()
            name = self.table.item(r, 1).text()
            show = (kw in sid) or (kw in name) or (kw == "")
            self.table.setRowHidden(r, not show)

    def _update_stats(self):
        if self.df is None or self.df.empty:
            self.lblStats.setText("总数：0 | 已签到：0 | 未签到：0")
            self.progress.setValue(0)
            return
        total = len(self.df)
        present = (self.df["签到状态"] == "已签到").sum()
        absent = total - present
        self.lblStats.setText(f"总数：{total} | 已签到：{present} | 未签到：{absent}")
        self.progress.setValue(int(present * 100 / total) if total else 0)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.resize(1040, 700)   # 默认大小
    w.show()

    # ===== 让窗口居中 =====
    rect = w.frameGeometry()
    center = app.primaryScreen().availableGeometry().center()
    rect.moveCenter(center)
    w.move(rect.topLeft())
    # ====================

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
