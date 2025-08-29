#!/usr/bin/env python
# @File     : name_picker.py
# @Author   : 念安
# @Time     : 2025/8/29 10:53
# @Verison  : V1.0
# @Desctrion:

# name_picker.py  —— PySide6 + PyQt-Fluent-Widgets (1.x 兼容) 漂亮版
import sys, os, random, json
from datetime import datetime
import pandas as pd

from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, QTimer
from PySide6.QtWidgets import (
    QApplication, QFileDialog, QTableWidgetItem
)

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
from PySide6.QtWidgets import QTableWidgetItem, QAbstractItemView


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
            if a in df.columns:
                found = a
                break
            if a.lower() in lower_map:
                found = lower_map[a.lower()]
                break
        if not found:
            return None
        cols_map[std_col] = found
    return cols_map


class PandasModel(QAbstractTableModel):
    def __init__(self, df: pd.DataFrame):
        super().__init__()
        self._df = df

    def rowCount(self, parent=QModelIndex()):
        return len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None
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
        if role != Qt.DisplayRole:
            return None
        return self._df.columns[section] if orientation == Qt.Horizontal else section + 1

    def set_cell(self, row, col_name, value):
        if col_name in self._df.columns:
            # 统一文本列，避免 pandas dtype 警告
            if self._df[col_name].dtype.kind != "O":
                self._df[col_name] = self._df[col_name].astype(object)
            self._df.at[self._df.index[row], col_name] = value
            idx = self.index(row, self._df.columns.get_loc(col_name))
            self.dataChanged.emit(idx, idx, [Qt.DisplayRole])

    def df(self):
        return self._df


class MainWindow(FluentWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Roll Call · Fluent Edition (PyQt-Fluent-Widgets 1.x)")
        setTheme(Theme.LIGHT)
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
        self.roll_timer.setInterval(50)  # 滚动速度（ms）
        self.roll_timer.timeout.connect(self._roll_tick)

        self.auto_sign_timer = QTimer(self)
        self.auto_sign_timer.setSingleShot(True)
        self.auto_sign_timer.timeout.connect(self.sign_current_or_selected)

        self._load_state()
        self._build_ui()
        self._autoload_cache()

    # ========== 工具 ==========
    def _ensure_text(self, df):
        for c in ["学号", "姓名", "签到状态", "签到时间"]:
            if c in df.columns:
                df[c] = df[c].astype(object)
                df[c] = df[c].where(df[c].notna(), "")
        return df

    def _toast(self, title: str, content: str, level: str = "info"):
        kw = dict(
            title=title,
            content=content,
            orient=Qt.Horizontal,
            isClosable=True,
            position=InfoBarPosition.TOP_RIGHT,
            duration=1800,
            parent=self,
        )
        if level == "success":
            InfoBar.success(**kw)
        elif level == "warning":
            InfoBar.warning(**kw)
        elif level == "error":
            InfoBar.error(**kw)
        else:
            InfoBar.info(**kw)

    # ========== UI ==========
    def _build_ui(self):
        # 主/设置 两个页面
        self.addSubInterface(self._build_main_page(), FI.HOME, "Roll Call", NavigationItemPosition.TOP)
        self.addSubInterface(self._build_settings_page(), FI.SETTING, "Settings", NavigationItemPosition.BOTTOM)
        self.navigationInterface.setAcrylicEnabled(True)
        self.titleBar.raise_()

    def _build_main_page(self):
        page = CardWidget(self)
        page.setObjectName("mainPage")
        page.setMinimumSize(960, 640)

        # 顶部工具条
        self.btnImport = PrimaryPushButton(FI.FOLDER, "Import Excel", page)
        self.btnImport.clicked.connect(self.load_excel)

        self.btnToggle = PrimaryPushButton(FI.PLAY, "Start", page)
        self.btnToggle.clicked.connect(self.toggle_roll)

        self.btnSign = PrimaryPushButton(FI.CHECKBOX, "Sign", page)
        self.btnSign.clicked.connect(self.sign_current_or_selected)

        self.btnClearAll = PushButton(FI.DELETE, "Clear All", page)
        self.btnClearAll.clicked.connect(self.clear_all_sign)

        self.btnClearSel = PushButton(FI.REMOVE, "Clear Selected", page)
        self.btnClearSel.clicked.connect(self.clear_selected_sign)

        self.chkNoRepeat = CheckBox("No repeat", page)
        self.chkNoRepeat.setChecked(self.no_repeat)
        self.chkNoRepeat.stateChanged.connect(self._toggle_no_repeat)

        self.btnTheme = PushButton(FI.BRUSH, "Toggle Theme", page)
        self.btnTheme.clicked.connect(self._toggle_theme)

        self.searchBox = LineEdit(page)
        self.searchBox.setPlaceholderText("Search by ID/Name")
        self.searchBox.textChanged.connect(self._on_search)

        # 统计徽章 & 进度
        self.lblStats = StrongBodyLabel("Total: 0 | Present: 0 | Absent: 0", page)
        self.badgePresent = InfoBadge.success("0", parent=page, position=InfoBadgePosition.TOP_RIGHT)
        # self.badgePresent.setText("0")
        self.progress = ProgressBar(page)
        self.progress.setValue(0)

        # 大屏展示
        self.bigText = StrongBodyLabel("——", page)
        self.bigText.setAlignment(Qt.AlignCenter)
        self.bigText.setFixedHeight(120)
        self.bigText.setStyleSheet("""
            QLabel{
                font-size: 48px; font-weight: 800;
                border-radius: 18px; padding: 14px 20px;
                background: rgba(0,0,0,0.04);
            }
        """)

        # 速度 & 倒计时
        self.speedSlider = Slider(Qt.Horizontal, page)
        self.speedSlider.setRange(10, 200)  # 10~200ms
        self.speedSlider.setValue(50)
        self.speedSlider.valueChanged.connect(lambda v: self.roll_timer.setInterval(v))

        self.countdownSpin = SpinBox(page)
        self.countdownSpin.setRange(0, 10)
        self.countdownSpin.setValue(0)  # 0 不自动签到

        # 表格
        self.table = TableWidget(page)  # 只传 parent
        self.table.setRowCount(0)  # 初始 0 行
        self.table.setColumnCount(4)  # 4 列
        self.table.setHorizontalHeaderLabels(["学号", "姓名", "签到状态", "签到时间"])

        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)

        # 布局（直接定位，简单直观）
        self.btnImport.move(20, 20)
        self.btnToggle.move(150, 20)
        self.btnSign.move(260, 20)
        self.btnClearAll.move(360, 20)
        self.btnClearSel.move(480, 20)
        self.chkNoRepeat.move(610, 24)
        self.btnTheme.move(720, 20)
        self.searchBox.resize(200, 36); self.searchBox.move(820, 20)

        self.bigText.resize(920, 120); self.bigText.move(20, 70)
        self.lblStats.move(20, 200)
        self.progress.resize(400, 8); self.progress.move(20, 228)

        BodyLabel("Speed (ms)", page).move(460, 200)
        self.speedSlider.resize(180, 22); self.speedSlider.move(540, 200)

        BodyLabel("Auto sign after (s)", page).move(740, 200)
        self.countdownSpin.move(890, 196)

        self.table.resize(920, 380); self.table.move(20, 250)

        return page

    def _build_settings_page(self):
        page = CardWidget(self)
        page.setObjectName("settingsPage")
        BodyLabel("Settings", page).move(20, 20)
        BodyLabel("This page is reserved for future options.", page).move(20, 50)
        return page

    # ========== 数据/缓存 ==========
    def _toggle_theme(self):
        setTheme(Theme.DARK if self.theme() == Theme.LIGHT else Theme.LIGHT)

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
        if self.df.empty:
            return
        try:
            roster = self.df[["学号", "姓名"]].copy()
            roster["学号"] = roster["学号"].astype(str).str.strip()
            roster["姓名"] = roster["姓名"].astype(str).str.strip()
            roster.to_excel(CACHE_FILE, index=False)
        except Exception:
            pass

    def _autoload_cache(self):
        if not os.path.exists(CACHE_FILE):
            return
        try:
            df = pd.read_excel(CACHE_FILE, engine="openpyxl" if CACHE_FILE.endswith("xlsx") else None)
            if df.empty or not {"学号", "姓名"}.issubset(df.columns):
                return
            df = df.dropna(how="all").copy()
            df["学号"] = df["学号"].astype(str).str.strip()
            df["姓名"] = df["姓名"].astype(str).str.strip()
            # 每次启动都重置签到列为空
            df["签到状态"] = ""
            df["签到时间"] = ""
            df = self._ensure_text(df)
            self._use_df(df.reset_index(drop=True))
            self._rebuild_pool()
            self._toast("Loaded", "Using cached roster.", "success")
        except Exception:
            pass

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Choose Excel", "", "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        try:
            df = pd.read_excel(path, engine="openpyxl" if path.endswith("xlsx") else None)
        except Exception as e:
            self._toast("Read failed", str(e), "error")
            return
        if df.empty:
            self._toast("Empty file", "Excel has no content.", "warning")
            return

        cols = resolve_columns(df)
        if not cols:
            self._toast("Columns missing", "Need '学号' & '姓名' columns.", "error")
            return

        df = df.rename(columns={cols["学号"]: "学号", cols["姓名"]: "姓名"})
        df = df.dropna(how="all").copy()
        df["学号"] = df["学号"].astype(str).str.strip()
        df["姓名"] = df["姓名"].astype(str).str.strip()
        # 导入时也清空签到列
        df["签到状态"] = ""
        df["签到时间"] = ""
        df = self._ensure_text(df)

        self._use_df(df.reset_index(drop=True))
        self._rebuild_pool()
        self._save_cache()
        self._toast("Imported", f"Loaded {len(self.df)} students.", "success")

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
        col_index = {name: i for i, name in enumerate(["学号", "姓名", "签到状态", "签到时间"])}
        try:
            self.table.setColumnWidth(col_index["学号"], 120)
            self.table.setColumnWidth(col_index["姓名"], 140)
            self.table.setColumnWidth(col_index["签到状态"], 120)
            self.table.setColumnWidth(col_index["签到时间"], 180)
        except Exception:
            pass

        self._update_stats()

    # ========== 抽取/签到 ==========
    def _rebuild_pool(self):
        if self.df is None or self.df.empty:
            self.current_idx_pool = []
            return
        if self.no_repeat:
            self.current_idx_pool = self.df.index[self.df["签到状态"] != "已签到"].tolist()
        else:
            self.current_idx_pool = list(range(len(self.df)))
        random.shuffle(self.current_idx_pool)

    def toggle_roll(self):
        if self.df is None or self.df.empty:
            self._toast("Tip", "Import roster first.", "warning")
            return
        if not self.rolling:
            if self.no_repeat and (self.df["签到状态"] != "已签到").sum() == 0:
                self._toast("Done", "All students signed.", "success")
                return
            if not self.current_idx_pool:
                self._rebuild_pool()
            self.rolling = True
            self.roll_timer.start()
            self.btnToggle.setText("Pause")
            self.btnToggle.setIcon(FI.PAUSE.icon())
            # 倒计时自动签到
            sec = self.countdownSpin.value()
            if sec > 0:
                self.auto_sign_timer.start(sec * 1000)
        else:
            self.roll_timer.stop()
            self.auto_sign_timer.stop()
            self.rolling = False
            self.btnToggle.setText("Start")
            self.btnToggle.setIcon(FI.PLAY.icon())

    def _roll_tick(self):
        if not self.current_idx_pool:
            self._rebuild_pool()
            if not self.current_idx_pool:
                self.toggle_roll()
                return
        idx = random.choice(self.current_idx_pool)
        name = self.df.at[idx, "姓名"]
        sid = self.df.at[idx, "学号"]
        self.last_show_text = f"{sid}  {name}"
        self.bigText.setText(self.last_show_text)

    def _find_row_by_sid_or_name(self):
        # 签到前确保暂停
        if self.rolling:
            self.toggle_roll()

        sid = None
        name_from_big = None
        text = (self.last_show_text or "").strip()
        if text:
            parts = text.split()
            if parts:
                sid = parts[0].strip()
                if len(parts) > 1:
                    name_from_big = " ".join(parts[1:]).strip()

        row = None
        if sid:
            series_ids = self.df["学号"].astype(str).str.strip()
            m = series_ids[series_ids == str(sid)].index.tolist()
            if m:
                row = m[0]
        if row is None and name_from_big:
            series_names = self.df["姓名"].astype(str).str.strip()
            m = series_names[series_names == name_from_big].index.tolist()
            if len(m) == 1:
                row = m[0]
        if row is None:
            items = self.table.selectedIndexes()
            if items:
                row = items[0].row()
        return row

    def sign_current_or_selected(self):
        if self.df is None or self.df.empty:
            self._toast("Tip", "Import roster first.", "warning")
            return

        row = self._find_row_by_sid_or_name()
        if row is None:
            self._toast("Tip", "No target to sign. Start rolling or select a row.", "warning")
            return

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.model.set_cell(row, "签到状态", "已签到")
        self.model.set_cell(row, "签到时间", now)

        # 同步到表格
        self.table.item(row, 2).setText("已签到")
        self.table.item(row, 3).setText(now)

        # 更新统计、移出池
        self._update_stats()
        if self.no_repeat and row in self.current_idx_pool:
            try:
                self.current_idx_pool.remove(row)
            except ValueError:
                pass

        self._save_cache()
        self._toast("Signed", f"{self.df.at[row,'学号']} {self.df.at[row,'姓名']} ✓", "success")

    # ========== 清除 ==========
    def clear_all_sign(self):
        if self.df is None or self.df.empty:
            return
        m = MessageBox("Confirm", "Clear ALL attendance?", self)
        if m.exec():
            self.df["签到状态"] = ""
            self.df["签到时间"] = ""
            for r in range(len(self.df)):
                self.table.item(r, 2).setText("")
                self.table.item(r, 3).setText("")
            self._update_stats()
            self._rebuild_pool()
            self._save_cache()
            self._toast("Cleared", "All cleared.", "success")

    def clear_selected_sign(self):
        if self.df is None or self.df.empty:
            return
        rows = sorted(set(i.row() for i in self.table.selectedIndexes()))
        if not rows:
            self._toast("Tip", "Select at least one row.", "warning")
            return
        for r in rows:
            self.model.set_cell(r, "签到状态", "")
            self.model.set_cell(r, "签到时间", "")
            self.table.item(r, 2).setText("")
            self.table.item(r, 3).setText("")
        self._update_stats()
        self._rebuild_pool()
        self._save_cache()
        self._toast("Cleared", f"Cleared {len(rows)} row(s).", "success")

    # ========== 其它 ==========
    def _on_search(self, kw: str):
        kw = kw.strip()
        # 简单过滤：按学号/姓名隐藏行
        for r in range(self.table.rowCount()):
            sid = self.table.item(r, 0).text()
            name = self.table.item(r, 1).text()
            show = (kw in sid) or (kw in name) or (kw == "")
            self.table.setRowHidden(r, not show)

    def _update_stats(self):
        if self.df is None or self.df.empty:
            self.lblStats.setText("Total: 0 | Present: 0 | Absent: 0")
            self.progress.setValue(0)
            self.badgePresent.setText("0")
            return
        total = len(self.df)
        present = (self.df["签到状态"] == "已签到").sum()
        absent = total - present
        self.lblStats.setText(f"Total: {total} | Present: {present} | Absent: {absent}")
        self.progress.setValue(int(present * 100 / total) if total else 0)
        self.badgePresent.setText(str(present))
        self.badgePresent.move(900, 8)  # 固定位置

def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.resize(1000, 700)
    w.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
