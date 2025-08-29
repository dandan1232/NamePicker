#!/usr/bin/env python
# @File     : name_picker.py
# @Author   : 念安
# @Time     : 2025/8/29 10:53
# @Verison  : V1.0
# @Desctrion:

#!/usr/bin/env python
# 简洁稳定版：支持缓存、滚动抽取、签到、清空、列宽/居中、类型统一
import sys, os, random, json
from datetime import datetime
import pandas as pd
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, QTimer
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox,
    QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QTableView, QCheckBox, QAbstractItemView
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

    def rowCount(self, parent=QModelIndex()):
        return len(self._df)

    def columnCount(self, parent=QModelIndex()):
        return len(self._df.columns)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid(): return None

        if role in (Qt.DisplayRole, Qt.EditRole):
            val = self._df.iat[index.row(), index.column()]
            return "" if pd.isna(val) else str(val)

        # 学号/姓名 居中
        if role == Qt.TextAlignmentRole:
            col_name = self._df.columns[index.column()]
            if col_name in ("学号", "姓名"):
                return Qt.AlignCenter

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole: return None
        return self._df.columns[section] if orientation == Qt.Horizontal else section + 1

    def set(self, row, col_name, value):
        if col_name in self._df.columns:
            # 关键：先把该列转为 object（文本），避免 pandas dtype 警告
            if self._df[col_name].dtype.kind != "O":
                self._df[col_name] = self._df[col_name].astype(object)
            self._df.at[self._df.index[row], col_name] = value
            idx = self.index(row, self._df.columns.get_loc(col_name))
            self.dataChanged.emit(idx, idx, [Qt.DisplayRole])

    def df(self):
        return self._df

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("课堂抽签式签到（导入一次·可复用）")
        self.resize(1000, 660)

        self.df = pd.DataFrame()
        self.model = None
        self.roll_timer = QTimer(self)
        self.roll_timer.setInterval(50)  # 滚动速度
        self.roll_timer.timeout.connect(self._roll_tick)

        self.rolling = False
        self.current_idx_pool = []
        self.last_show_text = ""
        self.no_repeat = True

        self._load_state()
        self._build_ui()
        self._autoload_cache()

    # —— 工具：把关键列都当“文本”处理，消除 NaN/类型问题 ——
    def _ensure_text_dtypes(self, df):
        for c in ["学号", "姓名", "签到状态", "签到时间"]:
            if c in df.columns:
                df[c] = df[c].astype(object)
                df[c] = df[c].where(df[c].notna(), "")
        return df

    def _build_ui(self):
        central = QWidget(); self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # 顶部按钮区
        top = QHBoxLayout()
        self.btn_import = QPushButton("导入Excel")
        self.btn_import.clicked.connect(self.load_excel)

        # 开始/暂停 切换按钮
        self.btn_toggle = QPushButton("开始")
        self.btn_toggle.clicked.connect(self.toggle_roll)

        # 独立“签到”按钮：对大屏显示者签到；若无则对选中行签到
        self.btn_sign = QPushButton("签到")
        self.btn_sign.clicked.connect(self.sign_current_or_selected)

        # 清空所有 & 清除选中行
        self.btn_clear_all = QPushButton("清空所有签到")
        self.btn_clear_all.clicked.connect(self.clear_all_sign)
        self.btn_clear_selected = QPushButton("清除选中行签到")
        self.btn_clear_selected.clicked.connect(self.clear_selected_sign)

        self.chk_no_repeat = QCheckBox("不重复抽取（默认）")
        self.chk_no_repeat.setChecked(self.no_repeat)
        self.chk_no_repeat.stateChanged.connect(self._toggle_no_repeat)

        self.lbl_stats = QLabel("未载入名单")
        self.lbl_stats.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        top.addWidget(self.btn_import)
        top.addSpacing(10)
        top.addWidget(self.btn_toggle)
        top.addWidget(self.btn_sign)
        top.addSpacing(10)
        top.addWidget(self.btn_clear_all)
        top.addWidget(self.btn_clear_selected)
        top.addSpacing(10)
        top.addWidget(self.chk_no_repeat)
        top.addStretch(1)
        top.addWidget(self.lbl_stats)
        layout.addLayout(top)

        # 大号滚动显示区
        self.lbl_big = QLabel("——")
        self.lbl_big.setAlignment(Qt.AlignCenter)
        self.lbl_big.setStyleSheet("""
            QLabel{
                font-size: 48px;
                font-weight: 700;
                padding: 16px 24px;
                border-radius: 16px;
                border: 2px solid #ddd;
                background: qlineargradient(x1:0,y1:0,x2:1,y2:1,
                    stop:0 #fafafa, stop:1 #f0f0f0);
            }
        """)
        layout.addWidget(self.lbl_big)

        # 表格
        self.table = QTableView()
        # 行选择更顺手
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setAlternatingRowColors(True)
        layout.addWidget(self.table, stretch=1)

        # 底部导出
        bottom = QHBoxLayout()
        self.btn_export = QPushButton("导出签到结果")
        self.btn_export.clicked.connect(self.export_excel)
        bottom.addStretch(1)
        bottom.addWidget(self.btn_export)
        layout.addLayout(bottom)

    # ——— 状态与缓存 ———
    def _toggle_no_repeat(self, _):
        self.no_repeat = self.chk_no_repeat.isChecked()
        self._save_state()
        self._rebuild_pool()

    def _save_cache(self):
        if self.df.empty: return
        try:
            # 只缓存花名册两列，不缓存签到状态/时间
            cols = [c for c in ["学号", "姓名"] if c in self.df.columns]
            if not cols:
                return
            roster = self.df[cols].copy()
            roster["学号"] = roster["学号"].astype(str).str.strip()
            roster["姓名"] = roster["姓名"].astype(str).str.strip()
            roster.to_excel(CACHE_FILE, index=False)
        except Exception:
            pass

    def _autoload_cache(self):
        if os.path.exists(CACHE_FILE):
            try:
                df = pd.read_excel(CACHE_FILE, engine="openpyxl" if CACHE_FILE.endswith("xlsx") else None)
                if not df.empty and {"学号", "姓名"}.issubset(df.columns):
                    df = df.dropna(how="all").copy()
                    df["学号"] = df["学号"].astype(str).str.strip()
                    df["姓名"] = df["姓名"].astype(str).str.strip()
                    # 关键：每次打开都初始化签到列为空
                    df["签到状态"] = ""
                    df["签到时间"] = ""
                    df = self._ensure_text_dtypes(df)
                    self._use_df(df.reset_index(drop=True))
                    self._rebuild_pool()
                    QMessageBox.information(self, "已载入缓存", f"使用上次导入的名单（{len(self.df)} 人）。")
            except Exception:
                pass

    def _save_state(self):
        try:
            json.dump({"no_repeat": self.no_repeat}, open(STATE_FILE, "w", encoding="utf-8"))
        except Exception:
            pass

    def _load_state(self):
        if os.path.exists(STATE_FILE):
            try:
                d = json.load(open(STATE_FILE, "r", encoding="utf-8"))
                self.no_repeat = bool(d.get("no_repeat", True))
            except Exception:
                self.no_repeat = True

    # ——— 数据导入 ———
    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择学生名单 Excel", "", "Excel Files (*.xlsx *.xls)")
        if not path: return
        try:
            df = pd.read_excel(path, engine="openpyxl" if path.endswith("xlsx") else None)
        except Exception as e:
            QMessageBox.critical(self, "读取失败", f"无法读取文件：\n{e}"); return
        if df.empty:
            QMessageBox.warning(self, "提示", "Excel 内容为空。"); return

        cols = resolve_columns(df)
        if not cols:
            QMessageBox.critical(self, "列未识别", "需要包含“学号/姓名”列或其同义列。"); return

        df = df.rename(columns={cols["学号"]: "学号", cols["姓名"]: "姓名"})
        df = df.dropna(how="all").copy()
        df["学号"] = df["学号"].astype(str).str.strip()
        df["姓名"] = df["姓名"].astype(str).str.strip()
        df["签到状态"] = ""
        df["签到时间"] = ""
        df = self._ensure_text_dtypes(df)

        self._use_df(df.reset_index(drop=True))
        self._rebuild_pool()
        self._save_cache()
        QMessageBox.information(self, "成功", f"已导入 {len(self.df)} 条学生记录。")

    def _use_df(self, df: pd.DataFrame):
        self.df = df
        self.model = PandasModel(self.df)
        self.table.setModel(self.model)
        self._tune_table_ui()
        self._update_stats()

    def _tune_table_ui(self):
        """列宽默认设置：状态/时间更宽；学号/姓名保持适中"""
        self.table.resizeColumnsToContents()
        cols = list(self.df.columns)
        try:
            if "签到状态" in cols:
                self.table.setColumnWidth(cols.index("签到状态"), 120)
            if "签到时间" in cols:
                self.table.setColumnWidth(cols.index("签到时间"), 180)
            if "学号" in cols:
                self.table.setColumnWidth(cols.index("学号"), max(self.table.columnWidth(cols.index("学号")), 120))
            if "姓名" in cols:
                self.table.setColumnWidth(cols.index("姓名"), max(self.table.columnWidth(cols.index("姓名")), 120))
        except Exception:
            pass

    def _update_stats(self):
        if self.df is None or self.df.empty:
            self.lbl_stats.setText("未载入名单"); return
        total = len(self.df)
        signed = (self.df["签到状态"] == "已签到").sum()
        self.lbl_stats.setText(f"总数：{total} | 已签到：{signed} | 未签到：{total - signed}")

    # ——— 抽取池与滚动 ———
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
            QMessageBox.warning(self, "提示", "请先导入或使用缓存名单。"); return
        if not self.rolling:
            if self.no_repeat and (self.df["签到状态"] != "已签到").sum() == 0:
                QMessageBox.information(self, "完成", "所有学生都已签到。"); return
            if not self.current_idx_pool:
                self._rebuild_pool()
            self.rolling = True
            self.roll_timer.start()
            self.btn_toggle.setText("暂停")
        else:
            self.roll_timer.stop()
            self.rolling = False
            self.btn_toggle.setText("开始")

    def _roll_tick(self):
        if not self.current_idx_pool:
            self._rebuild_pool()
            if not self.current_idx_pool:
                self.roll_timer.stop(); self.rolling = False; self.btn_toggle.setText("开始"); return
        idx = random.choice(self.current_idx_pool)
        name = self.df.at[idx, "姓名"]
        sid = self.df.at[idx, "学号"]
        self.last_show_text = f"{sid}  {name}"
        self.lbl_big.setText(self.last_show_text)

    # ——— 签到逻辑 ———
    def sign_current_or_selected(self):
        """优先对当前大屏显示的人签到；若无则对表格选中行签到"""
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "提示", "请先导入名单。"); return

        # 若在滚动，先暂停
        if self.rolling:
            self.toggle_roll()

        # 从大屏文本解析学号和姓名
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
            matches = series_ids[series_ids == str(sid)].index.tolist()
            if matches:
                row = matches[0]

        # 兜底：按姓名匹配（如果唯一）
        if row is None and name_from_big:
            series_names = self.df["姓名"].astype(str).str.strip()
            matches = series_names[series_names == name_from_big].index.tolist()
            if len(matches) == 1:
                row = matches[0]

        # 再兜底：用表格选中行
        if row is None and self.table.selectionModel():
            sel = self.table.selectionModel().selectedRows()
            if sel:
                row = sel[0].row()

        if row is None:
            QMessageBox.information(self, "提示", "没有可签到的对象：请先开始滚动或在表格中选中一行。")
            return

        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.model.set(row, "签到状态", "已签到")
        self.model.set(row, "签到时间", now)
        self._update_stats()
        self._save_cache()

        if self.no_repeat and row in self.current_idx_pool:
            try: self.current_idx_pool.remove(row)
            except ValueError: pass

    # ——— 清除签到 ———
    def clear_all_sign(self):
        if self.df is None or self.df.empty: return
        if QMessageBox.question(self, "确认", "清空所有签到状态？") != QMessageBox.Yes:
            return
        self.df["签到状态"] = ""
        self.df["签到时间"] = ""
        self.model.layoutChanged.emit()
        self._update_stats()
        self._rebuild_pool()
        self._save_cache()

    def clear_selected_sign(self):
        if self.df is None or self.df.empty: return
        sel = self.table.selectionModel().selectedRows()
        if not sel:
            QMessageBox.information(self, "提示", "请在表格中选中至少一行。")
            return
        for idx in sel:
            r = idx.row()
            self.model.set(r, "签到状态", "")
            self.model.set(r, "签到时间", "")
        self._update_stats()
        self._rebuild_pool()
        self._save_cache()

    # ——— 导出 ———
    def export_excel(self):
        if self.df is None or self.df.empty:
            QMessageBox.warning(self, "提示", "暂无数据可导出。"); return
        default_name = f"签到结果_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        path, _ = QFileDialog.getSaveFileName(self, "导出为 Excel", default_name, "Excel Files (*.xlsx)")
        if not path: return
        try:
            self.df.to_excel(path, index=False)
            QMessageBox.information(self, "成功", f"已导出：{path}")
        except Exception as e:
            QMessageBox.critical(self, "导出失败", f"{e}")

def main():
    app = QApplication(sys.argv)
    mw = MainWindow()
    mw.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
