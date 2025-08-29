#!/usr/bin/env python
# @File     : name_picker_minimal.py
# @Author   : 念安
# @Time     : 2025/8/29
# @Version  : V3.0
# @Desc     : 极简版抽签程序（无表格）

#!/usr/bin/env python
# -*- coding: utf-8 -*-

import sys, os, random
import pandas as pd

from PySide6.QtCore import Qt, QTimer
from PySide6.QtWidgets import (
    QApplication, QFileDialog,
    QWidget, QVBoxLayout, QHBoxLayout, QFrame, QMainWindow
)

# 这些控件来自 qfluentwidgets，用在普通 QWidget/QMainWindow 上也没问题
from qfluentwidgets import (
    setTheme, Theme, setFont,
    FluentIcon as FI, PrimaryPushButton, StrongBodyLabel
)

CACHE_FILE = "roster_cache.xlsx"

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
                found = a; break
            if a.lower() in lower_map:
                found = lower_map[a.lower()]; break
        if not found:
            return None
        cols_map[std_col] = found
    return cols_map


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("课堂点名·星星版")
        setTheme(Theme.LIGHT)
        setFont(self, 14)

        self.df = pd.DataFrame()
        self.rolling = False
        self.last_show_text = ""
        self.current_idx_pool = []

        # 定时器：固定滚动速度 50ms
        self.roll_timer = QTimer(self)
        self.roll_timer.setInterval(50)
        self.roll_timer.timeout.connect(self._roll_tick)

        self._build_ui()
        self._autoload_cache()

    def _build_ui(self):
        page = QWidget(self)
        page.setObjectName("mainPage")
        page.setMinimumSize(600, 400)

        # 顶部按钮（仅 导入 / 开始-暂停）
        self.btnImport = PrimaryPushButton(FI.FOLDER, "导入Excel", page)
        self.btnToggle = PrimaryPushButton(FI.PLAY, "开始", page)
        self.btnImport.clicked.connect(self.load_excel)
        self.btnToggle.clicked.connect(self.toggle_roll)

        topBar = QHBoxLayout()
        topBar.setSpacing(8)
        topBar.addWidget(self.btnImport)
        topBar.addWidget(self.btnToggle)
        topBar.addStretch(1)

        # 大屏显示结果
        self.bigText = StrongBodyLabel("——", page)
        self.bigText.setAlignment(Qt.AlignCenter)
        bigFrame = QFrame(page)
        bigFrame.setFrameShape(QFrame.StyledPanel)
        bigFrame.setStyleSheet("QFrame{background:rgba(0,0,0,0.05); border-radius:18px;}")
        bigLay = QVBoxLayout(bigFrame)
        self.bigText.setStyleSheet("QLabel{font-size:56px; font-weight:800;}")
        self.bigText.setMinimumHeight(160)
        bigLay.addWidget(self.bigText)

        # 布局
        root = QVBoxLayout(page)
        root.setContentsMargins(20, 16, 20, 16)
        root.setSpacing(12)
        root.addLayout(topBar)
        root.addWidget(bigFrame, stretch=1)

        self.setCentralWidget(page)

    # ---------- 数据缓存 ----------
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
            self._use_df(df.reset_index(drop=True))
        except Exception:
            pass

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择 Excel 文件", "", "Excel 文件 (*.xlsx *.xls)")
        if not path:
            return
        try:
            df = pd.read_excel(path, engine="openpyxl" if path.endswith("xlsx") else None)
        except Exception:
            return
        if df.empty:
            return

        cols = resolve_columns(df)
        if not cols:
            return

        df = df.rename(columns={cols["学号"]: "学号", cols["姓名"]: "姓名"})
        df = df.dropna(how="all").copy()
        df["学号"] = df["学号"].astype(str).str.strip()
        df["姓名"] = df["姓名"].astype(str).str.strip()
        self._use_df(df.reset_index(drop=True))
        self._save_cache()

    def _use_df(self, df: pd.DataFrame):
        self.df = df
        self.current_idx_pool = list(range(len(self.df)))

    # ---------- 抽取 ----------
    def toggle_roll(self):
        if self.df is None or self.df.empty:
            return
        if not self.rolling:
            if not self.current_idx_pool:
                self.current_idx_pool = list(range(len(self.df)))
            self.rolling = True
            self.roll_timer.start()
            self.btnToggle.setText("暂停")
            self.btnToggle.setIcon(FI.PAUSE.icon())
        else:
            self.roll_timer.stop()
            self.rolling = False
            self.btnToggle.setText("开始")
            self.btnToggle.setIcon(FI.PLAY.icon())

    def _roll_tick(self):
        if not self.current_idx_pool:
            self.current_idx_pool = list(range(len(self.df)))
        idx = random.choice(self.current_idx_pool)  # 允许重复抽取
        name = self.df.at[idx, "姓名"]
        sid = self.df.at[idx, "学号"]
        self.last_show_text = f"{sid}  {name}"
        self.bigText.setText(self.last_show_text)


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.resize(640, 420)
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
