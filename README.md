# 🎓 课堂点名系统 · Fluent 风格

一个基于 **Python 3.11 + PySide6 + PyQt-Fluent-Widgets** 的桌面应用，支持从 Excel 导入学生名单，进行随机点名、签到记录和导出功能。  
界面使用 Fluent Design 风格，美观现代，支持浅色/深色主题切换。

---

## ✨ 功能特性

- 📂 **导入 Excel**：支持 .xlsx / .xls，需包含「学号」「姓名」列（自动识别常见别名）。  
- 🎲 **随机点名**：支持不重复抽取、滚动速度调节、自动签到延迟。  
- ✅ **签到管理**：一键签到 / 清空签到 / 清除选中行签到。  
- 🔍 **搜索功能**：按学号或姓名实时过滤。  
- 📊 **统计显示**：显示总数、已签到数、未签到数，进度条动态更新。  
- 🎨 **主题切换**：浅色 / 深色主题切换。  
- 🥚 **彩蛋功能**：点击左下角彩蛋按钮，会弹出彩蛋提示。  
- 💾 **缓存机制**：保存上次导入的名单，下次启动自动加载（签到状态自动清空）。  
- 📤 **导出已签到名单**（菜单里提供）。  

---

## 📦 环境依赖

项目基于 **Python 3.11**，主要依赖：

- `PySide6`
- `PySide6-Fluent-Widgets`
- `pandas`
- `openpyxl`
- `pyinstaller`（打包用）

安装依赖：

```bash
pip install -r requirements.txt
```

`requirements.txt` 内容示例：

```
pyside6
PySide6-Fluent-Widgets
pandas
openpyxl
pyinstaller
```

---

## ▶️ 运行方式

在项目根目录运行：

```bash
python name_picker.py
```

---

## 🖼 图标设置

项目支持自定义图标：

- 准备一个 `app.ico` 文件（推荐 256×256 多尺寸 ico）  
- 放在 `NamePicker` 根目录  
- 打包时会自动嵌入 exe  

如果你只有 png/jpg，可以用 [convertico.com](https://convertico.com/) 转换成 ico。

---

## 🔨 打包成 EXE

在 PowerShell 中运行（推荐 **非管理员权限**）：

```powershell
pyinstaller --noconfirm --onefile --windowed `
  --name 点名系统 `
  --icon=app.ico `
  --collect-all qfluentwidgets `
  --hidden-import openpyxl `
  name_picker.py
```

打包完成后：

- 可执行文件在 **`dist/点名系统.exe`**  
- 可以直接拷贝给其他 Windows 用户使用，无需安装 Python  

> ⚡ 如果不需要单文件，可以改用 `--onedir`，启动更快，方便调试。

---

## 📂 项目结构

```plaintext
NamePicker/
│── name_picker.py        # 主程序
│── requirements.txt      # 依赖列表
│── app.ico               # 应用图标（可选）
│── dist/                 # 打包后生成的 exe 目录
│── build/                # 打包过程临时文件
│── 点名系统.spec         # PyInstaller 配置
│── README.md             # 使用说明
```

---

## ⚠️ 注意事项

1. Excel 必须包含「学号」「姓名」列，否则无法导入。  
2. 第一次运行会在目录下生成 `roster_cache.xlsx`，用于缓存名单。  
3. 如果目标电脑打开 exe 没反应，尝试用 `--onedir` 模式打包，或在命令行里运行查看报错。  
4. 杀毒软件可能会误报，建议使用 `--onedir` 分发，或对 exe 做签名。  

---

## 📸 界面预览

![img.png](img.png)

![img_1.png](img_1.png)


![img_2.png](img_2.png)



---

## 👨‍💻 作者

开发：念安(dandna1232)  
风格：微软 Fluent Design  