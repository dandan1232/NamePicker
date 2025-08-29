#!/usr/bin/env python
# @File     : show_fluentIcon.py
# @Author   : 念安
# @Time     : 2025-08-29 14:45
# @Verison  : V1.0
# @Desctrion: PyQt-Fluent-Widgets 1.x，它里面的图标都定义在 FluentIcon 这个枚举类里

from qfluentwidgets import FluentIcon as FI
print([n for n in dir(FI) if n.isupper()])
