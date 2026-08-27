# Third-party notices / 第三方依赖说明

本文件记录公开仓库直接声明的 Python 构建依赖。公开版运行时没有第三方 `pip` 依赖；Tkinter 随 Python/Tcl/Tk 环境提供。本仓库未复制第三方 Logo、品牌图标、字体、模板或付费素材。

This file records the Python build dependencies declared directly by the public repository. The runtime has no third-party `pip` dependency. Tkinter is provided by the Python/Tcl/Tk installation. No third-party logo, brand icon, font, template, or paid asset is copied into this repository.

| Project | Declared use | Version range | License | Official source / license |
| --- | --- | --- | --- | --- |
| setuptools | PEP 517 build backend | `>=68` | MIT | https://github.com/pypa/setuptools · https://github.com/pypa/setuptools/blob/main/LICENSE |
| wheel | Build-system wheel support | build-system dependency | MIT | https://github.com/pypa/wheel · https://github.com/pypa/wheel/blob/main/LICENSE.txt |
| PyInstaller | Optional local Windows executable build | `>=6,<7` | GPL-2.0-or-later with PyInstaller Bootloader Exception | https://github.com/pyinstaller/pyinstaller · https://github.com/pyinstaller/pyinstaller/blob/develop/COPYING.txt |

Python itself and Tcl/Tk are prerequisites supplied by the user's Python distribution; they are not vendored in this repository. Users should review the license texts that accompany their chosen Python distribution.

The project itself intentionally has no `LICENSE` file yet. Third-party notices do not grant a license to this project's source code.
