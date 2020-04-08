# Vsto.Technology

Vsto.Technology将为Excel单元格右键菜单增加额外功能按钮。

## 功能
- **生成.** 将用户所选单元格应用至模板生成新的工作簿。 模板路径由app.config中`TemplatePath`指定。
- **生成子表.** Split user selection range by distinct values at Column B into different children sheets and update the origin sheet linking with these children sheets. Then add headers for each child sheet. The header is specificied as the top 4 rows of the origin sheet with a column width the same as user selection range. You can split the whole origin sheet or just split part of that, while the link will only be generate if it has been split.

## 安装准备
- **编辑hosts.** _这不是必须的。_ 如果你无法打开https://raw.githubusercontent.com ，请在本机上使用文本编辑器打开`C:\Windows\System32\drivers\etc\hosts`并添加`151.101.4.133 raw.githubusercontent.com`。
- **安装证书.** 下载Release页面中的 [JetSnail.cer](https://github.com/Snailya/vsto-technology/releases/download/1.0.0.0/JetSnail.cer) 。 安装方法见 https://jingyan.baidu.com/article/f0e83a255ab90222e59101ea.html 。

## 安装
运行Release页面中的`setup.exe`。

## 问题
- **System.ArgumentOutOfRangeException: 指定的参数已超出有效值范围** 清空`C:\Users\%Username%\AppData\Local\Apps\2.0`内文件.

---

# Vsto.Technology

Vsto.Technology is an excel add-in for generating quotation related excel sheets. It works by providing one or more commandbar controls on the right click menu of cell.

## Feature
- **Generate quotation workbook.** Create a new quotation workbook using template then add user selection range to each sheets. Workbook template can be specified by `TemplatePath` in app.config.
- **Generate children sheets.** Split user selection range by distinct values at Column B into different children sheets and update the origin sheet linking with these children sheets. Then add headers for each child sheet. The header is specificied as the top 4 rows of the origin sheet with a column width the same as user selection range. You can split the whole origin sheet or just split part of that, while the link will only be generate if it has been split.

## Installation Requirements
- **Edit your hosts.** This may not be neccessary but strongly relevant to your network settings because of some unspeakable reason. However, you may edit your hosts file at `C:\Windows\System32\drivers\etc\hosts` by adding a new line `151.101.4.133 raw.githubusercontent.com`.
- **Install certification.** Download [JetSnail.cer](https://github.com/Snailya/vsto-technology/releases/download/1.0.0.0/JetSnail.cer) in release page. See https://jingyan.baidu.com/article/f0e83a255ab90222e59101ea.html for installtion.

## Installation
Run `setup.exe`.

## Troubleshoot
- **System.ArgumentOutOfRangeException: Argument is out of range.** Please clear files inside `C:\Users\%Username%\AppData\Local\Apps\2.0`.
