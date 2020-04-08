# Vsto.Technology

Vsto.Technology is an excel add-in for generating quotation related excel sheets. It works by providing one or more commandbar controls on the right click menu of cell.

## Feature
- **Generate quotation workbook.** Create a new quotation workbook using template then add user selection range to each sheets. Workbook template can be specified by `TemplatePath` in app.config.
- **Generate children sheets.** Split user selection range by distinct values at Column B into different children sheets and update the origin sheet linking with these children sheets. Then add headers for each child sheet. The header is specificied as the top 4 rows of the origin sheet with a column width the same as user selection range. You can split the whole origin sheet or just split part of that, while the link will only be generate if it has been split.

## Installation Requirements
- **Edit your hosts.** This may not be neccessary but strongly relevant to your network settings because of some unspeakable reason. However, you may edit your hosts file at C:\Windows\System32\drivers\etc\hosts by adding a new line `151.101.4.133 raw.githubusercontent.com`.
- **Install certification.** Download [JetSnail.cer](https://github.com/Snailya/vsto-technology/releases/download/1.0.0.0/JetSnail.cer) in release page. See https://jingyan.baidu.com/article/f0e83a255ab90222e59101ea.html for installtion.

## Installation
Run setup.exe.

## Troubleshoot
- **System.ArgumentOutOfRangeException: Argument is out of range.** Please clear files inside `C:\Users\%Username%\AppData\Local\Apps\2.0`.
