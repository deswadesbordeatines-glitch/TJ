采购发票识别工具 Windows 离线版

一、如何用 GitHub Actions 生成 Windows 离线发布包

1. 把本项目推送到 GitHub 仓库。
2. 打开 GitHub 仓库页面。
3. 进入 Actions。
4. 选择工作流：Build Windows Offline Release。
5. 点击 Run workflow。
6. 等待任务完成。
7. 在任务页面底部 Artifacts 区域下载：
   procurement-invoice-tool-windows-offline
8. 下载后得到 zip 文件：
   采购发票识别工具-Windows离线版.zip

这个 zip 是在 GitHub 的 windows-latest 环境中用 PyInstaller 打包生成的 Windows 离线版。


二、如何复制到离线电脑

1. 在有网络的电脑上下载 GitHub Actions 生成的 zip。
2. 解压：
   采购发票识别工具-Windows离线版.zip
3. 解压后会看到文件夹：
   采购发票识别工具
4. 请把整个“采购发票识别工具”文件夹复制到 U 盘或移动硬盘。
5. 再把整个文件夹复制到目标离线 Windows 电脑。

注意：不要只复制 exe。必须复制整个文件夹，因为依赖库都在同一个文件夹内。


三、如何运行

在离线电脑上打开文件夹：

采购发票识别工具

双击：

采购发票识别工具.exe

目标电脑不需要安装 Python，不需要网络。


四、输出 Excel 在哪里

默认输出位置：

C:\Users\用户名\Documents\采购发票台账.xlsx

也可以在软件界面点击“选择输出位置”，手动选择保存路径。


五、如果识别失败如何查看 error_log.txt

识别失败时，会在输出 Excel 同一个文件夹下生成：

error_log.txt

例如默认位置：

C:\Users\用户名\Documents\error_log.txt

此外还会生成调试日志：

invoice_debug_log.txt

这个文件会记录 PyMuPDF 和 pdfplumber 读取到的 PDF 原始文字，以及每一步字段识别结果。


六、功能范围

本离线版只保留基础识别功能：

- 发票类型
- 发票号码
- 开票日期
- 购买方名称
- 销售方名称
- 金额
- 税额
- 价税合计

本工具不调用任何网络接口，不使用 OpenAI、DeepSeek、Qwen 等在线 API。
本工具不包含 OCR，只处理带可复制文字层的电子发票 PDF。


七、可选：在 Windows 打包机上本地打包

如果你以后有一台联网 Windows 打包机，也可以在该机器上进入本目录，执行：

build_windows.bat

本地打包完成后，会生成：

dist\采购发票识别工具\采购发票识别工具.exe
