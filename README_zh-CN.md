# PDF 转 Word

一个离线、跨平台的 PDF → DOCX 桌面程序，可在 Windows 和 macOS 上运行。文件只在本机处理，不需要安装 Microsoft Word 或 Adobe Acrobat。

## 三种转换模式

| 模式 | 适用情况 | 可编辑性 |
| --- | --- | --- |
| 保留版式（推荐） | 普通电子 PDF，包含文字、图片或表格 | 大部分内容可编辑 |
| 完全一致 | 排版特别复杂，优先保证外观 | 不可编辑，每页作为高清图片 |
| OCR 识别 | 扫描件、照片 PDF | 文字可编辑，排版会简化 |

## 直接运行源码版

需要先安装 64 位 Python 3.11 或 3.12。

### Windows

双击 `run_windows.bat`。第一次运行会自动建立独立环境并安装组件，以后仍可双击启动。

### macOS

双击 `run_macos.command`。若系统阻止打开，请按住 Control 点击文件，选择“打开”。也可以在终端运行：

```bash
chmod +x run_macos.command
./run_macos.command
```

## 打包成真正的 `.exe` / `.app`

PyInstaller 不能跨系统打包，因此 Windows 版本需要在 Windows 上构建，macOS 版本需要在 macOS 上构建。

- Windows：双击 `build_windows.bat`，结果位于 `dist\PDF-to-Word\PDF-to-Word.exe`。
- macOS：运行 `chmod +x build_macos.command && ./build_macos.command`，结果位于 `dist/PDF-to-Word.app`。

打包后的普通转换不要求用户安装 Python。若要公开分发，建议再进行代码签名；未签名程序可能触发 Windows SmartScreen 或 macOS Gatekeeper 提示。

如果要向其他人或公司公开分发打包程序，请同时查看 `THIRD_PARTY_NOTICES.md`，并确认依赖组件的许可证符合你的分发方式。

## 使用 GitHub Actions 自动生成三平台程序

项目包含 `.github/workflows/build-desktop.yml`。上传到 GitHub 后：

1. 打开仓库的 **Actions** 页面。
2. 选择 **Build desktop apps**。
3. 点击 **Run workflow**。
4. 构建完成后，在本次运行页面下载三个 Artifacts：
   - `PDF-to-Word-Windows-x64`
   - `PDF-to-Word-macOS-Apple-Silicon`
   - `PDF-to-Word-macOS-Intel`

每个 Artifact 内都是一个 ZIP；解压后可得到 Windows 程序目录或 macOS `.app`。这些程序未经代码签名，因此系统第一次打开时可能显示安全提示。

## 扫描件 OCR

OCR 模式需要本机安装 Tesseract：

- macOS：`brew install tesseract tesseract-lang`
- Windows：安装 Tesseract OCR，并在安装时选择需要的语言数据。

中文简体加英文使用 `chi_sim+eng`，中文繁体加英文使用 `chi_tra+eng`。如果只装了英文语言包，请选择 `eng`。

## 命令行用法

```bash
python app.py input.pdf -o output.docx
python app.py input.pdf -o output.docx --pages 1-3,5 --mode exact
python app.py scan.pdf -o scan.docx --mode ocr --ocr-language chi_sim+eng
```

## 注意事项

- PDF 和 DOCX 的排版模型不同，复杂分栏、公式、艺术字或特殊字体不一定能完全还原。
- “完全一致”模式最稳定，但输出页面是图片，不能直接修改文字。
- OCR 的准确率取决于扫描清晰度、方向和已安装的语言包。
- 程序支持选择页码和打开带密码的 PDF。
