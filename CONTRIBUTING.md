# 参与贡献

感谢你对 md2word 的兴趣。下面是本地开发和提交改动的简要说明。

## 环境

- Windows、macOS 或 Linux
- Python 3.10 或更高版本
- 可选：[Pandoc](https://pandoc.org/installing.html)、Microsoft Word 或 LibreOffice（公式粘贴和 PDF 导出会用到）

```bash
python -m pip install -r requirements-dev.txt
python main.py
```

Windows 也可双击 `start.bat`；macOS / Linux 可用 `./start.sh`。

## 测试

```bash
python -m unittest tests.test_md2word -v
```

需要 Pandoc 的用例在未安装时会自动跳过；需要 Word 的 PDF 用例同样如此。提交前请尽量跑完整套测试。

## 打包 exe（Windows）

```bat
python build_exe.py
```

产物在 `dist\md2word.exe`。构建日志、`build/`、`dist/` 和 exe 都已加入 `.gitignore`，请不要提交。

## 代码约定

- 包代码放在 `md2word/`，入口是 `main.py`
- 示例 Markdown 放在 `examples/`
- 不要把本机路径、构建日志、`.docx` / `.pdf` 产物写进仓库
- 改动转换逻辑时，请在 `tests/test_md2word.py` 补充对应用例
- Windows 专用能力放在 `clipboard_win.py` / `wordcom.py` 的 COM 分支；其它平台走 `clipboard.py` 和 AppleScript / LibreOffice

## 提交 Pull Request

1. Fork 本仓库并创建分支
2. 做最小必要改动
3. 运行测试
4. 打开 PR，说明动机、行为变化和验证方式
