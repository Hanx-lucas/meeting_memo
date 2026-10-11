# 运行与渲染

## 依赖

- Python 3.10+；审查脚本依赖 `python-docx`，可执行 `python -m pip install -r requirements.txt`。
- DOCX 转 PDF 需要 LibreOffice；PDF 转逐页 PNG 需要 Poppler 的 `pdftoppm`。
- 使用宿主提供的文档运行环境时遵循其工具路径；否则从 PATH 查找工具。渲染脚本支持 `--soffice`、`--pdftoppm` 或环境变量 `SOFFICE_PATH`、`PDFTOPPM_PATH` 指定路径，无个人机器路径。
- 转写录音和检索公开资料需要宿主提供相应能力；附带脚本不负责语音识别或联网核验。

## Windows 与中文路径

- 使用 UTF-8 文件或可靠的 Unicode 参数传递，不把中文文件路径拼进可能损坏编码的内联 Python。必要时设置 `PYTHONUTF8=1`。
- 渲染脚本为每次运行建立独立临时目录，复制成英文文件名再转 PDF；最终 DOCX 保留用户指定路径。
- 工具执行失败后先看已生成的文件及错误原因，不重复尝试同一失败命令。可用时从标准文档渲染器切换到附带脚本。

```bash
python scripts/render_docx_windows.py --docx FINAL.docx --output-dir QA_DIR
python scripts/render_docx_windows.py --docx FINAL.docx --output-dir QA_DIR --soffice /path/to/soffice --pdftoppm /path/to/pdftoppm
```

脚本输出 PDF 与每页 PNG，支持 `--dpi` 调整分辨率。名称保留 Windows 后缀以兼容原调用，同时支持其他平台的 PATH 工具。输出目录会创建独立子目录，避免把旧图片误认成当前结果。

沙箱默认临时目录不可写时，可用 `--temp-dir` 指定允许写入的目录。

## 检查与交付

- 逐页确认无截断、重叠、缺字、错位或多余空白页。文件成功生成不等于视觉检查通过。
- 审查脚本默认适配用户模板，不强制中文编号或默认字号。只有内置格式使用 `--default-format`；字号异常会导致机器检查失败。
- 审查覆盖正文与表格中的文字，数字和英文词删除可作为检查线索；中文事实完整性、录音顺序、语义等仍需人工或模型逐段核校。
- 用 `--allow-missing-token` 明确记录有依据的词元更改；不能批量忽略报告以强行通过。旧稿里无信息重复合并造成的出现次数减少不直接算丢失，仍需内容对照。
- 集中审查和修复，只有实际发现问题时才重复渲染。交付最终 DOCX 与真实待确认项，工作中间文件不作为成品。
