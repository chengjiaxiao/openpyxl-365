# openpyxl-365

继续使用熟悉的 `openpyxl` API 编辑 Excel 365 工作簿，同时保留普通读写可能丢失的部分 OOXML 功能。

[English](README.md) · [更新记录](CHANGELOG.md)

## 安装

需要 Python 3.10+：

```bash
pip install git+https://github.com/chengjiaxiao/openpyxl-365.git
```

## 用法

```python
from openpyxl365 import Excel365Workbook

doc = Excel365Workbook.open("原模型.xlsm")
doc.wb["Parameter"]["D11"] = 44_000_000_000
doc.set_dynamic_formula("Summary", "H2", '=FILTER(A2:D100,D2:D100<>"")')
doc.save("修改版.xlsm")
```

普通值、样式和公式通过 `doc.wb` 修改；直接赋值的 `=UNIQUE(A2:A100)` 等已识别动态公式也会自动处理。新建工作簿使用 `Excel365Workbook.create()`。

## 主要能力

- 保留动态数组、现代函数前缀、溢出引用和未改公式的缓存。
- 保留 VBA、外链包、Rich Data 和既有 Python-in-Excel 载荷。
- 保留 x14 下拉框、线程批注、DrawingML 形状及宏按钮。
- 使用 `rich_text=True` 打开时，保留富文本及其空格。

替换数组锚点时清理旧溢出缓存，并保留对原溢出区域的显式修改。
通过 `external_link_targets()`、`set_external_link_target()` 和
`set_external_reference()` 可查看、迁移或新建 HTTPS 工作簿外链；这些操作不验证访问权限，也不刷新数据。

保存先在本机缓存目录生成文件，再原子替换目标文件。可通过环境变量 `EXCEL365_STAGING_DIR` 指定暂存目录。

## 边界

- **不计算公式。** 使用结果前，应在 Microsoft Excel 中重算并保存。保存会移除计算链，默认不额外写入强制全量重算标记；需要显式请求时可传入 `recalculate_on_open=True`。
- 保留不代表支持编辑所有高级对象；线程对话编辑、ActiveX 和窗体控件尚不保证。
- 删除工作表等结构变更需要额外核验；修改工作簿可能使数字签名失效。

## 开发

```bash
pip install -e .
python -m unittest discover -s tests -p "test_*.py"
```

参与贡献见 [CONTRIBUTING.md](CONTRIBUTING.md)，Windows 原生 Excel 验证脚本见 `tests/validate_in_excel.ps1`。

MIT · 程加骁维护，GPT（OpenAI）协作实现与文档整理。与 Microsoft 及 openpyxl 项目无隶属关系。
