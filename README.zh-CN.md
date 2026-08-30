# openpyxl-365

现代 Excel 工作簿的 OOXML 结构很复杂。一个 `.xlsx` 或 `.xlsm` 文件实际上是
由许多相互关联的 XML、关系文件、元数据、绘图和二进制部件组成的压缩包。
`openpyxl` 只能理解其中一部分，因此即使只修改一个单元格，普通的
`openpyxl.load_workbook(...).save(...)` 也可能删除或重写它不认识的内容，
造成：

- `FILTER`、`UNIQUE`、`XLOOKUP` 等公式被加上隐式交叉符号 `@`；
- 动态数组不再溢出，或 Excel 打开时提示修复记录；
- 外部加载项链接及其缓存被改写；
- x14 下拉框、线程批注、富数据或宏按钮丢失。

`openpyxl-365` 让你继续使用熟悉的 `openpyxl` API，同时在保存过程中保护这些
现代 Excel 365 所需的 OOXML 部件。

## 它如何工作

本项目先让 `openpyxl` 完成常规的单元格、公式、样式和工作表编辑，同时记录
`openpyxl` 无法安全往返保存的 OOXML 内容。保存后，再把相应的包部件、关系、
单元格属性和内容类型声明恢复回去。

这也是为什么问题并不只是“把 `=FILTER(...)` 写进单元格”：现代公式和工作簿
功能可能同时依赖公式文本、cell metadata、关系文件、内容类型和扩展部件。

## 主要能力

- 新建并保留 Excel 365 动态数组公式；
- 自动补充 `_xlfn.` / `_xlfn._xlws.` 前缀；
- 将公式栏写法 `D1#` 转换为文件内部的
  `_xlfn.ANCHORARRAY(D1)`；
- 保留既有动态数组元数据、公式属性、缓存结果及 Rich Data；
- 原样保留外部链接包，避免 `openpyxl` 重写外链缓存；
- 编辑 `.xlsm` 时保留 VBA；
- 保留 x14 扩展数据验证下拉框；
- 保留线程批注、人员信息和兼容 VML；
- 保留 DrawingML 形状、位置、样式和宏绑定按钮；
- 原子保存，并可要求 Excel 下次打开时重新计算。

## 使用方法

```python
from openpyxl365 import Excel365Workbook

doc = Excel365Workbook.open("原模型.xlsm")
doc.wb["Parameter"]["D11"] = 44_000_000_000
doc.set_dynamic_formula(
    "Summary",
    "H2",
    '=FILTER(A2:D100,D2:D100<>"")',
)
doc.save("修改版.xlsm", recalculate_on_open=True)
```

发布 PyPI 版本前，可以直接从 GitHub 安装：

```bash
pip install git+https://github.com/chengjiaxiao/openpyxl-365.git
```

## 边界

- 本项目不计算公式；最终结果仍应由 Microsoft Excel 或其他兼容计算引擎刷新。
- “保留”不等于能够用 Python 新建或修改所有高级对象。
- 可保留既有线程批注，但不能安全修改线程对话内容。
- 支持 DrawingML 宏按钮保留；ActiveX 和窗体控件尚未全面保证。
- 删除工作表、重构大量关系等结构性操作仍需专项核验。
- 修改任何文件都可能使原有数字签名失效。

## 为什么开源

这个项目来自真实复杂工作簿的兼容需求。Excel 365 的 OOXML 扩展问题往往需要
逐个拆包、对比和验证，重复排查成本很高。把这些经验沉淀为可测试、可复用的
开源工具，可以让更多使用 Python 自动化 Excel 的开发者少踩一些坑，也欢迎
社区贡献新的样本、边界条件和跨版本验证结果。

## 作者

- 程加骁：项目发起与维护
- GPT（OpenAI）：协作实现与文档整理

## 许可证

MIT。本项目为独立社区项目，与 `openpyxl` 项目及 Microsoft 无隶属或背书关系。
