# docx 编辑技术手册（docx-editing.md）

本文档收录修改 .docx 施工方案时的核心技术细节与踩坑记录。涉及 OOXML 结构，操作前先通读本节相关部分。

## 目录

1. 文件探测与解析
2. 文本定位与替换
3. 含图形段落的处理
4. 表格操作
5. 图片与嵌入对象保护
6. 自动编号与目录
7. 章节重排（块级重排）
8. 验证与打包
9. 索引系统：dump 索引 vs 顶层元素（重排/切块必读）
10. 样式层级调整检查清单（对齐模板/重排/改标题层级）
11. 自动编号失效的兜底：显式编号
12. 表格单元格文本替换与"内容容器表"
13. WPS/Word 保存会"规范化"文档
14. 渲染级验证（转 PDF）
15. WPS 文件读取与 COM 转换（输入可为 WPS 格式，产物必须 Word）

---

## 1. 文件探测与解析

**扩展名不可信**：`.doc` 文件可能是 docx（ZIP）格式。先读文件头字节判断：
- `PK\x03\x04` → 真 docx（ZIP），直接用 zipfile 解析；
- `D0 CF 11 E0` → 真 .doc（OLE 二进制），需 COM 转换（Word/WPS，流程见第 15 节）——**输入可为 WPS 相关格式，必须能读取**；
- 其他 → 不是 Word 文档，报告用户。

**WPS 生成的 docx 暗坑**：`word/_rels/document.xml.rels` 中可能存在 `Target="../NULL"` 的悬空关系，导致 python-docx 打开时报 `KeyError: "There is no item named 'NULL' in the archive"`。处理：改用 `zipfile` + `lxml` 直接读写包；若确需 python-docx，可先移除该关系再打开。本技能统一用 zipfile+lxml 方案，不依赖 python-docx 打开（python-docx 仅用于最终验证）。

**解析流程**：
1. 打开 zip，读取 `word/document.xml`；
2. `root.iter("w:p")` 收集全部段落（含表格内、文本框内段落——注意文本框段落也在 document.xml 中）；
3. dump 文本：`[段落索引] <样式> 文本` 到文件，便于检索定位；
4. 读取 `word/styles.xml` 与 `word/numbering.xml` 确认标题样式与编号定义。

## 2. 文本定位与替换

**段落定位**：优先用"包含唯一关键子串"定位段落，不要依赖固定索引（插入/删除会漂移）。

**跨 run 拆分（必踩的坑）**：同一段文字可能被拆到多个 `<w:t>`（如 `"JGJ 46"` 和 `"-2005"` 分属不同 run）。逐 run 替换会漏改。正确做法——段落级拼接替换：

```python
def replace_in_para(p, old, new):
    ts = p.findall(".//w:t")
    merged = "".join(t.text or "" for t in ts)
    if old not in merged:
        return False
    merged = merged.replace(old, new)
    ts[0].text = merged
    for t in ts[1:]:
        t.text = None
    return True
```

注意：`.//w:t` 是后代遍历。**若段落内嵌图形（drawing/object），此方法会把图形文本框内的 w:t 也合并进来并清空，破坏图形**——见第 3 节。

**空单元格**：`set_text` 前若无 `w:t`，需先创建 `<w:r><w:t>`。

**删除段落**：`p.getparent().remove(p)`。

## 3. 含图形段落的处理（组织架构图等）

- 组织架构图/流程图常以一个"容器段落"呈现：段落内嵌 `<w:drawing>`（新格式 `<mc:Choice>` 内 `wpg` 组）和 `<mc:Fallback>`（VML `w:pict` 文本框），图形内每个文本框是一个独立 `<w:p>`。
- **危险操作**：对容器段落做 `.//w:t` 全量拼接替换，会把图形内文本框的 w:t 一起重写/清空，毁掉整张图。
- **安全做法**：
  - 文本替换只处理"段落直接子级 run"的 w:t：`p.findall("./w:r/w:t")`；
  - 图形内文字要单独定位：直接遍历 `root.iter("w:p")` 找到文本框段落（其祖先含 `w:txbxContent`），再对该段落做替换；
  - 判断段落是否含图：`p.findall(".//w:drawing")` 或 `.//w:object` 非空。
- 注意 `<mc:Choice>` 与 `<mc:Fallback>` 是同一图形的两种表示（Word 只渲染其一），**两处文字都要同步改**，否则切换兼容模式时出现旧文字。

## 4. 表格操作

- **行定位**：遍历 `tbl.findall("w:tr")`，按单元格文本特征匹配目标行；列内单元格用 `tr.findall("w:tc")`，`tc.findall("w:p")` 取段落。
- **改单元格**：`set_text(tc 内第一个 w:p, 新文本)`；空单元格先建 run。
- **删除行**：`tr.getparent().remove(tr)`。删除前检查该行是否有 `vMerge` 纵向合并：
  - 若删除的是合并**起始行**（vMerge 无 val 或 val="restart"，且含"序号""分类"等文本），需先把起始行单元格内容与 vMerge 标记迁移到新的首行（深拷贝起始行 tc 替换新首行 tc），否则合并关系断裂、内容丢失；
  - 合并续行（val="continue"）可直接删。
- **插入行**：`copy.deepcopy` 一个现有行做模板，清空各单元格文本后填入新内容，`tr.addnext(new_tr)`——保持列宽与格式一致。
- **表标题**：表格标题（"表 X.Y-Z ……"）通常在表格外的段落（样式 caption 类），定位表格时不要用标题文本，用表格内特征文本。

## 5. 图片与嵌入对象保护

- **默认不做任何图片操作**：不删除、不移动、不替换 `word/media/*`；重新打包时原样复制所有部件。
- **rId 关系完整性**：修改 document.xml 后，`word/_rels/document.xml.rels` 中所有被引用的 rId 必须存在；被删除内容引用的 rId 若不再使用可保留（无害）或清理。新增图片需：加 media 文件 → 加 rels 关系 → document.xml 加 `w:drawing` 引用 → 必要时补 `[Content_Types].xml`。
- **嵌入式对象（Visio/AutoCAD OLE）**：`<w:object>` 含 `v:imagedata r:id`（预览图）与 `o:OLEObject r:id`（oleObject 二进制）。本技能不编辑 OLE 内容；若意见要求改流程图等内容，向用户说明需提供新图或改用文本流程表述。

## 6. 自动编号与目录

- 正文标题编号来自**样式级编号**（styles.xml 中 heading 样式的 numPr + numbering.xml），Word 渲染时按文档顺序自动计数。改标题文本不影响编号；**调整标题顺序/层级（style）会改变渲染编号**。
- 定位用户引用的章节号（如"3.5.2.2节"）时：按文档顺序模拟编号计数器（样式 1/2/3/4 → 层级 0/1/2/3，父级变化时子级计数重置），得到每个标题的渲染编号，再与用户引用对照。
- **目录（TOC）可能是过期快照**：TOC 条目文本带字面编号（静态文字），与正文渲染编号不一致时以正文为准；修改完成后提示用户在 Word 中"更新域"刷新目录（TOC 域 `TOC \o "1-3"`）。
- **交叉引用**：方案正文中"详见 X.Y.Z 节"的引用在章节重排后可能失效，需全文档检索并修正为新的实际编号。

## 7. 章节重排（块级重排，用于对齐模板等大调整）

1. 按标题切块：body 直接子元素中，标题段（样式为标题层级）开始一个新块，块 = 标题段 + 其后所有兄弟元素直到下一个同级或更高级标题；
2. 定义新顺序（引用原 body 子元素区间 + 新增块），构建新元素序列；
3. 先收集所有元素引用，再清空 body 依序 append（避免边遍历边修改）；
4. 调整标题层级时同步改 pStyle（如 style=3 → style=2 提升一级，其子标题 style=4 → style=3），否则会出现 "5.0.1" 这类异常编号（父级计数为 0）；
5. 重排后检查：模拟编号无异常、交叉引用修正、表图编号（"表 3.31" 等手写编号）按新章节更新、vMerge 跨块行处理。

## 8. 验证与打包

- 验证清单见 SKILL.md；工具函数见 `scripts/docx_edit.py`（dump、定位、替换、表格行删改、编号模拟、rId 校验）。
- 重新打包：`zipfile.ZipFile(DST, "w", ZIP_DEFLATED)` 遍历原包所有条目，仅替换 `word/document.xml`（必要时 rels/Content_Types），其余原样写入。
- 写文件用临时文件 + `os.replace` 原子替换，避免同路径读写冲突。

## 9. 索引系统：dump 索引 vs 顶层元素（重排/切块必读）

**dump 索引的语义**：`dump_docx_text` 输出的 `[N]` 是 `root.iter("w:p")` 的**全文档顺序、0-based 索引**——它数的是**所有** `w:p`，包括表格单元格内的段落、文本框（`w:txbxContent`）内的段落。空段落也占索引（dump 只显示非空文本，索引会跳号）。

**两个绝不能混用的系统**：
1. **dump 索引**（0-based，全文档 p 顺序）——用于定位"某段在文档里排第几"；
2. **body 顶层元素序列**（body 直接子元素：顶层 w:p 与 w:tbl）——用于切块、移动、重排。

不要用"遍历 body 子元素 + 自增计数器"来复现 dump 索引：文本框/表格内的段落不会被计入，二者从文档中第一个含文本框或表格的位置开始就产生偏移（常为 ±1，且随内容漂移），会导致范围切块错乱（内容重复发射、错位）。

**正确做法**：先建精确映射——`pidx = {id(p): i for i, p in enumerate(root.iter("w:p"))}`（**0-based，与 dump 索引一致**，与工具函数 `build_top_level_index` 同口径，全程只用这一套）；顶层元素的"位置"取其**首个后代 w:p** 的索引；**表格是单一顶层元素**（位置 = 首单元格索引），切块时整表移动，不能按单元格切。

工具函数：`build_top_level_index(body)`（返回 `(first_p_idx, element)` 列表）与 `elems_between(body, start_idx, end_idx)`（取 `[start, end)` 区间内所有顶层元素，含表格）。区间边界引用 dump 索引时，注意先确认该索引是顶层段落（表格单元格索引不能作边界）。

## 10. 样式层级调整检查清单（对齐模板/重排/改标题层级）

调整标题样式层级（pStyle 改到别的 heading 样式，或把某个样式定义替换成 heading 语义）时，**编号不显示/错乱**的常见根因与检查项：

1. **basedOn 链**：替换/复制样式定义时，`<w:basedOn>` 会被一起拷贝。若新样式名与原 basedOn 相同 → **自引用**（如 heading 2 basedOn 自己），Word/WPS 无法解析样式链，整个编号体系静默失效。检查：任何样式的 basedOn 不得等于自身；被替换的样式若有其他样式基于它（`basedOn="3"` 的还有谁），要一并修正（通常改为基于 Normal 或上一级 heading）。
2. **numPr 的 ilvl**：`<w:numPr><w:ilvl w:val="N"/><w:numId w:val="M"/></w:numPr>`——ilvl 决定编号格式层级（ilvl 0→%1，ilvl 1→%1.%2，ilvl 2→%1.%2.%3……）。从另一个样式拷贝 numPr 时，ilvl 要按目标层级重设，不能照搬源样式。
3. **不要 deepcopy 正在被修改的样式**：如果在同一循环里"用样式 B 的定义生成样式 A，随后又修改 B"，A 拷到的是 B 的**最终状态**还是**中间状态**取决于 XML 元素顺序——踩过 ilvl 拷错的坑。做法：先完整收集所有样式定义，再统一生成新定义，最后一次性替换。
4. **outlineLvl 对齐**：heading N 的 `<w:outlineLvl w:val="N-1"/>`（导航窗格层级），与 ilvl 同步改。
5. **numbering.xml 校验**：确认 `numId → abstractNumId` 映射（`<w:num>` 下的 `<w:abstractNumId>`），并检查被引用 abstractNum 的各级 `<w:lvl><w:lvlText w:val="..."/>`——lvlText 可能是乱码（如 `第%5条` 残留）或缺失。标题最多用到 ilvl N 时，0..N 级的 lvlText 必须都是干净的 `%1`/`%1.%2`/`%1.%2.%3` 形式。
6. **验证渲染**：逻辑编号（COM ListString 或模拟编号）≠ 显示。改完样式后必须转 PDF 看真实渲染（见第 14 节）。

## 11. 自动编号失效的兜底：显式编号

当自动编号修复后仍无法保证渲染（样式链损坏严重、查看器兼容性问题、用户环境不可控）时，可降级为**显式编号**：

1. 遍历所有标题段落，按 `simulate_heading_numbers` 的计数器算出每个标题的编号；
2. 把 `编号 + 半角空格` 作为文本 run 插入标题段落开头（编号文本与标题文本之间必须保留空格）；
3. **`w:t` 尾部空格会被 XML 序列化丢弃**，编号 run 的 `w:t` 必须设 `xml:space="preserve"`（`{http://www.w3.org/XML/1998/namespace}space` 属性），否则 "1 工程概况" 变成 "1工程概况"；
4. 关闭自动编号：删除标题样式 pPr 里的 `<w:numPr>`（或 numId 置 0），否则自动编号+显式文本双重显示；
5. 目录（TOC 域）更新后按标题文本重建，会自然带上显式编号，格式与原方案目录一致。

代价：文档失去自动重编号能力（增删章节不会自动改号）。仅当自动编号修复失败或验证不过时使用，并告知用户。

## 12. 表格单元格文本替换与"内容容器表"

- **替换必须覆盖表格单元格**：段落级替换（`p.findall(".//w:t")` 合并）只处理单个段落；做全文替换时要遍历**元素的全部后代** `w:p`（`el.iter("w:p")`），否则表格里的旧文本（人名、规范号、数值）会漏改。
- **内容容器表**：有些方案会把一大段连续内容（如整个应急预案：信息表+正文+医院地址+线路图）包在**一个表格**里（几十个单元格）。表现：dump 里这些段落索引连续、样式统一，但按索引切块时它们属于**同一个 tbl 顶层元素**——只能整表移动，无法拆到不同章节。处理：要么整表放一个章节并在别处引用，要么在技能边界内告知用户"该内容位于一张大表格中，整块移动"。
- 行定位用单元格文本特征（`tc` 内 `w:p` 文本），不要用表格外的题注文本。

## 13. WPS/Word 保存会"规范化"文档

用 WPS COM（或 Word COM）打开文档再 `Save()`，会重写整个包，已知副作用：
- **合并/重排 run**：编号 run 与正文 run 合并（显式编号场景会破坏 run 结构）；
- **丢弃 w:t 尾部空格**（没有 `xml:space="preserve"` 时）；
- **清理未引用 media**（图片数可能减少）；
- **重写样式集**（样式 ID、TOC 样式可能变化）。

因此：**修改一律用 zipfile+lxml 直接写回**（只替换 document.xml / styles.xml / rels），不要用 COM 保存作为最终步骤。COM 只用于两类动作，且动作后必须重新验证（python-docx 打开 + media 计数 + 关键文本抽查）：
- 转换输入文件（.doc/.wps → .docx）；
- 更新目录域（`doc.TablesOfContents.Item(1).Update()`）后立即导出 PDF 验证，再决定是否保留该次保存。

## 14. 渲染级验证（转 PDF）

逻辑层验证（XML 良构、字符串断言、COM `ListFormat.ListString` 返回的编号）**不等于渲染显示**。对编号、表格、图片类修改，必须做渲染级验证：

1. 用 COM（Word 或 WPS，探测顺序见第 15 节）打开输出文件，`ExportAsFixedFormat(路径, 17)` 导出 PDF；
2. 用 pypdf/pdfplumber 提取 PDF 文本；
3. 断言关键行真实渲染：如标题行出现 `1 工程概况`、`1.1 基坑工程概况和特点` 形态；表格关键单元格文本出现；喷锚等应删除的内容不出现。

COM 不可用或导出失败时，退而求其次：用模拟编号 + 告知用户"已按结构验证，请打开确认渲染"。

## 15. WPS 文件读取与 COM 转换（输入可为 WPS 格式，产物必须 Word）

**输入兼容性**：用户给的文件可能是 (a) 真 .doc（OLE，头 `D0 CF 11 E0`）、(b) WPS 生成的 .docx（头 PK，但包内有 WPS 暗坑：`../NULL` 悬空关系、mc:AlternateContent 等）、(c) WPS 专有 `.wps` 老格式、(d) 标准 .docx。**全部必须能读取**；**最终产物必须是标准 Word .docx**（"原名（修改稿）.docx"）。

**探测顺序**：
1. 文件头：`PK\x03\x04` → zip 路线（zipfile+lxml 直接解析，天然绕开 WPS 悬空关系）；
2. `D0 CF 11 E0` 或 `.wps` 扩展名 → COM 转换；
3. 其他 → 报告用户。

**COM 转换（.doc/.wps → .docx）**：按可用性依次尝试：
1. Word COM（`New-Object -ComObject Word.Application`）——注意注册表 `HKLM\SOFTWARE\Classes\Word.Application\CurVer` 可能残留旧版 ProgID（如 `.12`）导致 `New-Object Word.Application` 报"未能引发事件"或 `CO_E_SERVER_EXEC_FAILURE`；可改试显式版本（`Word.Application.16`）或先 `GetActiveObject` 连接已启动实例；
2. **WPS COM**（`New-Object -ComObject KWPS.Application`，先查 `C:\Program Files (x86)\Kingsoft`、`C:\Program Files\Kingsoft`、`%LOCALAPPDATA%\Kingsoft` 确认安装了 WPS）——API 与 Word 兼容：`Documents.Open(路径, 只读, 只读)` + `SaveAs2(目标, 12)`（12 = wdFormatXMLDocument）；
3. 都失败 → 告知用户转存 .docx，不硬解。

**转换后立即验证（Q4 决策）**：转换产物必须通过——python-docx 能打开（WPS 产物常见 `KeyError: 'NULL'`，处理：从 `word/_rels/document.xml.rels` 移除 Target 为 `../NULL` 的关系）或 zipfile+lxml 能解析 + rId 无悬空；验证不过先修再继续。转换稿与最终稿都保持这个标准。
