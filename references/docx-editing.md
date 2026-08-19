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

---

## 1. 文件探测与解析

**扩展名不可信**：`.doc` 文件可能是 docx（ZIP）格式。先读文件头字节判断：
- `PK\x03\x04` → 真 docx（ZIP），直接用 zipfile 解析；
- `D0 CF 11 E0` → 真 .doc（OLE 二进制），需转换（Word COM 或 LibreOffice），本技能默认不支持，告知用户转存为 .docx；
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
