# construction-scheme-modify

![npm version](https://img.shields.io/npm/v/construction-scheme-modify)
![License: MIT](https://img.shields.io/badge/license-MIT-blue)
![npm downloads](https://img.shields.io/npm/dm/construction-scheme-modify)
![GitHub stars](https://img.shields.io/github/stars/PengYuyan888/construction-scheme-modify)

修改施工方案/专项方案（Word .docx）并编写"修改意见回复"的 Agent 技能（Skill）。

Revise construction schemes / special schemes (Word .docx) per review comments and produce a formal modification reply — an agent skill for DSH / Claude Code style skill runtimes.

---

## 功能特性 / Features

- **按意见修改方案**：删除 / 修改 / 补充 / 调整四类修改意见逐条落位，输出"原名（修改稿）.docx"，原文件不动。
- **修改意见回复**：按固定格式生成"问题N：… / 回复：已修改，详xxxxx"，每条两行以内，只陈述处理结果、不解释技术原因。
- **自动冲突检查**：识别文档内部矛盾（数值/单位/日期/名称不一致）、模板残留（旧地名、旧奖项、旧项目名）、废止规范（JGJ46-2005 等），列出问题清单待用户决策。
- **规范版本校核**：更新废止规范时经 web 搜索确认现行版本号，不凭记忆编造。
- **图片零改动**：不编辑任何图片/嵌入对象（Visio 流程图等），做不到的事在过程沟通中告知用户，不写进最终回复。
- **文件查看降级链**：zipfile+lxml 默认解析 → python-docx 备选 → 查文件头兜底，不卡住。
- **WPS 兼容**：处理 WPS 生成的悬空关系（`Target="../NULL"`）等 python-docx 打开失败的文档。
- **WPS 老格式输入**：真 .doc / .wps 专有格式经 COM 转换（Word 失败自动降级 WPS）读取，转换产物立即验证；最终产物始终是标准 Word .docx。
- **渲染级验证**：编号/表格/图片类修改转 PDF 核对真实渲染，逻辑编号不等于显示。

---

## 安装 / Installation

### 方式一：npm（推荐，自动更新）

```bash
npm install -g construction-scheme-modify
# 或项目内安装：
npm install construction-scheme-modify
```

配合 Vercel [skills CLI](https://github.com/vercel-labs/skills)（Agent Skills 生态的官方工具）：

```bash
npx skills add construction-scheme-modify   # 从 npm 包安装
# 或从 GitHub 直接安装：
npx skills add PengYuyan888/construction-scheme-modify
```

### 方式二：GitHub 仓库

```bash
git clone https://github.com/PengYuyan888/construction-scheme-modify.git
```

### 方式三：手动放置（DSH / Claude Code）

将 `construction-scheme-modify/` 目录放入技能目录（如 `~/.agents/skills/` 或 `~/.claude/skills/`），重启会话后自动生效。

### 依赖 / Dependencies

- Python 3 + `lxml`（docx 解析，必需）
- `python-docx`（仅最终验证用，可选）

---

## 使用方法 / Usage

直接以自然语言提出修改意见即可触发，例如：

> 项目里有一个"××项目地下防水施工方案.docx"文件，以下是修改意见：
> 1、方案中标注的"（根据项目特点删减）"请删除。
> 2、2.4.1节地下室防水概况特殊部位的章节细部做法图与后文不一致。
> 3、2.1节地下室面积41.3㎡与后文3.2.2节施工部署中卸粮坑地下室40㎡数值不统一，请统一为41.3㎡。
> 4、4.2节施工流程体现的是2道防水卷材，图纸为涂料+卷材，请按图纸做法修改。
> 请根据意见修改方案，并生成修改意见回复。

技能工作流：

```
通读方案 → 逐条定位意见 → 汇报修改计划与问题清单（等你确认）
→ 执行修改 → 验证 → 输出"（修改稿）.docx" + 修改意见回复
```

### 修改意见回复格式

每条意见按固定格式回复，位置引用修改后文档的实际章节号（最多三级标题），每条两行以内：

```
问题1：方案中标注的"（根据项目特点删减）"请删除。
回复：已修改，已删除1.2.1节"国家及行业规范标准"两处表标题后的"（根据项目特点删减）"标注。

问题2：2.4.1节地下室防水概况特殊部位的章节细部做法图与后文不一致。
回复：已修改，2.4.1节"地下室防水概况"表中9处"详见4.7.x"引用已统一改为后文实际章节号（详见4.4节"防水重要节点处理"及其子节）。
```

### 触发场景示例

- "帮我改一下这个施工方案，意见是……"
- "桩基方案的规范 JGJ46-2005 废止了，更新一下"
- "按公司模板重新组织这份方案的章节结构"
- "方案里面积前后不一致，统一一下"
- "把方案里的旧项目名换成现在的项目名"

---

## 工作原理 / How It Works

### 四类修改原则

| 类型 | 处理方式 |
|---|---|
| 删除类 | 只删意见要求的内容，确需额外删除先征得确认 |
| 修改类 | 改标题/文本/表格，格式与原内容一致，只改目标内容 |
| 补充/新增类 | 按"数据来源优先级"取数，不编造；拿不准就问用户 |
| 调整类 | 重排章节前先给修改计划，重排后核查交叉引用与自动编号 |

### 数据来源优先级

```
同方案其他章节 → 同项目其他已编方案 → 计算书 → 规范原文（web 搜索确认） → 问用户
```

### 主动检查的冲突与残留

- 文档内部矛盾：两处数值/单位/日期/名称不一致
- 模板残留：旧地名、旧奖项（"中州杯"等）、旧机械、旧年份、其他项目名
- 废止规范：JGJ46-2005、GB/T50082-2009、GB1499.1-2008、JC239-2001 等

### 验证清单（输出前全部执行）

1. 重新解析输出文件，XML 良构
2. 逐条断言：每个修改点已生效、旧内容无残留
3. rId 完整性：无未定义引用、无指向 NULL 的悬空目标
4. python-docx 能正常打开
5. 包内 media 文件与修改前一致（图片未动）
6. 章节结构改动时：模拟渲染编号无"5.0.1"类异常，交叉引用同步修正

---

## 目录结构 / Layout

```
construction-scheme-modify/
├── SKILL.md                    # 技能主文件（工作流、原则、验证清单、回复格式）
├── references/
│   └── docx-editing.md         # OOXML 编辑技术手册（踩坑记录）
├── scripts/
│   └── docx_edit.py            # 可复用编辑工具（解析/替换/表格/编号模拟/验证）
├── evals/
│   └── evals.json              # 评估用例（提示词与断言，已匿名化）
├── README.md
├── LICENSE                     # MIT
└── package.json                # npm 包元数据
```

---

## 评估 / Evaluation

使用 [skill-creator](https://github.com/anthropics/skill-creator) 方法论评估：带技能 vs 无技能基线（baseline），各自独立修改真实 docx 方案文件，程序化断言评分。

| 迭代 | 带技能 | 基线 | 说明 |
|---|---|---|---|
| 迭代 1 | 100% (29/29) | 100% (29/29) | 初版验证，断言无区分度 |
| 迭代 2 | 100% (30/30) | 96% (29/30) | 新增"回复不解释技术原因"规则后产生真实区分度 |

迭代 2 的关键差异（意见 4：流程图图片不可编辑）：

- **带技能**：`回复：未修改，待提供新流程图后替换。`（技术原因只在过程沟通中告知）
- **基线**：在最终回复里解释"嵌入的Visio图片对象无法直接编辑…"

其他带技能优势：耗时更低（319s vs 360s）、方差更小（±54s vs ±143s）、回复章节定位更精确、主动报告执行中发现的新问题。

> 注：eval 输入为真实项目文档，未随仓库分发；用你自己的方案文件即可复现。

---

## 常见问题 / FAQ

**Q: 技能会修改我的原文件吗？**
不会。输出始终是新文件"原名（修改稿）.docx"，原文件保持不动。

**Q: 遇到 WPS 生成的文档打不开怎么办？**
技能默认用 zipfile+lxml 直接解析 XML，天然绕开 WPS 悬空关系问题；python-docx 打不开不影响修改。

**Q: 图片/流程图能改吗？**
不能。技能不编辑任何图片与嵌入对象（OLE）；如果意见要求改流程图，会在过程沟通中告知你需提供新图。

**Q: 回复里的章节号怎么来的？**
正文标题多为自动编号（Word 样式级编号），技能按文档顺序模拟渲染编号，与你在 Word 里看到的一致；目录若为过期快照，以正文为准，改完提示你"更新域"刷新目录。

**Q: 需要联网吗？**
修改本身不需要；仅当意见涉及更新废止规范版本号时，会经 web 搜索确认现行版本。

---

## 贡献 / Contributing

欢迎提交 Issue 与 PR：

1. Fork 本仓库
2. 新建分支：`git checkout -b feature/xxx`
3. 提交修改：`git commit -m 'feat: xxx'`
4. 推送分支：`git push origin feature/xxx`
5. 发起 Pull Request

### 开发流程

```bash
# 本地验证技能格式
python scripts/docx_edit.py dump <your-scheme.docx>

# 评估（需要 skill-creator 工作区）
python grade.py <eval-dir>
```

---

## 变更记录 / Changelog

### v1.0.0 (2026-08-19)

- 首个发布版本
- 核心工作流：通读 → 定位 → 汇报 → 修改 → 验证 → 回复
- 文件查看方法降级链（zipfile+lxml → python-docx → 文件头兜底）
- 回复只陈述结果，不解释技术原因
- WPS 悬空关系兼容处理
- 评估：迭代 2 带技能 100% vs 基线 96%

---

## 许可证 / License

MIT License. See [LICENSE](LICENSE).
