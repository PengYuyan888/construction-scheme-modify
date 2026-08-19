# construction-scheme-modify

修改施工方案/专项方案（Word .docx）并编写"修改意见回复"的 Agent 技能（Skill）。

Revise construction schemes / special schemes (Word .docx) per review comments and produce a formal modification reply. An agent skill for DSH / Claude Code style skill runtimes.

## 功能特性 / Features

- **按意见修改方案**：删除 / 修改 / 补充 / 调整四类修改意见逐条落位，输出"原名（修改稿）.docx"，原文件不动。
- **修改意见回复**：按固定格式生成"问题N：… / 回复：已修改，详xxxxx"，每条两行以内，只陈述处理结果。
- **自动冲突检查**：识别文档内部矛盾（数值/单位/日期/名称不一致）、模板残留（旧地名、旧奖项、旧项目名）、废止规范（JGJ46-2005 等），列出问题清单待用户决策。
- **规范版本校核**：更新废止规范时经 web 搜索确认现行版本号，不凭记忆编造。
- **图片零改动**：不编辑任何图片/嵌入对象（Visio 流程图等），做不到的事在过程沟通中告知用户。
- **文件查看降级链**：zipfile+lxml 默认解析 → python-docx 备选 → 查文件头兜底，不卡住。
- **WPS 兼容**：处理 WPS 生成的悬空关系（`Target="../NULL"`）等 python-docx 打开失败的文档。

## 安装 / Installation

1. 将 `construction-scheme-modify/` 目录（或 `.skill` 压缩包解压）放入技能目录，例如：
   ```
   ~/.agents/skills/construction-scheme-modify/
   ```
2. 重启会话或刷新技能列表后，技能自动生效。
3. 依赖：Python 3 + `lxml`（解析 docx 时使用）；`python-docx` 仅用于最终验证（可选）。

## 使用方法 / Usage

直接以自然语言提出修改意见即可触发，例如：

> 项目里有一个"××项目地下防水施工方案.docx"文件，以下是修改意见：
> 1、方案中标注的"（根据项目特点删减）"请删除。
> 2、2.4.1节地下室防水概况特殊部位的章节细部做法图与后文不一致。
> 3、2.1节地下室面积41.3㎡与后文3.2.2节施工部署中卸粮坑地下室40㎡数值不统一，请统一为41.3㎡。
> 4、4.2节施工流程体现的是2道防水卷材，图纸为涂料+卷材，请按图纸做法修改。
> 请根据意见修改方案，并生成修改意见回复。

技能会：通读方案 → 逐条定位 → 汇报修改计划与问题清单（等你确认）→ 执行修改 → 验证 → 输出修改稿与意见回复。

## 目录结构 / Layout

```
construction-scheme-modify/
├── SKILL.md                    # 技能主文件（工作流、原则、验证清单、回复格式）
├── references/
│   └── docx-editing.md         # OOXML 编辑技术手册（踩坑记录）
├── scripts/
│   └── docx_edit.py            # 可复用编辑工具（解析/替换/表格/编号模拟/验证）
└── evals/
    └── evals.json              # 评估用例（提示词与断言，已匿名化）
```

## 评估 / Evaluation

使用 [skill-creator](https://github.com/anthropics/skill-creator) 方法评估（带技能 vs 无技能基线，真实 docx 方案文件）：

| 迭代 | 带技能 | 基线 | 说明 |
|---|---|---|---|
| 迭代 1 | 100% (29/29) | 100% (29/29) | 初版验证，断言无区分度 |
| 迭代 2 | 100% (30/30) | 96% (29/30) | 新增"回复不解释技术原因"规则后产生区分度 |

（eval 输入为真实项目文档，未随仓库分发；用你自己的方案文件即可复现。）

## 许可证 / License

MIT License. See [LICENSE](LICENSE).
