# 模板关键字清单

把 ai-ppt-server 现有模板里用到的 `{{关键字}}` 全部收集起来，供模板生成页选用，也为以后加新的关键字和版式留底。
占位符的写法规则见 [template-placeholders.md](template-placeholders.md)。

- 收集时间：2026-10-19
- 来源：`resource/ai_skills/ppt-master/templates/layouts/` 下 25 个 SVG 模板，`pptx_templates/` 下 25 个 PPTX 模板
- 共 **142 个**关键字（`TOC_ITEM_1_TITLE`、`TOC_ITEM_2_TITLE` 这类只差序号的，算作一个）
- PPTX 模板是从 SVG 模板转来的，两边用的关键字几乎一样

## 引擎怎么对待一个关键字

引擎（`src/api/template.ts`）把关键字分成四种，**任何关键字都能用**，不需要事先登记：

| 种类 | 关键字 | 引擎怎么处理 |
|---|---|---|
| 角色标记 | `TOC_ITEM_*`、`CONTENT_AREA`、`CHAPTER_*`、`SECTION_NAME`、`THANK_YOU`、`TITLE` | 决定这一页是什么角色，见 [判断顺序](template-placeholders.md#引擎怎么判断页面角色) |
| 自动填 | `PAGE_NUM`、`TOTAL_PAGES`、`CHAPTER_NUM`、`SECTION_NUM` | 引擎生成时自己填页码、总页数、章节号 |
| 不填 | `LOGO`、`LOGO_*` | 不作为字段；写着它的文本框在生成时被删掉 |
| 普通字段 | 其他所有关键字 | 成为这一页的一个字段，Agent 在 `build_deck` 里按字段名填写；不填的话所在文本框被删掉 |

字段名的规则：几个常用的有短名（`TITLE`、`PAGE_TITLE`、`CHAPTER_TITLE`、`THANK_YOU` → `title`；
`ENDING_SUBTITLE`、`CHAPTER_SUBTITLE` → `subtitle`；`CHAPTER_DESC` → `desc`；`CLOSING_MESSAGE` → `message`；
`CONTACT_INFO` → `contact`），其余按"下划线转驼峰"：`KEY_MESSAGE` → `keyMessage`，`CARD_1_TITLE` → `card1Title`。

所以加一个新的关键字，引擎不用改；要考虑的是 **Agent 能不能从名字看懂该填什么**——`get_template` 只给 Agent 字段名，不给说明。

## 分类整理

### 1. 通用字段（多个模板都在用，建议模板生成页直接提供）

| 页面 | 关键字（括号里是用到它的模板数） |
|---|---|
| 封面 | `TITLE`(21)、`SUBTITLE`(21)、`DATE`(19)、`AUTHOR`(16)、`INSTITUTION`(4)、`ADVISOR`(3)、`PRESENTER`(3)、`AUTHOR_EN`(2)、`DEPARTMENT`(2)、`ORGANIZATION`(2) |
| 目录 | `TOC_ITEM_{n}_TITLE`(21)、`TOC_ITEM_{n}_DESC`(15) |
| 章节页 | `CHAPTER_TITLE`(20)、`CHAPTER_DESC`(9)、`CHAPTER_SUBTITLE`(7)、`CHAPTER_TITLE_EN`(2) |
| 内容页 | `PAGE_TITLE`(21)、`CONTENT_AREA`(21)、`SOURCE`(9，数据来源)、`KEY_MESSAGE`(6，本页结论)、`SECTION_NAME`(6，页眉上的所属章节)、`ORG_SHORT`(2) |
| 结束页 | `THANK_YOU`(20)、`CONTACT_INFO`(20)、`ENDING_SUBTITLE`(15)、`COPYRIGHT`(11)、`THANK_YOU_EN`(7)、`CLOSING_MESSAGE`(4)、`CONTACT_EMAIL`(2)、`EMAIL`(2) |
| 任何页 | `PAGE_NUM`(22)、`CHAPTER_NUM`(21)、`SECTION_NUM`(3)、`TOTAL_PAGES`(2)、`LOGO`(4) |

内容页上常放 `CHAPTER_NUM`、`SECTION_NAME`，结束页上常放 `DATE`、`AUTHOR`：这些都没问题，
因为引擎先判断目录和内容页，再判断章节页（见判断顺序）。会出问题的只是**封面上放 `CHAPTER_*`**（封面会被当成章节页）、
以及**目录以外的页放 `TOC_ITEM_*`**。

### 2. 同一个意思、写法不一样（建议统一）

| 现有写法 | 建议统一成 | 说明 |
|---|---|---|
| `COVER_TITLE` | `TITLE` | `中国联通_产互` 的封面只有 `COVER_TITLE`，引擎认不出它是封面 |
| `COVER_SUBTITLE` | `SUBTITLE` | 同上 |
| `ENDING_TITLE` | `THANK_YOU` | `中国联通_产互` 的结束页认不出 |
| `END_SUBTITLE` | `ENDING_SUBTITLE` | |
| `ENDING_MESSAGE` | `CLOSING_MESSAGE` | |
| `EMAIL`、`CONTACT_EMAIL` | `CONTACT_EMAIL` | |
| `STAT_LABEL_1`、`STAT_1_LABEL` | `STAT_1_LABEL` | 序号放在前面，和 `TOC_ITEM_1_TITLE`、`CARD_1_TITLE` 一致 |
| `STAT_NUMBER_1`、`STAT_1` | `STAT_1_VALUE` | |
| `SPEAKER_NAME`、`PRESENTER` | `PRESENTER` | |
| `ORGANIZATION`、`INSTITUTION` | 保留两个 | 前者是机关单位，后者是学校、医院，模板语气不同 |

### 3. 图片位置（引擎目前不填）

| 关键字 | 模板数 | 说明 |
|---|---|---|
| `LOGO`、`LOGO_HEADER`、`LOGO_LARGE` | 4、3、3 | 放 logo 的位置。4 个 PPTX 模板里还留着写有 `{{LOGO}}` 的文本框，生成时这个框会被删掉；`LOGO_HEADER`、`LOGO_LARGE` 只在 SVG 里 |
| `COVER_BG_IMAGE` | 4 | 封面背景图，只在 SVG 里；PPTX 模板直接把背景图做进页面 |

以后要让 Agent 或用户换图（比如换成上传的 logo），可以约定"图片名字叫 `{{LOGO}}` 的图片可替换"，再在 `build_deck` 里加一个图片字段。现在没有这个功能。

### 4. 固定版式字段（只在一两个模板里，跟着这页的设计走）

这些是"这一页就是这样排的"：设计师把几张卡片、几个数字摆好，Agent 只往里填字。

| 一组关键字 | 所在页面 | 模板 | 这一页长什么样 |
|---|---|---|---|
| `CARD_{1-3}_TITLE`、`_SUBTITLE`、`_TAG`、`_LINE1`、`_LINE2`、`_NOTE1`、`_NOTE2` | 内容 | anthropic | 三张卡片并排，每张有标签、标题、两行要点、两条注释 |
| `STAT_{n}`、`STAT_{n}_LABEL`、`STAT_LABEL_{n}`、`STAT_NUMBER_{n}`、`STATS_AREA` | 目录 | pixel_retro、中国电建、government_red、中汽研_常规、google_style | 目录旁边几个大数字 + 说明 |
| `GOAL_{1-4}`、`GOAL_{n}_DESC`、`KEY_FOCUS` | 目录 | psychology_attachment | 四个学习目标 + 本次重点 |
| `TOC_ITEM_{n}_SUB_{n}`、`TOC_SUMMARY`、`TOC_TITLE` | 目录 | china_unicom_uniai | 目录项下面再分子项 |
| `FEATURE_{1-4}`、`ROW_{1-3}_A/B`、`COMPARE_A/B`、`SUMMARY_{n}_TITLE`、`SUMMARY_1_LINE_{1-3}`、`SUMMARY_1_STAT` | 结束 | pixel_retro | 结束页上的总结卡片、对比表 |
| `THANKS_PERSON_{1-3}`、`THANKS_REASON_{1-3}` | 结束 | google_style | 致谢名单 |
| `QUOTE`、`QUOTE_AUTHOR`、`TAG_{1-4}` | 封面、章节 | psychology_attachment | 引言和标签 |
| `COVER_QUOTE`、`BRAND_LABEL`、`PAGE_LABEL` | 封面、内容 | anthropic | 品牌标签、引言 |
| `TITLE_LINE2`、`PAGE_TITLE_LINE2`、`TITLE_EN`、`PAGE_TITLE_EN`、`CHAPTER_EN` | 多种 | mckinsey 等 | 标题分两行、中英双语标题 |
| `CONTACT_NAME`、`CONTACT_TITLE`、`CONTACT_PHONE`、`CONTACT_URL`、`CONTACT_LINE_2`、`WEBSITE`、`CTA_TEXT` | 结束 | 多种 | 拆开的联系方式 |
| `VERSION`、`PROJECT_ID`、`PROJECT_CODE`、`PAGE_SUBTITLE`、`NOTE`、`ENDING_LABEL`、`ENDING_FOOTNOTE` | 多种 | 各一个 | 其他 |

引擎现在已经能填这些：它们都是普通字段，比如 anthropic 的内容页在 `get_template` 里就有 `card1Title` 等字段。
但有两个限制：

- **每种角色只能有一页**。所以一个模板不能同时有"自由排版的内容页"和"三张卡片的内容页"。
- **字段名就是全部说明**。`card1Line1` 还能猜，`rowOneA`、`compareA` 这类 Agent 很难知道该写什么。

## 以后支持"某页就要这样的布局"：版式变体（建议，未实现）

你说的"我就想要某页这样的布局"，就是上面第 4 类的做法，再往前走一步：**同一种角色可以有几页不同的版式，每页有自己的字段**。

建议的约定：

1. 模板里每种角色仍有一页"默认版式"；另外可以加几页同角色的"变体"。
2. 变体页用一个标记关键字写名字，例如在页面角落放一个文本框 `{{LAYOUT:三卡片}}`（生成时删除）。
3. 变体页的字段照常用关键字写，序号统一放在前面：`CARD_1_TITLE`、`CARD_1_TEXT`、`STAT_1_VALUE`、`STAT_1_LABEL`。
4. 每个字段可以在 `design_spec.md`（或模板旁的一个小 JSON）里写一句说明，`get_template` 把说明一起交给 Agent。
5. Agent 在 `build_deck` 里写 `{page: "content", layout: "三卡片", card1Title: …}`；不写 `layout` 时用默认版式。

引擎要改的地方：`inferTemplate` 按"角色 + 版式名"存页面，`describeTemplate` 列出每个版式的字段和说明，
`compileDeck` 按 `layout` 选页。模板生成页则让用户多选几页标成"内容页·变体"并起名字。
第一版模板生成页可以先不做变体，但关键字命名按上面第 3 条来，以后不用改模板。

## 引擎现在认不出的模板页

收集时发现两个 PPTX 模板有问题，`build_deck` 用不到这些页（`中国联通_产互` 实际只认出一页内容页，`china_unicom_uniai` 只认出封面和目录）：

| 模板 | 页面 | 页面上的关键字 | 原因 |
|---|---|---|---|
| `china_unicom_uniai` | 第 3 页（内容页） | `PAGE_TITLE`、`PAGE_NUM` | 没有 `CONTENT_AREA` |
| `china_unicom_uniai` | 第 4 页（结束页） | `ENDING_LABEL`、`ENDING_MESSAGE`、`ENDING_SUBTITLE`、`ENDING_FOOTNOTE` | 没有 `THANK_YOU` |
| `中国联通_产互` | 第 1 页（封面） | `COVER_TITLE`、`COVER_SUBTITLE`、`DATE` | 没有 `TITLE` |
| `中国联通_产互` | 第 5 页（结束页） | `ENDING_TITLE` | 没有 `THANK_YOU` |
| `中国联通_产互` | 第 2、3 页 | 和第 4 页一样（`PAGE_TITLE`、`CONTENT_AREA`…） | 这两页本该是目录和章节页，转成 PPTX 时变成了内容页的副本；引擎只用第一页内容页 |

改法：按上面"建议统一"把关键字改成引擎认识的那个，内容页加上 `CONTENT_AREA`；`中国联通_产互` 的目录和章节页要从 SVG 重新转。

## 附：全部关键字

按用到它的模板数从多到少排列。"出现的页面"按 SVG 文件名（`01_cover` 等）和 PPTX 页面上引擎判断的角色合并。

| 关键字 | 字段名 | SVG 模板数 | PPTX 模板数 | 出现的页面 | 例子 |
|---|---|---|---|---|---|
| `{{PAGE_NUM}}` | （自动填） | 22 | 22 | 内容、目录、结束、章节、封面、未识别 | academic_defense、ai_ops、anthropic |
| `{{CHAPTER_NUM}}` | （自动填） | 21 | 21 | 章节、内容、目录 | academic_defense、ai_ops、anthropic |
| `{{CONTENT_AREA}}` | （内容区标记） | 21 | 21 | 内容、章节、目录 | academic_defense、ai_ops、anthropic |
| `{{PAGE_TITLE}}` | `title` | 21 | 21 | 内容、目录、章节、未识别 | academic_defense、ai_ops、anthropic |
| `{{SUBTITLE}}` | `subtitle` | 21 | 21 | 封面 | academic_defense、ai_ops、anthropic |
| `{{TITLE}}` | `title` | 21 | 21 | 封面、章节 | academic_defense、ai_ops、anthropic |
| `{{TOC_ITEM_{n}_TITLE}}` | （目录项） | 21 | 21 | 目录 | academic_defense、ai_ops、anthropic |
| `{{CHAPTER_TITLE}}` | `title` | 20 | 20 | 章节 | academic_defense、ai_ops、anthropic |
| `{{CONTACT_INFO}}` | `contact` | 20 | 20 | 结束 | academic_defense、ai_ops、anthropic |
| `{{THANK_YOU}}` | `title` | 20 | 20 | 结束 | academic_defense、ai_ops、anthropic |
| `{{DATE}}` | `date` | 19 | 19 | 封面、结束、目录、未识别 | academic_defense、ai_ops、anthropic |
| `{{AUTHOR}}` | `author` | 16 | 16 | 封面、结束 | academic_defense、ai_ops、china_unicom_uniai |
| `{{ENDING_SUBTITLE}}` | `subtitle` | 15 | 15 | 结束、未识别 | academic_defense、ai_ops、anthropic |
| `{{TOC_ITEM_{n}_DESC}}` | （目录项） | 15 | 15 | 目录 | academic_defense、anthropic、exhibit |
| `{{COPYRIGHT}}` | `copyright` | 11 | 11 | 结束 | academic_defense、ai_ops、anthropic |
| `{{CHAPTER_DESC}}` | `desc` | 9 | 9 | 章节 | academic_defense、ai_ops、anthropic |
| `{{SOURCE}}` | `source` | 9 | 9 | 内容、封面 | academic_defense、ai_ops、anthropic |
| `{{CHAPTER_SUBTITLE}}` | `subtitle` | 7 | 7 | 章节 | government_blue、government_red、中国电建_常规 |
| `{{THANK_YOU_EN}}` | `thankYouEn` | 7 | 7 | 结束 | government_blue、government_red、psychology_attachment |
| `{{KEY_MESSAGE}}` | `keyMessage` | 6 | 6 | 内容 | academic_defense、exhibit、mckinsey |
| `{{SECTION_NAME}}` | `sectionName` | 6 | 6 | 内容 | academic_defense、ai_ops、mckinsey |
| `{{CLOSING_MESSAGE}}` | `message` | 4 | 4 | 结束 | psychology_attachment、smart_red、中国电建_现代 |
| `{{COVER_BG_IMAGE}}` | `coverBgImage` | 4 | 0 | 封面 | psychology_attachment、中汽研_商务、中汽研_常规 |
| `{{INSTITUTION}}` | `institution` | 4 | 4 | 封面、结束、章节、目录 | academic_defense、medical_university、招商银行 |
| `{{LOGO}}` | （不填） | 4 | 4 | 封面、内容、目录、结束 | academic_defense、medical_university、招商银行 |
| `{{ADVISOR}}` | `advisor` | 3 | 3 | 封面 | academic_defense、medical_university、重庆大学 |
| `{{LOGO_HEADER}}` | （不填） | 3 | 0 | 内容、目录 | 中汽研_商务、中汽研_常规、中汽研_现代 |
| `{{LOGO_LARGE}}` | （不填） | 3 | 0 | 封面、结束、内容、目录 | 中汽研_商务、中汽研_常规、中汽研_现代 |
| `{{PRESENTER}}` | `presenter` | 3 | 3 | 封面 | government_blue、中国电建_常规、中国电建_现代 |
| `{{SECTION_NUM}}` | （自动填） | 3 | 3 | 内容 | academic_defense、mckinsey、medical_university |
| `{{STAT_1}}` | `stat1` | 3 | 3 | 目录 | pixel_retro、中国电建_常规、中国电建_现代 |
| `{{STAT_2}}` | `stat2` | 3 | 3 | 目录 | pixel_retro、中国电建_常规、中国电建_现代 |
| `{{STAT_3}}` | `stat3` | 3 | 3 | 目录 | pixel_retro、中国电建_常规、中国电建_现代 |
| `{{AUTHOR_EN}}` | `authorEn` | 2 | 2 | 封面、结束 | government_red、中汽研_常规 |
| `{{CHAPTER_TITLE_EN}}` | `chapterTitleEn` | 2 | 2 | 章节 | google_style、pixel_retro |
| `{{CONTACT_EMAIL}}` | `contactEmail` | 2 | 2 | 结束 | anthropic、mckinsey |
| `{{DEPARTMENT}}` | `department` | 2 | 2 | 封面、结束 | medical_university、招商银行 |
| `{{EMAIL}}` | `email` | 2 | 2 | 结束 | academic_defense、重庆大学 |
| `{{ORGANIZATION}}` | `organization` | 2 | 2 | 封面、章节、结束 | government_blue、government_red |
| `{{ORG_SHORT}}` | `orgShort` | 2 | 2 | 内容、封面 | government_blue、government_red |
| `{{STAT_1_LABEL}}` | `stat1Label` | 2 | 2 | 目录 | 中国电建_常规、中国电建_现代 |
| `{{STAT_2_LABEL}}` | `stat2Label` | 2 | 2 | 目录 | 中国电建_常规、中国电建_现代 |
| `{{STAT_3_LABEL}}` | `stat3Label` | 2 | 2 | 目录 | 中国电建_常规、中国电建_现代 |
| `{{STAT_LABEL_1}}` | `statLabel1` | 2 | 2 | 目录 | government_red、中汽研_常规 |
| `{{STAT_LABEL_2}}` | `statLabel2` | 2 | 2 | 目录 | government_red、中汽研_常规 |
| `{{STAT_NUMBER_1}}` | `statNumber1` | 2 | 2 | 目录 | government_red、中汽研_常规 |
| `{{STAT_NUMBER_2}}` | `statNumber2` | 2 | 2 | 目录 | government_red、中汽研_常规 |
| `{{TOTAL_PAGES}}` | （自动填） | 2 | 2 | 内容、章节、目录、封面、结束 | anthropic、pixel_retro |
| `{{BRAND_LABEL}}` | `brandLabel` | 1 | 1 | 封面 | anthropic |
| `{{CARD_1_LINE1}}` | `card1Line1` | 1 | 1 | 内容 | anthropic |
| `{{CARD_1_LINE2}}` | `card1Line2` | 1 | 1 | 内容 | anthropic |
| `{{CARD_1_NOTE1}}` | `card1Note1` | 1 | 1 | 内容 | anthropic |
| `{{CARD_1_NOTE2}}` | `card1Note2` | 1 | 1 | 内容 | anthropic |
| `{{CARD_1_SUBTITLE}}` | `card1Subtitle` | 1 | 1 | 内容 | anthropic |
| `{{CARD_1_TAG}}` | `card1Tag` | 1 | 1 | 内容 | anthropic |
| `{{CARD_1_TITLE}}` | `card1Title` | 1 | 1 | 内容 | anthropic |
| `{{CARD_2_LINE1}}` | `card2Line1` | 1 | 1 | 内容 | anthropic |
| `{{CARD_2_LINE2}}` | `card2Line2` | 1 | 1 | 内容 | anthropic |
| `{{CARD_2_NOTE1}}` | `card2Note1` | 1 | 1 | 内容 | anthropic |
| `{{CARD_2_NOTE2}}` | `card2Note2` | 1 | 1 | 内容 | anthropic |
| `{{CARD_2_SUBTITLE}}` | `card2Subtitle` | 1 | 1 | 内容 | anthropic |
| `{{CARD_2_TAG}}` | `card2Tag` | 1 | 1 | 内容 | anthropic |
| `{{CARD_2_TITLE}}` | `card2Title` | 1 | 1 | 内容 | anthropic |
| `{{CARD_3_LINE1}}` | `card3Line1` | 1 | 1 | 内容 | anthropic |
| `{{CARD_3_LINE2}}` | `card3Line2` | 1 | 1 | 内容 | anthropic |
| `{{CARD_3_NOTE1}}` | `card3Note1` | 1 | 1 | 内容 | anthropic |
| `{{CARD_3_NOTE2}}` | `card3Note2` | 1 | 1 | 内容 | anthropic |
| `{{CARD_3_SUBTITLE}}` | `card3Subtitle` | 1 | 1 | 内容 | anthropic |
| `{{CARD_3_TAG}}` | `card3Tag` | 1 | 1 | 内容 | anthropic |
| `{{CARD_3_TITLE}}` | `card3Title` | 1 | 1 | 内容 | anthropic |
| `{{CHAPTER_EN}}` | `chapterEn` | 1 | 1 | 章节 | psychology_attachment |
| `{{COMPARE_A}}` | `compareA` | 1 | 1 | 结束 | pixel_retro |
| `{{COMPARE_B}}` | `compareB` | 1 | 1 | 结束 | pixel_retro |
| `{{CONTACT_LINE_2}}` | `contactLine2` | 1 | 1 | 结束 | psychology_attachment |
| `{{CONTACT_NAME}}` | `contactName` | 1 | 1 | 结束 | exhibit |
| `{{CONTACT_PHONE}}` | `contactPhone` | 1 | 1 | 结束 | mckinsey |
| `{{CONTACT_TITLE}}` | `contactTitle` | 1 | 1 | 结束 | mckinsey |
| `{{CONTACT_URL}}` | `contactUrl` | 1 | 1 | 结束 | anthropic |
| `{{COVER_QUOTE}}` | `coverQuote` | 1 | 1 | 封面 | anthropic |
| `{{COVER_SUBTITLE}}` | `coverSubtitle` | 1 | 1 | 封面、未识别 | 中国联通_产互 |
| `{{COVER_TITLE}}` | `coverTitle` | 1 | 1 | 封面、未识别 | 中国联通_产互 |
| `{{CTA_TEXT}}` | `ctaText` | 1 | 1 | 结束 | pixel_retro |
| `{{ENDING_FOOTNOTE}}` | `endingFootnote` | 1 | 1 | 结束、未识别 | china_unicom_uniai |
| `{{ENDING_LABEL}}` | `endingLabel` | 1 | 1 | 结束、未识别 | china_unicom_uniai |
| `{{ENDING_MESSAGE}}` | `endingMessage` | 1 | 1 | 结束、未识别 | china_unicom_uniai |
| `{{ENDING_TITLE}}` | `endingTitle` | 1 | 1 | 结束、未识别 | 中国联通_产互 |
| `{{END_SUBTITLE}}` | `endSubtitle` | 1 | 1 | 结束 | pixel_retro |
| `{{FEATURE_1}}` | `feature1` | 1 | 1 | 结束 | pixel_retro |
| `{{FEATURE_2}}` | `feature2` | 1 | 1 | 结束 | pixel_retro |
| `{{FEATURE_3}}` | `feature3` | 1 | 1 | 结束 | pixel_retro |
| `{{FEATURE_4}}` | `feature4` | 1 | 1 | 结束 | pixel_retro |
| `{{GOAL_1}}` | `goal1` | 1 | 1 | 目录 | psychology_attachment |
| `{{GOAL_1_DESC}}` | `goal1Desc` | 1 | 1 | 目录 | psychology_attachment |
| `{{GOAL_2}}` | `goal2` | 1 | 1 | 目录 | psychology_attachment |
| `{{GOAL_2_DESC}}` | `goal2Desc` | 1 | 1 | 目录 | psychology_attachment |
| `{{GOAL_3}}` | `goal3` | 1 | 1 | 目录 | psychology_attachment |
| `{{GOAL_3_DESC}}` | `goal3Desc` | 1 | 1 | 目录 | psychology_attachment |
| `{{GOAL_4}}` | `goal4` | 1 | 1 | 目录 | psychology_attachment |
| `{{GOAL_4_DESC}}` | `goal4Desc` | 1 | 1 | 目录 | psychology_attachment |
| `{{KEY_FOCUS}}` | `keyFocus` | 1 | 1 | 目录 | mckinsey |
| `{{NOTE}}` | `note` | 1 | 1 | 内容 | mckinsey |
| `{{PAGE_LABEL}}` | `pageLabel` | 1 | 1 | 内容 | anthropic |
| `{{PAGE_SUBTITLE}}` | `pageSubtitle` | 1 | 1 | 内容 | pixel_retro |
| `{{PAGE_TITLE_EN}}` | `pageTitleEn` | 1 | 1 | 内容、目录 | pixel_retro |
| `{{PAGE_TITLE_LINE2}}` | `pageTitleLine2` | 1 | 1 | 内容 | anthropic |
| `{{PROJECT_CODE}}` | `projectCode` | 1 | 1 | 封面、章节 | mckinsey |
| `{{PROJECT_ID}}` | `projectId` | 1 | 1 | 封面 | exhibit |
| `{{QUOTE}}` | `quote` | 1 | 1 | 封面、章节 | psychology_attachment |
| `{{QUOTE_AUTHOR}}` | `quoteAuthor` | 1 | 1 | 封面、章节 | psychology_attachment |
| `{{ROW_1_A}}` | `row1A` | 1 | 1 | 结束 | pixel_retro |
| `{{ROW_1_B}}` | `row1B` | 1 | 1 | 结束 | pixel_retro |
| `{{ROW_2_A}}` | `row2A` | 1 | 1 | 结束 | pixel_retro |
| `{{ROW_2_B}}` | `row2B` | 1 | 1 | 结束 | pixel_retro |
| `{{ROW_3_A}}` | `row3A` | 1 | 1 | 结束 | pixel_retro |
| `{{ROW_3_B}}` | `row3B` | 1 | 1 | 结束 | pixel_retro |
| `{{SPEAKER_NAME}}` | `speakerName` | 1 | 1 | 封面 | google_style |
| `{{SPEAKER_TITLE}}` | `speakerTitle` | 1 | 1 | 封面 | google_style |
| `{{STATS_AREA}}` | `statsArea` | 1 | 1 | 目录 | google_style |
| `{{SUMMARY_1_LINE_1}}` | `summary1Line1` | 1 | 1 | 结束 | pixel_retro |
| `{{SUMMARY_1_LINE_2}}` | `summary1Line2` | 1 | 1 | 结束 | pixel_retro |
| `{{SUMMARY_1_LINE_3}}` | `summary1Line3` | 1 | 1 | 结束 | pixel_retro |
| `{{SUMMARY_1_STAT}}` | `summary1Stat` | 1 | 1 | 结束 | pixel_retro |
| `{{SUMMARY_1_TITLE}}` | `summary1Title` | 1 | 1 | 结束 | pixel_retro |
| `{{SUMMARY_2_TITLE}}` | `summary2Title` | 1 | 1 | 结束 | pixel_retro |
| `{{SUMMARY_3_TITLE}}` | `summary3Title` | 1 | 1 | 结束 | pixel_retro |
| `{{TAG_1}}` | `tag1` | 1 | 1 | 封面、章节 | psychology_attachment |
| `{{TAG_2}}` | `tag2` | 1 | 1 | 封面、章节 | psychology_attachment |
| `{{TAG_3}}` | `tag3` | 1 | 1 | 封面、章节 | psychology_attachment |
| `{{TAG_4}}` | `tag4` | 1 | 1 | 封面、章节 | psychology_attachment |
| `{{THANKS_PERSON_1}}` | `thanksPerson1` | 1 | 1 | 结束 | google_style |
| `{{THANKS_PERSON_2}}` | `thanksPerson2` | 1 | 1 | 结束 | google_style |
| `{{THANKS_PERSON_3}}` | `thanksPerson3` | 1 | 1 | 结束 | google_style |
| `{{THANKS_REASON_1}}` | `thanksReason1` | 1 | 1 | 结束 | google_style |
| `{{THANKS_REASON_2}}` | `thanksReason2` | 1 | 1 | 结束 | google_style |
| `{{THANKS_REASON_3}}` | `thanksReason3` | 1 | 1 | 结束 | google_style |
| `{{TITLE_EN}}` | `titleEn` | 1 | 1 | 内容、封面 | psychology_attachment |
| `{{TITLE_LINE2}}` | `titleLine2` | 1 | 1 | 封面 | mckinsey |
| `{{TOC_ITEM_{n}_SUB_{n}}}` | （目录项） | 1 | 1 | 目录 | china_unicom_uniai |
| `{{TOC_SUMMARY}}` | `tocSummary` | 1 | 1 | 目录 | ai_ops |
| `{{TOC_TITLE}}` | `tocTitle` | 1 | 1 | 目录 | 重庆大学 |
| `{{VERSION}}` | `version` | 1 | 1 | 封面 | pixel_retro |
| `{{WEBSITE}}` | `website` | 1 | 1 | 结束 | 科技蓝商务 |
