# 模板制作页设计（第一版）

用户上传一份自己的 PPT，在网页里把它做成一份模板：选出（或做出）封面、目录、内容页、结束页，
写上占位符，检查、试生成，最后保存为**只有自己能用**的模板。之后新建 PPT 时，Agent 就能用 `build_deck` 按它生成。

相关约定：

- [template-placeholders.md](template-placeholders.md)：页面组成和占位符的写法规则
- [template-keywords.md](template-keywords.md)：现有模板用到的全部关键字
- [page-tools.md](page-tools.md)：Agent 的工具在用户打开的页面里执行

> 状态：设计稿，还没有实现。写于 2026-10-19。

## 1. 目标和范围

**第一版要做到**

1. 上传一份 .pptx，在编辑器里打开。
2. 给页面指定角色：封面、目录、内容页、结束页各一页必选，章节页可选。
3. 上传的稿子缺某种角色时（很常见），可以**拿某一页当底，复制出一页做成这个角色**。
4. 指定角色后，自动写入这个角色的占位符：能对应上的文本框直接换成关键字，示例文字删掉，缺的关键字新加文本框。
5. AI 辅助：一键让 Agent 建议每页的角色和关键字对应关系，并直接改好，用户在画布上确认、调整。
6. 保存前检查，试生成一份示例稿看效果。
7. 保存为私有模板，生成缩略图和 `design_spec.md`；新建 PPT 时在模板列表里能选到它。

**第一版不做**

- 发布给所有用户（数据结构里预留，见 4.3）。
- 版式变体（同一角色多页不同版式，见 template-keywords.md 最后一节）。
- 可替换的图片占位（logo 换图）。
- 自定义关键字：只能从 template-keywords.md 的"基本关键字 + 通用字段"里选。

**拿什么来试**

engine 仓库根目录的 `元景行业智能体方案-20260713(2) - 副本.pptx`：8 页，17 MB，**全是内容页**，没有封面、目录和结束页。
正好用来走通"复制某页做成封面、目录、结束页"这条路。

## 2. 和现有"模板制作"页的关系

ai-ppt-view 已经有一个"模板制作"页（路由 `/template-authoring`，`src/view/templateAuthoring/`，菜单"模板制作"），
后端是 ai-ppt-server 的 `/aiPpt/templateAuthoring/*`（`service/system/sys_ppt_template_authoring.go`）。
它是 **SVG 版**的做法：上传 .pptx → 用 Python 把每页转成 SVG → 逐页编辑 SVG → 保存到 `ppt_template` / `ppt_template_page` 表。

产品已经改成真 PPTX，这套 SVG 做法**直接删掉不要了**（2026-10-19 定）：

- **路由和菜单沿用** `/template-authoring`、"模板制作"，页面换成本文的新实现。
- 删掉旧代码：
  - ai-ppt-view：`src/view/templateAuthoring/` 的旧实现、`src/api/pptTemplateAuthoring.ts`；
  - ai-ppt-server：`api/v1/system/sys_ppt_template_authoring.go`、`service/system/sys_ppt_template_authoring.go`、
    `model/system/request/ppt_template_authoring.go`、`model/example/ppt_template_manage.go`、`model/example/ppt_template_page.go`，
    `router/system/sys_ppt.go` 里的 `/aiPpt/templateAuthoring/*` 路由，以及 `source/system/api.go`、`source/system/casbin.go`、
    `initialize/ensure_doc_modules_casbin.go` 里对应的登记；还有它调用的 PPTX 转 SVG 的 Python 脚本（删之前确认没有别处在用）。
- 只有这个模块用 `ppt_template_category`、`ppt_template`、`ppt_template_version`、`ppt_template_page` 四张表
  （"模板分类与版本"接口 `templateCategoryVersion` 是 Word 文档模板，用的是 `doc_template*`，不受影响）。
  表在共用的 MySQL 里，代码不会自动建或删它们；**删表由你手动执行**，代码先删。
- 旧页面的分类第一版不做；私有模板数量少，按更新时间排列即可。

## 3. 整体结构

模板草稿就是**一个 PPT 项目**，只是标记为"模板草稿"。这样现有的机制全部复用：

| 需要的能力 | 复用什么 |
|---|---|
| 打开、编辑、保存、历史版本 | 项目的 deck：`GET/PUT aiPpt/plugin/deck`、`AutoSaver`、`DeckEditor` |
| 缩略图 | 项目 deck 的缩略图（渲染服务） |
| AI 辅助 | 宿主 Agent + 页面工具（`get_outline`、`apply_edits`、`render_slide`），在用户打开的页面里执行 |
| 读取图片 | `aiPpt/plugin/deck/asset` |

```
上传 .pptx ──▶ 新建项目（kind = template_draft，deck = 上传的文件）
                 │
                 ▼
     模板制作页：DeckEditor + 角色面板 + Agent 对话
       │  用户/Agent 改页面 → AutoSaver 存回项目 deck（带版本）
       │  检查：engine checkTemplate（页面里执行）
       │  试生成：在页面里另开一份副本跑 build_deck，不动草稿
       ▼
     保存为模板 ──▶ 服务端读草稿当前版本，再检查一次，写入模板目录（owner = 当前用户）
                      生成缩略图、保存 design_spec.md
                 ▼
     新建 PPT 时 list_templates 返回：公共模板 + 自己的模板
```

## 4. 后端（ai-ppt-server）

### 4.1 模板草稿 = 项目

`SysPptDraft` 加一列：

| 列 | 说明 |
|---|---|
| `kind` | `deck`（默认，普通稿子）或 `template_draft` |

- 模板制作页的草稿列表只列 `template_draft`。
- "历史会话"页（`/ppt-history`）列的是对话，不是项目。模板草稿里用了 AI 对话，也会多出一条对话。
  这些对话**保留在历史会话里，标成"模板制作"，点开回到模板制作页**（2026-10-19 定）：
  - 历史会话的每一条已经带着它的草稿 id（`draftId`，就是项目 id），现在点"编辑"会去 `/edit-ppt/<draftId>`；
  - 历史会话页同时取一次 `aiPpt/plugin/templateDraft/list`，`draftId` 属于模板草稿的对话加上"模板制作"标签，
    点开去 `/template-authoring/<draftId>?sessionId=<对话id>`，不去编辑页；
  - 这样宿主不需要知道什么是模板草稿，判断都在插件的后端和页面里。
- 草稿发布成模板后保留，下次可以继续改、再保存（成为模板的新版本）。

新接口（都需要登录，按当前用户过滤，路由/API/casbin 和现有 deck 接口一样登记）：

| 接口 | 说明 |
|---|---|
| `POST aiPpt/plugin/templateDraft`，请求体为 .pptx，`?name=` | 新建模板草稿；文件用 `inspectDeck` 校验，最大 64 MB。返回 `project_id` |
| `GET aiPpt/plugin/templateDraft/list` | 自己的模板草稿：id、名字、页数、缩略图、是否已保存为模板、更新时间 |
| `DELETE aiPpt/plugin/templateDraft?project_id=` | 删除草稿（已保存的模板不受影响） |
| `POST aiPpt/plugin/templateDraft/publish?project_id=&expected_revision=` | 保存为模板，见 4.2 |
| `GET aiPpt/plugin/template/mine` | 自己的模板列表 |
| `DELETE aiPpt/plugin/template?id=` | 删除自己的模板（已用它生成的稿子不受影响：项目里存的是生成后的 deck） |

### 4.2 保存为模板

`publish` 做的事：

1. 读草稿 deck 的**当前版本**；`expected_revision` 不一致返回 409（页面有没存的改动时先存）。
2. 服务端不能跑引擎的 `inferTemplate`（它是 TypeScript）。所以检查分两层：
   - 页面在调用前用引擎做完整检查（第 5 节），不通过不允许保存；
   - 服务端只做便宜的复查：解压 .pptx，用正则从每页文字里取 `{{关键字}}`，按[判断顺序](template-placeholders.md#引擎怎么判断页面角色)确认 4 种必选角色各有一页。
     这和 `template-keywords.md` 收集时用的方法相同。
3. 只保留模板用到的页：页面在保存前已经把多余的页删掉，服务端复查时页数必须是 4 或 5。
4. 写入模板目录（4.3），版本号 = 内容的哈希（和 `PublishLocalTemplates` 一样）。
5. 生成缩略图：调用渲染服务 `/thumbnail` 渲染第 1 页。
6. 保存页面一起提交的 `design_spec.md` 和元数据（名字、摘要、风格、明暗、关键词）。

### 4.3 私有模板放在哪

模板目录正在从本地文件夹迁到数据库 + 对象存储（`ppt_deck_templates` / `ppt_deck_template_versions`，
`object_templates.go`，这部分是别人还没提交的改动，**实现前要和做迁移的人对齐**）。私有模板直接放进这套目录：

| 表 | 新增的列 | 说明 |
|---|---|---|
| `ppt_deck_templates` | `owner_id`（0 = 公共） | 私有模板的主人 |

- 私有模板的 id 用 `u<用户ID>-<随机串>`，不会和公共模板重名；显示名放在 `label` 里，可以重复、可以改。
- 对象存储路径照常用 `TemplatePrefix(id, version)`。
- 以后"发布给所有人"：把 `owner_id` 改成 0（或复制一份公共模板），不需要改结构。

需要改的地方：

| 位置 | 改动 |
|---|---|
| `ListObjectTemplates` / `Templates()` | 加上用户参数：返回公共模板 + 这个用户的模板 |
| `objectTemplate(ctx, id, version)` | 加上用户参数：私有模板只有主人能取到 |
| `templateDeck(id)`（本地模式） | 同上；本地模式下私有模板放在 `ProjectRoot/templates/<id>/` |
| `list_templates` 工具 | 自己的模板标出 `"mine": true`，Agent 可以告诉用户"这是你的模板" |
| `skill:pptx_templates/<id>/design_spec.md` | 私有模板的说明用另一个前缀（如 `template:<id>/design_spec.md`），读取时检查主人 |
| `init_project`、`use_template` | 传入当前用户，取模板时检查主人 |

### 4.4 Agent 怎么知道在做模板

插件只有一个 skill（`skill:SKILL.md`），宿主不能加 PPT 相关的逻辑。所以：

- 现在的服务端工具（`init_project`、`list_templates`、`use_template`、`read_file`、`read_files`、`list_dir`、`commit_design_spec`）
  里没有一个能告诉 Agent 项目是什么类型。新增一个 `get_project` 工具，返回标题、`kind` 和 deck 的版本、页数
  （数据就是现有 `Runtime.View` 返回的那些，加上 `kind`）。
- `SKILL.md` 加一节"模板制作"：开始时先调 `get_project`，`kind` 是 `template_draft` 时按那一节做，
  不要用 `build_deck` 和 `use_template` 改这份稿子。
- 模板制作页发出的第一条消息也写明"这是一份模板草稿"，Agent 没调 `get_project` 也不会弄错。
- 那一节的内容就是 template-placeholders.md 的规则，加上第 6 节的工作步骤。

## 5. 引擎（pptx-editor-engine）

新增 `src/api/template-authoring.ts`，在页面和 Agent 工具里都能用：

### 5.1 `checkTemplate(outline)`

按 template-placeholders.md 的"保存前的检查"返回问题列表，每条带页码、对象 id 和说明，格式和 `checkLayout` 一样：

- 每种必选角色有且只有一页，章节页最多一页，一共 4 或 5 页；
- 每页有角色的必须关键字，角色标记没有放错页；
- 关键字都在纯文本框里（不在有填充或边框的形状里）；
- 关键字在清单里，用的是建议的写法（`COVER_TITLE` 提示改成 `TITLE`）；
- 目录页有"目录"标题在第一项之前，项编号从 1 连续，每项的对象在图层上连续；
- 内容页的内容区存在，而且至少放得下两块并排的内容；
- `inferTemplate` 认出的角色和每页的关键字一致。

### 5.2 `suggestPlaceholders(outline, slide, role)`

不用 AI 的默认建议，返回一批 `apply_edits` 指令：

- 字号最大、位置靠上的文本框 → 标题（封面 `TITLE`、内容页 `PAGE_TITLE`、章节页 `CHAPTER_TITLE`、结束页 `THANK_YOU`）；
- 紧挨着标题、字号次之的 → 副标题；
- 像日期的文字 → `DATE`；像页码的数字 → `PAGE_NUM`；
- 目录页：按位置找出重复排列的几组文字，依次写 `TOC_ITEM_n_TITLE` / `TOC_ITEM_n_DESC`；
- 内容页：标题下方剩下的区域画一个无填充的虚线框，中间放 `{{CONTENT_AREA}}`，原来区域里的文字、表格、图表、示例图片删掉；
- 其余示例文字删掉；有填充的形状里的文字拆出来（形状留着当底色，上面新放一个纯文本框）；
- 缺的必须关键字新加文本框，放在页面空白处，字号按这个模板里同类文字取。

用户点"按角色自动写入"时用它；AI 辅助时 Agent 也可以先拿它的结果再改。

### 5.3 `copySlideAsRole(slide, role)`

"拿某一页当底做成某角色"：返回 `duplicateSlide` + 删除页面上的文字、表格、图表、内容图片（背景、装饰形状、logo 保留）+ `suggestPlaceholders` 的指令。

> 判断"装饰"还是"内容"：图片按面积和位置（贴边的、很小的算装饰，页面中部的大图算内容）；形状没有文字的算装饰。
> 判断不准时用户在画布上删掉或撤销即可（整批是一步撤销）。

### 5.4 `templateDesignFacts(outline)`

从模板里取出写 `design_spec.md` 用的事实：页面尺寸（厘米）、主色（`accentOf`）和用到的颜色、各级文字的字体和字号、
边距、内容区大小、每页的明暗。AI 只写用途、风格这些描述，数字以这里为准。

### 5.5 新的页面工具 `check_template`

把 `checkTemplate` 包成一个工具，和 `check_layout` 并列，`executor: "page"`。Agent 改完占位符后调用它。
加到 `TOOLS` 后重新生成 `page_tools.json`。

## 6. 前端（ai-ppt-view）

### 6.1 页面

**入口**：`/template-authoring` 先显示"我的模板草稿"和"我的模板"两个列表，加一个"上传 PPT 制作模板"按钮。
点开一份草稿进入制作页 `/template-authoring/:projectId`。

**制作页布局**

```
┌──────────────┬──────────────────────────────┬────────────────────┐
│ 页面列表      │           画布                │  这一页            │
│ ▢1 [封面]  ✓ │      （DeckEditor，可直接编辑）  │  角色：[内容页 ▾]   │
│ ▢2 [目录]  ! │                              │  关键字：           │
│ ▢3 [内容]  ✓ │                              │   PAGE_TITLE ✓ 已绑定│
│ ▢4 未指定     │                              │   CONTENT_AREA ! 缺 │
│ ▢5 [结束]  ✓ │                              │   [绑定到选中的框]   │
│              │                              │   [新加文本框]      │
│ [+ 复制为…]   │                              │  [按角色自动写入]    │
│              │                              │  检查结果（3 条）    │
├──────────────┴──────────────────────────────┤                    │
│ Agent 对话：[让 AI 帮我标注] …                 │                    │
└─────────────────────────────────────────────┴────────────────────┘
顶部：草稿名 · 保存状态 · 历史版本 · [试生成] · [保存为模板]
```

- **页面列表**：每页显示角色标签和检查状态（✓ / ! 有问题 / 未指定）。角色就是从页面上的关键字算出来的（`inferTemplate`），
  不另外存，所以用户手动改了关键字，标签也跟着变。
  - "复制为…"：选一种角色，调用 `copySlideAsRole`。
  - 未指定角色的页在保存为模板时会被删掉，删之前提示用户。
- **画布**：现有的 `DeckEditor`，用户可以拖动、改字、删对象。
- **这一页**面板：
  - 角色下拉：改角色时先提示"会清掉这页的示例文字并写入占位符"，再调用 `suggestPlaceholders`；
  - 关键字列表：这个角色的必须和可选关键字，显示是否已经在页面上；"绑定到选中的框"把选中文本框的文字换成这个关键字；
  - 这一页的检查结果，点一条就在画布上选中对应的对象。
- **Agent 对话**：和编辑页同一个对话组件（`pptWorkflowStore`），页面工具在这个页面里执行。常用的一句做成按钮："帮我标注这份模板"。

### 6.2 试生成

不能直接在草稿上跑 `build_deck`（它会把整份稿子换掉）。做法：

1. 导出草稿当前的 .pptx（内存里）；
2. 在页面里另开一个不显示的 viewer 或 Web Worker 打开这份副本；
3. 用一份内置的示例内容（覆盖 template-placeholders.md 列的情况：带副标题的封面、5 项目录、章节页、两块并排的内容页、带图表的内容页、结束页）跑 `build_deck`，再跑 `check_layout`；
4. 在弹窗里显示生成后每页的缩略图和版面问题。关闭弹窗即丢弃副本。

### 6.3 保存为模板

1. 先跑 `checkTemplate`，有问题时列出来，不允许保存；
2. 弹窗填名字；摘要、风格、明暗、关键词和 `design_spec.md` 由 AI 根据 `templateDesignFacts` 和页面截图起草，用户可以改；
3. 删掉未指定角色的页（一步可撤销），等自动保存完成；
4. 调 `publish`，成功后提示"已保存为你的模板，新建 PPT 时可以选它"。

### 6.4 新建 PPT 页

新建页的模板列表来自 `/aiPpt/templates`、`/aiPpt/templateDetail`。这两个接口挂在**不需要登录**的路由上（`router/system/sys_ppt.go`），
不能返回私有模板。所以新建页另外调需要登录的 `aiPpt/plugin/template/mine`，把结果合并进列表：自己的模板排在前面，标"我的"。
Agent 那边不受影响：`list_templates` 工具带着用户身份，直接返回公共模板加自己的模板。

## 7. AI 辅助怎么工作

用户点"帮我标注这份模板"时，Agent 按 SKILL.md 的"模板制作"一节：

1. `get_outline` + 每页 `render_slide`，看每页像什么；
2. 告诉用户建议：哪页做封面、目录、内容页、结束页；缺的角色拿哪页当底复制；哪些页删掉。**等用户确认**；
3. 每页用 `apply_edits` 写占位符（可以参考 `suggestPlaceholders` 的做法）：沿用原来的文本框，删示例文字，内容页画内容区；
4. `check_template` + `render_slide` 检查，修到没有问题；
5. 告诉用户可以试生成、保存。

页面工具在用户打开的页面里执行，改动立刻显示在画布上，每一步都能撤销，也会存进历史版本。

## 8. 分步实现

| 步骤 | 仓库 | 内容 | 完成的标准 |
|---|---|---|---|
| 1 | pptx-editor-engine | `checkTemplate`、`suggestPlaceholders`、`copySlideAsRole`、`templateDesignFacts`、`check_template` 工具 | vitest：用元景那份 PPT 做出 4 页模板，`checkTemplate` 没有问题，`build_deck` 能生成 |
| 2 | ai-ppt-server | `kind` 列、模板草稿接口、`publish`、`owner_id` 和按用户过滤的模板目录 | Go 测试：别的用户看不到、用不了私有模板；复查拦得住缺角色的模板 |
| 3 | ai-ppt-view、ai-ppt-server | 草稿列表、制作页、试生成、保存为模板、新建页显示"我的"；删掉旧的 SVG 模板制作代码 | `npm run test:ppt-agent` 里加一条：上传 → 自动写入 → 保存 → 新建时能选到 |
| 4 | ai-ppt-server | SKILL.md"模板制作"一节、新的 `get_project` 工具 | 用真模型对元景那份 PPT 走一遍 AI 标注 |
| 5 | 全部 | 用元景那份 PPT 在真实页面里完整走一遍 | 做出的模板能生成一份 10 页左右的稿子，`check_layout` 没有严重问题 |

步骤 1 和 2 互不依赖，可以同时做；3 依赖 1 和 2；4 依赖 1。

## 9. 已定和待定的问题

已定（2026-10-19）：

- 模板先只给制作者自己用；以后再加发布给所有人（结构已预留）。
- 旧的 SVG 模板制作代码直接删掉，表由你手动删（见第 2 节）。
- 每个用户最多 50 个私有模板，每个最大 64 MB；草稿也算在 50 个里。超过时上传和保存都提示先删掉旧的。
- "历史会话"里模板制作产生的对话保留，标成"模板制作"，点开回到模板制作页（见 4.1）。
- 上传的 PPT 用了非开源字体时，**保存为模板前提醒用户**：列出这些字体和页面上实际显示用的替代字体，说明生成的文件里保留原字体名，
  在装了这些字体的电脑上打开效果更准；用户确认后继续保存。

待定：

1. 模板目录的数据库 + 对象存储迁移还没提交，私有模板要建在它上面。实现前要确认这套表结构不会再变。
