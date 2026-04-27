# Composition Layouts (战略叙事组合版式)

10 种专门为**战略/董事会/咨询备忘录**叙事沉淀的组合版式,从一份咨询公司级别的
"西井战略思考 · 内部咨询版" 参考稿抽象而来。全部复用现有 Westwell 色板
(navy/teal/light-gray)和 `custom slide1-{light,dark}` 母版,所以可以和
`layouts-guide.md` / `layouts-analytic.md` 里的版式自由混搭。

---

## Editorial Framing(所有版式共享的 4 个可选参数)

所有版式方法(基础叙事版式 + 分析型版式 + 本文组合版式)都接受 4 个**可选**
参数。这是"咨询公司级别"材料的签名动作:

| 参数 | 作用 | 使用频率 |
|------|------|---------|
| `eyebrow='...'`   | 标题上方的 mono caps 小眉标(语义锚点) | **高**(几乎每页) |
| `subtitle='...'`  | 标题下的 14pt 灰色解释段(论点→论证) | 中(~60% 页面) |
| `footnote='...'`  | 底部 italic 灰色 so-what 金句 | 中(~30% 页面) |
| `notes='...'`     | 演讲者备注(不渲染到 slide,在演讲视图显示) | 低但非常有用 |

```python
ppt.pyramid(
    title='底盘、制高点、北极星 — 三层战略的分工',
    eyebrow='战略分层',               # mono caps 小眉标
    subtitle='战略不是并列罗列方向,而是明确分层关系:车守底盘,ReeWell '
             '争中期制高点,AI Operator 指向长期北极星。',
    tiers=[...],
    footnote='全球复制不是独立主线,而是标准层成立之后自然放大的结果。',
    notes='讲稿提示:强调"三层分工"vs"并列罗列" 的区别。管理层最容易误判的点。',
)
```

**Eyebrow 是你第一个应该加的**。它让一份 20 页的材料读起来像一份精心编织的
咨询报告,而不是一堆独立 slide 的拼盘。

---

## Bottom_callout — 战略备忘录的杀手锏

`two_col` 和 `three_col` 额外支持 `bottom_callout` 参数 — 在页面底部渲染
一个**横跨全幅**的强视觉 callout 框,承载这一页的终极 so-what。

```python
bottom_callout = {
    'label': 'BOTTOM LINE',   # mono caps 眉标(或 'FORMULA' / '关键判断')
    'text':  '车要继续做强,但公司不能止于卖车;项目要继续拿,但公司不能止于做项目。',
    'dark':  True,            # True=深蓝底白字(戏剧);False=浅灰+navy 左条(温和)
}
```

**两种典型用法**:

- **BOTTOM LINE**(dark=True):管理层建议页底部收束,一页自带结论。
- **FORMULA**(dark=False):三栏 + 底部等式(如 "Ainery = AI × 能源维度")。

**架构优势**:省掉"bottom line 独立成页"的稀释 — 一页到位,节奏紧。

---

## Editorial card body — 编号列表 vs 段落(两种内部布局)

每个 two_col / three_col 的 column body 可以是 **str 或 list**:

- **str** → 渲染为段落(适合叙述性说明)
- **list** → 渲染为**编号列表**(01/02/03 青色 mono + 每项一行)

列表风格搭配 `left_eyebrow` / `right_eyebrow`(two_col)或 column dict
里的 `eyebrow` 字段(three_col)使用,效果就是 `control_matrix` 左侧
MUST CONTROL 栏那种 editorial 卡片结构:

```
┌──────────────────────────────────┐
│ MUST CONTROL                     │  ← 小眉标(mono caps, teal)
│ 必须自控的能力                    │  ← 大标题(22pt bold)
│ ──────                           │  ← 青色 teal 短线
│ 01  场景运营与调度智能            │
│ 02  对客户场景的语义理解          │  ← 编号列表
│ 03  物理执行闭环                  │
│ 04  能源协同与补能组织            │
│ 05  关键执行层适配                │
└──────────────────────────────────┘
```

**使用方式**:

```python
ppt.two_col(
    ...,
    left_eyebrow='MUST ALIGN',        # 小眉标
    left_head='统一的三件事',         # 大标题
    left_body=[                        # list → 编号列表
        '**主线上移** — 不是卖更多车',
        '**ReeWell 平台化** — 按标准化组织',
        '**复制逻辑** — 先站住标准',
    ],
    right_eyebrow='DO NOT MISJUDGE',
    right_head='不要误判的三件事',
    right_body=[
        '咨询 ≠ 成熟打法',
        'AI Operator ≠ 近两年主战略',
        '探索 ≠ 主航道',
    ],
    contrast=True, emphasis='left',   # 左栏深蓝强调
)
```

**何时用 editorial list vs 长段落**:
- list → 多条并列的短论点、原则、建议、clauses(3–5 条)
- str  → 单一连贯的叙述、解释段落、因果论证

这是咨询备忘录的**签名卡片样式** — 比"长段落"清爽,读者扫一眼就能数清"有几件事"。

---

## 卡片对比度(`contrast` / `highlight_idx`)— 让某一栏突出

默认卡片是"全浅灰+teal 顶条"的均匀样式 — 适合平行并列。但当一栏承载整页
的"重点 / 结论 / 选择"时,均匀卡片视觉权重不够。参考 `control_matrix`
左右反差(深 navy vs 浅 lgray)的效果,**`two_col` 和 `three_col` 现在都
支持对比卡片**:

**two_col 的 `contrast=True`**:指定一栏深 navy(白字),另一栏浅 lgray(navy 字)。
```python
ppt.two_col(
    ...,
    contrast=True,
    emphasis='right',   # 'left' 或 'right' — 哪一栏深蓝抓眼
)
```

典型用法:
- **PATH A vs PATH B** / **过去 vs 现在** / **错的做法 vs 对的做法** → 强调对的那栏
- 一边陈述事实(浅色),一边给出判断(深色)→ 强调判断栏

**three_col 的 `highlight_idx=N`**:指定第 N 栏(0/1/2)为深 navy,其余仍为浅 lgray。
```python
ppt.three_col(
    ...,
    highlight_idx=1,    # 中间栏深蓝(像 Ainery 页的"成为标准层一部分")
)
```

典型用法:
- **并列三项,其中一项是"终局" / "关键"**(例如 Ainery 三栏中"成为标准层一部分")
- 三个阶段,最终阶段是目标(结合 FORMULA callout 效果最好)

**何时用对比卡片 vs 均匀卡片**:
- 均匀(默认)→ 三件事真正并列,没有谁更重要
- 对比(contrast / highlight_idx)→ 页面有"重点"需要视觉锚定

---

**与其他两份版式文档的分工**

| 文件 | 覆盖什么 | 典型场景 |
|------|---------|---------|
| `layouts-guide.md`      | 基础叙事版式 — 封面、章节、目录、bullets、stats、text_image、two_col、three_col、table、image、end | 产品方案、项目汇报、市场分析 |
| `layouts-analytic.md`   | 分析型版式 — 单图表、2×2 矩阵、瀑布图、时间轴等(多数需 PNG 预生成) | 数据论证、逻辑推理、趋势展示 |
| **`layouts-composition.md`** (本文) | 战略叙事组合版式 — 金字塔、价值链、控制矩阵、反向清单、Before/After、价值阶梯、Big Number、Quote、Number List | 战略备忘录、董事会材料、IPO 内部讨论、高阶提报 |
| `visual-grammar.md`     | 2页样张机制、节奏表、Westwell-compatible style directions | 5页以上 deck 开始批量生成前 |

`layouts-guide.md` now also includes visual grammar layouts:
`hero`, `big_numbers`, `image_grid`, `pipeline`, `rowlines`,
`lead_image`, and `quote_editorial`. Use these when the deck needs stronger
rhythm or a Guizang/Huashu-inspired editorial feel, then return to the
composition layouts here for strategy-specific argument structures.

---

## 适用场景决策树

```
需要展示"分层战略"(底盘 / 制高点 / 北极星)?             → 1. pyramid()
需要展示"从 A 到 Z 的升级路径"(同类多阶段)?              → 2. value_chain()  /  6. value_ladder()
  ├─ 只展示序列 + 说明                                     → value_chain()
  └─ 需要标注"已建立 / 建设中 / 待布局"的进度状态           → value_ladder()
需要划清"能力自控 vs 合作获取"的边界?                      → 3. control_matrix()
需要明确"哪些方向 ≠ 主战略"(反向表达)?                    → 4. not_list()
需要展示"客户痛点 / 行业状态 从过去迁移到未来"?            → 5. before_after()
需要一个单一关键数字作为全页焦点?                          → 7. big_number()
需要一句话作为战略宣言 / 管理层 bottom-line?               → 8. quote()
需要 3–5 条有序建议 / 举措 / 结论?                         → 9. number_list()
```

---

## 1. `pyramid(title, tiers, caption='', dark=False)`

**何时用**:三层战略分层 — 底盘 / 制高点 / 北极星 式结构,或任何"基础 → 中期 → 长期"
的递进分层关系。

**视觉**:3 个水平条,**从下到上由宽到窄**。底层最深蓝(foundation),中间 navy,
顶层亮蓝(`C_DARK`)。每层左侧有青色细竖条作为视觉锚点。标签约定:底=TIER 1,
顶=TIER 3(符合中文"第一层 = 最基础"的直觉)。

**入参**

```python
tiers = [
    # 顺序:底层(foundation) → 中层 → 顶层(north star)
    {"label": "当前底盘",   "en": "车 · 系统 · 项目 · 全链路交付",
     "sub": "现金流 · 场景入口 · 数据闭环"},
    {"label": "中期制高点", "en": "ReeWell · 调度标准 · 数据接口",
     "sub": "工作流收敛 + 能源维度"},
    {"label": "长期北极星", "en": "AI Operator · 按结果收费",
     "sub": "运营层替代"},
]
```

**内容规则**

- 必须 3 层。2 层用 `two_col`,4 层以上改用 `bullets` 或 `number_list`
- `label` ≤ 6 字(中文),`sub` ≤ 20 字,`en` ≤ 30 字(西文 / 专有名词为主)
- `caption` 是底部一行引言,用于收束或加一句金句 — 可省

---

## 2. `value_chain(title, steps, highlight_last=True, dark=False)`

**何时用**:展示一条**价值/工作流链**上的顺序步骤 — "我们做什么 → 客户怎么受益",
或商业模式的 N 个阶段。

**视觉**:3–5 个等宽竖栏,栏间用细竖线隔开(首栏深蓝线,其余青色)。每栏顶部
"STEP 0N" mono 青色,标题 navy 粗体,body 灰色。最后一栏可选高亮(淡灰底)
暗示"这是我们要去的方向"。

**入参**

```python
steps = [
    {"title": "卖车 / 卖项目",    "body": "执行层,现金流底盘"},
    {"title": "争调度 + 标准层", "body": "中期制高点,组织位置上移"},
    {"title": "进入运营层",       "body": "长期方向,绑定深度提升"},
    {"title": "按结果收费",       "body": "北极星,价值上限打开"},
]
```

**内容规则**

- 栏数 3–5 最佳。每栏 title ≤ 10 字,body ≤ 25 字
- 标题用名词短语,不写完整句子(栏窄会挤)
- 和 `value_ladder()` 的区别:本方法强调**序列**,不表达"当前进度"

---

## 3. `control_matrix(title, must_control, can_partner, principle='', dark=False)`

**何时用**:能力边界讨论 — "哪些能力必须自己做,哪些可以合作"。战略备忘录里用
来明确组织优先级。

**视觉**:左右 1.3 : 1 分栏。左栏深 navy 底,`MUST CONTROL` 青色 mono 眉标 +
"必须自控的能力"粗体标题 + 编号列表(01/02/03 青色)。右栏浅灰底,`CAN PARTNER`
+ "可以合作获取" + 青色方块 dot 列表。右下带分割线 + 斜体 `principle` 原则收束。

**入参**

```python
must_control = [
    "场景运营与调度智能 — 最核心的 AI 能力",
    "对客户场景的语义理解与接口组织能力",
    "物理执行闭环与调度闭环的耦合能力",
    # 3–5 条最佳
]
can_partner = [
    "部分通用大模型能力",
    "部分机器人与具身智能底层能力",
    # 2–4 条最佳
]
principle = "原则:凡决定能否进入标准层的能力必须自控;通用化程度高的可合作。"
```

**内容规则**

- 左右条数不必相等(左一般多些,突显"必须自控"的重)
- 每条 ≤ 30 字,否则换行影响对齐
- `principle` 是点睛一句 — 缺了也行,但加上才有"原则先行"的咨询味

---

## 4. `not_list(title, items, dark=False)`

**何时用**:反向表达 — **"什么 ≠ 主战略 + 正确定位是什么"**。比正向列举"主战略
有哪些"更有说服力,因为否定+修正是强认知动作。

**视觉**:逐行条目。每行三栏 — `01` 大灰数字 + 左栏 "what"(navy 粗体) + 右栏
`PROPER ROLE` 青色 mono 眉标 + 灰色 body。行间浅灰分隔线。

**入参**

```python
items = [
    {"what": "咨询能力本身",
     "correct": "是放大器,是前端入口,不是独立主线。"},
    {"what": "企业内部 AI 化升级",
     "correct": "是内部放大器,不能替代客户价值主线。"},
    # 2–7 条
]
```

**内容规则**

- `what` ≤ 15 字(短 label 式),`correct` ≤ 40 字(带原因/修正)
- 条数上限 7,超了拆两页
- 和 `bullets` 的区别:bullets 是正向并列,not_list 是否定 + 修正的**对仗结构**

---

## 5. `before_after(title, before, after, dark=False)`

**何时用**:展示**从一种状态迁移到另一种状态** — 客户痛点迁移、行业范式切换、
旧打法 → 新打法对比。

**视觉**:左右 1 : 1 分栏中间夹 `→` 青色大箭头。左"FROM · 过去"栏灰淡调;
右"TO · 未来"栏浅灰背景 + navy 粗左竖条 + `TO · 未来` 青色眉标,视觉上明显"
亮过"左栏,暗示方向。

**入参**

```python
before = {
    "title": "缺车 · 缺自动化设备",
    "body": "过去十年,客户核心诉求是通过采购自动化执行体,解决局部作业。",
}
after = {
    "title": "缺运营掌控力 + 能源协同",
    "body": "今天,客户要从局部自动化走向全场协同;能源从配套变成关键运营变量。",
}
```

**内容规则**

- 两边 `title` ≤ 15 字,`body` ≤ 50 字
- 左右信息量尽量对称,避免一边空一边挤
- 不要用 before_after 表达"对比两个方案"(那是 `two_col` 的事);它专门表达**时间迁移**

---

## 6. `value_ladder(title, stages, caption='', dark=False)`

**何时用**:展示一条**带阶段进度状态**的升级路径 — "第一阶段已建立,第二阶段
建设中,第三阶段早期探索,第四阶段待布局"。比 `value_chain` 多了"当前我们在
哪儿"的信息。

**视觉**:3–4 个等宽栏。每栏 `STAGE 0N` 青色 mono + 标题 navy + tag 灰色 +
进度条(青色填充,深 navy 填满表示"已建立")+ 状态 label。栏间 `›` 青灰
chevron 分隔符。

**入参**

```python
stages = [
    {"title": "卖车 / 卖项目",    "tag": "执行层",       "body": "底盘已站稳,现金流与客户入口都从这层出",       "progress": 100, "state": "已建立"},
    {"title": "争调度 + 标准层", "tag": "中期制高点",   "body": "ReeWell 从配套系统升级为全场调度大脑,先占写规则位置", "progress": 40,  "state": "建设中"},
    {"title": "进入运营层",       "tag": "长期方向",     "body": "嵌入客户运营流程,从系统供方到运营合伙人",             "progress": 10,  "state": "早期探索"},
    {"title": "按结果收费",       "tag": "北极星",       "body": "按可量化的物理操作结果计价,形成新的计量单位",           "progress": 0,   "state": "待布局"},
]
```

**内容规则**

- 3–4 个阶段最佳。再多视觉负担重
- `progress` 值决定填充颜色:100=navy(已建立),1–99=teal(进行中),0=空(待布局)
- `state` ≤ 6 字,和 `progress` 数值语义一致
- **`body` (可选) — 一行描述这个阶段"站住"意味着什么**。不写的话栏目只有
  title+tag+进度条,内容量少时会在中间留大块空白。填上 body 可以让栏内容
  均匀铺满 `ch`,页面密度显著改善。12–30 字为宜,超过会被 13pt 自动换行挤压
- 选用 `value_chain` 还是 `value_ladder`?看有没有"进度"维度 — 有就用 ladder

**版面自适应**:方法内部已把进度条锚定到 `ct + ch - 0.60`(而非固定
`ct + 1.38`),所以加了 subtitle/footnote 导致 `ch` 变大时,STAGE 标签在顶
部、进度条在底部,body 在中间呼吸 —— 不会再出现底部大块留白

---

## 7. `big_number(title, number, unit='', label='', body='', dark=False)`

**何时用**:一个关键数据作为全页焦点 — 渗透率、增长率、ARR、节省时间等。
替代"在 bullets 里塞一个加粗数字"的稀释写法。

**视觉**:左半幅巨大数字(navy,根据字符长度自动 70–150pt),右上角可选小号
青色 unit。右半幅 `LABEL` 青色 mono 眉标 + 洞察句(19pt,粗体 navy),顶部
一条青色短线做锚点。

**入参**

```python
number = "42%"        # 主数字
unit   = ""           # 可选:"亿" / "min" / "×" 等
label  = "KEY INSIGHT"  # mono 眉标,全大写英文
body   = "头部枢纽已进入**规模商用**,中小港口是下一波增量。"  # 支持 **bold**
```

**内容规则**

- `number` 保持简洁:"42%" / "1.2亿" / "<1s" / "83秒" 都行;太长 (>10字) 会自动缩到 70pt
- `body` ≤ 60 字。这是"so what" 的一句话洞察,不是数据解释
- 和 `stats()` 的区别:stats 是 2–4 个并列 KPI,big_number 是**单一焦点**

---

## 8. `quote(text, attribution='', title='', dark=True)`

**何时用**:一句宣言 / 金句 / 管理层 bottom-line。替代"在普通 slide 上加一
段斜体"的稀释写法,专门给关键句子一页的分量。

**视觉**:左上角巨大青色 `"` 引号 + 大号 navy/white 粗体引言 + 底部 "— 
`attribution`" 青色 mono。`dark=True`(默认)时为深蓝底白字,戏剧化; `dark=False` 时可做章节中段的回响 callout。

**入参**

```python
text        = "车要继续做强,但公司不能止于卖车;\n项目要继续拿,但公司不能止于做项目。"
attribution = "BOTTOM LINE · 给管理层的建议"
title       = ""       # 通常不写 title,让引言占满
dark        = True     # 深色版更震撼,浅色版作中段回响
```

**内容规则**

- `text` 最多 3 行(用 `\n` 分),每行 ≤ 25 字
- `attribution` ≤ 30 字(英文或中英混)
- 一份 deck 里用 quote 的页数 ≤ 2,多了稀释戏剧感

---

## 9. `number_list(title, items, dark=False)`

**何时用**:有序列表 — 3–5 条有编号的建议 / 行动项 / 结论。和 `bullets()`
是姊妹版式:bullets 是并列点(青色小方块),number_list 是**有序** / **有
优先级**(大号青色 01/02/03)。

**视觉**:每项左侧 `01` 大号青色 mono 数字 + 右侧标题(navy 粗体 17pt)+ 
body(灰色 14pt)。

**入参**

```python
items = [
    {"title": "主线不是卖更多车,而是上移到标准层和运营层",
     "body": "车必须继续做强,因为它是底盘、入口和承接层。"},
    {"title": "ReeWell 必须按标准化平台能力来组织",
     "body": "不再仅作为项目配套系统。"},
    # 3–5 条
]
```

**内容规则**

- 3–5 条最佳。2 条用 `two_col`,6+ 条拆两页
- `title` ≤ 25 字(一句完整论点),`body` ≤ 40 字(说明/原因/动作)
- 选 `bullets` 还是 `number_list`?有序 / 有优先级 / 要"第一、第二、第三"读感 → number_list

---

## 10. `step_grid(title, steps, highlight_last=True, dark=False)`

**何时用**:展示一条**协同体系**里的 3–6 个组成要素 — 比 bullets 更有
"系统感"、比 value_chain 更强调"并列而非递进"。Claude Design 参考稿的
p10 中期制高点页用的就是这个版式(5 个卡片方格,最后一个 ReeWell 格子高亮)。

**视觉**:3-6 张等宽卡片,浅灰底 + 青色顶部条 + STEP 01 mono 眉标 + 
20pt navy 粗体标题 + 14pt 灰色 body。最后一张如启用 `highlight_last`,
变为**深 navy 底 + 白字**(视觉焦点)。

**入参**

```python
steps = [
    {"title": "任务调度", "body": "车辆任务调度控制系统的标准版本"},
    {"title": "数据接口", "body": "数据接口与场景语义的定义权"},
    {"title": "工作流",   "body": "围绕客户工作流的收敛与标准化"},
    {"title": "能源协同", "body": "任务、设备、补能、能耗纳入同一调度框架"},
    {"title": "ReeWell",  "body": "从配套系统 → 全场运营标准层产品"},  # 高亮
]
```

**内容规则**

- 3–6 张卡片最佳(超过 6 个用 bullets + subtitle 说明)
- 每张 title ≤ 8 字(名词短语),body ≤ 25 字
- `highlight_last=True` 时最后一张应该是**整个体系的"终局"**(如 ReeWell 之
  于调度标准)
- 选 `step_grid` 还是 `value_chain`?
  - **协同关系**(5 个并列的组成要素)→ step_grid
  - **序列关系**(STEP 01 → STEP 04 递进)→ value_chain

**和 bullets 的区别**:bullets 是"5 个点",step_grid 是"5 个方格共同构成一个
协同体系"。视觉权重差 5×。

---

## 混搭策略

一份 20 页的战略备忘录里怎么分配?

- **每章 1 个**组合版式当作该章的"结构性锚点"(如金字塔、控制矩阵、价值阶梯)
- **章末 1 个** `quote()` 或 `number_list()` 收束,给读者一个明确的带走项
- **基础叙事版式** (bullets / text_image / stats) 仍然是主力,组合版式用来**放大
  关键节点**,不要铺满
- **组合版式 ≤ 40%** 的总页数,否则信息密度过载,读者记不住

**反模式**:

- 连续两页都是 control_matrix / pyramid / value_ladder — 视觉重复,失去新鲜感
- 把 quote 用在非关键位置 — 引号是"重兵器",每出现一次都应让读者多停留 10 秒
- not_list 的 `correct` 写成口号而非修正 — 失去"否定 + 修正"对仗的力量

---

## 和 `layouts-guide.md` / `layouts-analytic.md` 的交叉引用

| 你的需求                                     | 首选                  | 备选                |
|----------------------------------------------|----------------------|--------------------|
| 分层战略                                      | `pyramid` (composition) | — |
| 多阶段升级                                    | `value_chain` / `value_ladder` (composition) | `three_col` (guide) |
| 两个状态对比(方案 A vs B)                     | `two_col` (guide)     | `table` (guide)    |
| 两个状态迁移(过去 → 未来)                     | `before_after` (composition) | — |
| 单一焦点数字                                  | `big_number` (composition) | `stats` (guide,2–4 并列) |
| 2×2 象限矩阵                                  | `layouts-analytic` §3 (PNG + image) | — |
| 必控 vs 可合作 的定性边界                      | `control_matrix` (composition) | `two_col` (guide, 弱化版) |
| 章末结论 / 战略宣言                            | `quote` (composition) | `statement` (guide) |
| 3–5 条有序建议                                 | `number_list` (composition) | `bullets` (guide, 无序) |
| 正向陈述 "主战略有 X / Y / Z"                  | `number_list` (composition) / `three_col` (guide) | — |
| 反向陈述 "X / Y / Z 不是主战略 + 正确定位"     | `not_list` (composition) | — |
