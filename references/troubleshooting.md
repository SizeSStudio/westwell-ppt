# Westwell PPT · 常见问题与修复

Step 7 迭代优化时按需读取。用完即释放。

---

## 内容 / 叙事类问题

### 标题是 topic label 不是论点

**症状**:"市场现状 / 产品架构 / 实施计划"这种目录式标题。

**修复**:把"这一页要读者相信什么"写成完整句。
参考 SKILL.md 里的"weak vs strong"对照表。

### 一页塞了多个论点

**症状**:body 5 条 bullet,每条讲不同的东西;或 two_col 左右讲两个不相关的主题。

**修复**:一页一个论点。拆成两页。
如果拆完后其中一页内容太薄,说明原来的论点太弱,合并到相邻页或删掉。

### 依赖页还没做就想做总结页

**症状**:生成第 2 页执行摘要时 Claude 发现前面分析都没写过。

**修复**:看 `dummy-pages-spec.md` 的依赖检查话术,先问用户走哪条路径,不要硬生成。

### 数据编了 / 不确定来源

**症状**:body 里出现具体数字,但 data-trail.md 里查不到出处。

**修复**:要么去补搜索 + 回填 data-trail,要么把数字降级为定性描述("已有规模商用")。
**不要让数据无源头地出现在 slide 上。**

---

## 视觉类问题

### 标题被居中了

**修复**:Westwell 视觉的第一条铁律是 title top-left。检查 builder 调用,确认没有手动加 align=center。

### 内容密度过高

**症状**:一页超过 60 字的 body,或 7 条以上的 bullets。

**修复**:
- 文字能给演讲者说的就不要写在 slide 上
- bullets 超 7 条 → 拆成两页 或 改用 three_col / stats
- body 太长 → 改用 text_image,把说明丢给演讲

### 分析型版式看起来太密 / 不像 westwell

**症状**:连续几页都是 `table` / `two_col`,整份 deck 像 McKinsey 咨询报告。

**修复**:每章最多 2–3 张分析型,中间用 `statement` / 大图 / 单 KPI slide 做视觉缓冲。见 `layouts-analytic.md` 的"与西井视觉风格的平衡"。

### 图表配色不对

**症状**:matplotlib 默认蓝橙配色。

**修复**:图表生成脚本里强制用 westwell 色板:
```python
NAVY  = "#1A2B6D"
TEAL  = "#00B3B0"
GRAY  = "#8892A8"
ALTROW = "#F3F5F8"
```
配色参考 `design-system.md` 的完整色板。

### 图表文字缩放后看不清

**修复**:matplotlib 里字号统一放大:
```python
plt.rcParams.update({
    'font.size': 14,
    'axes.titlesize': 16,
    'axes.labelsize': 14,
    'xtick.labelsize': 12,
    'ytick.labelsize': 12,
    'legend.fontsize': 12,
})
```

### 深色 slide 上的文字看不清

**修复**:dark=True 时所有文字 builder 已自动换成白,除非你手工传了 `color=`。检查自定义 textbox 的颜色参数。

---

## Builder / 环境类问题

### 封面出现"click to add text"

**症状**:`cover()` 生成后预览里还有占位符文字。

**修复**:`cover()` 内部已经调用 `_suppress_placeholders` + 自由 textbox,如果还有,可能是 `.potx` 的 layout 里加了新占位符。临时修复:在 slide 生成后手动调用 `_suppress_placeholders(slide)`。

### `agenda()` 报错

**原因**:传了混合格式(有的是 string,有的是 tuple)。

**修复**:统一传 `List[str]`(让 builder 自动编号),或统一传 `List[Tuple[num, title, subtitle]]`,不要混。

### 预览报错找不到 `fitz`

**修复**:`pip install pymupdf`。缺依赖,不是 builder 的 bug。

### soffice 不在 PATH

**修复**:`brew install --cask libreoffice`,或在脚本里写死路径:
```python
SOFFICE = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
```

### .potx → .pptx 转换失败

**症状**:`save()` 报错或产出空 pptx。

**修复**:
1. 确认 `.potx` 文件路径正确且可读
2. 手动测试:`soffice --headless --convert-to pptx <template.potx>`
3. 如果 soffice 自己也打不开 → 换用 `.pptx` 模板(先手动转一次,保存 `assets/PPTTemplate.pptx`,builder 直接读)

---

## 流程类问题

### Claude 一次性想做太多页

**症状**:在 Step 5 直接"我把所有 20 页都写好了"。

**修复**:提醒 Claude 走 `data-collection.md` 的 9 步循环,一页一页来,每页暂停等用户确认。

### 跨对话续写对不上号

**症状**:新对话里生成的新页风格和旧的不一致。

**修复**:
1. 让用户把 Dummy + 已生成的 `.pptx` 都发过来
2. Claude 先 `Read` Dummy,再用 `preview_pptx` 看几张旧 slide
3. 如果有 data-trail.md,也读一下,避免重复搜索
4. 确认风格 / 版式后再开始新一页

### 用户说"再漂亮点"但没给具体方向

**对策**:具体化。常见选项:
- "是指视觉上更简洁(减字 / 加大图)?"
- "还是叙事更紧凑(砍掉重复的证据)?"
- "还是想改配色 / 版式?"
问清楚再改,不要瞎发散。
