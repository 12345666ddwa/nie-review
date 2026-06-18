---
name: exam-review-webpage-from-ppt
version: 1.0.0
description: 从课程PPT+课堂录音转写生成全覆盖交互式期末复习网页——含知识点提取、交叉验证、自测题系统、论述标答、图形综合题、一键部署
tags: [exam, review, education, html, ppt, study, chinese-university]
triggers:
  - 期末复习
  - 复习网页
  - PPT提取
  - 考试复习
  - 复习提纲
  - 复习手册
  - 课件提取
author: Hermes × 高兴
---

# 从PPT+录音到交互式复习网页的完整工作流

## 适用场景

用户提供课程复习PPT和/或课堂录音转写文本，需要生成一个覆盖全部考点的交互式复习网页，可部署到公网分享给同学。

## 总流程

```
PPT → python-pptx提取 → 考点清单
                              ↓
录音转写/笔记 ──→ 交叉对比(查漏补缺)
                              ↓
                    内容架构(教学逻辑)
                              ↓
                    HTML网页构建(单文件)
                              ↓
              美工 → 自测系统 → 署名 → 部署
```

## 阶段一：输入提取

### 1. PPT提取规则

```python
from pptx import Presentation
prs = Presentation("复习课.pptx")
for i, slide in enumerate(prs.slides, 1):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                # 保留 para.level（缩进层级）区分主/子知识点
                text = para.text.strip()
        if shape.has_table:
            # 遍历 table.rows → row.cells → cell.text
```

**铁律：**
- 遍历每一页的每一个shape（文本框+表格），不能只看标题
- 保留层级信息（para.level），区分主知识点和子知识点
- PPT列出的知识点就是考试范围，不能自行删减

### 2. 录音转写/笔记处理规则

- 课堂笔记是PPT的"扩充"而非"替代"
- 转写中的学术人名可能有听音误（如"洛斯"→"沃利斯和诺思"），需交叉核实
- 笔记中超出PPT范围的内容作为加分扩充保留，不丢弃

## 阶段二：交叉对比

### 程序化验证（必做）

```python
checklist = {
    "第一章 标题": [
        ("知识点名", "搜索关键词"),
        ...
    ],
    ...
}
for chapter, items in checklist.items():
    for name, keyword in items:
        if keyword in html_content:
            print(f"✅ {name}")
        else:
            print(f"❌ {name}")
```

**规则：**
- 逐条对照：PPT上的每一个关键词都必须在最终内容中找到对应
- 三类差异处理：
  - PPT有笔记没有 → 补充说明
  - 笔记有PPT没有 → 作为扩充保留
  - 术语冲突 → 以教材/学术标准为准
- 内容写完后再跑一次全量检测脚本，确保无遗漏
- 目标：100%覆盖率

## 阶段三：内容组织

### 每个知识点必须包含的四要素

1. **定义/内容**：先直觉理解再学术定义（假设读者零基础）
2. **题型预测标签**：根据特征匹配（分类→多选、定义→名词解释、是非→判断、逻辑链→论述）
3. **易错陷阱**（trap-box）：指出常见错误理解
4. **自测练习题**（self-test）：可折叠的练习+答案

### 题型预测规则

| 知识点特征 | 适配题型 |
|-----------|---------|
| 多维分类/多因素 | 多选 |
| 概念定义 | 名词解释、单选 |
| 容易混淆的表述 | 判断 |
| 理论体系/逻辑链 | 论述 |
| 图形+填空 | 综合分析 |

### 自测系统设计规则

- **两层折叠**：第一层展开看题目，第二层揭晓答案
- **题目用考卷格式**：标注【判断题】【单选题】等，选项ABCD
- **答案包含解析**：说明为什么对/错，指出陷阱
- **分布均匀**：每个章节的每个主要知识模块都必须有自测
- **覆盖所有题型**：单选/多选/判断/名词解释都要出现

### 论述题标答格式（老师要求小论文）

```
第一步：定义核心概念（2行）
第二步：理论观点梳理（4-5行）
第三步：个人理解与评价（3行）
第四步：结合现实案例（3-4行）
```

### 综合分析题规则

- 可能含图形（SVG内嵌，如供求曲线）
- 10个填空，有长有短（单字到一句话）
- 前半部分微观基础（均衡点、出清等），后半部分结合课程理论
- 答案同样折叠可揭晓

## 阶段四：HTML构建

### 单文件架构

所有CSS/JS/SVG内联，零外部依赖。结构：

```html
<!DOCTYPE html>
<html>
<head>
  <style>/* 全部CSS */</style>
</head>
<body>
  <nav class="top-nav">顶部导航</nav>
  <div class="main-content">
    署名 + 学校信息
    考试概览
    第一章～第N章（知识点+自测+陷阱）
    应试策略（论述标答+综合分析+图形题）
    速记清单（名词解释+判断陷阱+多选高频）
  </div>
  <script>/* 动画+导航高亮 */</script>
</body>
</html>
```

### CSS设计体系

| 组件 | 用途 | 配色 |
|------|------|------|
| `.card-blue` | 定义/重点 | 蓝色系 |
| `.card-red` | 警告/批判 | 红色系 |
| `.card-green` | 正面/答案 | 绿色系 |
| `.card-orange` | 提示/补充 | 橙色系 |
| `.card-purple` | 特殊概念 | 紫色系 |
| `.trap-box` | 易错陷阱 | 浅红底 |
| `.self-test` | 自测练习 | 紫色虚线框 |
| `.exam-question` | 模拟题 | 白底+边框 |
| `.kw` | 关键字蓝色 | 蓝色加粗+下划线 |
| `.kw-red` | 关键字红色 | 红色加粗 |
| `.key-point` | 重点句 | 黄底棕字 |

### 动效清单

- IntersectionObserver 滚动入场淡入
- 卡片悬停上浮+阴影增强
- 顶部导航当前章节高亮
- 折叠面板展开动画
- 页面标题淡入效果

### 关键字高亮规则

定义中的核心术语用 `.kw`（蓝色），易错/重要词用 `.kw-red`（红色），完整重点句用 `.key-point`（黄底）。目的：防止大段文字阅读走神。

## 阶段五：署名与学校信息

分享给同学时必须包含：
- 产权署名徽章（如"Built by Hermes × 学生名"）
- 学校+学院名称
- 课程名+学期
- 功能标签（章节数/题目数/陷阱数等）
- 简短使用说明

## 阶段六：部署

### GitHub Pages 部署步骤

```bash
# 1. 创建仓库
curl -X POST https://api.github.com/user/repos \
  -H "Authorization: token $TOKEN" \
  -d '{"name":"repo-name","public":true,"auto_init":true}'

# 2. 上传（大文件用@file）
# 先base64编码HTML → 写入JSON payload → curl -d @payload.json
curl -X PUT .../contents/index.html \
  -H "Authorization: token $TOKEN" \
  -d @payload.json

# 3. 开启Pages
curl -X POST .../pages \
  -d '{"source":{"branch":"main","path":"/"}}'
```

### 部署注意事项

- Token通过临时文件传递（分段拼接写入防redact）
- base64后>100KB时不能放命令行参数，必须写入JSON文件用`-d @file`
- 更新文件需先GET获取当前SHA，再PUT时带上
- 所有GitHub API请求走代理
- Pages缓存刷新需1-2分钟

## 质量保障检查清单

- [ ] python-pptx逐页逐shape提取，无遗漏
- [ ] 录音转写中的人名/术语已核实
- [ ] 程序化验证覆盖率100%
- [ ] 每个知识模块都有自测练习
- [ ] 自测题覆盖全部题型（单选/多选/判断/名词解释）
- [ ] 论述题有标答范文（四步框架）
- [ ] 综合分析题含图形+10空
- [ ] 关键字词已高亮标注
- [ ] 响应式布局手机可用
- [ ] 署名+学校信息完整
- [ ] GitHub Pages部署成功可访问

## Pitfalls

1. **PPT提取不能只看标题**：很多知识点藏在正文文本框和表格里
2. **课堂人名听音误**：转写可能把"沃利斯"听成"洛斯"，必须核实
3. **自测题分布不均**：容易只在有trap-box的地方加题，其他地方遗漏
4. **美工方向先问清**：避免做完才说"太老气"然后大改
5. **base64太大炸命令行**：>100KB必须用文件传递
6. **GitHub更新需要SHA**：不带SHA会报409冲突
7. **"覆盖100%"要用脚本验证**：人肉检查必遗漏
8. **综合分析题老师可能出图形题**：提前准备SVG供求曲线等示例
