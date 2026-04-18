---
name: image-first-5w-flashcard-generator
description: 输入主题 → 自动搜索图片 → AI识别内容 → 生成5W问题 → 输出Word图卡（图文完全匹配）
---

# 5W提问法图卡生成工具

## 核心原则
**先找图 → AI识别 → 根据实际内容设计问题 → 生成Word**。不要先写问题再配图，图片内容不可控。

## 输入主题，自动完成全流程

### 命令行用法

```bash
python generate_5w.py "公交车"        # 主题自动搜索图片
python generate_5w.py "超市" --img /tmp/x.jpg   # 指定本地图片
python generate_5w.py "餐厅" --url "https://..."  # 指定图片URL
python generate_5w.py                 # 交互式选择已有场景
```

### 自动流程（用户无感知）

```
输入主题
    ↓
Bing图片搜索（自动尝试3组关键词）
    ↓
下载第一个可用图片
    ↓
vision_analyze 识别图里有什么
    ↓
根据识别结果生成5W问题
    ↓
生成Word图卡到桌面（5W_主题.docx）
```

---

## 完整分步流程（详细步骤）

### Step 1 — Bing图片搜索

**关键词组合策略（自动执行）：**
1. `卡通 {主题} 小朋友`
2. `卡通 {主题} 儿童`
3. `儿童 {主题}`

**底层命令：**
```bash
curl -s --max-time 10 -A "Mozilla/5.0" \
  "https://cn.bing.com/images/search?q=$(python3 -c "import urllib.parse;print(urllib.parse.quote('卡通 公交车 小朋友'))")&first=1"
```

**解析方法：** 搜索结果HTML中的`murl`字段（真实图片URL），需解码HTML实体：
```python
text = raw_html.replace('&quot;', '"').replace('&amp;', '&')
imgs = re.findall(r'"murl":"(https://[^"]+\.(?:jpg|jpeg|png))"', text)
```

**优先选择标准：**
- 卡通/插画风格 > 真实照片
- 优先包含人物（小朋友）的场景
- 优先 .cn / .com 国内图库域名

### Step 2 — 下载图片

```bash
curl -s --max-time 15 -A "Mozilla/5.0" -o /tmp/topic.jpg "图片URL"
ls -lh /tmp/topic.jpg  # 验证大小 > 10KB
```

### Step 3 — AI识别（vision_analyze）

**分两步：**

**第一步：快速确认（判断是否换图）**
```
问题："图里有小朋友吗？是卡通还是真实照片？主要场景是什么？一句话描述。"
```
- ❌ 空场景、成人内容 → 换下一个候选URL
- ✅ 有小朋友 + 场景相关 → 进入第二步

**第二步：详细描述（决定问题内容）**
```
问题："详细描述图中所有人物的外貌、穿着、动作、表情，以及场景里的物品和细节。用中文。"
```

### Step 4 — 生成5W问题

根据Step 3的详细描述，按6个维度生成问题：

| 维度 | 问法 | 示例（基于描述提取） |
|------|------|------|
| 谁 Who | 人物身份、外貌、穿着 | "校车是什么颜色的？" |
| 什么 What | 物品、颜色、数量 | "校车前面有什么装饰？" |
| 哪里 Where | 位置、方向 | "校车在哪里行驶？" |
| 什么时候 When | 季节、时间 | "这是什么时间的画面？" |
| 为什么 Why | 原因、理由 | "校车为什么要画笑脸？" |
| 做什么 What doing | 动作、行为 | "校车在做什么？" |

**原则：每个维度至少3个问题，问题必须来自图中真实可见的元素。**

### Step 5 — 生成Word

```python
generate_5w_docx(topic, img_path, questions, output_path)
```

输出：`5W_{主题}.docx` 到桌面

---

## Bing搜索失效时的备选方案

如果Bing搜索无结果或下载失败：

1. **换搜索词**：去掉"卡通"，直接用主题
2. **换图片格式**：尝试`.png`而非`.jpg`
3. **手动搜索**：在浏览器打开 cn.bing.com/images 手动搜索，找到满意图片后复制URL，用`--url`参数

---

## 快速参考

| 任务 | 命令 |
|------|------|
| 搜索图片 | `python generate_5w.py "主题"` |
| 指定图片 | `python generate_5w.py "主题" --url "URL"` |
| 查看已有场景 | `python generate_5w.py`（无参数） |
| 测试搜索 | `python3 -c "from generate_5w import bing_search; print(bing_search('卡通 超市 小朋友'))"` |

---

## 关键文件位置

| 文件 | 路径 |
|------|------|
| 生成脚本 | `/mnt/c/Users/123/Desktop/小螺号会员体系/TC- 作业助手 - 02 hermes/generate_5w.py` |
| 场景数据 | `/mnt/c/Users/123/Desktop/小螺号会员体系/TC- 作业助手 - 02 hermes/5w_scenes.json` |
| 图片目录 | `/mnt/c/Users/123/Desktop/小螺号会员体系/TC- 作业助手 - 02 hermes/5w_images/` |
| 输出桌面 | `/mnt/c/Users/123/Desktop/5W_{主题}.docx` |

---

## 实战示例：公交车场景

```
用户输入: python generate_5w.py "公交车"

自动执行:
1. Bing搜索 "卡通 公交车 小朋友" → 找到8个候选URL
2. 下载第一个: https://img2.douhuiai.com/dhimg/20240111/659f4b914525f.jpg (352KB ✅)
3. vision识别:
   → 黄色卡通校车，蓝色笑脸，大眼睛
   → 城市背景，粉色/橙色/蓝色/白色楼房
   → 蓝天白云，一辆蓝色小汽车在路面
4. 生成5W问题（每个维度3-4个）:
   - 谁: 校车什么颜色？
   - 什么: 车头有什么装饰？
   - 哪里: 校车在哪里？
   - 什么时候: 这是白天还是晚上？
   - 为什么: 为什么要画笑脸？
   - 做什么: 校车在做什么？
5. 生成Word: 5W_公交车.docx (388KB)
```

---

## 注意事项

- **图片下载用curl而非urllib**：WSL环境下urllib的SSL兼容性差，curl更可靠
- **必须解码HTML实体**：Bing搜索结果中图片URL在`murl`字段里，需要`&quot;`→`"`解码
- **图片必须验证**：下载后先用vision_analyze确认内容，再生成问题
- **卡通优先**：卡通图画面清晰、内容针对性强、无版权风险
