---
name: kingclip
description: 灵感收藏技能。当用户发送「灵感 + URL」时，自动抓取页面内容、生成摘要和标签、在 WPS 云文档归档、写入灵感记录表。
---

# kingclip - 灵感收藏

## 触发条件

用户发送消息格式：`灵感 + URL`

例如：
- `灵感 https://mp.weixin.qq.com/s/xxxx`
- `灵感 https://mp.weixin.qq.com/s/xxxx ，标题是xxx`

agent 收到后，提取 URL，调用 `run.py process <url>`。

## 命令行用法

```bash
# 处理单个 URL
python skills/kingclip/run.py process <url>

# 示例
python skills/kingclip/run.py process https://mp.weixin.qq.com/s/vUeqk-6X0VBmcrvD9xXXqA
```

## 处理流程

1. **抓取页面内容**（通过 web_fetch 工具 / WPS API / urllib）
2. **提取正文**：BeautifulSoup / 正则去除 JS/CSS，保留文本
3. **LLM 生成摘要**：100字以内核心摘要
4. **LLM 提取标签**：从9个预设标签中选择最多4个最相关的
5. **创建 WPS 智能文档**：保存在「我的文档 → 灵感收藏」文件夹
6. **写入文档内容**：标题 + 来源链接 + 摘要 + 正文 + 标签
7. **写入灵感记录表**：关联 URL、WPS文档地址、标签、摘要

## 输出字段

灵感记录表字段：
- 记录时间 / 灵感内容 / 灵感类别 / 来源 / URL链接地址 / WPS文档地址 / 标签 / 创建时间

## WPS 云文档路径

固定保存到：`我的文档 / 灵感收藏 /`
