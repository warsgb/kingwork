---
name: kingwork
description: KingWork 工作智能管理技能。支持自然语言记录工作日记、AI自动分析聊天/会议记录、智能提醒待办和客户跟进、生成日报周报月报。触发词：kw、kingwork、工作日记、客户跟进、待办、周报、日报。
---

# KingWork - 售前工作智能管理

## 功能概览

| 子技能 | 命令 | 功能 |
|--------|------|------|
| kingrecord | `python skills/kingrecord/run.py` | 记录工作日记 |
| kingauto | `python skills/kingauto/run.py` | AI 自动分析聊天/会议记录 |
| kingclip | `python skills/kingclip/run.py` | 处理灵感 URL，创建 WPS 文档并写入灵感记录表 |
| kingbrowse | `python skills/kingbrowse/run.py` | 资料检索，搜索多维表内容返回名称+链接 |
| kingalert | `python skills/kingalert/run.py` | 智能提醒待办和客户跟进 |
| kingreflect | `python skills/kingreflect/run.py` | 生成日报周报月报 |
| kingquery | `python skills/kingquery/run.py` | 数据查询和统计 |
| kingupdate | `python skills/kingupdate/run.py` | 更新多维表记录 |
| kingteam | `python skills/kingteam/run.py` | 同步个人数据至团队（客户跟进/日报/周报） |

## 快速开始

### 1. 初始化配置

```bash
# 设置环境变量
export WPS_SID="你的WPS_SID"
export KINGWORK_FILE_ID="多维表文件ID"
export WPS365_SKILL_PATH="/path/to/wps365-skill"

# 初始化多维表（首次使用）
python scripts/init_tables.py
```

### 2. 记录工作日记

```bash
# 自然语言输入
python skills/kingrecord/run.py "今天和某某公司王总电话沟通了新项目需求"

# 快捷指令
python skills/kingrecord/run.py k1 "客户拜访内容"  # k1=客户跟进
python skills/kingrecord/run.py k2 "待办事项"      # k2=待办
python skills/kingrecord/run.py k3 "学习内容"      # k3=学习成长
python skills/kingrecord/run.py k4 "横向支持"      # k4=横向支持
python skills/kingrecord/run.py k5 "团队事务"      # k5=团队事务
python skills/kingrecord/run.py k6 "灵感记录"      # k6=灵感记录

# 结构化输入
python skills/kingrecord/run.py --type "客户跟进" --customer "某某公司" "内容"
```

### 3. AI 自动分析

```bash
# 分析当天聊天和会议记录
python skills/kingauto/run.py

# 指定日期范围
python skills/kingauto/run.py --start 2026-03-01 --end 2026-03-15
```

### 4. 灵感收藏（URL）

当用户发送「灵感 + URL」时，agent 自动调用 kingclip 处理：

```bash
# 处理单个 URL（agent 调用方式）
python skills/kingclip/run.py process <url> --title <标题> --content <预抓取内容>

# 示例
python skills/kingclip/run.py process "https://mp.weixin.qq.com/s/xxxxx" \
  --title "文章标题" \
  --content "文章正文内容..."
```

### 5. 资料检索

当用户询问「帮我找一下 xxx 相关的资料/方案/案例」时，agent 调用 kingbrowse 搜索：

```bash
# 搜索关键词
python skills/kingbrowse/run.py <关键词>

# 示例
python skills/kingbrowse/run.py 端云一体
python skills/kingbrowse/run.py 客户案例 --top 5
```

搜索范围由 `skills/kingbrowse/config.yaml` 配置（支持多文件多表）。

### 4. 智能提醒

```bash
# 查看所有提醒
python skills/kingalert/run.py

# 标记完成
python skills/kingalert/run.py --complete <todo_id>
```

### 5. 生成报告

```bash
python skills/kingreflect/run.py --period daily    # 日报
python skills/kingreflect/run.py --period weekly   # 周报
python skills/kingreflect/run.py --period monthly  # 月报
```

## 配置文件

- `config/kingwork.yaml` - 主配置（文件ID、路径等）
- `config/tables.yaml` - 数据表 Schema
- `config/prompts.yaml` - 大模型提示词模板

## 数据表结构

| 序号 | 数据表 | 说明 |
|------|--------|------|
| 01 | 日记记录 | 主入口表 |
| 02 | 待办记录 | 待办事项管理 |
| 03 | 客户档案 | 客户信息 |
| 04 | 项目档案 | 项目信息 |
| 05 | 客户跟进记录 | 跟进历史 |
| 06 | 学习成长记录 | 学习培训 |
| 07 | 横向支持记录 | 跨部门协作 |
| 08 | 团队事务记录 | 团队内部 |
| 09 | 灵感记录 | 灵感创意 |
| 10 | 惊喜文档记录 | 有价值文档 |
| 11 | 惊喜沟通记录 | 有价值沟通 |
| 12 | 惊喜会议记录 | 有价值会议（AI自动写入）|

## 部署场景与配置

KingWork 与 wps365-skill 的集成支持多种部署场景：

### 场景 1：统一运行环境（推荐）

在统一运行环境中，wps365-skill 已作为 Python 包预安装，KingWork 可以直接导入使用。

```bash
# 环境变量
export WPS_SID="你的WPS_SID"
export KINGWORK_FILE_ID="多维表文件ID"
# 无需设置 WPS365_SKILL_PATH
```

配置 `config/kingwork.yaml`：
```yaml
import_mode: "auto"  # 或 "direct"
skill_call_mode: "subprocess"  # 或 "direct"
```

### 场景 2：独立部署

KingWork 和 wps365-skill 分别部署在不同目录。

```bash
# 环境变量
export WPS_SID="你的WPS_SID"
export KINGWORK_FILE_ID="多维表文件ID"
export WPS365_SKILL_PATH="/path/to/wps365-skill"
```

配置 `config/kingwork.yaml`：
```yaml
import_mode: "path"
skill_call_mode: "subprocess"
```

### 场景 3：开发调试

开发模式下，可以更灵活地配置：

```yaml
import_mode: "auto"  # 自动检测
skill_call_mode: "direct"  # 直接导入，更高效
wps365_skill_path: "${WPS365_SKILL_PATH}"  # 可选配置
```

### 配置说明

| 配置项 | 可选值 | 说明 |
|--------|--------|------|
| `import_mode` | `auto`（默认）| 自动检测，优先直接导入 |
| | `direct` | 强制使用直接导入 |
| | `path` | 强制使用路径导入 |
| `skill_call_mode` | `subprocess`（默认）| 通过子进程调用 |
| | `direct` | 直接导入调用（更高效）|
| `wps365_skill_path` | 路径 | wps365-skill 目录路径 |

### 环境检测

运行以下命令检查当前环境配置：

```bash
python scripts/check_env.py
```
