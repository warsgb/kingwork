---
name: kingwork
description: KingWork 工作智能管理技能。支持自然语言记录工作日记、AI自动分析聊天/会议记录、智能提醒待办和客户跟进、生成日报周报月报。触发词：kw、kingwork、kr、记录、记一下、写日记、kwc、客户、拜访、kwt、待办、任务、kwl、学习、培训、kws、横向、kwm、团队、kwi、灵感、想法、kwp、方案、kwf、修复、排障、kwe、活动、接待、kwr、日报、周报、月报、kwa、我的任务、提醒、kwg、团队同步、kwq、查询、分析、kwu、更新、kurl、krun、auto、kwb、找资料、kwcfg、配置、kingrecord。
---

# KingWork - 工作智能管理

## 快捷指令路由表（Agent 必读）

用户输入的**第一个词**如果匹配以下触发词，直接路由到对应技能，**剩余内容**作为参数传入。

### 记录类（→ kingrecord）

| 触发词 | 目标类型 | 示例输入 |
|--------|---------|---------|
| `kr` / `记录` / `记一下` / `写日记` | AI 自动分类 | `记录 今天和客户开了个会` |
| `kwc` / `k1` / `客户` / `拜访` / `客户拜访` | 客户跟进 | `kwc 拜访了华为王总` |
| `kwt` / `k2` / `待办` / `任务` | 待办事项 | `待办 明天交周报` |
| `kwl` / `k3` / `学习` / `培训` / `成长` | 学习成长 | `学习 看了 RAG 论文` |
| `kws` / `k4` / `横向` / `支持` | 横向支持 | `横向 帮销售做了演示` |
| `kwm` / `k5` / `团队` / `管理` | 团队事务 | `团队 开了周会` |
| `kwi` / `k6` / `灵感` / `想法` | 灵感记录 | `灵感 可以做个自动摘要` |
| `kwp` / `k7` / `方案` | 方案编写 | `方案 写了信创适配文档` |
| `kwf` / `k8` / `修复` / `排障` | 问题服务 | `修复 解决了部署兼容性问题` |
| `kwe` / `k10` / `活动` / `接待` / `市场` | 活动接待 | `接待 信诺时代来访` |

### 其他技能

| 触发词 | 目标技能 | 说明 | 示例 |
|--------|---------|------|------|
| `kwr` / `日报` / `周报` / `月报` / `我的报告` | kingreflect | 生成报告 | `日报` / `周报` |
| `kwa` / `我的任务` / `提醒` / `待办提醒` | kingalert | 智能提醒 | `我的任务` |
| `kwg` / `团队同步` / `同步` | kingteam | 同步至团队 | `团队同步` |
| `kwq` / `查询` / `分析` / `统计` / `获取` | kingquery | 数据查询 | `查询 本周客户跟进` |
| `kwu` / `更新` / `修改记录` | kingupdate | 更新记录 | `更新 上条记录` |
| `kurl` / `灵感 http` | kingclip | URL 灵感收藏 | `灵感 https://...` |
| `krun` / `auto` / `自动分析` | kingauto | AI 自动分析 | `auto` |
| `kwb` / `找资料` / `搜索` | kingbrowse | 资料检索 | `找资料 信创方案` |
| `kwcfg` / `配置` | kingconfig | 配置管理 | `配置 list-work-types` |

### 路由规则

1. **精确匹配优先**：第一个词完全匹配触发词 → 直接路由
2. **中文触发词**：如果第一个词是中文触发词（如「客户」「待办」），后面的文字是日记内容
3. **无匹配时**：整句传给 kingrecord，由 AI 自动分类
4. **含 URL 的灵感**：如果内容以「灵感」开头且包含 http 链接 → kingclip

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
| kingconfig | `python skills/kingconfig/run.py` | 配置管理：工作类型增删改查、枚举字段管理 |

## Setup

Python 依赖（pyyaml、python-dateutil、requests）会在首次运行时**自动检测并安装**，无需手动操作。
如需手动安装：

```bash
pip install -r /Users/leo/.wpscomate/agent/skills/custom/kingwork/requirements.txt
```

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
# 自然语言输入（AI 自动分类）
python skills/kingrecord/run.py "今天和某某公司王总电话沟通了新项目需求"

# 字母快捷键（推荐）
python skills/kingrecord/run.py kwc "拜访了华为王总"      # kwc=客户跟进
python skills/kingrecord/run.py kwt "明天交周报"          # kwt=待办事项
python skills/kingrecord/run.py kwl "看了 RAG 论文"       # kwl=学习成长
python skills/kingrecord/run.py kws "帮销售做了演示"      # kws=横向支持
python skills/kingrecord/run.py kwm "开了周会"            # kwm=团队事务
python skills/kingrecord/run.py kwi "可以做个自动摘要"    # kwi=灵感记录
python skills/kingrecord/run.py kwp "写了信创适配文档"    # kwp=方案编写
python skills/kingrecord/run.py kwf "修了部署兼容问题"    # kwf=问题服务
python skills/kingrecord/run.py kwe "接待信诺时代来访"    # kwe=活动接待

# 中文快捷触发（第一个词即类型）
python skills/kingrecord/run.py 客户 "拜访了华为王总"
python skills/kingrecord/run.py 待办 "明天交周报"
python skills/kingrecord/run.py 灵感 "可以做个自动摘要"

# 数字快捷键（兼容旧版 k1~k10）
python skills/kingrecord/run.py k1 "客户拜访内容"

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
| 10 | 活动接待记录 | 市场活动、接待来访、展会、渠道宣讲 |
| 20 | 惊喜文档记录 | 有价值文档 |
| 21 | 惊喜沟通记录 | 有价值沟通 |
| 22 | 惊喜会议记录 | 有价值会议（AI自动写入）|

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
