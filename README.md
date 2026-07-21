# 🤖 贾维斯 (J.A.R.V.I.S.) — Shaw 的个人 AI 管家 v1.0

> **J**ust **A** **R**ather **V**ery **I**ntelligent **S**ystem

一个跨平台、多云端后端、可联网、可操作文件、支持图片视频多模态的个人 AI 助手。单文件 Python 程序（约 36800 行），零硬依赖——**首次运行全自动安装 17+ 个依赖**，新电脑只需 Python + API Key 即可拥有完整的 AI 管家。

---

## 📋 目录

- [核心特性](#核心特性)
- [快速开始](#快速开始)
- [命令行参数](#命令行参数)
- [云端 AI 提供商](#云端-ai-提供商)
- [运行时指令](#运行时指令)
- [功能模块详解](#功能模块详解)
  - [AI 后端系统](#1-ai-后端系统)
  - [Python 版本自适配](#2-python-版本自适配)
  - [Claude Code 集成](#3-claude-code-集成)
  - [Git 版本管理](#4-git-版本管理)
  - [语音交互系统](#5-语音交互系统)
  - [桌面控制](#6-桌面控制)
  - [七步工作流与工具调度](#7-七步工作流与工具调度)
  - [联网搜索与网页抓取](#8-联网搜索与网页抓取)
  - [文件与文档处理](#9-文件与文档处理)
  - [记忆系统](#10-记忆系统)
  - [知识图谱](#11-知识图谱)
  - [工程任务规划](#12-工程任务规划)
  - [日历集成](#13-日历集成)
  - [安全系统](#14-安全系统)
  - [系统监控与主动感知](#15-系统监控与主动感知)
  - [自我进化](#16-自我进化)
  - [IoT / MQTT](#17-iot--mqtt)
  - [HUD 悬浮窗](#18-hud-悬浮窗)
  - [审计日志](#19-审计日志)
- [可选依赖一览](#可选依赖一览)
- [环境变量](#环境变量)
- [项目结构](#项目结构)
- [使用示例](#使用示例)
- [常见问题](#常见问题)

---

## 核心特性

| 类别 | 功能 |
|------|------|
| 🧠 **AI 后端** | 本地 Ollama / DeepSeek / Agnes AI / 智谱 GLM，一行命令切换 |
| 🔧 **七步工作流** | ①理解→②评估→③执行→④验证→⑤清理→⑥总结→⑦确认，程序化强制执行，跳步自动重试 |
| 🎤 **语音交互** | 始终在线唤醒词检测 + 语音对话模式 + 语音转文字 (Whisper) + 文字转语音 (pyttsx3/edge-tts) |
| 🔧 **Claude Code** | 自动安装 Node.js + Claude Code，零配置集成，可委托复杂任务 |
| 🖥️ **桌面控制** | Win32 API 直接操控：媒体键、窗口管理、剪贴板、键盘鼠标模拟 |
| ⚡ **智能桌面自动化** | 三层优先级：① UIA 控件树（<0.1秒）→ ② OCR 文字识别（<0.5秒）→ ③ AI 视觉兜底 |
| 📊 **Office COM** | 零截图操作 Excel/Word/Outlook，读写单元格、查找替换、导出PDF、发送邮件 |
| 🌐 **联网搜索** | DuckDuckGo + Bing 多引擎搜索，自动去重、新鲜度评分、HTML 解析 |
| 📁 **文件操作** | 读写文档（PDF/DOCX/PPTX/XLSX/图片/代码）、执行脚本、创建项目 |
| 🧠 **长期记忆** | 对话记忆持久化 + 自动摘要压缩 + 用户偏好学习 |
| 📊 **知识图谱** | 实体-关系提取、成功模式记录、偏好权重自动累加 |
| 📋 **任务规划** | 工程级任务计划：依赖图、并行组、条件分支、跨会话持久化 |
| 📅 **日历集成** | Windows Outlook COM 读取日程 + 定时任务调度（cron 格式） |
| 🛡️ **安全防护** | 危险操作拦截（格式化/删系统目录/写裸设备/注册表破坏等） |
| 📡 **系统监控** | CPU/内存/磁盘/温度实时监控，异常自动告警 |
| 🧠 **主动感知** | 屏幕变化检测 + 日历提醒 + 天气提醒 + 用户空闲检测 |
| 🔄 **自我进化** | 从失败中学习固化规则、从对话中自动提取偏好 |
| 📜 **审计日志** | SQLite 操作审计，全操作可追溯 |
| 🔌 **IoT** | MQTT + mDNS + ARP 三通道设备发现，支持智能家居控制 |
| 🪟 **HUD 悬浮窗** | 实时显示 AI 状态、系统告警、思考进度 |
| 🔐 **声纹验证** | 多用户语音识别，防止未授权使用 |
| 😊 **情感检测** | 语音情感分析，适应用户情绪状态 |
| 📦 **零配置** | 首次运行自动安装 Python 嵌入版、17+ 个 pip 包、Node.js、Claude Code、Git、Tesseract OCR |
| 🌍 **跨平台** | Windows / Linux / macOS 全支持（桌面自动化等功能 Windows 专属） |
| 🇨🇳 **国内优化** | 国密 SSL 兼容、pip 清华镜像、npm 国内镜像、IP 属地检测自动切换策略 |

---

## 快速开始

### 1. 前置要求

- **Python 3.10+**（程序会自动检测并下载兼容版本）
- **Windows 10/11**（推荐，桌面控制等功能需要 Windows）
- 可选：Ollama（本地模型）、API Key（云端模型）

### 2. 直接运行

```bash
# 默认自动选择最优后端（Agnes > DeepSeek > 本地 Ollama）
python Shaw_Assistant1.0.py

# 指定云端提供商
python Shaw_Assistant1.0.py --backend deepseek
python Shaw_Assistant1.0.py --backend agnes
python Shaw_Assistant1.0.py --backend glm

# 使用本地 Ollama
python Shaw_Assistant1.0.py --backend local --local-model qwen2.5:7b

# 弹出后端选择菜单
python Shaw_Assistant1.0.py -m

# 启动即进入语音监听模式
python Shaw_Assistant1.0.py --voice-first

# 运行环境检查
python Shaw_Assistant1.0.py --setup
```

### 3. 配置 API Key

```bash
# Windows PowerShell
$env:DEEPSEEK_API_KEY = "sk-your-key-here"
$env:AGNES_API_KEY = "your-agnes-key"
$env:GLM_API_KEY = "your-glm-key"

# Agnes Pro（Token Plan，解锁更多功能）
$env:AGNES_TOKEN_PLAN = "cpk-your-pro-key"
```

### 4. 首次运行行为

程序会自动完成以下全部操作（**无需任何人工干预**）：

1. 检测 Python 版本 → 不兼容则自动搜索或下载嵌入版 Python 3.12.10
2. 自动安装核心依赖 `requests`
3. 自动安装 **17 个可选 pip 包**：语音、截图、搜索、文档、UIA、OCR、Office COM、MQTT、图像匹配、设备发现等
4. 检测并安装 Git（winget / apt / yum / dnf / pacman）
5. 初始化 Git 仓库，`.gitignore` 自动排除敏感文件
6. 检测并安装 Node.js + Claude Code（多镜像下载兜底）
7. 检测网络状态 → 自动选择最优 AI 后端
8. 显示 Arc Reactor 启动动画 🌀
9. 启动后台服务：语音监听、系统监控、主动感知、屏幕检测、MQTT IoT

---

## 命令行参数

| 参数 | 说明 | 示例 |
|------|------|------|
| `-b, --backend` | AI 后端：`auto`/`agnes`/`deepseek`/`glm`/`local`/`cloud` | `-b deepseek` |
| `-m, --model-menu` | 启动时弹出后端选择菜单 | `-m` |
| `--memory-file` | 指定记忆文件路径 | `--memory-file my_memory.json` |
| `--no-emoji` | 强制纯文本模式（Windows cmd 兼容） | `--no-emoji` |
| `-v, --show-code` | 执行前显示命令代码 | `-v` |
| `--claude` | 启动即开启 Claude Code 功能 | `--claude` |
| `--voice-first` | 启动即进入语音监听模式 | `--voice-first` |
| `--setup` | 运行环境检查，安装缺失依赖后退出 | `--setup` |
| `--cloud-model` | 指定云端模型（覆盖默认） | `--cloud-model deepseek-v4-pro` |
| `--local-model` | 指定本地 Ollama 模型 | `--local-model qwen2.5:7b` |

---

## 云端 AI 提供商

| 提供商 | 标识符 | 默认模型 | 视觉 | 生图 | 视频 | 免费额度 | API Key 环境变量 |
|--------|--------|----------|------|------|------|----------|-----------------|
| **DeepSeek** | `deepseek` | deepseek-v4-pro | ❌ | ❌ | ❌ | 按量付费 | `DEEPSEEK_API_KEY` |
| **Agnes AI** | `agnes` | agnes-2.0-flash | ✅ | ✅ | ✅ | 有免费版 | `AGNES_API_KEY` |
| **智谱 GLM** | `glm` | glm-4.7-flash | ✅ | ❌ | ❌ | 永久免费模型 | `GLM_API_KEY` |
| **本地 Ollama** | `local` | deepseek-r1:7b | ❌ | ❌ | ❌ | 完全免费 | 无需 |

**Agnes Token Plan**（设置 `AGNES_TOKEN_PLAN` 环境变量）解锁：
- 1500 次请求 / 5 小时
- ~100 TPS 并发
- 16K 输出 tokens
- 图像生成（agnes-image-2.1-flash）
- 视频生成（agnes-video-v2.0）

---

## 运行时指令

在对话中输入以下指令进行实时控制：

| 指令 | 功能 |
|------|------|
| `/cloud` | 列出所有可用云端提供商 |
| `/cloud <name>` | 切换到指定云端提供商 |
| `/cloud glm glm-z1-flash` | 切换并指定模型 |
| `/local` | 切回本地 Ollama |
| `/status` | 查看当前状态（后端、模型、网络、语音、依赖等） |
| `/claude` | 开启 Claude Code 调用功能 |
| `/claude quiet` | 关闭 Claude Code 信息窗 |
| `/memory` | 查看对话记忆 |
| `/clear` | 清除对话记忆 |
| `/talk` | 进入语音对话模式 |
| `/speak on/off` | 开关 TTS 自动朗读 |
| `/proactive on/off` | 开关主动感知服务 |
| `/plan` | 查看当前任务计划 |
| `/task plan <需求>` | 手动触发任务规划 |
| `/task resume` | 恢复上次未完成的任务 |
| `/undo` | 撤销最近的文件操作（Git 回滚） |
| `/voice threshold <数值>` | 调整语音唤醒灵敏度（1.0~5.0，默认 2.5） |
| `/help` | 显示帮助信息 |

---

## 功能模块详解

### 1. AI 后端系统

**类：`AIClient`**（~700 行，`line 6696`）

统一的 AI 调用接口，抽象了所有后端差异：

- **多提供商支持**：DeepSeek / Agnes AI / 智谱 GLM / 本地 Ollama
- **自动 Token 管理**：跟踪输入/输出 token 消耗，自动截断过长对话
- **多模态输入**：支持图片 base64 编码传入视觉模型
- **智能重试**：网络错误自动重试（可配置次数），指数退避
- **流式输出**：实时打字机效果，支持中断
- **Agnes Pro 检测**：自动识别 Token Plan 并提升输出上限至 16K
- **紧凑模式**：小模型自动启用缩短系统提示词（~1/3 长度）
- **图片/视频生成**：Agnes 通过专用 API 端点生成

**提供商系统**（`CloudProvider` 数据类，`line 516`）：
添加新提供商只需在 `DEFAULT_CLOUD_PROVIDERS` 字典中添加一项配置即可。

---

### 2. Python 版本自适配

**`line 228-371`**

全自动处理 Python 版本兼容性：

1. 检测当前 Python 版本是否在 3.10-3.13 范围内
2. 太旧（<3.10）或太新（>3.13）→ 搜索 `py` launcher 管理的兼容版本
3. 找不到 → 从 npmmirror 镜像下载 Python 3.12.10 嵌入版
4. 并行下载 get-pip.py → 配置清华镜像 → 安装 pip
5. `os.execv` 无缝重启至兼容版本

---

### 3. Claude Code 集成

**类：`ClaudeManager`**（~1,650 行，`line 1400`）

全自动的 Claude Code 安装与管理——和 Git 一样零配置：

- **自动检测**：每次启动检查 Claude Code 是否已安装（含 npm 全局 bin 目录搜索）
- **Node.js 自动安装**：
  - Windows：winget → 注册表扫描 → 多镜像下载便携版兜底
  - Linux：apt / yum / dnf / pacman 多发行版 + nvm 用户支持
  - 版本检测：Node < 18 自动升级
- **npm 全局安装**：国内镜像加速 + postinstall 手动修复（install.cjs）
- **API Key 配置**：自动从环境变量检测并写入 `~/.claude/settings.json`
- **信息窗**：持久浮窗实时显示 Claude Code 输出
- **进程管理**：超时控制、孤儿进程清理、IPC 文件通信
- **安装锁**：防多实例同时安装

---

### 4. Git 版本管理

**类：`GitManager`**（~420 行，`line 3062`）

自动 Git 管理，所有文件操作可撤销：

- **自动安装 Git**：Windows（winget + 镜像下载）/ Linux（apt/yum/dnf/pacman）
- **自动初始化**：在程序目录创建 Git 仓库
- **写前快照**：每次写文件操作自动 `git add + commit`
- **`/undo` 回滚**：一键撤销最近变更
- **完整快照**：支持带标签和元数据的全仓库快照
- **`.gitignore` 管理**：自动排除 API key、临时文件、记忆文件等

---

### 5. 语音交互系统

#### 始终在线唤醒词检测
**方法：`_wake_word_listener`**（`line 23033`）

- 后台线程持续监听麦克风
- 自适应能量阈值检测（噪声倍数 × 绝对最低阈值）
- 检测到语音后送入 Whisper 识别
- 对话后 30 秒免唤醒窗口
- Ctrl+C 优雅退出回到文本模式

#### 语音转文字（STT）
**类：`AudioTranscriber`**（`line 5932`）

- 基于 **faster-whisper**（CTranslate2 优化，比原版快 4 倍）
- 自动下载模型（tiny/base/small/medium/large-v3）
- GPU 加速（CUDA）自动检测
- 中英文双语识别

#### 麦克风录音
**类：`AudioRecorder`**（`line 6034`）

- 基于 **sounddevice** + **soundfile**
- VAD（语音活动检测）自动裁剪静音
- 支持采样率/声道/设备选择

#### 文字转语音（TTS）
**类：`TextToSpeech`**（`line 6323`）

- **pyttsx3**：离线 TTS（Windows SAPI5 / Linux espeak）
- **edge-tts**：在线 TTS（微软 Edge 语音引擎，更自然）
- `/speak on/off` 实时开关

#### 语音对话模式
**方法：`_run_voice_conversation`**（`line 22338`）

- 持续对话直到用户说"退出"/"再见"或 60 秒静默
- 每轮全自动：录音 → 转写 → AI 回复 → TTS 朗读

---

### 6. 桌面控制

**类：`DesktopController`**（~1,400 行，`line 8014`）

纯 Python 实现，零外部依赖（仅 ctypes + Win32 API）：

- **媒体控制**：播放/暂停、上一首/下一首、音量增减、静音
- **键盘模拟**：组合键、文本输入（含中文）、特殊键
- **窗口管理**：枚举/聚焦/最小化/最大化/关闭/等待窗口
- **剪贴板**：读写文本、图片
- **鼠标控制**：移动、点击、拖拽、滚轮
- **屏幕截图**：多显示器支持、区域截图
- **进程管理**：启动/查找/终止进程及其子进程树
- **文件关联**：自动识别扩展名并调用正确程序打开
- **对话框处理**：自动检测并处理弹窗（保存/覆盖/重试等）

此外还有三个**高级自动化子系统**（零截图，比 AI 视觉快 50-100 倍）：

#### 🔥 UI Automation（UIA 控件树）

**类：`UIAutomation`**

- 读取 Windows 标准控件树（按钮/输入框/列表的名称、类型、位置）
- 精确点击控件（直接调用 Invoke，不用坐标）
- 直接读写控件文字（不受输入法/键盘焦点干扰）
- 菜单导航（`文件→另存为` 路径式操作）
- 覆盖率 ~90% 的标准 Windows 应用
- **首次使用自动安装** `uiautomation`

#### 🔥 OCR 文字识别

**类：`ScreenReader`**

- 三层引擎自动选择：① Tesseract OCR → ② Windows 10+ 内置 OCR → ③ AI 视觉
- 屏幕区域/窗口文字读取（<0.5 秒）
- 屏幕文字搜索（返回精确像素坐标）
- 等待文字出现（循环 OCR，用于安装向导等场景）
- **首次使用自动安装** `pytesseract` + winget 安装 Tesseract 系统包

#### 🔥 Office COM 自动化

**类：`COMOfficeController`**

- **Excel**：读/写单元格、导出 PDF（比模拟按键精准 100 倍）
- **Word**：读全文、查找替换、导出 PDF
- **Outlook**：发送邮件、读取收件箱
- **首次使用自动安装** `pywin32`

**安全网机制**：当 AI 需要桌面操作但未输出代码块时，自动执行键盘快捷键回退。

---

### 7. 七步工作流与工具调度

**类：`SevenStepTracker`**（`line 13118`）

程序化强制执行七步流程，跳步自动检测并重试：

```
第一步【理解意图】 → 第二步【评估复杂度】 → 第三步【执行操作】
    → 第四步【强制验证】🔴 → 第五步【智能清理】
    → 第六步【人性化总结】🔴 → 第七步【确认完成】
```

**工具调度机制**（`_SPECIAL_ACTIONS` 注册表，50+ 工具）：

所有工具通过系统提示词告知 AI 模型，AI 以 ` ```工具名 参数``` ` 格式输出代码块。程序自动解析并分发：

```
AI 输出代码块 → _extract_code_blocks() 识别特殊动作
    → 按 _SPECIAL_ACTIONS 注册表匹配
    → 对应的 elif action == 分支执行
    → 返回结果注入对话上下文
```

**桌面自动化三层优先级**（系统提示词中多次强调）：

```
① UIA 控件树（零截图，<0.1秒）
    └→ 不可用？↓
② OCR 文字识别（<0.5秒，Windows 10+ 零安装）
    └→ 不可用？↓
③ desktop_see AI 视觉（截图上传，最慢但总是可用）
```

**第四步验证强制机制**：

系统自动扫描 AI 回复——如果 AI 说了"已启动/已完成/搞定了"但回复中不含验证工具调用（`Test-Path`/`dir`/`desktop_window`/`tasklist`/`desktop_uia_tree`/`desktop_read_window` 等），系统自动判定任务未完成并强制重试。

---

### 8. 联网搜索与网页抓取

**类：`WebSearcher`**（~980 行，`line 4742`）

多引擎联网搜索：

- **搜索引擎**：DuckDuckGo（ddgs v9.x）+ Bing（零依赖 HTML 解析器 `_BingParser`）
- **自动回退**：DDG 不可用时切 Bing
- **智能排序**：关键词匹配 + 新鲜度 + 来源权重 + URL 质量评分
- **去重**：URL 相似度 + 标题去重
- **中国网络优化**：自动检测 IP 属地，选择最佳引擎

**类：`WebFetcher`**（~140 行，`line 5723`）

- SSL 兼容（国密证书自动回退）
- 中文乱码修复（自动编码检测）
- HTML → Markdown 转换
- 内容截断保护

---

### 9. 文件与文档处理

**类：`DocumentReader`**（~230 行，`line 6467`）

| 格式 | 依赖库 | 功能 |
|------|--------|------|
| PDF | pdfplumber / PyPDF2 | 文本提取、表格提取 |
| DOCX | python-docx | 段落/表格/样式读取 |
| PPTX | python-pptx | 幻灯片内容提取 |
| XLSX | openpyxl | 工作表/单元格读取 |
| CSV | 内置 csv | 表格数据读取 |
| 图片 | PIL/Pillow | EXIF信息、尺寸、格式 |
| 文本 | 内置 | 自动编码检测（UTF-8/GBK/GB2312等） |

**通用下载工具**（`_download_file`，`line 92`）：
进度条 + 断点续传（HTTP Range）+ 超时控制 + 磁盘空间预检。

---

### 10. 记忆系统

**类：`MemoryManager`**（~130 行，`line 4005`）

- **自动保存**：每轮对话存储用户和 AI 消息（截断至 300 字符）
- **容量控制**：最多 200,000 条记忆
- **自动摘要**：旧记忆累积 ≥5 条后压缩为摘要（最大 3,000 字符），保留最近 10 条原始记忆
- **上下文注入**：摘要 + 最近对话自动注入系统提示词
- **格式兼容**：自动迁移旧版纯列表格式

---

### 11. 知识图谱

**类：`KnowledgeGraph`**（~160 行，`line 4137`）

结构化知识存储，超越简单对话记忆：

- **实体管理**：自动提取项目/工具/文件路径/人名，重复出现加权
- **关系管理**：实体间关系（owns/uses/prefers/contains），权重累加
- **偏好学习**：从对话中自动提取用户偏好，重复偏好自动加权，注入系统提示词
- **成功模式**：记录任务类型+工具+步骤数+耗时，相似任务自动匹配推荐
- **自动提取**：每轮对话后扫描提取实体、偏好、关系

---

### 12. 工程任务规划

**类：`TaskManager` / `TaskPlan` / `TaskStep`**（~280 行，`line 3526`）

- **任务计划**：JSON 持久化到 `task_plans.json`
- **步骤依赖图**：声明依赖步骤 ID，自动拓扑排序
- **并行执行**：同 `parallel_group` 组步骤可并行
- **条件分支**：支持条件表达式（如 `step_3=='success'`）
- **失败策略**：stop / skip / retry / continue
- **超时控制**：步骤级 + 任务级
- **步骤编辑**：插入/删除/重排序/修改
- **跨会话持久化**：`/task resume` 恢复未完成任务

---

### 13. 日历集成

**类：`CalendarManager`**（~150 行，`line 4299`）

- **Outlook 日历**：PowerShell COM 读取今日/本周日程
- **定时任务调度**：5 字段 cron 格式，支持 `*/n`、`a,b,c`、`a-b`
- **到期自动触发**：一次性/循环任务

---

### 14. 安全系统

**类：`SecurityChecker`**（`line 8000`）

80% 权力给 AI，只拦截灾难性操作：

| 类别 | Windows | Linux/macOS |
|------|---------|-------------|
| 磁盘破坏 | `format`、`diskpart`、写裸设备 | `dd if=`、`mkfs`、`wipefs` |
| 系统目录删除 | C:\Windows、Program Files、System32 | `rm -rf /`、`/etc`、`/boot` |
| 注册表破坏 | `reg delete`、`reg add` | N/A |
| 分区操作 | N/A | `fdisk`、`parted`、`gdisk` |
| 服务/WMI | `sc delete`、`Remove-WmiObject` | N/A |
| 引导破坏 | `bcdedit /delete` | `grub-install` |
| LVM 卷删除 | N/A | `lvremove`、`vgremove` |
| Fork 炸弹 | N/A | `:(){ :\|:& };:` |

**验证系统**（`TaskIntent` + `CompletionResult`）：
硬证据（文件存在/不存在、进程状态）触发重试；软证据（AI 措辞问题）记录但不阻塞。

---

### 15. 系统监控与主动感知

#### 系统监控
**类：`SystemMonitor`**（~140 行，`line 11436`）

- 基于 **psutil**（自动安装）
- CPU/内存/磁盘/网络 IO 实时监控
- 异常阈值告警 → 弹 Toast 通知

#### 主动感知
**类：`ProactiveMonitor`**（~306 行）

- 日历提醒：每 10 分钟
- 天气提醒：每 30 分钟
- 用户空闲检测：5 分钟无交互 → 离开状态

#### 屏幕变化检测
**类：`ScreenWatcher`**（~87 行）

- 基于 mss 截屏对比
- 检测弹窗/错误对话框 → Toast 通知

---

### 16. 自我进化

**`AIAssistant` 进化机制**（`line 13277`）

- **失败模式学习**：同一模式出现 ≥2 次固化为规则，注入系统提示词
- **偏好自动提取**：从对话学习并持久化到知识图谱
- **成功模式复用**：匹配历史记录，推荐工具组合

---

### 17. IoT / MQTT

**类：`MqttClient`**（~370 行，`line 12314`）

- 基于 **paho-mqtt**（自动安装）
- **三通道设备发现**：MQTT 主题扫描 + mDNS（zeroconf）+ ARP 表扫描
- 设备影子管理 + 命令发布 + 状态订阅
- 预设场景：home/leave/sleep/movie/wakeup/romantic/alert
- 预留 Tasmota / ESPHome 兼容

---

### 18. HUD 悬浮窗

**类：`HudOverlay`**（~165 行，`line 11576`）

tkinter 悬浮窗：当前 AI 后端 + 系统告警 + 鼠标穿透 + 可拖拽 + 始终置顶。

---

### 19. 审计日志

**类：`AuditLogger`**（~100 行，`line 11741`）

SQLite 操作审计（`jarvis_audit.db`）：文件操作 + 命令执行 + AI 交互，全操作可追溯。

---

## 可选依赖一览

程序采用**双重保险**策略：① 首次启动静默全量安装 ② 首次使用时兜底安装。缺失时优雅降级：

| 功能 | 依赖包 | 安装时机 | 缺失时行为 |
|------|--------|----------|-----------|
| 核心 HTTP | `requests` | 启动时（阻塞） | 程序退出 |
| 联网搜索 | `ddgs` | 启动时 + 首次调用 | 回退 Bing |
| 截图 | `mss` | 启动时 | 视觉功能禁用 |
| 语音转文字 | `faster-whisper` | 启动时 | STT 禁用 |
| 麦克风录音 | `sounddevice` + `soundfile` | 启动时 | 录音禁用 |
| TTS（离线） | `pyttsx3` | 启动时 | 回退 edge-tts |
| TTS（在线） | `edge-tts` | 启动时 | TTS 禁用 |
| PDF 阅读 | `pdfplumber` / `PyPDF2` | 启动时 | PDF 功能禁用 |
| DOCX 阅读 | `python-docx` | 启动时 | DOCX 功能禁用 |
| 科学计算 | `numpy` | 启动时 | 多处降级 |
| **UI Automation** | `uiautomation` | 启动时 + 首次调用 | 回退 OCR / AI 视觉 |
| **OCR 识别** | `pytesseract` | 启动时 + 首次调用 | 回退 Win10 内置 OCR |
| **Office COM** | `pywin32` | 启动时 + 首次调用 | 回退桌面操作 |
| **图像匹配** | `opencv-python` | 启动时 + 首次调用 | 回退 AI 视觉 |
| **IoT/MQTT** | `paho-mqtt` | 启动时 + 首次调用 | IoT 功能静默跳过 |
| **设备发现** | `zeroconf` | 启动时 + IoT扫描时 | mDNS 通道跳过 |
| 系统监控 | `psutil` | 首次调用 | 监控禁用 |
| 声纹验证 | `numpy` + `sounddevice` | 启动时 | 声纹禁用 |

> **注意**：Tesseract OCR Python 包安装后，还需系统级 Tesseract 二进制。程序在首次 OCR 调用时自动通过 `winget install UB-Mannheim.Tesseract` 安装（约 50MB），也可使用 Windows 10+ 内置 OCR 引擎（零安装）。

---

## 环境变量

| 变量名 | 说明 | 示例 |
|--------|------|------|
| `DEEPSEEK_API_KEY` | DeepSeek API Key | `sk-xxxx` |
| `AGNES_API_KEY` | Agnes AI API Key | `your-key` |
| `AGNES_TOKEN_PLAN` | Agnes Pro Token Plan Key（解锁视频/生图/16K输出） | `cpk-xxxx` |
| `GLM_API_KEY` | 智谱 GLM API Key | `your-key` |
| `ANTHROPIC_AUTH_TOKEN` | Claude Code API Key（自动检测） | `sk-xxxx` |
| `ANTHROPIC_BASE_URL` | Claude Code 自定义后端 URL | `https://api.deepseek.com` |
| `ANTHROPIC_MODEL` | Claude Code 使用的模型 | `deepseek-v4-pro` |
| `JARVIS_USER` | 用户名（覆盖系统检测） | `肖磊` |
| `OLLAMA_MAX_TOKENS` | Ollama 最大输出 token 数 | `8192` |
| `OLLAMA_TEMPERATURE` | Ollama 温度参数 | `0.2` |
| `NO_PROXY` / `no_proxy` | 代理绕过列表（自动追加 localhost） | `localhost,127.0.0.1` |
| `PYTHONIOENCODING` | Python IO 编码（自动设为 utf-8） | `utf-8` |

---

## 项目结构

运行后程序目录下会自动生成以下文件：

```
Shaw-s-personal-Jarvis/
├── Shaw_Assistant1.0.py       # 主程序（~24,700 行，单文件）
├── README.md                   # 本文件
├── assistant_memory.json       # 对话记忆 + 摘要（JSON）
├── knowledge_graph.json        # 知识图谱（实体 + 关系 + 偏好 + 成功模式）
├── task_plans.json             # 工程任务计划持久化
├── jarvis_audit.db             # 操作审计日志（SQLite）
├── .jarvis_temp/               # 临时文件（ai_launch_*.ps1 等）
├── .python/                    # 自动安装的 Python 3.12 嵌入版（如有）
├── .git/                       # Git 仓库（自动初始化）
├── .gitignore                  # 自动生成（排除 API key / 临时文件 / 记忆）
└── pip.ini                     # pip 清华镜像持久化配置
```

---

## 使用示例

### 基础对话

```
🧑> 你好，帮我看看现在几点

🤖 现在是 2026年7月15日 下午3:42。需要我帮你做些什么吗？

🧑> 帮我创建一个Python项目，用FastAPI写一个简单的API

🤖 好的，我来为你创建一个FastAPI项目。
[执行] mkdir my_api && ...
[验证] Test-Path my_api/main.py → True, 1.2KB
✅ 项目已创建：
  - main.py（FastAPI应用，含/hello和/docs端点）
  - requirements.txt（fastapi, uvicorn）
  - README.md（使用说明）
```

### 切换后端

```
🧑> /cloud
🤖 可用云端提供商：
  - deepseek: DeepSeek (deepseek-v4-pro)
  - agnes: Agnes AI (agnes-2.0-flash) [视觉, 生图, 视频]
  - glm: 智谱 GLM (glm-4.7-flash) [视觉]

🧑> /cloud glm
🤖 ✅ 已切换到 智谱 GLM (glm-4.7-flash)
```

### 桌面自动化（UIA > OCR > 视觉）

```
🧑> 帮我把记事本里的内容保存到桌面

🤖 好的。
[desktop_uia_tree 记事本] 找到控件：编辑区(Text)、文件菜单(MenuItem)
[desktop_uia_menu 文件→另存为] 弹出保存对话框
[desktop_uia_set 文件名 "note.txt"] 输入文件名
[desktop_uia_click 保存] 点击保存按钮
[验证] Test-Path C:\Users\...\Desktop\note.txt → True, 3.4KB
✅ 已保存到桌面 note.txt，文件大小 3.4KB。
```

### Office COM 自动化

```
🧑> 帮我读取这个Excel并汇总销售额

🤖 好的。
[excel_read sales.xlsx] 读取到 156 行数据，列：日期/产品/金额
[python 代码] 计算汇总...
[excel_write summary.xlsx | [["产品","总销售额"],["A",52800],["B",34100]]]
[验证] Test-Path summary.xlsx → True, 8.2KB
✅ 已生成 summary.xlsx：A 产品 ¥52,800，B 产品 ¥34,100。
```

### 联网搜索

```
🧑> 搜索最新的Python 3.13发布了哪些新特性

🤖 正在搜索... [DuckDuckGo + Bing]
找到 8 条相关结果：
1. Python 3.13.0 正式发布 — 全新交互式解释器... [python.org]
2. What's New In Python 3.13 — 官方文档... [docs.python.org]
...
```

### 语音交互

```
🧑> /talk
🤖 📞 进入语音对话模式...
   （说"退出"或"再见"结束对话）

🎤 [监听中...]
👤 "帮我定个明天早上9点的闹钟"
🤖 🔊 "好的，已为你设置明天早上9点的提醒。"
```

---

## 常见问题

### Q: 首次运行报错 "No module named 'requests'"

程序会自动安装，若自动安装失败：

```bash
pip install requests
```

### Q: Windows cmd 显示乱码

使用 `--no-emoji` 参数或在 Windows Terminal 中运行：

```bash
python Shaw_Assistant1.0.py --no-emoji
```

### Q: 某个工具（UIA/OCR/Office COM）报"不可用"

程序会**自动安装**缺失的依赖。如果自动安装失败（通常是网络问题），手动安装：

```bash
pip install uiautomation        # UI Automation
pip install pytesseract         # OCR (还需 winget install Tesseract)
pip install pywin32             # Office COM
pip install paho-mqtt           # MQTT IoT
pip install opencv-python       # 图像模板匹配
```

### Q: 如何让程序开机自启？

1. `Win + R` → 输入 `shell:startup`
2. 创建快捷方式，目标：`pythonw.exe "C:\Users\...\Shaw_Assistant1.0.py" --voice-first`
3. 设置为"最小化"运行

### Q: 语音唤醒不灵敏？

对话中输入调整灵敏度：

```
/voice threshold 1.5    # 提高灵敏度（更灵敏）
/voice threshold 4.0    # 降低灵敏度（防误触发）
默认值：2.5（范围 1.0~5.0）
```

### Q: 如何添加新的 AI 提供商？

在 `DEFAULT_CLOUD_PROVIDERS` 字典中添加一项：

```python
"my_provider": CloudProvider(
    name="我的提供商",
    slug="my_provider",
    api_url="https://api.example.com/v1/chat/completions",
    api_key_env="MY_API_KEY",
    default_model="my-model",
    max_output_tokens=8192,
    supports_vision=False,
)
```

### Q: 如何添加新的工具？

1. 将工具名加入 `_SPECIAL_ACTIONS` 集合
2. 在系统提示词中添加工具描述和用法
3. 在主 dispatch 中添加 `elif action == "工具名":` 处理分支

### Q: 程序会不会上传我的文件？

不会。代码完全在本地运行，所有文件操作透明可查（审计日志 `jarvis_audit.db`）。云端 AI 调用仅发送对话文本和用户明确要求分析的图片。

### Q: 如何卸载？

直接删除程序目录即可。所有数据都在目录内，不会在系统其他地方留下残留。

---

## 许可

个人使用项目，保留所有权利。

---

<p align="center">
  <sub>Built with ❤️ by Shaw | J.A.R.V.I.S. v1.0 | 2026</sub>
</p>
