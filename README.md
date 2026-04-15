# 抖音链接批量下载

从本地 **Excel 或 JSON 文件**中筛选「抖音 APP」来源的链接，自动判断是图集还是视频并下载到本地。

---

## 在 Claude Code 中使用

本工具附带一个 Claude Code 技能文件 `douyin-download.md`，安装后可在 Claude Code 中通过斜杠命令一键调用，无需手动敲命令行参数。

### 安装方式

**项目级（推荐）**：将技能文件放入当前项目的命令目录，仅对该项目生效：

```bash
# 在你的项目根目录下执行
mkdir -p .claude/commands
cp douyin-download.md .claude/commands/douyin-download.md
```

**用户级**：将技能文件放入用户全局目录，对所有项目生效：

```bash
# Windows
copy douyin-download.md "%USERPROFILE%\.claude\commands\douyin-download.md"

# macOS / Linux
cp douyin-download.md ~/.claude/commands/douyin-download.md
```

### 使用方式

安装后，在 Claude Code 中输入：

```
/douyin-download
```

Claude 会自动引导你完成整个流程：
1. 询问输入文件路径（Excel 或 JSON，及可选的输出目录）
2. 确认参数并运行脚本
3. 实时汇报进度（每隔约 1 分钟）
4. 下载完成后展示图集/视频/无效的汇总统计

---

## 使用方法

输入文件（Excel 或 JSON）可以放在电脑上的任意位置，直接把路径传给脚本即可，不需要修改脚本内容。
脚本根据文件扩展名（`.xlsx` / `.json`）自动切换读取方式。

```bash
# 基本用法（输出目录自动创建在输入文件同级目录下）
python batch-download-douyin.py "<输入文件路径>"

# 同时指定输出目录
python batch-download-douyin.py "<输入文件路径>" "<输出目录>"
```

**示例：**
```bash
# Excel 输入
python batch-download-douyin.py "C:/Users/张三/Desktop/yuqing_data.xlsx"

# JSON 输入（如 yuqing-sync-skill 导出的 JSON）
python batch-download-douyin.py "D:/data/yuqing_20260415.json"

# 指定输出目录
python batch-download-douyin.py "D:/data/yuqing_0415.xlsx" "D:/downloads/douyin"
```

**输出目录默认规则**：不指定时，在输入文件同级目录下自动创建 `douyin_downloads_<文件名>/`。

---

## 处理逻辑

```
对每条「抖音 APP」链接：
  │
  ├─ 步骤1：尝试图集下载（extract-douyin-images.js）
  │     成功 → 保存到 <output_dir>/<视频ID>/images/
  │
  └─ 图集失败 → 步骤2：尝试视频下载（parse-douyin-video.js）
        成功 → 保存到 <output_dir>/<视频ID>/video.mp4
        失败 → 标记为「无效」
```

---

## 输出结构

```
<output_dir>/
  7628800085900513994/
    video.mp4               ← 视频
  7628797289502562794/
    images/
      image-01.webp         ← 图集
      image-02.webp
      ...
  download_results.xlsx     ← 结果汇总（含类型/路径/状态，颜色高亮）
```

**子文件夹命名规则**：使用抖音视频 ID（从 URL 中提取）。
当输入文件的 case ID 列填写后，可通过 `download_results.xlsx` 中的对应关系查找。

---

## 结果表格说明

`download_results.xlsx` 包含以下列：

| 列 | 说明 |
|----|------|
| 序号 | 输入文件中的行顺序 |
| case_id | 原始 case ID |
| 视频ID（文件夹名） | 抖音视频 ID，即本地子文件夹名 |
| 原文URL | 原始链接 |
| 标题 | 原始标题 |
| 类型 | 图集 / 视频 / 无效（颜色高亮） |
| 文件路径 | 本地下载路径 |
| 备注 | 图片数量 / 视频大小 / 失败原因 |

---

## 依赖

| 依赖 | 说明 |
|------|------|
| Python 3.10+ | `openpyxl` 用于读写 Excel |
| Node.js 18+ | 运行同目录下的 JS 脚本，无需安装额外 npm 包 |
| `extract-douyin-images.js` | 图集下载（已包含在本文件夹中）|
| `parse-douyin-video.js` | 视频下载（已包含在本文件夹中）|
| `build-content-analysis.js` | 被两个 JS 脚本共同依赖（已包含在本文件夹中）|

安装 Python 依赖：
```bash
pip install openpyxl requests
```

---

## 关于「无效」链接

图集和视频均失败时标记为无效，常见原因：
- 视频已被作者删除
- 视频设置了「仅粉丝可见」或「私密」
- 抖音接口限制（无公开播放地址）

---

## 输入文件要求

### Excel（`.xlsx`）

输入文件需包含以下列名（表头在第 1 行）：

| 列名 | 必须 | 说明 |
|------|------|------|
| 来源渠道 | ✅ | 脚本筛选值为「抖音 APP」 |
| 原文URL | ✅ | 抖音视频链接 |
| 链接是否有效 | ✅ | 脚本运行后自动回填「是」/「否」|
| case ID | — | 可为空，填写后写入结果表格 |
| 标题 | — | 可为空，写入结果表格 |

> 列名必须一致，列的顺序无要求。运行后会在原 Excel 末尾**自动新增**「媒体类型」列（图集 / 视频）。

### JSON（`.json`）

JSON 格式为对象数组，每条记录需包含以下字段：

| 字段名 | 必须 | 说明 |
|--------|------|------|
| 来源渠道 | ✅ | 脚本筛选值为「抖音 APP」 |
| 原文URL | ✅ | 抖音视频链接 |
| case ID | — | 可为空 |
| 标题 | — | 可为空 |

下载完成后，脚本会**原地回写**原 JSON 文件，为每条匹配记录追加 `链接是否有效` 和 `媒体类型` 两个字段。
