# hanzi-cloud — 中文文档批量词频分析工具

一款基于 **PySide6 + jieba + WordCloud** 的桌面应用，可批量导入 Word 文档（`.docx`），在后台线程中完成中文分词、词频统计，并实时生成可视化词云图。

---

## 功能特性

- 📂 **批量导入**：支持一次性选择多个 `.docx` 文档
- ✂️ **中文分词**：使用 [jieba](https://github.com/fxsjy/jieba) 进行精准分词，自动过滤停用词
- 📊 **词频统计**：提取 Top 100 高频词并以表格形式展示
- ☁️ **词云生成**：利用 WordCloud 生成高清词云图，嵌入界面右侧展示
- 🧵 **非阻塞处理**：分词与渲染全程在 `QThread` 子线程执行，GUI 保持流畅响应
- 📈 **进度反馈**：进度条实时显示文档解析与词云生成进度
- 📤 **导出 Excel**：一键将 Top 100 词频结果导出为带样式的 `.xlsx` 文件
- 📝 **停用词可配置**：通过界面对话框或直接编辑 `stopwords.txt` 自定义过滤词。首次启动自动生成默认文件

---

## 环境要求

| 项目     | 版本                                  |
| -------- | ------------------------------------- |
| Python   | ≥ 3.11                                |
| 操作系统 | Windows 10 / 11（中文字体路径已适配） |

> **字体说明**：词云渲染默认使用 Windows 系统字体 `simhei.ttf`（黑体），请确保系统已安装该字体（通常位于 `C:\Windows\Fonts\simhei.ttf`）。

---

## 安装与运行

本项目使用 [uv](https://github.com/astral-sh/uv) 进行包管理。

### 1. 克隆项目

```bash
git clone <your-repo-url>
cd chinese_wordcloud
```

### 2. 安装依赖

```bash
uv sync
```

### 3. 运行程序

```bash
uv run main.py
```

---

## 依赖清单

| 包名          | 用途                                          |
| ------------- | --------------------------------------------- |
| `pyside6`     | GUI 框架（Qt for Python）                     |
| `jieba`       | 中文分词                                      |
| `python-docx` | 读取 `.docx` 文档                             |
| `wordcloud`   | 词云图生成                                    |
| `matplotlib`  | 词云图渲染（使用非交互式 Agg 后端，线程安全） |
| `openpyxl`    | 写入带样式的 Excel 文件（`.xlsx`）            |

---

## 项目结构

```
hanzi-cloud/
├── main.py          # 主程序（GUI 、后台处理、导出逻辑）
├── stopwords.py     # 停用词管理模块（加载 / 保存）
├── pyproject.toml   # 项目配置与依赖声明
├── uv.lock          # 依赖锁文件
└── README.md        # 本文档
```

> `stopwords.txt` 由程序自动生成，已列入 `.gitignore`，不跟随仓库提交。

---

## 使用说明

1. 点击 **"选择 Word 文档 (.docx)"** 按钮，选择一个或多个文档
2. （可选）点击 **"管理停用词"** ，在对话框中增删过滤词并保存
3. 点击 **"开始分析"** 启动后台处理
4. 等待进度条完成后，左侧表格显示 Top 100 词频，右侧展示词云图
5. 点击 **"导出 Excel"** 按钮，选择保存路径，即可生成带蓝色表头样式的 `.xlsx` 文件

---

## License

MIT
