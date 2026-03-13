# Weak Short 批量审核工具

从面板/芯片测试结果文件中，自动提取 **Weak Short-Circuit Test** 的 NG/OK 状态及关键数据，批量归类并生成 Excel 报表。替代人工逐文件比对，提升审核效率。

## 功能

- **批量解析**：支持 `.txt`、`.docx` 格式的测试结果文件
- **智能提取**：自动定位 `Weak Short-Circuit Test` 区段，正则匹配 NG/OK 结果
- **数据提取**：NG 文件自动提取 `Mul Short` 和 `Mutual Short` 下的测试数据
- **文件归类**：NG / OK 文件自动分拣到对应文件夹
- **Excel 报表**：生成 `output.xlsx`，Sheet1 为 NG 结果（含详细数据），Sheet2 为 OK 结果
- **跨平台**：支持 macOS 和 Windows

## 输出结构

```
桌面/审核结果_2026-03-13_150000/
├── NG/              # 失败的测试文件
├── OK/              # 通过的测试文件
└── output.xlsx      # 汇总报表
```

### output.xlsx 格式

**Sheet1（NG）：**

| 文件名称 | 结果1（Mul Short:） | 结果2（Mutual Short:） |
|---------|-------------------|---------------------|
| xxx_Fail.txt | Tx10: 46.63(kΩ) / Rx16: 46.63(kΩ) | Tx10 with Rx16: 46.63(kΩ) |

**Sheet2（OK）：**

| 文件名称 |
|---------|
| xxx_Pass.txt |

## 提取规则

1. 定位 `======Test Item: -------- Weak Short-Circuit Test` 到 `======Test Item: -------- INT Pin Test` 之间的内容
2. 正则匹配 `Weak Short` + `NG` → 失败；`Weak Short` + `OK` → 通过
3. NG 时提取 `Mul Short:` 和 `Mutual Short:` 下的数据行

## 快速开始

### 安装依赖

```bash
pip3 install -r requirements.txt
```

### 运行

```bash
python3 main.py
```

### 打包

**macOS：**

```bash
./build_mac.sh
# 输出: dist/WeakShortTool.app
```

**Windows：**

```bat
build_win.bat
# 输出: dist\WeakShortTool.exe
```

> 注意：需要在对应系统上分别打包，不支持跨平台编译。

## 技术栈

- **Python 3.9+**
- **pywebview** — 原生 WebView 渲染 GUI（macOS 用 WKWebView，Windows 用 Edge WebView2）
- **openpyxl** — Excel 读写
- **python-docx** — Word 文件解析
- **PyInstaller** — 打包为独立可执行文件

## 项目结构

```
weak-short-tool/
├── main.py              # 入口 + pywebview GUI
├── core/
│   ├── parser.py        # 文件解析 & 正则提取
│   ├── exporter.py      # Excel 输出
│   └── classifier.py    # 文件归类
├── requirements.txt
├── build_mac.sh         # macOS 打包脚本
└── build_win.bat        # Windows 打包脚本
```
