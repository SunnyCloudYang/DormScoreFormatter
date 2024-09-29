# 卫生成绩表格式化工具

## 描述

DormScoreFormatter.py 是一个用于处理和格式化宿舍检查评分的 Python 脚本。它可以将下载的多个包含每周宿舍评分的 CSV 文件合并成一个格式化的 Excel 文件和可以直接打印的PDF文件。

## 功能

- 合并多个包含每周宿舍评分的 CSV 文件
- 去重和格式化数据
- 创建 Excel 文件，包括：
  - 包含宿舍楼号和周数的标题
  - 楼长联系信息
  - 空白单元格标红，便于检查修正
- 导出PDF文件，可以直接打印

## 要求

运行此脚本需要：

1. Python 3.6 或更高版本
2. 以下 Python 库：
   - pandas
   - openpyxl
   - win32com.client (仅在 Windows 上导出 PDF 时需要)

可以使用 pip 安装这些库：

```
pip install pandas openpyxl pywin32
```

## 使用方法

1. 将所有 `WeekScoreManage_*.csv` 文件放在同一个文件夹中。

2. 打开命令提示符或终端。

3. 导航到包含 DormScoreFormatter.py 脚本的文件夹。

4. 使用以下命令查看脚本的帮助信息和可选参数：

   ```
   python DormScoreFormatter.py --help
   ```

5. 使用以下命令运行脚本（根据自己需要添加参数）：

   ```
   python DormScoreFormatter.py --folder \path\to\csv\folder
   ```

   将 `\path\to\csv\folder` 替换为包含 CSV 文件的实际文件夹路径，如果不指定文件夹，则默认使用脚本所在的文件夹。

6. 脚本将处理文件并在脚本所在的文件夹中创建一个以宿舍楼号和周数命名的 Excel 文件（例如："紫荆公寓2号楼第1周.xlsx"）。

## 输出

生成的 Excel 文件将包含：

- 包含宿舍楼号和周数的标题
- 楼长联系信息
- 格式化的表格，包括房间号、床位号、总分和整改意见

如果成绩表中出现空白单元格，将在控制台中打印出空白单元格的位置，方便检查。

## 故障排除

如果遇到任何问题：

1. 确保已安装 Python 和所有必需的库。
2. 检查 CSV 文件格式是否正确，并且文件名正确（以 'WeekScoreManage_' 开头）。
3. 确保提供了正确的 CSV 文件夹路径。

如有其他问题，请[联系我](mailto:sunnycloudyang@outlook.com)。

## Python 新手注意事项

如果您是 Python 新手，可能需要先设置 Python 环境。以下是入门步骤：

1. 从 [python.org](https://www.python.org/downloads/) 下载并安装 Python。
2. 安装时，请确保勾选"Add Python to PATH"选项。
3. 安装完成后，打开命令提示符或终端，输入 `python --version` 以验证 Python 是否正确安装。
4. 使用前面提到的 pip 命令安装所需的库。

设置好 Python 环境并安装所需库后，您应该能够按照"使用方法"部分的说明运行脚本。
