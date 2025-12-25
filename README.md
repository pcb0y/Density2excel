# Density2excel

一个用于“密度检测仪（串口输出）→ 自动提取密度值 → 写入 Excel”的小工具，提供 GUI 界面，支持按产品型号批量检测并回写结果到指定的 `xlsx` 文件。

## 功能

- 从串口读取原始数据，并提取密度值（优先匹配 `Density : 1.329`，否则尝试提取第一个浮点数）
- 每个产品默认采集 5 次密度值并计算平均值
- 从 Excel 读取待测产品列表（产品型号、机台号、来样时间、班次）
- 将检测结果回写到 Excel 对应产品行（检测时间、密度1~5、平均值）
- GUI 支持选择 Excel 文件、查看原始串口数据、查看每次检测结果与日志、配置串口参数并保存到 `config.ini`

## 环境依赖

- Python 3.x
- 第三方库：
  - `pyserial`
  - `openpyxl`
- `tkinter`：Python 自带（Windows 通常默认包含）

安装依赖：

```bash
pip install pyserial openpyxl
```

## 快速开始（GUI）

1. 准备 Excel 模板（或使用已有的 `density_data.xlsx`）
   - 你可以运行脚本生成示例文件：
     ```bash
     python create_test_excel.py
     ```
2. 连接密度仪到电脑，确认串口号（例如 `COM2` / `COM3`）
3. 启动 GUI：
   ```bash
   python main.py
   ```
4. 在界面中选择 Excel 文件（`*.xlsx`），加载产品列表
5. 根据设备设置串口参数（右上角“串口配置”），必要时点“保存配置”
6. 点击“开始检测”，程序会对当前产品进行 5 次检测并回写 Excel

## Excel 文件格式

程序默认读取工作表的前 5 列作为产品信息，并回写第 2 列和第 6~11 列的检测结果。表头建议如下（与 `create_test_excel.py` 一致）：

| 列 | 字段 |
|---|---|
| A | 来样时间 |
| B | 检测时间 |
| C | 机台号 |
| D | 产品型号 |
| E | 班次 |
| F~J | 密度1~密度5 |
| K | 平均值 |

注意：回写时以“产品型号”（第 D 列）作为匹配键。

## 串口配置（config.ini）

程序会读取/保存 `config.ini` 中的串口参数，示例：

```ini
[SerialConfig]
port = COM2
baudrate = 9600
bytesize = 7
stopbits = 1
parity = NONE
timeout = 2
max_attempts = 10
```

## 常见问题

- 读取不到密度值
  - 确认串口号与波特率是否正确
  - 确认设备输出中包含 `Density` 行（或至少包含可识别的浮点数）
- Excel 没有回写
  - 确认 Excel 中“产品型号”与界面显示一致（回写按产品型号匹配）
  - 确认 Excel 文件未被其它程序以独占方式占用

## 项目文件

- `main.py`：主程序（GUI + 串口读取 + Excel 读写）
- `config.ini`：串口配置
- `create_test_excel.py`：生成示例 `density_data.xlsx`
- `check_excel.py` / `check_result.py`：辅助检查 Excel 内容

