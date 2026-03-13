# 订单汇总生成器

从销售订单明细（Sheet A）自动生成按「产品线 x 国家」汇总的交叉表（Sheet B）。

## 功能

- 读取 Excel 文件中的 Sheet A 订单数据
- 按产品线和国家自动分类汇总金额（以千为单位）
- 生成格式化的 Sheet B 汇总表，包含 SUM 公式
- 特殊规则处理：CIS 国家归类、Parts 订单归入 Service
- 完整 GUI 界面，实时显示处理日志
- 支持自动检查更新，从 GitHub Releases 下载新版本

## 使用方法

### 直接运行

双击 `订单汇总生成器.exe`，无需安装 Python 环境。

1. 点击「浏览」选择输入的 xlsx 文件（需包含 Sheet A）
2. 输出路径会自动生成，也可手动修改
3. 点击「生成」，等待处理完成

### 命令行

```bash
python 生成订单汇总.py 输入文件.xlsx [输出文件.xlsx]
```

## Sheet A 输入格式

| 列 | 字段 | 说明 |
|---|---|---|
| L | Country Code Description | 国家 |
| O | Sales Analysis 3 | 产品线 |
| Q | Sales Analysis 5 | 值为 "Parts" 时归入 Service |
| W | Outstanding Value In USD | 汇总金额 |

## Sheet B 输出格式

- **列**：India, Singapore, Vietnam, Philippines, Indonesia, Thailand, Malaysia, Japan, Korea, New Zealand, Australia, CIS Countries, Other, APAC(Total)
- **行**：三大产品组（Flexible Packaging / Product Integrity & Material Test / Food & Beverage）下的各产品线，加 3rd party、Service、Total
- 所有金额除以 1000，以千为单位显示

## 自动更新

应用启动时会后台检查 GitHub Releases 是否有新版本，发现新版后提示下载更新并自动重启。也可以手动点击界面右上角的「检查更新」按钮。

## 技术栈

- Python 3 + openpyxl
- tkinter（GUI）
- PyInstaller（打包 exe）
- GitHub Releases（自动更新）
