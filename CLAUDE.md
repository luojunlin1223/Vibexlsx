# Vibexlsx - 订单汇总生成器

## 项目概述

将销售订单明细（Sheet A）自动汇总为按「产品线 × 国家」的交叉表（Sheet B）。
带 tkinter GUI 界面，可打包为独立 exe 分发给无 Python 环境的用户。

## 技术栈

- Python 3 + openpyxl（Excel 读写）
- tkinter（GUI）
- PyInstaller（打包 exe）

## 文件结构

- `生成订单汇总.py` — 主程序（业务逻辑 + GUI）
- `订单汇总生成器.exe` — 打包好的独立可执行文件

## 核心业务逻辑

### 输入：Sheet A 列映射

| 列 | 字段 | 用途 |
|---|---|---|
| L | Country Code Description | 国家 → 映射到 Sheet B 的列 |
| O | Sales Analysis 3 | 产品线 → 映射到 Sheet B 的行 |
| Q | Sales Analysis 5 | 若值为 `"Parts"` 则归入 Service 行 |
| W | Outstanding Value In USD | 汇总的数值（求和） |

### 输出：Sheet B 布局

- **列**：India, Singapore, Vietnam, Philippines, Indonesia, Thailand, Malaysia, Japan, Korea, New Zealand, Australia, CIS Countries, Other, APAC(Total)
- **行**：按三大产品组分组（Flexible Packaging / Product Integrity&Material Test / Food&Beverage），加 3rd party、Service、Total
- APAC(Total) 列和 Total 行使用 SUM 公式

### 特殊映射规则

- CIS 国家（Russia, Kazakhstan, Uzbekistan 等 12 国）→ "CIS Countries" 列
- 未匹配的国家 → "Other" 列
- Q 列为 "Parts" 的行 → 归入 "Service" 行（忽略原产品线）

## 开发注意事项

- Excel 文件必须用 `data_only=True` 加载，因为实际文件的 W 列可能是引用外部工作簿的公式
- 对 W 列值做类型检查，非数值（公式字符串等）跳过并记录警告
- GUI 耗时操作在子线程中执行（`threading.Thread`），通过 `root.after()` 回调更新 UI，避免界面阻塞

## 自动更新机制

- 基于 GitHub Releases，应用启动时后台检查 `https://api.github.com/repos/luojunlin1223/Vibexlsx/releases/latest`
- 比较本地 `VERSION` 与远程 `tag_name`，有新版则弹窗提示
- 更新流程：下载新 exe → 重命名旧 exe 为 `.old` → 放入新 exe → 自动重启
- Release 中的 exe 资产名称必须为 `订单汇总生成器.exe`

## 发版流程

1. 修改 `生成订单汇总.py` 中的 `VERSION`（如 `"1.1.0"` → `"1.2.0"`）
2. 打包 exe：
   ```bash
   python -m PyInstaller --onefile --windowed --name "订单汇总生成器" 生成订单汇总.py
   cp dist/订单汇总生成器.exe .
   rm -rf build/ dist/ 订单汇总生成器.spec
   ```
3. Commit & push：
   ```bash
   git add 生成订单汇总.py
   git commit -m "Bump version to vX.Y.Z"
   git push
   ```
4. 创建 GitHub Release 并上传 exe：
   ```bash
   gh release create vX.Y.Z 订单汇总生成器.exe --title "vX.Y.Z" --notes "更新说明"
   ```

注意：`gh` 需要 `repo` scope 权限（`gh auth refresh -s repo`）。
