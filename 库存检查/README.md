# ShowZ Store 每日库存自动检查脚本

自动登录 ShowZ Store 管理后台，扫描 APC Toys / Iron Factory / Gear Factory 三个品牌的所有产品，对库存为 0 的产品按规则补库存或开启销售检查，并通过 Telegram Bot 发送汇报。

每天定时运行五次（03:00 / 09:00 / 12:00 / 17:00 / 22:00），由本机 Windows 任务计划程序触发，无需人工干预。

---

## 功能说明

- 遍历三个品牌的所有产品页面（自动翻页到底）
- 库存为 0 时，按以下规则自动处理：
  - **In Stock**（无前缀）+ 销售检查未开启 → 触发销售检查 AJAX
  - **Pre-Order / Coming Soon** + SKU 在 Excel 中 → 按 Excel「剩余可加库存」列补库存
  - **Pre-Order** + SKU 不在 Excel → 补库存 20
  - **Coming Soon** + SKU 不在 Excel → 跳过
- Telegram 通知内容：
  - 📦 需要关注的产品（库存=0 且销售检查未开启）
  - 🔧 今日改动（本次补库存 / 开启销售检查的产品）

---

## 依赖安装

```bash
pip install playwright openpyxl requests
python -m playwright install chromium
```

---

## 使用前需要修改的地方

打开 `inventory_check.py`，找到顶部配置区：

```python
# ── Configuration ─────────────────────────────────────────────────────────────
MANAGE_BASE = "https://showzstore.com/manage/"
USERNAME    = "chantia@showz.store"        # 管理后台账号
PASSWORD    = "SS27650942"                 # 管理后台密码
BRANDS      = ["APC Toys", "Iron Factory", "Gear Factory"]
EXCEL_PATH  = r"C:\Users\XuQian\Desktop\产品报数文档.xlsx"  # ← 改成你的 Excel 路径
TG_TOKEN    = "8841015387:AAEJUhOZDKgHp84GZ0NwujqI-e-2Ao5Q71I"
TG_CHAT_ID  = "8965386696"                 # ← 改成你的 Telegram Chat ID
LOG_FILE    = r"C:\Users\XuQian\秘书\库存检查\inventory_check.log"
```

必须修改的项：

| 配置项 | 说明 |
|--------|------|
| `USERNAME` / `PASSWORD` | ShowZ Store 管理后台登录凭据 |
| `EXCEL_PATH` | 产品报数 Excel 文件的完整路径 |
| `TG_CHAT_ID` | 你自己的 Telegram Chat ID（发消息给 [@userinfobot](https://t.me/userinfobot) 可获取） |
| `TG_TOKEN` | Telegram Bot Token（从 [@BotFather](https://t.me/BotFather) 创建） |
| `LOG_FILE` | 日志文件保存路径（可按需修改） |

### Excel 格式要求

每个品牌一个 Sheet，至少包含以下列（列名可含中文关键词）：

| 列名 | 说明 |
|------|------|
| `SKU` / `编号` / `型号` | 产品 SKU |
| `剩余可加库存` / `可加库存` | 本次可补的库存数量 |

---

## 手动运行一次测试

```bash
cd C:\Users\XuQian\秘书\库存检查
python inventory_check.py
```

或直接双击 `run_inventory_check.bat`（日志会追加写入 `inventory_check.log`）。

---

## 设置 Windows 定时任务

以下命令以**管理员身份**在 PowerShell 中运行，创建每天 5 个触发时间的任务（03:00 / 09:00 / 12:00 / 17:00 / 22:00）：

```powershell
$bat = "C:\Users\XuQian\秘书\库存检查\run_inventory_check.bat"
foreach ($t in "03:00","09:00","12:00","17:00","22:00") {
    $action  = New-ScheduledTaskAction -Execute $bat
    $trigger = New-ScheduledTaskTrigger -Daily -At $t
    $name    = "ShowZ_InventoryCheck_" + ($t -replace ':','')
    Register-ScheduledTask -TaskName $name -Action $action -Trigger $trigger -RunLevel Highest -Force
}
```

创建后可在「任务计划程序」中验证，或手动右键→「运行」测试。

---

## 文件说明

| 文件 | 说明 |
|------|------|
| `inventory_check.py` | 主脚本，包含所有业务逻辑 |
| `run_inventory_check.bat` | 批处理包装层，供任务计划程序调用 |
| `inventory_check.log` | 运行日志（自动生成，不含敏感信息） |
