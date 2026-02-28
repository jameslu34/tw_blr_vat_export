# 臺灣營業稅BLR網路申報匯出模組（Odoo 19）

**臺灣營業稅BLR網路申報匯出模組**：從 Odoo 的會計資料產生營業稅(BLR)媒體申報匯入檔（TXT / TET_U），並提供下載、壓縮打包功能。

## 功能特點
- 公司層級欄位：設定營業人資料（統一編號、申報相關欄位）
- 發票/ 分錄欄位：補充 BLR 匯出所需資料
- 匯出精靈：依申報期間產出 TET_U 並可打包 ZIP

## 支援版本
- Odoo 19.0 Community / Enterprise

## 安裝與使用
1. 將 `tw_blr_vat_export/` 放到你的 Odoo `addons_path`
2. 重啟 Odoo，在 Apps 搜尋並安裝本模組
3. 於公司設定及發票/分錄中填入必要資料，使用匯出精靈選擇期間後匯出

## 授權
LGPL-3.0-or-later

---

# Taiwan VAT BLR Online Filing Export Module (Odoo 19)

**Taiwan VAT BLR Online Filing Export Module**: generate VAT (BLR) media filing import file (TXT/TET_U) from Odoo accounting data and provide downloadable zip.

## Features
- Company-level fields: configure VAT ID and related fields
- Invoice/Journal entry fields: add BLR-specific data
- Export wizard: produce TET_U file for a filing period, optionally zipped

## Supported Versions
- Odoo 19.0 Community / Enterprise

## Installation & Usage
1. Place `tw_blr_vat_export/` in your Odoo `addons_path`
2. Restart Odoo and install the module from Apps
3. Populate company and invoice fields, then export via wizard

## License
LGPL-3.0-or-later
