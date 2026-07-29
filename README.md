# SAP BTP Integration Hub — Frontend (Fiori Apps)

> **Đồ án tốt nghiệp** — SAP UI5 Fiori Frontend cho hệ thống tích hợp tải lên chứng từ SAP hàng loạt từ Excel, triển khai trên SAP Business Technology Platform.

[![FE Repo](https://img.shields.io/badge/repo-Frontend%20Fiori-2563EB?logo=sap)](https://github.com/tata-nguyen-BA/SAP490_SU26SAP03_GSU26SAP03_FE)
[![BE Repo](https://img.shields.io/badge/repo-Backend%20ABAP-0553CE?logo=sap)](https://github.com/tata-nguyen-BA/SAP490_SU26SAP03_GSU26SAP03)

---

## Liên kết Repository

| Repo | Mô tả | Link |
|------|-------|------|
| **Frontend (repo này)** | SAP UI5 Fiori apps — zfi_pk_zup, zup_rpt | [SAP490_SU26SAP03_GSU26SAP03_FE](https://github.com/tata-nguyen-BA/SAP490_SU26SAP03_GSU26SAP03_FE) |
| **Backend** | ABAP objects — ZFI, ZPP, ZIH_POGR | [SAP490_SU26SAP03_GSU26SAP03](https://github.com/tata-nguyen-BA/SAP490_SU26SAP03_GSU26SAP03) |

---

## Tổng quan

Hai Fiori apps triển khai trên SAP BTP, kết nối với SAP S/4HANA On-Premise qua OData V4:

| App | Chức năng chính |
|-----|----------------|
| **zfi_pk_zup** | Upload Hub — tải lên chứng từ FI, PP, GR từ Excel |
| **zup_rpt** | Analytics Hub — báo cáo, KPI, biểu đồ cho FI, PP, GR |

---

## Kiến trúc kết nối

```
┌───────────────────────────────────────────────────────┐
│              SAP Business Technology Platform          │
│                                                        │
│  ┌──────────────────────┐  ┌─────────────────────────┐ │
│  │  zfi_pk_zup           │  │  zup_rpt                │ │
│  │  Upload Hub           │  │  Analytics Hub          │ │
│  │                       │  │                         │ │
│  │  Tab FI  → mainService│  │  FiRpt  → mainService   │ │
│  │  Tab PP  → pp model   │  │  PpRpt  → pp model      │ │
│  │  Tab GR  → gr model   │  │  GrRpt  → gr model      │ │
│  └──────────┬────────────┘  └────────────┬────────────┘ │
│             │  OData V4                   │              │
│    BTP Connectivity Destination: S4H_BACKEND            │
└─────────────┼───────────────────────────┼──────────────┘
              │                           │
┌─────────────▼───────────────────────────▼──────────────┐
│                 SAP S/4HANA On-Premise                  │
│    ZFI_UI_ZUP_FIDOC_O4  ·  ZPP_UI_ZUPLSX_O4            │
│    ZMM_UI_POGR_O4  (GR Upload + History + KPI)          │
└─────────────────────────────────────────────────────────┘
```

---

## Cấu trúc Repository

```
📦 SAP490_SU26SAP03_GSU26SAP03_FE/
 ├── zfi_pk_zup/                        # Upload Hub
 │   └── webapp/
 │       ├── manifest.json              # data sources (mainService FI + pp + gr) + models + routes
 │       ├── view/
 │       │   ├── App.view.xml
 │       │   ├── Main.view.xml          # Landing page — chọn loại chứng từ
 │       │   ├── Fi.view.xml            # FI Upload (88-col table + history)
 │       │   ├── Pp.view.xml            # PP Upload + history
 │       │   ├── Gr.view.xml            # GR Upload + history (async APJ)
 │       │   └── fragment/
 │       │       ├── Busy.fragment.xml
 │       │       └── GrItemDetail.fragment.xml  # Dialog chi tiết item của 1 GR
 │       ├── controller/
 │       │   ├── App.controller.js
 │       │   ├── Main.controller.js
 │       │   ├── Fi.controller.js       # FI upload logic + SheetJS + _loadAuth
 │       │   ├── Pp.controller.js       # PP upload logic + SheetJS + _loadAuth
 │       │   ├── Gr.controller.js       # GR upload logic + group by GR# + _loadAuth + polling
 │       │   ├── xlsx/                  # xlsx.js / xlsx.bundle.js (SheetJS, vendor)
 │       │   └── helper/
 │       │       ├── ApiService.js      # OData V4 bound action calls + loadMyAuth/getUserEmail
 │       │       ├── ColumnConfig.js    # 88-col config cho FI
 │       │       ├── ColumnSettingsDialog.js  # Toggle "All Field" cho bảng FI
 │       │       ├── ErrorDialog.js     # Hiển thị message lỗi SAP/BAPI
 │       │       ├── ExcelParser.js     # Parse workbook chung (FI/PP)
 │       │       ├── ExcelTemplate.js   # Sinh template Excel FI
 │       │       ├── GrExcelTemplate.js # Sinh template Excel GR
 │       │       ├── GrHistoryExport.js # Export tab Lịch sử GR ra Excel
 │       │       ├── HistoryDetailDialog.js   # Dialog chi tiết 1 dòng lịch sử FI/PP
 │       │       ├── HistoryFileExport.js     # Export lịch sử FI/PP ra Excel
 │       │       ├── P13nHelper.js      # Personalization (ẩn/hiện, sắp cột)
 │       │       ├── PpExcelTemplate.js # Sinh template Excel PP
 │       │       ├── ResultsDialog.js   # Dialog kết quả Check/Upload
 │       │       └── UploadValidator.js # Validate file trước khi gửi lên OData
 │       ├── model/
 │       ├── i18n/
 │       ├── css/
 │       └── test/                      # OPA + QUnit scaffolding (mặc định UI5, chưa viết test case riêng)
 └── zup_rpt/                           # Analytics Hub
     └── webapp/
         ├── manifest.json              # mainService (ZUP_UI_RPT_O4) + grService (ZMM_UI_POGR_O4) + routes
         ├── view/
         │   ├── App.view.xml
         │   ├── Main.view.xml          # Dashboard — KPI tiles + nav cards
         │   ├── FiRpt.view.xml         # FI Analytics: filter + KPI + chart + table
         │   ├── PpRpt.view.xml         # PP Analytics
         │   ├── GrRpt.view.xml         # GR Analytics (7 cột + StatusCriticality)
         │   └── KpiRpt.view.xml        # KPI tổng hợp GR+FI+PP (route "kpirpt")
         ├── controller/
         │   ├── Main.controller.js
         │   ├── FiRpt.controller.js
         │   ├── PpRpt.controller.js
         │   ├── GrRpt.controller.js
         │   ├── KpiRpt.controller.js   # đọc entity UploadKPI, tính successRate/errorRate/avgSeconds
         │   ├── xlsx/                  # xlsx.js / xlsx.bundle.js (SheetJS, vendor)
         │   └── helper/
         │       ├── Formatter.js       # criticalityToState, statusText
         │       ├── SimpleChart.js     # SVG chart tự vẽ
         │       └── ExcelExport.js
         ├── model/
         ├── i18n/
         └── css/
```

---

## Ứng dụng 1: `zfi_pk_zup` — Upload Hub

### Chạy local

```bash
cd zfi_pk_zup
npm install
npm start
# → http://localhost:8080
```

### Các Tab

#### Tab FI — Upload Chứng từ Kế toán

- FileUploader chọn `.xlsx/.xls/.csv`
- SheetJS (`xlsx.bundle.js`) parse client-side → 88 cột
- Cột cấu hình trong `ColumnConfig.js` — mặc định hiển thị 13 cột chính, toggle "All Field" để xem đủ 88
- **Check**: gọi `uploadExcel(testmode=true)` → validate, hiển thị lỗi từng dòng
- **Post**: gọi `uploadExcel(testmode=false)` → post thật, kết quả per-row đồng bộ
- Tab "Lịch sử": binding `{path: '/UploadHistory', parameters: {'$orderby': 'pst_date desc'}}`

#### Tab PP — Upload Lệnh Sản xuất

- Validate ngày DD/MM/YYYY → convert YYYYMMDD trước khi gọi OData
- Retry từng dòng lỗi riêng lẻ
- Kết quả: Production Order number + status từng dòng

#### Tab GR — Upload Phiếu Nhập kho

- `_processRows`: normalize header keys (trim/toLowerCase), group dòng theo `gr_number`
- `_buildGrDocs()` trong `ApiService.js`: tạo `{grnumber, documentdate, movementtype, items: [...]}`
- Gọi Action `GrUpload/uploadExcel` với `payload_json`, `mapping_id="POGR001"`, `testmode`
- **Result**: `[1]` summary — hiển thị Panel `batch_id / total / success / error`
- APJ background job chạy async → user refresh Tab "Lịch sử GR" để xem `material_document`
- History binding: `{path: 'gr>/GrUpload', parameters: {'$orderby': 'CreatedAt desc'}}`
- Status hiển thị bằng `ObjectStatus` với `criticalityToState` formatter

### Authorization (cả 3 tab FI/PP/GR)

Mỗi controller (`Fi.controller.js`, `Pp.controller.js`, `Gr.controller.js`) gọi
`_loadAuth()` trong `onInit`, dùng `ApiService.loadMyAuth()` để lấy
`can_upload_fi/pp/gr` (action `getMyAuth` đặt trên service GR nhưng trả quyền của
cả 3 phân hệ). Nếu `can_upload_<module> === false`, ẩn nút Check và Post của tab
đó. Email dùng để check lấy từ `ApiService.getUserEmail()` — đọc
`sap.ushell.Container` `UserInfo` service (SSO thật của Work Zone), cache lại sau
lần gọi đầu; chạy ngoài launchpad (ví dụ `npm start` local) trả email rỗng và
backend sẽ từ chối kèm lý do. Đây chỉ là gate ở UI (ẩn nút) — quyền thật chặn ở
backend khi gọi action `uploadFromExcel`/`uploadExcel`/`retryPost` (xem
README_BE.md, mục Authorization).

### OData V4 Action Pattern (GR)

```javascript
// ApiService.js
const ACTION_FQN = "com.sap.gateway.srvd.zmm_ui_pogr_o4.v0001.uploadExcel";
const oOp = oGrModel.bindContext("/GrUpload/" + ACTION_FQN + "(...)");
oOp.setParameter("payload_json", JSON.stringify({ filename, doc: aDocs }));
oOp.setParameter("mapping_id",  "POGR001");
oOp.setParameter("testmode",    bTestMode);
await oOp.execute();
const oResult = oOp.getBoundContext().getObject() || {};
```

---

## Ứng dụng 2: `zup_rpt` — Analytics Hub

### Chạy local

```bash
cd zup_rpt
npm install
npm start
# → http://localhost:8080
```

### Các trang phân tích

4 trang: `FiRpt` / `PpRpt` / `GrRpt` (báo cáo riêng từng module) và `KpiRpt`
(route `kpirpt`, dashboard **tổng hợp cả 3 module**, KPI thứ 4 trong app này).

`KpiRpt` đọc entity `UploadKPI` của service `ZUP_UI_RPT_O4` (CDS
`ZUP_C_UPLOAD_KPI`, `UNION ALL` GR+FI+PP+bảng lỗi) — filter theo khoảng ngày/loại
chứng từ/status/created by, tính `total`, `success`, `error`, `pending`,
`successRate`, `errorRate`, và **`avgSeconds`** (thời gian xử lý trung bình, chỉ
tính được cho GR vì chỉ GR có mốc thời gian tạo/hoàn tất qua job nền — FI/PP post
đồng bộ nên không có khoảng cách thời gian đáng kể để đo). Đây chính là KPI
"thời gian xử lý" theo yêu cầu #7 của đề bài. Có chart theo dimension (DocType/
Status/Month/CreatedBy) và export Excel (Summary + Data, có ghi rõ filter đang áp
dụng).

Mỗi trang FiRpt / PpRpt / GrRpt có cấu trúc giống nhau:

```
Filter Bar
  └── DateRange (CreatedAt) · Status Select · CreatedBy Input · [Go] [Clear]

KPI Row (tính client-side trên tập đang lọc)
  └── Total · Success (valueColor: Good) · Error (valueColor: Critical) · Pending (Neutral)

SimpleChart
  └── Dimension selector: by Status / by Month / by User
  └── SVG bar chart tự vẽ

Table
  └── Columns: GrNumber · Status (ObjectStatus + criticalityToState) · DocumentDate
               MaterialDocument · Message · CreatedBy · CreatedAt
  └── Toolbar: "Xuất Excel" → ExcelExport.export() (sheet Summary + Data)
```

### Formatter

```javascript
// helper/Formatter.js
criticalityToState(iCrit) {
    const map = { "3": "Success", "1": "Error", "2": "Warning" };
    return map[String(iCrit)] || "None";
}
```

### Filter pattern (GrRpt)

```javascript
_buildFilters() {
    const aFilters = [];
    if (dFrom && dTo) {
        aFilters.push(new Filter("CreatedAt", FilterOperator.BT,
            dFrom.toISOString().replace("T00:00:00.000Z","T00:00:00Z"),
            dTo.toISOString().replace("T00:00:00.000Z","T23:59:59Z")));
    }
    if (sStatus) aFilters.push(new Filter("Status", FilterOperator.EQ, sStatus));
    if (sUser)   aFilters.push(new Filter("CreatedBy", FilterOperator.EQ, sUser));
    return aFilters;
}
```

---

## Deploy lên BTP & SAP Work Zone

### Bước 1 — Build

```bash
cd zfi_pk_zup && npm run build
cd zup_rpt    && npm run build
```

### Bước 2 — Login BTP Cloud Foundry

```bash
cf login -a https://api.cf.<region>.hana.ondemand.com \
  -o <your-org> -s <your-space>
```

### Bước 3 — Deploy

```bash
# cf push
cf push zfi-pk-zup -f zfi_pk_zup/manifest.yml
cf push zup-rpt    -f zup_rpt/manifest.yml

# Hoặc MTA
mbt build
cf deploy mta_archives/*.mtar
```

> Nếu dùng **SAP Fiori Tools** trong VS Code: `Fiori: Deploy Application` → BTP target → space.

### Bước 4 — SAP Work Zone

1. BTP Cockpit → **SAP Work Zone** (hoặc Launchpad Service)
2. **Provider Manager** → refresh HTML5 Apps
3. **Content Manager** → **Content Explorer** → tìm `zfi_pk_zup` và `zup_rpt`
4. Thêm vào **My Content** → gán **Role**
5. **Site Manager** → thêm vào **Group** → **Publish**

### Bước 5 — Cấu hình Destination (BTP Cockpit)

| Property | Giá trị |
|----------|---------|
| Name | `S4H_BACKEND` |
| Type | HTTP |
| URL | `https://<your-s4hana-host>:<port>` |
| Authentication | BasicAuthentication hoặc PrincipalPropagation |
| `sap-client` | `<client-number>` |
| `WebIDEEnabled` | `true` |
| `HTML5.DynamicDestination` | `true` |

---

## Công nghệ sử dụng

| Layer | Technology |
|-------|------------|
| Framework | SAP UI5 (Freestyle) — `sap.m`, `sap.ui.table`, `sap.ui.layout` |
| OData | OData V4 Model — bound Actions, `requestContexts`, `$orderby` |
| Excel Parse | SheetJS (`xlsx.bundle.js`) — client-side |
| Chart | SVG chart tự vẽ qua `SimpleChart.js` |
| Export | `ExcelExport.js` — XLSX client-side (Summary + Data sheet) |
| Deploy | SAP BTP Cloud Foundry, HTML5 Application Repository |
| Launchpad | SAP Work Zone Standard (Launchpad Service) |

---

## Yêu cầu môi trường

```bash
node --version   # >= 18
npm install -g @sap/ux-ui5-tooling
```

- SAP BTP account (trial hoặc enterprise)
- Cloud Foundry environment với org + space
- HTML5 Application Repository service instance
- Connectivity + Destination service instance
- SAP Work Zone Standard hoặc Advanced

---

## Bảo mật & `.gitignore`

Không commit vào repo này:

```
.env
node_modules/
dist/
.fiorirc           # có thể chứa credentials BTP
ui5-local.yaml     # local proxy config
```

Không đưa vào file tracked: username/password BTP thật, tenant URL thật, API key thật.

---

## Backend Repository

ABAP Backend (OData V4 services, BAPI, APJ Job) nằm ở repo riêng:

**[SAP490_SU26SAP03_GSU26SAP03](https://github.com/tata-nguyen-BA/SAP490_SU26SAP03_GSU26SAP03)**

Package gốc `ZPK_GSU26SAP03` (chứa luôn class nghiệp vụ FI) + 4 sub-package:
- `ZFI_PK_FIDOC_LEGACY` — chỉ còn HTTP wrapper cho SOAP post của FI
- `ZPP_PK_ZUPLSX` — PP Production Order
- `ZIH_POGR` — MM Goods Receipt + APJ Background Job (module chính)
- `ZPK_ZUP_RPT` — CDS báo cáo/KPI dùng chung cho app `zup_rpt` ở trên

---

## Tác giả

**Đồ án tốt nghiệp — Hệ thống Tích hợp SAP BTP**

- Nhóm: SU26SAP03_GSU26SAP03
- GVHD: Nguyễn Thị Cẩm Hương
- Trường: Đại học FPT
- Năm: 2026
