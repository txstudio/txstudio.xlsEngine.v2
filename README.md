# txstudio.xlsEngine.v2

使用 NPOI 對既有 EXCEL 檔案進行操作的類別庫 (v2) -支援表頭與表尾

## 方案架構說明

> 此方案包含一個「類別庫專案」與「主控台應用程式專案」

專案名稱|描述
--|--
xlsEngine.v2|進行 EXCEL 操作的類別庫專案（使用NPOI）
xlsEngineApp|類別庫操作結果的主控台應用程式（範例程式碼）

### 類別庫簡短使用說明

xlsEngineProvider 為實作部分 EXCEL 操作會使用的程式碼區塊，要建立新的 EXCEL 報表時需繼承至此抽象類別。

實作 SetOption、InsertHeaderRow、InsertFooterRow 方法設定相關內容

### EXCEL 範本檔案建立項目

@ 與 # 為此類別庫保留字元
- [@] 開頭的儲存格可在實作項目中指定對應的數值
- [#] 開頭的儲存格文字會自動對應 DataRow 物件中資料欄位名稱對應的資料
> #Username => DataRow["Username"]

### 範例程式碼說明

範例程式碼會依照實作項目產生對應的 EXCEL 報表檔案，並在匯出後透過 CMD 指令直接開啟 EXCEL 檔案（請確認執行環境包含 EXCEL 應用程式）。

範例程式碼會建立兩種類型的 EXCEL 銷售報表內容：
- 依照訂購人員分隔設定表頭與表尾的 EXCEL
- 將整份資料匯出的 EXCEL，並設定指定欄位重複時不顯示

若要變更匯出的 EXCEL 請調整下面程式碼的註解內容
```
_xlsEngine = new SalesXlsEngine();

_xlsEngine = new SalesSimpleXlsEngine();
```

詳細實作項目與執行程式碼區塊請參考範例程式碼檔案

