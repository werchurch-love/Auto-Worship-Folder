# Auto-Worship-Folder

在 Windows 中快速建立一次崇拜所需的全部文件，並按播放次序放在同一個資料夾內。

## 好處

- 一鍵建立宣召、經訓、讀經、禱告、講道、奉獻、報告、祝福及詩歌等崇拜文件。
- 不用每星期手動複製大量範本，節省時間，亦減少漏檔、放錯檔案或播放次序混亂。
- 如臨時加入、刪除或改名 PowerPoint，只需執行 `refresh-master.vbs`，即可重新更新 `master.pptx`。

## 首次使用

### 1. 下載及解壓

1. 在本頁上方按 **Code**。
2. 按 **Download ZIP**。
3. 解壓縮下載的 ZIP 檔案。
4. 將整個 `Auto-Worship-Folder-main` 資料夾放到你想儲存的位置。

> 請不要隨意搬走或改名 `templates` 資料夾內的檔案，否則程式可能無法正常運作。

### 2. 設定詩歌資料夾捷徑

程式會按你輸入的歌名，到詩歌 PowerPoint 資料夾尋找同名檔案。

1. 在檔案總管開啟你存放詩歌 PowerPoint 的資料夾。
2. 按一下檔案總管上方的路徑列，然後複製完整路徑。
3. 開啟：

   ```text
   Auto-Worship-Folder-main\templates
   ```

4. 在 `songs-folder.lnk` 按右鍵，選 **內容**。
5. 在 **目標** 欄貼上剛才複製的詩歌資料夾路徑。
6. 按 **套用**，再按 **確定**。

> 請填寫資料夾路徑，不是單一 PowerPoint 檔案的路徑。

## 每次使用

### 1. 先更新經文內容

開啟並按原有格式更新以下檔案：

```text
內容-宣召及啟應
內容-經訓
內容-讀經
```

### 2. 建立本週崇拜文件夾

1. 雙擊：

   ```text
   create-worship-folder.vbs
   ```

2. 輸入崇拜名稱或日期，例如：

   ```text
   20260906
   ```

3. 輸入全部詩歌名稱，歌名之間用逗號分隔，例如：

   ```text
   歡欣, 生趣, 慈繩愛索, 十架的冠冕
   ```

4. 等候程式完成。

程式會：

- 在詩歌資料夾尋找與你輸入歌名**完全相同**的 PowerPoint 檔案。
- 找到時，複製該 PowerPoint 到本週崇拜資料夾。
- 找不到時，建立一個 `(找不到) 歌名.txt` 檔案作提示。

例如找不到「十架的冠冕」時，程式會建立：

```text
(找不到) 十架的冠冕.txt
```

完成時，畫面會先顯示：

```text
完成。master.pptx 已更新
```

按 **確定** 後，再顯示：

```text
完成
```

之後，程式所在位置會新增一個以你輸入名稱命名的資料夾，內有該次崇拜所需的全部文件。

## 更新 master.pptx

`master.pptx` 是把資料夾內的 PowerPoint 按檔名次序整合後的播放檔案。

如你在已建立的崇拜資料夾內做了以下改動，就要重新執行 `refresh-master.vbs`：

- 加入 PowerPoint 檔案
- 刪除 PowerPoint 檔案
- 更改 PowerPoint 檔名
- 更改播放次序

### 操作方法

1. 關閉所有 PowerPoint，尤其是 `master.pptx`。
2. 如有需要，先複製 `master.pptx` 作備份。
3. 在該次崇拜資料夾內，雙擊：

   ```text
   refresh-master.vbs
   ```

4. 等候程式完成。
5. 開啟 `master.pptx`，檢查投影片內容和播放次序。

> `refresh-master.vbs` 會重新讀取資料夾內的 PowerPoint，並更新／覆蓋 `master.pptx` 和 PowerPoint 播放檔案（`.ppsx`）。執行前請先關閉 PowerPoint。

## 自訂範本

你可以自行修改字體、顏色、背景、圖片和預設文字。

### 固定崇拜投影片

以下資料夾內的檔案，會原樣複製到每次新建立的崇拜資料夾：

```text
templates\worship-files
```

例如：安靜、禱告、奉獻、歡迎你、報告、三一頌和祝福。

要修改這些投影片的格式或內容，直接修改此資料夾內對應的檔案。

### 宣召、經訓及讀經

以下三個檔案是建立相關 PowerPoint 的範本：

```text
templates\宣召及啟應-template.pptx
templates\經訓-template.pptx
templates\讀經-template.pptx
```

你可以修改字體、顏色、背景和版面。

> 不要刪除範本內的佔位符，例如 `{{CALL_RESPONSE}}`。刪除後，程式可能無法正確建立投影片。

## VBS 與安全

VBS（VBScript）是 Windows 的自動化腳本檔。本工具主要使用：

```text
create-worship-folder.vbs
refresh-master.vbs
```

### 查看或修改 VBS

不要直接雙擊 `.vbs` 檔案來閱讀，因為 Windows 通常會立即執行。

1. 在 `.vbs` 檔案按右鍵。
2. 選 **開啟檔案方式**。
3. 選「記事本」、Notepad++ 或其他文字編輯器。
4. 修改前先備份原始檔。
5. 修改後先在測試資料夾試用，確認正常才正式使用。

### 用 AI 檢查安全性

如你不確定 script 是否安全：

1. 用記事本開啟 `create-worship-folder.vbs` 或 `refresh-master.vbs`。
2. 複製全部程式碼。
3. 貼到你信任的 AI，並輸入：

```text
請用繁體中文解釋這段 VBS 的功能。

請檢查它有沒有刪除、覆蓋、移動或複製檔案，
執行 CMD／PowerShell、下載資料、連接網絡、
讀取密碼或上傳資料。

請列出它會讀取、建立、修改或刪除的檔案和資料夾。

請不要修改程式碼，只報告實際功能和風險。
```

> 不要把密碼、API Key、個人資料、教會內部資料或完整電腦路徑貼到公開 AI 對話。

## 注意

- 本工具只適用於 Windows。
- 執行 `create-worship-folder.vbs` 或 `refresh-master.vbs` 前，請先關閉所有 PowerPoint。
- 程式會更新或覆蓋 `master.pptx`；重要檔案請先備份。
- 公開分享前，請檢查 `.vbs`、`.txt` 和 `.pptx` 內沒有密碼、個人資料、內部文件或未獲授權公開的內容。
