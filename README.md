# Auto-Worship-Folder

在 Windows 環境中，創建整個崇拜流程所需用的文件，並按順序全部放置在同一資料夾內。

## 好處
1. 一鍵創建崇拜所用 PPTX 包括宣召、經訓、讀經、禱告、講道、奉獻、報告和祝福等內容。不用每次手動複製全部範本，可節省時間並減少遺漏檔案。
2. 面對臨時加插 PPTX，只需運行 `refresh-master.vbs` 重新併合 master.pptx 便可，減低臨場突發產生的混亂。


## 首次使用

1. **下載**
* 1.1 <> Code 按 Download ZIP
* 1.2 解壓後，將整個資料夾放在你想要的地方便可
2. **建立詩歌資料夾捷徑**
* 2.1 複製你存放詩歌 PowerPoint 的資料夾路徑
* 2.2 在剛下載的 Auto-Worship-Folder-main\templates 內找到 songs-folder
* 2.3 按右鍵，然後選 Properties
* 2.4 將 2.1 的路徑貼在 Target
* 2.5 按 Apply，再按 OK 完成


## 如何使用

1. 打開 `宣召及啟應.txt`、`經訓.txt`、`讀經.txt` 並按現有格式更新經文與相關內容。
2. 雙擊 `create-worship-folder.vbs`。
3. 輸入要預備的崇拜名稱，例如 20260906。
4. 輸入領詩的全部歌名，以逗號分隔，例如 歡欣, 生趣, 慈繩愛索, 十架的冠冕。
> 程式會在 songs-folder 的資料夾中尋找與歌名完全相符的 PPTX，如找到便會將之複製過來
> 如找不到，便會創建一個以(找不到) 開首的txt檔，例如(找不到) 十架的冠冕.txt
5. 等待約數十秒，畫面會出現 `完成。master.pptx 已更新` 字樣。
6. 按 OK 後再次見到 `完成`。
7. 此時資料夾中會新增了一個你在 #3 輸入的資料夾，裡面便是崇拜所用的全部檔案。

## refresh-master.vbs
重新讀取一次資料夾內的全部 PPTX，並按順序更新 master.pptx 和播放檔 (PPSX)

### 使用方法

1. 關閉所有 PowerPoint，特別是 `master.pptx`。
2. 先複製一份 `master.pptx` 作備份。
3. 進入：

   ```text
   templates\worship-files
   ```

4. 雙擊 `refresh-master.vbs`。
5. 完成後重新開啟 `master.pptx`，檢查內容是否正確。

> 此 script 可能會修改或覆蓋 `master.pptx`。首次使用或修改 script 後，必須先備份。

## VBS 是甚麼？

VBS（VBScript）是 Windows 的自動化腳本檔。

本專案的主要檔案是：

```text
create-worship-folder.vbs
```

它會建立崇拜資料夾並複製範本。

## 如何查看或修改 VBS

不要直接雙擊 `.vbs` 檔來閱讀，因為 Windows 通常會立刻執行它。

1. 在 `.vbs` 檔按右鍵。
2. 選「開啟檔案方式」。
3. 選擇「記事本」、Notepad++ 或 Visual Studio Code。
4. 修改前先備份原始檔。
5. 修改後先在測試資料夾執行，確認正常才正式使用。

## 用 AI 檢查安全

執行 `.vbs` 前，先用記事本或 VS Code 開啟它。

把程式碼貼到 AI，並輸入：

```text
請用繁體中文解釋這段 VBS 的功能。

請檢查它有沒有刪除、覆蓋、移動或複製檔案，
執行 CMD／PowerShell、下載資料、連接網絡、
讀取密碼或上傳資料。

請列出它會讀取、建立、修改或刪除的檔案和資料夾。

請不要修改程式碼，只報告實際功能和風險。
```

特別留意以下字眼：

```text
DeleteFile
DeleteFolder
MoveFile
CopyFile
Run
Exec
cmd.exe
powershell
XMLHTTP
WinHttpRequest
```

這些字眼不一定代表危險，但執行前必須明白它的用途。

> 不要把密碼、API Key、個人資料、教會內部資料或完整電腦路徑貼到公開 AI 對話。

## 注意

- 本工具只適用於 Windows。
- 公開分享前，請檢查 `.vbs`、`.txt` 和 `.pptx` 沒有密碼、個人資料或內部文件。
