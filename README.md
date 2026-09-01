# Auto-Worship-Folder

在 Windows 環境中，創建整個崇拜流程所需用的文件，並按順序全部放置在同一資料夾內。

## 好處
1. 一鍵創建崇拜所用 PPTX 包括宣召、經訓、讀經、禱告、講道、奉獻、報告和祝福等內容。不用每次手動複製全部範本，可節省時間並減少遺漏檔案。
2. 面對臨時加插 PPTX，只需運行 `refresh-master.vbs` 重新併合 master.pptx 便可，減低臨場突發產生的混亂。


## 如何使用

1. 不要更改 `create-worship-folder.vbs` 和 `templates` 資料夾的位置關係。
2. 雙擊 `create-worship-folder.vbs`。
3. 按畫面指示完成操作。
4. 到指定位置查看新建立的崇拜文件夾。

> 第一次使用前，請先備份整個專案。

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
- `templates\songs-folder.lnk` 是 Windows 捷徑。其他使用者下載後，可能因電腦路徑不同而失效。

如捷徑失效：

1. 找到你的詩歌 PowerPoint 資料夾。
2. 在該資料夾按右鍵。
3. 選「傳送到」→「桌面（建立捷徑）」。
4. 將新捷徑改名為 `songs-folder.lnk`。
5. 放到：

   ```text
   templates\
   ```