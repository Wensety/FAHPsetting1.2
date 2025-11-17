# GitHub 亂碼問題修正指南

## 問題原因

Git 提交訊息在 GitHub 上顯示亂碼，通常是因為：
1. Git 編碼設定不正確
2. Windows 終端編碼與 Git 編碼不一致
3. 提交時未使用 UTF-8 編碼

## 已執行的修正

### 1. Git 編碼設定（已完成）

```bash
# 全局設定
git config --global i18n.commitencoding utf-8
git config --global i18n.logoutputencoding utf-8
git config --global core.quotepath false

# 本地倉庫設定
git config --local i18n.commitencoding utf-8
git config --local i18n.logoutputencoding utf-8
```

### 2. Windows 終端編碼設定

在 PowerShell 中執行：
```powershell
chcp 65001  # 設置為 UTF-8
$env:LANG="zh_TW.UTF-8"
```

## 未來提交的正確方式

### 方法 1：使用環境變數（推薦）

在提交前設置環境變數：
```powershell
$env:LANG="zh_TW.UTF-8"
chcp 65001
git commit -m "提交訊息"
```

### 方法 2：使用 Git 編輯器

```bash
git commit
# 在編輯器中輸入中文訊息
# 確保編輯器使用 UTF-8 編碼保存
```

### 方法 3：從檔案讀取訊息

```bash
# 創建 UTF-8 編碼的訊息檔案
echo "提交訊息" > commit_msg.txt
git commit -F commit_msg.txt
```

## 修正歷史提交（可選）

如果需要修正已推送的提交訊息：

```bash
# 修正最後一個提交
git commit --amend -m "正確的提交訊息"
git push --force-with-lease origin main

# 注意：強制推送會改寫歷史，請謹慎使用
```

## 驗證設定

檢查編碼設定：
```bash
git config --get i18n.commitencoding
git config --get i18n.logoutputencoding
```

應該顯示：`utf-8`

## 檔案編碼檢查

確保所有文字檔案使用 UTF-8 編碼：
- Python: `# -*- coding: utf-8 -*-`
- Markdown: 以 UTF-8 編碼保存
- 其他文字檔案: UTF-8 編碼

## 建議

1. **使用英文提交訊息**（最安全）
2. **如果使用中文，確保終端和 Git 都使用 UTF-8**
3. **在提交前檢查編碼設定**

