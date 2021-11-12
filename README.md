# 概要

対象フォルダ内のエクセルファイルを以下の内容で保存し、資料作成の仕上げを短縮できます。

- 全シート"A1"セルを選択
- 拡縮を100%
- 1番左のシートを選択

# Badge

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/9d7eb6134ff24e8f8b18a7c205bbe770)](https://www.codacy.com/gh/ishi720/excelFinishSave/dashboard?utm_source=github.com&amp;utm_medium=referral&amp;utm_content=ishi720/excelFinishSave&amp;utm_campaign=Badge_Grade)

# 実行方法

## 引数なしで実行する（カレントディレクトリ内を処理）

1. `excelFinishSave.ps1`を処理したいフォルダに設置
1. PowerShellで以下のコマンド実行

```ps1
PowerShell -ExecutionPolicy RemoteSigned ".\excelFinishSave.ps1"
```

## 引数ありで実行する（対象フォルダ内を処理）

1. PowerShellで以下のコマンド実行

```ps1
PowerShell -ExecutionPolicy RemoteSigned ".\excelFinishSave.ps1" {フォルダPATH}
```
