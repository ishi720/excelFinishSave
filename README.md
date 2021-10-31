# 概要

カレントディレクトリ内のエクセルファイル(.xlsx)を
「全シート"A1"セルを選択」かつ「一番左のシートを表示」して保存する。

# 実行方法

1. `excelFinishSave.ps1`を処理したいフォルダに設置
1. PowerShellで以下のコマンド実行

```ps1
PowerShell -ExecutionPolicy RemoteSigned ".\excelFinishSave.ps1"
```
