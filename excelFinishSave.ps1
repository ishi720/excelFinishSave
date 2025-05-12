###
# 指定したのエクセルファイル(.xlsm,.xlsx)を
# 「全シート"A1"セルを選択」かつ「一番左のシートを表示」して保存する。
###

###
# 関数
###

# ダイアログを出して、ファイルを選択する
# @return fileList ファイルリスト
function fileSelect() {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excelファイル形式|*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xls;*.xml;*.xlam;*.xla;*.xlw;*.xlr;"

    # 起動時のディレクトリPath
    $dialog.InitialDirectory = Convert-Path .

    # ダイアログウインドウタイトル
    $dialog.Title = "ファイル選択"

    # 複数選択
    $dialog.Multiselect = $true

    # ダイアログ表示
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        return $dialog.FileNames
    } else {
        return $null
    }
}

###
# メイン処理
###

# エクセル操作初期化
$excel = New-Object -ComObject Excel.Application

# エクセル可視化
$excel.Visible = $False

# 処理ファイルの選択
$itemList = fileSelect
foreach($targetFile in $itemList) {

    # 処理対象ファイル名表示
    Write-Host "FileName:" $targetFile

    # エクセルを開く
    $book = $excel.Workbooks.Open($targetFile)

    # 存在するシート分処理する
    foreach ($s in $book.sheets){
        if ($s.Visible) {
            $sheet = $book.Sheets.item($s.name)
            $sheet.Activate()
            if ($excel.ActiveWindow.FreezePanes) {
                $excel.ActiveWindow.SmallScroll(
                    0,
                    $excel.ActiveWindow.ScrollRow,
                    0,
                    $excel.ActiveWindow.ScrollColumn
                ) | out-null
            }
            $excel.ActiveWindow.Zoom = 100
            $sheet.Range("A1").Select() | out-null
            Write-Host "  SheetName:" $s.name " [Processing completed.]"
        } else {
            Write-Host "  SheetName:" $s.name " [Hidden sheet skip.]"
        }
    }

    # 一番左のシートをアクティブにする
    $book.Sheets.item(1).Activate()

    # 保存
    $book.Save()

    # 閉じる
    $book.Close()

    Write-Host "  Saved`r`n"
}

# 後始末
$excel.Quit()
$excel = $null
[GC]::Collect()

Pause
