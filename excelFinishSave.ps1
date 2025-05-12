###
# エクセルファイルをカーソル位置などを初期状態に戻し保存する。
# ・全シート"A1"セルを選択
# ・一番左のシートを表示
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
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileNames
    } else {
        return $null
    }
}

# シートの処理を行う関数
# @param $sheet 処理対象のシート
# @param $book エクセルブック
function processSheet($sheet, $book) {
    if ($sheet.Visible) {
        $sheet.Activate()
        if ($book.Application.ActiveWindow.FreezePanes) {
            $book.Application.ActiveWindow.SmallScroll(
                0,
                $book.Application.ActiveWindow.ScrollRow,
                0,
                $book.Application.ActiveWindow.ScrollColumn
            ) | Out-Null
        }
        $book.Application.ActiveWindow.Zoom = 100
        $sheet.Range("A1").Select() | Out-Null
        Write-Host "  SheetName: $($sheet.Name) [Processing completed.]"
    } else {
        Write-Host "  SheetName: $($sheet.Name) [Hidden sheet skipped.]"
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
foreach ($targetFile in $itemList) {

    # 処理対象ファイル名表示
    Write-Host "FileName:" $targetFile

    # エクセルを開く
    $book = $excel.Workbooks.Open($targetFile)

    # 各シートに毎に処理を行う
    foreach ($sheet in $book.sheets) {
        processSheet $sheet $book
    }

    # 一番左のシートをアクティブにする
    $book.Sheets.item(1).Activate()

    # 保存
    $book.Save()

    # 閉じる
    $book.Close($false)

    Write-Host "  Saved`r`n"
}

# 後始末
$excel.Quit()
$excel = $null
[GC]::Collect()

Pause
