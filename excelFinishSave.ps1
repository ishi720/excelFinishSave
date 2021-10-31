###
# カレントディレクトリ内のエクセルファイル(.xlsx)を
# 「全シート"A1"セルを選択」かつ「一番左のシートを表示」して保存する。
###

# カレントディレクトリを変数にセット
$targetDir = [System.IO.Directory]::GetCurrentDirectory()

# エクセル操作初期化
$excel = New-Object -ComObject Excel.Application

# エクセル可視化
$excel.Visible = $False

# カレントディレクトリ内のファイル分処理を行う
$itemList = Get-ChildItem "./"
foreach($item in $itemList) {

    # 処理対象のファイルを変数にセット
    $targetFile = Join-Path $targetDir $item.Name

    # 拡張子が「.xlsx」以外は処理の対象外にする
    if ([System.IO.Path]::GetExtension($targetFile) -ne ".xlsx"){
        continue
    }

    # 処理対象ファイル名表示
    echo $targetFile

    # エクセルを開く
    $book = $excel.Workbooks.Open($targetFile)

    # 存在するシート分処理する
    foreach ($s in $book.sheets){
        $sheet = $book.Sheets.item($s.name)
        $sheet.Activate()
        $sheet.Range("A1").Select()
        echo $s.name
    }

    # 一番左のシートをアクティブにする
    $book.Sheets.item(1).Activate()

    # 保存
    $book.Save()

    # 閉じる
    $book.Close()
}

# 後始末
$excel.Quit()
$excel = $null
[GC]::Collect()
