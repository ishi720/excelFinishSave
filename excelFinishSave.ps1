###
# カレントディレクトリ内のエクセルファイル(.xlsx)を
# 「全シート"A1"セルを選択」かつ「一番左のシートを表示」して保存する。
###

# 拡張子のチェック関数
function extensionCheck($fileName) {
    $extArr = @('.xlsm','.xlsx')
    $targetExt = [System.IO.Path]::GetExtension("$fileName")
    return $extArr.Contains($targetExt)
}

# 存在しているディレクトリを判定する関数
function directoryCheck($Path) {
    $result = $False
    if (Test-Path $Path) {
        if ((Get-Item $Path).PSIsContainer) {
            $result = $True
        }
    }
    return $result
}

if ($Args[0]) {
    # 引数が指定されている場合
    if (directoryCheck($Args[0])) {
        # ディレクトリが存在している場合
        $targetDir = $Args[0]
    } else {
        # ディレクトリが存在していない場合
        Write-Host "[Error] The process ends because the target directory does not exist."
        Write-Host Directory: $Args[0]
        exit
    }
} else {
    # 引数が指定されていない場合
    # カレントディレクトリを変数にセット
    $targetDir = [System.IO.Directory]::GetCurrentDirectory()
}

# エクセル操作初期化
$excel = New-Object -ComObject Excel.Application

# エクセル可視化
$excel.Visible = $False

# カレントディレクトリ内のファイル分処理を行う
$itemList = Get-ChildItem "./"
foreach($item in $itemList) {

    # 処理対象のファイルを変数にセット
    $targetFile = Join-Path $targetDir $item.Name


    # 拡張子のチェック
    if (extensionCheck($targetFile)) {

        # 処理対象ファイル名表示
        echo $targetFile

        # エクセルを開く
        $book = $excel.Workbooks.Open($targetFile)

        # 存在するシート分処理する
        foreach ($s in $book.sheets){
            echo $s.name
            if ($s.Visible) {
                $sheet = $book.Sheets.item($s.name)
                $sheet.Activate()
                $sheet.Range("A1").Select()
            } else {
                echo "非表示シートのためスキップ"
            }
        }

        # 一番左のシートをアクティブにする
        $book.Sheets.item(1).Activate()

        # 保存
        $book.Save()

        # 閉じる
        $book.Close()
    }
}

# 後始末
$excel.Quit()
$excel = $null
[GC]::Collect()
