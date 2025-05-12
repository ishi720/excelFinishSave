###
# �w�肵���̃G�N�Z���t�@�C��(.xlsm,.xlsx)��
# �u�S�V�[�g"A1"�Z����I���v���u��ԍ��̃V�[�g��\���v���ĕۑ�����B
###

###
# �֐�
###

# �_�C�A���O���o���āA�t�@�C����I������
# @return fileList �t�@�C�����X�g
function fileSelect() {
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Filter = "Excel�t�@�C���`��|*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xls;*.xml;*.xlam;*.xla;*.xlw;*.xlr;"

    # �N�����̃f�B���N�g��Path
    $dialog.InitialDirectory = Convert-Path .

    # �_�C�A���O�E�C���h�E�^�C�g��
    $dialog.Title = "�t�@�C���I��"

    # �����I��
    $dialog.Multiselect = $true

    # �_�C�A���O�\��
    if($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK){
        return $dialog.FileNames
    } else {
        return $null
    }
}

###
# ���C������
###

# �G�N�Z�����쏉����
$excel = New-Object -ComObject Excel.Application

# �G�N�Z������
$excel.Visible = $False

# �����t�@�C���̑I��
$itemList = fileSelect
foreach($targetFile in $itemList) {

    # �����Ώۃt�@�C�����\��
    Write-Host "FileName:" $targetFile

    # �G�N�Z�����J��
    $book = $excel.Workbooks.Open($targetFile)

    # ���݂���V�[�g����������
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

    # ��ԍ��̃V�[�g���A�N�e�B�u�ɂ���
    $book.Sheets.item(1).Activate()

    # �ۑ�
    $book.Save()

    # ����
    $book.Close()

    Write-Host "  Saved`r`n"
}

# ��n��
$excel.Quit()
$excel = $null
[GC]::Collect()

Pause
