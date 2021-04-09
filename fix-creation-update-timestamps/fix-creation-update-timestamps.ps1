Write-Host ""
Write-Host "======================================================================"
Write-Host ""
Write-Host "Fix Creation Timestamp and Update Timestamps"
Write-Host ""
Write-Host "Copyright (C) 2021 tag. All rights reserved."
Write-Host ""
Write-Host "======================================================================"
Write-Host ""

Add-Type -AssemblyName System.Drawing
$ShellApplicationObject = New-Object -COMObject Shell.Application

# �f�o�b�O���[�h
$DebugPreference = "SilentlyContinue"
#$DebugPreference = "Continue"

# �X�N���v�g�f�B���N�g���ւ̈ړ�
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location $ScriptDir

# �s�������̍폜
Function Remove-InvisibleCharacter([string]$Text) {
    $Text = $Text -replace "\u200e", ""    # LEFT-TO-RIGHT MARK
    $Text = $Text -replace "\u200f", ""    # RIGHT-TO-LEFT MARK
    $Text = $Text -replace "\u202a", ""    # LEFT-TO-RIGHT EMBEDDING
    $Text = $Text -replace "\u202b", ""    # RIGHT-TO-LEFT EMBEDDING
    $Text = $Text -replace "\u202c", ""    # POP DIRECTIONAL FORMATTING
    $Text = $Text -replace "\u202d", ""    # LEFT-TO-RIGHT OVERRIDE
    $Text = $Text -replace "\u202e", ""    # RIGHT-TO-LEFT OVERRIDE
    $Text = $Text -replace "\u2066", ""    # LEFT-TO-RIGHT ISOLATE
    $Text = $Text -replace "\u2067", ""    # RIGHT-TO-LEFT ISOLATE
    $Text = $Text -replace "\u2068", ""    # FIRST STRONG ISOLATE
    $Text = $Text -replace "\u2069", ""    # POP DIRECTIONAL ISOLATE
    return $Text
}

Get-ChildItem -File $ScriptDir | % {

    Write-Host "----------------------------------------------------------------------"
    $DirectoryPath = $_.DirectoryName
    $FullFilePath = $_.FullName
    $FileName = $_.Name
    $BaseName = $_.BaseName
    $FileExtension = [System.IO.Path]::GetExtension($FullFilePath).ToLower()
    
    Write-Debug "�t�@�C���p�X: ${FullFilePath}"
    Write-Debug "�f�B���N�g��: ${DirectoryPath}"
    Write-Host "�t�@�C����: ${FileName}"
    Write-Debug "�x�[�X��: ${BaseName}"
    Write-Debug "�g���q: ${FileExtension}"

    $ShellFolder = $ShellApplicationObject.Namespace($DirectoryPath)

    if (@(".jpg", ".jpeg", ".png", ".gif", ".arw", ".heic", ".avi", ".mov", ".mp4").Contains($FileExtension.ToLower())) {
        Write-Debug "** �����Ώۃt�@�C��"

        $image = $null
        $DateTaken = ""

        # Drawing.Bitmap �ɂ��u�N/��/�� ��:��:�b�v�̎擾
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            try {
                Write-Debug "** �擾: �B�e���� (�b�܂�)"
                $image = New-Object System.Drawing.Bitmap($FullFilePath.ToString())
                $DateTaken = [System.Text.Encoding]::ASCII.GetString($image.GetPropertyItem(36867).Value).ToString() -replace "`0", ""
                $image.Dispose()
                $DateTaken = [DateTime]::ParseExact($DateTaken, "yyyy:MM:dd HH:mm:ss", $null)
            } catch {
                Write-Debug "** ���s"
            }
        }

        # Amazon Photos �ł̎����������uyyyy-MM-dd_HH-mm-ss_nnn�v����擾
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            Write-Debug "** �擾: Amazon Photos"
            try {
                $DateTaken = $BaseName.Substring(0, 19)
                $DateTaken = [DateTime]::ParseExact($DateTaken, "yyyy-MM-dd_HH-mm-ss", $null)
            } catch {
                Write-Debug "** ���s"
            }
        }

        # GetDetailsOf �ɂ��u�N/��/�� ��:���v�̎擾
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            Write-Debug "** �擾: �B�e����"
            $DateTaken = $ShellFolder.GetDetailsOf($ShellFolder.ParseName($FileName), 12)    # �B�e����
            $DateTaken = Remove-invisibleCharacter($DateTaken)
            Write-Debug "����: $DateTaken"
            if ($DateTaken -ne "") { $DateTaken = [DateTime]::ParseExact($DateTaken, "yyyy/MM/dd HH:mm", $null) } else { Write-Debug "** ���s" }
        }
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            Write-Debug "** �擾: ���f�B�A�̍쐬����"
            $DateTaken = $ShellFolder.GetDetailsOf($ShellFolder.ParseName($FileName), 208)    # ���f�B�A�̍쐬����
            $DateTaken = Remove-invisibleCharacter($DateTaken)
            Write-Debug "����: $DateTaken"
            if ($DateTaken -ne "") { $DateTaken = [DateTime]::ParseExact($DateTaken, "yyyy/MM/dd HH:mm", $null) } else { Write-Debug "** ���s" }
        }
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            Write-Debug "** �擾: �R���e���c�̍쐬����"
            $DateTaken = $ShellFolder.GetDetailsOf($ShellFolder.ParseName($FileName), 152)    # �R���e���c�̍쐬����
            $DateTaken = Remove-invisibleCharacter($DateTaken)
            Write-Debug "����: $DateTaken"
            if ($DateTaken -ne "") { $DateTaken = [DateTime]::ParseExact($DateTaken, "yyyy/MM/dd HH:mm", $null) } else { Write-Debug "** ���s" }
        }

        Write-Host "����: $DateTaken"

        if ($DateTaken -ne $null -and $DateTaken -ne "") {
            Write-Debug "** �X�V: �쐬����"
            Set-ItemProperty $FileName -Name CreationTime -Value $DateTaken.ToString()
    
            Write-Debug "** �X�V: �X�V����"
            Set-ItemProperty $FileName -Name LastWriteTime -Value $DateTaken.ToString()
        } else {
            Write-Debug "** ���t���擾�ł��Ȃ��������߁A�X�V���܂���ł���"
        }
    } else {
        Write-Debug "** �����ΏۊO�t�@�C��"
    }

}

pause