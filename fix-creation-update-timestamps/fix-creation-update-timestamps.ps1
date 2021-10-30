# ======================================================================
#
# 日付修正スクリプト
#
# Copyright (C) 2021 tag. All rights reserved.
#
# ======================================================================

Add-Type -AssemblyName System.Drawing
$ShellApplicationObject = New-Object -COMObject Shell.Application

# デバッグモード
$DebugPreference = "SilentlyContinue"
$DebugPreference = "Continue"

# スクリプトディレクトリへの移動
$ScriptDir = Split-Path $MyInvocation.MyCommand.Path -Parent
Set-Location $ScriptDir

# 不可視文字の削除
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

Get-ChildItem -File | ForEach-Object {

    Write-Debug "----------------------------------------------------------------------"
    $DirectoryPath = $_.DirectoryName
    $FullFilePath = $_.FullName
    $FileName = $_.Name
    $BaseName = $_.BaseName
    $FileExtension = [System.IO.Path]::GetExtension($FullFilePath).ToLower()
    
    Write-Debug "ファイルパス: ${FullFilePath}"
    Write-Debug "ディレクトリ: ${DirectoryPath}"
    Write-Debug "ファイル名: ${FileName}"
    Write-Debug "ベース名: ${BaseName}"
    Write-Debug "拡張子: ${FileExtension}"

    $ShellFolder = $ShellApplicationObject.Namespace($DirectoryPath)

    if (@(".jpg", ".jpeg", ".png", ".gif", ".arw", ".heic", ".avi", ".mov", ".mp4").Contains($FileExtension)) {
        Write-Debug "** 処理対象ファイル"

        $image = $null
        $DateTaken = ""

        # Drawing.Bitmap による「年/月/日 時:分:秒」の取得
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            try {
                Write-Debug "** 取得: 撮影日時 (秒まで)"
                $image = New-Object System.Drawing.Bitmap($FullFilePath.ToString())
                $DateTaken = [System.Text.Encoding]::ASCII.GetString($image.GetPropertyItem(36867).Value).ToString() -replace "`0", ""
                $DateTaken = [DateTime]::ParseExact($DateTaken, "yyyy:MM:dd HH:mm:ss", $null)
            }
            catch {
                Write-Debug "** 失敗"
            }
            finally {
                if ($image -ne $null) {
                    $image.Dispose()
                }
            }
        }

        # よくありそうなファイル名「yyyyMMdd-HHmmss」から取得
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            $Pattern = "yyyyMMdd-HHmmss"
            Write-Debug "** 取得: $Pattern"
            try {
                $DateTaken = $BaseName.Substring(0, $Pattern.Length)
                $DateTaken = [DateTime]::ParseExact($DateTaken, $Pattern, $null)
            }
            catch {
                Write-Debug "** 失敗"
                $DateTaken = $null
            }
        }

        # よくありそうなファイル名「yyyyMMdd_HHmmss」から取得
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            $Pattern = "yyyyMMdd_HHmmss"
            Write-Debug "** 取得: $Pattern"
            try {
                $DateTaken = $BaseName.Substring(0, $Pattern.Length)
                $DateTaken = [DateTime]::ParseExact($DateTaken, $Pattern, $null)
            }
            catch {
                Write-Debug "** 失敗"
                $DateTaken = $null
            }
        }

        # Amazon Photos での自動生成名「yyyy-MM-dd_HH-mm-ss_nnn」から取得
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            $Pattern = "yyyy-MM-dd_HH-mm-ss"
            Write-Debug "** 取得: $Pattern"
            try {
                $DateTaken = $BaseName.Substring(0, $Pattern.Length)
                $DateTaken = [DateTime]::ParseExact($DateTaken, $Pattern, $null)
            }
            catch {
                Write-Debug "** 失敗"
                $DateTaken = $null
            }
        }

        # GetDetailsOf による「年/月/日 時:分」の取得
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            Write-Debug "** 取得: 撮影日時"
            $DateTaken = $ShellFolder.GetDetailsOf($ShellFolder.ParseName($FileName), 12)    # 撮影日時
            $DateTaken = Remove-invisibleCharacter($DateTaken)
            Write-Debug "日時: $DateTaken"
            if ($DateTaken -ne "") { $DateTaken = [DateTime]::ParseExact($DateTaken, "yyyy/MM/dd HH:mm", $null) } else { Write-Debug "** 失敗" }
        }
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            Write-Debug "** 取得: メディアの作成日時"
            $DateTaken = $ShellFolder.GetDetailsOf($ShellFolder.ParseName($FileName), 208)    # メディアの作成日時
            $DateTaken = Remove-invisibleCharacter($DateTaken)
            Write-Debug "日時: $DateTaken"
            if ($DateTaken -ne "") { $DateTaken = [DateTime]::ParseExact($DateTaken, "yyyy/MM/dd HH:mm", $null) } else { Write-Debug "** 失敗" }
        }
        if (($DateTaken -eq $null) -or ($DateTaken -eq "")) {
            Write-Debug "** 取得: コンテンツの作成日時"
            $DateTaken = $ShellFolder.GetDetailsOf($ShellFolder.ParseName($FileName), 152)    # コンテンツの作成日時
            $DateTaken = Remove-invisibleCharacter($DateTaken)
            Write-Debug "日時: $DateTaken"
            if ($DateTaken -ne "") { $DateTaken = [DateTime]::ParseExact($DateTaken, "yyyy/MM/dd HH:mm", $null) } else { Write-Debug "** 失敗" }
        }

        Write-Debug "日時: $DateTaken"

        if ($DateTaken -ne $null -and $DateTaken -ne "") {
            Write-Debug "** 更新: 作成日時"
            Set-ItemProperty $FileName -Name CreationTime -Value $DateTaken.ToString()
    
            Write-Debug "** 更新: 更新日時"
            Set-ItemProperty $FileName -Name LastWriteTime -Value $DateTaken.ToString()
        }
        else {
            Write-Debug "** 日付が取得できなかったため、更新しませんでした"
        }
    }
    else {
        Write-Debug "** 処理対象外ファイル"
    }

}

pause
