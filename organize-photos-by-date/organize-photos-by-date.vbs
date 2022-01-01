Option Explicit

' ======================================================================
'
' Organize Photos by Taken Date
' Copyright (C) 2017-2020 tag. All rights reserved.
'
' ======================================================================

' Objects
Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
Dim objShellApp : Set objShellApp = CreateObject("Shell.Application")
Dim objShell : Set objShell = WScript.CreateObject("WScript.Shell")

' スクリプトフォルダパス
Dim strScriptDir : strScriptDir = Left(WScript.ScriptFullName, Len(WScript.ScriptFullName) - Len(WScript.ScriptName))
Dim strSubScriptDir : strSubScriptDir = WScript.ScriptName
Dim objScriptDir : Set objScriptDir = objShellApp.NameSpace(strScriptDir)

Dim objDirItems, objItem, strCommand, strFileExtension, strDateOfTaken, strDateFolderPath

Function removeInvisibleCharacter(str)
	str = Replace(str, ChrW(&H200E), "")	' LEFT-TO-RIGHT MARK
	str = Replace(str, ChrW(&H200F), "")	' RIGHT-TO-LEFT MARK
	str = Replace(str, ChrW(&H202A), "")	' LEFT-TO-RIGHT EMBEDDING
	str = Replace(str, ChrW(&H202B), "")	' RIGHT-TO-LEFT EMBEDDING
	str = Replace(str, ChrW(&H202C), "")	' POP DIRECTIONAL FORMATTING
	str = Replace(str, ChrW(&H202D), "")	' LEFT-TO-RIGHT OVERRIDE
	str = Replace(str, ChrW(&H202E), "")	' RIGHT-TO-LEFT OVERRIDE
	str = Replace(str, ChrW(&H2066), "")	' LEFT-TO-RIGHT ISOLATE
	str = Replace(str, ChrW(&H2067), "")	' RIGHT-TO-LEFT ISOLATE
	str = Replace(str, ChrW(&H2068), "")	' FIRST STRONG ISOLATE
	str = Replace(str, ChrW(&H2069), "")	' POP DIRECTIONAL ISOLATE
	removeInvisibleCharacter = str
End Function

Set objDirItems = objScriptDir.Items()
For Each objItem in objDirItems
	If Not objItem.IsFolder Then
		strFileExtension = LCase(objFso.GetExtensionName(objItem))
		If strFileExtension = "jpg" Or strFileExtension = "png" Or strFileExtension = "3gp" Or strFileExtension = "raw" Or strFileExtension = "avi" Or strFileExtension = "mp4" Then
			'WScript.Echo objScriptDir.ParseName(objItem.Name)

			' Exif 情報の撮影日時を取得
			strDateOfTaken = objScriptDir.GetDetailsOf(objScriptDir.ParseName(objItem.Name), 12)	' 12: 撮影日時
			'WScript.Echo strDateOfTaken

			' Exif 情報の撮影日時が取得できない場合、ファイルの更新日時を使用
			If Len(strDateOfTaken) = 0 Then
				strDateOfTaken = objScriptDir.GetDetailsOf(objScriptDir.ParseName(objItem.Name), 3)	' 3: 更新日時
				'WScript.Echo strDateOfTaken
			End If

			' 何かしら日時が取得できた場合、日付フォルダを作成し、移動
			If Len(strDateOfTaken) <> 0 Then
				strDateOfTaken = Replace(Split(strDateOfTaken, " ")(0), "/", "")
				strDateOfTaken = removeInvisibleCharacter(strDateOfTaken)
				strDateFolderPath = objFso.BuildPath(strScriptDir, strDateOfTaken)
				If Not objFso.FolderExists(strDateFolderPath) Then
					objFso.CreateFolder(strDateFolderPath)
				End If
				objFso.MoveFile objFso.BuildPath(strScriptDir, objItem.Name), objFso.BuildPath(strDateFolderPath, objItem.Name)
			End If

		End If
	End If
Next

WScript.Quit
