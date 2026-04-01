Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count = 0 Then
    MsgBox "ファイルが選択されていません"
    WScript.Quit
End If

strFile = WScript.Arguments(0)

strShortcutName = InputBox("ショートカット名を入力してください", "ショートカット作成")

If strShortcutName = "" Then
    MsgBox "名前が入力されませんでした。"
    WScript.Quit
End If

strShortcutName = Trim(strShortcutName)

strShortcutName = Replace(strShortcutName, "\", "")
strShortcutName = Replace(strShortcutName, "/", "")
strShortcutName = Replace(strShortcutName, ":", "")
strShortcutName = Replace(strShortcutName, "*", "")
strShortcutName = Replace(strShortcutName, "?", "")
strShortcutName = Replace(strShortcutName, "<", "")
strShortcutName = Replace(strShortcutName, ">", "")
strShortcutName = Replace(strShortcutName, "|", "")

basePath = "C:\Users\PC_User\Downloads\work"
strShortcutPath = basePath & "\" & strShortcutName & ".lnk"

Set objShortcut = objShell.CreateShortcut(strShortcutPath)
objShortcut.TargetPath = strFile
objShortcut.Save

' ★ブロック解除（Zone.Identifier削除）
zoneFile = strShortcutPath & ":Zone.Identifier"
If fso.FileExists(zoneFile) Then
    fso.DeleteFile zoneFile
End If

MsgBox "作成完了（ブロック解除済み）"