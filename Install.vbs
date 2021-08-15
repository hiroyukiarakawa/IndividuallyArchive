Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")
s7z = "C:\Program Files\7-Zip\7z.exe"

If fso.FileExists(s7z) = False Then
    rc = MsgBox(s7z & "が見つかりません", vbInformation)
    WScript.Quit
End If

' Dim rc As VbMsgBoxResult
rc = MsgBox("個別圧縮を送るメニューに登録しますか？", vbYesNo + vbQuestion)
If rc = vbNo Then
    WScript.Quit
End If

SendTo = ws.SpecialFolders("SendTo")

MyFolder = fso.getParentFolderName(WScript.ScriptFullName)

' ショートカット作成
Set shortcut = ws.CreateShortcut(SendTo & "\個別圧縮.lnk")
With shortcut
    .TargetPath = MyFolder & "\IndividuallyArchive.vbs"
    .WorkingDirectory = MyFolder
    .Save
End With

Set shortcut_del = ws.CreateShortcut(SendTo & "\個別圧縮(圧縮後削除).lnk")
With shortcut_del
    .TargetPath = MyFolder & "\IndividuallyArchive_delete.vbs"
    .WorkingDirectory = MyFolder
    .Save
End With


rc = MsgBox("送るメニューに登録しました", vbInformation)
