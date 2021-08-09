Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")


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

rc = MsgBox("送るメニューに登録しました", vbInformation)
