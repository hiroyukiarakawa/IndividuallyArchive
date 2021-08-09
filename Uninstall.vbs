Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")


' Dim rc As VbMsgBoxResult
rc = MsgBox("個別圧縮を送るメニューから解除しますか？", vbYesNo + vbQuestion)
If rc = vbNo Then
    WScript.Quit
End If

SendTo = ws.SpecialFolders("SendTo")

shortcut = SendTo & "\個別圧縮.lnk"

If fso.FileExists(shortcut) Then
    fso.DeleteFile(shortcut)
    rc = MsgBox("送るメニューから解除しました",vbInformation)
Else
    rc = MsgBox("個別圧縮が送るメニューに登録されていませんでした",vbInformation)
End If


