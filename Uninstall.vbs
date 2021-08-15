Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")


' Dim rc As VbMsgBoxResult
rc = MsgBox("個別圧縮を送るメニューから解除しますか？", vbYesNo + vbQuestion)
If rc = vbNo Then
    WScript.Quit
End If

SendTo = ws.SpecialFolders("SendTo")

shortcut = SendTo & "\個別圧縮.lnk"
shortcut_del = SendTo & "\個別圧縮(圧縮後削除).lnk"

If fso.FileExists(shortcut) Then
    fso.DeleteFile(shortcut)
End If    
If fso.FileExists(shortcut_del) Then
    fso.DeleteFile(shortcut_del)
End If

rc = MsgBox("送るメニューから解除しました",vbInformation)


