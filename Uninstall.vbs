Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")


' Dim rc As VbMsgBoxResult
rc = MsgBox("�ʈ��k�𑗂郁�j���[����������܂����H", vbYesNo + vbQuestion)
If rc = vbNo Then
    WScript.Quit
End If

SendTo = ws.SpecialFolders("SendTo")

shortcut = SendTo & "\�ʈ��k.lnk"
shortcut_del = SendTo & "\�ʈ��k(���k��폜).lnk"

If fso.FileExists(shortcut) Then
    fso.DeleteFile(shortcut)
End If    
If fso.FileExists(shortcut_del) Then
    fso.DeleteFile(shortcut_del)
End If

rc = MsgBox("���郁�j���[����������܂���",vbInformation)


