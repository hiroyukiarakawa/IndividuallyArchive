Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")


' Dim rc As VbMsgBoxResult
rc = MsgBox("�ʈ��k�𑗂郁�j���[����������܂����H", vbYesNo + vbQuestion)
If rc = vbNo Then
    WScript.Quit
End If

SendTo = ws.SpecialFolders("SendTo")

shortcut = SendTo & "\�ʈ��k.lnk"

If fso.FileExists(shortcut) Then
    fso.DeleteFile(shortcut)
    rc = MsgBox("���郁�j���[����������܂���",vbInformation)
Else
    rc = MsgBox("�ʈ��k�����郁�j���[�ɓo�^����Ă��܂���ł���",vbInformation)
End If


