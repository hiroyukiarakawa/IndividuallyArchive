Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")
s7z = "C:\Program Files\7-Zip\7z.exe"

If fso.FileExists(s7z) = False Then
    rc = MsgBox(s7z & "��������܂���", vbInformation)
    WScript.Quit
End If

' Dim rc As VbMsgBoxResult
rc = MsgBox("�ʈ��k�𑗂郁�j���[�ɓo�^���܂����H", vbYesNo + vbQuestion)
If rc = vbNo Then
    WScript.Quit
End If

SendTo = ws.SpecialFolders("SendTo")

MyFolder = fso.getParentFolderName(WScript.ScriptFullName)

' �V���[�g�J�b�g�쐬
Set shortcut = ws.CreateShortcut(SendTo & "\�ʈ��k.lnk")
With shortcut
    .TargetPath = MyFolder & "\IndividuallyArchive.vbs"
    .WorkingDirectory = MyFolder
    .Save
End With

Set shortcut_del = ws.CreateShortcut(SendTo & "\�ʈ��k(���k��폜).lnk")
With shortcut_del
    .TargetPath = MyFolder & "\IndividuallyArchive_delete.vbs"
    .WorkingDirectory = MyFolder
    .Save
End With


rc = MsgBox("���郁�j���[�ɓo�^���܂���", vbInformation)
