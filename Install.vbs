Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")


' Dim rc As VbMsgBoxResult
rc = MsgBox("Do you want to create SHORTCUT at SendTo for IndividuallyArchive.vbs", vbYesNo + vbQuestion)
If rc = vbNo Then
    WScript.Quit
End If

SendTo = ws.SpecialFolders("SendTo")

MyFolder = fso.getParentFolderName(WScript.ScriptFullName)

' ショートカット作成
Set shortcut = ws.CreateShortcut(SendTo & "\IndividuallyArchive.lnk")
With shortcut
    .TargetPath = MyFolder & "\IndividuallyArchive.vbs"
    .WorkingDirectory = MyFolder
    .Save
End With

MsgBox("Installed")
