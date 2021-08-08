Set fso = CreateObject("Scripting.FileSystemObject")
Set ws = CreateObject("WScript.Shell")


' Dim rc As VbMsgBoxResult
rc = MsgBox("Do you want to Uninstall IndividuallyArchive.vbs?", vbYesNo + vbQuestion)
If rc = vbNo Then
    WScript.Quit
End If

SendTo = ws.SpecialFolders("SendTo")

shortcut = SendTo & "\IndividuallyArchive.lnk"

fso.DeleteFile(shortcut)

MsgBox("Uninstalled")


