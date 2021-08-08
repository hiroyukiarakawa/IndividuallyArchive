'個別圧縮
'
'
Set WshShell = WScript.CreateObject("WScript.Shell")

Set objArgs = Wscript.Arguments

For Each arg In objArgs
    'https://ginpro.winofsql.jp/article/267579611.html
    ' cmdline = "C:\Program Files\7-Zip\7z.exe a " & arg & " " & arg & ".7z"
    cmdline = """C:\Program Files\7-Zip\7z.exe"" a " & arg & " " & arg & ".7z"
    ' cmdline = "notepad.exe " & arg
    Wscript.echo cmdline
    Call WshShell.Run( cmdline , , True )

Next


' For I = 0 To objArgs.Count - 1 
'     arg = objArgs(I)
'     Wscript.echo arg
' Next
