'個別圧縮
'引数で渡されたファイルを圧縮するバッチファイルを作成して実行する。
'個別ファイルごとに、WshShell.Runを行うと結果を受け取るのも一苦労なので。
'なぜ、全部バッチファイルで行わないのか？　コマンドライン引数の処理で()が使われたファイル名が指定されると誤作動を起こすので。

Set WshShell = WScript.CreateObject("WScript.Shell")
'https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/filesystemobject-object
Set fso = CreateObject("Scripting.FileSystemObject")
Set objArgs = Wscript.Arguments

'バッチファイル名を決める。
MyFolder = fso.getParentFolderName(WScript.ScriptFullName)
I = 0
TempBatFile = MyFolder & "\temp" & I & ".bat"
Do While fso.FileExists(TempBatFile) = True
    I = I + 1
    TempBatFile = MyFolder & "\temp" & I & ".bat"
Loop
'バッチファイルを作成する
Set tso = fso.CreateTextFile(TempBatFile, true)
'コマンドライン引数で渡されたファイルをバッチファイルに登録する
For Each arg In objArgs
    'https://ginpro.winofsql.jp/article/267579611.html
    cmdline = """C:\Program Files\7-Zip\7z.exe"" a """ & arg  & ".7z"" """ & arg & """"
    ' Wscript.echo cmdline
    tso.WriteLine( cmdline )
Next
tso.WriteLine("PAUSE")
tso.Close

'作成したバッチファイルを実行する
Call WshShell.Run( TempBatFile , 1 , True )

'作成したバッチファイルを削除する
fso.DeleteFile(TempBatFile)
