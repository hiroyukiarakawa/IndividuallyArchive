'�ʈ��k
'�����œn���ꂽ�t�@�C�������k����o�b�`�t�@�C�����쐬���Ď��s����B
'�ʃt�@�C�����ƂɁAWshShell.Run���s���ƌ��ʂ��󂯎��̂����J�Ȃ̂ŁB
'�Ȃ��A�S���o�b�`�t�@�C���ōs��Ȃ��̂��H�@�R�}���h���C�������̏�����()���g��ꂽ�t�@�C�������w�肳���ƌ�쓮���N�����̂ŁB

Set WshShell = WScript.CreateObject("WScript.Shell")
'https://docs.microsoft.com/ja-jp/office/vba/language/reference/user-interface-help/filesystemobject-object
Set fso = CreateObject("Scripting.FileSystemObject")
Set objArgs = Wscript.Arguments

'�o�b�`�t�@�C���������߂�B
MyFolder = fso.getParentFolderName(WScript.ScriptFullName)
I = 0
TempBatFile = MyFolder & "\temp" & I & ".bat"
Do While fso.FileExists(TempBatFile) = True
    I = I + 1
    TempBatFile = MyFolder & "\temp" & I & ".bat"
Loop
'�o�b�`�t�@�C�����쐬����
Set tso = fso.CreateTextFile(TempBatFile, true)
'�R�}���h���C�������œn���ꂽ�t�@�C�����o�b�`�t�@�C���ɓo�^����
For Each arg In objArgs
    'https://ginpro.winofsql.jp/article/267579611.html
    cmdline = """C:\Program Files\7-Zip\7z.exe"" a """ & arg  & ".7z"" """ & arg & """"
    ' Wscript.echo cmdline
    tso.WriteLine( cmdline )
Next
tso.WriteLine("PAUSE")
tso.Close

'�쐬�����o�b�`�t�@�C�������s����
Call WshShell.Run( TempBatFile , 1 , True )

'�쐬�����o�b�`�t�@�C�����폜����
fso.DeleteFile(TempBatFile)
