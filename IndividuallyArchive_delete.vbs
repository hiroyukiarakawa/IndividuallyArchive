'�ʈ��k
'���k��ɁA���̃t�@�C�����폜����o�[�W����

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
    '���k�o������폜���� (���k�������������ǂ����́AERRORLEVEL���m�F����)
    cmdline = "IF %ERRORLEVEL%==0 DEL """  & arg  & """"
    ' Wscript.echo cmdline
    tso.WriteLine( cmdline )
Next
tso.WriteLine("PAUSE")
tso.Close

'�쐬�����o�b�`�t�@�C�������s����
Call WshShell.Run( TempBatFile , 1 , True )

'�쐬�����o�b�`�t�@�C�����폜����
fso.DeleteFile(TempBatFile)
