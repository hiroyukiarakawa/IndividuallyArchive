# IndividuallyArchive
選択したファイル/フォルダを１つの書庫ファイルに圧縮するのではなく、個別の書庫ファイルに圧縮したい場合に使うと便利です。
7zが
C:\Program Files\7-Zip\7z.exe
にインストールされている必要があります。

## Install.vbs
実行するとIndividuallyArchiveのショートカットを送るに作成します。

## Uninstall.vbs
実行するとショートカットを削除します。

## IndividuallyArchive.vbs
Install.vbsにより、送るフォルダにショートカットが作成されます。
複数ファイルを選択して右クリック→送るフォルダのIndividuallyArchiveを選ぶと、
選択したファイルを個別に圧縮した7z書庫ファイルが作成されます。

C:\MyFolder\test → C:\MyFolder\test.7z


