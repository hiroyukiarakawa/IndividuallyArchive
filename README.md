# IndividuallyArchive
選択したファイル/フォルダを１つの書庫ファイルに圧縮するのではなく、個別の書庫ファイルに圧縮したい場合に使うと便利です。
圧縮に7zを使用します。
7zのインストール先は、デフォルトの以下で決め打ちしていますので、別の場所にインストールしているのであればIndividuallyArchive.vbsの修正が必要です。
C:\Program Files\7-Zip\7z.exe

## Install.vbs
実行すると"個別圧縮"のショートカットを送るメニューに作成します。

## Uninstall.vbs
実行すると送るメニューから"個別圧縮"のショートカットを削除します。

## IndividuallyArchive.vbs
Install.vbsにより、送るフォルダにショートカットが作成されます。
複数ファイルを選択して右クリック→送るフォルダの個別圧縮を選ぶと、
選択したファイル毎に個別に圧縮した7z書庫ファイルが作成されます。

例）
C:\MyFolder\test1.txt
C:\MyFolder\test2.txt
C:\MyFolder\test3.txt
を、個別圧縮に送ると、
C:\MyFolder\test1.txt.7z
C:\MyFolder\test2.txt.7z
C:\MyFolder\test3.txt.7z
という7z書庫を作成します。
