# Klefki-Conflu
<!-- Python (3.10.8) -->
![Python](https://img.shields.io/badge/language-Python-3776AB?style=flat-square&logo=python&logoColor=white)
![version](https://img.shields.io/badge/version-3.10.8-3776AB?style=flat-square&logo=python&logoColor=white)

<img width="1381" height="451" alt="Header_logo_large" src="https://github.com/user-attachments/assets/aaf74a19-d520-46de-a40e-520abc0c679d" />


## Download
<a href="https://github.com/Sadc2h4/Klefki-Conflu/releases/tag/v1.40a">
  <img
    src="https://raw.githubusercontent.com/Sadc2h4/brand-assets/main/button/Download_Button_1.png"
    alt="Download .zip"
    height="48"
  />
</a>
<br>
<a href="https://www.dropbox.com/scl/fi/cguqt5sppbx231t4cm961/Ajust-Image-Converter.zip?rlkey=r7d0n6e5fljcepq7mb2njmgan&st=oqsddfgy&dl=1">
  <img
    src="https://raw.githubusercontent.com/Sadc2h4/brand-assets/main/button/Download_Button_2.png"
    alt="Download .zip"
    height="48"
  />
</a>
<br>

## Features
本アプリケーションはConfluenceのバックアップZIPファイルを可読性の高いHTML形式に復元します．  
標準のConfluenceバックアップファイルはローカルサーバー環境の構築とConfluenceシステム内での閲覧を必要としますが，  
本アプリケーションでは改訂履歴，画像やドキュメントファイルを含むデータの約90%を復元可能です．

----------------------------------------------------------------------------------------------------
This Python application restores Confluence backup ZIP files into a readable HTML format.  
While standard Confluence backup files require setting up a local server environment and viewing them within the Confluence system,   
this application enables restoration of approximately 90% of the data, including revision history.

## Usage



https://github.com/user-attachments/assets/06cf6bcd-7c35-4015-b5e2-6572ede4a46e

https://github.com/user-attachments/assets/bc921688-05b5-4636-9a84-bd5ca8475e53


1. アプリケーションを起動したらロゴが表示されることを確認してください．
2. ウィンドウに直接ConfluenceのバックアップZipファイルをドラッグアンドドロップすると変換が開始します．  
   参照ボタンからファイル指定をした上で変換を実施しても実施可能です．
3. 変換中はコンソールウィンドウに変換中のファイル名と実施状況が記載されます．  
   進捗状況はプログレスバーから確認可能です．
4. ファイルの重さや種類によりますが，変換には1分程度時間が掛かります．
5. 変換が完了すると『YYYYMMDDhhmm_ファイル名』の形式で出力結果フォルダが作成されます．  
   出力場所はバックアップZipファイルと同じ階層になります．
6. 変換された情報はindex.htmlから閲覧できるほか，添付ファイルは階層ごとにフォルダ分けされて格納されます．

----------------------------------------------------------------------------------------------------
1. After launching the application, verify that the logo appears.  
2. Drag and drop the Confluence backup ZIP file directly into the window to start conversion.  
   Alternatively, you can click the Browse button to select the file before initiating conversion.  
3. During conversion, the console window displays the filenames being processed and the conversion status.  
   Progress can be monitored via the progress bar.
4. Depending on file size and type, conversion takes approximately one minute.
5. Upon completion, an output folder named in the format ‘YYYYMMDDhhmm_filename’ is created.  
   The output location is the same directory as the backup ZIP file.
6. Converted information can be viewed via index.html. Attachments are stored in separate folders organized by directory level.

## Deletion Method
・Please delete the entire file.

## Disclaimer
・I assume no responsibility whatsoever for any damages incurred through the use of this file.
B) Run as a standalone EXE
Use a prebuilt .exe if provided, or build your own with PyInstaller (see Build Notes).  

