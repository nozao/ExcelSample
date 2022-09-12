## はじめに

まず初めに、このサンプルプログラムの動作を説明します。  
また、実際のマクロファイルとダミーのCSVデータは[https://github.com/nozao/ExcelSample/ExtractCSV/SampleCode.zip](https://github.com/nozao/ExcelSample/ExtractCSV/SampleCode.zip)からダウンロードできます。  
実際のサンプルプログラムは一番下にあります  
サンプルプログラムの詳細解説は別のテキストに記述します。

## サンプルプログラムの前提環境
### CSVファイル構成
こんな感じで規則性のあるファイル名をつけられたcsvファイルがフォルダ内にまとまって入っています。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/153238/22326ced-ca96-ec66-898d-9e7713f2c68a.png)

`DeviceLog_`の後ろの6桁数値は西暦4桁＋月(2桁)となっています。

<div style="page-break-before:always"></div>

### CSVファイルの中身
試しに1つ開いてみると中身はこのようなデータになっています。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/153238/765b6bc1-72d2-4222-cdda-30e74b7bfeb0.png)

左から計測日時、設備No、測定した結果です。
今回は1つのファイルにそれぞれ20行づつ格納されていることとします。

### 処理対象ファイル
今回のサンプルマクロでは、各年の1月分のファイルだけを処理対象とします。
つまり、抽出処理されるファイルは下記の3ファイルだけです。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/153238/29b0bb35-5793-4544-225b-fabe2977d503.png)

これはファイル名から判断しています。

<div style="page-break-before:always"></div>  

### 言い訳
今回はサンプルプログラムなのでエラー処理や設定周りなどをあまり丁寧に作っていません。
コピペ部分も横着してSelectionを使ったので、マクロ実行中にExcelをクリックなどするとおかしな動作をする可能性があります。

## マクロ動作
### マクロ本体の説明
マクロ本体はこんな感じです。
`青枠`のボタンを押すとCSVを格納しているフォルダの場所を指定できます。
指定が終わったら`赤枠`のボタンでマクロを実行します。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/153238/b0a4989e-7ad9-d5b5-6d7f-bbfc48c5ae38.png)

<div style="page-break-before:always"></div>  
  

### 動作結果
動作させるとこのようになります。
右半分が集計結果です。
今回の処理対象ファイル3つ x 各ファイルに20行のデータがあったので合計60行までデータが集計されています。

![image.png](https://qiita-image-store.s3.ap-northeast-1.amazonaws.com/0/153238/2b4ae27c-96d4-9d09-c8c1-94a8713a52d1.png)

## マクロコード
下記にマクロコードを示します。
メインは中段くらいにある`Public Function MainProccess()`の部分です。
長いように見えますがコメントを入れまくっているせいです。

フォーム上のボタンと連動させる部分は記入していません。
```vb
'=============================各種設定====================
'このマクロの設定用シート名
Const SETTING_SHEET = "CSV抽出マクロ"
'CSVをコピーする範囲(固定の場合)
Const STATIC_COPY_RANGE = "B2:C21"
'処理対象にするファイル名のパターン
Const FILENAME_FILTER = "DeviceLog_####01.csv"
'フォルダの場所を取得して現在のシートの指定した場所に結果を出力する関数
Public Function GetTargetFolderPath(RowNumber As Long, ColumnNumber As Long)
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            ThisWorkbook.Sheets(SETTING_SHEET).Cells(RowNumber, ColumnNumber).Value = .SelectedItems(1)
        Else
            ThisWorkbook.Sheets(SETTING_SHEET).Cells(RowNumber, ColumnNumber).Value = ""
        End If
    End With
End Function

'メイン処理部分 
Public Function MainProccess()

    Dim TargetCSV As Workbook
    Dim ResultBook As Workbook
    
    '結果集計用のファイルを新規作成
    Set ResultBook = Workbooks.Add
    'ファイルオープンダイアログを開いてフォルダ情報を取得する準備
    Dim objFSO As Object, objFolder As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    '指定されたフォルダを掴む
    Set objFolder = objFSO.Getfolder(ThisWorkbook.Sheets(SETTING_SHEET).Cells(2, 2).Value)
    
    '結果集計シートの現在の最終行数カウンタ
    Dim ResultLastRow As Long
    '初期値として1行目を設定しておく
    ResultLastRow = 1
    
    Dim CurrentFile
    '指定されたフォルダ下にあるすべてのファイルをチェック
    For Each CurrentFile In objFolder.Files
        'ファイル名がDeviceLog_####01.csvの規則に合致するときだけ処理する
        If CurrentFile.Name Like FILENAME_FILTER Then
            'フリーズしたように見える現象の対策
            DoEvents
            '対象のCSVを開く
            Set TargetCSV = Workbooks.Open(CurrentFile.Path)
            '======== 固定範囲をコピペする場合の処理。変動する範囲(2,3列目の最初から最後までとか)をコピペしたい場合はここを変更する必要がある。======
            '固定範囲をコピーする。CSVは必ずSheetが1枚しかないのでSheets(1)で指定できるはず。
            TargetCSV.Sheets(1).Range(STATIC_COPY_RANGE).Copy
            '結果集計用ファイルをアクティブにする
            ResultBook.Activate
            '値のみ貼り付け。結果集計用ファイルもこのマクロ内で作成したのでシート名を指定せず、1枚目のシート(Sheets(1))という指定をする。
            'B列から貼り付ける。A列にはファイル名を入れるため。
            ActiveWorkbook.Sheets(1).Cells(ResultLastRow, 2).PasteSpecial Paste:=xlPasteValues
            'A列にファイル名入れる
            ActiveWorkbook.Sheets(1).Range("A" & ResultLastRow & ":A" & ResultLastRow + Selection.Rows.Count - 1).Value = CurrentFile.Name
            'コピペした行数を最終行数カウンタに足しておく
            ResultLastRow = ResultLastRow + Selection.Rows.Count
            '==========固定範囲コピペ処理部分はここまで
                        
            '使い終わったCSVを閉じる。閉じるときに「保存しますか？」とか出てきてマクロが止まってしまうので
            '確認画面や警告を出さないように抑制→ファイル閉じる→抑制解除を行う
            Application.DisplayAlerts = False
            TargetCSV.Close
            Application.DisplayAlerts = True
        End If
    Next
End Function
```
