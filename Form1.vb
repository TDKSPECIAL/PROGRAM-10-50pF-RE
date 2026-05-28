Imports Microsoft.Office.Interop
Public Class Form1
    'エクセルファイルxlsm化対応　2022年10月25日
    Dim strFileName As String = "C:\ABC\DATA10-50\ANYCSIZE-3M-6M-10-50pF-4DG.xlsm"
    Dim addkakuchosi As String
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Error処理追加　2022年10月25日
        On Error GoTo ErrorHandler

        'EXCEL-OPEN-MEASUREボタン
        'Excelアプリケーション起動
        xlsApplication = New Excel.Application
        'ExcelのWorkbooks取得
        xlsWorkbooks = xlsApplication.Workbooks

        'Excel Visible = true:表示,Visible = false:非表示
        xlsApplication.Visible = True
        xlsApplication.DisplayAlerts = False

        '既存 Excel ファイルを開く
        xlsWorkbook = xlsWorkbooks.Open(strFileName)
        'Excel の Worksheets 取得
        xlsWorkSheets = xlsWorkbook.Worksheets
        'Excel の Worksheet 取得
        xlsWorkSheet = CType(xlsWorkSheets.Item(1), Excel.Worksheet)
        'シート名称
        ' xlsWorkSheet.Name = "シート名test"
        '    xlsWorkSheet.Name = "CAPACITOR-1-5pF"
        'セル選択
        '    xlsRange = xlsWorkSheet.Range("A1")
        'セルに値設定
        '   xlsRange.Value = "TEST123"

        '*******************************
        ' Public CSIZE As String
        ' Public CAPA As String
        ' Public LIMITMONTH As String
        ' Public KOSU As String
        '*******************************
        'CSIZE = xlsWorkSheet.Application.Cells(1, 1).ToString
        'CSIZE = xlsWorkSheet.Application.Range("V8").Value.ToString  '"C0603"
        'CAPA = xlsWorkSheet.Application.Range("X8").Value.ToString   '"5pF"
        'LIMITMONTH = xlsWorkSheet.Application.Range("U8").Value.ToString  '"3M" or "6M"
        '原図を修正してから読取検討

        'YEAR-MONTH-DAY 用変数　NENGAPPI
        'NENGAPPI = xlsWorkSheet.Application.Range("S6").Value.ToString '"2022/09/13"

        '個数を出すルーチンから



        Me.Hide()
        '********************************************
        'Form2表示へ移る

        Form2.ShowDialog()

        MRComObject(xlsRange)

        '//////////////////////////////////////////////////////////////////////////////
        '********************************************
        xlsApplication.DisplayAlerts = False

        '********************************************
        '保存ダイアログを開く（ボタン3のルーチン）
        'Call Button3_Click(sender, e)
        Call D_open()

        addkakuchosi = objSFD.FileName & ".xlsx"　'マクロ無し原本のBUCKUP保存ファイル名C2012～

        xlsWorkbook.SaveAs(Filename:=addkakuchosi, FileFormat:=Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook)

        xlsApplication.DisplayAlerts = True
        '//////////////////////////////////////////////////////////////////////////////

        '保存時の問合せダイアログを非表示に設定
        ' xlsApplication.DisplayAlerts = False
        'ファイルに保存 (Excel 2007～ブック形式)
        'xlsWorkbook.SaveAs(Filename:=strFileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbook)
        '保存時の問合せダイアログを表示に戻す
        'xlsApplication.DisplayAlerts = True

        '終了処理
        'xlsWorkSheet の解放
        MRComObject(xlsWorkSheet)
        'xlsWorkSheets の解放
        MRComObject(xlsWorkSheets)
        'xlsWorkbookを閉じる
        xlsWorkbook.Close(False)
        'xlsWorkbook の解放
        MRComObject(xlsWorkbook)
        'xlsWorkbooks の解放
        MRComObject(xlsWorkbooks)
        'Excelを閉じる 
        xlsApplication.Quit()
        'xlsApplication を解放
        MRComObject(xlsApplication)

        'End 処理ボタンコールして終了
        Call Button2_Click(sender, e)

ErrorHandler:
        Select Case Err.Number
            Case 1004
                MsgBox("errorNo.= " & Err.Number & "エラーが発生しました。" & vbCrLf &
                       "エクセルのプラットフォームまで開けましたが、下記の場所に読み込む” & vbCrLf &
                       "エクセル原図ファイルが有りませんでした。" &
                       vbCrLf & vbCrLf &
                       "パソコン階層フォルダ　➡　C:\ABC\DATA10-50" &
                       vbCrLf & vbCrLf &
                       "上記階層フォルダ内に ANYCSIZE-3M-6M-10-50pF-4DG.xls " & vbCrLf & vbCrLf &
                       "エクセルファイル原図の設置保管を確認してください。" & vbCrLf &
                       "このエラーメッセージボックスを閉じてから、保管処理解決ののち" & vbCrLf &
                       "このままSUB-STD測定プログラムを継続使用できます。" & vbCrLf & vbCrLf &
                       "エクセル原図保管設置OK後、エクセルOPEN測定開始ボタンで処理再開できます。")

                '***************************************************************
                '開いたところまでのbookをメモリから開放する処理
                MRComObject(xlsWorkbooks)
                'Excelを閉じる 
                xlsApplication.Quit()
                'xlsApplication を解放
                MRComObject(xlsApplication)
                '***************************************************************
                Exit Sub

            Case Else
                MsgBox("errorNo.= " & Err.Number & " " & Err.Description)

                '***************************************************************
                '開いたところまでのbookをメモリから開放する処理
                MRComObject(xlsWorkbooks)
                'Excelを閉じる 
                xlsApplication.Quit()
                'xlsApplication を解放
                MRComObject(xlsApplication)
                '***************************************************************
                Exit Sub
        End Select


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        End

    End Sub
End Class
