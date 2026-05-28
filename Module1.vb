
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices

Module Module1
    '******************************************************************************
    'プログラム名：PROGRAM-10-50pF
    '使用言語   :VISUAL STUDIO 2022 VB.net
    'OS         :Windows 10 Pro
    'Version    :21H2
    'System     :64bit operationsystem
    'CPU        :Inter(R)Core(TM)i5-8250U CPU @1.6GHz
    'Memory     :8.00GB
    '参照       ：VISA COM488.2 Formatted I/O 5.13     version 5.13
    '参照       ：VISA COM 5.13 Type Library             version 5.13
    'Programed by H.W on 20.oct.2022
    '*****************************************************************************
    '【改訂履歴】
    '■Changed 2022.oct.21:自動VISA ADDRESS取得ルーチンに変更　確認して設定する事無し
    'でVISAアドレス自動指定可能
    '■Changed 2022.oct 25:①Form1のクラス中のエクセルファイル読み込み時のエラー処理
    '追加。②Form2中のCp,D表示用テキストボックス、Row,Column表示用テキストボックスの
    'イネーブルfalse化。③Module1中のMSGボックスの表示修正変更、”に”取り、ファイル
    '名の下にスペース行1行追加
    '■changed 2022.nov.17:High Accuracy mode value auto Setting coding
    '
    '
    '******************************************************************************
    'Excel のアプリケーション参照用オブジェクト 
    '名前空間により省略記載 （１例ｻﾝﾌﾟﾙ）
    'Public xlsApplication As Microsoft.Office.Interop.Excel.Application = Nothing

    Public xlsApplication As Excel.Application = Nothing
    'Excel の Workbooks 参照用オブジェクト (Workbook の Collection)
    Public xlsWorkbooks As Excel.Workbooks = Nothing
    'Excel の Workbooks 内の1個の Workbook 参照用オブジェクト
    Public xlsWorkbook As Excel.Workbook = Nothing
    'Excel の Workbook 内の Worksheets 参照用オブジェクト (Worksheet の Collection)
    Public xlsWorkSheets As Excel.Sheets = Nothing
    'Excel の Sheets 内の1個の Worksheet 参照用オブジェクト
    Public xlsWorkSheet As Excel.Worksheet = Nothing
    'Excel の Sheet 内の1個のセル Range 参照用オブジェクト
    Public xlsRange As Excel.Range = Nothing
    '******************************************************************************

    'Form2での測定時にエクセルのCOLUMN,ROW
    Public cichi As Integer 'COLUMN位置
    Public richi As Integer 'ROW位置

    Public sngC As Single 'Cd  Form1,2
    Public sngD As Single 'D   Form1,2

    Public superslim As Integer '実験用切替　MY測定器時:1、4278A時:0

    'ANYCSIZE-3M-6M-1-5pF-5DG.xlsの各設定後のセル値を取り込む変数
    Public CSIZE As String
    Public CAPA As String
    Public LIMITMONTH As String
    Public KOSU As String
    Public R_NENGAPPI As String 'セル読み込み年月日

    Public NENGAPPI As String　'ファイル名構成用年月日
    Public Yearda As String 'year 年4桁　"2022"　ファイル名の年構成用
    Public monthda As String 'month 月2桁 "01～12" "09"9月　ファイル名の月構成用
    Public dayda As String 'day 日2桁 "01～31" "22"22日 ファイル名の日構成用
    Public MATOMENENN As String 'MATOMENENN = yearda & monthda & dayda

    Public automatic As Integer '0：手動ダイアログファイル名入力、1：自動ダイアログファイル名入力

    Public IRAIMOTO As String 'コンボボックス3からの依頼元情報
    Public D_FILENAME As String 'ダイアログに表示する保存用ファイル名
    Public Disp_filename As String 'ダイアログ用ファイルネーム

    'FOR Dialog subroutine
    Public objSFD As New SaveFileDialog()

    'for Automatic Visa Address Get 2022/oct/21 add
    Public VisaAdds() As String
    Public plnterfaceType As Integer
    Public plnterfaceNumber As Integer
    Public pSessionType As String
    Public pUnaliasedExpandedResourceName As String
    Public pAliaslfExists As String
    Public GPIBAD As String
    Public setVisaAddress As String '実際のコマンドにて使用するときに使用

    Public Sub MRComObject(ByRef objCom As Object)　'Objectの解放処理
        'If Not objCom Is Nothing Then ’ﾒｯｾｰｼﾞの候補（次行）に修正　'2022年9月27日（火）
        If objCom IsNot Nothing Then
            Try
                'System.Runtime.InteropServices.Marshal.ReleaseComObject(objCom)
                Marshal.ReleaseComObject(objCom)
            Catch
                '
            Finally
                '参照を解除する
                objCom = Nothing
            End Try
        End If
    End Sub

    Public Sub D_open()
        '*****************************************************************
        'automatic→1 自動ダイアログファイル名　下行
        'automatic→1
        '*****************************************************************
        'automatic→0 手動エクセルファイル名ダイアログ入力　下行
        automatic = 1
        '*****************************************************************

        'Button1.Visible = False 'FORM1上に表示しない　→プロパティ設定済みの為不要

        ' Dim objSFD As New SaveFileDialog()　
        'ファイルダイアログのウィンドウタイトル
        objSFD.Title = "EXCELファイルの名前を付けて保存"
        'ダイアログ初期表示のディレクトリ
        objSFD.InitialDirectory = "C:\ABC\DATA10-50\"
        'ファイルの種類に表示される拡張子を指定
        objSFD.Filter = "EXCELファイル(*.xls)|*.xls|全てのファイル(*.*)|*.*"
        'ファイルの種類のリストで初期表示されるもの（規定は１）
        objSFD.FilterIndex = 2
        '前回開いたディレクトリを復元（ただし、InitialDirectoryが優先される）
        objSFD.RestoreDirectory = True

        '***************************************************************************

        If automatic = 1 Then '自動ダイアログ構成

            '*******************************
            ' Public CSIZE As String
            ' Public CAPA As String
            ' Public LIMITMONTH As String
            ' Public KOSU As String
            '*******************************
            'CSIZE = xlsWorkSheet.Application.Cells(1, 1).ToString
            CSIZE = xlsWorkSheet.Application.Range("V8").Value.ToString  '"C0603"
            Debug.WriteLine(CSIZE)
            CAPA = xlsWorkSheet.Application.Range("X8").Value.ToString   '"5pF"
            Debug.WriteLine(CAPA)
            LIMITMONTH = xlsWorkSheet.Application.Range("U8").Value.ToString  '"3M" or "6M"
            Debug.WriteLine(LIMITMONTH)
            '原図を修正してから読取検
            'YEAR-MONTH-DAY 用変数　NENGAPPI
            'NENGAPPI = xlsWorkSheet.Application.Range("S6").Value.ToString '"2022/09/13"
            R_NENGAPPI = xlsWorkSheet.Application.Range("S6").Value.ToString '"2022/09/13"
            NENGAPPI = Mid(R_NENGAPPI, 1, 10) '時間もあるので年月日のみ抜き取り
            Debug.WriteLine(NENGAPPI)
            '--------------------------------FILE年用　年月日再構成 -------------------------
            '"2022/09/22"から　"2022","09","22"を取り出し再構成で　"20220922"とする
            Yearda = Mid(NENGAPPI, 1, 4)  '年4桁抜き出し
            monthda = Mid(NENGAPPI, 6, 2)　'月2桁抜き出し
            dayda = Mid(NENGAPPI, 9, 2)　'日2桁抜き出し

            MATOMENENN = Yearda & monthda & dayda '"20220922"
            Debug.WriteLine(MATOMENENN)

            '/////////////////////////////////////////////////////////////////////////
            'IRAIMOTO情報はダイアログを開いた時に読み込む（最新変更反映させる為）
            IRAIMOTO = xlsWorkSheet.Application.Range("AA8").Value.ToString '依頼元設定情報
            Debug.WriteLine(IRAIMOTO)

            '個数を出すルーチンから
            KOSU = xlsWorkSheet.Application.Range("Z8").Value.ToString '5個の場合 "5"
            Debug.WriteLine(KOSU)

            '最終のファイル構成もダイアログを開いたときに読み込み構成する
            'D_FILENAME 拡張子無しのファイル名のみ
            D_FILENAME = CSIZE & "-" & CAPA & "-" & LIMITMONTH & "-" & KOSU & "PCS" & "-" & IRAIMOTO & "-" & MATOMENENN
            Debug.WriteLine(D_FILENAME)
            '/////////////////////////////////////////////////////////////////////////

            '*************************************************************************
            'D_FILENAMEは拡張子無しのファイル名のみ　
            '自動で、このファイル名に拡張子　".xls"付で保存される。
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            objSFD.FileName = D_FILENAME 'ダイアログに表示　rem解除で自動ファイル名構成
            '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
        Else
        End If

        If objSFD.ShowDialog() = DialogResult.OK Then
            'MessageBox.Show("保存先" & vbCrLf &
            'objSFD.FileName & vbCrLf &
            '               "の名前を付けて保存します。")

            Dim dnagasadir As Integer
            dnagasadir = Len(objSFD.InitialDirectory)
            Dim dnagasadirname As Integer
            dnagasadirname = Len(objSFD.FileName)
            Dim sabun As Integer
            sabun = dnagasadirname - dnagasadir

            Dim Disp_filename As String
            Disp_filename = Mid(objSFD.FileName, dnagasadir + 1, sabun)

            If automatic = 1 Then '自動ダイアログファイル名構成
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                Disp_filename = D_FILENAME
                '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
            Else
            End If

            'MessageBox.Show("保存先パス " & objSFD.InitialDirectory & "に" & vbCrLf &
            '                "ファイル名" & vbCrLf &
            '                "【" & Disp_filename & ".xlsx】" & vbCrLf &
            '                "で保存します。")

            MessageBox.Show("保存先パス   " & objSFD.InitialDirectory & vbCrLf & vbCrLf &
                            "ファイル名 ：" & vbCrLf &
                            "【" & Disp_filename & ".xlsx】" & vbCrLf & vbCrLf &
                            "で保存します。")

        End If
    End Sub


End Module
