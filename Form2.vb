Imports Ivi.Visa.Interop
Imports Microsoft.Office.Interop
Public Class Form2
    Dim i As Integer '2022oct21 add

    Dim myPos As Integer '4278ACdとDのデータ値区切り","位置
    Dim measdata As String 'GPIB測定データ
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '測定ボタン
        MsgBox("■書き込みセルをクリックしましたか？" & vbCrLf &
               "■測定端子にチップコンデンサをセットしましたか？" & vbCrLf &
               "OKで測定します。")

        richi = xlsApplication.ActiveCell.Column
        cichi = xlsApplication.ActiveCell.Row
        TextBox4.Text = richi.ToString
        TextBox5.Text = cichi.ToString

        Dim ioMgr As Ivi.Visa.Interop.ResourceManager
        Dim instrument As Ivi.Visa.Interop.FormattedIO488
        Dim idn As String

        '***************************************************************
        'High Accuracy Mode Auto Setting routine 
        'ADD 2022NOV17
        Dim HI_CRead As String
        Dim HI_CSet As String
        HI_CRead = ""
        HI_CSet = ""

        HI_CRead = xlsApplication.Cells(8, 24).Text '"150pF"

        Dim pfsakujo As Integer
        pfsakujo = Len(HI_CRead)
        Select Case pfsakujo
            Case 4
                HI_CSet = Trim(Mid(HI_CRead, 1, 2)) '"10-50pF:ex)300pF" -> "300"
                'Case 5
                'HI_CSet = Trim(Mid(HI_CRead, 1, 3)) '"100-300pF:ex)300pF" -> "300"
                'Case 6
                'HI_CSet = Trim(Mid(HI_CRead, 1, 4)) '"1000pF" -> "1000"
            Case Else
                MsgBox("EXCEL CELL READING ERROR" & vbCrLf &
                       "Check Excel CAPA Setting CELL Value" & vbCrLf &
                       "Cells(8,24)")
        End Select
        '***************************************************************
        'MsgBox(HI_CSet)

        'Dim GPIBDAT As String  'No need setting for automatically get GPIBDAT visa address  "17"
        'Dim GPIBAD As String   'No need setting for automatically get GPIBAD "GPIB::17::INSTR"
        '
        'GPIBDAT = "17"
        'GPIBDAT = Trim(TextBox1.Text)

        '******************************************************************
        '2022年9月27日(火）GPIBADはIOモニターで確認して選択するのがベター？
        'GPIBAD = "GPIB0::" & GPIBDAT & "::INSTR"  '現行SUB-STD用4278A
        'GPIBAD = "GPIB1::" & GPIBDAT & "::INSTR"
        'GPIBAD = "GPIB2::" & GPIBDAT & "::INSTR"   'テスト用4278A
        '******************************************************************

        'ioMgr = New Ivi.Visa.Interop.ResourceManager
        ioMgr = New ResourceManager
        'instrument = New Ivi.Visa.Interop.FormattedIO488
        instrument = New FormattedIO488
        instrument.IO = ioMgr.Open(GPIBAD)

        '4278A設定 2022年10月14日（金）測定時設定値　********************************
        '
        instrument.WriteString("MPAR1")     'Measurement parameter Cp-D 
        instrument.WriteString("FREQ2")     'FREQ1:1kHz,FREQ2:1MHz
        instrument.WriteString("OSC=1.0")   'OSC level setting is 1.0V

        instrument.WriteString("HIAC1") 'SM-11S96 Line 4278A also  setting ok!
        '***************************************************
        instrument.WriteString("RC=" & HI_CSet & "E-12")
        '***************************************************

        'instrument.WriteString("RB0")       'Measurement Range Setting RB0:AUTO
        instrument.WriteString("ITIM3")     'Measurement time Setting ITIM3:LONG,ITIM2:MEDIUM,ITIM1:SHORT
        instrument.WriteString("DTIM=0")    'Delay time setting DTIM=0:0msec

        instrument.WriteString("AVE=32")    'Averaging times setting AVE=32:32

        instrument.WriteString("TRIG1")     'Trigger mode setting TRIG1:INTERNAL,TRIG2:EXTERNAL,TRIG3:MANUAL
        instrument.WriteString("CABL0")     'CABLE length setting CABL0:0m ,CABL1:1m,CABL2:2m

        '****************************************************************************

        instrument.WriteString("DATA?")     'DATA OUTPUT
        idn = instrument.ReadString()       'SENDING COMMAND & RECIEVE REPLY DATA

        'MsgBox(idn)  '
        measdata = Trim(idn) '              'DELETE SPACE OF STRING OF VARIABLE 
        '                                    123456789012345678    9
        'for check                           ":DATA +15.0690E+03" & vbLf
        Dim lmeasdata As Integer
        lmeasdata = Len(measdata)

        '*****************************************************
        'デバッグモード設定用　
        superslim = 0      ' Set 1:7555MultiMeter, Set 0:4278A
        '*****************************************************
        '                                123456789012345678    9
        'for check                      ":DATA +15.0690E+03" & vbLf

        Select Case superslim
            Case 1
                sngC = Mid(measdata, 8, 18) '"15.0690E+03"
                TextBox2.Text = sngC
                sngD = 0.001
                TextBox3.Text = sngD
            Case 0
                myPos = InStr(1, measdata, ",", vbTextCompare)　'データ区切り位置
                sngC = Mid(measdata, 1, myPos - 1) 'Cd値抜き取り
                sngC = sngC * 1000000000000.0# 'pF単位に変換処理
                '      sngC = 1.234567
                TextBox2.Text = sngC
                sngD = Mid(measdata, myPos + 1) 'D値抜き取り
                sngD = sngD * 100
                '    sngD = 0.12345
                TextBox3.Text = sngD
            Case Else
        End Select

        xlsRange = xlsWorkSheet.Cells(cichi, richi)
        xlsRange.Value = sngC
        xlsRange = xlsWorkSheet.Cells(cichi + 1, richi)
        xlsRange.Value = sngD / 100 '元に戻す　REAL D vlue So that cell setting ％（1/100)display

        'tooltipにて説明に変更
        'MsgBox("測定継続→次のセルクリック→コンデンサ入替→測定” & vbCrLf & vbCrLf & "測定終了→「名前を付けて保存」")


    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '名前を付けて保存処理ボタン
        Me.Hide()
        Form1.Show()

    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        ToolTip1.ShowAlways = True

        ToolTip1.SetToolTip(Button1, "測定終了時は" & vbCrLf & "  [名前を付けて保存処理]ボタン  " & vbCrLf &
                                     "を押してください。")
        ToolTip1.SetToolTip(Button2, "データを入れるセルをクリックしてから " & vbCrLf &
                                     "測定ボタンを押して下さい。")


        '*************************************************************************
        'Automatically Get Visa Address & Visa Alias if any.
        Dim VisaCount As Integer
        VisaCount = 0

        Dim RM = New Ivi.Visa.Interop.ResourceManager
        VisaAdds = RM.FindRsrc("GPIB?*INSTR")          'input VisaAdds serch for "GPIB?*INSTR "
        GPIBAD = ""                                    'initialize the text variable "GPIBAD" 
        VisaCount = LBound(VisaAdds)

        For i = 0 To UBound(VisaAdds)                   'Maximum subscript of VisaAdds
            RM.ParseRsrcEx(VisaAdds(i), plnterfaceType, plnterfaceNumber,
                           pSessionType, pUnaliasedExpandedResourceName, pAliaslfExists)
        Next
        Me.TextBox1.Text = VisaAdds(0)
        Me.TextBox6.Text = VisaCount.ToString

        'GPIBのアドレスをグローバル変数GPIBADに設定
        'EX)設定VISAアドレス："GPIB2::17::INSTR"
        GPIBAD = VisaAdds(0)
        ' GPIBAD = VisaAdds(1)
        RM = Nothing


    End Sub
End Class