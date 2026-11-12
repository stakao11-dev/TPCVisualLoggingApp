Attribute VB_Name = "Module1"
Option Explicit
'ここにはシステム変更ボタンとグラフ描画ボタンがあります。
Public Const SHEETNAME1    As String = "ﾛｷﾞﾝｸﾞﾃﾞｰﾀ"
Public Const SHEETNAME2    As String = "測定結果"
Public Const CHARTNAME11   As String = "第１列データ速度"
Public Const CHARTNAME12   As String = "第２列データ速度"
Public Const CHARTNAME13   As String = "第1,2列データ速度"
Public Const CHARTNAME21   As String = "円全体"
Public Const CHARTNAME31   As String = "軌跡"
Public Const CHARTNAME32   As String = "軸別直線精度"
Public Const BUTTONSTR1    As String = "グラフ表示設定変更"
Public Const BUTTONSTR2    As String = "設定項目変更"
Public Const BUTTONSTR3    As String = "グラフ印刷"
Public Const FILENAMEPOINT As String = "g4"
Public Const RTCPOINT      As String = "h6"
Public Const RPOINT        As String = "h6"
Public Const XCENTER       As String = "h7"
Public Const YCENTER       As String = "h8"
Public Const AMPPOINT      As String = "h9"
Public Const STARTCSVX     As String = "a3"
Public Const STARTCSVY     As String = "b3"
'Public Const ENDCSVX       As String = "a4097"
'Public Const ENDCSVY       As String = "b4097"
'Public Const ENDCSVX       As String = "A31999"
'Public Const ENDCSVY       As String = "B31999"
Public Const RXB           As String = "g7"
Public Const RXE           As String = "g8"
Public Const RYB           As String = "h7"
Public Const RYE           As String = "h8"
Public Const STARTPOINT    As String = "h10"
Public Const ENDPOINT      As String = "h11"
Public Const CSVXCP        As String = "a3"
Public Const CSVYCP        As String = "b3"
Public Const ERRORX0       As String = "c3"
Public Const ERRORY0       As String = "d3"
Public Const ERRORL0       As String = "e3"
Public Const SUMPLINGN     As String = "g3"
Public Const XPOINT        As String = "h3"
Public Const YPOINT        As String = "i3"
Public Const ERRORX1       As String = "j3"
Public Const ERRORY1       As String = "k3"
Public Const ERRORL1       As String = "l3"
Public Const RRTCPOINT     As String = "aa20"
Public Const RRPOINT       As String = "ab20"
Public Const RXCENTER      As String = "ab21"
Public Const RYCENTER      As String = "ab22"
Public Const RAMPPOINT     As String = "ab23"
Public Const RRXB          As String = "ac20"
Public Const RRXE          As String = "ac21"
Public Const RRYB          As String = "ac22"
Public Const RRYE          As String = "ac23"
Public Const RMABIKIPOINT  As String = "ac24"
Public Const RSTARTPOINT   As String = "ac25"
Public Const RENDPOINT     As String = "ac26"
Public Const TXTBOXNAME11  As String = "ad20"
Public Const TXTBOXNAME12  As String = "ad21"
Public Const TXTBOXNAME13  As String = "ad22"
Public Const TXTBOXNAME21  As String = "ae20"
Public Const TXTBOXNAME31  As String = "af20"
Public Const TXTBOXNAME32  As String = "af21"
Public Const TXTBOXEN11    As String = "ad25"
Public Const TXTBOXEN12    As String = "ad26"
Public Const TXTBOXEN13    As String = "ad27"
Public Const TXTBOXEN21    As String = "ae25"
Public Const TXTBOXEN31    As String = "af25"
Public Const TXTBOXEN32    As String = "af26"
Public Const MBTNNAME00    As String = "設定画面へ"
Public Const MBTNNAME11    As String = "第1列グラフへ"
Public Const MBTNNAME12    As String = "第2列グラフへ"
Public Const MBTNNAME13    As String = "第1,2列グラフへ"
Public Const MBTNNAME21    As String = "円全体グラフへ"
Public Const MBTNNAME31    As String = "軌跡グラフへ"
Public Const MBTNNAME32    As String = "軸別直線精度グラフへ"
Public Const TXTBOXSTR11  As String = "ad6000"
Public Const TXTBOXSTR12  As String = "ad6001"
Public Const TXTBOXSTR13  As String = "ad6002"
Public Const TXTBOXSTR21  As String = "ae6000"
Public Const TXTBOXSTR31  As String = "af6000"
Public Const TXTBOXSTR32  As String = "af6001"
Public Const TXTBOXNAME111  As String = "ad30"
Public Const TXTBOXNAME121  As String = "ad31"
Public Const TXTBOXNAME131  As String = "ad32"
Public Const ERRTXT1 As String = "※量子化誤差"
Public Const ERRTXT2 As String = "pps"
'Rowが4096から50000に増やしPublicに定数化する
Public Const STARTROW As Integer = 3
Public Const ENDROW As Double = 32000
Public Const ENDADD1 As Double = ENDROW + 1
Public Const ENDADD3 As Double = ENDROW + 3

Sub chg_sys()
With UserForm1
  If (1 = sheet_chk()) Then
   Select Case ActiveSheet.Cells(5, 6).Value
    Case "サーボ応答精度"
      .OptionButton1.Value = True
    Case "真円精度"
      .OptionButton2.Value = True
    Case "直線精度"
      .OptionButton3.Value = True
    Case Else
      .OptionButton1.Value = False
      .OptionButton2.Value = False
      .OptionButton3.Value = False
   End Select
  Else
      .OptionButton1.Value = False
      .OptionButton2.Value = False
      .OptionButton3.Value = False
  End If
  .Show
End With
End Sub
Function prmchk(sysname As String) As Integer
Dim ret As Integer
 With Worksheets(SHEETNAME1)
  Select Case sysname
   Case "サーボ応答精度"
    If IsEmpty(.Range(RTCPOINT).Value) Then
     ret = 0
    Else
     ret = 1
    End If
   Case "真円精度"
    If ((IsEmpty(.Range(RPOINT).Value)) Or _
       (IsEmpty(.Range(XCENTER).Value)) Or _
       (IsEmpty(.Range(YCENTER).Value)) Or _
       (IsEmpty(.Range(AMPPOINT).Value))) Then
        ret = 0
    Else
     ret = 2
    End If
   Case "直線精度"
    If ((IsEmpty(.Range(RXB).Value)) Or _
       (IsEmpty(.Range(RXE).Value)) Or _
       (IsEmpty(.Range(RYB).Value)) Or _
       (IsEmpty(.Range(RYE).Value)) Or _
       (IsEmpty(.Range(STARTPOINT).Value)) Or _
       (IsEmpty(.Range(ENDPOINT).Value))) Then
        ret = 0
    Else
     ret = 3
    End If
  End Select
 End With
prmchk = ret
End Function
Sub draw_graph()
Dim csvfname As String, sysnum As Integer
 sysnum = prmchk(ActiveSheet.Cells(5, 6).Value)
 If (0 < sysnum) Then
  csvfname = Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value
  Select Case Worksheets(SHEETNAME1).Cells(5, 6).Value
   Case "サーボ応答精度"
    Call speedckeck(csvfname)
   Case "真円精度"
    Call circleckeck(csvfname)
   Case "直線精度"
    Call linearity_chk(csvfname)
   Case Else
    MsgBox "システム名が間違っています"
  End Select
 Else
  MsgBox "パラメータに未設定のものがあります。"
 End If
End Sub

Sub retain_data()
Dim sysname As String
 With Worksheets(SHEETNAME1)             'シートの選択
  sysname = .Cells(5, 6).Value
  If ("サーボ応答精度" = sysname) Then
   .Range(RRTCPOINT).Value = .Range(RTCPOINT).Value
  ElseIf ("真円精度" = sysname) Then
   .Range(RRPOINT).Value = .Range(RPOINT).Value
   .Range(RXCENTER).Value = .Range(XCENTER).Value
   .Range(RYCENTER).Value = .Range(YCENTER).Value
   .Range(RAMPPOINT).Value = .Range(AMPPOINT).Value
  ElseIf ("直線精度" = sysname) Then
   '自動計算が走るので不要
  End If
  .Range(RRTCPOINT & ":" & RENDPOINT).Font.ColorIndex = 2
 End With
End Sub

Sub load_retain_data()
Dim sysname As String
 With Worksheets(SHEETNAME1)             'シートの選択
  sysname = .Cells(5, 6).Value
  If ("サーボ応答精度" = sysname) Then
   .Range(RTCPOINT).Value = .Range(RRTCPOINT).Value
  ElseIf ("真円精度" = sysname) Then
   .Range(RPOINT).Value = .Range(RRPOINT).Value
   .Range(XCENTER).Value = .Range(RXCENTER).Value
   .Range(YCENTER).Value = .Range(RYCENTER).Value
   .Range(AMPPOINT).Value = .Range(RAMPPOINT).Value
  ElseIf ("直線精度" = sysname) Then
   '自動計算が走るので不要
  End If
 End With
End Sub
Sub form3_ini()
Dim txboxname_str As String, retaintxboxname As String, i As Integer
Dim retaintxboxvisible As String, retaintxboxstr As String

With UserForm3
  .TextBox1.Value = ActiveChart.Axes(xlCategory).MinimumScale
  .TextBox2.Value = ActiveChart.Axes(xlCategory).MaximumScale
  .TextBox3.Value = ActiveChart.Axes(xlValue).MinimumScale
  .TextBox4.Value = ActiveChart.Axes(xlValue).MaximumScale

  .CheckBox1.Value = ActiveChart.Axes(xlCategory).MinimumScaleIsAuto
  .CheckBox3.Value = ActiveChart.Axes(xlCategory).MaximumScaleIsAuto
  .CheckBox2.Value = ActiveChart.Axes(xlValue).MinimumScaleIsAuto
  .CheckBox4.Value = ActiveChart.Axes(xlValue).MaximumScaleIsAuto
  .CheckBox5.Value = ActiveChart.Axes(xlValue).HasMajorGridlines
  .CheckBox6.Value = ActiveChart.Axes(xlValue).HasTitle
  
  'コメント挿入前処理(関数化したい)
    txboxname_str = ""
    If (ActiveChart.Name = CHARTNAME11) Then
     retaintxboxname = TXTBOXNAME11
     retaintxboxvisible = TXTBOXEN11
     retaintxboxstr = TXTBOXSTR11
    ElseIf (ActiveChart.Name = CHARTNAME12) Then
     retaintxboxname = TXTBOXNAME12
     retaintxboxvisible = TXTBOXEN12
     retaintxboxstr = TXTBOXSTR12
    ElseIf (ActiveChart.Name = CHARTNAME13) Then
     retaintxboxname = TXTBOXNAME13
     retaintxboxvisible = TXTBOXEN13
     retaintxboxstr = TXTBOXSTR13
    ElseIf (ActiveChart.Name = CHARTNAME21) Then
     retaintxboxname = TXTBOXNAME21
     retaintxboxvisible = TXTBOXEN21
     retaintxboxstr = TXTBOXSTR21
    ElseIf (ActiveChart.Name = CHARTNAME31) Then
     retaintxboxname = TXTBOXNAME31
     retaintxboxvisible = TXTBOXEN31
     retaintxboxstr = TXTBOXSTR31
    ElseIf (ActiveChart.Name = CHARTNAME32) Then
     retaintxboxname = TXTBOXNAME32
     retaintxboxvisible = TXTBOXEN32
     retaintxboxstr = TXTBOXSTR32
    End If
    
    If (Not IsEmpty(Worksheets(SHEETNAME1).Range(retaintxboxname))) Then
      txboxname_str = Worksheets(SHEETNAME1).Range(retaintxboxname).Value
      i = 1
      Do While i <= ActiveChart.Shapes.Count
       If (ActiveChart.Shapes.Item(i).Name = txboxname_str) Then
        i = 0
        Exit Do
       End If
       i = i + 1
      Loop
      If (i <> 0) Then
       txboxname_str = ""
      End If
    End If
    If (txboxname_str <> "") Then
     i = ActiveChart.Shapes.Item(txboxname_str).Visible
     ActiveChart.Shapes.Item(txboxname_str).Visible = msoCTrue
     ActiveChart.Shapes.Item(txboxname_str).Select
     .TextBox5.Text = Selection.Characters.Text
     ActiveChart.Shapes.Item(txboxname_str).Visible = i
    Else
     Worksheets(SHEETNAME1).Range(retaintxboxname).Value = ""
     Worksheets(SHEETNAME1).Range(retaintxboxvisible).Value = 0
     .TextBox5.Text = Worksheets(SHEETNAME1).Range(retaintxboxstr).Value
    End If
   If (0 = Worksheets(SHEETNAME1).Range(retaintxboxvisible).Value) Then
    .CheckBox6.Value = False
   Else
    .CheckBox6.Value = True
   End If
   .TextBox5.Enabled = .CheckBox6.Value
  .Show
  End With
End Sub
Function sheet_chk()
Dim sheetnum As Integer, chartnum As Integer
Dim csvfname As String, sysnum As Integer
Dim cnt As Integer, chkflg As Integer
Dim name_str As String, sysname As String, ret As Integer

ret = 0
With Sheets(SHEETNAME1)
 '各解析方法選択時にシート/グラフがあるかチェック。
 sysname = .Cells(5, 6).Value
 sheetnum = Application.Worksheets.Count
 chartnum = Application.Charts.Count
 chkflg = 0
 If (2 <= sheetnum) Then
   cnt = 1
   Do While (cnt <= sheetnum)
    name_str = ThisWorkbook.Sheets.Item(cnt).Name
    If ((Not (chkflg And 1)) And (SHEETNAME1 = name_str)) Then
     chkflg = chkflg Or 1
    ElseIf ((Not (chkflg And 2)) And (SHEETNAME2 = name_str)) Then
      chkflg = chkflg Or 2
    End If
    cnt = cnt + 1
   Loop
 End If
 If ("サーボ応答精度" = sysname) Then
  If (3 <= chartnum) Then
   cnt = 1
   Do While (cnt <= chartnum)
    name_str = ThisWorkbook.Charts.Item(cnt).Name
    If ((Not (chkflg And 4)) And (CHARTNAME11 = name_str)) Then
     chkflg = chkflg Or 4
    ElseIf ((Not (chkflg And 8)) And (CHARTNAME12 = name_str)) Then
      chkflg = chkflg Or 8
    ElseIf ((Not (chkflg And 16)) And (CHARTNAME13 = name_str)) Then
      chkflg = chkflg Or 16
    End If
    cnt = cnt + 1
   Loop
   If (31 = chkflg) Then
    ret = 1
   End If
  End If
 ElseIf ("真円精度" = sysname) Then
  If (1 <= chartnum) Then
   cnt = 1
   Do While (cnt <= chartnum)
    name_str = ThisWorkbook.Charts.Item(cnt).Name
    If ((Not (chkflg And 4)) And (CHARTNAME21 = name_str)) Then
     chkflg = chkflg Or 4
    End If
    cnt = cnt + 1
   Loop
   If (7 = chkflg) Then
    ret = 1
   End If
  End If
 ElseIf ("直線精度" = sysname) Then
  If (2 <= chartnum) Then
   cnt = 1
   Do While (cnt <= chartnum)
    name_str = ThisWorkbook.Charts.Item(cnt).Name
    If ((Not (chkflg And 4)) And (CHARTNAME31 = name_str)) Then
     chkflg = chkflg Or 4
    ElseIf ((Not (chkflg And 8)) And (CHARTNAME32 = name_str)) Then
      chkflg = chkflg Or 8
    End If
    cnt = cnt + 1
   Loop
   If (15 = chkflg) Then
    ret = 1
   End If
  End If
 End If
End With
 sheet_chk = ret
End Function
Sub graph_print()
    If (vbYes = MsgBox(ActivePrinter & Chr(10) & "で印刷します。よろしいですか", vbYesNo)) Then
     SendKeys "{esc}", True
     Sheets(ActiveSheet.Name).PrintOut
    End If
End Sub
Sub move_to_csv()
 Sheets(SHEETNAME1).Activate
End Sub
Sub move_to_c11()
  If (1 = sheet_chk()) Then
     Sheets(CHARTNAME11).Activate
  End If
End Sub
Sub move_to_c12()
 If (1 = sheet_chk()) Then
  Sheets(CHARTNAME12).Activate
 End If
End Sub
Sub move_to_c13()
 If (1 = sheet_chk()) Then
  Sheets(CHARTNAME13).Activate
 End If
End Sub
Sub move_to_c21()
 If (1 = sheet_chk()) Then
  Sheets(CHARTNAME21).Activate
 End If
End Sub
Sub move_to_c31()
 If (1 = sheet_chk()) Then
  Sheets(CHARTNAME31).Activate
 End If
End Sub
Sub move_to_c32()
 If (1 = sheet_chk()) Then
  Sheets(CHARTNAME32).Activate
 End If
End Sub
Sub Data_Clear()
 Dim i As Integer
 Application.DisplayAlerts = False
 Sheets(SHEETNAME1).Select
 ActiveSheet.Range("a3:b65536").Select
 Selection.ClearContents
 ActiveSheet.Range("a3").Select
 Sheets(SHEETNAME2).Select
 ActiveSheet.Range("a3:u65536").Select
 Selection.ClearContents
 ActiveSheet.Range("a3").Select
 For i = Sheets.Count To 1 Step -1
  If Not (Sheets(i).Name = SHEETNAME1 Or Sheets(i).Name = SHEETNAME2) Then
  Sheets(i).Delete
  End If
 Next i
 Sheets(SHEETNAME1).Activate
 Sheets(SHEETNAME1).CommandButton2.Enabled = False
 Sheets(SHEETNAME1).CommandButton3.Enabled = False
 Sheets(SHEETNAME1).Range("a3").Select
 Application.DisplayAlerts = True
End Sub
Sub cmment_insert(ChartName As String)
Dim txboxname_str As String, retaintxboxname As String, i As Integer
Dim retaintxboxvisible As String, retaintxboxstr As String

    txboxname_str = ""
    If (ChartName = CHARTNAME11) Then
     retaintxboxname = TXTBOXNAME11
     retaintxboxvisible = TXTBOXEN11
     retaintxboxstr = TXTBOXSTR11
    ElseIf (ChartName = CHARTNAME12) Then
     retaintxboxname = TXTBOXNAME12
     retaintxboxvisible = TXTBOXEN12
     retaintxboxstr = TXTBOXSTR12
    ElseIf (ChartName = CHARTNAME13) Then
     retaintxboxname = TXTBOXNAME13
     retaintxboxvisible = TXTBOXEN13
     retaintxboxstr = TXTBOXSTR13
    ElseIf (ChartName = CHARTNAME21) Then
     retaintxboxname = TXTBOXNAME21
     retaintxboxvisible = TXTBOXEN21
     retaintxboxstr = TXTBOXSTR21
    ElseIf (ChartName = CHARTNAME31) Then
     retaintxboxname = TXTBOXNAME31
     retaintxboxvisible = TXTBOXEN31
     retaintxboxstr = TXTBOXSTR31
    ElseIf (ChartName = CHARTNAME32) Then
     retaintxboxname = TXTBOXNAME32
     retaintxboxvisible = TXTBOXEN32
     retaintxboxstr = TXTBOXSTR32
    End If
    
    If (Not IsEmpty(Worksheets(SHEETNAME1).Range(retaintxboxname))) Then
      txboxname_str = Worksheets(SHEETNAME1).Range(retaintxboxname).Value
      i = 1
      Do While i <= Charts(ChartName).Shapes.Count
       If (Charts(ChartName).Shapes.Item(i).Name = txboxname_str) Then
        i = 0
        Exit Do
       End If
       i = i + 1
      Loop
      If (i <> 0) Then
       txboxname_str = ""
      End If
    End If

    If Worksheets(SHEETNAME1).Range(retaintxboxvisible).Value Then
     If (txboxname_str = "") Then
      Charts(ChartName).Shapes.AddTextbox(msoTextOrientationHorizontal, 230.25, 70.25, 7.2, 7.2).Select
      Selection.ShapeRange.ScaleWidth 6.91, msoFalse, msoScaleFromTopLeft
     Else
      Charts(ChartName).Shapes.Item(txboxname_str).Visible = msoCTrue
      Charts(ChartName).Shapes(txboxname_str).Select
     End If
     Selection.Characters.Text = Worksheets(SHEETNAME1).Range(retaintxboxstr).Value
     Selection.Characters.Text = Replace(Selection.Characters.Text, Chr(13), "")
     Selection.AutoScaleFont = False
     Selection.AutoSize = True

    End If

End Sub
