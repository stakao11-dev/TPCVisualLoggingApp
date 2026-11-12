Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{AC0A6D7F-A937-4B68-A32B-8D719B9080B1}{71D079E7-1D06-4929-A13C-2B9D586D5029}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub CommandButton1_Click()
Dim i As Integer, j As Integer
Dim datX As String, datY As String
Dim rngRead As Range, endcell As Range

If ((0 = UserForm1.OptionButton1.Value) _
 And (0 = UserForm1.OptionButton2.Value) _
 And (0 = UserForm1.OptionButton3.Value)) Then
 MsgBox "解析方法を選択してください。"
 Exit Sub
End If

With Application
  .ScreenUpdating = False
  .DisplayAlerts = False
End With
UserForm1.Hide
Worksheets(SHEETNAME1).Range("f5:h11").Delete (xlShiftToLeft)
Worksheets(SHEETNAME1).Range("g8").ColumnWidth = Worksheets(SHEETNAME1).StandardWidth
For i = Sheets.Count To 1 Step -1
  If Not (Sheets(i).Name = SHEETNAME1 Or Sheets(i).Name = SHEETNAME2) Then
  Sheets(i).Delete
  End If
Next i

'移動ボタンの消去
With Worksheets(SHEETNAME1)
 For i = .Shapes.Count To 1 Step -1
  If ((.Shapes(i).AlternativeText = MBTNNAME00) _
   Or (.Shapes(i).AlternativeText = MBTNNAME11) _
   Or (.Shapes(i).AlternativeText = MBTNNAME12) _
   Or (.Shapes(i).AlternativeText = MBTNNAME13) _
   Or (.Shapes(i).AlternativeText = MBTNNAME21) _
   Or (.Shapes(i).AlternativeText = MBTNNAME31) _
   Or (.Shapes(i).AlternativeText = MBTNNAME32)) Then
   .Shapes(i).Delete
  End If
 Next i
End With

If UserForm1.OptionButton1.Value Then
  With Worksheets(SHEETNAME1)
    .Range("g4:h4").MergeCells = True    '表づくり
    .Range("f5:h5").MergeCells = True
    .Range("f6:f7").MergeCells = True
    .Range("f6:f7").WrapText = True
    .Range("f4:h7").Interior.ColorIndex = 45
    .Range("f4:h7").Borders.LineStyle = xlContinuous
    .Cells(4, 6).Value = "ファイル名"
    .Cells(5, 6).Value = "サーボ応答精度"
    .Cells(6, 6).Value = "サーボ応答設定"
    .Cells(6, 7).Value = "制御周期"
    '入力制限が入る予定
    
    '移動ボタンの追加
    .Buttons.Add(Left:=445, top:=424, width:=75.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c11"
     .Characters.Text = MBTNNAME11
     .Placement = xlFreeFloating
    End With
    .Buttons.Add(Left:=525.25, top:=424, width:=75.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c12"
     .Characters.Text = MBTNNAME12
     .Placement = xlFreeFloating
    End With
    .Buttons.Add(Left:=605.25, top:=424, width:=80.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c13"
     .Characters.Text = MBTNNAME13
     .Placement = xlFreeFloating
    End With
  End With
  Worksheets(SHEETNAME2).Activate
  With Worksheets(SHEETNAME2)
'    .Range("a1:u4097").Delete (xlShiftUp)
    .Range("A1:U" & ENDROW).Delete (xlShiftUp)
    .Range("a1:b2").Interior.ColorIndex = 37  '２枚めの表づくり
    .Cells(1, 1).Value = "位置データ"
    .Cells(2, 1).Value = "X"
    .Cells(2, 2).Value = "Y"
    .Range("d1:e2").Interior.ColorIndex = 33
    .Cells(1, 4).Value = "速度データ"
    .Cells(1, 5).Value = "単位:pulse/RTC"
    .Cells(2, 4).Value = "X"
    .Cells(2, 5).Value = "Y"
    .Range("g1:h2").Interior.ColorIndex = 45
    .Range("j1:k2").Interior.ColorIndex = 45
    .Cells(1, 7).Value = "速度データ"
    .Cells(1, 10).Value = "速度データ"
    .Cells(1, 8).Value = "単位:pps"
    .Cells(1, 11).Value = "単位:pps"
    .Cells(2, 7).Value = "時間"
    .Cells(2, 10).Value = "時間"
    .Cells(2, 8).Value = "第１行データ"
    .Cells(2, 8).ColumnWidth = 11
    .Cells(2, 11).Value = "第2行データ"
    .Cells(2, 11).ColumnWidth = 11
  End With
  ActiveWorkbook.Worksheets.Add after:=Worksheets(SHEETNAME2)  'なぜか最後のシートには
  ActiveWorkbook.Charts.Add after:=Worksheets(SHEETNAME2)      'chartはだめらしい。
  With ActiveChart
    .HasTitle = True
    .Name = CHARTNAME13
    .ChartTitle.Characters.Font.Size = 11
    
    .ChartType = xlXYScatterLines
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("l3:l4097"), PlotBy:=xlColumn
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumn
    .SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumns
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "msec"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "pps"
  End With
  ActiveChart.PlotArea.Select
  Selection.Height = 383
  Selection.top = 38
'  With ActiveChart.PlotArea
'    .Height = 383
'    .top = 38
'  End With
  ActiveSheet.Buttons.Add(Left:=120, top:=5, width:=100, Height:=30).Select
  With Selection
    .OnAction = "左ボタン"
    .Characters.Text = BUTTONSTR1
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=400, top:=5, width:=100, Height:=30).Select
  With Selection
    .OnAction = "右ボタン"
    .Characters.Text = BUTTONSTR2
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=20, top:=5, width:=70, Height:=30).Select
  With Selection
    .OnAction = "graph_print"
    .Characters.Text = BUTTONSTR3
    .Placement = xlFreeFloating
  End With
    ActiveSheet.Buttons.Add(Left:=455.25, top:=424.5, width:=65.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_csv"
     .Characters.Text = MBTNNAME00
     .Placement = xlFreeFloating
    End With
    ActiveSheet.Buttons.Add(Left:=525.25, top:=424.5, width:=75.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c11"
     .Characters.Text = MBTNNAME11
     .Placement = xlFreeFloating
    End With
    ActiveSheet.Buttons.Add(Left:=605.25, top:=424.5, width:=75.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c12"
     .Characters.Text = MBTNNAME12
     .Placement = xlFreeFloating
    End With
  ActiveChart.ChartArea.Select
  With ActiveChart.Axes(xlCategory)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.Axes(xlValue)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.PlotArea
    .Border.LineStyle = xlAutomatic
    .Interior.ColorIndex = 2
  End With
  cmment_insert (CHARTNAME13)
  
  ActiveWorkbook.Charts.Add after:=Worksheets(SHEETNAME2)
  With ActiveChart
    .HasTitle = True
    .Name = CHARTNAME12
    .ChartTitle.Characters.Font.Size = 11
    
    .ChartType = xlXYScatterLines
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("l3:l4097"), PlotBy:=xlColumn
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumn
    .SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumns
    .HasLegend = False
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "msec"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "pps"
  End With
  ActiveChart.PlotArea.Select
  Selection.Height = 383
  Selection.top = 38
'  With ActiveChart.PlotArea
'    .Height = 383
'    .top = 38
'  End With
  ActiveSheet.Buttons.Add(Left:=120, top:=5, width:=100, Height:=30).Select
  With Selection
    .OnAction = "左ボタン"
    .Characters.Text = BUTTONSTR1
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=400, top:=5, width:=100, Height:=30).Select
  With Selection
    .OnAction = "右ボタン"
    .Characters.Text = BUTTONSTR2
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=20, top:=5, width:=70, Height:=30).Select
  With Selection
    .OnAction = "graph_print"
    .Characters.Text = BUTTONSTR3
    .Placement = xlFreeFloating
  End With
    ActiveSheet.Buttons.Add(Left:=455.25, top:=424.5, width:=65.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_csv"
     .Characters.Text = MBTNNAME00
     .Placement = xlFreeFloating
    End With
    ActiveSheet.Buttons.Add(Left:=525.25, top:=424.5, width:=75.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c11"
     .Characters.Text = MBTNNAME11
     .Placement = xlFreeFloating
    End With
    ActiveSheet.Buttons.Add(Left:=605.25, top:=424.5, width:=80.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c13"
     .Characters.Text = MBTNNAME13
     .Placement = xlFreeFloating
    End With
  ActiveChart.ChartArea.Select  '初期状態でボタンがアクティブになっているのが嫌なので
  With ActiveChart.Axes(xlCategory)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.Axes(xlValue)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.PlotArea
    .Border.LineStyle = xlAutomatic
    .Interior.ColorIndex = 2
  End With
  cmment_insert (CHARTNAME12)
  
  ActiveWorkbook.Charts.Add after:=Worksheets(SHEETNAME2)
  With ActiveChart
    .HasTitle = True
    .Name = CHARTNAME11
    .ChartTitle.Characters.Font.Size = 11
    
    .ChartType = xlXYScatterLines
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("l3:l4097"), PlotBy:=xlColumn
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumn
    .SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumns
    .HasLegend = False
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "msec"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "pps"
  End With
  ActiveChart.PlotArea.Select
  Selection.Height = 383
  Selection.top = 38
'  With ActiveChart.PlotArea
'    .Height = 383
'    .top = 38
'  End With
  ActiveSheet.Buttons.Add(Left:=120, top:=5, width:=100, Height:=30).Select
  With Selection
    .OnAction = "左ボタン"
    .Characters.Text = BUTTONSTR1
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=400, top:=5, width:=100, Height:=30).Select
  With Selection
    .OnAction = "右ボタン"
    .Characters.Text = BUTTONSTR2
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=20, top:=5, width:=70, Height:=30).Select
  With Selection
    .OnAction = "graph_print"
    .Characters.Text = BUTTONSTR3
    .Placement = xlFreeFloating
  End With
    ActiveSheet.Buttons.Add(Left:=455.25, top:=424.5, width:=65.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_csv"
     .Characters.Text = MBTNNAME00
     .Placement = xlFreeFloating
    End With
    ActiveSheet.Buttons.Add(Left:=525.25, top:=424.5, width:=75.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c12"
     .Characters.Text = MBTNNAME12
     .Placement = xlFreeFloating
    End With
    ActiveSheet.Buttons.Add(Left:=605.25, top:=424.5, width:=80.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c13"
     .Characters.Text = MBTNNAME13
     .Placement = xlFreeFloating
    End With
  ActiveChart.ChartArea.Select
  With ActiveChart.Axes(xlCategory)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.Axes(xlValue)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.PlotArea
    .Border.LineStyle = xlAutomatic
    .Interior.ColorIndex = 2
  End With
  cmment_insert (CHARTNAME11)
  i = Sheets.Count
  Sheets(i).Delete






ElseIf UserForm1.OptionButton2.Value Then
  With Worksheets(SHEETNAME1)
    .Range("g4:h4,f5:h5,f6:f9").MergeCells = True
    .Range("f6:f9").WrapText = True
    .Range("f4:h9").Interior.ColorIndex = 38
    .Range("f4:h9").Borders.LineStyle = xlContinuous
    .Cells(4, 6).Font.ColorIndex = 1
    .Cells(4, 6).Value = "ファイル名"
    .Cells(5, 6).Value = "真円精度"
    .Cells(6, 6).Value = "真円精度設定"
    .Cells(6, 7).Value = "半径"
    .Cells(7, 7).Value = "X中心点"
    .Cells(8, 7).Value = "Y中心点"
    .Cells(9, 7).Value = "標準偏差スケール"
    .Cells(9, 7).ColumnWidth = 16
  
    '移動ボタンの追加
    .Buttons.Add(Left:=605.25, top:=424, width:=80.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c21"
     .Characters.Text = MBTNNAME21
     .Placement = xlFreeFloating
    End With
  End With
 ' With ActiveSheet.Range("h6,h9").Validation
 '   .Delete
 '   .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
 '   .IgnoreBlank = False
 '   .ErrorTitle = "入力エラー"
 '   .ErrorMessage = "このセルには0より大きい数値を入力してください"
 '   .IMEMode = xlIMEModeDisable
 '   .ShowInput = False
 '   .ShowError = True
 ' End With
  With ActiveSheet.Range("h7:h8").Validation
    .Delete
 '   .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="-9999999999999999"
  '  .IgnoreBlank = False
  '  .ErrorTitle = "入力エラー"
  '  .ErrorMessage = "このセルには小数点数を入力してください"
  '  .IMEMode = xlIMEModeDisable
  '  .ShowInput = False
  '  .ShowError = True
  End With
  
  Worksheets(SHEETNAME2).Activate
  With Worksheets(SHEETNAME2)
'    .Range("a1:u4097").Delete (xlShiftUp)
    .Range("A1:U" & ENDROW).Delete (xlShiftUp)
    .Range("a1:b2").Interior.ColorIndex = 37
    .Range("d1:f2").Interior.ColorIndex = 33
    .Range("h1:i2").Interior.ColorIndex = 45
    .Range("k1:l2").Interior.ColorIndex = 46
    .Range("q1:r2").Interior.ColorIndex = 43
    .Cells(1, 1).Value = "シフトﾛｷﾞﾝｸﾞﾃﾞｰﾀ"
    .Cells(1, 4).Value = "理想円"
    .Cells(1, 8).Value = "内側"
    .Cells(1, 11).Value = "外側"
    .Cells(1, 17).Value = "拡大誤差"
    .Range("a2,e2,h2,k2,q2").Value = "X"
    .Range("b2,f2,i2,l2,r2").Value = "Y"
    .Range("d2").Value = "角度"
  End With
  ActiveWorkbook.Worksheets.Add after:=Worksheets(SHEETNAME2)
  ActiveWorkbook.Charts.Add after:=Worksheets(SHEETNAME2)
  With ActiveChart
    .Name = CHARTNAME21
    .ChartType = xlXYScatterSmoothNoMarkers
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("l3:l4097"), PlotBy:=xlColumn
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumn
    .SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumns
    .HasLegend = False
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "X"
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Font.FontStyle = "太字"
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Font.Size = 14
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Y"
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Font.FontStyle = "太字"
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Font.Size = 14

  End With
  With ActiveChart.PlotArea
'    .width = 500
'    .Left = 150
    .Border.LineStyle = xlAutomatic
    .Interior.ColorIndex = 2
  End With
  
  ActiveChart.PlotArea.Select
  Selection.width = 500
  Selection.Left = 150
  
  With ActiveChart.Axes(xlValue)
    .MinimumScale = -5833.33333333333
    .MaximumScale = 5833.33333333333
    .MinorUnit = 3000
    .MajorUnit = 3000
    .HasMajorGridlines = False
    .HasMinorGridlines = False
    .MajorTickMark = xlNone
    .MinorTickMark = xlNone
    .TickLabelPosition = xlNone
    .AxisTitle.Left = 379
    .AxisTitle.top = 15
    .AxisTitle.Orientation = xlHorizontal
    .AxisTitle.Font.Size = 14
  End With
  With ActiveChart.Axes(xlCategory)
    .MinimumScale = -5833.33333333333
    .MaximumScale = 5833.33333333333
    .MinorUnit = 3000
    .MajorUnit = 3000
    .HasMajorGridlines = False
    .HasMinorGridlines = False
    .MajorTickMark = xlNone
    .MinorTickMark = xlNone
    .TickLabelPosition = xlNone
    .AxisTitle.Left = 637
    .AxisTitle.top = 220
    .AxisTitle.Font.Size = 14
  End With
  
  ActiveSheet.Buttons.Add(Left:=5, top:=10, width:=100, Height:=30).Select
  With Selection
    .OnAction = "表示項目変更"
    .Characters.Text = BUTTONSTR1
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=5, top:=60, width:=100, Height:=30).Select
  With Selection
    .OnAction = "設定項目変更"
    .Characters.Text = BUTTONSTR2
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=5, top:=110, width:=100, Height:=30).Select
  With Selection
    .OnAction = "graph_print"
    .Characters.Text = BUTTONSTR3
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=615.25, top:=424.5, width:=80.25, Height:=20.25).Select
  With Selection
   .OnAction = "move_to_csv"
   .Characters.Text = MBTNNAME00
   .Placement = xlFreeFloating
  End With
  cmment_insert (CHARTNAME21)
  ActiveChart.ChartArea.Select
  i = Sheets.Count
  Sheets(i).Delete
  
  
  
    
ElseIf UserForm1.OptionButton3.Value Then
  With ActiveSheet                  'ユーザーが罫線の色などを設定していることを
    .Range("g4:h4,f5:h5,f9:h9,f10:g10,f11:g11").MergeCells = True '想定していない
    .Range("f4:h11").Interior.ColorIndex = 6
    .Range("f4:h11").Borders.LineStyle = xlContinuous
    .Cells(4, 6).Value = "ファイル名"
    .Cells(5, 6).Value = "直線精度"
    .Cells(6, 6).Value = "基準直線"
    .Cells(6, 7).Value = "X座標"
    .Cells(6, 8).Value = "Y座標"
    .Cells(7, 6).Value = "始点"
    .Cells(8, 6).Value = "終点"
    .Cells(9, 6).Value = "解析データ範囲"
    .Cells(10, 6).Value = "開始(1～)"
'    .Cells(11, 6).Value = "終了(～4097)"
    .Cells(11, 6).Value = "終了(～" & ENDROW & ")"
    If Not IsEmpty(.Cells(3, 1).Value) Then
      .Range(RXB).Value = .Cells(3, 1).Value
      .Range(RYB).Value = .Cells(3, 2).Value
'      Set endcell = .Range("a3:a4100").find(what:=Empty)
      Set endcell = .Range("a3:a" & ENDADD3).find(what:=Empty)
      .Range(RXE).Value = endcell.Offset(-1, 0).Value
      .Range(RYE).Value = endcell.Offset(-1, 1).Value
      .Range(STARTPOINT).Value = 1
      .Range(ENDPOINT).Value = (endcell.Row - 3)
    End If
  
    .Buttons.Add(Left:=525.25, top:=424, width:=75.25, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c31"
     .Characters.Text = MBTNNAME31
     .Placement = xlFreeFloating
    End With
    .Buttons.Add(Left:=605.25, top:=424, width:=110, Height:=20.25).Select
    With Selection
     .OnAction = "move_to_c32"
     .Characters.Text = MBTNNAME32
     .Placement = xlFreeFloating
    End With
  
  End With
  'With ActiveSheet.Range("G7:H8").Validation    '入力制限
   ' .Delete
   ' .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="-99999999999999"
   ' .IgnoreBlank = False
   ' .ErrorTitle = "入力エラー"
   ' .ErrorMessage = "このセルには小数点数を入力してください"
   ' .IMEMode = xlIMEModeDisable
   ' .ShowInput = False
   ' .ShowError = True
  'End With
  'With ActiveSheet.Range(MABIKIPOINT, STARTPOINT).Validation
  '  .Delete
  '  .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
  '  .ErrorTitle = "入力エラー"
  '  .ErrorMessage = "このセルには正の整数を入力してください"
  '  .IMEMode = xlIMEModeDisable
  '  .ShowInput = False
  '  .ShowError = True
  'End With
  'With ActiveSheet.Range(ENDPOINT).Validation
  '  .Delete
  '  .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=h10", Formula2:="4097"
  '  .ErrorTitle = "入力エラー"
  '  .ErrorMessage = "このセルには正の整数を入力してください。また、開始より小さい数を入力することはできません。"
  '  .IMEMode = xlIMEModeDisable
  '  .ShowInput = False
  '  .ShowError = True
  'End With
  With Worksheets(SHEETNAME2)
'    .Range("a1:u4097").Delete (xlShiftUp)
    .Range("A1:U" & ENDROW).Delete (xlShiftUp)
    .Range("a1:b2").Interior.ColorIndex = 50
    .Range("c1:e2").Interior.ColorIndex = 45
    .Range("g1:g2").Interior.ColorIndex = 39
    .Range("h1:i2").Interior.ColorIndex = 8
    .Range("j1:l2").Interior.ColorIndex = 3
    .Range("a1:b1,c1:e1,g1:g2,h1:i1,j1:l1").MergeCells = True
    .Cells(1, 1).Value = "CVSデータ"
    .Cells(1, 3).Value = "誤差（間引く前）"
    .Cells(1, 7).Value = "サンプリングデータ"
    .Cells(1, 7).ColumnWidth = 18
    .Cells(1, 8).Value = "位置データ"
    .Cells(1, 10).Value = "誤差(間引き後)"
    .Range("a2,c2,h2").Value = "X"
    .Range("b2,d2,i2").Value = "Y"
    .Range("e2,l2").Value = "L"
    .Cells(2, 10).Value = "EX"
    .Cells(2, 11).Value = "EY"
  End With
  
'  ActiveWorkbook.Worksheets.Add after:=Worksheets(SHEETNAME2)
'  ActiveSheet.Name = CHARTNAME31
  
  ActiveWorkbook.Worksheets.Add after:=Worksheets(SHEETNAME2)
  ActiveSheet.Name = CHARTNAME32
  
  ActiveSheet.Shapes.AddChart.Select
  
  'ActiveWorkbook.Charts.Add after:=Worksheets(SHEETNAME2)
  'ActiveChart.Name = CHARTNAME32
  
  ActiveChart.ChartType = xlXYScatterLines
'  ActiveChart.SetSourceData Source:=Sheets(SHEETNAME2).Range("l3:l4097"), PlotBy:=xlColumn
'  ActiveChart.SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumn
'  ActiveChart.SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumns
'  ActiveChart.SetSourceData Source:=Sheets(SHEETNAME2).Range("L" & STARTROW & ":M" & ENDROW), PlotBy:=xlColumns
  ActiveChart.SetSourceData Source:=Sheets(SHEETNAME2).Range("l3:M4097"), PlotBy:=xlColumns
'  ActiveChart.SeriesCollection(1).XValues = Range("L3:L4097") '元々 CommentOut されていた
'  ActiveChart.SeriesCollection(1).YValues = Range("M3:M4097") '元々 CommentOut されていた
'  ActiveChart.SeriesCollection(1).XValues = Range("g3:g4097") '元々 CommentOut されていた
'  Worksheets(CHARTNAME32).Select
'  ActiveChart.Location Where:=xlLocationAsObject, Name:=CHARTNAME32
  ActiveChart.HasLegend = False
  ActiveChart.HasTitle = True
  With ActiveChart.Axes(xlCategory)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.Axes(xlValue)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.PlotArea
    .Border.LineStyle = xlAutomatic
    .Interior.ColorIndex = 2
  End With
  
  With ActiveSheet.Shapes("グラフ 1")
'widthとheightの値がうまく設定できない。
'left,topとheight,widthの順序を変えると
'大きさが変わる
    .Height = 160
    .width = 750
    .Left = 1                               'copy
    .top = 1
'    .ScaleWidth 4.06, msoFalse, msoScaleFromTopLeft
'    .ScaleHeight 1.3, msoFalse, msoScaleFromTopLeft
   .Duplicate.Select
  End With
  With ActiveSheet.Shapes("グラフ 2")
    .Height = 130
    .Left = 1
    .top = 160
    .Duplicate.Select
  End With
  With ActiveSheet.Shapes("グラフ 3")
    .Height = 130
    .Left = 1
    .top = 290
  End With
  ActiveSheet.Buttons.Add(Left:=120, top:=15, width:=100, Height:=30).Select
  With Selection
    .OnAction = "ボタン"
    .Characters.Text = BUTTONSTR1
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=600, top:=15, width:=100, Height:=30).Select
  With Selection
    .OnAction = "ボタン2"
    .Characters.Text = BUTTONSTR2
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=20, top:=15, width:=70, Height:=30).Select
  With Selection
    .OnAction = "graph_print"
    .Characters.Text = BUTTONSTR3
    .Placement = xlFreeFloating
  End With
  'ActiveSheet.ChartObjects("グラフ 3").Chart.Buttons.Add(Left:=525.25, Top:=424.5, width:=75.25, Height:=20.25).Select
  ActiveSheet.Buttons.Add(Left:=555.25, top:=424.5, width:=75.25, Height:=20.25).Select
  ActiveSheet.Shapes("Button 4").OnAction = "move_to_csv"
  ActiveSheet.Shapes("Button 4").DrawingObject.Caption = MBTNNAME00
  ActiveSheet.Shapes("Button 4").Placement = xlFreeFloating
  ActiveSheet.Buttons.Add(Left:=645.25, top:=424.5, width:=80.25, Height:=20.25).Select
  ActiveSheet.Shapes("Button 5").OnAction = "move_to_c31"
  ActiveSheet.Shapes("Button 5").DrawingObject.Caption = MBTNNAME31
  ActiveSheet.Shapes("Button 5").Placement = xlFreeFloating
'  ActiveChart.PlotArea.Select '右ボタンがおかしくなる現象防止
  ActiveSheet.ChartObjects("グラフ 1").Select
  cmment_insert (CHARTNAME32)
  
  
'  ActiveWorkbook.Worksheets.Add after:=Worksheets(SHEETNAME2)
'  ActiveSheet.Name = CHARTNAME31
  
'  Worksheets(CHARTNAME31).Select
'  ActiveSheet.Shapes.AddChart.Select
  
  ActiveWorkbook.Charts.Add after:=Worksheets(SHEETNAME2)
  With ActiveChart
    .Name = CHARTNAME31
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("i3:i4097"), PlotBy:=xlColumn
'    .SetSourceData Source:=Sheets(SHEETNAME2).Range("I" & STARTROW & ":I" & ENDROW), PlotBy:=xlColumn
    .SetSourceData Source:=Sheets(SHEETNAME2).Range("I" & STARTROW & ":I" & ENDROW), PlotBy:=xlColumns
    .HasLegend = False
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "pls"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "pls"
    
    '.ChartTitle.Characters.Text = "軌跡"
  End With
  ActiveChart.PlotArea.Select
  Selection.Height = 383
  Selection.top = 38
'  With ActiveChart.PlotArea
'    .Height = 383
'    .top = 38
'  End With
  ActiveSheet.Buttons.Add(Left:=400, top:=10, width:=100, Height:=30).Select
  With Selection
    .OnAction = "ボタン2"
    .Characters.Text = BUTTONSTR2
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=10, top:=10, width:=70, Height:=30).Select
  With Selection
    .OnAction = "graph_print"
    .Characters.Text = BUTTONSTR3
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=120, top:=10, width:=100, Height:=30).Select
  'なぜか表示範囲設定ボタンが勝手にアクティブになる。
  With Selection
    .OnAction = "ボタン3"
    .Characters.Text = BUTTONSTR1
    .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=525.25, top:=424.5, width:=75.25, Height:=20.25).Select
  With Selection
   .OnAction = "move_to_csv"
   .Characters.Text = MBTNNAME00
   .Placement = xlFreeFloating
  End With
  ActiveSheet.Buttons.Add(Left:=605.25, top:=424.5, width:=110, Height:=20.25).Select
  With Selection
   .OnAction = "move_to_c32"
   .Characters.Text = MBTNNAME32
   .Placement = xlFreeFloating
  End With

'  ActiveSheet.ChartObjects("グラフ 1").Select
  ActiveChart.ChartArea.Select
  With ActiveChart.Axes(xlCategory)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.Axes(xlValue)
    .HasMajorGridlines = False
    .HasMinorGridlines = False
  End With
  With ActiveChart.PlotArea
    .Border.LineStyle = xlAutomatic
    .Interior.ColorIndex = 2
  End With
  cmment_insert (CHARTNAME31)
'  i = Sheets.Count
'  Sheets(i).Delete
End If

With Worksheets(SHEETNAME1)
 .Activate
 .CommandButton2.Enabled = True
 .CommandButton3.Enabled = True
 .Range("A3").Select
End With

With Application
  .DisplayAlerts = True
  .ScreenUpdating = True
End With
End Sub

