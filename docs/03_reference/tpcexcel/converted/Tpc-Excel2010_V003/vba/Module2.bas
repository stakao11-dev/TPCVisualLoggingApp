Attribute VB_Name = "Module2"
Option Explicit
'このモジュールにはサーボ応答のspeedckeckと右ボタン、左ボタンがあります。
'speedckeckはメインの計算。左ボタンはデータ範囲設定を表示、右は設定項目変更です。
Sub speedckeck(csvfname As String)
Dim rtctim As Integer
 Application.ScreenUpdating = False  '画面再描画禁止
  With Worksheets(SHEETNAME1)
   .Activate    'シートの選択
   rtctim = .Range(RTCPOINT).Value
  
   '中間データ消去
   Worksheets(SHEETNAME2).Activate             'シートの選択
   Worksheets(SHEETNAME2).Range("A3:K65536").Select
   Selection.ClearContents
  '位置データ用意"G" & STARTROW & ":G" & ENDROW
   .Activate             'シートの選択
'   .Range("A3:B4097").Select                  'シートの選択
   .Range("A" & STARTROW & ":B" & ENDROW).Select                  'シートの選択
   Selection.Copy
   Worksheets(SHEETNAME2).Activate            'シートの選択
   Worksheets(SHEETNAME2).Range("A3").Select
   ActiveSheet.Paste
   .Activate                                  'シートの選択
   .Range("A3").Select
   Application.CutCopyMode = False
   'Worksheets(SHEETNAME2).Activate            'シートの選択
  End With
  Call speedcalc(csvfname, rtctim)
 Application.ScreenUpdating = True   '画面描画有効
End Sub
Sub speedcalc(csvfname As String, rtctim As Integer)
         
    '速度データ用意(pulse/rtc,pps)
    Dim posX, posY, preposX, preposY, timX, timY As Range
    Dim feedX, feedY, ppsfeedX, ppsfeedY As Range
'    Dim cnt_X As Integer, cnt_Y As Integer, errstr As String
    Dim cnt_X As Double, cnt_Y As Double, errstr As String
    cnt_X = 0
    cnt_Y = 0
    
    With Worksheets(SHEETNAME2)
        Set posX = .Range("A4")     '理想円用角度データ作成
        Set posY = .Range("B4")     '理想円用角度データ作成
        Set preposX = .Range("A3")     '理想円用角度データ作成
        Set preposY = .Range("B3")     '理想円用角度データ作成
        Set feedX = .Range("D4")     '理想円用角度データ作成
        Set feedY = .Range("E4")     '理想円用角度データ作成
        Set timX = .Range("G4")   '理想円用角度データ作成
        Set ppsfeedX = .Range("H4")   '理想円用角度データ作成
        Set timY = .Range("J4")   '理想円用角度データ作成
        Set ppsfeedY = .Range("K4")   '理想円用角度データ作成
        Do While Not IsEmpty(posX)
            timX.Value = cnt_X
            timY.Value = cnt_Y
            feedX.Value = posX.Value - preposX.Value
            feedY.Value = posY.Value - preposY.Value
            ppsfeedX.Value = feedX.Value * 1000 / rtctim
            ppsfeedY.Value = feedY.Value * 1000 / rtctim
            Set posX = posX.Offset(1, 0)            'X軸用データ挿入セルとして一つ下のセルを選択
            Set posY = posY.Offset(1, 0)            'Y軸用データ挿入セルとして一つ下のセルを選択
            Set preposX = preposX.Offset(1, 0)      'X軸用データ挿入セルとして一つ下のセルを選択
            Set preposY = preposY.Offset(1, 0)      'Y軸用データ挿入セルとして一つ下のセルを選択
            Set feedX = feedX.Offset(1, 0)          'X軸用データ挿入セルとして一つ下のセルを選択
            Set feedY = feedY.Offset(1, 0)          'X軸用データ挿入セルとして一つ下のセルを選択
            Set timX = timX.Offset(1, 0)            'X軸用データ挿入セルとして一つ下のセルを選択
            Set ppsfeedX = ppsfeedX.Offset(1, 0)    'X軸用データ挿入セルとして一つ下のセルを選択
            Set timY = timY.Offset(1, 0)            'X軸用データ挿入セルとして一つ下のセルを選択
            Set ppsfeedY = ppsfeedY.Offset(1, 0)    'X軸用データ挿入セルとして一つ下のセルを選択
            cnt_X = cnt_X + rtctim
            cnt_Y = cnt_Y + rtctim
        Loop
    End With
    errstr = ERRTXT1 & (1000 / rtctim) & ERRTXT2
    Sheets(CHARTNAME11).Select  '従来のやり方ではエラーがでるsetsourcedataが必要"G" & STARTROW & ":G" & ENDROW
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "サーボ応答精度" & Chr(10) & Now & Chr(10) & csvfname
'    ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("h4:h4097"), PlotBy:=xlColumns 'Yデータを選択
'    ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G4:G4097")   'Xデータを選択
    ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("H4:H" & ENDROW), PlotBy:=xlColumns 'Yデータを選択
    ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G4:G" & ENDROW)   'Xデータを選択
    ActiveChart.SeriesCollection(1).MarkerSize = 2
    errmes_ins (errstr)


    Sheets(CHARTNAME12).Select
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "サーボ応答精度" & Chr(10) & Now & Chr(10) & csvfname
 '   ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("k4:k4097"), PlotBy:=xlColumns 'Yデータを選択
 '   ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("J4:J4097")  'Xデータを選択
    ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("K4:K" & ENDROW), PlotBy:=xlColumns 'Yデータを選択
    ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("J4:J" & ENDROW)  'Xデータを選択
    ActiveChart.SeriesCollection(1).MarkerSize = 2
    errmes_ins (errstr)
    
    
    Sheets(CHARTNAME13).Select
    ActiveChart.HasTitle = True
    ActiveChart.ChartTitle.Text = "サーボ応答精度" & Chr(10) & Now & Chr(10) & csvfname
'    ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("h4:h4097"), PlotBy:=xlColumns  'Yデータを選択
'    ActiveChart.SeriesCollection(1).XValues = Worksheets("測定結果").Range("G4:G4097")  'Xデータを選択
    ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("H4:H" & ENDROW), PlotBy:=xlColumns  'Yデータを選択
    ActiveChart.SeriesCollection(1).XValues = Worksheets("測定結果").Range("G4:G" & ENDROW)  'Xデータを選択
    ActiveChart.SeriesCollection(1).MarkerSize = 2
    ActiveChart.SeriesCollection(1).Name = "=""第1列ﾃﾞｰﾀ"""
    ActiveChart.SeriesCollection.NewSeries
'    ActiveChart.SeriesCollection(2).XValues = Worksheets("測定結果").Range("G4:G4097")  'Xデータを選択
'    ActiveChart.SeriesCollection(2).Values = Worksheets("測定結果").Range("K4:K4097")  'Yデータを選択
    ActiveChart.SeriesCollection(2).XValues = Worksheets("測定結果").Range("G4:G" & ENDROW)  'Xデータを選択
    ActiveChart.SeriesCollection(2).Values = Worksheets("測定結果").Range("K4:K" & ENDROW)  'Yデータを選択
    ActiveChart.SeriesCollection(2).Name = "=""第2列ﾃﾞｰﾀ"""
    ActiveChart.SeriesCollection(2).MarkerSize = 2
    ActiveChart.HasLegend = True
    ActiveChart.HasLegend = True
    ActiveChart.Legend.Position = xlCorner
    ActiveChart.Legend.Left = 639
    ActiveChart.Legend.top = 6
    ActiveChart.PlotArea.width = 696
    errmes_ins (errstr)
End Sub
Sub 左ボタン()
Call form3_ini
End Sub
Sub 右ボタン()
With Worksheets(SHEETNAME1)            'シートの選択
    UserForm6.TextBox1.Text = .Range(RTCPOINT).Value  'X軸中心をコントロールへ数値の書込み
End With
    UserForm6.Show
End Sub

Sub errmes_ins(errstr As String)
Dim txboxname_str As String, retaintxboxname As String, i As Integer
Dim retaintxboxvisible As String, retaintxboxstr As String

    If (ActiveChart.Name = CHARTNAME11) Then
     retaintxboxname = TXTBOXNAME111
    ElseIf (ActiveChart.Name = CHARTNAME12) Then
     retaintxboxname = TXTBOXNAME121
    ElseIf (ActiveChart.Name = CHARTNAME13) Then
     retaintxboxname = TXTBOXNAME131
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
     ActiveChart.Shapes.Item(txboxname_str).Select
    Else
     ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 570.25, 380.25, 7.2, 7.2).Select
     Selection.ShapeRange.ScaleWidth 6.91, msoFalse, msoScaleFromTopLeft
     Worksheets(SHEETNAME1).Range(retaintxboxname).Value = ""
    End If
    Worksheets(SHEETNAME1).Range(retaintxboxname).Value = Selection.Name
    Selection.Characters.Text = errstr
    Selection.AutoScaleFont = False
    Selection.AutoSize = True
End Sub
