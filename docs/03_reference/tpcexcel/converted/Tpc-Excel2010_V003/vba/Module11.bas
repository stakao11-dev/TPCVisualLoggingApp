Attribute VB_Name = "Module11"
Option Explicit
'ここにはボタン2_clickと直線精度、再計算、ボタン、ボタン２、ボタン３、～changeがあります。
'ボタン２_clickはｃｓｖファイルを読込、直線精度はメインの計算です。
'再計算は設定を変更して再計算を使用したときに動作します。
'ボタン２は再計算のデータ範囲設定を表示するためのものです。
'ボタンは軸別直線精度の表示範囲設定を表示するためのものです。
'ボタン３は軌跡の表示範囲設定を表示するためのものです。
Public start As Integer, mabiki As Integer, endnum As Double
Public rx0 As Double, rx1 As Double, ry0 As Double, ry1 As Double
Sub make_locus_graph()
Dim cnt As Integer, chrtname As String
Application.ScreenUpdating = False
 cnt = ActiveSheet.ChartObjects.Count
 If (0 < cnt) Then
  chrtname = ActiveSheet.ChartObjects(1).Name
  ActiveSheet.ChartObjects(chrtname).Activate
  With ActiveChart
'   .SetSourceData Source:=Sheets(SHEETNAME1).Range(STARTCSVX & ":" & ENDCSVY), PlotBy:=xlColumns
   .SetSourceData Source:=Sheets(SHEETNAME1).Range(STARTCSVX & ":B" & ENDROW), PlotBy:=xlColumns
   .HasTitle = True
   .ChartTitle.Characters.Text = "簡易軌跡表示"
   .PlotArea.Interior.ColorIndex = xlNone
   .Axes(xlValue).TickLabels.Font.ColorIndex = 3
   .Axes(xlValue).TickLabels.Orientation = 10
   .Axes(xlCategory, xlPrimary).HasTitle = True
   .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "X"
   .Axes(xlValue, xlPrimary).HasTitle = True
   .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Y"
  End With
 Else
  Charts.Add
  With ActiveChart
   .ChartType = xlXYScatterLines
'   .SetSourceData Source:=Sheets(SHEETNAME1).Range(STARTCSVX & ":" & ENDCSVY), PlotBy:=xlColumns
   .SetSourceData Source:=Sheets(SHEETNAME1).Range(STARTCSVX & ":B" & ENDROW), PlotBy:=xlColumns
   .Location Where:=xlLocationAsObject, Name:="ﾛｷﾞﾝｸﾞﾃﾞｰﾀ"
   .HasTitle = True
   .ChartTitle.Characters.Text = "簡易軌跡表示"
   .PlotArea.Interior.ColorIndex = xlNone
   .Axes(xlValue).TickLabels.Font.ColorIndex = 3
   .Axes(xlValue).TickLabels.Orientation = 10
   .Axes(xlCategory, xlPrimary).HasTitle = True
   .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "X"
   .Axes(xlValue, xlPrimary).HasTitle = True
   .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Y"
  End With
 End If
Application.ScreenUpdating = True
Sheets(SHEETNAME1).Range(STARTCSVX).Select
End Sub
Sub read_csv()
Dim FileName As Variant, F_Num As Integer, rngRead As Range, rngRead2 As Range
Dim strBuf As String
Dim obj As New DataObject
'Dim i As Integer
Dim i As Double
FileName = Application.GetOpenFilename("Log File,*.csv", 1, "ﾛｷﾞﾝｸﾞﾃﾞｰﾀﾌｧｲﾙを開く", "開く", False)
If TypeName(FileName) = "Boolean" Then Exit Sub

On Error Resume Next
Set rngRead = ActiveSheet.Range("a3")
On Error GoTo 0
If rngRead Is Nothing Then Exit Sub

Application.ScreenUpdating = False

Range("a3:B65536").Select
Selection.ClearContents
Range("A3").Select

'CSV読み込み処理開始
 'PasteSpecialを使った方法
 F_Num = FreeFile
 Open FileName For Binary As F_Num
 ActiveSheet.Range(FILENAMEPOINT).Value = Right(FileName, Len(FileName) - InStrRev(FileName, "\"))  '追加
 strBuf = Space(LOF(F_Num))
 Get #F_Num, , strBuf
 Close F_Num
 obj.SetText strBuf
 obj.PutInClipboard
 Set obj = Nothing: strBuf = ""

 With rngRead
  .Worksheet.Activate
  With .Cells(1, 1)
   .Activate
   .Value = "dummy,dummy"
   Application.DisplayAlerts = False
    .TextToColumns , xlDelimited, xlTextQualifierDoubleQuote, False, False, False, True, False, False
   Application.DisplayAlerts = True
    .Resize(, 2).Clear
  End With
 End With
 Range(STARTCSVX).PasteSpecial

 'CSVﾌｧｲﾙを新しいｼｰﾄとして読み込んだ場合
  'Workbooks.Open FileName:= FileName
  'Windows(Right(FileName, Len(FileName) - InStrRev(FileName, "\"))).Activate
  'ActiveSheet.Range("A1:B4097").Select
  'Selection.Copy
  'Windows(ThisWorkbook.FullName & ".xls").Activate
  'Range("A3").Select
  'ActiveSheet.Paste
  'Windows(Right(FileName, Len(FileName) - InStrRev(FileName, "\"))).Activate
  'ActiveWindow.Close
  'Windows(ThisWorkbook.FullName & ".xls").Activate

'CSV読み込み処理終了

Worksheets(SHEETNAME1).Activate
With Worksheets(SHEETNAME1) '読み込んだときに各種設定を自動で記入
  If .Cells(5, 6).Value = "直線精度" Then
    .Range(RXB).Value = .Cells(3, 1).Value
    .Range(RYB).Value = .Cells(3, 2).Value
    .Range(STARTPOINT).Value = 1
    i = .Range("a3:a4100").find(what:=Empty).Row
    .Range(RXE) = .Cells(i - 1, 1).Value
    .Range(RYE).Value = .Cells(i - 1, 2).Value
    .Range(ENDPOINT).Value = i - 3
  End If
End With
 make_locus_graph

'Sheets(SHEETNAME1).CommandButton3.Enabled = True
Application.ScreenUpdating = True
End Sub
Function linearity_chk(csvfname As String)
Application.ScreenUpdating = False
Dim x0 As Double, y0 As Double, ex0 As Double, ey0, l0 As Double
Dim X As Range, Y As Range, EX As Range, EY As Range, SN As Range, l As Range
'Dim i As Integer
Dim i As Double

Worksheets(SHEETNAME2).Activate
Range("a3:l65536").Select
Selection.ClearContents '二枚目のシートを空に
Cells(3, 1).Select

Worksheets(SHEETNAME1).Activate
With Worksheets(SHEETNAME1)  'x座標始点　　　　　　x座標終点　　　　　　　　　　y座標始点　　　　　　　　　　　　y座標終点　　　　　　　　　　　　　　　     開始　　　　　　　　　　　        終了
  If Not (IsNumeric(.Range(RXB).Value) And IsNumeric(.Range(RXE).Value) And IsNumeric(.Range(RYB).Value) And IsNumeric(.Range(RYE).Value) And IsNumeric(.Range(STARTPOINT).Value) And IsNumeric(.Range(ENDPOINT).Value)) Then
    MsgBox "数値を入力してください"    ' 入力部分に数値が入っていないときにエラーを出す。入力制限のためいらないかも。
    Exit Function
  End If
  
  If IsEmpty(.Range(RXB).Value) Or IsEmpty(.Range(RXE).Value) Or IsEmpty(.Range(RYB).Value) Or IsEmpty(.Range(RYE).Value) Then
    MsgBox "基準直線を設定してください"    '基準直線が１つでもかけているとエラー
    Exit Function
  End If
  
  rx0 = .Range(RXB).Value   '数値の代入
  rx1 = .Range(RXE).Value
  ry0 = .Range(RYB).Value
  ry1 = .Range(RYE).Value
  mabiki = 1
  start = .Range(STARTPOINT).Value
  endnum = .Range(ENDPOINT).Value
  
  Select Case mabiki
  Case 0
    mabiki = 1    '間引きが0もしくはemptyのとき1を代入
  Case Is < 0
    MsgBox "「間引き」には正の整数を入力してください"  '負数はエラー、いらないかも。
    Exit Function
  End Select
  
  Select Case start
  Case 0
    start = 1 '0はデータの先頭からとする。いらない可能性大
    .Range(STARTPOINT).Value = start
  Case Is < 0
    MsgBox "「開始」には正の整数を入力してください"  '負数はエラー、いらないかも。
    Exit Function
  End Select
  
  Select Case endnum
  Case 0
'    endnum = 4067  '0またはemptyのときは最後のデータまで使用する。
'***Has***
  endnum = ENDROW
  Case Is < start
    MsgBox "「終了」には「開始」より小さい数値を入力することはできません"  '負数はエラーいらないかも。
    Exit Function
  End Select
  
  If rx0 = rx1 And ry0 = ry1 Then
    MsgBox "基準直線の始点終点が同じです。異なる値を設定して下さい。"    '基準直線の始点と終点の値が全く同じだとエラー
    Exit Function
  End If
  
'  .Range("a3:b4097").Select
  .Range("A" & STARTROW & ":B" & ENDROW).Select
   Selection.Copy
  .Cells(3, 1).Select
End With
  
Worksheets(SHEETNAME2).Activate
With Worksheets(SHEETNAME2)
'  Range("a3:b4097").PasteSpecial
  Range("A" & STARTROW & ":B" & ENDROW).PasteSpecial
  Application.CutCopyMode = False         'ﾛｷﾞﾝｸﾞﾃﾞｰﾀを２枚目にコピー
  i = 3
  Set X = .Range(XPOINT)
  Set Y = .Range(YPOINT)
  Set EX = .Range(ERRORX1)
  Set EY = .Range(ERRORY1)
  Set SN = .Range(SUMPLINGN)
  Set l = .Range(ERRORL1)
    While (Not IsEmpty(.Cells(i, 1))) And i < (endnum + 3) 'endnum+3は最終行+1
      x0 = .Cells(i, 1).Value
      y0 = .Cells(i, 2).Value
      ex0 = ((ry1 - ry0) ^ 2 * (x0 - rx0) - (rx1 - rx0) * (ry1 - ry0) * (y0 - ry0)) / ((rx1 - rx0) ^ 2 + (ry1 - ry0) ^ 2)
      ey0 = ((rx1 - rx0) ^ 2 * (y0 - ry0) - (rx1 - rx0) * (ry1 - ry0) * (x0 - rx0)) / ((rx1 - rx0) ^ 2 + (ry1 - ry0) ^ 2)
      l0 = (ex0 ^ 2 + ey0 ^ 2) ^ 0.5
      .Cells(i, 3).Value = ex0
      .Cells(i, 4).Value = ey0
      .Cells(i, 5).Value = l0
      If (i - 2 - mabiki * Int((i - 2) / mabiki)) = 0 And i >= start + 2 Then 'iが間引きで割り切れるときのみ
        X.Value = .Cells(i, 1).Value     'startはこのifないで反映
        Y.Value = .Cells(i, 2).Value
        EX.Value = .Cells(i, 3).Value
        EY.Value = .Cells(i, 4).Value
        l.Value = .Cells(i, 5).Value
        SN.Value = i - 2
        Set X = X.Offset(1, 0)
        Set Y = Y.Offset(1, 0)
        Set EX = EX.Offset(1, 0)
        Set EY = EY.Offset(1, 0)
        Set l = l.Offset(1, 0)
        Set SN = SN.Offset(1, 0)
      End If
    i = i + 1
    Wend
.Cells(3, 1).Select
End With
endnum = i - 3
Worksheets(SHEETNAME1).Range(ENDPOINT).Value = endnum 'データ数をシートに還元

'この辺はグラフを作るところでフォローしたい。userform4
If Not IsEmpty(Worksheets(SHEETNAME2).Range(YPOINT).Value) Then  'hがないとエラーに
  Charts(CHARTNAME31).Select
  
With ActiveChart
  .HasTitle = True
  .ChartTitle.Characters.Text = "軌跡" & Chr(10) & Now & Chr(10) & Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value
  .ChartTitle.Characters.Font.Size = 11
  .ChartTitle.top = 15
  
  .SeriesCollection(1).ChartType = xlXYScatter
'  .SetSourceData Source:=Worksheets(SHEETNAME2).Range("i3:i4097"), PlotBy:=xlColumns
  .SetSourceData Source:=Worksheets(SHEETNAME2).Range("I" & STARTROW & ":I" & ENDROW), PlotBy:=xlColumns
'  .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("h3:h4097")
  .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("H" & STARTROW & ":H" & ENDROW)
  .SeriesCollection(1).Name = "=""ｻﾝﾌﾟﾘﾝｸﾞﾃﾞｰﾀ"""
  .SeriesCollection.NewSeries
  .SeriesCollection(2).XValues = Worksheets(SHEETNAME1).Range("g7:g8")  '基準直線
  .SeriesCollection(2).Values = Worksheets(SHEETNAME1).Range("h7:h8")
  .SeriesCollection(2).Name = "=""理想直線"""
  
  .SeriesCollection(1).MarkerSize = 2
  .SeriesCollection(2).ChartType = xlXYScatterLinesNoMarkers
  .HasLegend = True
  .Legend.Position = xlCorner
  .Legend.Left = 639
  .Legend.top = 6
  .PlotArea.top = 65
  .PlotArea.width = 692
End With

End If

Sheets(CHARTNAME32).Activate
If Not IsEmpty(Worksheets(SHEETNAME2).Range(ERRORL1).Value) Then 'lに値がないとエラーになるので。
  ActiveSheet.ChartObjects("グラフ 1").Activate
  With ActiveChart
    .HasTitle = True
    .ChartTitle.Text = "軸別直線精度" & Chr(10) & Now & Chr(10) & Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value & Chr(10) & "L：ＸＹ合成誤差"
    .ChartTitle.Characters.Font.Size = 11
    .ChartTitle.top = 10
    
    .ChartArea.Border.Weight = 1
    .ChartArea.Border.LineStyle = 0
'    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("l3:l4097"), PlotBy:=xlColumns
    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumns
'    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
    .SeriesCollection(1).MarkerSize = 2
    .Axes(xlCategory).TickLabels.AutoScaleFont = False
    .Axes(xlValue).TickLabels.AutoScaleFont = False
    .Axes(xlCategory).TickLabels.Font.Size = 10.75
    .Axes(xlValue).TickLabels.Font.Size = 10.5
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "pls"
    .PlotArea.width = 750
    i = (.PlotArea.width - .PlotArea.InsideWidth)
    .PlotArea.width = 630 + i
    .PlotArea.Left = 25
    i = (.PlotArea.InsideLeft - .PlotArea.Left)
    .PlotArea.Left = 72 - i
    .PlotArea.top = 65
    .PlotArea.Height = 78
  End With
End If

If Not IsEmpty(Worksheets(SHEETNAME2).Range(ERRORX1).Value) Then  'jに値がないとエラーに
  ActiveSheet.ChartObjects("グラフ 2").Activate
  With ActiveChart
    .HasTitle = True
    .ChartTitle.Text = "Ex：Ｘ方向誤差"
    .ChartTitle.Characters.Font.Size = 11
    .ChartTitle.top = 30
    
    .ChartArea.Border.Weight = 1
    .ChartArea.Border.LineStyle = 0
'    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("j3:j4097"), PlotBy:=xlColumns
    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("J" & STARTROW & ":J" & ENDROW), PlotBy:=xlColumns
'    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
    .SeriesCollection(1).MarkerSize = 2
    .Axes(xlCategory).TickLabels.AutoScaleFont = False
    .Axes(xlValue).TickLabels.AutoScaleFont = False
    .Axes(xlCategory).TickLabels.Font.Size = 10.75
    .Axes(xlValue).TickLabels.Font.Size = 10.5
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "pls"
    .PlotArea.width = 750
    i = (.PlotArea.width - .PlotArea.InsideWidth)
    .PlotArea.width = 630 + i
    .PlotArea.Left = 25
    i = (.PlotArea.InsideLeft - .PlotArea.Left)
    .PlotArea.Left = 72 - i
    .PlotArea.top = 45
    .PlotArea.Height = 78
  End With
'どういうときにどれだけボタンがずれるのか謎。絶対座標指定はないのか
'ActiveChart.Shapes("Button 7").Select
'Selection.ShapeRange.Left 200
'ActiveChart.Shapes("Button 8").Select
'Selection.ShapeRange.Left 500
End If

If Not IsEmpty(Worksheets(SHEETNAME2).Range("k3").Value) Then  'kに値がないとエラーになるので。
  ActiveSheet.ChartObjects("グラフ 3").Activate
  With ActiveChart
    .HasTitle = True
    .ChartTitle.Text = "Ey：Ｙ方向誤差"
    .ChartTitle.Characters.Font.Size = 11
    .ChartTitle.top = 30
    
    .ChartArea.Border.Weight = 1
    .ChartArea.Border.LineStyle = 0
'    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("k3:k4097"), PlotBy:=xlColumns
'    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("K" & STARTROW & ":K" & ENDROW), PlotBy:=xlColumns
    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
    .SeriesCollection(1).MarkerSize = 2
    .Axes(xlCategory).TickLabels.AutoScaleFont = False
    .Axes(xlValue).TickLabels.AutoScaleFont = False
    .Axes(xlCategory).TickLabels.Font.Size = 10.75
    .Axes(xlValue).TickLabels.Font.Size = 10.5
    .Axes(xlCategory, xlPrimary).HasTitle = True
    .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "ﾃﾞｰﾀpoint"
    .Axes(xlValue, xlPrimary).HasTitle = True
    .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "pls"
    .PlotArea.width = 750
    i = (.PlotArea.width - .PlotArea.InsideWidth)
    .PlotArea.width = 630 + i
    .PlotArea.Left = 25
    i = (.PlotArea.InsideLeft - .PlotArea.Left)
    .PlotArea.Left = 72 - i
    .PlotArea.top = 45
    .PlotArea.Height = 78
  End With
End If

'ActiveSheet.ChartObjects("グラフ 1").Activate    'グラフが消える現象防止
'ActiveSheet.ChartObjects("グラフ 2").Activate
Application.ScreenUpdating = True
SendKeys "{esc}", False

End Function
Sub ボタン2()
With UserForm4
   .TextBox1.Value = Worksheets(SHEETNAME1).Range(RXB).Value    'シート１からデータを拾ってきて
   .TextBox2.Value = Worksheets(SHEETNAME1).Range(RXE).Value    'テキストボックスに表示
   .TextBox3.Value = Worksheets(SHEETNAME1).Range(RYB).Value
   .TextBox4.Value = Worksheets(SHEETNAME1).Range(RYE).Value
   .TextBox6.Value = Worksheets(SHEETNAME1).Range(STARTPOINT).Value
   .TextBox7.Value = Worksheets(SHEETNAME1).Range(ENDPOINT).Value
  .Show
End With
End Sub
Sub 再計算()
Dim X As Range, Y As Range, EX As Range, EY As Range, SN As Range, l As Range
Dim i As Integer
Dim n As String, csvfname As String

Application.ScreenUpdating = False
n = ActiveSheet.Name       'ボタンを押したシートに戻るように。
With UserForm4
  If Not (IsNumeric(.TextBox1.Value) And IsNumeric(.TextBox2.Value) And IsNumeric(.TextBox3.Value) And IsNumeric(.TextBox4.Value) And IsNumeric(.TextBox6.Value) And IsNumeric(.TextBox7.Value)) Then
    MsgBox "数値を入力してください"    'テキストボックスに文字が入っている場合×。基準直線と間引きは最低限必要
    Exit Sub
  End If
  rx0 = .TextBox1.Value
  rx1 = .TextBox2.Value
  ry0 = .TextBox3.Value
  ry1 = .TextBox4.Value
  mabiki = 1
  start = .TextBox6.Value
  endnum = .TextBox7.Value
  .Hide
End With
With Worksheets(SHEETNAME1)
  Select Case mabiki
  Case 0
    mabiki = 1    '間引きが0の時１
  Case Is < 0
    MsgBox "「間引き」には正の整数を入力してください"    '負数はエラー。いらないかも
    Exit Sub
  End Select
  
  Select Case start
  Case 0
    start = 1     '0は1とする。0がくることはないと思う。
  Case Is < 0
    MsgBox "「開始」には正の整数を入力してください"  '負数はエラー。いらないかも。
    Exit Sub
  End Select
  
  Select Case endnum
  Case 0
'    endnum = 4067    '０のときはデータすべて。いらんかも。
'*****Has
    endnum = ENDROW
  Case Is < start
    MsgBox "「終了」には「開始」より小さい数を入力することはできません"    '入力制限がちがち。いらないかも
    Exit Sub
  End Select
'  If IsEmpty(rx0) Or IsEmpty(rx1) Or IsEmpty(ry0) Or IsEmpty(ry1) Then
'    MsgBox "基準直線を設定してください"    'テキストボックスが空だとisnumericに引っかかる。
'    Exit Sub
'  End If
  
  .Range(STARTPOINT) = start
  .Range(ENDPOINT) = endnum
  csvfname = .Range(FILENAMEPOINT)
  If (Not (.Range(RXB).Value = rx0 And .Range(RXE).Value = rx1 And .Range(RYB).Value = ry0 And .Range(RYE).Value = ry1)) Then
    .Range(RXB).Value = rx0  '基準直線が一つでもシートの値と違う場合は直線精度を呼ぶ。
    .Range(RXE).Value = rx1
    .Range(RYB).Value = ry0
    .Range(RYE).Value = ry1
    Call linearity_chk(csvfname)
    Sheets(n).Activate    '帰ってきた後に最初にアクティブだったシートを選択。
    Exit Sub
  End If
End With

Worksheets(SHEETNAME2).Activate
With Worksheets(SHEETNAME2)
  If IsEmpty(.Cells(3, 1).Value) Or .Cells(3, 1) <> Worksheets(SHEETNAME1).Cells(3, 1) Then
    Call linearity_chk(csvfname) 'csv読み込んだ後グラフ描画ボタンを押さなかった時を想定。
    Sheets(n).Activate
    Exit Sub
  End If

'  .Range("f3:k4097").ClearContents
  .Range("F" & STARTROW & ":K" & ENDROW).ClearContents
  i = start + 2
  Set X = .Cells(3, 8)
  Set Y = .Cells(3, 9)
  Set EX = .Cells(3, 10)
  Set EY = .Cells(3, 11)
  Set SN = .Cells(3, 7)
  Set l = .Cells(3, 12)
    While (i < endnum + 3) And (Not IsEmpty(.Cells(i, 1).Value))
      If (i - 2 - mabiki * Int((i - 2) / mabiki)) = 0 Then
        X.Value = .Cells(i, 1).Value
        Y.Value = .Cells(i, 2).Value
        EX.Value = .Cells(i, 3).Value
        EY.Value = .Cells(i, 4).Value
        l.Value = .Cells(i, 5).Value
        SN.Value = i - 2
        Set X = X.Offset(1, 0)
        Set Y = Y.Offset(1, 0)
        Set EX = EX.Offset(1, 0)
        Set EY = EY.Offset(1, 0)    '移すだけの作業だが遅い気がする
        Set SN = SN.Offset(1, 0)
        Set l = l.Offset(1, 0)
      End If
    i = i + 1
    Wend
.Cells(1, 3).Select
End With
endnum = i - 3
Worksheets(SHEETNAME1).Range(ENDPOINT) = endnum

If Not IsEmpty(Worksheets(SHEETNAME2).Range("h3").Value) Then  'hにデータがないとエラーがでる。
  Charts(CHARTNAME31).Select
'  ActiveChart.SetSourceData Source:=Worksheets(2).Range("h3:h4097"), PlotBy:=xlColumns  'グラフにデータがないとき必要
  ActiveChart.SetSourceData Source:=Worksheets(2).Range("H" & STARTROW & ":H" & ENDROW), PlotBy:=xlColumns  'グラフにデータがないとき必要
'  ActiveChart.SeriesCollection(1).XValues = Worksheets(2).Range("g3:g4097")
  ActiveChart.SeriesCollection(1).XValues = Worksheets(2).Range("G" & STARTROW & ":G" & ENDROW)
  ActiveChart.SeriesCollection.NewSeries                    '基準直線を書く
  ActiveChart.SeriesCollection(2).XValues = Worksheets(1).Range("g4:g5")
  ActiveChart.SeriesCollection(2).Values = Worksheets(1).Range("h4:h5")
    With ActiveChart
  .SeriesCollection(1).ChartType = xlXYScatter
  .SeriesCollection(1).MarkerSize = 2
  .SeriesCollection(2).ChartType = xlXYScatterLinesNoMarkers
  .HasTitle = True
  .ChartTitle.Characters.Text = "軌跡" & Chr(10) & Now & Chr(10) & Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value
End With
End If

'この辺はグラフを作るところでフォローしたい。userform4
If Not IsEmpty(Worksheets(SHEETNAME2).Range(YPOINT).Value) Then  'hがないとエラーに
  Charts(CHARTNAME31).Select
'  ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("i3:i4097"), PlotBy:=xlColumns
'  ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("h3:h4097")
  ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("I" & STARTROW & ":I" & ENDROW), PlotBy:=xlColumns
  ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("H" & STARTROW & ":H" & ENDROW)
  ActiveChart.SeriesCollection.NewSeries
  ActiveChart.SeriesCollection(2).XValues = Worksheets(SHEETNAME1).Range("g7:g8")  '基準直線
  ActiveChart.SeriesCollection(2).Values = Worksheets(SHEETNAME1).Range("h7:h8")
With ActiveChart
  .SeriesCollection(1).ChartType = xlXYScatter
  .SeriesCollection(1).MarkerSize = 2
  .SeriesCollection(2).ChartType = xlXYScatterLinesNoMarkers
  .HasTitle = True
  .ChartTitle.Characters.Text = "軌跡" & Chr(10) & Now & Chr(10) & Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value
End With
End If

Sheets(CHARTNAME32).Activate
If Not IsEmpty(Worksheets(SHEETNAME2).Range(ERRORL1).Value) Then 'lに値がないとエラーになるので。
  ActiveSheet.ChartObjects("グラフ 1").Activate
  With ActiveChart
'    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("l3:l4097"), PlotBy:=xlColumns
'    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumns
    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
    .SeriesCollection(1).MarkerSize = 2
    .HasTitle = True
    .ChartTitle.Characters.Text = "軸別直線精度" & Chr(10) & Now & Chr(10) & Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value & Chr(10) & "L：ＸＹ合成誤差"
    .PlotArea.width = 650
    .PlotArea.Left = 25
    .PlotArea.top = 55
    .PlotArea.Height = 78
  End With
End If

If Not IsEmpty(Worksheets(SHEETNAME2).Range(ERRORX1).Value) Then  'jに値がないとエラーに
  ActiveSheet.ChartObjects("グラフ 2").Activate
'  ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("j3:j4097"), PlotBy:=xlColumns
  ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("J" & STARTROW & ":J" & ENDROW), PlotBy:=xlColumns
'  ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
  ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
  ActiveChart.HasTitle = True
  ActiveChart.ChartTitle.Characters.Text = "Ex：方向誤差"
  ActiveChart.PlotArea.width = 650
  ActiveChart.PlotArea.Left = 25
  ActiveChart.ChartTitle.top = 15
'どういうときにどれだけボタンがずれるのか謎。絶対座標指定はないのか
'ActiveChart.Shapes("Button 7").Select
'Selection.ShapeRange.Left 200
'ActiveChart.Shapes("Button 8").Select
'Selection.ShapeRange.Left 500
End If

If Not IsEmpty(Worksheets(SHEETNAME2).Range("k3").Value) Then  'kに値がないとエラーになるので。
  ActiveSheet.ChartObjects("グラフ 3").Activate
'  ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("k3:k4097"), PlotBy:=xlColumns
'  ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
  ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("K" & STARTROW & ":K" & ENDROW), PlotBy:=xlColumns
  ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
  ActiveChart.HasTitle = True
  ActiveChart.ChartTitle.Characters.Text = "Ey：Ｙ方向誤差"
  ActiveChart.PlotArea.width = 650
  ActiveChart.PlotArea.Left = 25
  ActiveChart.ChartTitle.top = 15
End If

'ActiveSheet.ChartObjects("グラフ 1").Activate    'グラフが消える現象防止
'ActiveSheet.ChartObjects("グラフ 2").Activate
Sheets(n).Activate
Application.ScreenUpdating = True
End Sub
' change～はUserForm4でテキストボックスに好ましくない数値が入力されて別の場所をクリックしたときに
Sub change1()          'シート１から値を拾ってきて変更する
With UserForm4
  If IsEmpty(.TextBox1.Value) Or (Not IsNumeric(.TextBox1.Value)) Then
    .TextBox1.Value = Worksheets(SHEETNAME1).Range(RXB).Value
  End If
End With
End Sub
Sub change2()
With UserForm4
  If IsEmpty(.TextBox2.Value) Or (Not IsNumeric(.TextBox2.Value)) Then
    .TextBox2.Value = Worksheets(SHEETNAME1).Range(RXE).Value
  End If
End With
End Sub
Sub change3()
With UserForm4
  If IsEmpty(.TextBox3.Value) Or (Not IsNumeric(.TextBox3.Value)) Then
    .TextBox3.Value = Worksheets(SHEETNAME1).Range(RYB).Value
  End If
End With
End Sub
Sub change4()
With UserForm4
  If IsEmpty(.TextBox4.Value) Or (Not IsNumeric(.TextBox4.Value)) Then
    .TextBox4.Value = Worksheets(SHEETNAME1).Range(RYE).Value
  End If
End With
End Sub
Sub change5()
'Dim a As Integer
'With UserForm4.TextBox5
'  If IsEmpty(.Value) Or (Not IsNumeric(.Value)) Or .Value < 0 Then
'    a = 1
'  ElseIf .Value <= 1 Then        'いったんint型に入れてから代入し直すことで
'    a = 1                        '小数が入力されたときに整数にもどる。
'  Else
'    a = 1
'  End If
'  .Value = a
'End With
End Sub
Sub change6()
Dim a As Integer
With UserForm4.TextBox6
  If IsEmpty(.Value) Or (Not IsNumeric(.Value)) Or .Value < 1 Then
     a = Worksheets(SHEETNAME1).Range(STARTPOINT).Value
  Else
     a = .Value
  End If
  .Value = a
End With
End Sub
Sub change7()
Dim a As Integer
With UserForm4.TextBox7
  If IsEmpty(.Value) Or (Not IsNumeric(.Value)) Or .Value <= UserForm4.TextBox6 Then
    a = Worksheets(SHEETNAME1).Range(ENDPOINT).Value
  Else
    a = .Value
  End If
  .Value = a
End With
End Sub
Sub ボタン()    '「軸別直線精度」の右のボタンに対応。ダイヤログを表示するまでの下準備
Application.ScreenUpdating = False
With UserForm2
  ActiveSheet.ChartObjects("グラフ 1").Activate
  .TextBox6.Value = ActiveChart.Axes(xlCategory).MinimumScale
  .TextBox7.Value = ActiveChart.Axes(xlCategory).MaximumScale
  .TextBox8.Value = ActiveChart.Axes(xlValue).MinimumScale
  .TextBox9.Value = ActiveChart.Axes(xlValue).MaximumScale

  .CheckBox1.Value = ActiveChart.Axes(xlCategory).MinimumScaleIsAuto
  .CheckBox3.Value = ActiveChart.Axes(xlCategory).MaximumScaleIsAuto
  .CheckBox2.Value = ActiveChart.Axes(xlValue).MinimumScaleIsAuto
  .CheckBox4.Value = ActiveChart.Axes(xlValue).MaximumScaleIsAuto
  .CheckBox5.Value = ActiveChart.Axes(xlValue).HasMajorGridlines

  ActiveSheet.ChartObjects("グラフ 2").Activate
  .TextBox10.Value = ActiveChart.Axes(xlValue).MinimumScale
  .TextBox11.Value = ActiveChart.Axes(xlValue).MaximumScale

  .CheckBox7.Value = ActiveChart.Axes(xlValue).MinimumScaleIsAuto
  .CheckBox8.Value = ActiveChart.Axes(xlValue).MaximumScaleIsAuto
  .CheckBox11.Value = ActiveChart.Axes(xlValue).HasMajorGridlines

  ActiveSheet.ChartObjects("グラフ 3").Activate
  .TextBox12.Value = ActiveChart.Axes(xlValue).MinimumScale
  .TextBox13.Value = ActiveChart.Axes(xlValue).MaximumScale

  .CheckBox9.Value = ActiveChart.Axes(xlValue).MinimumScaleIsAuto
  .CheckBox10.Value = ActiveChart.Axes(xlValue).MaximumScaleIsAuto
  .CheckBox12.Value = ActiveChart.Axes(xlValue).HasMajorGridlines
  
  .Show
End With
Application.ScreenUpdating = True
End Sub
Sub ボタン3()        '「軌跡」の右ボタンに対応。ダイヤログを表示するまでの下準備
Call form3_ini
End Sub
Sub 実験()    '高速化実験
'while isemptyの代わりに使う。データ数が多いときは倍くらい違う

Application.ScreenUpdating = False
Dim x0 As Double, y0 As Double, ex0 As Double, ey0 As Double, l0 As Double
Dim EX As Range, EY As Range, SN As Range, l As Range
Dim i As Integer, datnum As Integer, j As Integer
Dim dat As String

Worksheets(SHEETNAME2).Activate
Range("a3:l65536").Select
Selection.ClearContents '二枚目のシートを空に
Cells(3, 1).Select

Worksheets(SHEETNAME1).Activate
With Worksheets(SHEETNAME1)  'x座標始点　　　　　　x座標終点　　　　　　　　　　y座標始点　　　　　　　　　　　　y座標終点　　　　　　　　　　　　            開始　　　　　　　　　　　        終了
  If Not (IsNumeric(.Range(RXB).Value) And IsNumeric(.Range(RXE).Value) And IsNumeric(.Range(RYB).Value) And IsNumeric(.Range(RYE).Value) And IsNumeric(.Range(STARTPOINT).Value) And IsNumeric(.Range(ENDPOINT).Value)) Then
    MsgBox "数値を入力してください"    ' 入力部分に数値が入っていないときにエラーを出す。入力制限のためいらないかも。
    Exit Sub
  End If
  
  If IsEmpty(.Range(RXB).Value) Or IsEmpty(.Range(RXE).Value) Or IsEmpty(.Range(RYB).Value) Or IsEmpty(.Range(RYE).Value) Then
    MsgBox "基準直線を設定してください"    '基準直線が１つでもかけているとエラー
    Exit Sub
  End If
  
  rx0 = .Range(RXB).Value   '数値の代入
  rx1 = .Range(RXE).Value
  ry0 = .Range(RYB).Value
  ry1 = .Range(RYE).Value
  mabiki = 1
  start = .Range(STARTPOINT).Value
  endnum = .Range(ENDPOINT).Value
  
  Select Case mabiki
  Case 0
    mabiki = 1    '間引きが0もしくはemptyのとき1を代入
  Case Is < 0
    MsgBox "「間引き」には正の整数を入力してください"  '負数はエラー、いらないかも。
    Exit Sub
  End Select
  
  Select Case start
  Case 0
    start = 1 '0はデータの先頭からとする。いらない可能性大
    .Range(STARTPOINT).Value = start
  Case Is < 0
    MsgBox "「開始」には正の整数を入力してください"  '負数はエラー、いらないかも。
    Exit Sub
  End Select
  
  Select Case endnum
  Case 0
    endnum = ENDROW  '0またはemptyのときは最後のデータまで使用する。
  Case Is < start
    MsgBox "「終了」には「開始」より小さい数値を入力することはできません"  '負数はエラーいらないかも。
    Exit Sub
  End Select
  
  If rx0 = rx1 And ry0 = ry1 Then
    MsgBox "基準直線の始点と終点が同じになっています。再度設定してください。"    '基準直線の始点と終点の値が全く同じだとエラー
    Exit Sub
  End If
  
  .Range("a3:b4100").Select
   Selection.Copy
  .Cells(3, 1).Select
End With
  
Worksheets(SHEETNAME2).Activate
With Worksheets(SHEETNAME2)
  Range("a3:b4100").PasteSpecial
  Application.CutCopyMode = False         'ﾛｷﾞﾝｸﾞﾃﾞｰﾀを２枚目にコピー
  If (start Mod mabiki) = 0 Then
    i = start / mabiki
  Else
    i = Int(start / mabiki) + 1
  End If
  j = i
  datnum = .Range("a3:a4100").find(what:=Empty).Row - 3
  If endnum > datnum Then
    endnum = datnum
  End If
  dat = "a3:a" & Trim(str(endnum + 2))
  For Each EX In .Range(dat)
    x0 = EX.Value
    y0 = EX.Offset(0, 1).Value
    ex0 = ((ry1 - ry0) ^ 2 * (x0 - rx0) - (rx1 - rx0) * (ry1 - ry0) * (y0 - ry0)) / ((rx1 - rx0) ^ 2 + (ry1 - ry0) ^ 2)
    ey0 = ((rx1 - rx0) ^ 2 * (y0 - ry0) - (rx1 - rx0) * (ry1 - ry0) * (x0 - rx0)) / ((rx1 - rx0) ^ 2 + (ry1 - ry0) ^ 2)
    l0 = (ex0 ^ 2 + ey0 ^ 2) ^ 0.5
    EX.Offset(0, 2).Value = ex0
    EX.Offset(0, 3).Value = ey0
    EX.Offset(0, 4).Value = l0
  Next EX
  dat = "g3:g" & Trim(str(Int(endnum / mabiki + 2)))
  For Each SN In .Range(dat)
    SN.Value = i * mabiki
    i = i + 1
  Next SN
  dat = "h3:l" & Trim(str(Int(endnum / mabiki + 2)))
  For Each EY In .Range(dat)
    i = (EY.Row - j - 1) * (mabiki - 1)
    EY.Value = EY.Offset(i, -7).Value
  Next EY
End With
Worksheets(SHEETNAME1).Range(ENDPOINT).Value = endnum 'データ数をシートに還元

'この辺はグラフを作るところでフォローしたい。userform4
If Not IsEmpty(Worksheets(SHEETNAME2).Range(YPOINT).Value) Then  'hがないとエラーに
  Charts(CHARTNAME31).Select
'  ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("i3:i4097"), PlotBy:=xlColumns
'  ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("h3:h4097")
  ActiveChart.SetSourceData Source:=Worksheets(SHEETNAME2).Range("I" & STARTROW & ":I" & ENDROW), PlotBy:=xlColumns
  ActiveChart.SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("H" & STARTROW & ":H" & ENDROW)
  ActiveChart.SeriesCollection.NewSeries
  ActiveChart.SeriesCollection(2).XValues = Worksheets(SHEETNAME1).Range("g7:g8")  '基準直線
  ActiveChart.SeriesCollection(2).Values = Worksheets(SHEETNAME1).Range("h7:h8")
With ActiveChart
  .SeriesCollection(1).ChartType = xlXYScatter
  .SeriesCollection(1).MarkerSize = 2
  .SeriesCollection(2).ChartType = xlXYScatterLinesNoMarkers
  .HasTitle = True
  .ChartTitle.Characters.Text = "軌跡" & Chr(10) & Now & Chr(10) & Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value
End With
End If

Sheets(CHARTNAME32).Activate
If Not IsEmpty(Worksheets(SHEETNAME2).Range(ERRORL1).Value) Then 'lに値がないとエラーになるので。
  ActiveSheet.ChartObjects("グラフ 1").Activate
  With ActiveChart
'    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("l3:l4097"), PlotBy:=xlColumns
'    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW), PlotBy:=xlColumns
    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
    .SeriesCollection(1).MarkerSize = 2
    '.TickLabels.Font.Size = 10.75
    .HasTitle = True
    .ChartTitle.Characters.Text = "軸別直線精度" & Chr(10) & Now & Chr(10) & Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value & Chr(10) & "L：ＸＹ合成誤差"
    .PlotArea.width = 650
    .PlotArea.Left = 25
    .PlotArea.top = 55
    .PlotArea.Height = 78
  End With
End If

If Not IsEmpty(Worksheets(SHEETNAME2).Range(ERRORX1).Value) Then  'jに値がないとエラーに
  ActiveSheet.ChartObjects("グラフ 2").Activate
  With ActiveChart
 '   .SetSourceData Source:=Worksheets(SHEETNAME2).Range("j3:j4097"), PlotBy:=xlColumns
 '   .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("J" & STARTROW & ":J" & ENDROW), PlotBy:=xlColumns
    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
    .SeriesCollection(1).MarkerSize = 2
    '.TickLabels.Font.Size = 10.75
    .HasTitle = True
    .ChartTitle.Characters.Text = "Ex：Ｘ方向誤差"
    .Axes(xlCategory).TickLabelPosition = xlNone
    .PlotArea.width = 650
    .PlotArea.Left = 25
    .ChartTitle.top = 15
  End With
'どういうときにどれだけボタンがずれるのか謎。絶対座標指定はないのか
'ActiveChart.Shapes("Button 7").Select
'Selection.ShapeRange.Left 200
'ActiveChart.Shapes("Button 8").Select
'Selection.ShapeRange.Left 500
End If

If Not IsEmpty(Worksheets(SHEETNAME2).Range("k3").Value) Then  'kに値がないとエラーになるので。
  ActiveSheet.ChartObjects("グラフ 3").Activate
  With ActiveChart
'    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("k3:k4097"), PlotBy:=xlColumns
'    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
'    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("k3:k4097"), PlotBy:=xlColumns
'    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("g3:g4097")
    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("K" & STARTROW & ":K" & ENDROW), PlotBy:=xlColumns
    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
    .SetSourceData Source:=Worksheets(SHEETNAME2).Range("K" & STARTROW & ":K" & ENDROW), PlotBy:=xlColumns
    .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("G" & STARTROW & ":G" & ENDROW)
    .SeriesCollection(1).MarkerSize = 2
    '.TickLabels.Font.Size = 10.75
    .HasTitle = True
    .ChartTitle.Characters.Text = "Ey：Ｙ方向誤差"
    .Axes(xlCategory).TickLabelPosition = xlNone
    .PlotArea.width = 650
    .PlotArea.Left = 25
    .ChartTitle.top = 15
  End With
End If

'ActiveSheet.ChartObjects("グラフ 1").Activate
'ActiveSheet.ChartObjects("グラフ 2").Activate 'グラフが消える現象防止
Application.ScreenUpdating = True

End Sub
'------------------------------------------------------------------------
' Load CSV File
'------------------------------------------------------------------------
Public Function LoadCsvFile(LoadSheet As String, LoadArea As String, title As String, clmnum As Integer) As String
    Dim top As Range, area As Range
    Dim FileName As Variant, F_Num As Integer
    Dim str As String
    Dim cnt As Long, lmt As Long, maxclm As Long
    Dim linestr As String
    Dim tmp As Range
    Dim linnum As Long, maxlinnum As Long
    Dim scupflg As Integer, linchk As Integer
    Dim fs As Object, wkFSO As Object
    
    scupflg = 0
    If (False = Application.ScreenUpdating) Then
     Application.ScreenUpdating = True
     scupflg = 1
    End If
    FileName = Application.GetOpenFilename("Log File,*.csv", 1, title, "開く", False)
    
    If (1 = scupflg) Then
     Application.ScreenUpdating = False
    End If
    If TypeName(FileName) = "Boolean" Then
        LoadCsvFile = ""
        Exit Function
    End If

    On Error Resume Next

    Worksheets(LoadSheet).Activate
    Set area = Worksheets(LoadSheet).Range(LoadArea)
    Set top = area.Item(1, 1)

    area.Select
    Selection.ClearContents
    top.Select
    maxlinnum = area.Rows.Count
    
    'Way of 1line read
    'F_Num = FreeFile
    'Open FileName For Binary As F_Num
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set wkFSO = fs.OpenTextFile(FileName, 1, True)
    If (0 = clmnum) Then
        lmt = -1
    Else
        lmt = clmnum + 1
    End If
    If Not (IsNull(wkFSO)) Then
        maxclm = clmnum
        linnum = 0
        Do While wkFSO.AtEndOfStream <> True
         linestr = wkFSO.ReadLine()
         linchk = UBound(Split(linestr, ",", lmt, vbBinaryCompare)) + 1
         If (0 >= clmnum) Then
          maxclm = linchk
         ElseIf ((linchk) < (clmnum)) Then
          maxclm = linchk
         Else
          maxclm = clmnum
         End If
         Set tmp = top
         For cnt = 0 To maxclm - 1
          str = Split(linestr, ",", lmt, vbBinaryCompare)(cnt)
          If ("" <> str) Then
           tmp.Value = str
          End If
          Set tmp = tmp.Offset(0, 1)
         Next cnt
         Set top = top.Offset(1, 0)
         linnum = linnum + 1
         If (maxlinnum <= linnum) Then
          Exit Do
         End If
        Loop
        wkFSO.Close
    Else
     MsgBox "Internet Exploerのバージョンが古い可能性があります。"
    End If
    LoadCsvFile = CreateObject("Scripting.FileSystemObject").GetFileName(FileName)

End Function
