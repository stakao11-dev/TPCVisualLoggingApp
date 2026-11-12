Attribute VB_Name = "Module21"
Option Explicit
'ここにはcircleckeckとAnalizeErrがあります。
'どちらも真円精度のマクロを少しいじっただけのものです。
Sub circleckeck(csvfname As String)
'
' circleckeck Macro
' マクロ記録日 : 2000/8/20  ユーザー名 : kanetomo nozaki
'
    Application.ScreenUpdating = False  '画面再描画禁止

    Dim Pi, Ofs, Radius As Double
    Dim centerX, centerY, amp As Double

    Worksheets(SHEETNAME2).Activate            'シートの選択
    Range("A3:U65536").Select
    Selection.ClearContents
    Range("A3").Select

    Worksheets(SHEETNAME1).Activate            'シートの選択
    Radius = Range(RPOINT).Value                 '半径Rの取得
    centerX = Range(XCENTER).Value                'X軸中心の取得
    centerY = Range(YCENTER).Value                'Y軸中心の取得
    amp = Range(AMPPOINT).Value                    '拡大率の取得

    Worksheets(SHEETNAME2).Activate               'シートの選択
    With Worksheets(SHEETNAME2)
        Range("D1").Value = "理想円"  'セルの値を取得
        Range("D2").Value = "角度"    'セルの値を取得
        Range("E2").Value = "X"       'セルの値を取得
        Range("F2").Value = "Y"       'セルの値を取得
        Pi = Application.WorksheetFunction.Acos(-1)
        Ofs = Radius / 12 '実測値半径の±1/12を内側理想円・外側理想円の半径とする。'200
       
        'シフトﾛｷﾞﾝｸﾞﾃﾞｰﾀ作成(中心点が（０，０）になるようにシフト)
        Dim csvdatX, sftcsvdatX As Range
        Dim csvdatY, sftcsvdatY As Range
        Set csvdatX = Worksheets(SHEETNAME1).Range("A3")       '理想円用角度データ作成
        Set csvdatY = Worksheets(SHEETNAME1).Range("B3")       '理想円用角度データ作成
        Set sftcsvdatX = .Range("A3")
        Set sftcsvdatY = .Range("B3")
        Do While Not IsEmpty(csvdatX)
            sftcsvdatX.Value = csvdatX.Value - centerX             'Ｘ軸データ
            sftcsvdatY.Value = csvdatY.Value - centerY             'Ｙ軸データ
        
            Set csvdatX = csvdatX.Offset(1, 0)                     'X軸用データ挿入セルとして一つ下のセルを選択
            Set csvdatY = csvdatY.Offset(1, 0)                     'Y軸用データ挿入セルとして一つ下のセルを選択
            Set sftcsvdatX = sftcsvdatX.Offset(1, 0)               'X軸用データ挿入セルとして一つ下のセルを選択
            Set sftcsvdatY = sftcsvdatY.Offset(1, 0)               'Y軸用データ挿入セルとして一つ下のセルを選択
        Loop
        
        '理想円データ作成
        Dim cnt As Integer
        Dim temp1, temp2 As Double
        Dim degdata As Range
        Dim imgcirX, innerX, outerX As Range
        Dim imgcirY, innerY, outerY As Range
        Set degdata = .Range("D3")     '理想円用角度データ作成
        Set imgcirX = .Range("E3")     '理想円用X軸用データ作成
        Set imgcirY = .Range("F3")     '理想円用Y軸用データ作成
        Set innerX = .Range("H3")     '理想円用X軸用データ作成
        Set innerY = .Range("I3")     '理想円用X軸用データ作成
        Set outerX = .Range("K3")     '理想円用X軸用データ作成
        Set outerY = .Range("L3")     '理想円用X軸用データ作成
        cnt = 0
        
        Do While cnt <= 360
            degdata.Value = cnt                              '角度用データ
            temp1 = Sin(2 * Pi * cnt / 360)  'sin計算値データ
            temp2 = Cos(2 * Pi * cnt / 360)  'cos計算値データ
            imgcirX.Value = Radius * temp1
            imgcirY.Value = Radius * temp2
            innerX.Value = (Radius - Ofs) * temp1
            innerY.Value = (Radius - Ofs) * temp2
            outerX.Value = (Radius + Ofs) * temp1
            outerY.Value = (Radius + Ofs) * temp2
            
            Set imgcirX = imgcirX.Offset(1, 0)               'X軸用データ挿入セルとして一つ下のセルを選択
            Set imgcirY = imgcirY.Offset(1, 0)               'Y軸用データ挿入セルとして一つ下のセルを選択
            Set innerX = innerX.Offset(1, 0)               'X軸用データ挿入セルとして一つ下のセルを選択
            Set innerY = innerY.Offset(1, 0)               'Y軸用データ挿入セルとして一つ下のセルを選択
            Set outerX = outerX.Offset(1, 0)               'X軸用データ挿入セルとして一つ下のセルを選択
            Set outerY = outerY.Offset(1, 0)               'Y軸用データ挿入セルとして一つ下のセルを選択
            Set degdata = degdata.Offset(1, 0)               '次角度用データ挿入セルとして一つ下のセルを選択
            cnt = cnt + 1                                    '角度用データｲﾝｸﾘﾒﾝﾄ
        Loop
    End With
    
    Call AnalyzeErr(Radius, amp)
    
    ' フッターの設定(出来ればテキストBOXに入れたい)
    ActiveChart.PageSetup.RightFooter = Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value
    
    Application.ScreenUpdating = True   '画面描画有効
End Sub

Sub AnalyzeErr(ByVal Radius As Double, ByVal ErrScale As Double)
Dim difpntang As Double, tmpdat As Double, strtflg As Integer
    difpntang = strtflg = 0
    Worksheets(SHEETNAME2).Activate               'シートの選択
    With Worksheets(SHEETNAME2)
        Dim datX, amperrX As Range
        Dim datY, amperrY As Range
        Dim datX_V, datY_V As Double
        Dim ErrPitch, ExRadius, radian As Double, degree As Double
        Dim difpntX As Range, difpntY As Range
        Set datX = .Range("A3")        '理想円用角度データ作成
        Set datY = .Range("B3")        '理想円用角度データ作成
        Set amperrX = .Range("Q3")     '理想円用角度データ作成
        Set amperrY = .Range("R3")     '理想円用角度データ作成
        Set difpntX = .Range("T3")     '別色での描画データ作成
        Set difpntY = .Range("U3")     '別色での描画データ作成
        ErrPitch = Radius / 12 / ErrScale
        Do While Not IsEmpty(datX)
            datX_V = datX.Value
            datY_V = datY.Value
            'WorksheetfunctionのAtan2であれば分母０にも対応している
            'radian = Application.WorksheetFunction.Atan2(datX_V, datY_V)
            '********* add following at V15 *******
            radian = 0
            If ((datX_V * datY_V) <> 0) Then
                radian = Application.WorksheetFunction.Atan2(datX_V, datY_V)
            End If
            '*********
            ExRadius = Sqr((datX_V ^ 2) + (datY_V ^ 2))         ' 実半径の算出
            ExRadius = Radius + ErrPitch * (ExRadius - Radius)  ' 誤差拡大後の半径を算出
            amperrX.Value = Cos(radian) * ExRadius
            amperrY.Value = Sin(radian) * ExRadius
            If (2 <> strtflg) Then
                degree = Application.WorksheetFunction.Degrees(radian)
                degree = degree Mod 360
                If (degree < 0) Then: degree = degree + 360
                If (0 = strtflg) Then
                    difpntang = degree
                    strtflg = 1
                ElseIf ((degree > (difpntang + 30)) Or (degree < (difpntang - 30))) Then
                    strtflg = 2
                End If
                difpntX.Value = amperrX.Value
                difpntY.Value = amperrY.Value
                Set difpntX = difpntX.Offset(1, 0)            '一つ下のセルを選択
                Set difpntY = difpntY.Offset(1, 0)            '一つ下のセルを選択
            End If
            Set datX = datX.Offset(1, 0)            '一つ下のセルを選択
            Set datY = datY.Offset(1, 0)            '一つ下のセルを選択
            Set amperrX = amperrX.Offset(1, 0)      '一つ下のセルを選択
            Set amperrY = amperrY.Offset(1, 0)      '一つ下のセルを選択
        Loop
    End With
    Sheets(CHARTNAME21).Select
    With ActiveChart
      .HasTitle = True
      'グラフタイトルは自動で中央位置に置かれる。
      .ChartTitle.top = 1
      .ChartTitle.Characters.Text = "真円精度 ( 半径：" & Radius & " ｽｹｰﾙ：" & ErrScale & " )" & Chr(10) & Now & Chr(10) & Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value
      .ChartTitle.Characters.Font.Size = 11
      .PlotArea.Left = 160
      .PlotArea.Height = 400
      .PlotArea.width = 470
      .PlotArea.top = 45
      .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Font.Size = 14
      .Axes(xlCategory, xlPrimary).AxisTitle.Left = 580
      .Axes(xlCategory, xlPrimary).AxisTitle.top = 243
      .Axes(xlValue, xlPrimary).AxisTitle.Characters.Font.Size = 14
      .Axes(xlValue, xlPrimary).AxisTitle.Left = 363
      .Axes(xlValue, xlPrimary).AxisTitle.top = 38
    End With
    
    Sheets(CHARTNAME21).Select
    With ActiveChart
    'ｆに値がないとエラーとなる。まずないことですが。if必要？
'   .SetSourceData Source:=Worksheets(SHEETNAME2).Range("f3:f4097")
'   .SetSourceData Source:=Worksheets(SHEETNAME2).Range("f" & STRTROW & ":" & "f" & ENDROW)
   .SetSourceData Source:=Worksheets(SHEETNAME2).Range("F" & STARTROW & ":F" & ENDROW)
'      .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("e3:e4097")
      .SeriesCollection(1).XValues = Worksheets(SHEETNAME2).Range("E" & STARTROW & ":E" & ENDROW)
      .SeriesCollection(1).Border.ColorIndex = 7
      .SeriesCollection(1).Name = "=""理想円"""
    End With
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(2)
'      .XValues = Worksheets(SHEETNAME2).Range("h3:h4097")
'      .Values = Worksheets(SHEETNAME2).Range("i3:i4097")
      .XValues = Worksheets(SHEETNAME2).Range("H" & STARTROW & ":H" & ENDROW)
      .Values = Worksheets(SHEETNAME2).Range("I" & STARTROW & ":I" & ENDROW)
      '.Border.ColorIndex = 1
      .Border.LineStyle = xlNone
      .MarkerBackgroundColorIndex = 1
      .MarkerStyle = xlSquare
      .MarkerSize = 2
      .Name = "=""理想円 - " & ErrScale & """"
    End With
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(3)
'      .XValues = Worksheets(SHEETNAME2).Range("k3:k4097")
'      .Values = Worksheets(SHEETNAME2).Range("l3:l4097")
      .XValues = Worksheets(SHEETNAME2).Range("K" & STARTROW & ":K" & ENDROW)
      .Values = Worksheets(SHEETNAME2).Range("L" & STARTROW & ":L" & ENDROW)
      '.Border.ColorIndex = 1
      .Border.LineStyle = xlNone
      .MarkerBackgroundColorIndex = 1
      .MarkerStyle = xlSquare
      .MarkerSize = 2
      .Name = "=""理想円 + " & ErrScale & """"
    End With
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(4)
'      .XValues = Worksheets(SHEETNAME2).Range("q3:q4097")
'Ｑに値がないとエラーとなる。具体的にはｃｓｖファイルを読み込んでいないとき。if必要？
'      .Values = Worksheets(SHEETNAME2).Range("r3:r4097")
      .XValues = Worksheets(SHEETNAME2).Range("Q" & STARTROW & ":Q" & ENDROW)
'Ｑに値がないとエラーとなる。具体的にはｃｓｖファイルを読み込んでいないとき。if必要？
      .Values = Worksheets(SHEETNAME2).Range("R" & STARTROW & ":R" & ENDROW)
      .ChartType = xlXYScatter
      .MarkerSize = 2
      .MarkerBackgroundColorIndex = 55
      .MarkerForegroundColorIndex = 55
      .MarkerStyle = xlSquare
      .Name = "=""ﾛｷﾞﾝｸﾞﾃﾞｰﾀ"""
    End With
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(5) 'スタートから３０度軌跡表示
'      .XValues = Worksheets(SHEETNAME2).Range("t3:t4097")
'      .Values = Worksheets(SHEETNAME2).Range("u3:u4097")
      .XValues = Worksheets(SHEETNAME2).Range("T" & STARTROW & ":T" & ENDROW)
      .Values = Worksheets(SHEETNAME2).Range("U" & STARTROW & ":U" & ENDROW)
      .ChartType = xlXYScatter
      .Border.LineStyle = xlNone
      .MarkerBackgroundColorIndex = 9
      .MarkerForegroundColorIndex = 9
      .MarkerStyle = xlCircle
      .MarkerSize = 4
      .Name = "=""開始30°範囲"""
    End With
    ActiveChart.SeriesCollection.NewSeries
    With ActiveChart.SeriesCollection(6) 'スタート点表示
      .XValues = Worksheets(SHEETNAME2).Range("q3")
      .Values = Worksheets(SHEETNAME2).Range("r3")
      .ChartType = xlXYScatter
      .Border.LineStyle = xlNone
      .MarkerBackgroundColorIndex = xlNone
      .MarkerForegroundColorIndex = 9
      .MarkerStyle = xlCircle
      .MarkerSize = 10
      .Name = "=""スタート点"""
    End With
        
    ActiveChart.Axes(xlCategory).Select
    With ActiveChart.Axes(xlCategory)
        .MinimumScale = -(Radius * 7 / 6)
        .MaximumScale = (Radius * 7 / 6)
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNone
    End With
    ActiveChart.Axes(xlValue).Select
    With ActiveChart.Axes(xlValue)
        .MinimumScale = -(Radius * 7 / 6)
        .MaximumScale = (Radius * 7 / 6)
        .MajorTickMark = xlNone
        .MinorTickMark = xlNone
        .TickLabelPosition = xlNone
    End With

    With ActiveChart
        If .HasLegend Then
            .HasLegend = False
        End If
        .HasLegend = True
        With .Legend
            Dim i As Integer
            .Position = xlRight
            .Left = 592
            .top = 7
        End With
    End With

    CirGraph_SizeCalc

End Sub
Sub CirGraph_SizeCalc()
    ' プロットエリアの設定（等比）
    Dim pw As Long
    Dim ph As Long
    ph = ActiveChart.ChartArea.Height * 9 / 10
    pw = ActiveChart.ChartArea.width * 9 / 10
    If ph > pw Then
        ph = pw
    End If
    If pw > ph Then
        pw = ph
    End If
    ActiveChart.PlotArea.Height = ph
    ActiveChart.PlotArea.width = pw
End Sub


' 円全体シート[グラフ表示設定変更]ボタン押下時の処理
Sub 表示項目変更()
Call form3_ini
End Sub

' 円全体シート[設定項目変更]ボタン押下時の処理
Sub 設定項目変更()
    With Worksheets(SHEETNAME1)            'シートの選択
        UserForm8.TextBox1.Text = .Range(RPOINT).Value  '半径Rをコントロールへ数値の書込み
        UserForm8.TextBox2.Text = .Range(XCENTER).Value  'X軸中心をコントロールへ数値の書込み
        UserForm8.TextBox3.Text = .Range(YCENTER).Value  'Y軸中心をコントロールへ数値の書込み
        UserForm8.TextBox4.Text = .Range(AMPPOINT).Value  '拡大率をコントロールへ数値の書込み
    End With
    UserForm8.Show
End Sub


