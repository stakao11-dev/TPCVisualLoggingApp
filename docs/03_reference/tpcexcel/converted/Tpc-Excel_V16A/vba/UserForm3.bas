Attribute VB_Name = "UserForm3"
Attribute VB_Base = "0{5B38302A-348D-4942-92AB-31F91B19368B}{509138FE-84F6-4304-9747-4D243951742D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit
Private Sub CheckBox1_Click()   'チェックが付いているときはテキストボックスを書き込めなくする。
With UserForm3
  If .CheckBox1.Value Then
    .TextBox1.Enabled = False
    .TextBox1.Value = ActiveChart.Axes(xlCategory).MinimumScale  'いらないかも
  Else
    .TextBox1.Enabled = True
  End If
End With
End Sub
Private Sub CheckBox3_Click()  'チェックが付いているときはテキストボックスを書き込めなくする。
With UserForm3
  If .CheckBox3.Value Then
    .TextBox2.Enabled = False
    .TextBox2.Value = ActiveChart.Axes(xlCategory).MaximumScale
  Else
    .TextBox2.Enabled = True
  End If
End With
End Sub
Private Sub CheckBox2_Click()  'チェックが付いているときはテキストボックスを書き込めなくする。
With UserForm3
  If .CheckBox2.Value Then
    .TextBox3.Enabled = False
    .TextBox3.Value = ActiveChart.Axes(xlValue).MinimumScale
  Else
    .TextBox3.Enabled = True
  End If
End With
End Sub
Private Sub CheckBox4_Click()  'チェックが付いているときはテキストボックスを書き込めなくする。
With UserForm3
  If .CheckBox4.Value Then
    .TextBox4.Enabled = False
    .TextBox4.Value = ActiveChart.Axes(xlValue).MaximumScale
  Else
    .TextBox4.Enabled = True
  End If
End With
End Sub
Private Sub CheckBox6_Click()  'チェックが付いているときはテキストボックスを書き込めなくする。
With UserForm3
  .TextBox5.Enabled = .CheckBox6.Value
End With
End Sub
Private Sub CheckBox5_Click()                '縦軸の目盛り線の追加
If UserForm3.CheckBox5.Value Then
  ActiveChart.Axes(xlValue).HasMajorGridlines = True
  With ActiveChart.Axes(xlValue).MajorGridlines.Border
    .Weight = xlHairline
    .LineStyle = xlDot
  End With
Else
ActiveChart.Axes(xlValue).HasMajorGridlines = False
End If
End Sub
Private Sub CommandButton1_Click()
Dim txboxname_str As String, retaintxboxname As String, i As Integer
Dim retaintxboxvisible As String, retaintxboxstr As String
With UserForm3
  If Not (IsNumeric(.TextBox1.Value) And IsNumeric(.TextBox2.Value) And IsNumeric(.TextBox3.Value) And IsNumeric(.TextBox4.Value)) Then
    MsgBox "数値を入力してください"
    Exit Sub
  End If

  Application.ScreenUpdating = False    'チェックが付いていればオート、

  If .CheckBox1.Value Then
    ActiveChart.Axes(xlCategory).MinimumScaleIsAuto = True
  Else                                  '付いていなければテキストボックス参照
    ActiveChart.Axes(xlCategory).MinimumScale = .TextBox1.Value
  End If

  If .CheckBox3.Value Then
    ActiveChart.Axes(xlCategory).MaximumScaleIsAuto = True
  Else
    ActiveChart.Axes(xlCategory).MaximumScale = .TextBox2.Value
  End If

  If .CheckBox2.Value Then
    ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
  Else
    ActiveChart.Axes(xlValue).MinimumScale = .TextBox3.Value
  End If

  If .CheckBox4.Value Then
    ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
  Else
    ActiveChart.Axes(xlValue).MaximumScale = .TextBox4.Value
  End If
  
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
  'コメント挿入前処理(関数化したい)
  
  If .CheckBox6.Value Then            'キャプションをつけるかどうかの判定
'HasTitleの処理がエラーになるになる場合はこの処理を使う
    If (txboxname_str = "") Then
     ActiveChart.Shapes.AddTextbox(msoTextOrientationHorizontal, 230.25, 70.25, 7.2, 7.2).Select
     Selection.ShapeRange.ScaleWidth 6.91, msoFalse, msoScaleFromTopLeft
    Else
     ActiveChart.Shapes.Item(txboxname_str).Visible = msoCTrue
     ActiveChart.Shapes(txboxname_str).Select
    End If
    Worksheets(SHEETNAME1).Range(retaintxboxname).Value = Selection.Name
    Selection.Characters.Text = .TextBox5.Text
    Worksheets(SHEETNAME1).Range(retaintxboxvisible).Value = 1
    Selection.Characters.Text = Replace(Selection.Characters.Text, Chr(13), "")
    Selection.AutoScaleFont = False
    Selection.AutoSize = True
          
'この処理を入れるとエラーが発生し開けなくなる？
'     ActiveChart.Axes(xlValue, xlPrimary).HasTitle = True
'     ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = "Filename:" & Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Text & .TextBox5.Text
'     ActiveChart.Axes(xlValue).AxisTitle.Orientation = xlHorizontal
'     ActiveChart.Axes(xlValue).AxisTitle.Left = 125
'     ActiveChart.Axes(xlValue).AxisTitle.Top = 100
'     ActiveChart.PlotArea.Left = 1
'     ActiveChart.PlotArea.width = 750
  Else
    If (txboxname_str <> "") Then
     ActiveChart.Shapes.Item(txboxname_str).Visible = msoFalse
    End If
    Worksheets(SHEETNAME1).Range(retaintxboxvisible).Value = 0
    'ActiveChart.Axes(xlValue).HasTitle = False
  End If
  Worksheets(SHEETNAME1).Range(retaintxboxstr).Value = .TextBox5.Text
  .Hide
End With

If ("真円精度" = Sheets(SHEETNAME1).Cells(5, 6).Value) Then
 CirGraph_SizeCalc
End If
'ActiveChart.Shapes.Item(5).Select
Application.ScreenUpdating = True
End Sub
