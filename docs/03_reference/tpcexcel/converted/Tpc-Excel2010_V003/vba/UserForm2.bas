Attribute VB_Name = "UserForm2"
Attribute VB_Base = "0{8272DD97-5105-409E-A554-601A29A71AFA}{667635BA-15E0-4F6E-AEC5-EB9CCB895145}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub CheckBox1_Click()      'チェックボックスにチェックをしたときに
If UserForm2.CheckBox1.Value Then  'テキストボックスに書き込めなくする。
  UserForm2.TextBox6.Enabled = False  '「軸別」の右ボタンです。
Else
  UserForm2.TextBox6.Enabled = True
End If
End Sub
Private Sub CheckBox3_Click()
If UserForm2.CheckBox3.Value Then
  UserForm2.TextBox7.Enabled = False
Else
  UserForm2.TextBox7.Enabled = True
End If
End Sub
Private Sub CheckBox2_Click()
If UserForm2.CheckBox2.Value Then
  UserForm2.TextBox8.Enabled = False
Else
  UserForm2.TextBox8.Enabled = True
End If
End Sub
Private Sub CheckBox4_Click()
If UserForm2.CheckBox4.Value Then
  UserForm2.TextBox9.Enabled = False
Else
  UserForm2.TextBox9.Enabled = True
End If
End Sub
Private Sub checkbox7_click()
If UserForm2.CheckBox7.Value Then
  UserForm2.TextBox10.Enabled = False
Else
  UserForm2.TextBox10.Enabled = True
End If
End Sub
Private Sub checkbox8_click()
If UserForm2.CheckBox8.Value Then
  UserForm2.TextBox11.Enabled = False
Else
  UserForm2.TextBox11.Enabled = True
End If
End Sub
Private Sub checkbox9_click()
If UserForm2.CheckBox9.Value Then
  UserForm2.TextBox12.Enabled = False
Else
  UserForm2.TextBox12.Enabled = True
End If
End Sub
Private Sub CheckBox10_Click()
If UserForm2.CheckBox10.Value Then
  UserForm2.TextBox13.Enabled = False
Else
  UserForm2.TextBox13.Enabled = True
End If
End Sub
Private Sub CheckBox5_Click()    '縦軸の目盛り線の追加。
ActiveSheet.ChartObjects("グラフ 1").Activate
If UserForm2.CheckBox5.Value Then
  ActiveChart.Axes(xlValue).HasMajorGridlines = True
  With ActiveChart.Axes(xlValue).MajorGridlines.Border
    .Weight = xlHairline
    .LineStyle = xlDot
  End With
Else
  ActiveChart.Axes(xlValue).HasMajorGridlines = False
End If
End Sub
Private Sub CheckBox11_Click()
ActiveSheet.ChartObjects("グラフ 2").Activate
If UserForm2.CheckBox11.Value Then
  ActiveChart.Axes(xlValue).HasMajorGridlines = True
  With ActiveChart.Axes(xlValue).MajorGridlines.Border
    .Weight = xlHairline
    .LineStyle = xlDot
  End With
Else
  ActiveChart.Axes(xlValue).HasMajorGridlines = False
End If
End Sub
Private Sub CheckBox12_Click()
ActiveSheet.ChartObjects("グラフ 3").Activate
If UserForm2.CheckBox12.Value Then
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
With UserForm2
  If Not (IsNumeric(.TextBox6.Value) And IsNumeric(.TextBox7.Value) And IsNumeric(.TextBox8.Value) And IsNumeric(.TextBox9.Value)) And IsNumeric(.TextBox10.Value) And IsNumeric(.TextBox11.Value) And IsNumeric(.TextBox12.Value) And IsNumeric(.TextBox13.Value) Then
    MsgBox "数値を入力してください"
    Exit Sub
  End If
End With

Application.ScreenUpdating = False
ActiveSheet.ChartObjects("グラフ 1").Activate
If UserForm2.CheckBox1.Value Then        'チェックボックスが付いていればオート
  ActiveChart.Axes(xlCategory).MinimumScaleIsAuto = True
Else                                     'そうでないならテキストボックス参照
  ActiveChart.Axes(xlCategory).MinimumScale = UserForm2.TextBox6.Value
End If

If UserForm2.CheckBox3.Value Then
  ActiveChart.Axes(xlCategory).MaximumScaleIsAuto = True
Else
  ActiveChart.Axes(xlCategory).MaximumScale = UserForm2.TextBox7.Value
End If

If UserForm2.CheckBox2.Value Then
  ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
Else                                      'xlcategoryは横軸xlvalueは縦軸
  ActiveChart.Axes(xlValue).MinimumScale = UserForm2.TextBox8.Value
End If

If UserForm2.CheckBox4.Value Then
  ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
Else
  ActiveChart.Axes(xlValue).MaximumScale = UserForm2.TextBox9.Value
End If

ActiveChart.PlotArea.width = 650    'グラフの位置調節。もっともあまりずれなくはなった。
ActiveChart.PlotArea.Left = 25

ActiveSheet.ChartObjects("グラフ 2").Activate
If UserForm2.CheckBox1.Value Then
  ActiveChart.Axes(xlCategory).MinimumScaleIsAuto = True
Else
  ActiveChart.Axes(xlCategory).MinimumScale = UserForm2.TextBox6.Value
End If

If UserForm2.CheckBox3.Value Then
  ActiveChart.Axes(xlCategory).MaximumScaleIsAuto = True
Else
  ActiveChart.Axes(xlCategory).MaximumScale = UserForm2.TextBox7.Value
End If
If UserForm2.CheckBox7 Then
  ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
Else
  ActiveChart.Axes(xlValue).MinimumScale = UserForm2.TextBox10.Value
End If

If UserForm2.CheckBox8 Then
  ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
Else
  ActiveChart.Axes(xlValue).MaximumScale = UserForm2.TextBox11.Value
End If

ActiveChart.PlotArea.width = 650
ActiveChart.PlotArea.Left = 25

ActiveSheet.ChartObjects("グラフ 3").Activate
If UserForm2.CheckBox1.Value Then
  ActiveChart.Axes(xlCategory).MinimumScaleIsAuto = True
Else
  ActiveChart.Axes(xlCategory).MinimumScale = UserForm2.TextBox6.Value
End If

If UserForm2.CheckBox3.Value Then
  ActiveChart.Axes(xlCategory).MaximumScaleIsAuto = True
Else
  ActiveChart.Axes(xlCategory).MaximumScale = UserForm2.TextBox7.Value
End If
If UserForm2.CheckBox9 Then
  ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
Else
  ActiveChart.Axes(xlValue).MinimumScale = UserForm2.TextBox12.Value
End If

If UserForm2.CheckBox10 Then
  ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True
Else
  ActiveChart.Axes(xlValue).MaximumScale = UserForm2.TextBox13.Value
End If

ActiveChart.PlotArea.width = 650
ActiveChart.PlotArea.Left = 25

UserForm2.Hide
Application.ScreenUpdating = True
End Sub

