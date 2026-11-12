Attribute VB_Name = "UserForm6"
Attribute VB_Base = "0{16097412-D321-4486-93B3-7E426B9E0BB7}{6DE4A619-1D6B-4890-8A3B-8876F27ABD45}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton2_Click()
Application.ScreenUpdating = False  '画面再描画禁止
Dim rtctim As Integer
UserForm6.Hide
If Not IsNumeric(UserForm6.TextBox1.Value) Then
  Application.ScreenUpdating = False
  MsgBox "数値を入力してください"
  Exit Sub
End If
 Worksheets(SHEETNAME1).Range(RTCPOINT).Value = TextBox1.Text
 rtctim = TextBox1.Text
 Call speedcalc(Worksheets(SHEETNAME1).Range(FILENAMEPOINT).Value, rtctim)
 Call retain_data
Application.ScreenUpdating = True   '画面描画有効
End Sub

