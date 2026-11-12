Attribute VB_Name = "UserForm4"
Attribute VB_Base = "0{DD77EB25-EFFD-461D-BEE4-BDDF7598FC54}{71C9C5EE-71AA-4568-AFF3-A3C2E2D5E14B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton1_Click()    '再計算の反映ボタンで動作
Call 再計算
End Sub

Private Sub TextBox1_AfterUpdate()    'フォーカスが移り変わったときに値が
Call change1                          'よい値であるかを調べる
End Sub
Private Sub TextBox2_AfterUpdate()
Call change2
End Sub
Private Sub TextBox3_AfterUpdate()
Call change3
End Sub
Private Sub TextBox4_AfterUpdate()
Call change4
End Sub
Private Sub TextBox5_AfterUpdate()
Call change5
End Sub
Private Sub TextBox6_AfterUpdate()
Call change6
End Sub
Private Sub TextBox7_AfterUpdate()
Call change7
End Sub
