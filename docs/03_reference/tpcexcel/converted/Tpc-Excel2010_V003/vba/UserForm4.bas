Attribute VB_Name = "UserForm4"
Attribute VB_Base = "0{27DEFBEF-5F65-481B-A63D-79035748182B}{0D5C9AB6-E19F-4970-8200-54851E17DC83}"
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
