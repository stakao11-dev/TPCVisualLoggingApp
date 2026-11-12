Attribute VB_Name = "UserForm8"
Attribute VB_Base = "0{ABF90148-2EBC-4E8E-8908-C390E66465FF}{6C0A4A49-B874-455C-BDA7-984B53DA9446}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton1_Click()
    Label5.Caption = Now                '計算開始時間表示
    Application.ScreenUpdating = False  '画面再描画禁止
    UserForm8.Hide
    If Not IsNumeric(TextBox4.Value) Then
      Application.ScreenUpdating = True
      MsgBox "数値を入力してください"
      Exit Sub
    End If
    Dim amp As Long
    Dim Radius, centerX, centerY As Double
    
    amp = TextBox4.Text                     'コントロールから数値の読み出し
     
    With Worksheets(SHEETNAME1)          'シートの選択
        Radius = .Range(RPOINT).Value             '半径Rの取得
        centerX = .Range(XCENTER).Value            'X軸中心の取得
        centerY = .Range(YCENTER).Value            'Y軸中心の取得
        .Range(AMPPOINT).Value = Val(amp)           '拡大率の反映
    End With
    
    Call AnalyzeErr(Radius, amp)
    Call retain_data
    Application.ScreenUpdating = True   '画面描画有効
    Label6.Caption = Now                '計算終了時間表示
End Sub

Private Sub CommandButton2_Click()
 Dim csvfname As String
    Label5.Caption = Now  '計算開始時間表示
    UserForm8.Hide
    If Not (IsNumeric(TextBox1.Value) And IsNumeric(TextBox2.Value) And IsNumeric(TextBox3.Value) And IsNumeric(TextBox4.Value)) Then
      MsgBox "数値を入力してください"
      Exit Sub
    End If
    
    With Worksheets(SHEETNAME1)             'シートの選択
        .Range(RPOINT).Value = TextBox1.Text          '半径Rの反映
        .Range(XCENTER).Value = TextBox2.Text          'X軸中心の反映
        .Range(YCENTER).Value = TextBox3.Text          'Y軸中心の反映
        .Range(AMPPOINT).Value = TextBox4.Text          '拡大率の取得
        csvfname = .Range(FILENAMEPOINT).Value           '拡大率の取得
    End With

    Call circleckeck(csvfname)
    Call retain_data
    Label6.Caption = Now                '計算終了時間表示
End Sub

