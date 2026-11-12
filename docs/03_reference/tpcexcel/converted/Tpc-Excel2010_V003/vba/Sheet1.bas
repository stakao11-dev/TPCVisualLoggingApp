Attribute VB_Name = "Sheet1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "CommandButton3, 25, 0, MSForms, CommandButton"
Attribute VB_Control = "CommandButton2, 24, 1, MSForms, CommandButton"
Attribute VB_Control = "CommandButton1, 23, 2, MSForms, CommandButton"
Private Sub CommandButton1_Click()
 Call retain_data
 Call chg_sys
 Call load_retain_data
End Sub
Private Sub CommandButton2_Click()
    Dim fname As String
    Application.ScreenUpdating = False  '
'    fname = LoadCsvFile(SHEETNAME1, "A3:B4098", "ﾛｷﾞﾝｸﾞﾃﾞｰﾀﾌｧｲﾙを開く", 2)
    fname = LoadCsvFile(SHEETNAME1, "A3:B" & ENDADD1, "ﾛｷﾞﾝｸﾞﾃﾞｰﾀﾌｧｲﾙを開く", 2)
    If fname <> "" Then
        Worksheets(SHEETNAME1).Range("G4").Value = fname
    End If
    Application.ScreenUpdating = True  '
End Sub
Private Sub CommandButton3_Click()
 If ((IsEmpty(ActiveSheet.Cells(3, 1).Value)) Or _
    (IsEmpty(ActiveSheet.Cells(3, 2).Value))) Then
  MsgBox "データが未設定の可能性があります。"
 Else
  Call retain_data
  If (1 = sheet_chk()) Then
   Call draw_graph
  Else
   'MsgBox "シートの不足があります。解析方法を再度選択してください。"
   Application.DisplayAlerts = False
    For i = Sheets.Count To 1 Step -1
     If Not (Sheets(i).Name = SHEETNAME1 Or Sheets(i).Name = SHEETNAME2) Then
      Sheets(i).Delete
     End If
    Next i
   Application.DisplayAlerts = True
  End If
 End If
End Sub

Private Sub CommandButton4_Click()
 Application.ScreenUpdating = False
 Data_Clear
 Application.ScreenUpdating = True
End Sub
