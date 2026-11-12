Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit
Private Sub workbook_beforeclose(cancel As Boolean)
Dim i As Integer
'Application.ScreenUpdating = False  'ほかのブックが開いているときに
'Sheets(SHEETNAME1).Select            'アプリケーションごと閉じようとすると
'ActiveSheet.Range("a3:b65536").Select 'エラーがでる。
'Selection.ClearContents
'ActiveSheet.Range("a3").Select
'Sheets(SHEETNAME2).Select
'ActiveSheet.Range("a3:m65536").Select
'Selection.ClearContents
'ActiveSheet.Range("a3").Select
'For i = Sheets.Count To 1 Step -1
'  If Not (Sheets(i).Name = SHEETNAME1 Or Sheets(i).Name = SHEETNAME2) Then
'  Sheets(i).Delete
'  End If
'Next i
'Sheets(SHEETNAME1).Activate
'Application.ScreenUpdating = True
End Sub
Private Sub Workbook_Open()
Dim sheetnum As Integer, chartnum As Integer
Dim csvfname As String, sysnum As Integer
Dim cnt As Integer, chkflg As Integer
Dim name_str As String, sysname As String
Dim mydate As Variant, mymonth As Variant, myyear As Variant

 Application.ScreenUpdating = False

 '日付制限
 mydate = Date
 mymonth = Month(mydate)
 myyear = Year(mydate)
 With Sheets(SHEETNAME1)
  'If (Not ((myyear = 2004) And ((1 = mymonth) Or (2 = mymonth)))) Then
  If (0) Then
   .CommandButton1.Enabled = False
   .CommandButton2.Enabled = False
   .CommandButton3.Enabled = False
   Application.ScreenUpdating = True
   MsgBox "試用期間が過ぎてしまいました。継続してご利用になる場合は正式版をご注文下さい。"
   Application.ScreenUpdating = False
   Data_Clear
  Else
   .Activate
   .Range(STARTCSVX).Select
   .CommandButton1.Enabled = True
   .CommandButton2.Enabled = False
   .CommandButton3.Enabled = False
   If (1 = sheet_chk()) Then
    .CommandButton2.Enabled = True
    .CommandButton3.Enabled = True
   End If
  End If
 End With
 Application.ScreenUpdating = True
End Sub

Sub old_sub()
Dim sheetnum As Integer, chartnum As Integer
Dim csvfname As String, sysnum As Integer
Application.ScreenUpdating = False
Sheets(SHEETNAME1).Activate
Sheets(SHEETNAME1).Range(STARTCSVX).Select
Sheets(SHEETNAME1).CommandButton2.Enabled = False
Sheets(SHEETNAME1).CommandButton3.Enabled = False
 
 sheetnum = Application.Worksheets.Count
 chartnum = Application.Charts.Count
 
 If ("サーボ応答精度" = ActiveSheet.Cells(5, 6).Value) Then
  If ((2 <= sheetnum) And (3 <= chartnum)) Then
   If ((CHARTNAME11 = ThisWorkbook.Charts.Item(1).Name) And _
       (CHARTNAME12 = ThisWorkbook.Charts.Item(2).Name) And _
       (CHARTNAME13 = ThisWorkbook.Charts.Item(3).Name) And _
       (SHEETNAME1 = ThisWorkbook.Sheets.Item(1).Name) And _
       (SHEETNAME2 = ThisWorkbook.Sheets.Item(2).Name)) Then
        Sheets(SHEETNAME1).CommandButton2.Enabled = True
        Sheets(SHEETNAME1).CommandButton3.Enabled = True
   End If
  End If
 ElseIf ("真円精度" = ActiveSheet.Cells(5, 6).Value) Then
  If ((2 <= sheetnum) And (1 <= chartnum)) Then
   If ((CHARTNAME21 = ThisWorkbook.Charts.Item(1).Name) And _
       (SHEETNAME1 = ThisWorkbook.Sheets.Item(1).Name) And _
       (SHEETNAME2 = ThisWorkbook.Sheets.Item(2).Name)) Then
        Sheets(SHEETNAME1).CommandButton2.Enabled = True
        Sheets(SHEETNAME1).CommandButton3.Enabled = True
   End If
  End If
 ElseIf ("直線精度" = ActiveSheet.Cells(5, 6).Value) Then
  If ((2 <= sheetnum) And (2 <= chartnum)) Then
   If ((CHARTNAME31 = ThisWorkbook.Charts.Item(1).Name) And _
       (CHARTNAME32 = ThisWorkbook.Charts.Item(2).Name) And _
       (SHEETNAME1 = ThisWorkbook.Sheets.Item(1).Name) And _
       (SHEETNAME2 = ThisWorkbook.Sheets.Item(2).Name)) Then
        Sheets(SHEETNAME1).CommandButton2.Enabled = True
        Sheets(SHEETNAME1).CommandButton3.Enabled = True
   End If
  End If
 End If
Application.ScreenUpdating = True
End Sub
