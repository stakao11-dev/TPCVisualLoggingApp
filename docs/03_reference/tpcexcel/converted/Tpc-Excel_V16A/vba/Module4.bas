Attribute VB_Name = "Module4"
Sub Macro2()
Attribute Macro2.VB_Description = "マクロ記録日 : 2011/9/22  ユーザー名 : ASM5201"
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
' マクロ記録日 : 2011/9/22  ユーザー名 : ASM5201
'

'
    ActiveChart.SetSourceData Source:=Sheets("ﾛｷﾞﾝｸﾞﾃﾞｰﾀ").Range("A3:B31999"), _
        PlotBy:=xlColumns
End Sub
Sub Macro3()
Attribute Macro3.VB_Description = "マクロ記録日 : 2011/9/22  ユーザー名 : ASM5201"
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
' マクロ記録日 : 2011/9/22  ユーザー名 : ASM5201
'

'
    ActiveChart.SeriesCollection(1).XValues = "=ﾛｷﾞﾝｸﾞﾃﾞｰﾀ!R3C1:R31999C1"
End Sub
