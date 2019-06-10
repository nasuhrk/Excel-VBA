Attribute VB_Name = "Personal"
Option Explicit

Public Sub A00_Cleanup()
    
    Dim hiddenFlg As Boolean
    Dim i As Integer
    
    Call A03_VisibleNames

    For i = Sheets.Count To 1 Step -1
        
        Application.StatusBar = "処理中...あと [ " & i & "/" & Sheets.Count & " ] ファイル"
        
        hiddenFlg = False
        If (Sheets(i).Visible = False) Then
            Dim ret As Integer
            ret = MsgBox("[" & Sheets(i).name & "] シートが非表示です。シートを削除しますか？", vbYesNo + vbDefaultButton1)
            If ret = vbYes Then
                Sheets(i).Delete
            Else
                Sheets(i).Visible = True
                hiddenFlg = True
            End If
        End If
    
        Sheets(i).Select
        Call A01_SheetCleanup
        Call A02_PageCleanup
    
        If (hiddenFlg = True) Then
            Sheets(i).Visible = False
        End If
    
    Next i
        
    Application.StatusBar = "完了しました"

End Sub


Private Sub A01_SheetCleanup()
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ActiveWindow.DisplayGridlines = False
    
    ActiveWindow.View = xlNormalView
    ActiveWindow.Zoom = 100
    ActiveWindow.View = xlPageBreakPreview
    ActiveWindow.Zoom = 100
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ActiveWindow.ScrollColumn = 1 ' スクロール列の設定
    ActiveWindow.ScrollRow = 1    ' スクロール行の設定
    
    'カーソルを左上に設定
    Cells(1, 1).Select
    
    'タブバーを規定サイズに設定
    ActiveWindow.TabRatio = 0.6

End Sub

Private Sub A02_PageCleanup()
    
    Dim AutoFlg As Boolean
    AutoFlg = True
    
    With ActiveSheet.PageSetup
        Application.PrintCommunication = False
        
        If AutoFlg = True Then
            .FitToPagesWide = 1
            .FitToPagesTall = 0
        Else
           .Zoom = 80
        End If
    
    Application.PrintCommunication = True
    End With
    
    '全ての改ページ解除
    ActiveSheet.ResetAllPageBreaks

End Sub

Private Sub A03_VisibleNames()
    Dim n As name
    
    For Each n In Names
        If n.Visible = False Then
            n.Visible = True
        End If
    Next
    
'    For Each n In ActiveWorkbook.Names
'        ' Print_Area を残す
'        If Not n.name Like "*!Print_Area" Then
'            n.Delete
'        End If
'
'        ' Print_Titles を残す
'        If Not n.name Like "*!Print_Titles" Then
'            n.Delete
'        End If
'    Next
    
End Sub

Public Sub B_設計書列幅_初期化()
    Cells.Select
    Selection.ColumnWidth = 3
    Columns("A:A").Select
    Selection.ColumnWidth = 1
    Range("A1").Select
End Sub

Public Sub C_方眼紙_初期化()
    Cells.Select
    Selection.ColumnWidth = 2
    Cells(1, 1).Select
End Sub
