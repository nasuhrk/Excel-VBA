Attribute VB_Name = "modMain"
Option Explicit

'目盛線の表示/非表示
Dim display_gridlines As Boolean

'アクティブ ウィンドウのビューモード
Dim active_window_view(3) As Integer

' ============================================================
'  [setシート初期値]
' ============================================================
Function setシート初期値()
'
    '目盛線の表示/非表示
    display_gridlines = True '表示
    
    'アクティブ ウィンドウのビューモード
    active_window_view(0) = xlNormalView
    
    '表示倍率 (標準)
    active_window_view(xlNormalView) = 100
    
    '表示倍率 (改ページ プレビュー)
    active_window_view(xlPageBreakPreview) = 60
    
    '表示倍率 (ページ レイアウト ビュー)
    active_window_view(xlPageLayoutView) = 100

End Function

' ============================================================
'  [showフォーム]
' ============================================================
Public Sub showフォーム()
'
    'フォームを表示
    frmMain.Show vbModeless

End Sub

' ============================================================
'  [getシート設定値]
' ============================================================
Function getシート設定値()
'
    '目盛線の表示/非表示
    display_gridlines = ActiveWindow.DisplayGridlines
        
    'アクティブ ウィンドウのビューモード
    active_window_view(0) = ActiveWindow.View
    
    '表示倍率 (改ページ プレビュー)
    ActiveWindow.View = xlPageBreakPreview
    active_window_view(xlPageBreakPreview) = ActiveWindow.zoom
    
    '表示倍率 (標準)
    ActiveWindow.View = xlNormalView
    active_window_view(xlNormalView) = ActiveWindow.zoom

    '//表示倍率 (ページ レイアウト ビュー)
    '//ActiveWindow.View = xlPageLayoutView
    '//active_window_view(xlPageLayoutView) = ActiveWindow.Zoom

    'ビューモードを復元
    ActiveWindow.View = active_window_view(0)

End Function

' ============================================================
'  [getフォーム値]
' ============================================================
Function getフォーム値() As Boolean
'
    getフォーム値 = False
    
    display_gridlines = frmMain.chkGridlines.Value
    
    If frmMain.optWindowView1.Value Then
        active_window_view(0) = xlNormalView
    Else
        active_window_view(0) = xlPageBreakPreview
    End If

    Dim view1, view2 As Variant
    
    view1 = frmMain.txtWindoView1.Text
    view2 = frmMain.txtWindoView2.Text
    
    If Not (IsNumeric(view1) And IsNumeric(view2)) Then
        '入力値エラー
        MsgBox "入力値を正しく入力してください", vbExclamation
        getフォーム値 = True
        
        If Not IsNumeric(view1) Then
            frmMain.txtWindoView1.SetFocus
            Exit Function
        End If
        
        If Not IsNumeric(view2) Then
            frmMain.txtWindoView2.SetFocus
            Exit Function
        End If
    End If

    active_window_view(xlNormalView) = view1
    active_window_view(xlPageBreakPreview) = view2

End Function

' ============================================================
'  [setフォーム値]
' ============================================================
Function setフォーム値()
'
    If active_window_view(0) = xlNormalView Then
        frmMain.optWindowView1.Value = True
    Else
        frmMain.optWindowView2.Value = True
    End If

    frmMain.txtWindoView1.Text = active_window_view(xlNormalView)
    frmMain.txtWindoView2.Text = active_window_view(xlPageBreakPreview)
    frmMain.chkGridlines.Value = display_gridlines

End Function

' ============================================================
'  [actionシート仕上げ]
' ============================================================
Function actionシート仕上げ()
'
    Dim hiddenFlg As Boolean
    Dim i As Integer
    
    '処理の高速化(オン)
    Call screenUpdating(True)
        
    'タブバーを規定サイズに設定
    ActiveWindow.TabRatio = 0.6
    
    For i = Sheets.Count To 1 Step -1
        
        '非表示シートを一時的に表示
        hiddenFlg = False
        If (Sheets(i).Visible = False) Then
            Sheets(i).Visible = True
            hiddenFlg = True
        End If
    
        '対象のシート選択
        Sheets(i).Select
        
        Call sheetCleanup
    
        '一時的に表示したシートを元に戻す
        If (hiddenFlg = True) Then
            Sheets(i).Visible = False
        End If
    
    Next i
        
    '処理の高速化(オフ)
    Call screenUpdating(True)
    
    '結果表示
    MsgBox "完了しました", vbInformation
    
End Function

' ============================================================
'  [sheetCleanup]
' ============================================================
Private Sub sheetCleanup()
'
    'ページ レイアウト ビュー
    ActiveWindow.View = xlPageLayoutView
    ActiveWindow.zoom = active_window_view(xlPageLayoutView)
    
    If active_window_view(0) = xlNormalView Then
        '改ページ プレビュー
        ActiveWindow.View = xlPageBreakPreview
        ActiveWindow.zoom = active_window_view(xlPageBreakPreview)
        '標準
        ActiveWindow.View = xlNormalView
        ActiveWindow.zoom = active_window_view(xlNormalView)
    Else
        '標準
        ActiveWindow.View = xlNormalView
        ActiveWindow.zoom = active_window_view(xlNormalView)
        '改ページ プレビュー
        ActiveWindow.View = xlPageBreakPreview
        ActiveWindow.zoom = active_window_view(xlPageBreakPreview)
    End If

    '目盛線(枠線)を非表示
    ActiveWindow.DisplayGridlines = display_gridlines
    
    'スクロールバーを初期位置に設定
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    'カーソルを左上に設定
    Cells(1, 1).Select

End Sub

' ============================================================
'  [screenUpdating]
' ============================================================
Private Sub screenUpdating(ByVal mode As Boolean)
'
    If mode Then
        '処理の高速化(オン)
        With Application
            .screenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
        End With
    Else
        '処理の高速化(オフ)
        With Application
            .screenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
        End With
    End If

End Sub
