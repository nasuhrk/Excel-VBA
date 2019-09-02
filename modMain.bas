Attribute VB_Name = "modMain"
Option Explicit

'�ڐ����̕\��/��\��
Dim display_gridlines As Boolean

'�A�N�e�B�u �E�B���h�E�̃r���[���[�h
Dim active_window_view(3) As Integer

' ============================================================
'  [set�V�[�g�����l]
' ============================================================
Function set�V�[�g�����l()
'
    '�ڐ����̕\��/��\��
    display_gridlines = True '�\��
    
    '�A�N�e�B�u �E�B���h�E�̃r���[���[�h
    active_window_view(0) = xlNormalView
    
    '�\���{�� (�W��)
    active_window_view(xlNormalView) = 100
    
    '�\���{�� (���y�[�W �v���r���[)
    active_window_view(xlPageBreakPreview) = 60
    
    '�\���{�� (�y�[�W ���C�A�E�g �r���[)
    active_window_view(xlPageLayoutView) = 100

End Function

' ============================================================
'  [show�t�H�[��]
' ============================================================
Public Sub show�t�H�[��()
'
    '�t�H�[����\��
    frmMain.Show vbModeless

End Sub

' ============================================================
'  [get�V�[�g�ݒ�l]
' ============================================================
Function get�V�[�g�ݒ�l()
'
    '�ڐ����̕\��/��\��
    display_gridlines = ActiveWindow.DisplayGridlines
        
    '�A�N�e�B�u �E�B���h�E�̃r���[���[�h
    active_window_view(0) = ActiveWindow.View
    
    '�\���{�� (���y�[�W �v���r���[)
    ActiveWindow.View = xlPageBreakPreview
    active_window_view(xlPageBreakPreview) = ActiveWindow.zoom
    
    '�\���{�� (�W��)
    ActiveWindow.View = xlNormalView
    active_window_view(xlNormalView) = ActiveWindow.zoom

    '//�\���{�� (�y�[�W ���C�A�E�g �r���[)
    '//ActiveWindow.View = xlPageLayoutView
    '//active_window_view(xlPageLayoutView) = ActiveWindow.Zoom

    '�r���[���[�h�𕜌�
    ActiveWindow.View = active_window_view(0)

End Function

' ============================================================
'  [get�t�H�[���l]
' ============================================================
Function get�t�H�[���l() As Boolean
'
    get�t�H�[���l = False
    
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
        '���͒l�G���[
        MsgBox "���͒l�𐳂������͂��Ă�������", vbExclamation
        get�t�H�[���l = True
        
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
'  [set�t�H�[���l]
' ============================================================
Function set�t�H�[���l()
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
'  [action�V�[�g�d�グ]
' ============================================================
Function action�V�[�g�d�グ()
'
    Dim hiddenFlg As Boolean
    Dim i As Integer
    
    '�����̍�����(�I��)
    Call screenUpdating(True)
        
    '�^�u�o�[���K��T�C�Y�ɐݒ�
    ActiveWindow.TabRatio = 0.6
    
    For i = Sheets.Count To 1 Step -1
        
        '��\���V�[�g���ꎞ�I�ɕ\��
        hiddenFlg = False
        If (Sheets(i).Visible = False) Then
            Sheets(i).Visible = True
            hiddenFlg = True
        End If
    
        '�Ώۂ̃V�[�g�I��
        Sheets(i).Select
        
        Call sheetCleanup
    
        '�ꎞ�I�ɕ\�������V�[�g�����ɖ߂�
        If (hiddenFlg = True) Then
            Sheets(i).Visible = False
        End If
    
    Next i
        
    '�����̍�����(�I�t)
    Call screenUpdating(True)
    
    '���ʕ\��
    MsgBox "�������܂���", vbInformation
    
End Function

' ============================================================
'  [sheetCleanup]
' ============================================================
Private Sub sheetCleanup()
'
    '�y�[�W ���C�A�E�g �r���[
    ActiveWindow.View = xlPageLayoutView
    ActiveWindow.zoom = active_window_view(xlPageLayoutView)
    
    If active_window_view(0) = xlNormalView Then
        '���y�[�W �v���r���[
        ActiveWindow.View = xlPageBreakPreview
        ActiveWindow.zoom = active_window_view(xlPageBreakPreview)
        '�W��
        ActiveWindow.View = xlNormalView
        ActiveWindow.zoom = active_window_view(xlNormalView)
    Else
        '�W��
        ActiveWindow.View = xlNormalView
        ActiveWindow.zoom = active_window_view(xlNormalView)
        '���y�[�W �v���r���[
        ActiveWindow.View = xlPageBreakPreview
        ActiveWindow.zoom = active_window_view(xlPageBreakPreview)
    End If

    '�ڐ���(�g��)���\��
    ActiveWindow.DisplayGridlines = display_gridlines
    
    '�X�N���[���o�[�������ʒu�ɐݒ�
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    
    '�J�[�\��������ɐݒ�
    Cells(1, 1).Select

End Sub

' ============================================================
'  [screenUpdating]
' ============================================================
Private Sub screenUpdating(ByVal mode As Boolean)
'
    If mode Then
        '�����̍�����(�I��)
        With Application
            .screenUpdating = False
            .EnableEvents = False
            .Calculation = xlCalculationManual
        End With
    Else
        '�����̍�����(�I�t)
        With Application
            .screenUpdating = True
            .EnableEvents = True
            .Calculation = xlCalculationAutomatic
        End With
    End If

End Sub
