Attribute VB_Name = "modMain"
Option Explicit

'�ڐ����̕\��/��\��
Dim display_gridlines As Boolean

'�A�N�e�B�u �E�B���h�E�̃r���[���[�h
Dim active_window_view(3) As Integer

' ============================================================
'  [ExcelFinisher]
' ============================================================
Public Sub ExcelFinisher_START()
'
    '�t�H�[����\��
    frmMain.Show vbModeless

End Sub

' ============================================================
'  [setSheetValue]
' ============================================================
Function setSheetValue()
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
'  [getSheetValue]
' ============================================================
Function getSheetValue()
'
    '�ڐ����̕\��/��\��
    display_gridlines = ActiveWindow.DisplayGridlines
        
    '�A�N�e�B�u �E�B���h�E�̃r���[���[�h
    active_window_view(0) = ActiveWindow.View
    
    '�\���{�� (�W��)
    If ActiveWindow.View = xlNormalView Then
        active_window_view(xlNormalView) = ActiveWindow.Zoom
    End If
    
    '�\���{�� (���y�[�W �v���r���[)
    If ActiveWindow.View = xlPageBreakPreview Then
        active_window_view(xlPageBreakPreview) = ActiveWindow.Zoom
    End If
    
    '�\���{�� (�y�[�W ���C�A�E�g �r���[)
    If ActiveWindow.View = xlPageLayoutView Then
        active_window_view(xlPageLayoutView) = ActiveWindow.Zoom
    End If
    
    '���ʕ\��
    MsgBox "�l���擾���܂���", vbInformation

End Function

' ============================================================
'  [setFormValue]
' ============================================================
Function setFormValue()
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
'  [getFormValue]
' ============================================================
Function getFormValue() As Boolean
'
    getFormValue = False
    
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
        getFormValue = True
        
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
'  [actionFinisher]
' ============================================================
Function actionFinisher()
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
    Call screenUpdating(False)
    
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
    ActiveWindow.Zoom = active_window_view(xlPageLayoutView)
    
    If active_window_view(0) = xlNormalView Then
        '���y�[�W �v���r���[
        ActiveWindow.View = xlPageBreakPreview
        ActiveWindow.Zoom = active_window_view(xlPageBreakPreview)
        '�W��
        ActiveWindow.View = xlNormalView
        ActiveWindow.Zoom = active_window_view(xlNormalView)
    Else
        '�W��
        ActiveWindow.View = xlNormalView
        ActiveWindow.Zoom = active_window_view(xlNormalView)
        '���y�[�W �v���r���[
        ActiveWindow.View = xlPageBreakPreview
        ActiveWindow.Zoom = active_window_view(xlPageBreakPreview)
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
