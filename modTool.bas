Attribute VB_Name = "modTool"
Option Explicit

' ============================================================
'  [removeNameDefinition]
'  - �s�v�Ȗ��O�̒�`���폜���܂�
' ============================================================
Function removeNameDefinition()
'
    '�G���[�𖳎� (�폜�����Ɍv�サ�Ȃ�)
    On Error Resume Next
    
    Dim total As Integer: total = 0
    Dim cnt As Integer: cnt = 0
    Dim n As Name
    
    For Each n In ActiveWorkbook.Names
        If n.Visible = False Then
            n.Visible = True
        End If
        
        If InStr(n.Value, "#REF") > 0 Or InStr(n.Value, "\") > 0 Then
            '[DEBUG] MsgBox "Name=" & n.name & " Value=" & n.Value
            n.Delete
            cnt = cnt + 1
        End If
        
        total = total + 1
    Next n
    
    '���ʕ\��
    If (0 < cnt) Then
        MsgBox cnt & " / " & total & " ���̒�`���폜���܂���", vbInformation
    Else
        MsgBox "�Ώۂ͂���܂���", vbInformation
    End If

End Function

' ============================================================
'  [checkUnhideSheets]
'  - ��\���V�[�g���`�F�b�N���܂�
'  - �K�v�ɉ����đS�ẴV�[�g���ĕ\�����܂�
' ============================================================
Function checkUnhideSheets()
'
    Dim i As Integer
    Dim cnt As Integer
    
    cnt = 0
    
    For i = Sheets.Count To 1 Step -1
        If (Sheets(i).Visible = False) Then
            cnt = cnt + 1
        End If
    Next i
    
    If cnt = 0 Then
        MsgBox "�Ώۂ͂���܂���", vbInformation
        Exit Function
    End If

    If vbNo = MsgBox("��\����" & cnt & "�V�[�g�ł��B�S�ẴV�[�g���ĕ\�����܂����H", vbYesNo + vbDefaultButton2) Then
        Exit Function
    End If

    cnt = 0
    
    For i = Sheets.Count To 1 Step -1
        If (Sheets(i).Visible = False) Then
            Sheets(i).Visible = True
            MsgBox Sheets(i).Name & " ���ĕ\�����܂���", vbInformation
        End If
    Next i

End Function

' ============================================================
'  [resetAllPageBreaks]
'  - ���y�[�W���������܂�
' ============================================================
Function resetAllPageBreaks()
'
    '�S�Ẳ��y�[�W����
    ActiveSheet.resetAllPageBreaks

    '���ʕ\��
    MsgBox "�������܂���", vbInformation

End Function

' ============================================================
'  [createSheetList]
'  - �V�[�g���̈ꗗ���쐬���܂�
' ============================================================
Function createSheetList()
'
    Dim i As Integer
    
    Dim wb As Variant
    Set wb = ActiveWorkbook
    
    Workbooks.Add
    Sheets(1).Name = "�V�[�g���ꗗ"
    
    For i = 1 To wb.Worksheets.Count
        Cells(i, 1) = i                 '����
        Cells(i, 2) = wb.Sheets(i).Name '�V�[�g��
    Next
  
    Columns("A:B").EntireColumn.AutoFit
   
    '�J�[�\�����z�[���|�W�V�����Ɉړ�
    Cells(1, 1).Select

End Function

' ============================================================
'  [createGraphPaper]
'  - ���ᎆ���쐬���܂�
' ============================================================
Function createGraphPaper()
'
    '�V�[�g��ǉ�
    Worksheets.Add

    Cells.Select
    
    '��̕��� 2 �ɂ���
    Selection.ColumnWidth = 2
    
    '�\���`����"������"
    Selection.NumberFormatLocal = "@"
    
    '�J�[�\�����z�[���|�W�V�����Ɉړ�
    Cells(1, 1).Select

    '���ʕ\��
    MsgBox ActiveSheet.Name & " ��ǉ����܂���", vbInformation

End Function

' ============================================================
'  [removePhoneticCharacters]
'  - �I��͈͂̂ӂ肪�Ȃ��܂Ƃ߂č폜���܂�
' ============================================================
Function removePhoneticCharacters()
'
    Dim cnt As Integer: cnt = 0
    Dim r As Range
    
    For Each r In Selection
        '�󗓂͑ΏۊO
        If r.Value <> "" Then
            If r.Characters.PhoneticCharacters <> "" Then
                r.Characters.PhoneticCharacters = ""
                cnt = cnt + 1
            End If
        End If
    Next r
    
    If 0 < cnt Then
        MsgBox cnt & " �� �̂ӂ肪�Ȃ��폜���܂���", vbInformation
    Else
        Call popupMessage("�Ώۂ͂���܂���", vbInformation)
    End If

End Function

' ============================================================
'  [setPageStyle]
'  -
' ============================================================
Function setPageStyle(idx As String)
'
    'MsgBox (idx)
    If (0 <= idx And idx <= 2) Then
        Dim pagedata(2) As Variant
        
        pagedata(0) = Array(1, 1, 2, 1.5, 0.8, 0.8, xlLandscape, xlPaperA4, 1, False)  '�`�S�i���j
        pagedata(1) = Array(2, 0.5, 1.5, 1.5, 0.8, 0.8, xlPortrait, xlPaperA4, False, 1) '�`�S�i�c�j
        pagedata(2) = Array(2, 1, 1.5, 1.5, 0.8, 0.8, xlLandscape, xlPaperA3, 1, False) '�`�R�i���j
    
        With ActiveSheet.PageSetup
            .LeftMargin = Application.CentimetersToPoints(pagedata(idx)(0))    '�}�[�W��(��)
            .RightMargin = Application.CentimetersToPoints(pagedata(idx)(1))   '�}�[�W��(�E)
            .TopMargin = Application.CentimetersToPoints(pagedata(idx)(2))     '�}�[�W��(��)
            .BottomMargin = Application.CentimetersToPoints(pagedata(idx)(3))  '�}�[�W��(��)
            .HeaderMargin = Application.CentimetersToPoints(pagedata(idx)(4))  '�}�[�W��(�w�b�_�[)
            .FooterMargin = Application.CentimetersToPoints(pagedata(idx)(5))  '�}�[�W��(�t�b�^�[)
            .Orientation = pagedata(idx)(6)    '����̌���
            .PaperSize = pagedata(idx)(7)      '�p���T�C�Y
            .Zoom = False          '---�g��E�k�������w�肵�Ȃ�
            .FitToPagesWide = pagedata(idx)(8) '�����ɍ��킹��
            .FitToPagesTall = pagedata(idx)(9) '�c���ɍ��킹��
        End With
        
        ActiveSheet.PrintPreview '����v���r���[��\��
    Else
        Call popupMessage("�y�[�W�ݒ��I�����Ă�������", vbCritical)
    End If

End Function

' ============================================================
'  [popupMessage]
'  - ���b�Z�[�W���|�b�v�A�b�v���܂��B
' ============================================================
Function popupMessage(prompt As String, msgboxstyle As vbmsgboxstyle)
    MsgBox prompt, msgboxstyle, ""
End Function

