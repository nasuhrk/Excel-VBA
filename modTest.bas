Attribute VB_Name = "modTest"
Option Explicit

Function delete_name_and_style()

	Dimi As Integer

	Dim wb As Variant
	Set wb = ActiveWorkbook
	
	'�����i�X�^�C���j��`
	Dim styles As Variant
	styles = Array(
		"Normal", _
		"")

	Dim s As Variant
	Dim flg As Boolean
	Dim cnt As Byte

	cnt = 0
	i = 1
	While i < wb.styles.Count
		flg = False
		For Each s In styles
			If s = wb.styles(i).Name Then
				flg = True '��v�����甲���Ď�������
				i = i + 1 '���̑Ώ�
				Exit For
			End If
		Next

		'��v���Ȃ��ꍇ�͍폜
		If flg = False Then
			wb.styles(i).Delete
			cnt = cnt + 1 '�J�E���g
		End If
	Wend

	If 0 < cnt Then
		MsgBox cnt & "���̕s�v�ȃX�^�C�����폜���܂���", vbInformation
	Else
		Call popupMessage("�Ώۂ͂���܂���", vbInformation)
	End If

End Function

Sub get_name_and_style()

	Dim i As Integer

	Dim wb As Variant
	Set wb = ActiveWorkbook

	Workbooks.Add
	Sheets (1).Name = "�X�^�C���ꗗ"

	For i = 1 To wb.styles.Count
		Cells(i, 1) = i						'����
		Cells(i, 2) = wb.styles(i). Name	'�X�^�C����
	Next

	Columns ("A:B").Ent i reColumn.AutoFit 

	'�J�[�\�����z�[���|�W�V�����Ɉړ�
	Cells(1, 1).Select
End Sub
