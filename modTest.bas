Attribute VB_Name = "modTest"
Option Explicit

Function delete_name_and_style()

	Dimi As Integer

	Dim wb As Variant
	Set wb = ActiveWorkbook
	
	'書式（スタイル）定義
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
				flg = True '一致したら抜けて次を検証
				i = i + 1 '次の対象
				Exit For
			End If
		Next

		'一致しない場合は削除
		If flg = False Then
			wb.styles(i).Delete
			cnt = cnt + 1 'カウント
		End If
	Wend

	If 0 < cnt Then
		MsgBox cnt & "件の不要なスタイルを削除しました", vbInformation
	Else
		Call popupMessage("対象はありません", vbInformation)
	End If

End Function

Sub get_name_and_style()

	Dim i As Integer

	Dim wb As Variant
	Set wb = ActiveWorkbook

	Workbooks.Add
	Sheets (1).Name = "スタイル一覧"

	For i = 1 To wb.styles.Count
		Cells(i, 1) = i						'項番
		Cells(i, 2) = wb.styles(i). Name	'スタイル名
	Next

	Columns ("A:B").Ent i reColumn.AutoFit 

	'カーソルをホームポジションに移動
	Cells(1, 1).Select
End Sub
