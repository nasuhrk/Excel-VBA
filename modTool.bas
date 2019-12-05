Attribute VB_Name = "modTool"
Option Explicit

' ============================================================
'  [removeNameDefinition]
'  - 不要な名前の定義を削除します
' ============================================================
Function removeNameDefinition()
'
    'エラーを無視 (削除件数に計上しない)
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
    
    '結果表示
    If (0 < cnt) Then
        MsgBox cnt & " / " & total & " 件の定義を削除しました", vbInformation
    Else
        MsgBox "対象はありません", vbInformation
    End If

End Function

' ============================================================
'  [checkUnhideSheets]
'  - 非表示シートをチェックします
'  - 必要に応じて全てのシートを再表示します
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
        MsgBox "対象はありません", vbInformation
        Exit Function
    End If

    If vbNo = MsgBox("非表示は" & cnt & "シートです。全てのシートを再表示しますか？", vbYesNo + vbDefaultButton2) Then
        Exit Function
    End If

    cnt = 0
    
    For i = Sheets.Count To 1 Step -1
        If (Sheets(i).Visible = False) Then
            Sheets(i).Visible = True
            MsgBox Sheets(i).Name & " を再表示しました", vbInformation
        End If
    Next i

End Function

' ============================================================
'  [resetAllPageBreaks]
'  - 改ページを解除します
' ============================================================
Function resetAllPageBreaks()
'
    '全ての改ページ解除
    ActiveSheet.resetAllPageBreaks

    '結果表示
    MsgBox "完了しました", vbInformation

End Function

' ============================================================
'  [createSheetList]
'  - シート名の一覧を作成します
' ============================================================
Function createSheetList()
'
    Dim i As Integer
    
    Dim wb As Variant
    Set wb = ActiveWorkbook
    
    Workbooks.Add
    Sheets(1).Name = "シート名一覧"
    
    For i = 1 To wb.Worksheets.Count
        Cells(i, 1) = i                 '項番
        Cells(i, 2) = wb.Sheets(i).Name 'シート名
    Next
  
    Columns("A:B").EntireColumn.AutoFit
   
    'カーソルをホームポジションに移動
    Cells(1, 1).Select

End Function

' ============================================================
'  [createGraphPaper]
'  - 方眼紙を作成します
' ============================================================
Function createGraphPaper()
'
    'シートを追加
    Worksheets.Add

    Cells.Select
    
    '列の幅を 2 にする
    Selection.ColumnWidth = 2
    
    '表示形式は"文字列"
    Selection.NumberFormatLocal = "@"
    
    'カーソルをホームポジションに移動
    Cells(1, 1).Select

    '結果表示
    MsgBox ActiveSheet.Name & " を追加しました", vbInformation

End Function

' ============================================================
'  [removePhoneticCharacters]
'  - 選択範囲のふりがなをまとめて削除します
' ============================================================
Function removePhoneticCharacters()
'
    Dim cnt As Integer: cnt = 0
    Dim r As Range
    
    For Each r In Selection
        '空欄は対象外
        If r.Value <> "" Then
            If r.Characters.PhoneticCharacters <> "" Then
                r.Characters.PhoneticCharacters = ""
                cnt = cnt + 1
            End If
        End If
    Next r
    
    If 0 < cnt Then
        MsgBox cnt & " 件 のふりがなを削除しました", vbInformation
    Else
        Call popupMessage("対象はありません", vbInformation)
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
        
        pagedata(0) = Array(1, 1, 2, 1.5, 0.8, 0.8, xlLandscape, xlPaperA4, 1, False)  'Ａ４（横）
        pagedata(1) = Array(2, 0.5, 1.5, 1.5, 0.8, 0.8, xlPortrait, xlPaperA4, False, 1) 'Ａ４（縦）
        pagedata(2) = Array(2, 1, 1.5, 1.5, 0.8, 0.8, xlLandscape, xlPaperA3, 1, False) 'Ａ３（横）
    
        With ActiveSheet.PageSetup
            .LeftMargin = Application.CentimetersToPoints(pagedata(idx)(0))    'マージン(左)
            .RightMargin = Application.CentimetersToPoints(pagedata(idx)(1))   'マージン(右)
            .TopMargin = Application.CentimetersToPoints(pagedata(idx)(2))     'マージン(上)
            .BottomMargin = Application.CentimetersToPoints(pagedata(idx)(3))  'マージン(下)
            .HeaderMargin = Application.CentimetersToPoints(pagedata(idx)(4))  'マージン(ヘッダー)
            .FooterMargin = Application.CentimetersToPoints(pagedata(idx)(5))  'マージン(フッター)
            .Orientation = pagedata(idx)(6)    '印刷の向き
            .PaperSize = pagedata(idx)(7)      '用紙サイズ
            .Zoom = False          '---拡大・縮小率を指定しない
            .FitToPagesWide = pagedata(idx)(8) '横幅に合わせる
            .FitToPagesTall = pagedata(idx)(9) '縦幅に合わせる
        End With
        
        ActiveSheet.PrintPreview '印刷プレビューを表示
    Else
        Call popupMessage("ページ設定を選択してください", vbCritical)
    End If

End Function

' ============================================================
'  [popupMessage]
'  - メッセージをポップアップします。
' ============================================================
Function popupMessage(prompt As String, msgboxstyle As vbmsgboxstyle)
    MsgBox prompt, msgboxstyle, ""
End Function

