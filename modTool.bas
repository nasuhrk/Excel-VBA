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
        
        If InStr(n.Value, "#REF!") > 0 Or InStr(n.Value, ":\") > 0 Or InStr(n.Value, "'\\") > 0 Then
            '[DEBUG]
            Debug.Print "x 名前：" & n.Name '名前を取得
            Debug.Print "x 参照先：" & n.RefersTo '参照先を取得
            Debug.Print "x 親要素：" & n.Parent.Name '親要素を取得
            Debug.Print "x 値：" & n.Value '親要素を取得
            n.Delete
            cnt = cnt + 1
        Else
            '[DEBUG]
            'Debug.Print "  Name=" & n.Name; " Value=" & n.Value
            Debug.Print "名前：" & n.Name '名前を取得
            Debug.Print "参照先：" & n.RefersTo '参照先を取得
            Debug.Print "親要素：" & n.Parent.Name '親要素を取得
            Debug.Print "値：" & n.Value '親要素を取得
        End If

'        ListBox1.AddItem "名前:" + n.Name
'        ListBox1.ListIndex = 0

        total = total + 1
    Next n
    
    '結果表示
    If (0 < cnt) Then
        MsgBox cnt & " / " & total & " 件の定義を削除しました", vbInformation
    Else
        MsgBox "対象はありません", vbInformation
    End If

End Function

Function removeNameDefinition2()
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
        
        '[DEBUG]
        Debug.Print "x 名前：" & n.Name '名前を取得
        Debug.Print "x 参照先：" & n.RefersTo '参照先を取得
        Debug.Print "x 親要素：" & n.Parent.Name '親要素を取得
        Debug.Print "x 値：" & n.Value '親要素を取得
        n.Delete
        cnt = cnt + 1
        
'        ListBox1.AddItem "名前:" + n.Name
'        ListBox1.ListIndex = 0
        
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
    '印刷範囲の解除
    ActiveSheet.PageSetup.PrintArea = ""
    
    'すべての改ページを解除
    ActiveSheet.PageSetup.Zoom = 100
    ActiveSheet.resetAllPageBreaks
    
    '印刷結果の幅を指定
    With ActiveSheet.PageSetup
       .Zoom = False              '拡大・縮小率の指定なし (100 は表示パーセント)
       .FitToPagesWide = 1        '幅を 1 ページに縮小 （数字はページ数、自動はFalse）
       .FitToPagesTall = False    '縦のページ指定なし  （数字はページ数、自動はFalse）
    End With
   
    '結果表示
  '  MsgBox "完了しました", vbInformation

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
    
Dim a As Integer
a = Selection.RowHeight
MsgBox a '18
MsgBox PointToPixcel(a) '
MsgBox PixcelToPoint(a) '
'MsgBox LogicalPixcel
'MsgBox GetDpi()
    
    
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
'  [removeStyles]
'  - 不要なスタイルを削除して初期状態に戻します
' ============================================================
Function removeStyles()
'
    'エラーを無視 (例外的なスタイルは削除不可)
    On Error Resume Next

    Dim cnt As Integer: cnt = 0
    Dim oSty As Style
    
    Dim arySty As Variant

    '[DEBUG]
    Debug.Print ActiveWorkbook.Styles.Count; "件"

    MsgBox ("対象は " + ActiveWorkbook.Styles.Count + "件です。")

    For Each oSty In ActiveWorkbook.Styles
        'ビルトインスタイル（削除不可）以外を削除

    '[DEBUG]
   ' Debug.Print oSty.BuiltIn; oSty
        If oSty.BuiltIn = False Then
            frmMain.ListBox1.AddItem "oSty:" + oSty.Name
            frmMain.ListBox1.ListIndex = 0
            oSty.Delete
            cnt = cnt + 1
             
             '脱出
            If (cnt = 500) Then
                If vbNo = MsgBox("500 件を超えました。処理を継続しますか？", vbYesNo + vbDefaultButton1) Then
                    Exit For
                End If
            End If
        End If
    Next
 
    '結果表示
'    For Each oSty In ActiveWorkbook.Styles
'        If oSty.BuiltIn = False Then
'            MsgBox oSty & "は削除できませんでした"
'            cnt = cnt - 1
'        End If
'    Next
    
    If (0 < cnt) Then
        MsgBox cnt & " 件の不要なスタイルを削除しました", vbInformation
    Else
        MsgBox "対象はありません", vbInformation
    End If
   

End Function

' ============================================================
'  [removeDocumentInformation]
'  - ファイルに記録されている個人情報を削除します
' ============================================================
Function removeDocumentInformation()
'
    On Error Resume Next
    
    
'TODO: 削除前に確認ダイアログを表示
    
    Dim wb As Variant
    Set wb = ActiveWorkbook
    
    'ダイアログを非表示
    Application.DisplayAlerts = False

    '個人情報の削除を許可する
    wb.RemovePersonalInformation = True

    '個人情報を削除
    wb.removeDocumentInformation (xlRDIDocumentProperties)

    'ダイアログを表示
    Application.DisplayAlerts = True

    Debug.Print "OK"
    
    '結果表示
    MsgBox "個人情報を削除しました", vbInformation

    
End Function


' ============================================================
'  [removePhoneticCharacters]
'  - 選択範囲のふりがなをまとめて削除します
' ============================================================
Function removePhoneticCharacters_temp()
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

' ============================================================
'  [開発計画]
'  - TODO: 選択したセルの先頭のアポストロフィー削除する
' ============================================================
'VBAでシングルクォーテーションを取り除くVBA
Call prefixDelete(Range("A:B"))

Public Sub prefixDelete()
    '対象セル範囲を使用セル範囲に限定
    Dim ws As Worksheet
    Set ws = argRange.Worksheet
    Set argRange = Intersect(ws.Range(argRange.Item(1), _
                             ws.UsedRange(ws.UsedRange.Count)), _
                             argRange)

    'プレフィックス文字がある場合のみValueをコピー
    Dim myRange As Range
    For Each myRange In argRange
        If myRange.PrefixCharacter <> "" Then
            myRange.Value = myRange.Value
        End If
    Next
End Sub
'PrefixCharacterが空白以外の時だけValueをコピーしています｡
'広いセル範囲が指定された場合の無駄なコピーを省くために､UsedRangeの範囲に限定しています｡

' ============================================================
'  [開発計画]
'  - TODO: Print_Area 削除
' ============================================================

' ============================================================
'  [開発計画]
'  - TODO: フィルタ設定削除
' ============================================================

' ============================================================
'  [開発計画]
'  - TODO: ESCで終了
' ============================================================

'TODO: 余白を簡単設定
