VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Excel Finisher"
   ClientHeight    =   4365
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7755
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================
' フォーム初期化
'================
Private Sub UserForm_Initialize()
    Call setSheetValue
    Call setFormValue
    Call setPageList
'    Application.OnKey "{ESCAPE}", "cmdEnd_Click"
End Sub

'========
' ボタン
'========
'[このシートの設定取得]ボタン
Private Sub cmdGetValue_Click()
    Call getSheetValue
    Call setFormValue
End Sub

'[全シート統一]ボタン
Private Sub cmdExecute_Click()
    If Not getFormValue Then
        Call actionFinisher
    End If
End Sub

'[不要な名前の定義を削除]ボタン
Private Sub cmdFunc1_Click()
    Call removeNameDefinition
End Sub
Private Sub CommandButton1_Click()
    Call removeNameDefinition2
End Sub

'[非表示シートチェック]ボタン
Private Sub cmdFunc2_Click()
    Call checkUnhideSheets
End Sub

'[改ページ解除]ボタン
Private Sub cmdFunc3_Click()
    Call resetAllPageBreaks
End Sub

'[シート名一覧出力]ボタン
Private Sub cmdFunc4_Click()
    Call createSheetList
End Sub

'[方眼紙シート追加]ボタン
Private Sub cmdFunc5_Click()
    Call createGraphPaper
End Sub

'[選択範囲のふりがなをまとめて削除]ボタン
Private Sub cmdFunc6_Click()
'    Call removePhoneticCharacters
    Call removeDocumentInformation
End Sub

Private Sub cmdFunc7_Click()
    Call removeStyles
End Sub

'[ページ設定]ボタン
Private Sub cmdPageSet_Click()
    Call setPageStyle(frmMain.cmbPageList.ListIndex)
End Sub

'[終了]ボタン
'Private Sub cmdEnd_Click()
'    End
'End Sub

'TODO: Application.EnableCancelKey
