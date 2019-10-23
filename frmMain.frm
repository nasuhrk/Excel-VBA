VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Excel Finisher"
   ClientHeight    =   3756
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6420
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
    Call removePhoneticCharacters
End Sub

'[修了]ボタン
Private Sub cmdEnd_Click()
    End
End Sub

