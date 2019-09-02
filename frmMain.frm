VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Excel Finisher"
   ClientHeight    =   4056
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6420
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Call resetAllPageBreaks
End Sub

Private Sub UserForm_Initialize()
    Call setシート初期値
    Call setフォーム値
End Sub

Private Sub cmdGetValue_Click()
    Call getシート設定値
End Sub

Private Sub cmdExecute_Click()
    If Not getフォーム値 Then
        Call actionシート仕上げ
    End If
End Sub

Private Sub cmdFunc1_Click()
    Call checkUnhideSheets
End Sub

Private Sub cmdFunc2_Click()
    Call removeNameDefinition
End Sub

Private Sub cmdFunc3_Click()
    Call resetAllPageBreaks
End Sub

Private Sub cmdFunc4_Click()
    Call createSheetList
End Sub

Private Sub cmdFunc5_Click()
    Call createGraphPaper
End Sub

Private Sub cmdFunc6_Click()
    Call removePhoneticCharacters
End Sub

Private Sub cmdEnd_Click()
    End
End Sub

