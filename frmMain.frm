VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "Excel Finisher"
   ClientHeight    =   3756
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   6420
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   2  '��ʂ̒���
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================
' �t�H�[��������
'================
Private Sub UserForm_Initialize()
    Call setSheetValue
    Call setFormValue
End Sub

'========
' �{�^��
'========
'[���̃V�[�g�̐ݒ�擾]�{�^��
Private Sub cmdGetValue_Click()
    Call getSheetValue
    Call setFormValue
End Sub

'[�S�V�[�g����]�{�^��
Private Sub cmdExecute_Click()
    If Not getFormValue Then
        Call actionFinisher
    End If
End Sub

'[�s�v�Ȗ��O�̒�`���폜]�{�^��
Private Sub cmdFunc1_Click()
    Call removeNameDefinition
End Sub

'[��\���V�[�g�`�F�b�N]�{�^��
Private Sub cmdFunc2_Click()
    Call checkUnhideSheets
End Sub

'[���y�[�W����]�{�^��
Private Sub cmdFunc3_Click()
    Call resetAllPageBreaks
End Sub

'[�V�[�g���ꗗ�o��]�{�^��
Private Sub cmdFunc4_Click()
    Call createSheetList
End Sub

'[���ᎆ�V�[�g�ǉ�]�{�^��
Private Sub cmdFunc5_Click()
    Call createGraphPaper
End Sub

'[�I��͈͂̂ӂ肪�Ȃ��܂Ƃ߂č폜]�{�^��
Private Sub cmdFunc6_Click()
    Call removePhoneticCharacters
End Sub

'[�C��]�{�^��
Private Sub cmdEnd_Click()
    End
End Sub

