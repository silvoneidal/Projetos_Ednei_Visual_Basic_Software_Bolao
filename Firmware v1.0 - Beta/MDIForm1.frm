VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12600
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    frmTreino.Show
    Me.Caption = App.Title & "_v" & App.Major & "." & App.Minor & " Beta " & " by DALÇOQUIO AUTOMAÇÃO" & "  " & "[ " & frmTreino.addressRegisters & " ]"
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    StartBackup  ' Executa backup do bando de dados
    
End Sub
