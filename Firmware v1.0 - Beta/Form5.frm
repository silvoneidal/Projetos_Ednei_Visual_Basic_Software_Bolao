VERSION 5.00
Begin VB.Form frmControle 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DALÇÓQUIO AUTOMAÇÃO"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "Form5.frx":0000
      Top             =   1080
      Width           =   6855
   End
   Begin VB.CommandButton cmdControle 
      Caption         =   "Habilitado"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.ComboBox cboTime 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      ItemData        =   "Form5.frx":0006
      Left            =   5040
      List            =   "Form5.frx":0008
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frmControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim x As Integer
    For x = 1 To 9
        cboTime.AddItem (x) * 1000
    Next x
    
    cboTime.Text = frmTreino.tmrDownload.Interval
    
    ' Texto de Informação
    txtInfo.Text = "Se controle habilitado, a atualização será automática de" & vbCrLf
    txtInfo.Text = txtInfo.Text & "acordo com o tempo (ms) selecionado, caso desabilitado a" & vbCrLf
    txtInfo.Text = txtInfo.Text & "atualização será seguido de cada ponto atualizado pelo usuário."
    
End Sub

Private Sub cboTime_Change()
    frmTreino.tmrDownload.Interval = cboTime.Text
    
End Sub

Private Sub cmdControle_Click()
    If cmdControle.Caption = "Habilitado" Then
        frmTreino.tmrDownload.Enabled = False
        cmdControle.Caption = "Desabilitado"
    Else
        frmTreino.tmrDownload.Enabled = True
        cmdControle.Caption = "Habilitado"
    End If
End Sub


