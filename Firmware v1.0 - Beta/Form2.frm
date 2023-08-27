VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCadastro 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8055
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   8055
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   4320
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form2.frx":0000
      OLEDBString     =   $"Form2.frx":00C5
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TabelaCadastro"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7815
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   3840
      Width           =   2535
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Novo"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   2535
   End
End
Attribute VB_Name = "frmCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Notas:
' If Adodc1.Recordset.EOF = True porque chegou a pecorrer todo o registro
' If Adodc1.Recordset.EOF = False porque não chegou a pecorrer todo o resgistro

Option Explicit

Dim query As String

'//////////////////////////////////////////////////////////////////////////////////////////////
' SQL PARA CONSULTA NO BANCO DE DADOS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub queryString(query As String)
    ' Comando para SQL
    Adodc1.RecordSource = query
    Adodc1.CommandType = adCmdText
    Adodc1.Refresh
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' INICIO DO FORMULÁRIO CADASTRO
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Form_Load()
    ' Atualiza a lista com valores de registro
    Call UpdateList1
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' COMANDO PARA CADASTRAR NOVO NOME
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdNovo_Click()
On Error GoTo Erro

    Dim newName As String
    newName = InputBox("Digite o nome do jogador.", "DALCOQUIO AUTOMAÇÃO")
    
    If newName <> "" Then
        ' Configurações para Registros
        query = "SELECT * FROM TabelaCadastro"
        Call queryString(query)
        
        ' Busca no registro se nome existe
        Do While Not Adodc1.Recordset.EOF
            If newName = Adodc1.Recordset("NOME") Then
                Beep
                MsgBox "Nome já cadastrado.", vbInformation, "DALCOQUIO AUTOMAÇÃO"
                Exit Sub
            End If
            Adodc1.Recordset.MoveNext
        Loop
        
        ' Grava novo registro
        If Adodc1.Recordset.EOF = True Then
            Adodc1.Recordset.AddNew
            Adodc1.Recordset!Data = Date
            Adodc1.Recordset!HORA = Time
            Adodc1.Recordset!Nome = newName
            Adodc1.Recordset.Update
            Call UpdateList1
        End If
    End If
    
Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' COMANDO PARA EDITAR UM NOME
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdEditar_Click()
On Error GoTo Erro

    ' Configurações para Registros
    query = "SELECT * FROM TabelaCadastro"
    Call queryString(query)
        
    ' Verifica se nome selecionado
    Dim valueSelected As String
    If List1.ListIndex >= 0 Then
        valueSelected = List1.List(List1.ListIndex)
        ' Busca no registro nome selecionado
        Do While Not Adodc1.Recordset.EOF
            If valueSelected = Adodc1.Recordset("NOME") Then
                Exit Do ' Nome localizado
            End If
            Adodc1.Recordset.MoveNext
        Loop
        If Adodc1.Recordset.EOF = False Then
            ' Solicitação ao usuário para edição do nome
            Dim newName As String
            newName = InputBox("Edite o nome do jogador.", "DALCOQUIO AUTOMAÇÃO", Adodc1.Recordset!Nome)
            ' Se ok atualiza registro
            If newName <> "" Then
                Adodc1.Recordset!Data = Date
                Adodc1.Recordset!HORA = Time
                Adodc1.Recordset!Nome = newName
                Adodc1.Recordset.Update
                Call UpdateList1
            End If
        End If
    Else
        Beep
        MsgBox "Nenhum nome selecionado.", vbInformation, "DALCOQUIO AUTOMAÇÃO"
    End If

Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    

End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' COMANDO PARA EXCLUIR UM NOME
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdExcluir_Click()
On Error GoTo Erro

    ' Configurações para Registros
    query = "SELECT * FROM TabelaCadastro"
    Call queryString(query)
        
    'Verifica se nome selecionado
    Dim valueSelected As String
    If List1.ListIndex >= 0 Then
        valueSelected = List1.List(List1.ListIndex)
        ' Busca no registro nome selecionado
        Do While Not Adodc1.Recordset.EOF
            If valueSelected = Adodc1.Recordset("NOME") Then
                Exit Do ' Nome localizado
            End If
            Adodc1.Recordset.MoveNext
        Loop
        ' Exclui registro selecionado
        If Adodc1.Recordset.EOF = False Then
            Adodc1.Recordset.Delete
            Call UpdateList1
        End If
    Else
        Beep
        MsgBox "Nenhum nome selecionado !!!", , "DALCOQUIO AUTOMAÇÃO"
    End If

Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' ATUALIZA A LISTA DE REGISTROS
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub UpdateList1()
On Error GoTo Erro
    List1.Clear
    
    ' Configurações para Registros
    query = "SELECT * FROM TabelaCadastro ORDER by NOME ASC"
    Call queryString(query)
    
    ' Atualiza lista
    Do While Not Adodc1.Recordset.EOF
        List1.AddItem Adodc1.Recordset("NOME")
        Adodc1.Recordset.MoveNext
    Loop
    
    ' Fecha conexão com o registro
    'Adodc1.Recordset.Close
    
Exit Sub

Erro:
    Beep
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
    
End Sub

'//////////////////////////////////////////////////////////////////////////////////////////////
' TRATAMENTO PARA EXECUÇÃO ANTES DO FECHAMENTO O FORMULÁRIO
'//////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Form_Unload(Cancel As Integer)
   ' Atualiza nomes do form principal
   Call frmTreino.updateCboNames

End Sub



