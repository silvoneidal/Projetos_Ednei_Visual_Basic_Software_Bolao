Attribute VB_Name = "ModuleBKP"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub StartBackup()
On Error GoTo Erro

    MDIForm1.Caption = "Aguarde, efetuando backup do banco de dados em: " & frmTreino.addressBackups
    
    DoEvents ' Permite que o sistema continue respondendo
    Sleep (3000) ' Aguarda um tempo...
    

    ' Efetuar Bakcup do Banco de Dados
    Dim deletFilePath As String
    Dim copyFilePath As String
    Dim destinationPath As String
    
    ' Caminho completo do arquivo a ser deletado
    deletFilePath = frmTreino.addressBackups & "\bkpRegistrosBolao.mdb"
    
    ' Caminho completo do arquivo a ser copiado
    copyFilePath = frmTreino.addressRegisters & "\RegistrosBolao.mdb"
    
    ' Caminho de destino para onde o arquivo será copiado
    destinationPath = frmTreino.addressBackups & "\bkpRegistrosBolao.mdb"
    
    ' Deletar o arquivo original
    If Dir(deletFilePath) <> "" Then ' Verifica se o arquivo existe
        Kill deletFilePath
    End If

    ' Copiar o novo arquivo para o diretório de destino
    FileCopy copyFilePath, destinationPath
    
Exit Sub
Erro:
    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"

End Sub


