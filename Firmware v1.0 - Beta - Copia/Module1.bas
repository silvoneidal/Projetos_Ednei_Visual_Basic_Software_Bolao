Attribute VB_Name = "Module1"
'Option Explicit
'
'Public Sub ColunasListView()
'On Error GoTo Erro
'
'    ' JOGADOR 1
'    Form1.ListView1.View = lvwReport ' Defina o estilo para exibir colunas
'    With Form1.ListView1.ColumnHeaders
'        .Add , , "Valor", 1200 ' Coluna de Itens
'        .Add , , "E", 700      ' Colunas de subItens
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T1", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T2", 1200
'        .Add , , "Sub Total", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T3", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T4", 1200
'        .Add , , "Sub Total", 1200
'        .Add , , "Total Final", 1200
'    End With
'
'    ' JOGADOR 2
'    Form1.ListView2.View = lvwReport ' Defina o estilo para exibir colunas
'    With Form1.ListView2.ColumnHeaders
'        .Add , , "Valor", 1200 ' Coluna de Itens
'        .Add , , "E", 700      ' Colunas de subItens
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T1", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T2", 1200
'        .Add , , "Sub Total", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T3", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T4", 1200
'        .Add , , "Sub Total", 1200
'        .Add , , "Total Final", 1200
'    End With
'
'    ' JOGADOR 3
'    Form1.ListView3.View = lvwReport ' Defina o estilo para exibir colunas
'    With Form1.ListView3.ColumnHeaders
'        .Add , , "Valor", 1200 ' Coluna de Itens
'        .Add , , "E", 700      ' Colunas de subItens
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T1", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T2", 1200
'        .Add , , "Sub Total", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T3", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T4", 1200
'        .Add , , "Sub Total", 1200
'        .Add , , "Total Final", 1200
'    End With
'
'    ' JOGADOR 4
'    Form1.ListView4.View = lvwReport ' Defina o estilo para exibir colunas
'    With Form1.ListView4.ColumnHeaders
'        .Add , , "Valor", 1200 ' Coluna de Itens
'        .Add , , "E", 700      ' Colunas de subItens
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T1", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T2", 1200
'        .Add , , "Sub Total", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T3", 1200
'        .Add , , "E", 700
'        .Add , , "1", 700
'        .Add , , "2", 700
'        .Add , , "3", 700
'        .Add , , "4", 700
'        .Add , , "5", 700
'        .Add , , "T4", 1200
'        .Add , , "Sub Total", 1200
'        .Add , , "Total Final", 1200
'    End With
'
'Exit Sub
'Erro:
'    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
'
'End Sub
'
'Public Sub ValoresListView()
'On Error GoTo Erro
'
'    'ListView1
'    Form1.ListView1.ListItems.Add , , Empty
'    With Form1.ListView1.ListItems(1)
'        .SubItems(7) = "0" ' T1
'        .SubItems(14) = "0" ' T2
'        .SubItems(15) = "0" ' T1 + T2
'        .SubItems(22) = "0" ' T3
'        .SubItems(29) = "0" ' T4
'        .SubItems(30) = "0" ' T3 + T4
'        .SubItems(31) = "0" ' Total Final
'    End With
'    'ListView2
'    Form1.ListView2.ListItems.Add , , Empty
'    With Form1.ListView2.ListItems(1)
'        .SubItems(7) = "0" ' T1
'        .SubItems(14) = "0" ' T2
'        .SubItems(15) = "0" ' T1 + T2
'        .SubItems(22) = "0" ' T3
'        .SubItems(29) = "0" ' T4
'        .SubItems(30) = "0" ' T3 + T4
'        .SubItems(31) = "0" ' Total Final
'    End With
'    'ListView3
'    Form1.ListView3.ListItems.Add , , Empty
'    With Form1.ListView3.ListItems(1)
'        .SubItems(7) = "0" ' T1
'        .SubItems(14) = "0" ' T2
'        .SubItems(15) = "0" ' T1 + T2
'        .SubItems(22) = "0" ' T3
'        .SubItems(29) = "0" ' T4
'        .SubItems(30) = "0" ' T3 + T4
'        .SubItems(31) = "0" ' Total Final
'    End With
'    'ListView4
'    Form1.ListView4.ListItems.Add , , Empty
'    With Form1.ListView4.ListItems(1)
'        .SubItems(7) = "0" ' T1
'        .SubItems(14) = "0" ' T2
'        .SubItems(15) = "0" ' T1 + T2
'        .SubItems(22) = "0" ' T3
'        .SubItems(29) = "0" ' T4
'        .SubItems(30) = "0" ' T3 + T4
'        .SubItems(31) = "0" ' Total Final
'    End With
'
'Exit Sub
'Erro:
'    MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
'
'End Sub
'
