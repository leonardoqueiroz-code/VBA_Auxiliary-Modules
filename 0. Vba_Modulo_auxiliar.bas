Sub Funcion_date_now()
    'Data atual
    date_now = Date
    MsgBox date_now
End Sub
Sub Função_Time_now()
    'Horario atual
    time_now = Time
    MsgBox time_now
End Sub
Sub COW_ROW_INS_DEL()
    'Inserção e delete de linha
    Rows(1).Insert 'Inserirá uma linha na Linha 1.
    Rows(1).Delete 'Deletará a Linha 1.
    'Inserção e delete de coluna
    Columns("A").Insert 'Inserirá uma Coluna a partir da A.
    Columns("A").Delete 'Deletará a Coluna A.
End Sub
Sub Insert_values()
    Planilha1.Range("A1") = "Value"
    'Inserirá um valor (VALUE) na celula A1
End Sub
Sub Update_values()
    Planilha1.Range("A1:G37").ClearContents
    Planilha1.Range("A1") = "Value"
    'Inserirá um valor (VALUE) na celula A1
End Sub
Sub Read_values()
    read_value = Planilha1.Range("A1").Value
    MsgBox read_value
    'Realiza a leitura de uma celula ou range
End Sub
Sub Delete_values()
    Planilha1.Range("A1:G37").ClearContents
    'Realiza a exclusão os dados de um range
End Sub
