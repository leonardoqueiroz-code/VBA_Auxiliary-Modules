Sub COW_ROW_INS_DEL()
    'Inserção e delete de linha
    Rows(1).Insert 'Inserirá uma linha na Linha 1.
    Rows(1).Delete 'Deletará a Linha 1.
    'Inserção e delete de coluna
    Columns("A").Insert 'Inserirá uma Coluna a partir da A.
    Columns("A").Delete 'Deletará a Coluna A.
End Sub
