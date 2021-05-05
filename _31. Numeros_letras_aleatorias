Sub Numb_let()
    Dim obj As Object
    Dim linha As Integer, coluna As Integer
    Dim r As Integer, c As Integer
    Set obj = Range("A1")
    linha = obj.Row
    coluna = obj.Column
    For r = 0 To 10
        For c = 0 To 5
            Cells(linha + r, coluna + c).Value = Chr(65 + Int(10 * Rnd)) _
            & Chr(65 + Int(10 * Rnd)) & Chr(65 + Int(10 * Rnd)) _
            & Chr(48 + Int(10 * Rnd)) & Chr(48 + Int(10 * Rnd)) _
            & Chr(48 + Int(10 * Rnd))
        Next c
    Next r
End Sub
