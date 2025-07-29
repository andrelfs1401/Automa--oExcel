Sub ColetarInformacoes()

    Dim wbOrigem As Workbook
    Dim wsOrigem As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaLinhaOrigem As Long
    Dim ultimaLinhaDestino As Long
    Dim i As Long
    Dim col As Long

    On Error GoTo TrataErro

    ' Abre o arquivo de origem
    Set wbOrigem = Workbooks.Open("C:\Users\ANDRE\Downloads\teste.xlsx")
    Set wsOrigem = wbOrigem.Sheets("Planilha1")
    Set wsDestino = ThisWorkbook.Sheets("Enviados 1° Janeiro2021")

    ' Encontra a última linha da planilha de origem e a próxima linha disponível na planilha destino
    ultimaLinhaOrigem = wsOrigem.Cells(wsOrigem.Rows.Count, "A").End(xlUp).Row
    ultimaLinhaDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row + 1

    ' Percorre cada linha da origem
    For i = 2 To ultimaLinhaOrigem
        ' Percorre as colunas de B (2) até AD (30)
        For col = 1 To 30
            If col <> 9 And col <> 27 Then ' Ignora as colunas I (9) e AA (27)
                wsDestino.Cells(ultimaLinhaDestino, col).Value = wsOrigem.Cells(i, col).Value
            End If
        Next col
        ultimaLinhaDestino = ultimaLinhaDestino + 1
    Next i

    MsgBox "Informações coletadas com sucesso!"
    wbOrigem.Close SaveChanges:=False
    Exit Sub

TrataErro:
    MsgBox "Ocorreu um erro: " & Err.Description, vbCritical

End Sub