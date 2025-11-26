Sub UnificarMacroRREO()
    Dim caminhoOrigem As String
    caminhoOrigem = "C:\caminho\para\planilha_auditoria.xls" ' <<< AJUSTE AQUI O CAMINHO CORRETO

    ' Abrir a planilha de origem
    Dim wbOrigem As Workbook
    On Error Resume Next
    Set wbOrigem = Workbooks.Open(caminhoOrigem, ReadOnly:=True)
    If wbOrigem Is Nothing Then
        MsgBox "Erro ao abrir a planilha de origem.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    Dim wbDestino As Workbook: Set wbDestino = ThisWorkbook

    ' === Copiar dados entre planilhas ===
    ' Anexo 01
    CopiarLinhas wbOrigem.Sheets("RREO-Anexo 01"), wbDestino.Sheets("RREO-Anexo 01"), _
        Array(Array(21, 98, Array("B", "C", "D", "F")), _
              Array(107, 129, Array("B", "C", "D", "E", "G", "H", "J", "K")), _
              Array(139, 201, Array("B", "C", "D", "F")), _
	      Array(210, 219, Array("B", "C", "D", "E", "G", "H", "J", "K")))

    ' Anexo 02
    CopiarLinhas wbOrigem.Sheets("RREO-Anexo 02"), wbDestino.Sheets("RREO-Anexo 02"), _
        Array(Array(19, 212, Array("B", "C", "D", "E", "H", "I", "L")), _
              Array(221, 413, Array("B", "C", "D", "E", "H", "I", "L")))


    ' Anexo 03
    CopiarLinhas wsOrigem:=wbOrigem.Sheets("RREO-Anexo 03"), wsDestino:=wbDestino.Sheets("RREO-Anexo 03"), _
                 intervalo:=Array(Array(21, 55, Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O")))

    ' Anexo 04
    CopiarLinhas wsOrigem:=wbOrigem.Sheets("RREO-Anexo 04"), wsDestino:=wbDestino.Sheets("RREO-Anexo 04"), _
                 intervalo:=Array(Array(20, 42, Array("B", "C")), _
                                 Array(51, 58, Array("B", "C", "D", "E", "F")), _
                                 Array(67, 67, Array("B")), _
                                 Array(76, 76, Array("B")), _
                                 Array(85, 88, Array("B")), _
                                 Array(97, 99, Array("B")), _
                                 Array(108, 129, Array("B", "C")), _
                                 Array(138, 145, Array("B", "C", "D", "E", "F")), _
                                 Array(154, 155, Array("B")), _
                                 Array(164, 166, Array("B")), _
                                 Array(175, 176, Array("B", "C")), _
                                 Array(185, 190, Array("B", "C", "D", "E", "F")), _
                                 Array(199, 201, Array("B")), _
                                 Array(210, 212, Array("B", "C")), _
                                 Array(221, 225, Array("B", "C", "D", "E", "F")))

    ' Anexo 06
    CopiarLinhas wsOrigem:=wbOrigem.Sheets("RREO-Anexo 06"), wsDestino:=wbDestino.Sheets("RREO-Anexo 06"), _
                 intervalo:=Array(Array(21, 63, Array("B", "C")), _
                                 Array(74, 94, Array("B", "C", "D", "E", "F", "G", "H")), _
                                 Array(103, 104, Array("B", "C")), _
                                 Array(113, 113, Array("B")), _
                                 Array(122, 123, Array("B")), _
                                 Array(132, 132, Array("B")), _
                                 Array(141, 148, Array("B", "C")), _
                                 Array(157, 157, Array("B")), _
                                 Array(166, 166, Array("B")), _
                                 Array(175, 181, Array("B")), _
                                 Array(190, 190, Array("B")), _
                                 Array(199, 202, Array("B")))

    ' Anexo 07
    CopiarLinhas wsOrigem:=wbOrigem.Sheets("RREO-Anexo 07"), wsDestino:=wbDestino.Sheets("RREO-Anexo 07"), _
                 intervalo:=Array(Array(22, 28, Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M")), _
                                 Array(39, 43, Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M")))

    ' Anexo 14
    CopiarLinhas wsOrigem:=wbOrigem.Sheets("RREO-Anexo 14"), wsDestino:=wbDestino.Sheets("RREO-Anexo 14"), _
                 intervalo:=Array(Array(20, 32, Array("B")), _
                                 Array(41, 42, Array("B")), _
                                 Array(51, 53, Array("B")), _
                                 Array(62, 73, Array("B")), _
                                 Array(82, 83, Array("B", "C", "D")), _
                                 Array(92, 104, Array("B", "C", "D", "E")), _
                                 Array(114, 117, Array("B", "C", "D")), _
                                 Array(126, 127, Array("B", "C")), _
                                 Array(136, 143, Array("B", "C", "D", "E")), _
                                 Array(152, 153, Array("B", "C")), _
                                 Array(163, 163, Array("B", "C", "D")), _
                                 Array(172, 172, Array("B")))

    ' Fecha a planilha de origem
    wbOrigem.Close SaveChanges:=False

    ' === Preencher fórmulas nos Anexos ===
    Call PreencherFormulasAnexos

    MsgBox "Planilha Copiada com sucesso.", vbInformation
End Sub

Sub PreencherFormulasAnexos()
    Dim ws1 As Worksheet: Set ws1 = ThisWorkbook.Sheets("RREO-Anexo 01")
    Dim ws2 As Worksheet: Set ws2 = ThisWorkbook.Sheets("RREO-Anexo 02")
    Dim i As Long

    ' Anexo 01 - Coluna E, G, H, F, I
    For i = 21 To 98
        ws1.Cells(i, "E").Formula = "=D" & i & "/C" & i & "*100"
        ws1.Cells(i, "G").Formula = "=F" & i & "/C" & i & "*100"
        If IsEmpty(ws1.Cells(i, "H")) Then ws1.Cells(i, "H").Formula = "=C" & i & "-F" & i
    Next i

    For i = 139 To 201
        ws1.Cells(i, "E").Formula = "=D" & i & "/C" & i & "*100"
        ws1.Cells(i, "G").Formula = "=F" & i & "/C" & i & "*100"
        If IsEmpty(ws1.Cells(i, "H")) Then ws1.Cells(i, "H").Formula = "=C" & i & "-F" & i
    Next i

    For i = 107 To 129
        If IsEmpty(ws1.Cells(i, "F")) Then ws1.Cells(i, "F").Formula = "=C" & i & "-E" & i
        If IsEmpty(ws1.Cells(i, "I")) Then ws1.Cells(i, "I").Formula = "=C" & i & "-H" & i
    Next i

    For i = 210 To 219
        If IsEmpty(ws1.Cells(i, "F")) Then ws1.Cells(i, "F").Formula = "=C" & i & "-E" & i
        If IsEmpty(ws1.Cells(i, "I")) Then ws1.Cells(i, "I").Formula = "=C" & i & "-H" & i
    Next i

    ' Anexo 02 - Fórmulas F, G, J, K
    For i = 19 To 212
        If IsEmpty(ws2.Cells(i, "F")) Then ws2.Cells(i, "F").FormulaLocal = "=ARRED(E" & i & "/$E$213*100;2)"
        If IsEmpty(ws2.Cells(i, "G")) Then ws2.Cells(i, "G").Formula = "=C" & i & "-E" & i
        If IsEmpty(ws2.Cells(i, "J")) Then ws2.Cells(i, "J").FormulaLocal = "=ARRED(I" & i & "/$I$213*100;2)"
        If IsEmpty(ws2.Cells(i, "K")) Then ws2.Cells(i, "K").Formula = "=C" & i & "-I" & i
    Next i

    For i = 221 To 413
        If IsEmpty(ws2.Cells(i, "F")) Then ws2.Cells(i, "F").FormulaLocal = "=ARRED(E" & i & "/$E$213*100;2)"
        If IsEmpty(ws2.Cells(i, "G")) Then ws2.Cells(i, "G").Formula = "=C" & i & "-E" & i
        If IsEmpty(ws2.Cells(i, "J")) Then ws2.Cells(i, "J").FormulaLocal = "=ARRED(I" & i & "/$I$213*100;2)"
        If IsEmpty(ws2.Cells(i, "K")) Then ws2.Cells(i, "K").Formula = "=C" & i & "-I" & i
    Next i
End Sub

Sub CopiarLinhas(wsOrigem As Worksheet, wsDestino As Worksheet, intervalo As Variant)
    Dim i As Long, j As Long
    Dim linhaIni As Long, linhaFim As Long
    Dim colunas As Variant, letra As Variant

    For j = LBound(intervalo) To UBound(intervalo)
        linhaIni = intervalo(j)(0)
        linhaFim = intervalo(j)(1)
        colunas = intervalo(j)(2)

        For i = linhaIni To linhaFim
            For Each letra In colunas
                If IsEmpty(wsDestino.Range(letra & i)) Then
                    wsDestino.Range(letra & i).Value = wsOrigem.Range(letra & i).Value
                End If
            Next letra
        Next i
    Next j
End Sub
