Sub CopiarRREO()
    Dim wbOrigem As Workbook, wbDestino As Workbook
    Dim caminhoOrigem As String

    caminhoOrigem = "C:\caminho\planilha_auditoria.xls" ' <<< AJUSTE O CAMINHO AQUI

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wbDestino = ThisWorkbook
    Set wbOrigem = Workbooks.Open(caminhoOrigem)

    ' === Cópia por Anexo ===
    CopiarIntervalos wbOrigem, wbDestino, "RREO-Anexo 01", Array( _
        Array("B:D,E,H:I,L", 21, 98), _
        Array("B:D,E,G:H,J:K", 107, 129), _
        Array("B:D,F", 139, 201), _
        Array("B:D,E,G:H,J:K", 210, 219) _
    )

    CopiarIntervalos wbOrigem, wbDestino, "RREO-Anexo 02", Array( _
        Array("B:E,H:I,L", 19, 212), _
        Array("B:E,H:I,L", 221, 413) _
    )

    CopiarIntervalos wbOrigem, wbDestino, "RREO-Anexo 03", Array( _
        Array("B:O", 21, 55) _
    )

    CopiarIntervalos wbOrigem, wbDestino, "RREO-Anexo 04", Array( _
        Array("B:C", 20, 42), _
        Array("B:F", 51, 58), _
        Array("B:B", 67, 67), _
        Array("B:B", 76, 76), _
        Array("B:B", 85, 88), _
        Array("B:B", 97, 99), _
        Array("B:C", 108, 129), _
        Array("B:F", 138, 145), _
        Array("B:B", 154, 155), _
        Array("B:B", 164, 166), _
        Array("B:C", 175, 176), _
        Array("B:F", 185, 190), _
        Array("B:B", 199, 201), _
        Array("B:C", 210, 212), _
        Array("B:F", 221, 225) _
    )

    CopiarIntervalos wbOrigem, wbDestino, "RREO-Anexo 06", Array( _
        Array("B:C", 21, 63), _
        Array("B:H", 74, 94), _
        Array("B:C", 103, 104), _
        Array("B:B", 113, 113), _
        Array("B:B", 122, 123), _
        Array("B:B", 132, 132), _
        Array("B:C", 141, 148), _
        Array("B:B", 157, 157), _
        Array("B:B", 166, 166), _
        Array("B:B", 175, 181), _
        Array("B:B", 190, 190), _
        Array("B:B", 199, 202) _
    )

    CopiarIntervalos wbOrigem, wbDestino, "RREO-Anexo 07", Array( _
        Array("B:M", 22, 28), _
        Array("B:M", 39, 43) _
    )

    CopiarIntervalos wbOrigem, wbDestino, "RREO-Anexo 13", Array( _
        Array("B:B", 22, 30), _
        Array("B:L", 67, 72) _
    )

    CopiarIntervalos wbOrigem, wbDestino, "RREO-Anexo 14", Array( _
        Array("B:B", 20, 32), _
        Array("B:B", 41, 42), _
        Array("B:B", 51, 53), _
        Array("B:B", 62, 73), _
        Array("B:D", 82, 83), _
        Array("B:E", 92, 103), _
        Array("B:D", 114, 117), _
        Array("B:C", 126, 127), _
        Array("B:E", 136, 142), _
        Array("B:C", 152, 153), _
        Array("B:D", 163, 163), _
        Array("B:B", 172, 172) _
    )

    wbOrigem.Close SaveChanges:=False

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Cópia concluída com sucesso!", vbInformation
End Sub

Sub CopiarIntervalos(wbOrigem As Workbook, wbDestino As Workbook, nomeAba As String, intervalos As Variant)
    Dim i As Integer
    Dim cols As String, linInicio As Long, linFim As Long
    Dim subCol As Variant, colStr As String
    Dim wsOrigem As Worksheet, wsDestino As Worksheet
    Dim rngOrigem As Range, cel As Range

    On Error Resume Next
    Set wsOrigem = wbOrigem.Sheets(nomeAba)
    Set wsDestino = wbDestino.Sheets(nomeAba)
    On Error GoTo 0

    If wsOrigem Is Nothing Or wsDestino Is Nothing Then
        MsgBox "Aba '" & nomeAba & "' não encontrada em uma das planilhas.", vbExclamation
        Exit Sub
    End If

    For i = LBound(intervalos) To UBound(intervalos)
        cols = intervalos(i)(0)
        linInicio = intervalos(i)(1)
        linFim = intervalos(i)(2)

        For Each subCol In Split(cols, ",")
            colStr = Trim(subCol)
            On Error Resume Next
            Set rngOrigem = wsOrigem.Range(colStr & linInicio & ":" & colStr & linFim)
            On Error GoTo 0

            If Not rngOrigem Is Nothing Then
                For Each cel In rngOrigem.Cells
                    If IsEmpty(wsDestino.Cells(cel.Row, cel.Column).Value) Then
                        wsDestino.Cells(cel.Row, cel.Column).Value = cel.Value
                    End If
                Next cel
                Set rngOrigem = Nothing
            End If
        Next subCol
    Next i
End Sub
