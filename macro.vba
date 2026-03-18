Option Explicit

Sub PrepararEnvioPAF()

    On Error GoTo Erro

    Dim wsForm As Worksheet
    Dim wsBase As Worksheet
    Dim proximaLinha As Long
    
    Set wsForm = ThisWorkbook.Sheets("PAF")
    Set wsBase = ThisWorkbook.Sheets("BASE_ENVIO")
    
    ' =========================
    ' VALIDAR CAMPOS
    ' =========================
    If Not CamposValidos(wsForm) Then
        MsgBox "Preencha todos os campos obrigatórios.", vbExclamation
        Exit Sub
    End If
    
    ' =========================
    ' DEFINIR LINHA
    ' =========================
    proximaLinha = ObterProximaLinha(wsBase)
    
    ' =========================
    ' SALVAR DADOS
    ' =========================
    SalvarDados wsForm, wsBase, proximaLinha
    
    ' =========================
    ' GERAR EMAIL
    ' =========================
    GerarEmail wsForm
    
    ' =========================
    ' LIMPAR CAMPOS
    ' =========================
    LimparCampos wsForm
    
    MsgBox "Processo concluído com sucesso!", vbInformation
    Exit Sub

Erro:
    MsgBox "Erro: " & Err.Description, vbCritical

End Sub

' =========================
' VALIDAÇÃO
' =========================
Function CamposValidos(ws As Worksheet) As Boolean

    CamposValidos = _
        ws.Range("H4").Value <> "" And _
        ws.Range("H5").Value <> "" And _
        ws.Range("H6").Value <> "" And _
        ws.Range("J12").Value <> ""

End Function

' =========================
' PRÓXIMA LINHA
' =========================
Function ObterProximaLinha(ws As Worksheet) As Long

    Dim linha As Long
    linha = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If ws.Cells(linha, "A").Value <> "" Then
        linha = linha + 1
    End If
    
    If linha < 2 Then linha = 2
    
    ObterProximaLinha = linha

End Function

' =========================
' SALVAR DADOS
' =========================
Sub SalvarDados(wsForm As Worksheet, wsBase As Worksheet, linha As Long)

    Dim dados As Variant
    
    dados = Array( _
        wsForm.Range("H4").Value, _
        wsForm.Range("H5").Value, _
        wsForm.Range("H6").Value, _
        wsForm.Range("H7").Value, _
        wsForm.Range("H8").Value, _
        wsForm.Range("H9").Value, _
        wsForm.Range("H10").Value, _
        wsForm.Range("J12").Value _
    )
    
    Dim i As Integer
    For i = 0 To UBound(dados)
        wsBase.Cells(linha, i + 1).Value = dados(i)
    Next i
    
    wsBase.Cells(linha, 9).Value = _
        wsForm.Range("F15").Value & " - " & _
        wsForm.Range("I15").Value & " - " & _
        wsForm.Range("I16").Value
        
    wsBase.Cells(linha, 10).Value = "PENDENTE"

End Sub

' =========================
' GERAR EMAIL
' =========================
Sub GerarEmail(wsForm As Worksheet)

    Dim assunto As String
    Dim corpo As String
    Dim link As String
    
    assunto = wsForm.Range("N7").Value
    
    corpo = MontarCorpoEmail(wsForm)
    
    link = "https://outlook.office.com/mail/deeplink/compose?" & _
           "to=exemplo@empresa.com" & _
           "&subject=" & UrlEncode(assunto) & _
           "&body=" & UrlEncode(corpo)
    
    ThisWorkbook.FollowHyperlink link

End Sub

' =========================
' CORPO DO EMAIL
' =========================
Function MontarCorpoEmail(ws As Worksheet) As String

    MontarCorpoEmail = _
        "Bom dia," & vbCrLf & vbCrLf & _
        "Solicitação de emissão de nota:" & vbCrLf & vbCrLf & _
        "FILIAL: " & ws.Range("H4").Value & vbCrLf & _
        "DESTINO: " & ws.Range("H5").Value & vbCrLf & _
        "TRANSPORTE: " & ws.Range("H6").Value & vbCrLf & _
        "FRETE: " & ws.Range("H7").Value & vbCrLf & _
        "VOLUMES: " & ws.Range("H8").Value & vbCrLf & _
        "PESO: " & ws.Range("H9").Value & vbCrLf & _
        "CENTRO DE CUSTO: " & ws.Range("H10").Value & vbCrLf & vbCrLf & _
        "Saídas:" & vbCrLf & _
        "- " & ws.Range("J12").Value & vbCrLf & _
        "- " & ws.Range("J13").Value & vbCrLf & _
        "- " & ws.Range("J14").Value & vbCrLf & vbCrLf & _
        "Observações:" & vbCrLf & _
        ws.Range("F15").Value & " - " & _
        ws.Range("I15").Value & " - " & _
        ws.Range("I16").Value

End Function

' =========================
' LIMPAR CAMPOS
' =========================
Sub LimparCampos(ws As Worksheet)

    Dim cel As Range
    For Each cel In ws.Range("F15,J12,J13,J14")
        cel.MergeArea.ClearContents
    Next cel

End Sub

' =========================
' URL ENCODE
' =========================
Function UrlEncode(ByVal texto As String) As String
    UrlEncode = Application.WorksheetFunction.EncodeURL(texto)
End Function
