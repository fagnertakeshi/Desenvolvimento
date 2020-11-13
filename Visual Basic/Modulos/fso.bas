Attribute VB_Name = "fso"
Dim fso As New FileSystemObject
Dim arqtxt As TextStream
Dim arq1 As TextStream
Dim arq3 As TextStream
Dim arq2 As TextStream
Dim texto As String


Sub CriarArquivoTexto(pArqTexto As String, pCaminhoTexto As String)

Set arqtxt = fso.CreateTextFile(pCaminhoTexto, True)
With arqtxt
    .Write pArqTexto
    .Close
End With

End Sub


Function LerArquivoTexto(pArqTexto As String, pCaminhoTexto As String) As String


On Error GoTo TrataErro
Set arq2 = fso.OpenTextFile(pCaminhoTexto, ForReading, True)
texto = arq2.ReadAll

'mostrando o conteúdo do arquivo
LerArquivoTexto = texto

arq2.Close
Exit Function

TrataErro:
If Err.Number = 53 Then
   MsgBox "Arquivo <<" & pCaminhoTexto & ">> não encontrado !", vbCritical
Else
   MsgBox Err.Description & " - " & Err.Number, vbCritical
End If


End Function

Sub EscreveLogDelete(pArqTexto As String, pCaminhoTexto As String)

On Error GoTo TrataErro


pArqTexto = "[" & Format(Now, "dd/mm/yyyy hh:mm:ss") & "] - " & pArqTexto

If fso.FileExists(pCaminhoTexto) Then
   Set arqtxt = fso.OpenTextFile(pCaminhoTexto, 8, True)
   arqtxt.WriteLine pArqTexto
   arqtxt.WriteBlankLines 1
   arqtxt.Close

Else
    Set arqtxt = fso.CreateTextFile(pCaminhoTexto, True)
    With arqtxt
        .WriteLine pArqTexto
        .WriteBlankLines 1
        .Close
    End With
    
    Set arqtxt = Nothing
End If

Exit Sub

TrataErro:
            glLog.EscreveLog "[Escreve] - Erro na função do módulo fso com erro: " & Err.Description
            Exit Sub
            

End Sub



Sub CopiarArquivoTexto(pCaminhoOrigem As String, pCaminhoDestino As String)

On Error GoTo TrataErro

fso.CopyFile pCaminhoOrigem, pCaminhoDestino, True


Exit Sub

TrataErro:
If Err.Number = 53 Then
  MsgBox "Arquivo <<" & Text4.Text & ">> não encontrado !", vbCritical
Else
  MsgBox Err.Description & " - " & Err.Number, vbCritical
End If
End Sub









