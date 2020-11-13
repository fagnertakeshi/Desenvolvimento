Attribute VB_Name = "Modfso"
Dim fso As New FileSystemObject
Dim arqtxt As TextStream
Dim arq1 As TextStream
Dim arq3 As TextStream
Dim arq2 As TextStream
Dim texto As String

Public Function CriarArquivoTexto(pCaminhoTexto As String, pArqTexto As String, Optional pmsg As String) As Long

On Error GoTo TrataErro


Set arqtxt = fso.CreateTextFile(pCaminhoTexto, True)
With arqtxt
    .Write pArqTexto
    .Close
End With

CriarArquivoTexto = 0

Exit Function

TrataErro:
            gllog.EscreveLog "[CriarArquivoTexto] - Erro:" & Err.Description
            CriarArquivoTexto = -1
            Exit Function
            
End Function


Function LerArquivoTexto(pCaminhoTexto As String) As String


On Error GoTo TrataErro
Set arq2 = fso.OpenTextFile(pCaminhoTexto, ForReading, True)
texto = arq2.ReadAll

'mostrando o conte�do do arquivo
LerArquivoTexto = texto

arq2.Close
Exit Function

TrataErro:
If Err.Number = 53 Then
   gllog.EscreveLog "Arquivo <<" & pCaminhoTexto & ">> n�o encontrado !", vbCritical
Else
   gllog.EscreveLog " - " & Err.Number, vbCritical
End If


End Function



Sub CopiarArquivoTexto(pCaminhoOrigem As String, pCaminhoDestino As String)

On Error GoTo TrataErro

fso.CopyFile pCaminhoOrigem, pCaminhoDestino, True


Exit Sub

TrataErro:
If Err.Number = 53 Then
  gllog.EscreveLog "Arquivo <<" & Text4.Text & ">> n�o encontrado !", vbCritical
Else
  gllog.EscreveLog Err.Description & " - " & Err.Number, vbCritical
End If
End Sub


Sub CopiarDiretorio(pCaminho As String, pDestino As String)

On Error GoTo TrataErro

fso.CopyFolder pCaminho, pDestino


TrataErro:
    gllog.EscreveLog "[CopiarDiretorio]- Erro ao copiar o diretorio"
    Exit Sub
    

End Sub


Sub CriarDiretorio(pCaminho As String)

On Error GoTo TrataErro

If Not (fso.FolderExists(pCaminho)) Then

    fso.CreateFolder pCaminho
    
Else
    MsgBox "Diret�rio j� existe"

End If

TrataErro:
    gllog.EscreveLog "[CriarDiretorio] -  N�o foi poss�vel criar o diret�rio" & pCaminho
    Exit Sub
    

End Sub



Sub ApagaDiretorio(pCaminho As String)

On Error GoTo TrataErro

fso.DeleteFolder pCaminho

TrataErro:
    gllog.EscreveLog "[ApagaDiretorio] - N�o foi poss�vel apagar o diret�rio"
    Exit Sub
    
End Sub


Function ArquivoExisteFSO(pCaminho As String) As Boolean

On Error GoTo TrataErro

ArquivoExisteFSO = fso.FileExists(pCaminho)

TrataErro:
    gllog.EscreveLog "[ArquivoExiste]- N�o foi poss�vel verificar a exist�ncia do arquivo"
    Exit Function
    

End Function


Sub MoverDiretorio(pCaminhoOrigem As String, pCaminhoDestino As String)

On Error GoTo TrataErro

fso.MoveFolder pCaminhoOrigem, pCaminhoDestino

TrataErro:
    gllog.EscreveLog "[MoverDiretorio]- N�o foi poss�vel mover o diret�rio"
    Exit Sub
    
End Sub



Function DataCriacaoPasta(pCaminho As String) As Date


Dim f As Scripting.folder
Set f = fso.GetFolder(pCaminho)


On Error GoTo TrataErro

DataCriacaoPasta = f.DateCreated

TrataErro:
    gllog.EscreveLog "[DataCriacaoPasta]- N�o foi poss�vel verificar a data da cria��o da pasta."
    Exit Function
    

End Function

Function DataAcessoPasta(pCaminho As String) As Date


Dim f As Scripting.folder
Set f = fso.GetFolder(pCaminho)


On Error GoTo TrataErro

DataAcessoPasta = f.DateLastAccessed

TrataErro:
    gllog.EscreveLog "[DataAcessoPasta]-N�o foi poss�vel verificar a data de acesso da pasta."
    Exit Function
    

End Function

Function DataModificacaoPasta(pCaminho As String) As Date


Dim f As Scripting.folder
Set f = fso.GetFolder(pCaminho)


On Error GoTo TrataErro

DataModificacaoPasta = f.DateLastModified

TrataErro:
    gllog.EscreveLog "[DataModificacaoPasta]- N�o foi poss�vel verificar a data de modifica��o da pasta."
    Exit Function
    
End Function
Function ProcuraArquivos(pCaminho As folder) As Long

Dim arquivo As File
Dim subdiretorio As folder
Dim diretorio As folder

On Error GoTo TrataErro

Set diretorio = fso.GetFolder(pCaminho)

For Each arquivo In diretorio.Files
    If arquivo.Name Like Text1.Text Then
        List1.AddItem arquivo.Name
        achei = True
        Contador = Contador + 1
    End If
Next

ProcuraArquivos = Contador

TrataErro:
    gllog.EscreveLog "[ProcuraArquivos] - N�o foi poss�vel localizar o(s) arquivo(s)."

    Exit Function

End Function




