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

'mostrando o conteúdo do arquivo
LerArquivoTexto = texto

arq2.Close
Exit Function

TrataErro:
If Err.Number = 53 Then
   gllog.EscreveLog "Arquivo <<" & pCaminhoTexto & ">> não encontrado !", vbCritical
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
  gllog.EscreveLog "Arquivo <<" & Text4.Text & ">> não encontrado !", vbCritical
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
    MsgBox "Diretório já existe"

End If

TrataErro:
    gllog.EscreveLog "[CriarDiretorio] -  Não foi possível criar o diretório" & pCaminho
    Exit Sub
    

End Sub



Sub ApagaDiretorio(pCaminho As String)

On Error GoTo TrataErro

fso.DeleteFolder pCaminho

TrataErro:
    gllog.EscreveLog "[ApagaDiretorio] - Não foi possível apagar o diretório"
    Exit Sub
    
End Sub


Function ArquivoExisteFSO(pCaminho As String) As Boolean

On Error GoTo TrataErro

ArquivoExisteFSO = fso.FileExists(pCaminho)

TrataErro:
    gllog.EscreveLog "[ArquivoExiste]- Não foi possível verificar a existência do arquivo"
    Exit Function
    

End Function


Sub MoverDiretorio(pCaminhoOrigem As String, pCaminhoDestino As String)

On Error GoTo TrataErro

fso.MoveFolder pCaminhoOrigem, pCaminhoDestino

TrataErro:
    gllog.EscreveLog "[MoverDiretorio]- Não foi possível mover o diretório"
    Exit Sub
    
End Sub



Function DataCriacaoPasta(pCaminho As String) As Date


Dim f As Scripting.folder
Set f = fso.GetFolder(pCaminho)


On Error GoTo TrataErro

DataCriacaoPasta = f.DateCreated

TrataErro:
    gllog.EscreveLog "[DataCriacaoPasta]- Não foi possível verificar a data da criação da pasta."
    Exit Function
    

End Function

Function DataAcessoPasta(pCaminho As String) As Date


Dim f As Scripting.folder
Set f = fso.GetFolder(pCaminho)


On Error GoTo TrataErro

DataAcessoPasta = f.DateLastAccessed

TrataErro:
    gllog.EscreveLog "[DataAcessoPasta]-Não foi possível verificar a data de acesso da pasta."
    Exit Function
    

End Function

Function DataModificacaoPasta(pCaminho As String) As Date


Dim f As Scripting.folder
Set f = fso.GetFolder(pCaminho)


On Error GoTo TrataErro

DataModificacaoPasta = f.DateLastModified

TrataErro:
    gllog.EscreveLog "[DataModificacaoPasta]- Não foi possível verificar a data de modificação da pasta."
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
    gllog.EscreveLog "[ProcuraArquivos] - Não foi possível localizar o(s) arquivo(s)."

    Exit Function

End Function




