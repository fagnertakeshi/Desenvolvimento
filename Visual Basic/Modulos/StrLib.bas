Attribute VB_Name = "StrLib"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long


Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, _
                                                                                  ByVal nFolder As Long, _
                                                                                  ByVal hToken As Long, _
                                                                                  ByVal dwFlags As Long, _
                                                                                  ByVal pszPath As String) As Long

Public Const CSIDL_WINDOWS As Long = &H24 'Windows directory Or SYSROOT()
Public Const CSIDL_PROGRAM_FILES As Long = &H26 'Program Files folder
Private Const SHGFP_Type_CURRENT = &H0 'current value For user, verify it exists
Private Const S_OK = 0
Private Const S_False = 1



Public Function ComputerName() As String
  Dim sBuffer As String
  
  Dim lAns As Long
 
  sBuffer = Space$(255)
  lAns = GetComputerName(sBuffer, 255)
  If lAns <> 0 Then
        'read from beginning of string to null-terminator
        ComputerName = Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1)
   Else
        Err.Raise Err.LastDllError, , _
          "A system call returned an error code of " _
           & Err.LastDllError
   End If

End Function


Public Sub DeletaArquivosTemp()


Dim fso As New FileSystemObject
Dim oFolder1 As Folder
Dim dirwin As Folder
Dim oFile As File
Dim ofile2 As File
Dim userProfile
Dim Windir
Dim WshShell As Object

On Error GoTo DizErro

'Set up environment
Set WshShell = CreateObject("WScript.Shell")
'Set WshShell = New WshShell
'Set fso = CreateObject("Scripting.FileSystemObject")
Set fso = New FileSystemObject


userProfile = WshShell.ExpandEnvironmentStrings("%temp%")
Windir = WshShell.ExpandEnvironmentStrings("%windir%")

'start deleting files
Set oFolder1 = fso.GetFolder(userProfile)
 For Each oFile In oFolder1.Files
        If UCase(Left(oFile.Name, 3)) = "LAB" Then
            oFile.Delete True
        Else
            If UCase(Right(oFile.Name, 3)) = "TIF" Or UCase(Right(oFile.Name, 3)) = "BMP" Or UCase(Right(oFile.Name, 3)) = "JPG" Then
                oFile.Delete True
            End If
        End If
 Next
 
Set dirwin = fso.GetFolder(Windir & "\temp")

 For Each ofile2 In dirwin.Files
        If UCase(Left(ofile2.Name, 3)) = "LAB" Then
            ofile2.Delete True
        Else
            If UCase(Right(ofile2.Name, 3)) = "TIF" Or UCase(Right(ofile2.Name, 3)) = "BMP" Or UCase(Right(ofile2.Name, 3)) = "JPG" Then
                ofile2.Delete True
            End If
        End If
 Next

Fim:
    Set oFile = Nothing
    Set ofile2 = Nothing
    Set Windir = Nothing
    Set fso = Nothing

Exit Sub

DizErro:
        'glLog.EscreveLog "[GetNomeComputer]- Erro: " & Err.Description
        Resume Fim


End Sub
Public Function GetNomeComputer(MachineName As String) As Long
    
    Dim NameSize As Long
    Dim X As Long
    Dim pos As Long
On Error GoTo DizErro:

    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
    MachineName = Trim(MachineName)
    pos = InStr(1, MachineName, vbNullChar)
    MachineName = Mid(MachineName, 1, pos - 1)

Fim:
    Exit Function
DizErro:
        'glLog.EscreveLog "[GetNomeComputer]- Erro: " & Err.Description
        Resume Fim

End Function
Public Function ConverteStringData(Data As String) As Date

Dim mes As String
Dim ano As String
Dim dia As String

On Error GoTo DizErro


If Data <> "" Then
    ano = Mid(Data, 1, 4)
    mes = Mid(Data, 5, 2)
    dia = Mid(Data, 7, 2)
    Data = ano & "/" & mes & "/" & dia
End If

ConverteStringData = CDate(Data)

Fim:
    Exit Function

DizErro:
    'glLog.EscreveLog "[ConverteDataString]- Erro na conversão da data"
    Resume Fim
End Function


Public Function GetSystemDrive() As String

On Error GoTo DizErro

Dim vWindowDir As String

    vWindowDir = GetSpecialFolderPath(CSIDL_WINDOWS)
    GetSystemDrive = Mid(vWindowDir, 1, InStr(vWindowDir, "\"))

Fim:
    Exit Function

DizErro:
    GetSystemDrive = ""
    Resume Fim

End Function

Public Function GetProgramFiles() As String

On Error GoTo DizErro

    GetProgramFiles = GetSpecialFolderPath(CSIDL_PROGRAM_FILES)

Fim:
    Exit Function

DizErro:
    GetProgramFiles = ""
    Resume Fim

End Function


Public Function RetirarBarraInterrogacao(ByVal pFilePath As String) As String

On Error GoTo DizErro
    
    RetirarBarraInterrogacao = Trim(pFilePath)

    If Mid(pFilePath, 1, 4) = "\\?\" Then
        RetirarBarraInterrogacao = Mid(pFilePath, 5)
    End If

Fim:
    Exit Function

DizErro:
    Resume Fim

End Function

Public Function TrataLinhaDeComandoArquivos(ByVal pComando As String) As String
    
    Dim strAux As String
    On Error GoTo Erro
    
    strAux = pComando
    If InStr(1, strAux, Chr(34)) <> 0 Then
        strAux = Replace(pComando, Chr(34) & " ", Chr(31))
        strAux = Replace(strAux, Chr(34), "")
    Else
        strAux = Replace(pComando, " ", Chr(31))
    End If
    
    TrataLinhaDeComandoArquivos = strAux

Fim:
    Exit Function
Erro:
    TrataLinhaDeComandoArquivos = pComando
    Resume Fim
    
End Function

Private Function GetSpecialFolderPath(CSIDL As Long) As String

On Error GoTo DizErro

Dim strPath As String
Dim iReturn As Long
Const cMaxPath = 260

    strPath = String(cMaxPath, 0)
    iReturn = SHGetFolderPath(0, CSIDL, 0, SHGFP_Type_CURRENT, strPath)
    
    Select Case iReturn
        Case S_OK
           GetSpecialFolderPath = Left$(strPath, InStr(1, strPath, Chr(0)) - 1)
        Case S_False
           GetSpecialFolderPath = "Pasta não existe."
        Case Else
           GetSpecialFolderPath = "Pasta inválida para este sistema operacional."
    End Select
    
Fim:
    Exit Function
    
DizErro:
    alerta "Erro na função GetSpecialFolderPath - " & Err.Description & " (" & Err.Number & ")"
    Resume Fim
    
End Function

Public Function TemBarraFinal(ByVal pStr As String) As Boolean

Dim vLen As Byte

On Error GoTo Erro
    
    If Not StringVazia(pStr) Then
        If Len(pStr) > 1 Then
            vLen = 1
        Else
            vLen = 0
        End If
        TemBarraFinal = Mid(pStr, Len(pStr), Len(pStr) - vLen) = "/"
    End If
    
    Exit Function
    
Erro:
    alerta "Erro na função: TemBarraFinal - " & Err.Description & " (" & Err.Number & ")"
    Exit Function

End Function

Public Function AddBarraFinal(ByVal pStr As String) As String

On Error GoTo Erro

    AddBarraFinal = pStr

    If StringVazia(AddBarraFinal) Then
        AddBarraFinal = "/"
    Else
        If Not TemBarraFinal(AddBarraFinal) Then AddBarraFinal = AddBarraFinal & "/"
    End If
    
    Exit Function
    
Erro:
    If Err.Number <> 0 Then
        alerta "Erro na função: AddBarraFinal - " & Err.Description & " (" & Err.Number & ")"
    End If
    Exit Function

End Function

Public Function TemContraBarraFinal(ByVal pStr As String) As Boolean

Dim vLen As Byte

On Error GoTo Erro
    
    If Not StringVazia(pStr) Then
        If Len(pStr) > 1 Then
            vLen = 1
        Else
            vLen = 0
        End If
        TemContraBarraFinal = Mid(pStr, Len(pStr), Len(pStr) - 1) = "\"
    End If
    
    Exit Function

Erro:
    alerta "Erro na função: AddBarraFinal - " & Err.Description & " (" & Err.Number & ")"
    Exit Function


End Function

Public Function AddContraBarraFinal(ByVal pStr As String) As String

On Error GoTo Erro

    AddContraBarraFinal = pStr

    If StringVazia(AddContraBarraFinal) Then
        AddContraBarraFinal = "\"
    Else
        If Not TemContraBarraFinal(pStr) Then AddContraBarraFinal = AddContraBarraFinal & "\"
    End If
    
    Exit Function
    
Erro:
    If Err.Number <> 0 Then
        alerta "Erro na função: AddContraBarraFinal - " & Err.Description & " (" & Err.Number & ")"
    End If
    Exit Function

End Function


Public Function StringVazia(ByVal pStr As String) As Boolean
    
    StringVazia = (Trim(pStr) = "")
    
End Function


Public Function CriarPasta(ByVal pCaminhoPasta As String, Optional pMsgErro As Boolean = False) As Long

On Error GoTo Erro

    CriarPasta = 0
   
    If Not PastaExiste(pCaminhoPasta) Then MkDir pCaminhoPasta
    
    Exit Function
    
Erro:
    If Err.Number <> 0 Then
        CriarPasta = Err.Number
        If pMsgErro Then alerta "Erro ao criar a pasta: " & Err.Description & " (" & Err.Number & ")"
    End If
    Exit Function
    
End Function

Public Function Confirmacao(ByVal pStr As String, Optional pCaption As String = "") As Boolean

    If StringVazia(pCaption) Then pCaption = app.Title
    
    Confirmacao = MsgBox(pStr, vbYesNo + vbQuestion, pCaption) = vbYes

End Function

Public Sub Mensagem(ByVal pStr As String, Optional pCaption As String = "")

    If StringVazia(pCaption) Then pCaption = app.Title

    MsgBox pStr, vbExclamation + vbOKOnly, pCaption

End Sub



Public Sub alerta(ByVal pStr As String, Optional pCaption As String = "")

On Error GoTo Erro

    If StringVazia(pCaption) Then pCaption = app.Title

    MsgBox pStr, vbCritical + vbOKOnly, pCaption
    
    
    Exit Sub
    
Erro:
     'MsgBox "Erro na função: Alerta - " & Err.Description & " (" & Err.Number & ")", vbCritical + vbOKOnly, "Erro"

End Sub

Public Function DesejaSair() As Boolean

    DesejaSair = Confirmacao("Deseja sair da aplicação?")
    
End Function

Public Function ExtrairNomeDoArquivo(pArquivo As String) As String

    ExtrairNomeDoArquivo = Right(pArquivo, Len(pArquivo) - InStrRev(pArquivo, "\"))

End Function

Public Function ExtrairCaminhoDoArquivoComBarra(ByVal pArquivo As String) As String

    ExtrairCaminhoDoArquivoComBarra = Left(pArquivo, InStrRev(pArquivo, "\"))

End Function

' kann 25-04-2007
Public Function ExtencaoArquivo(ByVal pArquivo As String, Optional ByRef pPos As Integer) As String
Dim vPos As Integer
On Error GoTo Erro
    vPos = InStr(pArquivo, ".")
    If vPos = 0 Then
        ExtencaoArquivo = ""
    Else
        pPos = vPos ' posição do ponto da direita para a esquerda
        ExtencaoArquivo = Mid(pArquivo, vPos + 1)
    End If
    Exit Function
Erro:
    alerta "Erro na função: ExtencaoArquivo - " & Err.Description & " (" & Err.Number & ")"
End Function

' 'kann 25 - 4 - 2007
'Public Function MontaNomeArquivo(ByVal pArq As String) As String
'
'Dim vTemp As String
'Dim vPos As Integer
'On Error GoTo Erro
'
'    vTemp = ExtencaoArquivo(pArq, vPos)
'    If vPos = 0 Then
'        MontaNomeArquivo = pArq & gHoje_expira
'    Else
'        MontaNomeArquivo = Left(pArq, vPos - 1) & gHoje_expira & "." & vTemp
'    End If
'    Exit Function
'
'Erro:
'    alerta "Erro na função: ExtencaoArquivo - " & Err.Description & " (" & Err.Number & ")"
'
'End Function


Public Function ExtrairCaminhoDoArquivoSemBarra(ByVal pArquivo As String) As String

    ExtrairCaminhoDoArquivoSemBarra = Left(pArquivo, InStrRev(pArquivo, "\") - 1)

End Function

Public Function CopiarArquivo(pOrigem, pDestino As String, _
                              Optional pReescreve As Boolean = False, _
                              Optional pAlerta As Boolean = False) As Long

Dim vStr As String

On Error GoTo Erro
    
    If pOrigem <> pDestino Then
    
        vStr = Dir(pDestino, vbArchive + vbHidden)
        
        If Not (StringVazia(vStr) Or pReescreve) Then
            If Confirmacao("Arquivo " & Chr(34) & pDestino & Chr(34) & " já existe. " & vbCrLf & _
                            "Deseja sobrepor o arquivo?") Then
                FileCopy pOrigem, pDestino
            End If
        Else
            FileCopy pOrigem, pDestino
        End If
        
    End If
    
Fim:
    Exit Function
    
Erro:
    If pAlerta Then alerta "Erro na função: CopiarArquivo - " & Err.Description & " (" & Err.Number & ")"
    CopiarArquivo = Err.Number
    Resume Fim

End Function
 
Public Function ArquivoExiste(ByVal pNomeArq As String) As Boolean
    
    ArquivoExiste = False
    
    If Not StringVazia(pNomeArq) Then ArquivoExiste = Dir(pNomeArq) <> ""

End Function

Public Function PastaExiste(ByVal pNomePasta As String) As Boolean

On Error GoTo Erro1
    
    PastaExiste = False
    
    If Not StringVazia(pNomePasta) Then
        PastaExiste = Dir(pNomePasta, vbDirectory) <> ""
        
On Error GoTo Erro1

        'Se não existir tentar procurar com a barra no final
        If Not PastaExiste Then
            pNomePasta = AddContraBarraFinal(pNomePasta)
            PastaExiste = Dir(pNomePasta, vbDirectory) <> ""
        End If
        
    End If
    
    Exit Function

Erro1:
    If Err.Number = 52 Then
        Resume Next
    Else
        GoTo Erro
    End If
    
Erro:
    If Err.Number <> 0 Then
        alerta "Erro na função: PastaExiste - " & Err.Description & " (" & Err.Number & ")"
    End If

End Function


' kann 10-04-2007
Public Function validaNumero(ByVal KeyAscii As Integer) As Integer

Dim conj As String
    
On Error GoTo Erro

    conj = "0123456789" & Chr(8)
    
    If InStr(conj, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    
    validaNumero = KeyAscii
    
Exit Function

Erro:
    alerta "Erro na função: validaNumero - " & Err.Description & " (" & Err.Number & ")"
    
End Function


' kann 10-04-2007
Public Function validaTextBoxNum(ByVal vTextBox As TextBox) As Boolean

Dim Ok As Boolean

On Error GoTo Erro

    Ok = IsNumeric(vTextBox.Text)
    
    If Not Ok Then
        Mensagem "Favor preencher o campo (" & vTextBox.ToolTipText & ") com valores numéricos"
    End If
    
    validaTextBoxNum = Ok
    
    'vTextBox.SetFocus
    
Fim:
    
    Exit Function

Erro:
    
    alerta "Erro na função: validaTextBoxNum - " & Err.Description & " (" & Err.Number & ")"
    Resume Fim
    
End Function


' kann 20-04-2007
Public Function deletaArquivo(pPathCompleto As String, Optional pSemAlerta As Boolean) As Long

On Error GoTo Erro

    deletaArquivo = 0

    Kill pPathCompleto
    
Exit Function

Erro:
    If Not pSemAlerta Then alerta "Erro na função: deletaArquivo - " & Err.Description & " (" & Err.Number & ")"
    deletaArquivo = Err.Number
    
End Function

Public Function deletaPasta(pPasta As String, Optional pSemAlerta As Boolean) As Long

On Error GoTo Erro
    
    deletaPasta = 0
    
     If PastaExiste(pPasta) Then RmDir pPasta
    
Exit Function

Erro:
    deletaPasta = Err.Number
    If Not pSemAlerta Then alerta "Erro na função: deletaPasta - " & Err.Description & " (" & Err.Number & ")"
    
End Function

Public Function validaPastaServidor(ByRef pPasta As String, Optional pSemAlerta As Boolean) As String

Dim vValido As Boolean
Dim vTam As Integer, vIndice As Integer
Dim vTemp As String

On Error GoTo Erro

    validaPastaServidor = "\\"
    If Dir(pPasta, vbDirectory) <> "" Then
        validaPastaServidor = pPasta
    Else
        Exit Function
    End If
    
    vTam = Len(pPasta)
    ' pasta servidor inicia com \\
    vIndice = 3
    vTemp = ""
    While vIndice <= vTam
        If Mid(pPasta, vIndice, 1) = "\" Then
            If Mid(pPasta, vIndice - 1, 1) <> "\" Then
                vTemp = vTemp + Mid(pPasta, vIndice, 1)
            End If
        Else
            vTemp = vTemp + Mid(pPasta, vIndice, 1)
        End If
        vIndice = vIndice + 1
    Wend
    pPasta = "\\" & vTemp
    validaPastaServidor = pPasta
Exit Function

Erro:
    If Err.Number = 52 Then
        validaPastaServidor = "SemContraBarraNoFinal"
        Exit Function
    Else
        If Not pSemAlerta Then alerta "Erro na função: validaPastaServidor - " & Err.Description & " (" & Err.Number & ")"
    End If
    
End Function

Public Function limpaString(ByVal strTexto As String) As String

    On Error GoTo DizErro
    limpaString = Replace(strTexto, "/", "-")
    limpaString = Replace(limpaString, "\", "-")
    limpaString = Replace(limpaString, "|", "-")
    limpaString = Replace(limpaString, "?", "-")
    limpaString = Replace(limpaString, "<", "-")
    limpaString = Replace(limpaString, ">", "-")
    limpaString = Replace(limpaString, "*", "-")
    limpaString = Replace(limpaString, ":", "-")
    limpaString = Replace(limpaString, Chr(34), "-")

Fim:
    Exit Function
    
DizErro:
    Resume Fim
    
End Function

Public Function AlterarExtensaoArquivo(ByVal pNomeArq As String, ByVal pExtensao As String) As String

On Error GoTo DizErro
    
    pExtensao = Replace(pExtensao, ".", "")
    AlterarExtensaoArquivo = RetirarExtensaoArquivo(pNomeArq) & "." & pExtensao

Fim:
    Exit Function

DizErro:
    AlterarExtensaoArquivo = ""
    Resume Fim

End Function

Public Function RetirarExtensaoArquivo(ByVal pNomeArq As String) As String

On Error GoTo DizErro

Dim vPos As Long

    vPos = InStrRev(pNomeArq, ".")
    If vPos > 0 Then
        RetirarExtensaoArquivo = Mid(pNomeArq, 1, vPos - 1)
    Else
        RetirarExtensaoArquivo = pNomeArq
    End If
Fim:
    Exit Function

DizErro:
    RetirarExtensaoArquivo = ""
    Resume Fim

End Function

Public Function RenomearArquivo(ByVal pCaminhoENomeArq As String, _
                                ByVal pNovoNome As String) As String

On Error GoTo DizErro

    Dim vNovoNome As String
    Dim vDiretorio As String
    Dim vStr As String

    RenomearArquivo = ""
    
    vStr = Dir(pCaminhoENomeArq)
    If vStr = "" Then Exit Function
    
    vDiretorio = ExtrairCaminhoDoArquivoComBarra(pCaminhoENomeArq)
    vNovoNome = ExtrairNomeDoArquivo(pNovoNome)
    vNovoNome = vDiretorio & vNovoNome
    
    Name pCaminhoENomeArq As vNovoNome
    RenomearArquivo = vNovoNome
    

Fim:
    Exit Function

DizErro:
    RenomearArquivo = ""
    Resume Fim

End Function

Function ExtrairExtensaoArquivo(ByVal pNomeArq As String, _
                                Optional pComPonto As Boolean = True) As String
    
On Error GoTo DizErro

Dim vPos As Long

    ExtrairExtensaoArquivo = ""
    vPos = InStrRev(pNomeArq, ".")
    If vPos > 0 Then
        ExtrairExtensaoArquivo = Mid(pNomeArq, vPos, Len(pNomeArq))
        If Not pComPonto Then
            ExtrairExtensaoArquivo = Mid(ExtrairExtensaoArquivo, 2, Len(ExtrairExtensaoArquivo))
        End If
    End If
Fim:
    Exit Function

DizErro:
    ExtrairExtensaoArquivo = ""
    Resume Fim

End Function

Function AddTextoNoNomeDoArquivo(ByVal pNomeArq As String, _
                                 ByVal pTexto As String) As String
    
On Error GoTo DizErro

Dim vPos As Long
Dim vExtensao As String
Dim vDiretorio As String
Dim NomeArquivo As String

    AddTextoNoNomeDoArquivo = pNomeArq

    vDiretorio = ExtrairCaminhoDoArquivoComBarra(pNomeArq)
    
    NomeArquivo = RetirarExtensaoArquivo(ExtrairNomeDoArquivo(pNomeArq))
    NomeArquivo = NomeArquivo & pTexto
    
    vExtensao = ExtrairExtensaoArquivo(pNomeArq)
    
    AddTextoNoNomeDoArquivo = vDiretorio & NomeArquivo & vExtensao

Fim:
    Exit Function

DizErro:
    Resume Fim
    
End Function


Function RetirarChrFinalString(ByVal pStr As String, ByVal pChr As String) As String

On Error GoTo DizErro

    RetirarChrFinalString = pStr

    While Right(RetirarChrFinalString, 1) = pChr
        RetirarChrFinalString = Left(RetirarChrFinalString, Len(RetirarChrFinalString) - 1)
    Wend
    
Fim:
    Exit Function

DizErro:
    Resume Fim

End Function

Function RetirarChrInicioString(ByVal pStr As String, ByVal pChr As String) As String

On Error GoTo DizErro

    RetirarChrInicioString = pStr

    While Left(RetirarChrInicioString, 1) = pChr
        RetirarChrInicioString = Right(RetirarChrInicioString, Len(RetirarChrInicioString) - 1)
    Wend
    
Fim:
    Exit Function

DizErro:
    Resume Fim

End Function

Function BuscaStringChrSeparador(ByVal pStr As String, _
                                 ByVal pChrSep As String, _
                                 ByVal pStrBusca As String) As Long
                                  
On Error GoTo DizErro

    Dim aVetorTemp() As String
    Dim vTamVetor As Long
    Dim vInd As Long
    
    BuscaStringChrSeparador = -1
    
    aVetorTemp = Split(pStr, pChrSep)
    vTamVetor = UBound(aVetorTemp)
    
    For vInd = 0 To vTamVetor
        If Trim(UCase(aVetorTemp(vInd))) = Trim(UCase(pStrBusca)) Then
            BuscaStringChrSeparador = vInd
            Exit For
        End If
    Next vInd

Fim:
    Exit Function

DizErro:
    BuscaStringChrSeparador = Err.Number
    Resume Fim

End Function

'Heliomar kann
Public Function RetiraAcento(ByVal pStr As String) As String
    
Dim vTemp As String
    
    vTemp = UCase(pStr)
    
    vTemp = Replace(vTemp, "Á", "A")
    vTemp = Replace(vTemp, "À", "A")
    vTemp = Replace(vTemp, "Ã", "A")
    vTemp = Replace(vTemp, "Â", "A")
    
    vTemp = Replace(vTemp, "É", "E")
    vTemp = Replace(vTemp, "Ê", "E")
    vTemp = Replace(vTemp, "È", "E")
    
    vTemp = Replace(vTemp, "Í", "I")
    vTemp = Replace(vTemp, "Ì", "I")
    vTemp = Replace(vTemp, "Î", "I")
    
    vTemp = Replace(vTemp, "Ò", "O")
    vTemp = Replace(vTemp, "Ó", "O")
    vTemp = Replace(vTemp, "Õ", "O")
    vTemp = Replace(vTemp, "Ô", "O")
    
    vTemp = Replace(vTemp, "Ú", "U")
    vTemp = Replace(vTemp, "Ù", "U")
    vTemp = Replace(vTemp, "Û", "U")
    
    RetiraAcento = vTemp
    
End Function
Public Function RetirarRepetidosVetString(pVet() As String) As Long
    
    Dim vArrAux() As String
    Dim vIndI As Long
    Dim vIndJ As Long
    Dim vBoolExiste As Boolean
    
    On Error GoTo DizErro
    
    RetirarRepetidosVetString = 0
    
    ReDim vArrAux(0)
    
    For vIndI = 0 To UBound(pVet)
        If vIndI > 0 Then
            For vIndJ = 0 To UBound(vArrAux) - 1
                vBoolExiste = (pVet(vIndI) = vArrAux(vIndJ))
                If vBoolExiste Then Exit For
            Next vIndJ
        Else
            vBoolExiste = False
        End If
        
        If Not vBoolExiste Then
            vArrAux(UBound(vArrAux)) = pVet(vIndI)
            
            If Not vIndI = UBound(pVet) Then
                ReDim Preserve vArrAux(UBound(vArrAux) + 1)
            End If
        End If
        
    Next vIndI
    
    ReDim pVet(0)
    pVet = vArrAux
    
Fim:
    Exit Function
DizErro:
    RetirarRepetidosVetString = Err.Number
    Resume Fim
End Function

'Heliomar kann
Public Function OrdenaVetString(pVet() As String) As Long

On Error GoTo DizErro
    
    Dim vIni As Integer
    Dim vI As Integer
    Dim vJ As Integer
    Dim vFim As Integer
    Dim vPosMemor As Integer
    
    
    Dim vMemor As String
    Dim vAux As String
    Dim vTroca As String
    
    OrdenaVetString = 0
    vIni = 0
    vPosMemor = 0
    vFim = UBound(pVet)
    
    For vIni = 0 To vFim
        vPosMemor = vIni
        vMemor = RetiraAcento(pVet(vPosMemor))
        For vJ = vIni + 1 To vFim
            vAux = RetiraAcento(pVet(vJ))
            If vAux < vMemor Then
                vMemor = vAux
                vPosMemor = vJ
            End If
        Next vJ
        
        'atualizando o vetor original
        vTroca = pVet(vIni)
        pVet(vIni) = pVet(vPosMemor)
        pVet(vPosMemor) = vTroca
        
    Next vIni
    
Fim:
    Exit Function
DizErro:
    OrdenaVetString = Err.Number
    Resume Fim
    
End Function

Public Function OrdenaIndices(pDesordenados() As String, pOrdenados() As String, pVetIndicesLong() As Long)

On Error GoTo DizErro

    ' Ordena um vetor de Indices do tipo long de acordo com
    ' a posição de um vetor ordenado em um desordenado
    ' ex.
    ' pDesordenados =   { j , x , a }  onde o j = 8, x = 3, a = 1
    ' pVetIndicesLong = { 8 , 3 , 1 }
    ' pOrdenados =      { a , j , x }
    'a Função retorna = { 1 , 8 , 3 }
    
    
    Dim vVetind() As String
    Dim vI As Long
    Dim vJ As Long
    Dim vAux As Long
    Dim vFim As Long
    Dim vRetorno() As Long
    
    vFim = UBound(pOrdenados)
    
    ReDim vRetorno(vFim)
    
    For vI = 0 To vFim - 1
        vAux = vI
        If pOrdenados(vI) <> pDesordenados(vI) Then
            For vJ = 0 To vFim - 1
                If pOrdenados(vI) = pDesordenados(vJ) Then
                    vAux = vJ
                    Exit For
                End If
            Next vJ
        End If
        vRetorno(vI) = pVetIndicesLong(vAux)
    Next vI
    
    OrdenaIndices = vRetorno

Fim:
    Exit Function
DizErro:
     alerta "Erro no método OrdenaIndices: " & Err.Description & "  " & Err.Number
End Function

Public Function Replicar(ByVal pChr As String, ByVal pQtde As Long) As String

    Dim vInd As Long
    
    Replicar = ""
    For vInd = 1 To pQtde
        Replicar = Replicar & pChr
    Next vInd

End Function

Public Function SubtituirCaseInsensitive(ByVal pTexto As String, _
                                         ByVal pSubstituido As String, _
                                         ByVal pSubstituidor As String, _
                                         Optional ByRef pErro As Long = 0) As String

Dim vPontoInicio As Long
Dim vPontoFim As Long
Dim vTextoComparativo As String
    
On Error GoTo TrataErro
    
    pSubstituido = UCase(pSubstituido)
    SubtituirCaseInsensitive = pTexto
    
    While InStr(UCase(SubtituirCaseInsensitive), pSubstituido) > 0
        vPontoInicio = InStr(UCase(SubtituirCaseInsensitive), pSubstituido)
        vPontoFim = vPontoInicio + Len(pSubstituido)
        SubtituirCaseInsensitive = Mid(SubtituirCaseInsensitive, 1, vPontoInicio - 1) & pSubstituidor & Mid(SubtituirCaseInsensitive, vPontoFim)
    Wend
    
Fim:
    Exit Function
    
TrataErro:
    pErro = Err.Number
    SubtituirCaseInsensitive = pTexto
    Resume Fim

End Function

Public Function DataFormatada() As String
    DataFormatada = Format(Now, "YYYY-MM-DD-HH-MM-SS")
End Function

Public Function GetLoginWindows() As String
    
Dim vInd As Long
    
    On Error GoTo DizErro
    
    GetLoginWindows = Space(40)
    vInd = Len(GetLoginWindows)
    GetUserName GetLoginWindows, vInd
    GetLoginWindows = Left(GetLoginWindows, vInd - 1)

Fim:
    Exit Function
    
DizErro:
    Resume Fim
    
End Function

Public Sub RetirarTag(ByRef pConteudo As String, _
                      ByVal pTagI As String, _
                      ByVal pTagF As String)
                                        
Dim vIndI As Long
Dim vIndF As Long
Dim vstrAux As String

Dim vStrIni As String
Dim vStrFim As String

On Error GoTo TrataErro
    
    pTagI = UCase(pTagI)
    pTagF = UCase(pTagF)

    vstrAux = UCase(pConteudo)
    
    vIndI = InStr(1, vstrAux, pTagI)
    If vIndI > 0 Then
        vIndF = InStr(vIndI, vstrAux, pTagF)
        
        If vIndF > 0 Then
            vStrIni = Mid(pConteudo, 1, vIndI - 1)
            vStrFim = Mid(pConteudo, vIndF + Len(pTagF), Len(vstrAux))
        End If
        pConteudo = vStrIni & vStrFim
    End If
    
Fim:
    Exit Sub
    
TrataErro:
    Resume Fim
    
End Sub

Public Sub AlterarValorEntreTags(ByRef pConteudo As String, _
                                 ByVal pTagI As String, _
                                 ByVal pTagF As String, _
                                 ByVal pValor As String)
                                        
Dim vIndI As Long
Dim vIndF As Long
Dim vstrAux As String

Dim vStrIni As String
Dim vStrFim As String

On Error GoTo TrataErro

    pTagI = UCase(pTagI)
    pTagF = UCase(pTagF)

    vstrAux = UCase(pConteudo)
    
    vIndI = InStr(1, vstrAux, pTagI)
    If vIndI > 0 Then
        vIndF = InStr(vIndI, vstrAux, pTagF)
        
        If vIndF > 0 Then
            vStrIni = Mid(pConteudo, 1, vIndI + Len(pTagI) - 1)
            vStrFim = Mid(pConteudo, vIndF, Len(vstrAux))
            pConteudo = vStrIni & pValor & vStrFim
        End If
    
    End If
    
Fim:
    Exit Sub
    
TrataErro:
    Resume Fim
 

End Sub

Public Function ValorEntreTags(ByVal pValor As String, _
                               ByVal pTagI As String, _
                               ByVal pTagF As String) As String
                                        
Dim vIndI As Long
Dim vIndF As Long
Dim vstrAux As String

On Error GoTo TrataErro

    ValorEntreTags = ""

    pTagI = UCase(pTagI)
    pTagF = UCase(pTagF)

    vstrAux = UCase(pValor)
    
    vIndI = InStr(1, vstrAux, pTagI)
    If vIndI > 0 Then
        vIndI = vIndI + Len(pTagI)
        
        vIndF = InStr(vIndI, vstrAux, pTagF)
        If vIndF > 0 Then
            ValorEntreTags = Mid(pValor, vIndI, vIndF - vIndI)
        End If
    
    End If
    
Fim:
    Exit Function
    
TrataErro:
    Resume Fim

End Function
