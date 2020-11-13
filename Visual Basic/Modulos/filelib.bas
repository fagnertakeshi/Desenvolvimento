Attribute VB_Name = "Filelib"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Option Explicit
'////////////////////////////////////////////////////////////////////////////
'Uses: StrLib
'////////////////////////////////////////////////////////////////////////////

Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" _
    (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, _
    lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias _
    "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock _
    As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Public Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 _
    As Any, ByVal lpString2 As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, _
    Source As Any, ByVal Length As Long)
Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersion As Long
    dwFileVersionMS As Long
    dwFileVersionLS As Long
    dwProductVersionMS As Long
    dwProductVersionLS As Long
    dwFileFlagsMask As Long
    dwFileFlags As Long
    dwFileOS As Long
    dwFileType As Long
    dwFileSubtype As Long
    dwFileDateMS As Long
    dwFileDateLS As Long
End Type


Public Function GetConteudoArquivoTexto(ByVal pNomeArquivo As String, _
                                        ByRef pConteudo As String, _
                                        Optional pVerificarArqExiste = True) As Long
                                     
On Error GoTo DizErro

    Dim vFlag As Boolean
    Dim vNumDisponivel As Long
    
    GetConteudoArquivoTexto = 0
    pConteudo = ""
    
    If pVerificarArqExiste Then
        If Not ArquivoExiste(pNomeArquivo) Then
            GetConteudoArquivoTexto = -1
            Exit Function
        End If
    End If
    
    vNumDisponivel = FreeFile
    Open pNomeArquivo For Input Access Read As #vNumDisponivel
    vFlag = True
    pConteudo = Input(LOF(vNumDisponivel), #vNumDisponivel)
    Close #vNumDisponivel

Fim:
    Exit Function
    
DizErro:
    GetConteudoArquivoTexto = Err.Number
    'Se o arquivo foi aberto, devo fecha-lo
    If vFlag Then Close #vNumDisponivel
    Resume Fim
    
End Function

Public Function HIWORD(ByVal dwValue As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(dwValue), 8)
    HIWORD = CLng("&H" & Left(hexstr, 4))
End Function
Public Function LOWORD(ByVal dwValue As Long) As Long
    Dim hexstr As String
    hexstr = Right("00000000" & Hex(dwValue), 8)
    LOWORD = CLng("&H" & Right(hexstr, 4))
End Function

Function RetornaVersao(Arquivo As String)

    Dim vffi As VS_FIXEDFILEINFO  ' version info structure
    Dim buffer() As Byte          ' buffer for version info resource
    Dim pData As Long             ' pointer to version info data
    Dim nDataLen As Long          ' length of info pointed at by pData
    Dim cpl(0 To 3) As Byte       ' buffer for code page & language
    Dim cplstr As String          ' 8-digit hex string of cpl
    Dim dispstr As String         ' string used to display version information
    Dim retval As Long            ' generic return value
    
    ' First, get the size of the version info resource.  If this function fails, then Text1
    ' identifies a file that isn't a 32-bit executable/DLL/etc.
    nDataLen = GetFileVersionInfoSize(Arquivo, pData)
    If nDataLen = 0 Then
        Debug.Print "Not a 32-bit executable!"
        Exit Function
    End If
    ' Make the buffer large enough to hold the version info resource.
    ReDim buffer(0 To nDataLen - 1) As Byte
    ' Get the version information resource.
    retval = GetFileVersionInfo(Arquivo, 0, nDataLen, buffer(0))
    
    ' Get a pointer to a structure that holds a bunch of data.
    retval = VerQueryValue(buffer(0), "\", pData, nDataLen)
    ' Copy that structure into the one we can access.
    CopyMemory vffi, ByVal pData, nDataLen
    ' Display the full version number of the file.
    dispstr = Trim(Str(HIWORD(vffi.dwFileVersionMS))) & "." & _
        Trim(Str(HIWORD(vffi.dwFileVersionLS))) & _
        Trim(Str(LOWORD(vffi.dwFileVersionMS))) & "." & _
        Format(Trim(Str(LOWORD(vffi.dwFileVersionLS))), "0000")
    RetornaVersao = dispstr

End Function

'////////////////////////////////////////////////////////////////////////
'Autor: Isnard (2007.09.16)
'OBS: É necessário verificar se o arquivo existe antes de
'     utilizar esta função.
'////////////////////////////////////////////////////////////////////////
Public Function ArquivoEmUso(ByVal pNomeArquivo As String, Optional pMaxTentativas As Integer = 0)

On Error GoTo DizErro

    Dim vFileNumber As Long
    
    vFileNumber = FreeFile
    
    If FileLen(pNomeArquivo) = 0 Then
        ArquivoEmUso = True
        GoTo Fim
    End If
    
    Open pNomeArquivo For Binary Access Write Lock Read Write As #vFileNumber
    Close #vFileNumber
    ArquivoEmUso = False

Fim:
    Exit Function
    
DizErro:
    ArquivoEmUso = True
    Close #vFileNumber
    Resume Fim

End Function

'////////////////////////////////////////////////////////////////////////
'Autor: Isnard (2007.09.16)
'Retorno da função:
'   0: O arquivo não está em uso
'   -1: Excedeu o número de tentativas (caso tenha sido passado o parâmetro "pNumTentativas")
'   <> 0 ou -1: Erro na função
'
'OBS: Esta função não funciona com arquivos abertos através
'     do bloco de notas.
'
'////////////////////////////////////////////////////////////////////////
Public Function AguardarArquivoLiberado(ByVal pNomeArquivo As String, _
                                        Optional pTempoEspera As Long = 1200, _
                                        Optional pNumTentativas As Integer = 200) As Long

Dim vNumTentativas As Integer
Dim vFlag As Boolean

On Error GoTo DizErro
    
    AguardarArquivoLiberado = 0
    vNumTentativas = 0
    vFlag = True
    
    vNumTentativas = 0
    
    While vFlag
        vFlag = ArquivoEmUso(pNomeArquivo)
        'Arquivo está em uso. Deve verificar o número de tentativas
        If vFlag Then
            If pNumTentativas > 0 Then
                vNumTentativas = vNumTentativas + 1
                If vNumTentativas > pNumTentativas Then
                    vFlag = False
                    AguardarArquivoLiberado = -1
                End If
            End If
        
            Sleep (pTempoEspera)
            DoEvents
        End If
    Wend
    
Fim:
    Exit Function
    
DizErro:
    AguardarArquivoLiberado = Err.Number
    Resume Fim
    
End Function
'
'Public Function CriarArquivoTexto(ByVal pNomeArquivo As String, _
'                                  ByVal pConteudo As String, _
'                                  ByRef pMsgErro As String) As Long
'
'Dim vErro As Long
'Dim vNumDisp As Long
'
'On Error GoTo TrataErro
'
'    CriarArquivoTexto = 0
'    pMsgErro = ""
'
'    vNumDisp = FreeFile
'    Open pNomeArquivo For Output Access Write As #vNumDisp
'    Print #vNumDisp, pConteudo
'    Close #vNumDisp
'
'Fim:
'    Exit Function
'
'TrataErro:
'    vErro = Err.Number
'    pMsgErro = Err.Description
'    CriarArquivoTexto = vErro
'    Resume Fim
'
'
'End Function


Public Function CriarArquivoBinario(ByVal pNomeArquivo As String, _
                                    ByVal pConteudo As String, _
                                    ByRef pMsgErro As String) As Long
                                  
Dim vErro As Long
Dim vNumDisp As Long

On Error GoTo TrataErro

    CriarArquivoBinario = 0
    pMsgErro = ""
    
    vNumDisp = FreeFile
    Open pNomeArquivo For Binary As #vNumDisp
    Put #vNumDisp, 1, pConteudo
    Close #vNumDisp
    
Fim:
    Exit Function

TrataErro:
    vErro = Err.Number
    pMsgErro = Err.Description
    CriarArquivoBinario = vErro
    Resume Fim
    

End Function


