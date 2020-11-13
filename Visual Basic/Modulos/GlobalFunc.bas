Attribute VB_Name = "GlobalFunc"
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Declare Sub IniciaTiffHandler Lib "TiffHandler" ()
Public Declare Sub FinalizaTiffHandler Lib "TiffHandler" ()

Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private priv_lngErro As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Declare Function SetProcessWindowStation Lib "user32" (ByVal hWinSta As Long) As Long
Declare Function OpenWindowStation Lib "user32" Alias "OpenWindowStationA" (ByVal lpszWinSta As String, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Declare Function OpenDesktop Lib "user32" Alias "OpenDesktopA" (ByVal lpszDesktop As String, ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long

Public Const ERROR_FILE_NOT_FOUND = 2&
Public Const ERROR_PATH_NOT_FOUND = 3&
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_NOT_ENOUGH_MEMORY = 8 '  dderror
Public Const ERROR_SHARING_VIOLATION = 32&

Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const DESKTOP_WRITEOBJECTS = &H80&
Public Const DESKTOP_READOBJECTS = &H1&

Private Entrada(21) As String

'Global glLog As New TLog

Function Estrutura_Dados(Optional ByRef pMsgErro As String) As Long

Dim vErro As Long

On Error GoTo DizErro

    Estrutura_Dados = 0
    pMsgErro = ""

    Entrada(0) = "Topico"
    Entrada(1) = "Arquivos"
    Entrada(2) = "NumTopico"
    Entrada(3) = "Indices"
    Entrada(4) = "Visualizar"
    Entrada(5) = "MargemEsq"
    Entrada(6) = "MargemDir"
    Entrada(7) = "MargemSup"
    Entrada(8) = "MargemInf"
    Entrada(9) = "Copias"
    Entrada(10) = "Unidade"
    Entrada(11) = "Paginas"
    Entrada(12) = "Orientacao"
    Entrada(13) = "Tamanho"
    Entrada(14) = "Login"
    Entrada(15) = "Senha"
    Entrada(16) = "Chave"
    Entrada(17) = "FormatoDocMarcaDagua"
    Entrada(18) = "JuntarPDF"
    Entrada(19) = "ExecutandoComoCGI"
    Entrada(20) = "GerarXML"
    Entrada(21) = "ValoresCriptografados"
    
    NumCampos = 21

Fim:
    Exit Function
    
DizErro:
    vErro = Err.Number
    Estrutura_Dados = vErro
    pMsgErro = Err.Description
    glLog.EscreveLog "[Estrutura_Dados] - Erro inesperado: " & pMsgErro & " (" & vErro & ")", True
    Resume Fim

End Function

Public Function DefineCamTemp(strDir As String, strCaminho As String, arqTemp As String, Optional strExtensao As String = "") As Long

    On Error GoTo DizErro

    Dim strDirTemp2 As String * 255
    Dim lngTamDirTemp As Long
    Dim strDirTemp As String
    Dim strArqTemp As String * 255
    Dim lngNumDisponivel As Long
    Dim Passo As Integer

    DefineCamTemp = 0
    lngTamDirTemp = GetTempPath(255, strDirTemp2)
    strDirTemp2 = "C:\"
    strDirTemp = Left(strDirTemp2, lngTamDirTemp - 1)
    strDirTemp = strDir & "\"
    strArqTemp = String(255, Chr(11))

Passo = 0
    
    lngTamDirTemp = GetTempFileName(strDirTemp, "PDF", 0, strArqTemp)
Passo = 1
    strCaminho = Left(strArqTemp, InStr(1, strArqTemp, Chr(11)) - 2)

Passo = 2
    If strExtensao <> "" Then
        Kill strCaminho

Passo = 3
        strCaminho = Left(strCaminho, Len(strCaminho) - 3) & strExtensao
        lngNumDisponivel = FreeFile
Passo = 4
        Open strCaminho For Binary Access Write As #lngNumDisponivel
        Close #lngNumDisponivel

       Kill strCaminho
    End If

Passo = 5
    arqTemp = Trim(Right(strCaminho, Len(strCaminho) - InStrRev(strCaminho, "\")))

    glLog.EscreveLog "[DefineCamTemp] Arquivo Temporário: " & arqTemp

Fim:
    Exit Function

DizErro:
    DefineCamTemp = Err.Number
    glLog.EscreveLog "[DefineCamTemp] Erro: " & Err.Description & ". Passo: " & Passo, True
    Resume Fim

End Function
