Attribute VB_Name = "Module1"
Option Explicit

'ATENÇÃO!!!!!!


'base de senhas FTP criptografadas
Global dbSenhasFTP As DAO.Database

Global usrRemotoUnLog As Boolean

'Para modificação de documento já indexado
Global GLDirTemp As String


'Para DEBUG sete a variável para true, para cliente, sete variável false
'Global Const DebugBit = False

'habilitar a skin
Global Const boolHabilitaSkin As Boolean = False

Global GLControleCamposIndexacao As New TControleCamposIndexacao
Global GLCamposReindexacao As New TCamposReindexacao
Global GLConstantesLab245 As New TConstantesLab245
Global GLGerentePicklists As New TGerentePickList
Global GLContruindoPicklist As New TConstruindoPicklist
Global GLCriptografia As New TCriptografiaLab245
'Global GLMarcasDagua As New TMarcasDagua
Global GLControlaGravarPara As New TControlaGravarPara
Global GLControlaSubstituir As New TControlaSubstituir
Global GlAssinaturaDigital As New TAssinaturaDigital
Global GLIniciaRegistro As New TIniciaRegistroPlusDrag

'função para pegar serial do computador
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Global Tempo As Double

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Global Const SW_SHOWNORMAL = 1
Global Chamada As Integer
Global NomeNovoArquivo As String
Global BaseHTML As String
Global nos As node
Global numero As Integer
Global Indice As Long
Global TextoEstrutura(23) As String
Global NumeroCampos As Integer
Global EnderecoDoc As String
Global i As Integer
Global Const Nulo = 0
Global Const Estrutura = 1
Global Const Imagem = 3
Global Const ImagemLote = 5
Global Const DocumentoUniversal = 4
Global Const DragDrop = 7
Global Const ArquivosExterno = 8

'Isnard: 2008-05-12
Global Const gArquivoAudio = 9
Global Const gArquivoVideo = 10
Global Const gArquivoPadrao = 11
'Fim - Isnard: 2008-05-12

Global Tipo As Integer
Global GLTiffMulti As Boolean  'define se o documento TIFF será gerado como multipágina ou várias páginas
Global valores(23) As String
Global ArquivoExecutar As String
Global ItemSelecionado
Global EnderecoPagina As String
Global Acao As Integer
Public Const NovoDoc = 0
Public Const Modifica = 1
Public Const Topico = 2
Public Const Subtopico = 3
Public Const Drag = 4

'Após junção

Public Const InserindoLink = 6
Public Const InsMascara = 7
Global DocumentoLink As String

'Fim

Public Const Ver = 4
Public Const Sem = -1
Global DirAnterior As String
Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Global topicos As Recordset
Global Servidor As Recordset
Global Login_Windows As Recordset
Global Servidor_Login As Recordset
Global Topicos_Servidor_Login As Recordset
Global Configuracao As Database
Global RedeWeb As String
Global ModeloFrame As String
Global ModeloCabe As String
Global ModeloRoda As String
Global NomeTopico As String
Global codServidor As String
Global NoNode As Integer
'Global inicio As String
Global NomeLote As String
Global ItemIndex As Integer
Global ImagemTP As Integer
Public Const Tiff = 1
Public Const BMP = 3
Public Const JPEG = 6
Global ext As String
Global ImagemArq(999) As String
Global PaginaAtual As Integer
Global NumeroPaginas As Integer
Global Ok As Integer
Global Tabela(5) As String
Global Quality As Long
Global Repete As Integer
Global Persiste As Integer
Global TamanhoOriginal As Integer
Global cont As Integer
Global anexa As Integer
Global EnderecoLink As String
Global PaginasTiff As Integer
Global DirImporta As String
Global TipoPicklist As Integer
Global ValorCampo As String
Global ErroFlag As Integer
Global habilitado As Boolean        'variável que determina se o acesso a determinadas funções é restrito ou não
Global ControleRestricao As New TControleREstricao

'do Folder245 Plus

Global Base As Integer
Global NBaseMax  As Integer
Global Query As String
Global VerThumbnail As String
Global SemTopico As Boolean
Global CaminhoTemp As String
Global ViewerFora As Integer
Global FimIndices As Integer
Global EnderecoInicial As String
Global NumeroPagIndexTif As Integer
Global DirExporta As String
Global DiretoriodeTrabalho As String
Global TopicoJaExiste As Boolean
Global blExcluirImagemAposImportacao As Boolean

'do Folder245 Input

Global DirLote As String
Global GLLoteAtual As String
Global Gerais As Recordset
Global NImagensImport As Integer
Global ImagemImporta(10000) As String
Global Importa As Boolean
Global NImagem As Integer

'do workflow
Global Operacao As Integer
Global Const Workflow = 6
Global InserirBase As Integer

'declaração para pegar numero do registro
'Declare Function RegQueryValueExNum Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Long, lpcbData As Long) As Long        ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'Para ler no Registry
'Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long       ' Note that if you declare the lpData parameter as String, you must pass it By Value.
'Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
'Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
'Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long        ' Note that if you declare the lpData parameter as String, you must pass it By Value.

'Public Const HKEY_LOCAL_MACHINE = &H80000002
'Public Const KEY_QUERY_VALUE = &H1

'Variáveis de configuração no registry
Global RegTipoImagem As String


'Variáveis para Drag and Drop
Global ArquivoDrag As String
Global TopicoDestino As New TTopico

'Variáveis para Thumbanil
Global ImagemVista As String


Global TopicoAtual As New TTopico
Global LoteAtual As New TLote
Global DocumentoAtual As New TDocumento

'Sentinelas: objetos nulos, são usados como se para
'se demonstrar que a referência ao objeto é nula
'Já que o VB é limitado e fica difícil usar
'o nothing
Global SentinelaTopico As New TTopico
Global SentinelaDocumento As New TDocumento
Global SentinelaLote As New TLote
Global SentinelaGrupoFRHU As New TGrupoFRHU

Global ImagemImportada() As Integer
Global ImagemMarcada() As Integer

Global TipoIndexacao As Integer

'Casos de uso
Global OEnviaFTP As New TEnviaFTP
'Global GLVerificaLicenca As TVerificaLicenca
'Global GLRegistraLicenca As TRegistraLicenca
Global GLIniciaPrograma As TIniciaPrograma


Enum MSGEventoWz
    MSGSAIR = 1
    MSGANTERIOR = 2
    msgproximo = 3
    MSGFINALIZAR = 4
    msgoutro = 5
    MSGHELP1 = 6
    MSGHELP2 = 7
    MSGHELP3 = 8
    MSGOUTRO2 = 9
    MSGHELP4 = 10
End Enum

Enum EventosMarcas
    Okk = 1
    SelCorTexto = 3
    SelCorTarja = 4
    Novo = 5
    Modificar = 6
    MudaCombo = 7
End Enum

'Gerentes
Global GLGerenteTopicos As New TGerenteTopicos

'Constantes e funções para FTP
Declare Function GetProcessHeap Lib "kernel32" () As Long
Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Public Const HEAP_ZERO_MEMORY = &H8
Public Const HEAP_GENERATE_EXCEPTIONS = &H4

Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" ( _
         hpvDest As Long, hpvSource As Any, ByVal cbCopy As Long)

Public Const MAX_PATH = 260
Public Const NO_ERROR = 0
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_OFFLINE = &H1000




Public Const ERROR_NO_MORE_FILES = 18

    
Public Declare Function FtpFindFirstFile Lib "wininet.dll" Alias "FtpFindFirstFileA" _
(ByVal hFtpSession As Long, ByVal lpszSearchFile As String, _
      lpFindFileData As WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent As Long) As Long

Public Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal lpszRemoteFile As String, _
      ByVal lpszNewFile As String, ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes As Long, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal lpszLocalFile As String, _
      ByVal lpszRemoteFile As String, _
      ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias "FtpSetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByVal lpszDirectory As String) As Boolean
' Initializes an application's use of the Win32 Internet functions
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Public Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias "FtpGetCurrentDirectoryA" _
    (ByVal hFtpSession As Long, ByRef lpszDirectory As String, ByRef lpdwDirectory As Long) As Boolean


' User agent constant.
Public Const scUserAgent = "vb wininet"

' Use registry access settings.
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const FTP_TRANSFER_TYPE_ASCII = &H1
Public Const FTP_TRANSFER_TYPE_BINARY = &H2
Public Const INTERNET_FLAG_PASSIVE = &H8000000

' Opens a HTTP session for a given site.
Public Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUsername As String, ByVal sPassword As String, ByVal lService As Long, _
ByVal lFlags As Long, ByVal lContext As Long) As Long
                
Public Const ERROR_INTERNET_EXTENDED_ERROR = 12003
Public Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" ( _
    lpdwError As Long, _
    ByVal lpszBuffer As String, _
    lpdwBufferLength As Long) As Boolean

' Number of the TCP/IP port on the server to connect to.
Public Const INTERNET_DEFAULT_FTP_PORT = 21
Public Const INTERNET_DEFAULT_GOPHER_PORT = 70
Public Const INTERNET_DEFAULT_HTTP_PORT = 80
Public Const INTERNET_DEFAULT_HTTPS_PORT = 443
Public Const INTERNET_DEFAULT_SOCKS_PORT = 1080
Public Const INTERNET_FLAG_IGNORE_CERT_DATE_INVALID = &H2000
Public Const INTERNET_FLAG_IGNORE_CERT_CN_INVALID = &H1000
Public Const INTERNET_FLAG_SECURE = &H800000

Public Const INTERNET_OPTION_CONNECT_TIMEOUT = 2
Public Const INTERNET_OPTION_RECEIVE_TIMEOUT = 6
Public Const INTERNET_OPTION_SEND_TIMEOUT = 5

Public Const INTERNET_OPTION_USERNAME = 28
Public Const INTERNET_OPTION_PASSWORD = 29
Public Const INTERNET_OPTION_PROXY_USERNAME = 43
Public Const INTERNET_OPTION_PROXY_PASSWORD = 44

' Type of service to access.
Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

' Opens an HTTP request handle.
Public Declare Function HttpOpenRequest Lib "wininet.dll" Alias "HttpOpenRequestA" _
(ByVal hHttpSession As Long, ByVal sVerb As String, ByVal sObjectName As String, ByVal sVersion As String, _
ByVal sReferer As String, ByVal something As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

' Brings the data across the wire even if it locally cached.
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Public Const INTERNET_FLAG_MULTIPART = &H200000

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

' Sends the specified request to the HTTP server.
Public Declare Function HttpSendRequest Lib "wininet.dll" Alias "HttpSendRequestA" (ByVal _
hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, ByVal sOptional As _
String, ByVal lOptionalLength As Long) As Integer


' Queries for information about an HTTP request.
Public Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" _
(ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, _
ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer

' The possible values for the lInfoLevel parameter include:
Public Const HTTP_QUERY_CONTENT_TYPE = 1
Public Const HTTP_QUERY_CONTENT_LENGTH = 5
Public Const HTTP_QUERY_EXPIRES = 10
Public Const HTTP_QUERY_LAST_MODIFIED = 11
Public Const HTTP_QUERY_PRAGMA = 17
Public Const HTTP_QUERY_VERSION = 18
Public Const HTTP_QUERY_STATUS_CODE = 19
Public Const HTTP_QUERY_STATUS_TEXT = 20
Public Const HTTP_QUERY_RAW_HEADERS = 21
Public Const HTTP_QUERY_RAW_HEADERS_CRLF = 22
Public Const HTTP_QUERY_FORWARDED = 30
Public Const HTTP_QUERY_SERVER = 37
Public Const HTTP_QUERY_USER_AGENT = 39
Public Const HTTP_QUERY_SET_COOKIE = 43
Public Const HTTP_QUERY_REQUEST_METHOD = 45
Public Const HTTP_STATUS_DENIED = 401
Public Const HTTP_STATUS_PROXY_AUTH_REQ = 407

' Add this flag to the about flags to get request header.
Public Const HTTP_QUERY_FLAG_REQUEST_HEADERS = &H80000000
Public Const HTTP_QUERY_FLAG_NUMBER = &H20000000
' Reads data from a handle opened by the HttpOpenRequest function.
Public Declare Function InternetReadFile Lib "wininet.dll" _
(ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, _
lNumberOfBytesRead As Long) As Integer

Public Declare Function InternetWriteFile Lib "wininet.dll" _
        (ByVal hFile As Long, ByVal sBuffer As String, _
        ByVal lNumberOfBytesToRead As Long, _
        lNumberOfBytesRead As Long) As Integer

Public Declare Function FtpOpenFile Lib "wininet.dll" Alias _
        "FtpOpenFileA" (ByVal hFtpSession As Long, _
        ByVal sFileName As String, ByVal lAccess As Long, _
        ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function FtpDeleteFile Lib "wininet.dll" _
    Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
    ByVal lpszFileName As String) As Boolean
Public Declare Function InternetSetOption Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByVal lBufferLength As Long) As Integer
Public Declare Function InternetSetOptionStr Lib "wininet.dll" Alias "InternetSetOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByVal sBuffer As String, ByVal lBufferLength As Long) As Integer

' Closes a single Internet handle or a subtree of Internet handles.
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
(ByVal hInet As Long) As Integer

' Queries an Internet option on the specified handle
Public Declare Function InternetQueryOption Lib "wininet.dll" Alias "InternetQueryOptionA" _
(ByVal hInternet As Long, ByVal lOption As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long) As Integer

' Returns the version number of Wininet.dll.
Public Const INTERNET_OPTION_VERSION = 40

' Contains the version number of the DLL that contains the Windows Internet
' functions (Wininet.dll). This structure is used when passing the
' INTERNET_OPTION_VERSION flag to the InternetQueryOption function.
Private Type tWinInetDLLVersion
    lMajorVersion As Long
    lMinorVersion As Long
End Type

' Adds one or more HTTP request headers to the HTTP request handle.
Public Declare Function HttpAddRequestHeaders Lib "wininet.dll" Alias "HttpAddRequestHeadersA" _
(ByVal hHttpRequest As Long, ByVal sHeaders As String, ByVal lHeadersLength As Long, _
ByVal lModifiers As Long) As Integer

' Flags to modify the semantics of this function. Can be a combination of these values:

' Adds the header only if it does not already exist; otherwise, an error is returned.
Public Const HTTP_ADDREQ_FLAG_ADD_IF_NEW = &H10000000

' Adds the header if it does not exist. Used with REPLACE.
Public Const HTTP_ADDREQ_FLAG_ADD = &H20000000

' Replaces or removes a header. If the header value is empty and the header is found,
' it is removed. If not empty, the header value is replaced
Public Const HTTP_ADDREQ_FLAG_REPLACE = &H80000000



Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Declare Function InternetFindNextFile Lib "wininet.dll" Alias "InternetFindNextFileA" _
    (ByVal hFind As Long, lpvFindData As WIN32_FIND_DATA) As Long

' Variável que indica se o form de configuração de picklist
' está sendo carregado
Global ConfPickFirstTime As Boolean
Global BAseConfPick As String
Global TipoConfPick As Integer
Global NomeArquivosPick() As String

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'declarações para usar o TCriptografia245

'Base e Conexao utilizados pelo sistema
Global BaseFlow As String
Global ConexaoFlow As String

'Declare Function RegOpenKeyEx Lib "ADVAPI32.DLL" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
'Declare Function RegQueryValueEx Lib "ADVAPI32.DLL" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long        ' Note that if you declare the lpData parameter as String, you must pass it By Value.
'Declare Function RegCloseKey Lib "ADVAPI32.DLL" (ByVal hKey As Long) As Long
'Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const KEY_QUERY_VALUE = &H1
'Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not Synchronize))
Public Const STANDARD_RIGHTS_ALL = &H1F0000

Global Diretorio As String

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Public Const LOGON32_LOGON_BATCH = 4
Public Const LOGON32_LOGON_INTERACTIVE = 2
Public Const LOGON32_LOGON_SERVICE = 5
Public Const LOGON32_PROVIDER_DEFAULT = 0
Public Const LOGON32_PROVIDER_WINNT35 = 1
' Dados do Criptografia
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved when system is rebooted
Public Const SYNCHRONIZE = &H100000
Public Const KEY_SET_VALUE = &H2

Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
Public Const REG_DWORD = 4                      ' 32-bit number
Public Const REG_BINARY = 3                     ' Free form binary


'declarações das novas funções de scaneamento
Public Declare Function SelecionarScanner Lib "TwainShell" () As Boolean
Public Declare Function Scan Lib "TwainShell" (NomeArquivo As String, PaletaDeCores As Long, ExibirSetup As Boolean) As Boolean


'Luiz em 02-11-2004
'Public Declare Sub ConverteFormato Lib "TiffHandler" (ArqDestino As String, ArqOrigem As String, Pagina As Long)
Public Declare Sub ConverteFormato Lib "TiffShell" (ArqDestino As String, ArqOrigem As String, Pagina As Long)
Public Declare Sub ConverteImagemParaTiffPB Lib "TiffHandlerPB" (ArqOrigem As String, ArgPagina As Long, ArqDestino As String)



Public Declare Sub AdicionarImagem Lib "TiffHandler" (ArqDestino As String, DestPagina As Long, ArqOrigem As String, OrgnPagina As Long)
Public Declare Sub ExcluirImagem Lib "TiffHandler" (Arquivo As String, Pagina As Long)
Declare Function NumDeImagens Lib "TiffHandler" (NomeArquivo As String) As Long
Public Declare Sub SalvaPrimeiraPagina Lib "TiffHandler" (ArqDest As String, ArqOrgn As String, Pagina As Long, PaginaUnica As Boolean)
Public Declare Sub AdicionaPagina Lib "TiffHandler" (ArqOrgn As String, Pagina As Long, Finaliza As Boolean)

'Public Declare Sub SalvaNoArquivo Lib "TiffHandler" (ArqDest As String)
Public Declare Sub SalvaNoArquivo Lib "TiffHandler" (ArqDest As String, Compressao As Long)
Public Declare Sub SalvaPagina Lib "TiffHandler" (ArqOrgn As String, Pagina As Long)
Public Declare Sub FechaArquivo Lib "TiffHandler" ()
Public Declare Function TiffHandlerPegaCompressao Lib "TiffHandler" (Arquivo As String) As Long
Public Declare Function TiffHandlerPegaCor Lib "TiffHandler" (Arquivo As String) As Long

'declarações para o novo componente de exibição
Global ListaDeUs() As String
Global NumUs As Long

Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long



Global GLRegistro As New TRegistro
Global glBiblioteca245 As New TBiblioteca245
Global GLCompressaoArquivo As TCompressao
