Attribute VB_Name = "Global"
Option Explicit

'instancia para os objetos do aplicativo
Global FuncoesTela As New TFuncoesTela
Global FuncoesCreator As New TFuncoesCreator
Global Biblioteca As New TBiblioteca245
Global glLog As New TLog
Global glRegistro As New TRegistro
Global glGerenteUsuarioDS As TGerenteUsuario
Global glBiblioteca245 As New TBiblioteca245
Global GLConstantesLab245 As New TConstantesLab245


Global Inicio As String 'variável de compatibilidade do módulo CriptografiaPlusDrag

'bit para determinar se é modo de debug
Global Const DebugBit As Boolean = False

'bit para determinar se exibe mensagens para usuário
Global GLExibeErro As Boolean

'guarda o caminho do CDCreator
Global DirCDCreator As String

'guarda o caminho do cfw
Global dirCFW As String

'guarda o caminho do arquivo de log
Global GLCaminhoLog As String

Declare Function WNetConnectionDialog Lib "mpr.dll" (ByVal hwnd As Long, ByVal dwType As Long) As Long

Sub EscreveLog(Mensagem As String)
    
    Dim NumDisponivel As Long
    
    On Error GoTo DizErro
    
    NumDisponivel = FreeFile
    Open GLCaminhoLog For Append Access Write As #NumDisponivel
    Print #NumDisponivel, Mensagem
    Close #NumDisponivel
    
    Exit Sub
    
DizErro:
    Resume Fim
    
Fim:
    
End Sub
