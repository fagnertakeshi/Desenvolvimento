Attribute VB_Name = "GlobalVar"
Option Explicit

Global Entrada(40) As String
Global NumCampos As Long
Global Dado(40) As String

Global ArquivoEnt As String
Global ArquivoOut As String
Global FormatoDocMarcaDagua As String
Global NomeTopico As String
Global NumTopico As Integer
Global ArquivoF As String
Global EnderecoWeb As String
Global TamanhoPapel As Long
Global Orientacao As Long
Global MarcadAguaPDFTopico As Long
Global ModeloMarcadAguaPDFTopico As String
Global strPastaTemp As String
Global strPastaVirtualTemp As String
Global MD As Double
Global MEE As Double
Global MS As Double
Global MI As Double
Global glCriptografia245 As TCriptografiaLab245
Global glConstantes245 As TConstantesLab245
Global glRegistro As New TRegistro
Global Inicio As String 'Variável de compatibilidade do módulo CriptografiaPlusDrag

Global glBiblioteca245 As New TBiblioteca245
Global Const gImpressorasPDF = "PDF995#PDF_995"
'Global Const gImpressorasPDF = "PDF995#ADOBE#ADOBEPDF#ACROBAT#PDF"
Global gNomeArqIndices As String
