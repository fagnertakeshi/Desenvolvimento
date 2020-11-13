Attribute VB_Name = "formlib"
Option Explicit

Public Function FormCarregado(ByVal pNomeForm As String) As Boolean

Dim vInd As Long
Dim vNomeForm As String
On Error GoTo TrataErro

    FormCarregado = False
    vNomeForm = UCase(pNomeForm)
    For vInd = 0 To Forms.Count - 1
        If UCase(Forms(vInd).name) = vNomeForm Then
            FormCarregado = True
            Exit For
        End If
    Next vInd
    
Fim:
    Exit Function
    
TrataErro:
    FormCarregado = False
    Resume Fim

End Function

Public Function SalvarInfoControles(ByVal pForm As Form, _
                                    Optional pNomeArquivo As String = "") As Long

Dim vControle As Control
Dim vNumFile As Long
Dim vNomeArq As String

On Error GoTo TrataErro

    SalvarInfoControles = 0
    vNumFile = FreeFile
    If pNomeArquivo <> "" Then
        vNomeArq = pNomeArquivo
    Else
        vNomeArq = AddContraBarraFinal(app.path) & pForm.name & ".TXT"
    End If
    
    Open vNomeArq For Output Access Write As #vNumFile
    Print #vNumFile, pForm.name & "=" & pForm.Caption
    For Each vControle In pForm
        Print #vNumFile, vControle.name & ".Caption=" & vControle.Caption
        Print #vNumFile, vControle.name & ".ToolTip=" & vControle.ToolTipText
    Next
    Close #vNumFile

Fim:
    Exit Function
    
TrataErro:
    If Err.Number = 438 Then
        Resume Next
    Else
        SalvarInfoControles = Err.Number
    End If
    Resume Fim
    
End Function
