Attribute VB_Name = "modHTML"
Option Explicit

Public Function DivMarcaLab() As String
    DivMarcaLab = "<div class=" & Chr(34) & "F245Marca" & Chr(34) & ">By <a href=" & Chr(34) & "http://www.lab245.com" & Chr(34) & " target=" & Chr(34) & "_top" & Chr(34) & ">Folder245</a></div>"
End Function

Public Function DivDica() As String
    DivDica = "<div id=" & Chr(34) & "dica" & Chr(34) & "class=" & Chr(34) & "divDica" & Chr(34) & "style=" & Chr(34) & "display:none;" & Chr(34) & "></div>"
End Function

Public Function TableQtdeDocumentos(ByVal pMsg As String) As String
    TableQtdeDocumentos = "<table width=100%><tr><td align = 'left'><font face = 'verdana' size =2>" & _
                          pMsg & "</td><td align = 'right'><input type = 'CheckBox' name = 'chkTodos' id = 'chkTodos' onclick = 'SelecionarTodos(chkTodos);'>Marcar/Desmarcar Todos</td></tr></font></table>"
End Function
