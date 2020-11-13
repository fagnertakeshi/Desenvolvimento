Attribute VB_Name = "FuncoesAuxiliares"
'Isnard: 2007-06-11
Private Type TTipoCampo
    Tabela As String
    nomecampo As String
    Tipo As Long
    'Caracter As String
End Type

'Isnard: 2007-06-11
Public agTipoCampo() As TTipoCampo
'Leandro: 2007-06-11
Private glBoolDefinido As Boolean

Public Sub ResetaLog(ByVal NomeArquivo As String)
    
    Dim NumDisp As Long
    NumDisp = FreeFile
    Open NomeArquivo For Output As #NumDisp
    Close #NumDisp
    
End Sub

Public Sub Log(ByVal NomeArquivo As String, ByVal NomeRotina As String, Optional ByVal DescErro As String, Optional ByVal MsgAdic As String)
    Dim NumDisp As Long
    NumDisp = FreeFile
    Open NomeArquivo For Append As #NumDisp
        Print #NumDisp, "Rotina / Hora: " & NomeRotina & " / " & Time$
        If Not (IsMissing(DescErro) Or (DescErro = "") Or IsNull(DescErro)) Then
            Print #NumDisp, vbTab & "Erro: " & DescErro
        End If
        If Not (IsMissing(MsgAdic) Or (MsgAdic = "") Or IsNull(MsgAdic)) Then
            Print #NumDisp, vbTab & MsgAdic
        End If
    Close #NumDisp
End Sub

Function DeterminaFROM(Query As String)

    ' Tenho que pegar o que vem depois do FROM
    ' Antes, vou procurando SELECTs que possam estar após o primeiro SELECT
    ' Ex.: SELECT Campo1, (SELECT Campo2 FROM Tabela2) as Campo2 FROM Tabela1
    
    Dim expr As String
    Dim Ok As Boolean
    Dim possel As Long
    Dim posfrom As Long
    Dim posini As Long
    
    expr = Trim(Query)
    
    ' Pegando o que tem depois do select
    expr = Mid(Query, 8)
    
    posini = 1
    Ok = False
    While Not Ok
        possel = InStr(posini, UCase(expr), "SELECT ")
        posfrom = InStr(posini, UCase(expr), " FROM ")
        ' Se possel = 0, basta pegar o que está depois do FROM
        ' Se possel > posfrom, basta pegar o que está depois do FROM
        If (possel > posfrom Or possel = 0) Then
            Ok = True
        Else
            ' Se possel < posfrom, tenho que achar o próximo FROM e contar a partir dele
            posini = InStr(possel, UCase(expr), " FROM ") + 4
        End If
    Wend
    DeterminaFROM = UCase(Mid(UCase(expr), posfrom + 1))

End Function


Sub DefineBase(NomeTopico As String, ByRef NomeBase As String, ByRef nometabela As String, CandBase As String, CandTabela As String)
    
    'CandBase = valor de Dado(i) correspondente ao parametro "Base"
    'CandTabela = valor de Dado(i) correspondente ao parametro "Tabela"
    Dim RecTopico As DAO.Recordset
    Dim BaseTopico As DAO.Database
    Dim QueryTopico As String
    Dim Erro As Long
    Dim Constantes As TConstantesLab245
    Dim Criptografia As TCriptografiaLab245
    Dim diret As String
    Dim existeFuncPesquisa As Boolean

  
    Set Constantes = New TConstantesLab245
    Set Criptografia = New TCriptografiaLab245
    
    'adicionei essa variável pq não existia!!!!!
    Dim ChaveGLOBAL As String
    
    On Error GoTo Fim
        
    If NomeTopico <> "" Then
        Call Criptografia.CarregaSenhaLogin(Login, Senha, True, NomeTopico, Constantes.Folder245Database, ChaveGLOBAL)
        diret = Criptografia.DefineDiretorio
        If diret <> "" Then diret = diret & "\"
        glLog.EscreveLog "[DefineBase]: A base está em " & diret & "folder.cfc"
        On Error GoTo ErroBase:
        
Continua:
        Set BaseTopico = OpenDatabase(diret & "folder.cfc", False, False, "")
     
       
        Erro = 0
        existeFuncPesquisa = True
        QueryTopico = "SELECT Base,Tabela,FuncaoPesquisa FROM Topicos WHERE Nome = '" & NomeTopico & "'"
        On Error GoTo erroFuncaoPesquisa
continuaFuncaoPesquisa:
        Set RecTopico = BaseTopico.OpenRecordset(QueryTopico, dbOpenDynaset)
        If Not RecTopico.EOF Then
            If Not IsNull(RecTopico("Base")) Then
                If Trim(RecTopico("Base")) <> "" Then
                    NomeBase = RecTopico("Base")
                Else
                    NomeBase = CandBase
                End If
            Else
                NomeBase = CandBase
            End If
            
            'luiz acrescentou a busca pelo campo FuncaoPesquisa em 2005-08-01
            If existeFuncPesquisa Then
                If Not IsNull(RecTopico("FuncaoPesquisa")) Then
                    FuncaoPesquisa = Trim(RecTopico("FuncaoPesquisa"))
                Else
                    FuncaoPesquisa = ""
                End If
            Else
                FuncaoPesquisa = ""
            End If
            
            If Not IsNull(RecTopico("Tabela")) Then
                If RecTopico("Tabela") <> "" Then
                    nometabela = RecTopico("Tabela")
                Else
                    nometabela = CandTabela
                End If
            Else
                nometabela = CandTabela
            End If
        Else
            NomeBase = CandBase
            nometabela = CandTabela
        End If
    Else
        NomeBase = CandBase
        nometabela = CandTabela
    End If
    
    RecTopico.Close
    Set RecTopico = Nothing
    
    BaseTopico.Close
    Set BaseTopico = Nothing
    
    Exit Sub
        
Fim:
    glLog.EscreveLog "[DefineBase] -  FuncoesAuxiliares.DefineBase: Erro:" & Err.Description
    Exit Sub
    
ErroBase:
        glLog.EscreveLog "[DefineBase] - Erro:" & Err.Description
        Resume Continua
    
        
erroFuncaoPesquisa:
    existeFuncPesquisa = False
    Resume proxFuncaoPesquisa
    
proxFuncaoPesquisa:
    On Error GoTo Fim
    QueryTopico = "SELECT Base,Tabela FROM Topicos WHERE Nome = '" & NomeTopico & "'"
    GoTo continuaFuncaoPesquisa
        
End Sub

'Function PegaValoresCamposChave(ListaCampos As String, NewDin As DAO.Recordset) As String
Function PegaValoresCamposChave(ListaCampos As String, NewDin As ADODB.Recordset) As String
    Dim acabou As Boolean
    Dim poscam As Long
    Dim camchave, Campo As String
    Dim resp
    'Pega os valores dos campos chave e os separa por '#'
    'São os valores que vão diferenciar os registros, em caso
    'de modificação ou remoção
        
    camchave = ListaCampos
    ordem = 1
    acabou = False
    resp = ""
    If camchave = "" Then acabou = True
    While Not acabou
        poscam = InStr(camchave, "#")
        If (poscam <> 0) Then
            Campo = Left(camchave, poscam - 1)
            camchave = Mid(camchave, poscam + 1)
        Else
            Campo = camchave
            acabou = True
        End If
        If (ordem > 1) Then resp = resp & "#"
        resp = resp & NewDin(Campo)
        ordem = ordem + 1
    Wend
    PegaValoresCamposChave = resp
    Exit Function
    
Fim:
    glLog.EscreveLog "[FuncoesAuxiliares.PegaValoresCamposChave] - Erro:" & Err.Description
    glLog.EscreveLog "[FuncoesAuxiliares.PegaValoresCamposChave] - ListaCampos: " & ListaCampos
    Exit Function
End Function

Sub EscreveCampos(Tipo As Long, Optional QueryPar As String)
    Dim i As Long
    Dim valoropcao As String
    Dim pode As Boolean
    Dim Constantes As TConstantesLab245
    
    On Error GoTo Erro
    
    Set Constantes = New TConstantesLab245
    'Definindo a opção, de acordo com o tipo
    Select Case Tipo
        Case Constantes.ModificarCon
            valoropcao = "modificar"
        Case Constantes.ApagarCon
            valoropcao = "apagar"
        Case Else
            valoropcao = ""
    End Select
    
    For i = 0 To Campos
        pode = True
        If Entrada(i) <> "" And Not IsNull(Entrada(i)) Then
            'Valor deve passar vazio se for pra modificar
            Select Case i
                Case 41 To 60
                    If Tipo = Constantes.ModificarCon Then
                        pode = False
                    End If
                'opcao deve ser vazio se já foi feita a modificação
                Case 86
                    Print #3, "<input type=hidden name=""" & Replace(Entrada(i), Chr(34), "") & """ value=""" & Replace(valoropcao, Chr(34), "") & """>"
                    pode = False
                'se confirmar modificacao regnod, RegApag devem ser vazios
                Case 87, 88, 91
                    If Tipo = Constantes.ControlCon Then
                        Print #3, "<input type=hidden name=""" & Replace(Entrada(i), Chr(34), "") & """>"
                        pode = False
                    End If
                ' Atenção à query!
                Case 95
                    If QueryPar <> "" Then
                        Print #3, "<input type=hidden name=""" & Replace(Entrada(i), Chr(34), "") & """ value=""" & Replace(Query, Chr(34), "") & """>"
                        pode = False
                    End If
                'Login, senha e chavecontrol:
                
                'Senha NUNCA é passada
                
                'Chave SEMPRE é passado como ChaveGlobal
                
                'Se for control e chavecontrol é vazio
                '   não passo login
                'senão
                '   login é normalmente passados
                
                Case 161 'Login
                    'If Tipo = Constantes.ControlCon And (Dado(204) = "" Or IsNull(Dado(204))) Then
                    'Mudei a linha acima para a abaixo pois a estrutura Dado só
                    'tem 201 posições (0 a 200)
                    If Tipo = Constantes.ControlCon Then
                        pode = False
                    End If
                Case 162 'senha
                    pode = False
                Case 163 'modificacampos (deve estar vazio quando for confirmar modificação)
                    If Tipo = Constantes.ControlCon Then
                        Print #3, "<input type=hidden name=""" & Replace(Entrada(i), Chr(34), "") & """>"
                        pode = False
                    End If
                'O case abaixo não ocorre, pois a variável Campos é inicializada
                'com o valor 189
'                Case 204 'chavecontrol
'                    Log "Connecta.log", "FuncoesAuxiliares.EscreveCampos", , "Cheguei no chavecontrol -> vou imprimir com o valor de '" & ChaveGLOBAL & "'"
'
'                    Print #3, "<input type=hidden name=""" & Replace(Entrada(i), Chr(34), "") & """ value=""" & Replace(ChaveGLOBAL, Chr(34), "") & """>"
'
'                    Log "Connecta.log", "FuncoesAuxiliares.EscreveCampos", , "Imprimi <input type=hidden name=""" & Replace(Entrada(i), Chr(34), "") & """ value=""" & Replace(ChaveGLOBAL, Chr(34), "") & """>"
'
'                    pode = False
            End Select
            If pode And Dado(i) <> "" And Not IsNull(Dado(i)) Then
                Print #3, "<input type=hidden name=""" & Replace(Entrada(i), Chr(34), "") & """ value=""" & Replace(Dado(i), Chr(34), "") & """>"
            End If
        End If
     Next i
    Exit Sub
    
Erro:
    glLog.EscreveLog "[EscreveCampos] - Erro:" & Err.Description
    Resume Next
End Sub

Function ComparaChave(Chave As String, Campo) As String
    Dim resp, camchave, possivel
    ' Diz se CAMPO faz parte da chave primária
    ' Retorna CAMPO caso sim
    ' Retorna "" caso não
    
    resp = ""
    camchave = Chave
    
    While Not acabou
        poscam = InStr(camchave, "#")
        If (poscam <> 0) Then
            possivel = Left(camchave, poscam - 1)
            camchave = Mid(camchave, poscam + 1)
        Else
            possivel = camchave
            acabou = True
        End If
        If UCase(possivel) = UCase(Campo) Then
            resp = possivel
            acabou = True
        End If
    Wend
    
    glLog.EscreveLog "[FuncoesAuxiliares.ComparaChave] - Campo aqui em comparachave: " & Campo
    
    ComparaChave = resp
    Exit Function
    
Fim:
   glLog.EscreveLog "[FuncoesAuxiliares.ComparaChave] - Erro:" & Err.Description
   glLog.EscreveLog "[FuncoesAuxiliares.ComparaChave] - Chave: " & Chave & " / Campo: " & Campo
   Exit Function
End Function


Function DefineOrdenacao(NomeBase As String, nometabela As String, Optional vOrdenaAgrupado As String = "")
    Dim Query As String
    Dim i As Long
    
    Query = ""
    'Vou ordenar apenas pelo campo 1
    If CampoSubtitulo <> "" Then
        Query = " ORDER BY " & CampoSubtitulo
    Else
        For i = 1 To 20
            If Trim(Dado(i)) <> "" Then
                pos = InStr(Dado(i), "#*")
                If pos = 0 Then
                    Query = " ORDER BY " & Dado(i)
                Else
                    Query = " ORDER BY " & Left(Dado(i), pos - 1)
                End If
                Exit For
            Else
                i = i + 1
            End If
        Next
    End If
        
    DefineOrdenacao = Query
    
    Exit Function
    
Fim:
    glLog.EscreveLog "[FuncoesAuxiliares.DefineOrdenacao] - Erro:" & Err.Description
     glLog.EscreveLog "[FuncoesAuxiliares.DefineOrdenacao] - Query: " & Query
    Exit Function
End Function

Function DadosEmMemoria(ByVal strTabela As String) As Boolean

    Dim vInd As Long

    On Error GoTo DizErro
    
    For vInd = LBound(agTipoCampo) To UBound(agTipoCampo)
    
        If agTipoCampo(vInd).Tabela = strTabela Then
            
            DadosEmMemoria = True
            
            'Sai da função
            Exit Function
            
        End If
    
    Next vInd
    
    DadosEmMemoria = False
    Exit Function
     
     
DizErro:

    glLog.EscreveLog "[FuncoesAuxiliares.DefineAspas.DadosEmMemoria] - Erro:" & Err.Description
    Exit Function

End Function
Function RetornaCaracter(pTipo As Long) As String

    On Error GoTo DizErro
    
    If pTipo = adChar Or pTipo = adVarChar Or pTipo = adLongVarChar Or pTipo = adVarWChar Or pTipo = adBSTR Or pTipo = adLongVarWChar Or pTipo = adWChar Then
        RetornaCaracter = "'"
    Else
        If pTipo = adDBTimeStamp Or pTipo = adDate Or pTipo = adDBDate Or pTipo = adDBTime Then
            RetornaCaracter = "#"
        Else
            RetornaCaracter = ""
        End If
    End If
    
    glBoolDefinido = True
    
    Exit Function
    
DizErro:
    glLog.EscreveLog "[FuncoesAuxiliares.DefineAspas.RetornaCaracter] - erro:", Err.Description
    Exit Function
    
End Function
Function PegaCaracter(strTabela As String, strCampo As String) As String
    
    Dim Tipo As Long
    Dim i As Long
    Dim vInicio As Long
    Dim vFinal As Long
    
    On Error GoTo DizErro
    
    vInicio = LBound(agTipoCampo)
    vFinal = UBound(agTipoCampo)
    For i = vInicio To vFinal
        If UCase(agTipoCampo(i).nomecampo) = UCase(strCampo) And _
           UCase(agTipoCampo(i).Tabela) = UCase(strTabela) Then
            
            RetornaCaracter agTipoCampo(i).Tipo
            Exit For
            
        End If
    Next i
    
    Exit Function
    

DizErro:

   glLog.EscreveLog "FuncoesAuxiliares.DefineAspas.Pegacaracter] - Erro:" & Err.Description
    End

End Function

Function CarregaCampos(ByVal rs As ADODB.Recordset, _
                       ByVal strTabela As String, _
                       ByVal strCampo As String) As String
    
    Dim arrCampos() As TTipoCampo
    Dim Campo As ADODB.Field
    Dim vstrAux As String

    On Error GoTo DizErro
    vstrAux = UCase(strCampo)
    For Each Campo In rs.Fields
        
        If UCase(Campo.Name) = vstrAux Then
            ReDim agTipoCampo(UBound(agTipoCampo) + 1)
            With agTipoCampo(UBound(agTipoCampo))
                .Tabela = strTabela
                .nomecampo = strCampo
                .Tipo = Campo.Type
                CarregaCampos = RetornaCaracter(Campo.Type)
                glBoolDefinido = True
            End With
            Exit For
        End If
   Next Campo
   
   Exit Function

DizErro:

    glLog.EscreveLog "[FuncoesAuxiliares.DefineAspas.CarregaCampos] - Erro:" & Err.Description
    Exit Function
    
End Function

Function DefineAspas(ByVal NomeBase As String, ByVal nometabela As String, ByVal nomecampo As String) As String
    
'    Dim Tabelas() As String
'
'    Dim DBLocal As Database
'
'    Dim i As Long
'    Dim Tipo As Long
'
'    Dim Query As String
'    Dim strBaseParaUsoADODB As String
'
'    Dim ConfigADOX As New ADOX.Catalog
'    Dim ConfigADODB As New ADODB.Connection
'    Dim Campo As ADODB.Field
'    Dim TabelaAAnalisar As New ADODB.Recordset
'
'    On Error GoTo Fim
'
'    Log "Connecta.log", "FuncoesAuxiliares.DefineAspas", , "SGBD: " & SGBD
'
'    Tabelas = Split(nometabela, ",")
'    For i = LBound(Tabelas) To UBound(Tabelas)
'        If Trim$(Tabelas(i)) = "" Then
'            Tabelas(i) = ","
'        Else
'            Tabelas(i) = Trim$(Tabelas(i))
'        End If
'    Next
'
'    Tabelas = Filter(Tabelas, ",", False)
'    For i = LBound(Tabelas) To UBound(Tabelas)
'        If Not (Left(Tabelas(i), 1) = "[" And Right(Tabelas(i), 1) = "]") Then
'            If InStr(Tabelas(i), " ") > 0 Then
'                Tabelas(i) = Left(Tabelas(i), InStr(Tabelas(i), " ") - 1)
'            End If
'        End If
'    Next
'
'    If (InStr(nomecampo, ".") > 0) And Len(nomecampo) > InStr(nomecampo, ".") Then
'        nomecampo = Mid(nomecampo, InStr(nomecampo, ".") + 1, Len(nomecampo))
'    End If
'
'    'traverso mudou 2006-03-20
'    If SGBD = "" Or SGBD = "acces" Then
'        ConfigADODB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NomeBase
'        ConfigADOX.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & NomeBase
'    Else
'        strBaseParaUsoADODB = Replace(NomeBase, "ODBC;", "")
'
'        ConfigADODB.Open strBaseParaUsoADODB
'        ConfigADOX.ActiveConnection = strBaseParaUsoADODB
'    End If
'
'    For i = LBound(Tabelas) To UBound(Tabelas)
'
'        glBoolDefinido = False
'
'        If DadosEmMemoria(Tabelas(i)) Then
'            'Se o vetor já estiver preenchido, procurar no vetor retornar o caracter
'            DefineAspas = PegaCaracter(Tabelas(i), nomecampo)
'        Else
'            Query = "SELECT * FROM " & Tabelas(i) & " WHERE 1 = 2"
'            TabelaAAnalisar.Open Query, ConfigADODB, , , adCmdText
'            DefineAspas = CarregaCampos(TabelaAAnalisar, Tabelas(i), nomecampo)
'            TabelaAAnalisar.Close
'        End If
'
'        If glBoolDefinido Then
'            ConfigADODB.Close
'            Set ConfigADOX = Nothing
'            Exit Function
'        End If
'
''        For Each Campo In TabelaAAnalisar.Fields
''            If UCase(Campo.Name) = UCase(nomecampo) Then
''                Tipo = Campo.Type
''
''                If Tipo = adChar Or Tipo = adVarChar Or Tipo = adLongVarChar Or Tipo = adVarWChar Or Tipo = adBSTR Or Tipo = adLongVarWChar Or Tipo = adWChar Then
''                    DefineAspas = "'"
''
''                Else
''                    If Tipo = adDBTimeStamp Or Tipo = adDate Or Tipo = adDBDate Or Tipo = adDBTime Then
''                        DefineAspas = "#"
''                    Else
''                        DefineAspas = ""
''                    End If
''
''                End If
''
''                TabelaAAnalisar.Close
''                ConfigADODB.Close
''                Set ConfigADOX = Nothing
''
''                Exit Function
''            End If
''        Next Campo
'
'   Next i
'
'
'   ConfigADODB.Close
'   Set ConfigADOX = Nothing

   DefineAspas = "'"
   Exit Function
   
Fim:
    glLog.EscreveLog "[FuncoesAuxiliares.DefineAspas] - Erro:" & Err.Description
    glLog.EscreveLog "[FuncoesAuxiliares.DefineAspas] - Base: " & NomeBase & "campo" & nomecampo
    Exit Function
            
End Function


Function ContaCores(ListaCores) As Long
    pos = InStr(ListaCores, "-")
    numpos = 1
    While (pos <> 0)
        numpos = numpos + 1
        posant = pos
        pos = InStr(posant + 1, ListaCores, "-")
    Wend
    ContaCores = numpos
End Function

Function PodeEscrever(NumControl As Long, TipoAcao As Long) As Boolean
    
    ' NumControl é um binário
    ' Cada bit representa um direito a um tipo de acesso
    If (NumControl And (2 ^ TipoAcao)) = 2 ^ TipoAcao Then
        PodeEscrever = True
    Else
        PodeEscrever = False
    End If
    
End Function

Function DefineEstilo(EntradaPura As String) As Long
    
    Dim Constantes As TConstantesLab245
    Dim Ent As String
    
    Set Constantes = New TConstantesLab245

    Ent = UCase(EntradaPura)
    If IsNumeric(Ent) Then
        DefineEstilo = Val(Ent)
    Else
        Select Case Ent
            Case "LISTA"
                DefineEstilo = Constantes.EstiloListaCon
            Case "RELATORIOP"
                DefineEstilo = Constantes.EstiloRelatPCon
            Case "TABELA"
                DefineEstilo = Constantes.EstiloTabelaCon
            Case "QUANTIDADE"
                DefineEstilo = Constantes.EstiloQuantCon
            Case Else
                DefineEstilo = Constantes.EstiloRelatCon
        End Select
    End If

End Function
Function DefineCorLinha(Contador As Long) As String
    
    Dim c As Long

    c = (Contador - 1) Mod Alternadas
    DefineCorLinha = CoresLinhas(c + 1)
    
End Function

Function DefineFormato(NomeBase As String, nometabela As String, nomecampo As String) As String
    
    Dim DBLocal As DAO.Database

    'On Error GoTo Fim
    On Error Resume Next
    Set DBLocal = OpenDatabase(NomeBase, False, False, SGBD)
    Select Case DBLocal(nometabela)(nomecampo).Type
        Case dbDate
            DefineFormato = "(formato aaaa-dd-mm)"
        Case dbText, dbMemo
            DefineFormato = ""
        Case Else
            DefineFormato = "(somente números)"
    End Select
    DBLocal.Close
    Exit Function
    
Fim:
    glLog.EscreveLog "[FuncoesAuxiliares.DefineFormato] - Erro:" & Err.Description
    glLog.EscreveLog "NomeTabela: " & nometabela & " / nomecampo: " & nomecampo
    Exit Function
End Function

Function TabelaExiste(NomeBase As String, nometabela As String, ErrDesc As String) As Long
'Valores de retorno:
'0 - Nenhum erro
'1 - Erro no acesso à tabela (prov. tabela inexistente)
'2 - Erro na abertura da base (prov. base inexistente)
    Dim QueryAux As String, RSAux As DAO.Recordset, BaseAux As DAO.Database
    Dim lngAuxExiste As Long, tdTeste As DAO.TableDef, qdTeste As DAO.QueryDef
    ErrDesc = ""
    
    lngAuxExiste = 1
    If nometabela <> "" Then
        On Error GoTo ErroBase
        If SGBD = "" Then
            Set BaseAux = OpenDatabase(NomeBase, False, False, "")
        Else
            Set BaseAux = OpenDatabase(NomeBase, False, False, "ODBC")
        End If
        
        'Melhorar a forma de verificar a existência da tabela
        'Feito assim para permitir que o usuário passe "tabelaA, tabelaB"
        If SGBD = "" Then
            QueryAux = "SELECT COUNT(*) FROM " & nometabela
            Set RSAux = BaseAux.OpenRecordset(QueryAux, dbOpenDynaset)
            Set RSAux = Nothing
            lngAuxExiste = 0
        Else
            QueryAux = "SELECT COUNT(*) FROM " & nometabela
            Set RSAux = BaseAux.OpenRecordset(QueryAux, dbOpenDynaset, dbSQLPassThrough)
            Set RSAux = Nothing
            lngAuxExiste = 0
        End If
            
    End If
    
    TabelaExiste = lngAuxExiste
    
    Exit Function
    
ErroTabela:
    ErrDesc = Err.Description
    TabelaExiste = 1
    Resume Erro

ErroBase:
    ErrDesc = Err.Description
    TabelaExiste = 2
    Resume Erro

Erro:
   glLog.EscreveLog "[FuncoesAuxiliares.TabelaExiste] - Erro:" & Err.Description
   glLog.EscreveLog "SGBD: " & SGBD & " / Tabela: " & nometabela
   Exit Function
End Function

Function EscreveFormDados() As String
    Dim resp As String
    resp = "<form name=ResultadoConnect method=POST action=foldcon.dll>" & vbCrLf
    For i = 0 To Campos
        '86 é opção: requerido por causa de javascript
        '87 é RegApag: requerido por causa de javascript
        '88 é RegMod: requerido por causa de javascript
        '91 é RegCri: requerido por causa de javascript
        '101 é RegistroIni: requerido por causa de javascript
        If i = 86 Or i = 87 Or i = 88 Or i = 91 Or i = 101 Then
            If Entrada(i) <> "" And Not IsNull(Entrada(i)) Then
                Print #3, "   <input type=hidden name=""" & Replace(Entrada(i), Chr(34), "") & """ value=""" & Replace(Dado(i), Chr(34), "") & """>"
            End If
        Else
            '161 => Login: tratado abaixo
            '162 => Senha: não é exibido
            '164 => ChaveLogOn: tratado abaixo
            If i <> 161 And i <> 162 And i <> 164 Then
                If Entrada(i) <> "" And Not IsNull(Entrada(i)) And Dado(i) <> "" And Not IsNull(Dado(i)) Then
                    resp = resp & "<input type=hidden name=""" & Replace(Entrada(i), Chr(34), "") & """  value=""" & Replace(Dado(i), Chr(34), "") & """>" & vbCrLf
                End If
            End If
        End If
    Next
    resp = resp & "<hr><font size=-1 face=Verdana,Arial>" & vbCrLf
    If Login <> "" And ChaveLogOn <> "" Then
        resp = resp & "<input type=hidden name=""Login"" value=""" & Replace(Login, Chr(34), "") & """>" & vbCrLf
        resp = resp & "<input type=hidden name=""ChaveLogOn"" value=""" & Replace(ChaveLogOn, Chr(34), "") & """>" & vbCrLf
    Else
        resp = resp & "<BR>Login" & vbCrLf
        resp = resp & "<BR><input type=text name=Login value=''>" & vbCrLf
        resp = resp & "<BR>Password" & vbCrLf
        resp = resp & "<BR><input type=password name=Senha>" & vbCrLf
        resp = resp & "<br><input type=submit value='Confirm'>" & vbCrLf
    End If
    resp = resp & "</form>" & vbCrLf
    EscreveFormDados = resp
End Function
