Attribute VB_Name = "PADRAO"
Option Explicit
Global Const G_GridHeaderColor = &HD8E9EC    '&HBFDDD3
Global Const gCODSIS = "Cadastro"

'Public Function F_LOV(pXDb As DS_BANCO, Tabela$, Optional pQry, Optional pMultRows, Optional pMerge, Optional pCab, Optional pIdCampo, Optional pTitulo, Optional pisTree)
Public Function F_LOV(pXDb As Object, Tabela$, Optional pQry, Optional pMultRows, Optional pMerge, Optional pCab, Optional pIdCampo, Optional pTitulo, Optional pisTree)
  '* Propriedades da Classe
   Dim Cab, IdCampo
   Dim SQL        As Variant 'Coluna de ordenação (Padrao = 2 ) ou Query de acesso
   Dim Tit        As String
   Dim mPointer   As Integer
   Dim lAux       As Integer
   Dim MultRows   As Boolean
   Dim Merge      As Boolean
   Dim MergeRow   As Collection
   Dim MergeCol   As Collection
   Dim WidthScr   As Double
   
   Dim MyLOV      As Object
   
   mPointer = Screen.MousePointer
   
   Screen.MousePointer = vbHourglass
   
   If Not IsMissing(pCab) Then Cab = pCab
   If Not IsMissing(pIdCampo) Then IdCampo = pIdCampo
   If Not IsMissing(pTitulo) Then Tit = pTitulo
   If Not IsMissing(pQry) Then SQL = pQry
   If Not IsMissing(pMultRows) Then MultRows = pMultRows
   If Not IsMissing(pMerge) Then Merge = pMerge
   
   Set MyLOV = CriarObjeto("XActive.XLOV")
   Tabela = UCase(Tabela)
   
   Select Case UCase(Tabela)
      
      Case "COLIGADA"
         '**************
         '* COLIGADA
         '**************
         
         If IsMissing(pQry) Then
            SQL = "SELECT "
            SQL = SQL & " IDCOLIGADA,"
            SQL = SQL & " NMCOLIGADA"
            SQL = SQL & " FROM COLIGADA"
         Else
            SQL = pQry
         End If
         
         Cab = Array("Código", "IDCOLIGADA", 8, vbLeftJustify, _
                     "Descrição", "NMCOLIGADA", 30, vbLeftJustify)
         
         IdCampo = Array("IDCOLIGADA", "NMCOLIGADA")
         Tit$ = "Coligadas"
      
      Case "PSCHEDULE"
         '**************
         '* SCHEDULE
         '**************
         
         If IsMissing(pQry) Then
            SQL = "SELECT "
            SQL = SQL & " IDSCHEDULE,"
            SQL = SQL & " CODSCHEDULE"
            SQL = SQL & " FROM PSCHEDULE"
         Else
            SQL = pQry
         End If
         
         Cab = Array("Código", "IDSCHEDULE", 8, vbLeftJustify, _
                     "Descrição", "CODSCHEDULE", 15, vbLeftJustify)
         
         IdCampo = Array("IDFAM", "DSCFAM", "ACESSO")
         Tit$ = "Schedule"
   
      Case "ACENTRAL"
         '**************
         '* Tabela de Central telefonica
         '**************
         SQL = "Select IDCENTRAL,DSCCENTRAL"
         SQL = SQL & " From ACENTRAL"
         
         Cab = Array("Código", "IDCENTRAL", 5, vbLeftJustify, _
                     "Descrição", "DSCCENTRAL", 30, vbLeftJustify)
         IdCampo = Array("IDCENTRAL", "DSCCENTRAL")
         Tit$ = "Central Telefônica"
         
         
      Case "APGTO"
         '**************
         '* Tabela Autorização de Pagamento
         '**************
         SQL = "Select SETORAP, Right('00000' + cast(IDAP as varchar),5) as [IDAP]"
         SQL = SQL & ", ANOAP, DSCAP, DTEMISSAO, FAVORECIDO "
         SQL = SQL & " From APGTO"
         
         Cab = Array("Setor", "SETORAP", 10, vbLeftJustify, _
                     "Código", "IDAP", 8, vbLeftJustify, _
                     "Ano", "ANOAP", 4, vbCenter, _
                     "Descrição", "DSCAP", 40, vbLeftJustify, _
                     "Emissão", "DTEMISSAO", 10, vbCenter, _
                     "Favorecido", "FAVORECIDO", 40, vbLeftJustify)
         IdCampo = Array("SETORAP", "IDAP", "ANOAP")
         Tit$ = "Autorização de Pagamento"
      
      Case "APLICACAO"
         '**************
         '* Tabela Aplicação
         '**************
         Cab = Array("IDAPLIC", "IDAPLIC", 0, vbLeftJustify, _
                     "Descrição", "DSCAPLIC", 30, vbLeftJustify, _
                     "", "IDPROJ", 0, vbLeftJustify, _
                     "", "IDSUB", 0, vbLeftJustify)
         IdCampo = Array("IDAPLIC", "DSCAPLIC", "IDPROJ", "IDSUB")
         Tit$ = "Aplicação"
      
      Case "BANCO"
         '**************
         '* Tabela Banco
         '**************
         Cab = Array("Código", "CODBANCO", 8, vbLeftJustify, _
                     "Descrição", "DSCBANCO", 60, vbLeftJustify)
         IdCampo = Array("CODBANCO", "DSCBANCO")
         Tit$ = "Bancos"
      
      Case "CCUSTO"
         '**************
         '* Tabela Centro de Custo
         '**************
         If IsMissing(pisTree) Then
            pisTree = True
            MyLOV.CAMPO_ID = "IDCCUSTO"
            MyLOV.CAMPO_CODIGO = "CODCCUSTO"
            MyLOV.CAMPO_NOME = "DSCCCUSTO"
            MyLOV.ExibeCodigo = True
         End If
         If IsMissing(pQry) Then SQL = 3
         
         If IsMissing(pCab) Then
            Cab = Array("Seq.", "IDCCUSTO", 4, vbLeftJustify, _
                        "Código", "CODCCUSTO", 10, vbLeftJustify, _
                        "Centro de Custo", "DSCCCUSTO", 50, vbLeftJustify, _
                        "IDPAI", "IDPAI", 0, vbLeftJustify)
         End If
         IdCampo = Array("IDCCUSTO", "CODCCUSTO", "DSCCCUSTO")
         Tit$ = "Centros de Custo"
      
      Case "CLIENTE"
         '**************
         '* Tabela Cliente
         '**************
         Cab = Array("Seq.", "IDCLIENTE", 5, vbLeftJustify, _
                     "Razão Social", "RZCLI", 60, vbLeftJustify, _
                     "CNPJ", "CNPJCLI", 0, vbLeftJustify)
         IdCampo = Array("IDCLIENTE", "RZCLI", "CNPJCLI")
         Tit$ = "Cliente"
      
      Case "CONDVENDA"
         '**************
         '* Tabela Condição de vendas
         '**************
         Cab = Array("Código", "IDCONDVENDA", 10, vbLeftJustify, _
                     "Condição de Venda", "CONDVENDA", 20, vbLeftJustify)
         IdCampo = Array("IDCONDVENDA", "CONDVENDA")
         Tit$ = "Condição de Venda"
         
         SQL = "Select IDCONDVENDA, CONDVENDA"
         SQL = SQL & " From CONDVENDA "
         SQL = SQL & " Where Not ( CONDVENDA is Null or CONDVENDA = '')"
         SQL = SQL & " order by CONDVENDA"
         
      Case "CONTRATADA"
         '**************
         '* Tabela Contratada
         '**************
         Cab = Array("Cód.", "IDCONTRATADA", 5, vbLeftJustify, _
                     "Nome", "NOME", 30, vbLeftJustify, _
                     "CNPJ", "CNPJ", 15, vbLeftJustify)
         IdCampo = Array("IDCONTRATADA", "NOME", "CNPJ")
         Tit$ = "Contratada"
         
      Case "DESPESA"
         '**************
         '* Tabela Despesas
         '**************
         WidthScr = 7500
         If IsMissing(pisTree) Then
            pisTree = True
            MyLOV.CAMPO_ID = "IDDESP"
            MyLOV.CAMPO_CODIGO = "CODDESP"
            MyLOV.CAMPO_NOME = "DSCDESP"
            MyLOV.ExibeCodigo = True
         End If
         If IsMissing(pQry) Then pQry = 3
         
         Cab = Array("Seq.", "IDDESP", 8, vbLeftJustify, _
                     "Código", "CODDESP", 10, vbLeftJustify, _
                     "Despesa", "DSCDESP", 50, vbLeftJustify) ', _
                     "IDPAI", "IDPAI", 0, vbLeftJustify)
         IdCampo = Array("IDDESP", "CODDESP", "DSCDESP")
         
         Tit$ = "Despesas"
      
      Case "DISCIPLINA"
         '**************
         '* Tabela Disciplina
         '**************
         If IsMissing(pQry) Then
            SQL = "Select IDDISCIPLINA, DSCDISCIPLINA"
            SQL = SQL & " From DISCIPLINA"
            SQL = SQL & " Order By 1"
         End If
         If IsMissing(pCab) Then
            Cab = Array("Seq.", "IDDISCIPLINA", 6, vbLeftJustify, _
                        "Código", "DSCDISCIPLINA", 20, vbLeftJustify)
         End If
         If IsMissing(pIdCampo) Then IdCampo = Array("IDDISCIPLINA")
         If IsMissing(pTitulo) Then Tit$ = "DISCIPLINA"

      Case "DPA"
         '**************
         '* Tabela DPA
         '**************
         Cab = Array("Setor", "SETORDPA", 10, vbCenter, _
                     "NºDPA", "DPA", 10, vbCenter, _
                     "Ano", "ANODPA", 5, vbCenter, _
                     "Data", "DTEMISSAO", 10, vbCenter, _
                     "Usuário", "NOME", 20, vbLeftJustify)
         IdCampo = Array("SETORDPA", "DPA", "ANODPA")
         Tit$ = "D.P.A."
      
      Case "ENDENTRG"
         '**************
         '* Tabela Endereços
         '**************
         Cab = Array("Código", "CODEND", 8, vbLeftJustify, _
                     "Endereço", "DSCEND", 60, vbLeftJustify)
         IdCampo = Array("CODEND", "DSCEND")
         Tit$ = "Endereços"
         
      Case "EMPRESA"
         '**************
         '* Tabela Empresa
         '**************
         Cab = Array("Seq.", "IDEMPRESA", 5, vbLeftJustify, _
                     "Razão Social", "NMEMPRESA", 60, vbLeftJustify, _
                     "CNPJ", "CNPJ", 0, vbLeftJustify)
         IdCampo = Array("IDEMPRESA", "NMEMPRESA", "CNPJ")
         Tit$ = "Empresa"
        
      
         
      Case "FAMILIAPROD"
         '**************
         '* Tabela Família de Produtos
         '**************
         Cab = Array("Código", "IDFAM", 8, vbLeftJustify, _
                     "Descrição", "DSCFAM", 25, vbLeftJustify, _
                     "Controle de Acesso", "ACESSO", 13, vbCenter)
         IdCampo = Array("IDFAM", "DSCFAM", "ACESSO")
         Tit$ = "Família de Produtos"
      
      Case "FORNECEDOR"
         '**************
         '* Tabela Fornecedor
         '**************
         If IsMissing(pCab) Then
            Cab = Array("Código", "IDFOR", 6, vbLeftJustify, _
                        "Descrição", "NMFOR", 40, vbLeftJustify, _
                        "CNPJ", "CNPJFOR", 15, vbLeftJustify, _
                        "País", "SIGLAPAIS", 5, vbLeftJustify)
         End If
         If IsMissing(pIdCampo) Then
            IdCampo = Array("IDFOR", "NMFOR", "CNPJFOR")
         End If
         Tit$ = "Fornecedores"
         
         If IsMissing(pQry) Then
            SQL = "Select IDFOR, NMFOR "
            SQL = SQL & " , SUBSTRING(CNPJFOR,1,2) + '.' + SUBSTRING(CNPJFOR,3,3) + '.'"
            SQL = SQL & " + SUBSTRING(CNPJFOR,6,3) + '/' + SUBSTRING(CNPJFOR,9,4) + '-'"
            SQL = SQL & " + SUBSTRING(CNPJFOR,13,2) as CNPJFOR"
            SQL = SQL & " , SIGLAPAIS"
            SQL = SQL & " From FORNECEDOR "
            SQL = SQL & " Where Not DTCADASTRO is Null"
            SQL = SQL & " order by 2"
         End If
         
      Case "FUNCAOPROD"
         '**************
         '* Tabela de Função do Produto
         '**************
         Cab = Array("Código", "FUNPROD", 10, vbLeftJustify, _
                     "Descrição", "DSCFUNPROD", 50, vbLeftJustify)
         IdCampo = Array("FUNPROD", "DSCFUNPROD")
         Tit$ = "Função do Produto"
      
      Case "IDESENHO"
         '********************
         '* Tabela de Interface de desenhos
         '********************
         
         Cab = Array("Spool", "DSCSPOOL", 7, vbLeftJustify, _
                   "Isometrico", "DSCISOMETRICO", 7, vbLeftJustify, _
                   "Projeto", "IDPROJ", 0, vbLeftJustify, _
                   "Sub-Projeto", "IDSUB", 0, vbLeftJustify, _
                   "Aplicação", "IDAPLIC", 0, vbLeftJustify, _
                   "Tipo Desenho", "TPDESENHO", 6, vbLeftJustify, _
                   "Revisão", "REVSPOOL", 6, vbLeftJustify)
         IdCampo = Array("DSCSPOOL", "DSCISOMETRICO", "IDPROJ", "IDSUB", "IDAPLIC", "TPDESENHO", "REVSPOOL")
         Tit$ = "Desenhos"
         
      Case "PSPOOL"
         '*******************
         '* Tabela de Spools
         '*******************
         
         Cab = Array("Spool", "CODSPOOL", 20, vbLeftJustify, _
                  "Rev", "REV", 10, vbLeftJustify, _
                  "Isométrico", "DSCISOMETRICO", 20, vbLeftJustify, _
                  "Projeto", "IDPROJ", 0, vbLeftJustify, _
                  "Sub-Projeto", "IDSUB", 0, vbLeftJustify, _
                  "Elevação", "IDAPLIC", 0, vbLeftJustify)
         IdCampo = Array("CODSPOOL", "REV", "DSCISOMETRICO", "IDPROJ", "IDSUB", "IDAPLIC")
         Tit$ = "Spools"
         
      Case "PFUNCAO"
         '**************
         '* Tabela de Funções
         '**************
         Cab = Array("IDFUNCAO", "IDFUNCAO", 0, vbLeftJustify, _
                     "CODFUNCAO", "CODFUNCAO", 10, vbLeftJustify, _
                     "DSCFUNCAO", "DSCFUNCAO", 40, vbLeftJustify)
         IdCampo = Array("CODFUNCAO")
         Tit$ = "FUNÇÕES"
         Tabela = "FUNÇÕES"
         
      Case "GRPFOR"
         '**************
         '* Tabela Grupo de Fornecimento
         '**************
         WidthScr = 7500
         pisTree = True
         MyLOV.CAMPO_ID = "IDGRPFOR"
         MyLOV.CAMPO_CODIGO = "CODGRPFOR"
         MyLOV.CAMPO_NOME = "DSCGRPFOR"
         
         MyLOV.ExibeCodigo = True
         If IsMissing(pQry) Then pQry = 3
         
         If IsMissing(pQry) Then
            SQL = "Select IDGRPFOR, CODGRPFOR, DSCGRPFOR, IDPAI "
            SQL = SQL & " From GRPFOR "
            SQL = SQL & " Where IDGRPFOR <> 0 "
         End If
         If IsMissing(pCab) Then
            Cab = Array("Seq.", "IDGRPFOR", 0, vbLeftJustify, _
                        "Código", "CODGRPFOR", 10, vbLeftJustify, _
                        "Descrição", "DSCGRPFOR", 30, vbLeftJustify)
         End If
         If IsMissing(pIdCampo) Then
            IdCampo = Array("IDGRPFOR", "CODGRPFOR", "DSCGRPFOR")
         End If
         Tit$ = "Grupo de Fornecimento"
      
      Case "GRPSMS"
         '****************
         '* Tabela Grupos de SMS
         '****************
         Cab = Array("Código", "IDGRPSMS", 8, vbLeftJustify, _
                     "Descrição", "DSCGRPSMS", 25, vbLeftJustify, _
                     "Observações", "OBSSMS", 13, vbLeftJustify)
         IdCampo = Array("IDGRPSMS", "DSCGRPSMS", "OBSSMS")
         Tit$ = "Grupos SMS"
      
      Case "GRPUSU"
         '**************
         '* Tabela Grupo de Usuários
         '**************
         Cab = Array("Código", "IDGRUPO", 10, vbLeftJustify, _
                     "Descrição", "DSCGRUPO", 30, vbLeftJustify)
         IdCampo = Array("IDGRUPO", "DSCGRUPO")
         Tit$ = "Grupo de Usuários"
      
      Case "NOTAFISCAL"
         '**************
         '* Tabela Pais
         '**************
         Cab = Array("IDFOR", "IDFOR", 0, vbLeftJustify, _
                     "Fornecedor", "NMFOR", 20, vbLeftJustify, _
                     "Número", "IDNOTA", 10, vbLeftJustify, _
                     "Série", "SERIE", 5, vbLeftJustify, _
                     "Emissão", "DTEMISSAO", 10, vbLeftJustify)
         IdCampo = Array("IDNOTA", "SERIE", "IDFOR")
         Tit$ = "Notas Fiscais"
         If IsMissing(pQry) Then
            SQL = "Select N.IDFOR, F.NMFOR , N.IDNOTA, N.SERIE, N.DTEMISSAO"
            SQL = SQL & " From NOTAFISCAL N left join FORNECEDOR F "
            SQL = SQL & " on N.IDFOR = F.IDFOR"
            SQL = SQL & " order by 1"
         End If
         
      Case "NVCARGO"
         '**************
         '* Tabela Níveis de Hierarquia
         '**************
         Cab = Array("Código", "IDNVCARGO", 10, vbLeftJustify, _
                     "Nível", "DSCNVCARGO", 30, vbLeftJustify)
         IdCampo = Array("IDNVCARGO", "DSCNVCARGO")
         Tit$ = "Níveis de Hierarquia"
      
      Case "MODULO"
         '**************
         '* Tabela MODULOS
         '**************
         If IsMissing(pisTree) Then
            pisTree = True
            MyLOV.CAMPO_ID = "ID"
            MyLOV.CAMPO_CODIGO = "IDMODU"
            MyLOV.CAMPO_NOME = "DSCMODU"
            'MyLOV.CAMPO_PAI = "MODUPAI"
            MyLOV.ExibeCodigo = False
         End If
         
         Cab = Array("ID", "ID", 15, vbLeftJustify, _
                     "Código", "IDMODU", 15, vbLeftJustify, _
                     "Descrição", "DSCMODU", 30, vbLeftJustify, _
                     "Módulo Superior", "IDPAI", 0, vbLeftJustify, _
                     "Ativo", "SITMODU", 5, vbLeftJustify)
         IdCampo = Array("ID", "IDMODU", "DSCMODU")
         Tit$ = "Módulos"
      
      Case "OCOMPRA"
         '**************
         '* Tabela Ordem de Compra
         '**************
         Cab = Array("Número", "IDOC", 10, vbLeftJustify, _
                     "Ano", "ANOOC", 6, vbLeftJustify, _
                     "Emissão", "DTEMISSAO", 10, vbLeftJustify, _
                     "Fornecedor", "NMFOR", 30, vbLeftJustify)
         IdCampo = Array("IDOC", "ANOOC")
         Tit$ = "Pedidos de Fornecimento"

      Case "PAIS"
         '**************
         '* Tabela Pais
         '**************
         Cab = Array("Código", "IDPAIS", 8, vbLeftJustify, _
                     "Descrição", "DSCPAIS", 30, vbLeftJustify, _
                     "Sigla", "SIGLAPAIS", 5, vbLeftJustify)
         IdCampo = Array("IDPAIS")
         Tit$ = "Paises"
         
      Case "PARAM"
         '**************
         '* Tabela Parâmetros
         '**************
         Cab = Array("Sistema", "CODSIS", 10, vbLeftJustify, _
                     "Código", "CODPARAM", 10, vbLeftJustify, _
                     "Descrição", "DSCPARAM", 30, vbLeftJustify, _
                     "Valor", "VLPARAM", 30, vbLeftJustify)
         IdCampo = Array("CODSIS", "CODPARAM")
         Tit$ = "Parâmetros"
      
      Case "PATIVIDADEEAP"
         '*******************
         '* Tabela Atividades
         '*******************
         If IsMissing(pQry) Then
            SQL = "SELECT P.IDPROJ, PR.NMPROJ, P.IDATVEAP,P.CODATVEAP"
            SQL = SQL & " FROM PATIVIDADEEAP P"
            SQL = SQL & " INNER JOIN PROJETO PR ON"
            SQL = SQL & "        P.IDPROJ = PR.IDPROJ"
            SQL = SQL & " ORDER BY PR.NMPROJ, P.CODATVEAP"
         End If
         
         Cab = Array("IDATVEAP", "IDATVEAP", 0, vbLeftJustify, _
                     "Cód. Atividade", "CODATVEAP", 20, vbLeftJustify, _
                     "IDPROJ", "IDPROJ", 0, vbLeftJustify, _
                     "Projeto", "NMPROJ", 0, vbLeftJustify, _
                     "Descriçao Atividade", "DSCATVEAP", 40, vbLeftJustify)
                     
                     
         IdCampo = Array("IDPROJ", "IDATVEAP")
         If IsMissing(Trim(pTitulo)) Then
            Tit$ = "Projeto"
         Else
            Tit$ = pTitulo
         End If
         
      Case "PDOCUMENTOS"
         '*******************
         '* Tabela Documentos
         '*******************
         If IsMissing(pQry) Then
            SQL = "SELECT P.IDPROJ, PR.NMPROJ, P.TIPODOC, P.NUMERODOC, P.TITULODOC FROM"
            SQL = SQL & " PDOCUMENTOS P INNER JOIN PROJETO PR ON"
            SQL = SQL & "     P.IDPROJ = PR.IDPROJ"
            SQL = SQL & " INNER JOIN PTIPODOC PT ON"
            SQL = SQL & "     P.TIPODOC = PT.TIPODOC"
            SQL = SQL & " ORDER BY PR.NMPROJ, P.TIPODOC, P.NUMERODOC"
         End If
         
         Cab = Array("IDPROJ", "IDPROJ", 0, vbLeftJustify, _
                     "Projeto", "NMPROJ", 10, vbLeftJustify, _
                     "Cód. Tipo", "TIPODOC", 5, vbLeftJustify, _
                     "Número", "NUMERODOC", 25, vbLeftJustify, _
                     "Título", "TITULODOC", 30, vbLeftJustify)
         IdCampo = Array("IDPROJ", "NMPROJ", "TIPODOC", "NUMERODOC")
         Tit$ = "Documentos"
      
      Case "PDRCODPROD"
         '**************
         '* Tabela Padrão de Cód. de Produto
         '**************
         Cab = Array("Seq.", "IDPDRCOD", 6, vbLeftJustify, _
                     "Código", "NMPDR", 20, vbLeftJustify, _
                     "Descrição", "DSCPDRCOD", 30, vbLeftJustify, _
                     "Padrão", "EPADRAO", 5, vbLeftJustify)
         IdCampo = Array("IDPDRCOD")
         Tit$ = "Códigos de Produto"
      
      Case "PRODETIQUETA"
         '**************
         '* Tabela PRODETIQUETA
         '**************
         If Not IsMissing(pQry) Then SQL = pQry
         If IsMissing(pCab) Then
            Cab = Array("Identificador", "CODETIQUETA", 25, vbLeftJustify, _
                        "Dt. Prev. Receb.", "DTPREVRECEBE", 15, vbLeftJustify, _
                        "Dt. Receb.", "DTRECEBE", 15, vbLeftJustify)
         End If
         IdCampo = Array("CODETIQUETA")
         If IsMissing(pTitulo) Then Tit$ = "TAGS"
         
      Case "PESCOPO"
         '**************
         '* Tabela PEscopo
         '**************
         Cab = Array("Código", "IDESCOPO", 5, vbLeftJustify, _
                     "DESCRICAO", "DSCESCOPO", 20, vbLeftJustify, _
                     "COD.AUXILIAR", "CODAUX", 8, vbLeftJustify)
                     
         IdCampo = Array("IDESCOPO", "DSCESCOPO", "CODAUX")
         Tit$ = "Escopo"
         
      Case "PESPECIFICACAO"
         '**************
         '* Tabela PESPECIFICACAO
         '**************
         If Not IsMissing(pQry) Then SQL = pQry
         If IsMissing(pCab) Then
            Cab = Array("Identificador", "IDESPECIFICACAO", 10, vbLeftJustify, _
                        "Código", "CODESPECIFICACAO", 20, vbLeftJustify)
         End If
         IdCampo = Array("IDESPECIFICACAO", "CODESPECIFICACAO")
         If IsMissing(pTitulo) Then Tit$ = "Especificações"
         
      Case "PTEMPERATURA_INSPECAO"
         '**************
         '* Tabela PTEMPERATURA_INSPECAO
         '**************
         If Not IsMissing(pQry) Then SQL = pQry
         If IsMissing(pCab) Then
            Cab = Array("IDTEMPERATURAINSPECAO", "IDTEMPERATURAINSPECAO", 0, vbLeftJustify, _
                        "IDNUMEROP", "IDNUMEROP", 0, vbLeftJustify, _
                        "IDCLASSEINSPECAO", "IDCLASSEINSPECAO", 0, vbLeftJustify, _
                        "Número P", "CODNUMEROP", 10, vbLeftJustify, _
                        "Classe", "CODCLASSEINSPECAO", 10, vbLeftJustify, _
                        "Temp. Inicial", "TEMPERATURAINICIAL", 10, vbLeftJustify, _
                        "Temp. Final", "TEMPERATURAFINAL", 10, vbLeftJustify, _
                        "Pressão Inicial", "PRESSAOINICIAL", 10, vbLeftJustify, _
                        "Pressão Final", "PRESSAOFINAL", 10, vbLeftJustify)
         End If
         IdCampo = Array("IDTEMPERATURAINSPECAO", "IDNUMEROP", "IDCLASSEINSPECAO", "CODNUMEROP", "CODCLASSEINSPECAO", "TEMPERATURAINICIAL", "TEMPERATURAFINAL", "PRESSAOINICIAL", "PRESSAOFINAL")
         If IsMissing(pTitulo) Then Tit$ = "Temperatura de Inspeção"
         
      Case "PCLASSIFICACAOACO"
         '**************
         '* Tabela PCLASSIFICACAOACO
         '**************
         If Not IsMissing(pQry) Then SQL = pQry
         If IsMissing(pCab) Then
            Cab = Array("Identificador", "IDCLASSIFICACAOACO", 0, vbLeftJustify, _
                        "Código", "CODCLASSIFICACAOACO", 35, vbLeftJustify, _
                        "Número P", "CODNUMEROP", 10, vbLeftJustify, _
                        "", "IDNUMEROP", 0, vbLeftJustify)
         End If
         IdCampo = Array("IDCLASSIFICACAOACO", "CODCLASSIFICACAOACO", "CODNUMEROP", "IDNUMEROP")
         If IsMissing(pTitulo) Then Tit$ = "Classificação do Aço"
         
      Case "PSERVICO_FLUIDO"
         '**************
         '* Tabela SERVICO
         '**************
         If Not IsMissing(pQry) Then SQL = pQry
         If IsMissing(pCab) Then
            Cab = Array("", "IDSERVICO", 0, vbLeftJustify, _
                        "Código", "CODSERVICO", 10, vbLeftJustify, _
                        "Descrição", "DSCSERVICO", 20, vbLeftJustify)
         End If
         IdCampo = Array("IDSERVICO", "CODSERVICO", "DSCSERVICO")
         If IsMissing(pTitulo) Then Tit$ = "Serviço"
      
      Case "PESSOA"
         '**************
         '* Tabela Grupo de Usuários
         '**************
         Cab = Array("Id.", "IDPESSOA", 5, vbLeftJustify, _
                     "Usuário", "CODPESSOA", 10, vbLeftJustify, _
                     "Nome", "NMPESSOA", 40, vbLeftJustify)
         IdCampo = Array("IDPESSOA", "CODPESSOA", "NMPESSOA")
         Tit$ = "Pessoas"
         SQL = 3
      
      Case "PIRECEBE"
         '**************
         '* Tabela de Inspeção
         '**************
         Cab = Array("Número", "NUMPI", 10, vbLeftJustify, _
                     "Ano", "ANOPI", 7, vbLeftJustify)
         IdCampo = Array("NUMPI", "ANOPI")
         Tit$ = "Pedido de Inspeção"
         
      'CRIADO POR EDUARDO MEDINA - CLASSE A CONSULTORIA
      'ATENDE PEDIDO DO CLIENTE SEPARAR INFORMAÇÕES EQUIPES
      Case "PMULTIEQUIPE"
         '*********************
         '* Tabela PMULTIEQUIPE
         '*********************
         Cab = Array("IDMULTIEQUIPE", "IDMULTIEQUIPE", 0, vbLeftJustify, _
                     "Nome das Equipes", "NMMULTIEQUIPE", 30, vbLeftJustify)
                     
         IdCampo = Array("IDMULTIEQUIPE", "NMMULTIEQUIPE")
         Tit$ = "EQUIPES CADASTRADAS"
         
      Case "PJUNTAEQUIV"
         '**************
         '* Tabela PJUNTAEQUIV
         '**************
         If Not IsMissing(pQry) Then SQL = pQry
         If IsMissing(pCab) Then
            Cab = Array("ID Junta Equiv.", "IDJUNTAEQUIV", 0, vbLeftJustify, _
                        "Diâmetro", "DIAMETRO", 10, vbLeftJustify, _
                        "Esp. Mínima", "ESPESSURAMIN", 10, vbLeftJustify, _
                        "Esp. Máxima", "ESPESSURAMAX", 10, vbLeftJustify, _
                        "Valor", "VALOR", 10, vbLeftJustify)
         End If
         IdCampo = Array("IDJUNTAEQUIV", "DIAMETRO", "ESPESSURAMIN", "ESPESSURAMAX", "VALOR")
         If IsMissing(pTitulo) Then Tit$ = "Juntas Equivalentes"
      
      Case "PO"
         '**************
         '* Tabela de Requisição de Materiais
         '**************
         Cab = Array("P.O.", "IDPO", 25, vbLeftJustify, _
                     "Rev.", "REV", 5, vbLeftJustify, _
                     "Descrição", "DSCPO", 40, vbLeftJustify) ', _
                     "Emissão", "DTEMISSAO", 8, vbLeftJustify)
         IdCampo = Array("IDPO", "REV")
         Tit$ = "P.O."
               
      Case "PRODUTO", "PRODUTOS"
         Dim cAux As String
         '**************
         '* PRODUTO
         '**************
         If pisTree Then
            MyLOV.CAMPO_ID = "IDPROD"
            MyLOV.CAMPO_CODIGO = "CODPROD"
            MyLOV.CAMPO_NOME = "NMPROD"
            MyLOV.ExibeCodigo = True
         End If

         Cab = Array("Seq.", "IDPROD", 6, vbLeftJustify, _
                     "Código", "CODPROD", 6, vbLeftJustify, _
                     "Descrição", "NMPROD", 50, vbLeftJustify)
         IdCampo = Array("IDPROD", "NMPROD", "CODPROD")
         'cAux = FrmCadProdDet.CmbGrupo.Text
         Tit$ = "Produtos"
         If IsMissing(pQry) Then
            If UCase(Tabela) = "PRODUTO" Then
               '* Lista de Valores chamada de Tela de detalhamento da carga (FrmCadProdDet)
               SQL = "Select IDPROD,CODPROD,NMPROD"
               SQL = SQL & " From PRODUTO "
               SQL = SQL & " Where Not(DTCADASTRO is Null)"
               'Sql = Sql & " and IDPAI = " & GetTag(FrmCadProdDet.MskCODPROD, "IDPAI")
               SQL = SQL & " order by NMPROD"
               Tit$ = "Produtos de " & Trim(Mid(cAux, InStr(cAux, "-") + 1))
            Else
               '* Lista de Valores chamada de Tela de detalhamento da carga (FrmCadProdDet)
               SQL = "Select IDPROD,CODPROD,NMPROD"
               SQL = SQL & " From PRODUTO "
               SQL = SQL & " Where Not(DTCADASTRO is Null)"
               SQL = SQL & " order by NMPROD"
            End If
         End If
         Tabela = "PRODUTO"
        
      Case "PROJETO"
         '**************
         '* Tabela Projetos
         '**************
         If IsMissing(pQry) Then
            SQL = "Select IDPROJ, CODPROJ, NMPROJ," & vbCr
            SQL = SQL & " STATUS =" & vbCr
            SQL = SQL & "     case when SITPROJ = 'C' then 'Concluído'" & vbCr
            SQL = SQL & "     when SITPROJ = 'P' then 'Paralizado'" & vbCr
            SQL = SQL & "     when SITPROJ = 'A' then 'Andamento'" & vbCr
            SQL = SQL & "     when SITPROJ = 'E' then 'Encerrado c/ Pendência'" & vbCr
            SQL = SQL & "     Else 'NA'" & vbCr
            SQL = SQL & "     End" & vbCr
            SQL = SQL & "     From PROJETO Order By 1" & vbCr
         End If
         If IsMissing(pCab) Then
            Cab = Array("Seq.", "IDPROJ", 6, vbLeftJustify, _
                        "Código", "CODPROJ", 10, vbLeftJustify, _
                        "Descrição", "NMPROJ", 30, vbLeftJustify, _
                        "Status", "STATUS", 25, vbLeftJustify)
         End If
         IdCampo = Array("IDPROJ", "CODPROJ", "NMPROJ", "STATUS")
         'If IsMissing(pIdCampo) Then IdCampo = Array("IDPROJ")
         If IsMissing(pTitulo) Then Tit$ = "Projetos"
         
      Case "PTABELAPREAQUEC"
         '**************
         '* Tabela PTABELAPREAQUEC
         '**************
         If Not IsMissing(pQry) Then SQL = pQry
         If IsMissing(pCab) Then
            Cab = Array("Identificador", "IDTABPREAQUEC", 10, vbLeftJustify, _
                        "Descrição", "DSCTABELAPREAQUEC", 35, vbLeftJustify)
         End If
         IdCampo = Array("IDTABPREAQUEC", "DSCTABELAPREAQUEC")
         If IsMissing(pTitulo) Then Tit$ = "Especificações"
      
      Case "RECURSO"
         '**************
         '* Tabela de Recursos
         '**************
         If IsMissing(pQry) Then
            SQL = "Select R.IDRECURSO, U.NMUSU"
            SQL = SQL & " From RECURSO R, USUARIO U"
            SQL = SQL & " Where R.IDRECURSO=U.IDUSU"
         End If
         
         Cab = Array("Código", "IDRECURSO", 10, vbLeftJustify, _
                     "Recurso", "NMUSU", 40, vbLeftJustify)
         IdCampo = Array("IDRECURSO", "NMUSU")
         Tit$ = "Recursos"
      
      Case "REQUISICAO"
         '**************
         '* Tabela de Requisição de Produtos
         '**************
         If Not IsMissing(pQry) Then SQL = pQry
         Cab = Array("Setor", "SETORREQ", 5, vbLeftJustify, _
                     "Nº Req.", "REQ", 9, vbLeftJustify, _
                     "Emissão", "DTEMISSAO", 8, vbLeftJustify, _
                     "C.Custo", "CODCCUSTO", 8, vbLeftJustify, _
                     "Requisitante", "SOLICITANTE", 25, vbLeftJustify, _
                     "Aplicação", "APLIC", 35, vbLeftJustify)
         IdCampo = Array("SETORREQ", "REQ")
         If IsMissing(pTitulo) Then
            Tit$ = "Requisição de Produtos"
         End If
         Merge = True
      
      Case "RESPONSAVEL"
         '**************
         '* Tabela de Responsáveis por Projeto
         '**************
         If IsMissing(pQry) Then
            SQL = "Select PR.IDRESPONSAVEL, P.NMPESSOA "
            SQL = SQL & " From PRESPPROJETO PR "
            SQL = SQL & " Inner Join PESSOA P ON PR.IDRESPONSAVEL = P.IDPESSOA"
         End If
         
         Cab = Array("Responsável", "IDRESPONSAVEL", 10, vbLeftJustify, _
                     "Nome", "NMPESSOA", 40, vbLeftJustify)
         IdCampo = Array("IDRESPONSAVEL", "NMPESSOA")
         Tit$ = "Responsáveis"
      
      Case "RM"
         '**************
         '* Tabela de Requisição de Materiais
         '**************
         Cab = Array("IDPROJ", "IDPROJ", 0, vbLeftJustify, _
                     "Projeto", "CODPROJ", 6, vbLeftJustify, _
                     "R.M.", "IDRM", 25, vbLeftJustify, _
                     "Rev.", "REV", 5, vbLeftJustify, _
                     "Descrição", "DSCRM", 40, vbLeftJustify) ', _
                     "Emissão", "DTEMISSAO", 8, vbLeftJustify)
         IdCampo = Array("IDPROJ", "IDRM", "REV")
         Tit$ = "R.M."
      
      Case "RR"
         '**************
         '* Tabela Relatorio de Recebimento RR
         '**************
         Cab = Array("R.R.", "IDRR", 5, vbLeftJustify, _
               "Ano", "ANORR", 4, vbLeftJustify, _
               "Data", "DTRR", 8, vbLeftJustify, _
               "Cliente", "NMCLI", 30, vbLeftJustify)
         IdCampo = Array("IDRR", "ANORR")
         Tit$ = "Relatório de Recebimento"
         If IsMissing(pQry) Then
            SQL = "Select R.IDRR, R.ANORR, R.DTRR, C.NMCLI "
            SQL = SQL & " From (RR R left join CLIENTE C "
            SQL = SQL & " on R.IDCLIENTE=C.IDCLIENTE) "
         End If
         
      Case "SETORES"
         '**************
         '* Tabela Setores
         '**************
         If IsMissing(pQry) Then
            SQL = "Select IDSETOR, CODSETOR, DSCSETOR"
            SQL = SQL & " From SETORES "
            SQL = SQL & " Where IDSETOR <> 0 "
            SQL = SQL & " Order By 2"
         End If
         
         Cab = Array("Seq.", "IDSETOR", 4, vbLeftJustify, _
                     "Cód Setor", "CODSETOR", 10, vbLeftJustify, _
                     "Descrição", "DSCSETOR", 35, vbLeftJustify, _
                     "Auxiliar", "IDAUXILIAR", 35, vbLeftJustify, _
                     "Usuário", "NMUSU", 35, vbLeftJustify)
         IdCampo = Array("IDSETOR", "CODSETOR", "IDAUXILIAR")
         Tit$ = "Setores"
      
      Case "SISTEMA"
         '**************
         '* Tabela Sistemas
         '**************
         Cab = Array("Código", "CODSISTEMA", 15, vbLeftJustify, _
                     "Descrição", "DSCSISTEMA", 60, vbLeftJustify)
         IdCampo = Array("CODSISTEMA", "DSCSISTEMA")
         Tit$ = "Sistemas"
         
      Case "SITUACAOPROD"
         '**************
         '* Tabela Setores
         '**************
         Cab = Array("Módulo", "CODSIS", 10, vbLeftJustify, _
                     "Código", "SITPROD", 8, vbLeftJustify, _
                     "Situação", "DSCSITPROD", 50, vbLeftJustify)
         IdCampo = Array("SITPROD", "DSCSITPROD")
         Tit$ = "Situação de Produtos"
      
      Case "SUBPROJETO"
         '**************
         '* Tabela SubProjetos
         '**************
         Cab = Array("Nº", "IDSUB", 7, vbLeftJustify, _
                     "Código", "CODSUB", 10, vbLeftJustify, _
                     "Descrição", "DSCSUB", 50, vbLeftJustify)
         IdCampo = Array("IDSUB", "CODSUB", "DSCSUB")
         Tit$ = "SubProjetos"
         
      Case "SUBSISTEMA"
         '**************
         '* Tabela SubSistemas
         '**************
         Cab = Array("Código", "CODSUBSISTEMA", 10, vbLeftJustify, _
                     "Descrição", "DSCSUBSISTEMA", 50, vbLeftJustify)
         IdCampo = Array("CODSUBSISTEMA", "DSCSUBSISTEMA")
         Tit$ = "SubSitemas"

      Case "SPA"
         '**************
         '* Tabela SPA
         '**************
         Cab = Array("Setor", "SETORSPA", 10, vbCenter, _
                     "NºSPA", "SPA", 10, vbCenter, _
                     "Ano", "ANOSPA", 5, vbCenter, _
                     "Data", "DTEMISSAO", 10, vbCenter, _
                     "Usuário", "NOME", 20, vbLeftJustify)
         IdCampo = Array("SETORSPA", "SPA", "ANOSPA")
         Tit$ = "S.P.A."
      
      
      Case "TPDOC"
         '**************
         '* Tabela Pais
         '**************
         Cab = Array("Código", "IDTPDOC", 8, vbLeftJustify, _
                     "Descrição", "DSCTPDOC", 40, vbLeftJustify, _
                     "Validade (Dias)", "VALIDADE", 13, vbLeftJustify, _
                     "Obrigatório", "OBRIGATORIO", 8, vbCenter)
         IdCampo = Array("IDTPDOC", "DSCTPDOC", "VALIDADE")
         Tit$ = "Documentos"

      Case "TPNOME"
         '**************
         '* Tabela Pais
         '**************
         Cab = Array("Código", "IDTPNOME", 8, vbLeftJustify, _
                     "Descrição", "DSCTPNOME", 40, vbLeftJustify)
         IdCampo = Array("IDTPNOME", "DSCTPNOME")
         Tit$ = "Classif. de Pessoas"
      
      Case "TRADUCAO"
         '**************
         '* Tabela de Palavras Traduzidas
         '**************
         Cab = Array("IDLNG", "IDLNG", 0, vbLeftJustify, _
                     "IDPALAVRA", "IDPALAVRA", 0, vbLeftJustify, _
                     "Palavra/Expressão", "TRADUCAO", 60, vbLeftJustify)
         IdCampo = Array("IDLNG", "IDPALAVRA", "TRADUCAO")
         Tit$ = "Palavra/Expressão"
     
     
      Case "UNIDADE"
         '**************
         '* Tabela Setores
         '**************
         Cab = Array("Sigla", "SIGLAUNID", 8, vbLeftJustify, _
                     "Unidade", "DSCUNID", 20, vbLeftJustify)
         IdCampo = Array("SIGLAUNID", "DSCUNID")
         Tit$ = "Unidades"
      
      Case "USUARIO"
         '**************
         '* Tabela Grupo de Usuários
         '**************
         Cab = Array("Usuário", "IDUSU", 10, vbLeftJustify, _
                     "Nome", "NMUSU", 40, vbLeftJustify)
         IdCampo = Array("IDUSU", "NMUSU")
         Tit$ = "Usuários"
            
      Case "GSITUACAO"
         '**************
         '* Tabela de Situacao
         '**************
         Cab = Array("Código", "CODSIT", 10, vbLeftJustify, _
                     "Descrição", "DSCSIT", 40, vbLeftJustify)
         IdCampo = Array("CODSIT", "DSCSIT")
         Tit$ = "Situações"
         
     Case "FACILIDADES"
         '**************
         '* Tabela Setores
         '**************
         Cab = Array("Id.", "IDFACILIDADE", 8, vbLeftJustify, _
                     "Descrição", "DSCFACILIDADE", 20, vbLeftJustify)
         IdCampo = Array("IDFACILIDADE", "DSCFACILIDADE")
         Tit$ = "Facilidades"
         
      Case "PCRONOGRAMA"
         '**************
         '* Tabela CRONOGRAMA
         '**************
         Cab = Array("IDPROJ", "IDPROJ", 0, vbLeftJustify, _
                     "Código", "CODCRONOGRAMA", 10, vbLeftJustify, _
                     "Descrição", "DSCCRONOGRAMA", 20, vbLeftJustify, _
                     "Dt. Inicio", "DTINI", 8, vbLeftJustify, _
                     "Dt. Fim", "DTFIM", 8, vbLeftJustify)
                                          
         IdCampo = Array("IDPROJ", "CODCRONOGRAMA", "DSCCRONOGRAMA")
         
         Tit$ = "Cronograma"
      
      Case "TIPOCORTE"
         '**************
         '* Tabela TIPOCORTE
         '**************
         Cab = Array("CODTPC", "CODTPC", 10, vbLeftJustify, _
                     "DSCTPC", "DSCTPC", 20, vbLeftJustify)
                                          
         IdCampo = Array("CODTPC", "DSCTPC")
         
         Tit$ = "Tipo de Corte"
         
      Case "PTPMATERIAL"
         '**************
         '* Tabela Tipo de Material
         '**************
         Cab = Array("Código", "CODTPMATERIAL", 10, vbLeftJustify, _
                     "Compr.", "COMPRTPROD", 10, vbLeftJustify, _
                     "Tenacidade", "NVLTENACIDADE", 10, vbLeftJustify)
         IdCampo = Array("CODTPMATERIAL")
         Tit$ = "Tipo de Material"
         
      Case "TIPOPROCESSO"
         '**************
         '* Tabela TIPOPROCESSO
         '**************
         If IsMissing(pQry) Then
            SQL = ""
            SQL = SQL & "SELECT  D.DSCDISCIPLINA, T.IDDISCIPLINA, T.TPPROCESSO, T.DSCPROCESSO "
            SQL = SQL & " FROM       TIPOPROCESSO T"
            SQL = SQL & " INNER JOIN DISCIPLINA D ON T.IDDISCIPLINA = D.IDDISCIPLINA"
           'Sql = Sql & " Order By T.IDDISCIPLINA,T.ORDEM,T.TPPROCESSO"
            SQL = SQL & " Order By D.DSCDISCIPLINA,T.ORDEM,T.DSCPROCESSO"
         End If
         
         Dim tamanhoCol As Integer
         
         If IsMissing(pQry) Then
            tamanhoCol = 15
         Else
            tamanhoCol = 0
         End If
         
         Cab = Array("Disciplina", "DSCDISCIPLINA", tamanhoCol, vbLeftJustify, _
                     "IDDISC", "IDDISCIPLINA", 0, vbLeftJustify, _
                     "Código Processo", "TPPROCESSO", 9, vbLeftJustify, _
                     "Descrição", "DSCPROCESSO", 30, vbLeftJustify)
         IdCampo = Array("DSCDISCIPLINA", "IDDISCIPLINA", "TPPROCESSO", "DSCPROCESSO")
         
         If IsMissing(pTitulo) Then
            Tit$ = "Tipo de Processo"
         Else
            Tit$ = Trim(pTitulo)
         End If
         
      Case "TIPOSOLDA"
         '**************
         '* Tabela Tipo Solda
         '**************
         Cab = Array("Código", "TPSOLDA", 5, vbLeftJustify, _
                     "Descrição", "DSCTPSOLDA", 20, vbLeftJustify)
         IdCampo = Array("TPSOLDA", "DSCTPSOLDA")
         Tit$ = "Tipo de Solda"
         
      Case "TIPOTRATAMENTO"
         '**************
         '* Tabela Tipo de Tratamento
         '**************
         Cab = Array("Código", "IDTRAT", 10, vbLeftJustify, _
                     "Descrição", "DSCTRAT", 30, vbLeftJustify)
         IdCampo = Array("IDTRAT", "DSCTRAT")
         Tit$ = "Tipos de Tratamentos"
         
      Case "SUBDISCIPLINA"
         '**************
         '* Tabela Sub-Disciplina
         '**************
         If IsMissing(pQry) Then
            SQL = "SELECT     S.IDSUBDISC, S.DSCSUBDISC, D.DSCDISCIPLINA "
            SQL = SQL & " FROM       SUBDISCIPLINA S"
            SQL = SQL & " INNER JOIN DISCIPLINA D ON S.IDDISCIPLINA = D.IDDISCIPLINA"
         End If
         
         If IsMissing(pCab) Then
            Cab = Array("Código", "IDSUBDISC", 8, vbLeftJustify, _
                        "Descrição", "DSCSUBDISC", 22, vbLeftJustify, _
                        "Disciplina", "DSCDISCIPLINA", 20, vbLeftJustify)
         End If
         
         If IsMissing(pIdCampo) Then IdCampo = Array("IDSUBDISC", "DSCSUBDISC", "DSCDISCIPLINA")
         
         Tit$ = "Situação de Produtos"
      
      Case "MOT_ROMANEIO"
         '**************
         '* Tabela de Motivo de Romaneio
         '**************
         Cab = Array("Id", "IDMOTIVO", 15, vbLeftJustify, _
                     "Descrição", "DSCMOTIVO", 30, vbLeftJustify)
         
         IdCampo = Array("IDMOTIVO", "DSCMOTIVO")
         Tit$ = "Motivo de Romaneio"
         
      Case "OFICINA"
         '**************
         '* Tabela OFICINA
         '**************
         If IsMissing(pQry) Then
            SQL = "Select O.IDOFICINA, O.CODOFICINA, O.DSCOFICINA,C.NOME"
            SQL = SQL & " From OFICINAS O left join CONTRATADA C"
            SQL = SQL & " on O.IDCONTRATADA = C.IDCONTRATADA"
            SQL = SQL & " Order By C.NOME"
         End If
         
         Cab = Array("Seqüência", "IDOFICINA", 6, vbRightJustify, _
                     "Código", "CODOFICINA", 10, vbLeftJustify, _
                     "Descrição", "DSCOFICINA", 25, vbLeftJustify, _
                     "Contratada", "NOME", 16, vbLeftJustify)
                     
         IdCampo = Array("IDOFICINA", "CODOFICINA", "DSCOFICINA", "CONTRATADA")
         Tit$ = "Oficinas"
         
      Case "PTIPOOP"
         '**************
         '* Tabela de Funções
         '**************
         Cab = Array("IDTIPOOP", "IDTIPOOP", 0, vbLeftJustify, _
                     "Descrição", "DSCTIPOOP", 50, vbLeftJustify, _
                     "Código", "CODAUX", 10, vbLeftJustify)
         IdCampo = Array("IDTIPOOP")
         Tit$ = "TIPO OP"
         Tabela = "PTIPOOP"
         
      Case "EQUIPAMENTO"
         '**************
         '* Tabela de Equipamentos
         '**************
         If IsMissing(pQry) Then
            SQL = ""
            SQL = SQL & "Select CODEQUIPAMENTO, DESCRICAO"
            SQL = SQL & " From EQUIPAMENTO "
         End If
         Cab = Array("Código", "CODEQUIPAMENTO", 15, vbLeftJustify, _
                   "Descrição", "DESCRICAO", 30, vbLeftJustify)
         IdCampo = Array("CODEQUIPAMENTO", "DESCRICAO")
         Tit$ = "Equipamentos"
            
      Case "PNIVELCLASSEEAP"
         '**************
         '* Tabela de Nível de Classe
         '**************
         Cab = Array("IDPROJ", "IDPROJ", 0, vbLeftJustify, _
                     "IDNIVELCLS", "IDNIVELCLS", 0, vbLeftJustify, _
                     "Código", "CODNIVELCLS", 10, vbLeftJustify, _
                     "Descrição", "DSCNIVELCLS", 20, vbLeftJustify)
                     
         IdCampo = Array("IDPROJ", "IDNIVELCLS")
         Tit$ = "Níveis de Classe EAP"
            
      Case "PATIVIDADEOP"
         '**************
         '* Tabela ATIVIDADES OP
         '**************
         Cab = Array("IDATV", "IDATV", 0, vbLeftJustify, _
                     "Cód. ATIV.", "CODATIVIDADE", 25, vbLeftJustify, _
                     "Descrição", "DSCATIVIDADE", 50, vbLeftJustify)
         IdCampo = Array("IDATV")
         Tit$ = "Atividades"
      
      Case "PCLASSEEAP"
         '**************
         '* Tabela CLASSES EAP
         '**************
         Cab = Array("IDPROJ", "IDPROJ", 0, vbLeftJustify, _
                     "IDNIVELCLS", "IDNIVELCLS", 0, vbLeftJustify, _
                     "IDCLASSEEAP", "IDCLASSEEAP", 0, vbLeftJustify, _
                     "Código", "CODCLASSEEAP", 10, vbLeftJustify, _
                     "Descrição", "DSCCLASSEEAP", 20, vbLeftJustify)
                     
         IdCampo = Array("IDPROJ", "IDNIVELCLS", "IDCLASSEEAP")
         Tit$ = "Classes EAP"
         
      Case "PMODELO"
         '**************
         '* Tabela de OP
         '**************
         Cab = Array("IDMODELO", "IDMODELO", 0, vbLeftJustify, _
                     "Nome", "NMMODELO", 25, vbLeftJustify, _
                     "Unid. Fator", "UNIDFATOR", 7, vbLeftJustify, _
                     "Valor", "VLRFATOR", 7, vbRightJustify)
                     
         IdCampo = Array("IDMODELO")
         
         Tit$ = "Modelos de HH"
         
      Case "DEFEITOSOLDA"
         '**************
         '* Tabela Defeitos de Solda
         '**************
         
         Cab = Array("Código", "CODDEF", 5, vbLeftJustify, _
                     "Descrição", "DSCDEF", 30, vbLeftJustify, _
                     "Descrição Inglês", "DSCDEFING", 30, vbLeftJustify)
         IdCampo = Array("CODDEF", "DSCDEF", "DSCDEFING")
         Tit$ = "Defeitos de Solda"
         
      Case "CATINSPECAO"
         '**************
         '* Tabela Categoria de Inspeção
         '**************
         Cab = Array("Código", "CODCAT", 5, vbLeftJustify, _
                     "Descrição", "DSCCAT", 30, vbLeftJustify)
         IdCampo = Array("CODCAT", "DSCCAT")
         Tit$ = "Categoria de Inspeção"
         
      Case "TIPOENSAIO"
         '**************
         '* Tabela Tipo de Ensaio
         '**************
         Cab = Array("Código", "CODTPENSAIO", 6, vbLeftJustify, _
                     "Descrição", "DSCINSP", 25, vbLeftJustify)
         IdCampo = Array("CODTPENSAIO", "DSCINSP")
         Tit$ = "Tipo de Ensaio"
         
      Case "TIPOJUNTA"
         '**************
         '* Tabela Tipos de Junta de Solda
         '**************
         Cab = Array("Código", "CODTPJUNTA", 10, vbLeftJustify, _
                     "Descrição", "DSCTPJUNTA", 30, vbLeftJustify)
         IdCampo = Array("CODTPJUNTA", "DSCTPJUNTA")
         Tit$ = "Tipos de Junta de Solda"
         
      Case "TPPROCSOLDA"
         '**************
         '* Tabela Tipo de Processo de Solda
         '**************
         Cab = Array("Código", "PROCSOLDA", 5, vbLeftJustify, _
                     "Descrição", "DESCPROCSOLDA", 30, vbLeftJustify)
         IdCampo = Array("PROCSOLDA", "DESCPROCSOLDA")
         Tit$ = "Tipo de Processos de Soldagem"
         
      Case "FERRAMENTAS"
         '**************
         '* Tabela de Ferramentas
         '**************
         
         Cab = Array("Código", "CODFERR", 5, vbLeftJustify, _
                     "Dt. Aferição", "DTAFER", 10, vbLeftJustify, _
                     "Dt. Validade", "DTVALID", 10, vbLeftJustify, _
                     "Descrição", "DSCFERR", 30, vbLeftJustify)
         IdCampo = Array("CODFERR", "DTAFER", "DTVALID", "DSCFERR")
         Tit$ = "Ferramentas"
         
      Case "SOLDADOR"
         '**************
         '* Tabela de Soldador
         '**************
         Cab = Array("Sinete", "SINETE", 5, vbLeftJustify, _
                     "Nome", "SOLDNOME", 30, vbLeftJustify, _
                     "Norma", "Norma", 6, vbLeftJustify)
         IdCampo = Array("SINETE", "SOLDNOME", "Norma")
         Tit$ = "Cadastro de Soldador"
         
      Case "TIPOATRSINSP"
         '**************
         '* Tabela TIPOATRSINSP
         '**************
         Cab = Array("Código", "CODATRASO", 5, vbLeftJustify, _
                     "Descrição", "DESCATRASO", 30, vbLeftJustify, _
                     "Relevante", "RELEVANTE", 6, vbLeftJustify)
         IdCampo = Array("CODATRASO", "DESCATRASO")
         Tit$ = "Atrasos de Inspeção"
         
      Case "CERTIFICADORAS"
         '**************
         '* Tabela Contratada
         '**************
         Cab = Array("Cód.", "CODCERTF", 5, vbLeftJustify, _
                     "Nome", "NOMECERTIF", 30, vbLeftJustify, _
                     "Tipo", "TIPOCERTF", 5, vbLeftJustify)
         IdCampo = Array("CODCERTF")
         Tit$ = "Certificadora"
         
      Case "NORMASCQ"
         '**************
         '* Tabela Contratada
         '**************
         Cab = Array("Tipo", "TIPONORMA", 5, vbLeftJustify, _
                     "Norma", "CODNORMA", 10, vbLeftJustify, _
                     "Descrição", "DESCRNORMA", 20, vbLeftJustify)
         IdCampo = Array("TIPONORMA", "CODNORMA")
         Tit$ = "Normas CQ"
         
      Case "INSPETOR"
         '**************
         '* Tabela de Eventos
         '**************
         Cab = Array("Código", "CODINSPETOR", 8, vbLeftJustify, _
                     "Nome", "NOMEINSPETOR", 15, vbLeftJustify, _
                     "Especialidade", "ESPECIALIDADE", 12, vbLeftJustify)
         IdCampo = Array("CODINSPETOR", "NOMEINSPETOR", "ESPECIALIDADE")
         Tit$ = "Inspetores"
      
      Case "IEIS"
         '**********************************************
         '* Tabela de IEIS *
         '**********************************************
         Cab = Array("Código", "CODIEIS", 10, vbLeftJustify, _
                      "RQPS", "NUMRQPS", 10, vbLeftJustify, _
                      "EPS", "NUMEPS", 10, vbLeftJustify, _
                      "Revisão", "REVIEIS", 7, vbLeftJustify)
         IdCampo = Array("CODIEIS", "REVIEIS")
         Tit$ = "IEIS"
         
       Case "INSPECAOTABELA"
         '**************
         '* Tabela de Eventos
         '**************
         Cab = Array("Tabela", "IDTABINSPECAO", 8, vbLeftJustify, _
                     "Revisao", "REVTAB", 8, vbLeftJustify, _
                     "Data de Revisão", "DTREV", 10, vbLeftJustify)
         IdCampo = Array("IDTABINSPECAO", "REVTAB", "DTREV")
         Tit$ = "Tabelas de Inspeção"
         
      Case "NORMALINHA"
         '**************
         '* Tabela de Normas / Linhas
         '**************
         If IsMissing(pQry) Then
            SQL = ""
            SQL = SQL & "Select IDLinha, VlLinha,TipoDado"
            SQL = SQL & " from PNORMALINHA "
         End If
         
         Cab = Array("Linha", "IDLinha", 8, vbLeftJustify, _
                     "Valor", "VlLinha", 8, vbLeftJustify, _
                     "Tipo Dado", "TipoDado", 10, vbLeftJustify)
         IdCampo = Array("IDLinha", "VlLinha", "TipoDado")
         Tit$ = "Normas / Linhas"
       
      Case "NORMAS"
         '**************
         '* Tabela de Normas
         '**************
         If IsMissing(pQry) Then
            SQL = ""
            SQL = SQL & "Select IDNORMA,DSCNORMA"
            SQL = SQL & " from PNORMAS "
         End If
         
         Cab = Array("Código", "IDNORMA", 8, vbLeftJustify, _
                     "Descrição", "DSCNORMA", 16, vbLeftJustify)
                     
         IdCampo = Array("IDNORMA", "DSCNORMA")
         Tit$ = "Normas"
         
      Case "NORMACOLUNA"
         '***********************
         '* Tabela de NormaColuna
         '***********************
         If IsMissing(pQry) Then
            SQL = ""
            SQL = SQL & "Select IDCOLUNA,NMCOLUNA"
            SQL = SQL & " from PNORMACOLUNA "
         End If
         
         Cab = Array("Código", "IDCOLUNA", 8, vbLeftJustify, _
                     "Descrição", "NMCOLUNA", 16, vbLeftJustify)
                     
         IdCampo = Array("IDCOLUNA", "NMCOLUNA")
         Tit$ = "Normas/Colunas "
      
      Case "VLCOLUNA"
         '*****************************
         '* Tabela de Valores p/ Coluna
         '*****************************
         If IsMissing(pQry) Then
            SQL = ""
            SQL = SQL & "Select V.IDVALOR,V.DESCRICAO,C.NmColuna "
            SQL = SQL & "from PNORMAVLCOLUNA V inner join PnormaColuna C "
            SQL = SQL & "on "
            SQL = SQL & "V.IdColuna = C.IdColuna "
            SQL = SQL & "order by NmColuna "
            
           'Mostra o Nome da Coluna  *****************
            Cab = Array("Código", "IDVALOR", 0, vbLeftJustify, _
                 "Coluna", "NmColuna", 10, vbLeftJustify, _
                 "Descrição", "DESCRICAO", 40, vbLeftJustify)
         Else
           'O Nome da Coluna é exibido no Título *****
            Cab = Array("Código", "IDVALOR", 0, vbLeftJustify, _
                 "Coluna", "NmColuna", 0, vbLeftJustify, _
                 "Descrição", "DESCRICAO", 40, vbLeftJustify)
         End If
         
         
         IdCampo = Array("IDVALOR", "NmColuna", "DESCRICAO")
         Tit$ = pTitulo
      
      Case "CATEGORIAOS"
         '**************
         '* Tabela "CATEGORIA"
         '**************
         Cab = Array("Código", "IDCATEGORIAOS", 10, vbLeftJustify, _
                     "Nome", "NMCATEGORIAOS", 30, vbLeftJustify)
                     
                                          
         IdCampo = Array("IDCATEGORIAOS", "NMCATEGORIAOS")
         Tit$ = "Categorias"
      
      Case "PMODELOAVANCO"
         '**************
         '* Tabela CRONOGRAMA
         '**************
         Cab = Array("IDPROJ", "IDPROJ", 0, vbLeftJustify, _
                     "IDMODELOAVANCO", "IDMODELOAVANCO", 0, vbLeftJustify, _
                     "Descrição", "DSCMODELOAVANCO", 35, vbLeftJustify)
                                          
         IdCampo = Array("IDPROJ", "IDMODELOAVANCO", "DSCMODELOAVANCO")
         
         Tit$ = "Modelo de Avanço"
      
      
      Case "PROPRIETARIO"
         '**************
         '* Tabela PROPRIETARIO
         '**************
         Cab = Array("ID", "IDPESSOA", 0, vbLeftJustify, _
                     "Código", "CODPROP", 10, vbLeftJustify, _
                     "Nome", "NMPROP", 35, vbLeftJustify)
                                          
         IdCampo = Array("IDPESSOA", "CODPROP", "NMPROP")
         
         Tit$ = "Proprietário"
      
      Case "PLOCALARMAZANAGEM"
         '**********************************
         '* Tabela de Local de Armazenamento
         '**********************************
         If IsMissing(pQry) Then
            SQL = ""
            SQL = SQL & "Select P.IDLOCAL, P.CODLOCAL, P.DSCLOCAL"
            SQL = SQL & "  From PLOCALARMAZENAGEM P"
            SQL = SQL & " Order By P.CODLOCAL"
         End If
         
         Cab = Array("IDLOCAL", "IDLOCAL", 0, vbLeftJustify, _
                     "Codigo", "CODLOCAL", 10, vbLeftJustify, _
                     "Descrição", "DSCLOCAL", 30, vbLeftJustify)
         IdCampo = Array("IDLOCAL", "CODLOCAL", "DSCLOCAL")
         Tit$ = "Local de Armazanamento"
         
      Case "PDOCREVISAO"
         '****************
         '* Tabela Documento de Revisão
         '****************
         Cab = Array("Id", "IDDOCREVISAO", 0, vbLeftJustify, _
                     "Nº Documento", "NUMERODOCREV", 20, vbLeftJustify, _
                     "Tipo", "TIPODOCREV", 10, vbLeftJustify, _
                     "Dt. Emissão", "DTEMISSAO", 10, vbLeftJustify)
         IdCampo = Array("IDDOCREVISAO", "NUMERODOCREV", "TIPODOCREV", "DTEMISSAO")
         Tit$ = "Documento de Revisão"
         
      Case "PERFILSOLDA"
         '**************
         '* Tabela Perfil de Solda
         '**************
         Cab = Array("Perfil Soldado", "CODPS", 10, vbLeftJustify, _
                     "Revisão", "REV", 10, vbRightJustify)
         IdCampo = Array("CODPS", "REV")
         Tit$ = "Perfil Soldado"
         
      Case "PLANOCORTE"
         '**************
         '* Tabela Plano de Corte
         '**************
         Cab = Array("Plano de Corte", "CODPC", 10, vbLeftJustify, _
                     "Revisão", "REV", 10, vbRightJustify)
         IdCampo = Array("CODPC", "REV")
         Tit$ = "Plano de Corte"
         
       Case "DETFAB"
         '**************
         '* Tabela Detalhe de Fabricação
         '**************
         Cab = Array("Código", "CODDF", 10, vbLeftJustify, _
                     "Revisão", "REV", 10, vbRightJustify)
         IdCampo = Array("CODDF", "REV")
         Tit$ = "Detalhe de Fabricação"

      Case "DMONTAGEM"
         '**************
         '* Tabela Detalhe de Fabricação
         '**************
         Cab = Array("Diag. Montagem", "CODDM", 10, vbLeftJustify, _
                     "Revisão", "REV", 10, vbRightJustify)
         IdCampo = Array("CODDM", "REV")
         Tit$ = "Diagrama de Montagem"

      Case "PFOLHAVERIFICACAO"
         Cab = Array("F.V.I", "CODFOLHA", 20, vbLeftJustify)
         IdCampo = Array("CODFOLHA")
         Tit$ = "Folha de Verificação"
      
      Case "POPRODUCAO"
         Cab = Array("Nº da O.P", "CODOP", 20, vbLeftJustify)
         IdCampo = Array("CODOP")
         Tit$ = "Número da Ordem de Produção"

      Case "POSERVICO"
         Cab = Array("Nº do J.C", "CODOS", 20, vbLeftJustify)
         IdCampo = Array("CODOS")
         Tit$ = "Número do Job Card"
      
      
      Case "ESTOQUE"
         Cab = Array("ID", "IDEST", 4, vbLeftJustify, _
                     "Descrição", "DSCEST", 20, vbLeftJustify)
         IdCampo = Array("IDEST", "DSCEST")
         Tit$ = "Estoque"
      
      Case "USUARIO_SETORES"
          '**********************************
         '* Tabela de Local de Usuario_Setor
         '**********************************
         If IsMissing(pQry) Then
            SQL = ""
            SQL = SQL & "SELECT U.IDUSU,U.NMUSU,S.CODSETOR, (S.CODSETOR + '-' + S.DSCSETOR) AS SETOR "
            SQL = SQL & " FROM USUARIO U"
            SQL = SQL & " LEFT JOIN SETORES S"
            SQL = SQL & " ON U.IDSETOR = S.IDSETOR"
            SQL = SQL & " ORDER BY U.NMUSU"
         End If
         
         Cab = Array("ID", "IDUSU", 6, vbLeftJustify, _
                     "Nome", "NMUSU", 15, vbLeftJustify, _
                     "Cod Setor", "CODSETOR", 5, vbLeftJustify, _
                     "Setor", "SETOR", 15, vbLeftJustify)
         
         IdCampo = Array("IDUSU", "NMUSU", "CODSETOR", "SETOR")
         Tit$ = "Usuário/Setor"
      
      Case "" 'Lista de Valores Livre
      Case Else
         F_LOV = Empty
         GoTo Saida
   End Select
   
   If IsMissing(pisTree) Then pisTree = False
   With MyLOV
      .Aplic = App
      '.Idioma = Sys.Idioma
      '.FundoTela = Sys.FundoTela
      .Tipo = "LOV"
      
      .dBase = pXDb
      .Table = UCase(Tabela)
      .Query = SQL
      .Cab = Cab
      .IdField = IdCampo
      .MultRows = MultRows
      .Merge = Merge
      .Caption = Tit$
      .isTree = pisTree
      .WidthScr = WidthScr
      .Show
      F_LOV = .ID
   End With
   Set MyLOV = Nothing
Saida:
   
   Screen.MousePointer = mPointer
End Function

Public Function f_ExisteNoGrid(pGrd As iGrid, pCol As Long, pConteudo As Variant) As Boolean
   Dim n
   
   f_ExisteNoGrid = False
   For n = 1 To pGrd.RowCount
      If Trim(pGrd.CellValue(n, pCol)) = Trim(pConteudo) Then
         f_ExisteNoGrid = True
         Exit For
      End If
   Next
End Function

