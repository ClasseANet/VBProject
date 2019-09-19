VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_USUARIO"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb            As Object 
Private mvarRS             As Object 

Private mvarQryInsert      As String
Private mvarQryUpDate      As String
Private mvarQryDelete      As String
Private mvarQrySelect      As String
Private mvarQrySave        As String

Private mvarisDirt         As Boolean

Private mvarIDUSU         As String
Private mvarIDGRUPO       As String
Private mvarSENHAUSU      As String
Private mvarNMUSU         As String
Private mvarDTCADASTRO    As String
Private mvarDTEXPIRAUSU   As String
Private mvarDTEXPIRASENHA As String
Private mvarTPUSU         As String
Private mvarSITUSU        As String
Private mvarUSUCADASTRO   As String
Private mvarFEXPIRASENHA  As Double
Private mvarMATRUSU       As String
Private mvarICORREIO      As String
Private mvarECORREIO      As String
Private mvarASSINATURA    As Object
Private mvarTIMEOUT       As Double
Private mvarIDSETOR       As Double
Private mvarIDEMPR        As Double
Private mvarLOCALTRAB     As String
Private mvarNIVELOS       As Double
Private mvarNMAD          As String
Private mvarIDPESSOA      As Double
Private mvarIDCOLIGADA    As Double
Private mvarDTVALIDAUSU   As String
Private mvarTROCARSENHA   As Double
Public Property Let IDUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDUSU <> vData)
   mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
End Property
Public Property Let IDGRUPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDGRUPO <> vData)
   mvarIDGRUPO = vData
End Property
Public Property Get IDGRUPO() As String
   IDGRUPO = mvarIDGRUPO
End Property
Public Property Let SENHAUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSENHAUSU <> vData)
   mvarSENHAUSU = vData
End Property
Public Property Get SENHAUSU() As String
   SENHAUSU = mvarSENHAUSU
End Property
Public Property Let NMUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMUSU <> vData)
   mvarNMUSU = vData
End Property
Public Property Get NMUSU() As String
   NMUSU = mvarNMUSU
End Property
Public Property Let DTCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTCADASTRO <> vData)
   mvarDTCADASTRO = vData
End Property
Public Property Get DTCADASTRO() As String
   DTCADASTRO = mvarDTCADASTRO
End Property
Public Property Let DTEXPIRAUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEXPIRAUSU <> vData)
   mvarDTEXPIRAUSU = vData
End Property
Public Property Get DTEXPIRAUSU() As String
   DTEXPIRAUSU = mvarDTEXPIRAUSU
End Property
Public Property Let DTEXPIRASENHA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEXPIRASENHA <> vData)
   mvarDTEXPIRASENHA = vData
End Property
Public Property Get DTEXPIRASENHA() As String
   DTEXPIRASENHA = mvarDTEXPIRASENHA
End Property
Public Property Let TPUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPUSU <> vData)
   mvarTPUSU = vData
End Property
Public Property Get TPUSU() As String
   TPUSU = mvarTPUSU
End Property
Public Property Let SITUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSITUSU <> vData)
   mvarSITUSU = vData
End Property
Public Property Get SITUSU() As String
   SITUSU = mvarSITUSU
End Property
Public Property Let USUCADASTRO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUSUCADASTRO <> vData)
   mvarUSUCADASTRO = vData
End Property
Public Property Get USUCADASTRO() As String
   USUCADASTRO = mvarUSUCADASTRO
End Property
Public Property Let FEXPIRASENHA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFEXPIRASENHA <> vData)
   mvarFEXPIRASENHA = vData
End Property
Public Property Get FEXPIRASENHA() As Double
   FEXPIRASENHA = mvarFEXPIRASENHA
End Property
Public Property Let MATRUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMATRUSU <> vData)
   mvarMATRUSU = vData
End Property
Public Property Get MATRUSU() As String
   MATRUSU = mvarMATRUSU
End Property
Public Property Let ICORREIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarICORREIO <> vData)
   mvarICORREIO = vData
End Property
Public Property Get ICORREIO() As String
   ICORREIO = mvarICORREIO
End Property
Public Property Let ECORREIO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarECORREIO <> vData)
   mvarECORREIO = vData
End Property
Public Property Get ECORREIO() As String
   ECORREIO = mvarECORREIO
End Property
Public Property Set ASSINATURA(ByVal vData As Object)
   If Not mvarisDirt Then mvarisDirt = (mvarASSINATURA <> vData)
   Set mvarASSINATURA = vData
End Property
Public Property Get ASSINATURA() As Object
   Set ASSINATURA = mvarASSINATURA
End Property
Public Property Let TIMEOUT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTIMEOUT <> vData)
   mvarTIMEOUT = vData
End Property
Public Property Get TIMEOUT() As Double
   TIMEOUT = mvarTIMEOUT
End Property
Public Property Let IDSETOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSETOR <> vData)
   mvarIDSETOR = vData
End Property
Public Property Get IDSETOR() As Double
   IDSETOR = mvarIDSETOR
End Property
Public Property Let IDEMPR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDEMPR <> vData)
   mvarIDEMPR = vData
End Property
Public Property Get IDEMPR() As Double
   IDEMPR = mvarIDEMPR
End Property
Public Property Let LOCALTRAB(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarLOCALTRAB <> vData)
   mvarLOCALTRAB = vData
End Property
Public Property Get LOCALTRAB() As String
   LOCALTRAB = mvarLOCALTRAB
End Property
Public Property Let NIVELOS(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNIVELOS <> vData)
   mvarNIVELOS = vData
End Property
Public Property Get NIVELOS() As Double
   NIVELOS = mvarNIVELOS
End Property
Public Property Let NMAD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMAD <> vData)
   mvarNMAD = vData
End Property
Public Property Get NMAD() As String
   NMAD = mvarNMAD
End Property
Public Property Let IDPESSOA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPESSOA <> vData)
   mvarIDPESSOA = vData
End Property
Public Property Get IDPESSOA() As Double
   IDPESSOA = mvarIDPESSOA
End Property
Public Property Let IDCOLIGADA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCOLIGADA <> vData)
   mvarIDCOLIGADA = vData
End Property
Public Property Get IDCOLIGADA() As Double
   IDCOLIGADA = mvarIDCOLIGADA
End Property
Public Property Let DTVALIDAUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTVALIDAUSU <> vData)
   mvarDTVALIDAUSU = vData
End Property
Public Property Get DTVALIDAUSU() As String
   DTVALIDAUSU = mvarDTVALIDAUSU
End Property
Public Property Let TROCARSENHA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarTROCARSENHA <> vData)
   mvarTROCARSENHA = vData
End Property
Public Property Get TROCARSENHA() As Double
   TROCARSENHA = mvarTROCARSENHA
End Property
Public Property Get isDirt() As Boolean
   isDirt = mvarisDirt
End Property
Public Property Set XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Let XDb(ByVal vData As Object)
   Set mvarXDb = vData
End Property
Public Property Get XDb() As Object
   Set XDb = mvarXDb
End Property
Public Property Get RS() As Object
   Set RS = mvarRS
End Property
Public Property Get QryInsert(Optional pAutoId as Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Insert Into USUARIO (IDUSU, IDGRUPO, SENHAUSU, NMUSU" & VbNewLine 
   Sql = Sql & ", DTCADASTRO, DTEXPIRAUSU, DTEXPIRASENHA, TPUSU, SITUSU" & VbNewLine 
   Sql = Sql & ", USUCADASTRO, FEXPIRASENHA, MATRUSU, ICORREIO, ECORREIO, TIMEOUT, IDSETOR, IDEMPR, LOCALTRAB" & VbNewLine 
   Sql = Sql & ", NIVELOS, NMAD, IDPESSOA, IDCOLIGADA, DTVALIDAUSU" & VbNewLine 
   Sql = Sql & ", TROCARSENHA) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDGRUPO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSENHAUSU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMUSU) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTEXPIRAUSU, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTEXPIRASENHA, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTPUSU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSITUSU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUSUCADASTRO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarFEXPIRASENHA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMATRUSU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarICORREIO) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarECORREIO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarTIMEOUT) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDSETOR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDEMPR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarLOCALTRAB) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarNIVELOS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMAD) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPESSOA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTVALIDAUSU, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarTROCARSENHA) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDUSU, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From USUARIO" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDUSU)  Then 
      Sql = Sql & " IDUSU = " & SqlStr(mvarIDUSU) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDUSU) Then Sql = Sql & " IDUSU = " & SqlStr(Cstr(Ch_IDUSU)) & " AND " & VbNewLine 
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & VbNewLine
      End If
   End If 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update USUARIO Set " & VbNewLine 
   Sql = Sql & " IDUSU         = " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & " , IDGRUPO       = " & SqlStr(mvarIDGRUPO) & VbNewLine 
   Sql = Sql & " , SENHAUSU      = " & SqlStr(mvarSENHAUSU) & VbNewLine 
   Sql = Sql & " , NMUSU         = " & SqlStr(mvarNMUSU) & VbNewLine 
   Sql = Sql & " , DTCADASTRO    = " & SqlDate(mvarDTCADASTRO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTEXPIRAUSU   = " & SqlDate(mvarDTEXPIRAUSU, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTEXPIRASENHA = " & SqlDate(mvarDTEXPIRASENHA, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , TPUSU         = " & SqlStr(mvarTPUSU) & VbNewLine 
   Sql = Sql & " , SITUSU        = " & SqlStr(mvarSITUSU) & VbNewLine 
   Sql = Sql & " , USUCADASTRO   = " & SqlStr(mvarUSUCADASTRO) & VbNewLine 
   Sql = Sql & " , FEXPIRASENHA  = " & SqlNum(mvarFEXPIRASENHA) & VbNewLine 
   Sql = Sql & " , MATRUSU       = " & SqlStr(mvarMATRUSU) & VbNewLine 
   Sql = Sql & " , ICORREIO      = " & SqlStr(mvarICORREIO) & VbNewLine 
   Sql = Sql & " , ECORREIO      = " & SqlStr(mvarECORREIO) & VbNewLine 
   Sql = Sql & " , TIMEOUT       = " & SqlNum(mvarTIMEOUT) & VbNewLine 
   Sql = Sql & " , IDSETOR       = " & SqlNum(mvarIDSETOR) & VbNewLine 
   Sql = Sql & " , IDEMPR        = " & SqlNum(mvarIDEMPR) & VbNewLine 
   Sql = Sql & " , LOCALTRAB     = " & SqlStr(mvarLOCALTRAB) & VbNewLine 
   Sql = Sql & " , NIVELOS       = " & SqlNum(mvarNIVELOS) & VbNewLine 
   Sql = Sql & " , NMAD          = " & SqlStr(mvarNMAD) & VbNewLine 
   Sql = Sql & " , IDPESSOA      = " & SqlNum(mvarIDPESSOA) & VbNewLine 
   Sql = Sql & " , IDCOLIGADA    = " & SqlNum(mvarIDCOLIGADA) & VbNewLine 
   Sql = Sql & " , DTVALIDAUSU   = " & SqlDate(mvarDTVALIDAUSU, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , TROCARSENHA   = " & SqlNum(mvarTROCARSENHA) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'USUARIO'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDUSU           = " & SqlStr(mvarIDUSU) & VbNewLine 

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & VbNewLine 
   Sql = Sql & Me.QryUpDate(pSinc:=pSinc)
   Sql = Sql & " Else " & VbNewLine 
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId, pSinc:=pSinc)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect( Optional Ch_IDUSU, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDUSU, IDGRUPO, SENHAUSU, NMUSU" & VbNewLine 
   Sql = Sql & ", DTCADASTRO, DTEXPIRAUSU, DTEXPIRASENHA, TPUSU, SITUSU" & VbNewLine 
   Sql = Sql & ", USUCADASTRO, FEXPIRASENHA, MATRUSU, ICORREIO, ECORREIO, TIMEOUT, IDSETOR, IDEMPR, LOCALTRAB" & VbNewLine 
   Sql = Sql & ", NIVELOS, NMAD, IDPESSOA, IDCOLIGADA, DTVALIDAUSU" & VbNewLine 
   Sql = Sql & ", TROCARSENHA" & VbNewLine 
   Sql = Sql &" From USUARIO" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDUSU)  Then 
      Sql = Sql & " IDUSU = " & SqlStr(mvarIDUSU) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDUSU) Then Sql = Sql & " IDUSU = " & SqlStr(Cstr(Ch_IDUSU)) & " AND " & VbNewLine 
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & VbNewLine
      End If
   End If 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY 

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar( Optional Ch_IDUSU, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDUSU, Ch_WHERE, Ch_ORDERBY)
   bExiste = mvarXDb.AbreTabela(Sql, mvarRS)
   mvarisDirt = False
   With mvarRS
      If bExiste Then bExiste = Not .EOF
      If bExiste Then
         Me.Popula
         Pesquisar = True
      Else
         Pesquisar = False
      End If
   End With
   Exit Function
PesquisarErr:
    call ShowError(Sql)
    Pesquisar = False
End Function
Public Sub Popula(Optional pRcSet)
   If IsMissing(pRcSet) Then Set pRcSet = mvarRS
   With pRcSet
      mvarIDUSU = !IDUSU & ""
      mvarIDGRUPO = !IDGRUPO & ""
      mvarSENHAUSU = !SENHAUSU & ""
      mvarNMUSU = !NMUSU & ""
      mvarDTCADASTRO = xDate(!DTCADASTRO & "", True)
      mvarDTEXPIRAUSU = xDate(!DTEXPIRAUSU & "", True)
      mvarDTEXPIRASENHA = xDate(!DTEXPIRASENHA & "", True)
      mvarTPUSU = !TPUSU & ""
      mvarSITUSU = !SITUSU & ""
      mvarUSUCADASTRO = !USUCADASTRO & ""
      mvarFEXPIRASENHA = XVal(!FEXPIRASENHA & "")
      mvarMATRUSU = !MATRUSU & ""
      mvarICORREIO = !ICORREIO & ""
      mvarECORREIO = !ECORREIO & ""
      mvarECORREIO = !ECORREIO & ""
      mvarTIMEOUT = XVal(!TIMEOUT & "")
      mvarIDSETOR = XVal(!IDSETOR & "")
      mvarIDEMPR = XVal(!IDEMPR & "")
      mvarLOCALTRAB = !LOCALTRAB & ""
      mvarNIVELOS = XVal(!NIVELOS & "")
      mvarNMAD = !NMAD & ""
      mvarIDPESSOA = XVal(!IDPESSOA & "")
      mvarIDCOLIGADA = XVal(!IDCOLIGADA & "")
      mvarDTVALIDAUSU = xDate(!DTVALIDAUSU & "", True)
      mvarTROCARSENHA = XVal(!TROCARSENHA & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDUSU = ""
   mvarIDGRUPO = ""
   mvarSENHAUSU = ""
   mvarNMUSU = ""
   mvarDTCADASTRO = ""
   mvarDTEXPIRAUSU = ""
   mvarDTEXPIRASENHA = ""
   mvarTPUSU = ""
   mvarSITUSU = ""
   mvarUSUCADASTRO = ""
   mvarFEXPIRASENHA = 0
   mvarMATRUSU = ""
   mvarICORREIO = ""
   mvarECORREIO = ""
   mvarECORREIO = ""
   mvarTIMEOUT = 0
   mvarIDSETOR = 0
   mvarIDEMPR = 0
   mvarLOCALTRAB = ""
   mvarNIVELOS = 0
   mvarNMAD = ""
   mvarIDPESSOA = 0
   mvarIDCOLIGADA = 0
   mvarDTVALIDAUSU = ""
   mvarTROCARSENHA = 0

   On Error Resume Next
   Call Class_Initialize
End Sub
Public Function Salvar(Optional ComCOMMIT As Boolean = True, Optional pAutoId As Boolean = True, Optional pSinc As Boolean = False) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId, pSinc:=pSinc), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId as Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDUSU), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDUSU As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update USUARIO Set "
   Sql = Sql & " IDUSU = " & SqlStr(Ch_IDUSU)
   Sql = Sql & " Where 
   Sql = Sql & " IDUSU = " & SqlStr(mvarIDUSU)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Initialize()
   Set mvarRS = Nothing
   Set mvarXDb = Nothing
   mvarQryInsert = ""
   mvarQryUpDate = ""
   mvarQryDelete = ""
   mvarQrySelect = ""
   mvarQrySave = ""
   mvarisDirt = False
End Sub
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
