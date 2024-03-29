VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_LOGEVT"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit 
Private mvarXDb         As Object 
Private mvarRS          As Object 

Private mvarQryInsert   As String
Private mvarQryUpDate   As String
Private mvarQryDelete   As String
Private mvarQrySelect   As String
Private mvarQrySave     As String

Private mvarisDirt      As Boolean

Private mvarIDLOG    As Object
Private mvarCODEVT   As String
Private mvarDSCEVT   As String
Private mvarTPDOC    As String
Private mvarSETOR    As String
Private mvarIDDOC    As String
Private mvarANODOC   As String
Private mvarCAMPO    As String
Private mvarITEMDOC  As Double
Private mvarVLANTES  As String
Private mvarVLDEPOIS As String
Private mvarIDUSU    As String
Private mvarMAQ      As String
Private mvarDTLOG    As String
Private mvarHORALOG  As String
Private mvarMOTIVO   As String
Public Property Set IDLOG(ByVal vData As Object)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOG <> vData)
   Set mvarIDLOG = vData
End Property
Public Property Get IDLOG() As Object
   Set IDLOG = mvarIDLOG
End Property
Public Property Let CODEVT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODEVT <> vData)
   mvarCODEVT = vData
End Property
Public Property Get CODEVT() As String
   CODEVT = mvarCODEVT
End Property
Public Property Let DSCEVT(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCEVT <> vData)
   mvarDSCEVT = vData
End Property
Public Property Get DSCEVT() As String
   DSCEVT = mvarDSCEVT
End Property
Public Property Let TPDOC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPDOC <> vData)
   mvarTPDOC = vData
End Property
Public Property Get TPDOC() As String
   TPDOC = mvarTPDOC
End Property
Public Property Let SETOR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarSETOR <> vData)
   mvarSETOR = vData
End Property
Public Property Get SETOR() As String
   SETOR = mvarSETOR
End Property
Public Property Let IDDOC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDDOC <> vData)
   mvarIDDOC = vData
End Property
Public Property Get IDDOC() As String
   IDDOC = mvarIDDOC
End Property
Public Property Let ANODOC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarANODOC <> vData)
   mvarANODOC = vData
End Property
Public Property Get ANODOC() As String
   ANODOC = mvarANODOC
End Property
Public Property Let CAMPO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCAMPO <> vData)
   mvarCAMPO = vData
End Property
Public Property Get CAMPO() As String
   CAMPO = mvarCAMPO
End Property
Public Property Let ITEMDOC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarITEMDOC <> vData)
   mvarITEMDOC = vData
End Property
Public Property Get ITEMDOC() As Double
   ITEMDOC = mvarITEMDOC
End Property
Public Property Let VLANTES(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVLANTES <> vData)
   mvarVLANTES = vData
End Property
Public Property Get VLANTES() As String
   VLANTES = mvarVLANTES
End Property
Public Property Let VLDEPOIS(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVLDEPOIS <> vData)
   mvarVLDEPOIS = vData
End Property
Public Property Get VLDEPOIS() As String
   VLDEPOIS = mvarVLDEPOIS
End Property
Public Property Let IDUSU(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDUSU <> vData)
   mvarIDUSU = vData
End Property
Public Property Get IDUSU() As String
   IDUSU = mvarIDUSU
End Property
Public Property Let MAQ(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMAQ <> vData)
   mvarMAQ = vData
End Property
Public Property Get MAQ() As String
   MAQ = mvarMAQ
End Property
Public Property Let DTLOG(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTLOG <> vData)
   mvarDTLOG = vData
End Property
Public Property Get DTLOG() As String
   DTLOG = mvarDTLOG
End Property
Public Property Let HORALOG(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHORALOG <> vData)
   mvarHORALOG = vData
End Property
Public Property Get HORALOG() As String
   HORALOG = mvarHORALOG
End Property
Public Property Let MOTIVO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarMOTIVO <> vData)
   mvarMOTIVO = vData
End Property
Public Property Get MOTIVO() As String
   MOTIVO = mvarMOTIVO
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
 
   Sql = "Insert Into LOGEVT (CODEVT, DSCEVT, TPDOC" & VbNewLine 
   Sql = Sql & ", SETOR, IDDOC, ANODOC, CAMPO, ITEMDOC" & VbNewLine 
   Sql = Sql & ", VLANTES, VLDEPOIS, IDUSU, MAQ, DTLOG" & VbNewLine 
   Sql = Sql & ", HORALOG, MOTIVO) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &   & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCODEVT) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCEVT) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTPDOC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarSETOR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDDOC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarANODOC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarCAMPO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarITEMDOC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVLANTES) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVLDEPOIS) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMAQ) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTLOG, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarHORALOG) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarMOTIVO) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOG, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From LOGEVT" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOG)  Then 
      Sql = Sql & " IDLOG = " & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOG) Then Sql = Sql & " IDLOG = " & " AND " & VbNewLine 
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
 
   Sql = "Update LOGEVT Set " & VbNewLine 
   Sql = Sql & " , CODEVT   = " & SqlStr(mvarCODEVT) & VbNewLine 
   Sql = Sql & " , DSCEVT   = " & SqlStr(mvarDSCEVT) & VbNewLine 
   Sql = Sql & " , TPDOC    = " & SqlStr(mvarTPDOC) & VbNewLine 
   Sql = Sql & " , SETOR    = " & SqlStr(mvarSETOR) & VbNewLine 
   Sql = Sql & " , IDDOC    = " & SqlStr(mvarIDDOC) & VbNewLine 
   Sql = Sql & " , ANODOC   = " & SqlStr(mvarANODOC) & VbNewLine 
   Sql = Sql & " , CAMPO    = " & SqlStr(mvarCAMPO) & VbNewLine 
   Sql = Sql & " , ITEMDOC  = " & SqlNum(mvarITEMDOC) & VbNewLine 
   Sql = Sql & " , VLANTES  = " & SqlStr(mvarVLANTES) & VbNewLine 
   Sql = Sql & " , VLDEPOIS = " & SqlStr(mvarVLDEPOIS) & VbNewLine 
   Sql = Sql & " , IDUSU    = " & SqlStr(mvarIDUSU) & VbNewLine 
   Sql = Sql & " , MAQ      = " & SqlStr(mvarMAQ) & VbNewLine 
   Sql = Sql & " , DTLOG    = " & SqlDate(mvarDTLOG, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , HORALOG  = " & SqlStr(mvarHORALOG) & VbNewLine 
   Sql = Sql & " , MOTIVO   = " & SqlStr(mvarMOTIVO) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'LOGEVT'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOG      = " & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOG, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select , CODEVT, DSCEVT, TPDOC" & VbNewLine 
   Sql = Sql & ", SETOR, IDDOC, ANODOC, CAMPO, ITEMDOC" & VbNewLine 
   Sql = Sql & ", VLANTES, VLDEPOIS, IDUSU, MAQ, DTLOG" & VbNewLine 
   Sql = Sql & ", HORALOG, MOTIVO" & VbNewLine 
   Sql = Sql &" From LOGEVT" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOG)  Then 
      Sql = Sql & " IDLOG = " & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOG) Then Sql = Sql & " IDLOG = " & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOG, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOG, Ch_WHERE, Ch_ORDERBY)
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

      mvarCODEVT = !CODEVT & ""
      mvarDSCEVT = !DSCEVT & ""
      mvarTPDOC = !TPDOC & ""
      mvarSETOR = !SETOR & ""
      mvarIDDOC = !IDDOC & ""
      mvarANODOC = !ANODOC & ""
      mvarCAMPO = !CAMPO & ""
      mvarITEMDOC = XVal(!ITEMDOC & "")
      mvarVLANTES = !VLANTES & ""
      mvarVLDEPOIS = !VLDEPOIS & ""
      mvarIDUSU = !IDUSU & ""
      mvarMAQ = !MAQ & ""
      mvarDTLOG = xDate(!DTLOG & "", True)
      mvarHORALOG = !HORALOG & ""
      mvarMOTIVO = !MOTIVO & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()

   mvarCODEVT = ""
   mvarDSCEVT = ""
   mvarTPDOC = ""
   mvarSETOR = ""
   mvarIDDOC = ""
   mvarANODOC = ""
   mvarCAMPO = ""
   mvarITEMDOC = 0
   mvarVLANTES = ""
   mvarVLDEPOIS = ""
   mvarIDUSU = ""
   mvarMAQ = ""
   mvarDTLOG = ""
   mvarHORALOG = ""
   mvarMOTIVO = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOG), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOG As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update LOGEVT Set "
   Sql = Sql & " IDLOG = "
   Sql = Sql & " Where 
   Sql = Sql & " IDLOG = "
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
