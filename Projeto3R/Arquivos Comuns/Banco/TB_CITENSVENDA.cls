VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_CITENSVENDA"
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

Private mvarIDLOJA       As Double
Private mvarIDVENDA      As Double
Private mvarIDITEM       As Double
Private mvarIDPROD       As Double
Private mvarNMPROD       As String
Private mvarUNIDCONTROLE As String
Private mvarQTDVENDA     As Double
Private mvarVLUNIT       As Double
Private mvarALTERSTAMP   As Double
Private mvarTIMESTAMP    As String
Private mvarIDPACOTE     As Double
Private mvarVLDESC       As Double
Private mvarDSCDESC      As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let IDITEM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDITEM <> vData)
   mvarIDITEM = vData
End Property
Public Property Get IDITEM() As Double
   IDITEM = mvarIDITEM
End Property
Public Property Let IDPROD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROD <> vData)
   mvarIDPROD = vData
End Property
Public Property Get IDPROD() As Double
   IDPROD = mvarIDPROD
End Property
Public Property Let NMPROD(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarNMPROD <> vData)
   mvarNMPROD = vData
End Property
Public Property Get NMPROD() As String
   NMPROD = mvarNMPROD
End Property
Public Property Let UNIDCONTROLE(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarUNIDCONTROLE <> vData)
   mvarUNIDCONTROLE = vData
End Property
Public Property Get UNIDCONTROLE() As String
   UNIDCONTROLE = mvarUNIDCONTROLE
End Property
Public Property Let QTDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDVENDA <> vData)
   mvarQTDVENDA = vData
End Property
Public Property Get QTDVENDA() As Double
   QTDVENDA = mvarQTDVENDA
End Property
Public Property Let VLUNIT(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLUNIT <> vData)
   mvarVLUNIT = vData
End Property
Public Property Get VLUNIT() As Double
   VLUNIT = mvarVLUNIT
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update CITENSVENDA Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDVENDA    = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " And IDITEM     = " & SqlNum(mvarIDITEM) & VbNewLine 
   If Not mvarXDb Is Nothing Then
      If mvarXDb.Conectado Then
         If mvarXDb.Executa(Sql, True) Then
            mvarALTERSTAMP = vData
         End If
      End If
   End If
End Property
Public Property Get ALTERSTAMP() As Double
   ALTERSTAMP = mvarALTERSTAMP
End Property
Public Property Let TIMESTAMP(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTIMESTAMP <> vData)
   mvarTIMESTAMP = vData
End Property
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
End Property
Public Property Let IDPACOTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPACOTE <> vData)
   mvarIDPACOTE = vData
End Property
Public Property Get IDPACOTE() As Double
   IDPACOTE = mvarIDPACOTE
End Property
Public Property Let VLDESC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLDESC <> vData)
   mvarVLDESC = vData
End Property
Public Property Get VLDESC() As Double
   VLDESC = mvarVLDESC
End Property
Public Property Let DSCDESC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCDESC <> vData)
   mvarDSCDESC = vData
End Property
Public Property Get DSCDESC() As String
   DSCDESC = mvarDSCDESC
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
 
   Sql = "Insert Into CITENSVENDA (IDLOJA, IDVENDA, IDITEM, IDPROD" & VbNewLine 
   Sql = Sql & ", NMPROD, UNIDCONTROLE, QTDVENDA, VLUNIT, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP, IDPACOTE, VLDESC, DSCDESC) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDITEM),0)+1 From CITENSVENDA Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDVENDA = " & SqlNum(mvarIDVENDA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDITEM) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarNMPROD) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarUNIDCONTROLE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarQTDVENDA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLUNIT) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", " &  SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", " &  SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", GetDate()" & vbNewLine
   End If
   Sql = Sql & ", " & SqlNum(mvarIDPACOTE) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarDSCDESC) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_IDITEM, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From CITENSVENDA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDVENDA)  And IsMissing(Ch_IDITEM)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDVENDA = " & SqlNum(mvarIDVENDA) & " AND " & VbNewLine 
      Sql = Sql & " IDITEM = " & SqlNum(mvarIDITEM) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(Cstr(Ch_IDVENDA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDITEM) Then Sql = Sql & " IDITEM = " & SqlNum(Cstr(Ch_IDITEM)) & " AND " & VbNewLine 
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

   Dim MyDelRow As New TB_DELETEDROWS
   Dim sTag     As String

   sTag = "
   sTag = sTag & "|IDLOJA = " & mvarIDLOJA
   sTag = sTag & "|IDVENDA = " & mvarIDVENDA
   sTag = sTag & "|IDITEM = " & mvarIDITEM
   sTag = sTag & "|IDPROD = " & mvarIDPROD
   sTag = sTag & "|NMPROD = " & mvarNMPROD
   sTag = sTag & "|UNIDCONTROLE = " & mvarUNIDCONTROLE
   sTag = sTag & "|QTDVENDA = " & mvarQTDVENDA
   sTag = sTag & "|VLUNIT = " & mvarVLUNIT
   sTag = sTag & "|IDPACOTE = " & mvarIDPACOTE
   sTag = sTag & "|VLDESC = " & mvarVLDESC
   sTag = sTag & "|DSCDESC = " & mvarDSCDESC
   sTag = sTag & "|Where = " & IIf(IsMissing(Ch_WHERE), "", Ch_WHERE)
   sTag = sTag & "|"

   MyDelRow.Query = Sql & ";"
   MyDelRow.Tag = sTag
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate(Optional pSinc As Boolean = False) As String
   Dim Sql As String
 
   Sql = "Update CITENSVENDA Set " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDVENDA      = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " , IDITEM       = " & SqlNum(mvarIDITEM) & VbNewLine 
   Sql = Sql & " , IDPROD       = " & SqlNum(mvarIDPROD) & VbNewLine 
   Sql = Sql & " , NMPROD       = " & SqlStr(mvarNMPROD) & VbNewLine 
   Sql = Sql & " , UNIDCONTROLE = " & SqlStr(mvarUNIDCONTROLE) & VbNewLine 
   Sql = Sql & " , QTDVENDA     = " & SqlNum(mvarQTDVENDA) & VbNewLine 
   Sql = Sql & " , VLUNIT       = " & SqlNum(mvarVLUNIT) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP   = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP   = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP    = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP    = GetDate()" & vbNewLine
   End If
   Sql = Sql & " , IDPACOTE     = " & SqlNum(mvarIDPACOTE) & VbNewLine 
   Sql = Sql & " , VLDESC       = " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & " , DSCDESC      = " & SqlStr(mvarDSCDESC) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'CITENSVENDA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDVENDA    = " & SqlNum(mvarIDVENDA) & VbNewLine 
   Sql = Sql & " And IDITEM     = " & SqlNum(mvarIDITEM) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_IDITEM, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDVENDA, IDITEM, IDPROD" & VbNewLine 
   Sql = Sql & ", NMPROD, UNIDCONTROLE, QTDVENDA, VLUNIT, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP, IDPACOTE, VLDESC, DSCDESC" & VbNewLine 
   Sql = Sql &" From CITENSVENDA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDVENDA)  And IsMissing(Ch_IDITEM)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDVENDA = " & SqlNum(mvarIDVENDA) & " AND " & VbNewLine 
      Sql = Sql & " IDITEM = " & SqlNum(mvarIDITEM) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDVENDA) Then Sql = Sql & " IDVENDA = " & SqlNum(Cstr(Ch_IDVENDA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDITEM) Then Sql = Sql & " IDITEM = " & SqlNum(Cstr(Ch_IDITEM)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDVENDA, Optional Ch_IDITEM, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDVENDA, Ch_IDITEM, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDLOJA = XVal(!IDLOJA & "")
      mvarIDVENDA = XVal(!IDVENDA & "")
      mvarIDITEM = XVal(!IDITEM & "")
      mvarIDPROD = XVal(!IDPROD & "")
      mvarNMPROD = !NMPROD & ""
      mvarUNIDCONTROLE = !UNIDCONTROLE & ""
      mvarQTDVENDA = XVal(!QTDVENDA & "")
      mvarVLUNIT = XVal(!VLUNIT & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarIDPACOTE = XVal(!IDPACOTE & "")
      mvarVLDESC = XVal(!VLDESC & "")
      mvarDSCDESC = !DSCDESC & ""
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDVENDA = 0
   mvarIDITEM = 0
   mvarIDPROD = 0
   mvarNMPROD = ""
   mvarUNIDCONTROLE = ""
   mvarQTDVENDA = 0
   mvarVLUNIT = 0
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""
   mvarIDPACOTE = 0
   mvarVLDESC = 0
   mvarDSCDESC = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDVENDA, mvarIDITEM), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDVENDA As String, Ch_IDITEM As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CITENSVENDA Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDVENDA = " & SqlNum(Ch_IDVENDA)
   Sql = Sql & " , IDITEM = " & SqlNum(Ch_IDITEM)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDVENDA = " & SqlNum(mvarIDVENDA)
   Sql = Sql & " and IDITEM = " & SqlNum(mvarIDITEM)
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
