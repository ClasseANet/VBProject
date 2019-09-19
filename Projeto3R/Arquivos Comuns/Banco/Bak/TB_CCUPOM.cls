VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_CCUPOM"
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

Private mvarIDLOJA     As Double
Private mvarIDPROMO    As Double
Private mvarIDCUPOM    As Double
Private mvarCODCUPOM   As String
Private mvarDTEMISSAO  As String
Private mvarDTVENC     As String
Private mvarVLDESC     As Double
Private mvarATIVO      As Double
Private mvarALTERSTAMP As Double
Private mvarTIMESTAMP  As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDPROMO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDPROMO <> vData)
   mvarIDPROMO = vData
End Property
Public Property Get IDPROMO() As Double
   IDPROMO = mvarIDPROMO
End Property
Public Property Let IDCUPOM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCUPOM <> vData)
   mvarIDCUPOM = vData
End Property
Public Property Get IDCUPOM() As Double
   IDCUPOM = mvarIDCUPOM
End Property
Public Property Let CODCUPOM(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarCODCUPOM <> vData)
   mvarCODCUPOM = vData
End Property
Public Property Get CODCUPOM() As String
   CODCUPOM = mvarCODCUPOM
End Property
Public Property Let DTEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEMISSAO <> vData)
   mvarDTEMISSAO = vData
End Property
Public Property Get DTEMISSAO() As String
   DTEMISSAO = mvarDTEMISSAO
End Property
Public Property Let DTVENC(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTVENC <> vData)
   mvarDTVENC = vData
End Property
Public Property Get DTVENC() As String
   DTVENC = mvarDTVENC
End Property
Public Property Let VLDESC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLDESC <> vData)
   mvarVLDESC = vData
End Property
Public Property Get VLDESC() As Double
   VLDESC = mvarVLDESC
End Property
Public Property Let ATIVO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarATIVO <> vData)
   mvarATIVO = vData
End Property
Public Property Get ATIVO() As Double
   ATIVO = mvarATIVO
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update CCUPOM Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDPROMO  = " & SqlNum(mvarIDPROMO) & VbNewLine 
   Sql = Sql & " And IDCUPOM  = " & SqlNum(mvarIDCUPOM) & VbNewLine 
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
Public Property Get TIMESTAMP() As String
   TIMESTAMP = mvarTIMESTAMP
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
Public Property Get QryInsert(Optional pAutoId as Boolean = True) As String
   Dim Sql As String
 
   Sql = "Insert Into CCUPOM (IDLOJA, IDPROMO, IDCUPOM, CODCUPOM" & VbNewLine 
   Sql = Sql & ", DTEMISSAO, DTVENC, VLDESC, ATIVO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDPROMO) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDCUPOM),0)+1 From CCUPOM Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDPROMO = " & SqlNum(mvarIDPROMO) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDCUPOM) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarCODCUPOM) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTVENC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarATIVO) & VbNewLine 
   Sql = Sql & ", " & "1" & VbNewLine 
   Sql = Sql & ", " & "GetDate()" & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_IDCUPOM) As String
   Dim Sql As String
 
   Sql = "Delete From CCUPOM" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   if Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDPROMO) Then Sql = Sql & " IDPROMO = " & SqlNum(Cstr(Ch_IDPROMO)) & " AND " & VbNewLine 
   if Not isMissing(Ch_IDCUPOM) Then Sql = Sql & " IDCUPOM = " & SqlNum(Cstr(Ch_IDCUPOM)) & " AND " & VbNewLine 
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   MyDelRow.Query = Sql & ";"
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update CCUPOM Set " & VbNewLine 
   Sql = Sql & " IDLOJA     = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDPROMO    = " & SqlNum(mvarIDPROMO) & VbNewLine 
   Sql = Sql & " , IDCUPOM    = " & SqlNum(mvarIDCUPOM) & VbNewLine 
   Sql = Sql & " , CODCUPOM   = " & SqlStr(mvarCODCUPOM) & VbNewLine 
   Sql = Sql & " , DTEMISSAO  = " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTVENC     = " & SqlDate(mvarDTVENC, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VLDESC     = " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & " , ATIVO      = " & SqlNum(mvarATIVO) & VbNewLine 
   Sql = Sql & " , ALTERSTAMP = " & "1" & VbNewLine 
   Sql = Sql & " , TIMESTAMP  = " & "GetDate()" & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'CCUPOM'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA       = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDPROMO  = " & SqlNum(mvarIDPROMO) & VbNewLine 
   Sql = Sql & " And IDCUPOM  = " & SqlNum(mvarIDCUPOM) & VbNewLine 

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId as Boolean = True) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & VbNewLine 
   Sql = Sql & Me.QryUpDate
   Sql = Sql & " Else " & VbNewLine 
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_IDCUPOM, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDPROMO, IDCUPOM, CODCUPOM" & VbNewLine 
   Sql = Sql & ", DTEMISSAO, DTVENC, VLDESC, ATIVO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP" & VbNewLine 
   Sql = Sql &" From CCUPOM" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDPROMO)  And IsMissing(Ch_IDCUPOM)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDPROMO = " & SqlNum(mvarIDPROMO) & " AND " & VbNewLine 
      Sql = Sql & " IDCUPOM = " & SqlNum(mvarIDCUPOM) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDPROMO) Then Sql = Sql & " IDPROMO = " & SqlNum(Cstr(Ch_IDPROMO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDCUPOM) Then Sql = Sql & " IDCUPOM = " & SqlNum(Cstr(Ch_IDCUPOM)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDPROMO, Optional Ch_IDCUPOM, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDPROMO, Ch_IDCUPOM, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDPROMO = XVal(!IDPROMO & "")
      mvarIDCUPOM = XVal(!IDCUPOM & "")
      mvarCODCUPOM = !CODCUPOM & ""
      mvarDTEMISSAO = xDate(!DTEMISSAO & "", True)
      mvarDTVENC = xDate(!DTVENC & "", True)
      mvarVLDESC = XVal(!VLDESC & "")
      mvarATIVO = XVal(!ATIVO & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Function Salvar(Optional ComCOMMIT = True, Optional pAutoId as Boolean = True) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId as Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDPROMO, mvarIDCUPOM), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDPROMO As String, Ch_IDCUPOM As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update CCUPOM Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDPROMO = " & SqlNum(Ch_IDPROMO)
   Sql = Sql & " , IDCUPOM = " & SqlNum(Ch_IDCUPOM)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDPROMO = " & SqlNum(mvarIDPROMO)
   Sql = Sql & " and IDCUPOM = " & SqlNum(mvarIDCUPOM)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
