VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_OTPTRATATR"
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

Private mvarIDLOJA         As Double
Private mvarIDTPTRATAMENTO As Double
Private mvarIDATR          As Double
Private mvarDSCATR         As String
Private mvarTPATR          As String
Private mvarVLLISTA        As String
Private mvarVLPADRAO       As String
Private mvarATIVO          As Double
Private mvarALTERSTAMP     As Double
Private mvarTIMESTAMP      As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDTPTRATAMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPTRATAMENTO <> vData)
   mvarIDTPTRATAMENTO = vData
End Property
Public Property Get IDTPTRATAMENTO() As Double
   IDTPTRATAMENTO = mvarIDTPTRATAMENTO
End Property
Public Property Let IDATR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDATR <> vData)
   mvarIDATR = vData
End Property
Public Property Get IDATR() As Double
   IDATR = mvarIDATR
End Property
Public Property Let DSCATR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDSCATR <> vData)
   mvarDSCATR = vData
End Property
Public Property Get DSCATR() As String
   DSCATR = mvarDSCATR
End Property
Public Property Let TPATR(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarTPATR <> vData)
   mvarTPATR = vData
End Property
Public Property Get TPATR() As String
   TPATR = mvarTPATR
End Property
Public Property Let VLLISTA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVLLISTA <> vData)
   mvarVLLISTA = vData
End Property
Public Property Get VLLISTA() As String
   VLLISTA = mvarVLLISTA
End Property
Public Property Let VLPADRAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarVLPADRAO <> vData)
   mvarVLPADRAO = vData
End Property
Public Property Get VLPADRAO() As String
   VLPADRAO = mvarVLPADRAO
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
 
   Sql = "Update OTPTRATATR Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & VbNewLine 
   Sql = Sql & " And IDATR        = " & SqlNum(mvarIDATR) & VbNewLine 
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
 
   Sql = "Insert Into OTPTRATATR (IDLOJA, IDTPTRATAMENTO, IDATR, DSCATR" & VbNewLine 
   Sql = Sql & ", TPATR, VLLISTA, VLPADRAO, ATIVO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDTPTRATAMENTO) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDATR),0)+1 From OTPTRATATR Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDATR) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlStr(mvarDSCATR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarTPATR) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVLLISTA) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarVLPADRAO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarATIVO) & VbNewLine 
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

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDTPTRATAMENTO, Optional Ch_IDATR, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From OTPTRATATR" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDTPTRATAMENTO)  And IsMissing(Ch_IDATR)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & " AND " & VbNewLine 
      Sql = Sql & " IDATR = " & SqlNum(mvarIDATR) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDTPTRATAMENTO) Then Sql = Sql & " IDTPTRATAMENTO = " & SqlNum(Cstr(Ch_IDTPTRATAMENTO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDATR) Then Sql = Sql & " IDATR = " & SqlNum(Cstr(Ch_IDATR)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDTPTRATAMENTO = " & mvarIDTPTRATAMENTO
   sTag = sTag & "|IDATR = " & mvarIDATR
   sTag = sTag & "|DSCATR = " & mvarDSCATR
   sTag = sTag & "|TPATR = " & mvarTPATR
   sTag = sTag & "|VLLISTA = " & mvarVLLISTA
   sTag = sTag & "|VLPADRAO = " & mvarVLPADRAO
   sTag = sTag & "|ATIVO = " & mvarATIVO
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
 
   Sql = "Update OTPTRATATR Set " & VbNewLine 
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & VbNewLine 
   Sql = Sql & " , IDATR          = " & SqlNum(mvarIDATR) & VbNewLine 
   Sql = Sql & " , DSCATR         = " & SqlStr(mvarDSCATR) & VbNewLine 
   Sql = Sql & " , TPATR          = " & SqlStr(mvarTPATR) & VbNewLine 
   Sql = Sql & " , VLLISTA        = " & SqlStr(mvarVLLISTA) & VbNewLine 
   Sql = Sql & " , VLPADRAO       = " & SqlStr(mvarVLPADRAO) & VbNewLine 
   Sql = Sql & " , ATIVO          = " & SqlNum(mvarATIVO) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP     = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP     = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP      = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP      = GetDate()" & vbNewLine
   End If
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OTPTRATATR'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & VbNewLine 
   Sql = Sql & " And IDATR        = " & SqlNum(mvarIDATR) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDTPTRATAMENTO, Optional Ch_IDATR, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDTPTRATAMENTO, IDATR, DSCATR" & VbNewLine 
   Sql = Sql & ", TPATR, VLLISTA, VLPADRAO, ATIVO, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP" & VbNewLine 
   Sql = Sql &" From OTPTRATATR" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDTPTRATAMENTO)  And IsMissing(Ch_IDATR)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO) & " AND " & VbNewLine 
      Sql = Sql & " IDATR = " & SqlNum(mvarIDATR) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDTPTRATAMENTO) Then Sql = Sql & " IDTPTRATAMENTO = " & SqlNum(Cstr(Ch_IDTPTRATAMENTO)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDATR) Then Sql = Sql & " IDATR = " & SqlNum(Cstr(Ch_IDATR)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDTPTRATAMENTO, Optional Ch_IDATR, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDTPTRATAMENTO, Ch_IDATR, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDTPTRATAMENTO = XVal(!IDTPTRATAMENTO & "")
      mvarIDATR = XVal(!IDATR & "")
      mvarDSCATR = !DSCATR & ""
      mvarTPATR = !TPATR & ""
      mvarVLLISTA = !VLLISTA & ""
      mvarVLPADRAO = !VLPADRAO & ""
      mvarATIVO = XVal(!ATIVO & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDTPTRATAMENTO = 0
   mvarIDATR = 0
   mvarDSCATR = ""
   mvarTPATR = ""
   mvarVLLISTA = ""
   mvarVLPADRAO = ""
   mvarATIVO = 0
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDTPTRATAMENTO, mvarIDATR), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDTPTRATAMENTO As String, Ch_IDATR As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OTPTRATATR Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDTPTRATAMENTO = " & SqlNum(Ch_IDTPTRATAMENTO)
   Sql = Sql & " , IDATR = " & SqlNum(Ch_IDATR)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDTPTRATAMENTO = " & SqlNum(mvarIDTPTRATAMENTO)
   Sql = Sql & " and IDATR = " & SqlNum(mvarIDATR)
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
