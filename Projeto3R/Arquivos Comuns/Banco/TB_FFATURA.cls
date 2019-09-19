VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
END
Attribute VB_Name = "TB_FFATURA"
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

Private mvarIDLOJA        As Double
Private mvarIDFATURA      As Double
Private mvarIDCLIENTE     As Double
Private mvarDTEMISSAO     As String
Private mvarDTPREV        As String
Private mvarVALOR         As Double
Private mvarVLDESC        As Double
Private mvarIDSUBDESP     As Double
Private mvarHISTORICO     As String
Private mvarIDATENDIMENTO As Double
Private mvarIDVENDA       As Double
Private mvarSITFATURA     As Double
Private mvarIDDESP        As Double
Private mvarALTERSTAMP    As Double
Private mvarTIMESTAMP     As String
Private mvarIDVENDAORIGEM As Double
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDFATURA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDFATURA <> vData)
   mvarIDFATURA = vData
End Property
Public Property Get IDFATURA() As Double
   IDFATURA = mvarIDFATURA
End Property
Public Property Let IDCLIENTE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDCLIENTE <> vData)
   mvarIDCLIENTE = vData
End Property
Public Property Get IDCLIENTE() As Double
   IDCLIENTE = mvarIDCLIENTE
End Property
Public Property Let DTEMISSAO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTEMISSAO <> vData)
   mvarDTEMISSAO = vData
End Property
Public Property Get DTEMISSAO() As String
   DTEMISSAO = mvarDTEMISSAO
End Property
Public Property Let DTPREV(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTPREV <> vData)
   mvarDTPREV = vData
End Property
Public Property Get DTPREV() As String
   DTPREV = mvarDTPREV
End Property
Public Property Let VALOR(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVALOR <> vData)
   mvarVALOR = vData
End Property
Public Property Get VALOR() As Double
   VALOR = mvarVALOR
End Property
Public Property Let VLDESC(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarVLDESC <> vData)
   mvarVLDESC = vData
End Property
Public Property Get VLDESC() As Double
   VLDESC = mvarVLDESC
End Property
Public Property Let IDSUBDESP(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSUBDESP <> vData)
   mvarIDSUBDESP = vData
End Property
Public Property Get IDSUBDESP() As Double
   IDSUBDESP = mvarIDSUBDESP
End Property
Public Property Let HISTORICO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarHISTORICO <> vData)
   mvarHISTORICO = vData
End Property
Public Property Get HISTORICO() As String
   HISTORICO = mvarHISTORICO
End Property
Public Property Let IDATENDIMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDATENDIMENTO <> vData)
   mvarIDATENDIMENTO = vData
End Property
Public Property Get IDATENDIMENTO() As Double
   IDATENDIMENTO = mvarIDATENDIMENTO
End Property
Public Property Let IDVENDA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDA <> vData)
   mvarIDVENDA = vData
End Property
Public Property Get IDVENDA() As Double
   IDVENDA = mvarIDVENDA
End Property
Public Property Let SITFATURA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarSITFATURA <> vData)
   mvarSITFATURA = vData
End Property
Public Property Get SITFATURA() As Double
   SITFATURA = mvarSITFATURA
End Property
Public Property Let IDDESP(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDDESP <> vData)
   mvarIDDESP = vData
End Property
Public Property Get IDDESP() As Double
   IDDESP = mvarIDDESP
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update FFATURA Set ALTERSTAMP=" & vData & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDFATURA    = " & SqlNum(mvarIDFATURA) & VbNewLine 
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
Public Property Let IDVENDAORIGEM(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDVENDAORIGEM <> vData)
   mvarIDVENDAORIGEM = vData
End Property
Public Property Get IDVENDAORIGEM() As Double
   IDVENDAORIGEM = mvarIDVENDAORIGEM
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
 
   Sql = "Insert Into FFATURA (IDLOJA, IDFATURA, IDCLIENTE, DTEMISSAO" & VbNewLine 
   Sql = Sql & ", DTPREV, VALOR, VLDESC, IDSUBDESP, HISTORICO" & VbNewLine 
   Sql = Sql & ", IDATENDIMENTO, IDVENDA, SITFATURA, IDDESP, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP, IDVENDAORIGEM) " & VbNewLine 
   Sql = Sql & " Output Inserted.*" & VbNewLine 
   Sql = Sql & " Select 
   Sql = Sql &  SqlNum(mvarIDLOJA) & VbNewLine 
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDFATURA),0)+1 From FFATURA Where IDLOJA = " & SqlNum(mvarIDLOJA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDFATURA) & VbNewLine 
   End If
   Sql = Sql & ", " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlDate(mvarDTPREV, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVALOR) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDSUBDESP) & VbNewLine 
   Sql = Sql & ", " & SqlStr(mvarHISTORICO) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDATENDIMENTO) & VbNewLine 
   Sql = Sql & ", " & IIf(mvarIDVENDA = 0, "Null", SqlNum(mvarIDVENDA)) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarSITFATURA) & VbNewLine 
   Sql = Sql & ", " & SqlNum(mvarIDDESP) & VbNewLine 
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
   Sql = Sql & ", " & SqlNum(mvarIDVENDAORIGEM) & VbNewLine 

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete( Optional Ch_IDLOJA, Optional Ch_IDFATURA, Optional Ch_WHERE) As String
   Dim Sql As String
 
   Sql = "Delete From FFATURA" & VbNewLine 
   Sql = Sql & " Where" & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDFATURA)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDFATURA = " & SqlNum(mvarIDFATURA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDFATURA) Then Sql = Sql & " IDFATURA = " & SqlNum(Cstr(Ch_IDFATURA)) & " AND " & VbNewLine 
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
   sTag = sTag & "|IDFATURA = " & mvarIDFATURA
   sTag = sTag & "|IDCLIENTE = " & mvarIDCLIENTE
   sTag = sTag & "|DTEMISSAO = " & mvarDTEMISSAO
   sTag = sTag & "|DTPREV = " & mvarDTPREV
   sTag = sTag & "|VALOR = " & mvarVALOR
   sTag = sTag & "|VLDESC = " & mvarVLDESC
   sTag = sTag & "|IDSUBDESP = " & mvarIDSUBDESP
   sTag = sTag & "|HISTORICO = " & mvarHISTORICO
   sTag = sTag & "|IDATENDIMENTO = " & mvarIDATENDIMENTO
   sTag = sTag & "|IDVENDA = " & mvarIDVENDA
   sTag = sTag & "|SITFATURA = " & mvarSITFATURA
   sTag = sTag & "|IDDESP = " & mvarIDDESP
   sTag = sTag & "|IDVENDAORIGEM = " & mvarIDVENDAORIGEM
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
 
   Sql = "Update FFATURA Set " & VbNewLine 
   Sql = Sql & " IDLOJA        = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " , IDFATURA      = " & SqlNum(mvarIDFATURA) & VbNewLine 
   Sql = Sql & " , IDCLIENTE     = " & SqlNum(mvarIDCLIENTE) & VbNewLine 
   Sql = Sql & " , DTEMISSAO     = " & SqlDate(mvarDTEMISSAO, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , DTPREV        = " & SqlDate(mvarDTPREV, eSysDate.Data_Hora) & VbNewLine 
   Sql = Sql & " , VALOR         = " & SqlNum(mvarVALOR) & VbNewLine 
   Sql = Sql & " , VLDESC        = " & SqlNum(mvarVLDESC) & VbNewLine 
   Sql = Sql & " , IDSUBDESP     = " & SqlNum(mvarIDSUBDESP) & VbNewLine 
   Sql = Sql & " , HISTORICO     = " & SqlStr(mvarHISTORICO) & VbNewLine 
   Sql = Sql & " , IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & VbNewLine 
   Sql = Sql & " , IDVENDA       = " & IIf(mvarIDVENDA = 0, "Null", SqlNum(mvarIDVENDA)) & VbNewLine 
   Sql = Sql & " , SITFATURA     = " & SqlNum(mvarSITFATURA) & VbNewLine 
   Sql = Sql & " , IDDESP        = " & SqlNum(mvarIDDESP) & VbNewLine 
   If pSinc Then 
      Sql = Sql & ", ALTERSTAMP    = " & SqlNum(mvarALTERSTAMP) & vbNewLine
   Else
      Sql = Sql & ", ALTERSTAMP    = 1" & vbNewLine
   End If
   If pSinc Then 
      Sql = Sql & ", TIMESTAMP     = " & SqlDate(mvarTIMESTAMP) & vbNewLine
   Else
      Sql = Sql & ", TIMESTAMP     = GetDate()" & vbNewLine
   End If
   Sql = Sql & " , IDVENDAORIGEM = " & SqlNum(mvarIDVENDAORIGEM) & VbNewLine 
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FFATURA'" ) Then
      Sql = Sql & " Output Inserted.*" & VbNewLine 
   End If
   Sql = Sql & " Where " & VbNewLine 
   Sql = Sql & " IDLOJA          = " & SqlNum(mvarIDLOJA) & VbNewLine 
   Sql = Sql & " And IDFATURA    = " & SqlNum(mvarIDFATURA) & VbNewLine 

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
Public Property Get QrySelect( Optional Ch_IDLOJA, Optional Ch_IDFATURA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDFATURA, IDCLIENTE, DTEMISSAO" & VbNewLine 
   Sql = Sql & ", DTPREV, VALOR, VLDESC, IDSUBDESP, HISTORICO" & VbNewLine 
   Sql = Sql & ", IDATENDIMENTO, IDVENDA, SITFATURA, IDDESP, ALTERSTAMP" & VbNewLine 
   Sql = Sql & ", TIMESTAMP, IDVENDAORIGEM" & VbNewLine 
   Sql = Sql &" From FFATURA" & VbNewLine 
   Sql = Sql & " Where " & VbNewLine 
   If IsMissing(Ch_WHERE)  And IsMissing(Ch_IDLOJA)  And IsMissing(Ch_IDFATURA)  Then 
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & VbNewLine 
      Sql = Sql & " IDFATURA = " & SqlNum(mvarIDFATURA) & " AND " & VbNewLine 
   Else 
      If Not isMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(Cstr(Ch_IDLOJA)) & " AND " & VbNewLine 
      If Not isMissing(Ch_IDFATURA) Then Sql = Sql & " IDFATURA = " & SqlNum(Cstr(Ch_IDFATURA)) & " AND " & VbNewLine 
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
Public Function Pesquisar( Optional Ch_IDLOJA, Optional Ch_IDFATURA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect( Ch_IDLOJA, Ch_IDFATURA, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDFATURA = XVal(!IDFATURA & "")
      mvarIDCLIENTE = XVal(!IDCLIENTE & "")
      mvarDTEMISSAO = xDate(!DTEMISSAO & "", True)
      mvarDTPREV = xDate(!DTPREV & "", True)
      mvarVALOR = XVal(!VALOR & "")
      mvarVLDESC = XVal(!VLDESC & "")
      mvarIDSUBDESP = XVal(!IDSUBDESP & "")
      mvarHISTORICO = !HISTORICO & ""
      mvarIDATENDIMENTO = XVal(!IDATENDIMENTO & "")
      mvarIDVENDA = XVal(!IDVENDA & "")
      mvarSITFATURA = XVal(!SITFATURA & "")
      mvarIDDESP = XVal(!IDDESP & "")
      mvarALTERSTAMP = XVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
      mvarIDVENDAORIGEM = XVal(!IDVENDAORIGEM & "")
   End With
   mvarisDirt = False
End Sub
Public Sub Limpar()
   mvarIDLOJA = 0
   mvarIDFATURA = 0
   mvarIDCLIENTE = 0
   mvarDTEMISSAO = ""
   mvarDTPREV = ""
   mvarVALOR = 0
   mvarVLDESC = 0
   mvarIDSUBDESP = 0
   mvarHISTORICO = ""
   mvarIDATENDIMENTO = 0
   mvarIDVENDA = 0
   mvarSITFATURA = 0
   mvarIDDESP = 0
   mvarALTERSTAMP = 0
   mvarTIMESTAMP = ""
   mvarIDVENDAORIGEM = 0

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
   Excluir = mvarXDb.Executa(Me.QryDelete( mvarIDLOJA, mvarIDFATURA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar =  mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave( Ch_IDLOJA As String, Ch_IDFATURA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FFATURA Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDFATURA = " & SqlNum(Ch_IDFATURA)
   Sql = Sql & " Where 
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDFATURA = " & SqlNum(mvarIDFATURA)
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
