VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_FFATURA"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private mvarXDb            As DS_BANCO
Private mvarRS             As Object

Private mvarQryInsert      As String
Private mvarQryUpDate      As String
Private mvarQryDelete      As String
Private mvarQrySelect      As String
Private mvarQrySave        As String

Private mvarisDirt         As Boolean

Private mvarIDLOJA        As String
Private mvarIDFATURA       As Double
Private mvarIDCLIENTE     As Double
Private mvarDTPREV        As String
Private mvarVALOR         As Double
Private mvarIDSUBDESP     As Double
Private mvarHISTORICO     As String
Private mvarIDATENDIMENTO As Double
Private mvarIDVENDA       As Double
Private mvarSITFATURA     As Double
Private mvarIDDESP        As Double
Private mvarALTERSTAMP    As Double
Private mvarTIMESTAMP     As String
Public Property Let IDLOJA(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As String
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
Public Property Get QryInsert() As String
   Dim Sql As String
 
   Sql = "Insert Into FFATURA (IDLOJA, IDFATURA, IDCLIENTE, DTPREV" & vbNewLine
   Sql = Sql & ", VALOR, IDSUBDESP, HISTORICO, IDATENDIMENTO, IDVENDA" & vbNewLine
   Sql = Sql & ", SITFATURA, IDDESP, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Values " & vbNewLine
   Sql = Sql & "(" & vbNewLine
   Sql = Sql & SqlStr(mvarIDLOJA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDFATURA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDCLIENTE) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTPREV, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarVALOR) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDSUBDESP) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarHISTORICO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDATENDIMENTO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDVENDA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarSITFATURA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDDESP) & vbNewLine
   Sql = Sql & ", " & SqlNum(1) & vbNewLine
   Sql = Sql & ", " & SqlStr(mvarXDb.Sysdate(eSysDate.Data_Hora)) & vbNewLine
   Sql = Sql & ")" & vbNewLine

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_IDFATURA) As String
   Dim Sql As String
 
   Sql = "Delete From FFATURA" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlStr(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
   If Not IsMissing(Ch_IDFATURA) Then Sql = Sql & " IDFATURA = " & SqlNum(CStr(Ch_IDFATURA)) & " AND " & vbNewLine
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   mvarQryDelete = Sql

   Dim MyDelRow As New TB_DELETEDROWS
   MyDelRow.QUERY = Sql & ";"
   mvarQryDelete = mvarQryDelete & vbNewLine & MyDelRow.QryInsert & ";"
   Set MyDelRow = Nothing

   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql As String
 
   Sql = "Update FFATURA Set " & vbNewLine
   Sql = Sql & " IDLOJA        = " & SqlStr(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDFATURA       = " & SqlNum(mvarIDFATURA) & vbNewLine
   Sql = Sql & " , IDCLIENTE     = " & SqlNum(mvarIDCLIENTE) & vbNewLine
   Sql = Sql & " , DTPREV        = " & SqlDate(mvarDTPREV, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , VALOR         = " & SqlNum(mvarVALOR) & vbNewLine
   Sql = Sql & " , IDSUBDESP     = " & SqlNum(mvarIDSUBDESP) & vbNewLine
   Sql = Sql & " , HISTORICO     = " & SqlStr(mvarHISTORICO) & vbNewLine
   Sql = Sql & " , IDATENDIMENTO = " & SqlNum(mvarIDATENDIMENTO) & vbNewLine
   Sql = Sql & " , IDVENDA       = " & SqlNum(mvarIDVENDA) & vbNewLine
   Sql = Sql & " , SITFATURA      = " & SqlNum(mvarSITFATURA) & vbNewLine
   Sql = Sql & " , IDDESP        = " & SqlNum(mvarIDDESP) & vbNewLine
   Sql = Sql & " , ALTERSTAMP    = " & SqlNum(1) & vbNewLine
   Sql = Sql & " , TIMESTAMP     = " & SqlStr(mvarXDb.Sysdate(eSysDate.Data_Hora)) & vbNewLine
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'FFATURA'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA          = " & SqlStr(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDFATURA     = " & SqlNum(mvarIDFATURA) & vbNewLine

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave() As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & vbNewLine
   Sql = Sql & Me.QryUpDate
   Sql = Sql & " Else " & vbNewLine
   Sql = Sql & Me.QryInsert
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_IDFATURA, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDFATURA, IDCLIENTE, DTPREV" & vbNewLine
   Sql = Sql & ", VALOR, IDSUBDESP, HISTORICO, IDATENDIMENTO, IDVENDA" & vbNewLine
   Sql = Sql & ", SITFATURA, IDDESP, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From FFATURA" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDFATURA) Then
      Sql = Sql & " IDLOJA = " & SqlStr(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDFATURA = " & SqlNum(mvarIDFATURA) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlStr(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDFATURA) Then Sql = Sql & " IDFATURA = " & SqlNum(CStr(Ch_IDFATURA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_WHERE) Then
         If Trim(Ch_WHERE) = "" And Right(Trim(Replace(Sql, vbNewLine, "")), Len("Where")) = "Where" Then
            Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" Where  ")))
         Else
            Sql = Sql & Ch_WHERE
         End If
         Sql = Sql & " And " & vbNewLine
      End If
   End If
   Sql = Trim(Mid(Sql, 1, Len(Sql) - Len(" AND  ")))
   If Not IsMissing(Ch_ORDERBY) Then Sql = Sql & " Order By " & Ch_ORDERBY

   mvarQrySelect = Sql
   QrySelect = mvarQrySelect
End Property
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_IDFATURA, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_IDFATURA, Ch_WHERE, Ch_ORDERBY)
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
    Call ShowError(Sql)
    Pesquisar = False
End Function
Public Sub Popula(Optional pRcSet)
   If IsMissing(pRcSet) Then Set pRcSet = mvarRS
   With pRcSet
      mvarIDLOJA = !IDLOJA & ""
      mvarIDFATURA = xVal(!IDFATURA & "")
      mvarIDCLIENTE = xVal(!IDCLIENTE & "")
      mvarDTPREV = FormatDate(!DTPREV & "", True)
      mvarVALOR = xVal(!VALOR & "")
      mvarIDSUBDESP = xVal(!IDSUBDESP & "")
      mvarHISTORICO = !HISTORICO & ""
      mvarIDATENDIMENTO = xVal(!IDATENDIMENTO & "")
      mvarIDVENDA = xVal(!IDVENDA & "")
      mvarSITFATURA = xVal(!SITFATURA & "")
      mvarIDDESP = xVal(!IDDESP & "")
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = FormatDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Function Salvar(Optional ComCOMMIT = True) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarIDFATURA), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_IDFATURA As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update FFATURA Set "
   Sql = Sql & " IDLOJA = " & SqlStr(Ch_IDLOJA)
   Sql = Sql & " , IDFATURA = " & SqlNum(Ch_IDFATURA)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlStr(mvarIDLOJA)
   Sql = Sql & " and IDFATURA = " & SqlNum(mvarIDFATURA)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub
