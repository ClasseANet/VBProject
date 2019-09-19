VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TB_OMAQDISPAROS"
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
Private mvarIDMAQUINA      As Double
Private mvarIDDISPARO      As Double
Private mvarIDTPMANIPULO   As Double
Private mvarIDMANIPULO     As Double
Private mvarIDLAMP         As Double
Private mvarNUMDISPARO     As Double
Private mvarQTDDISPARO     As Double
Private mvarDTDISPARO      As String
Private mvarIDATENDIMENTO  As Double
Private mvarIDSESSAO       As Double
Private mvarIDTPTRATAMENTO As Double
Private mvarIDAREA         As Double
Private mvarFLGDELETE      As Double
Private mvarIDLOJAD        As Double
Private mvarALTERSTAMP     As Double
Private mvarTIMESTAMP      As String
Public Property Let IDLOJA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJA <> vData)
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Double
   IDLOJA = mvarIDLOJA
End Property
Public Property Let IDMAQUINA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMAQUINA <> vData)
   mvarIDMAQUINA = vData
End Property
Public Property Get IDMAQUINA() As Double
   IDMAQUINA = mvarIDMAQUINA
End Property
Public Property Let IDDISPARO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDDISPARO <> vData)
   mvarIDDISPARO = vData
End Property
Public Property Get IDDISPARO() As Double
   IDDISPARO = mvarIDDISPARO
End Property
Public Property Let IDTPMANIPULO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPMANIPULO <> vData)
   mvarIDTPMANIPULO = vData
End Property
Public Property Get IDTPMANIPULO() As Double
   IDTPMANIPULO = mvarIDTPMANIPULO
End Property
Public Property Let IDMANIPULO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDMANIPULO <> vData)
   mvarIDMANIPULO = vData
End Property
Public Property Get IDMANIPULO() As Double
   IDMANIPULO = mvarIDMANIPULO
End Property
Public Property Let IDLAMP(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLAMP <> vData)
   mvarIDLAMP = vData
End Property
Public Property Get IDLAMP() As Double
   IDLAMP = mvarIDLAMP
End Property
Public Property Let NUMDISPARO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarNUMDISPARO <> vData)
   mvarNUMDISPARO = vData
End Property
Public Property Get NUMDISPARO() As Double
   NUMDISPARO = mvarNUMDISPARO
End Property
Public Property Let QTDDISPARO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarQTDDISPARO <> vData)
   mvarQTDDISPARO = vData
End Property
Public Property Get QTDDISPARO() As Double
   QTDDISPARO = mvarQTDDISPARO
End Property
Public Property Let DTDISPARO(ByVal vData As String)
   If Not mvarisDirt Then mvarisDirt = (mvarDTDISPARO <> vData)
   mvarDTDISPARO = vData
End Property
Public Property Get DTDISPARO() As String
   DTDISPARO = mvarDTDISPARO
End Property
Public Property Let IDATENDIMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDATENDIMENTO <> vData)
   mvarIDATENDIMENTO = vData
End Property
Public Property Get IDATENDIMENTO() As Double
   IDATENDIMENTO = mvarIDATENDIMENTO
End Property
Public Property Let IDSESSAO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDSESSAO <> vData)
   mvarIDSESSAO = vData
End Property
Public Property Get IDSESSAO() As Double
   IDSESSAO = mvarIDSESSAO
End Property
Public Property Let IDTPTRATAMENTO(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDTPTRATAMENTO <> vData)
   mvarIDTPTRATAMENTO = vData
End Property
Public Property Get IDTPTRATAMENTO() As Double
   IDTPTRATAMENTO = mvarIDTPTRATAMENTO
End Property
Public Property Let IDAREA(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDAREA <> vData)
   mvarIDAREA = vData
End Property
Public Property Get IDAREA() As Double
   IDAREA = mvarIDAREA
End Property
Public Property Let FLGDELETE(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarFLGDELETE <> vData)
   mvarFLGDELETE = vData
End Property
Public Property Get FLGDELETE() As Double
   FLGDELETE = mvarFLGDELETE
End Property
Public Property Let IDLOJAD(ByVal vData As Double)
   If Not mvarisDirt Then mvarisDirt = (mvarIDLOJAD <> vData)
   mvarIDLOJAD = vData
End Property
Public Property Get IDLOJAD() As Double
   IDLOJAD = mvarIDLOJAD
End Property
Public Property Let ALTERSTAMP(ByVal vData As Double)
   Dim Sql As String
 
   Sql = "Update OMAQDISPAROS Set ALTERSTAMP=" & vData & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDMAQUINA    = " & SqlNum(mvarIDMAQUINA) & vbNewLine
   Sql = Sql & " And IDDISPARO    = " & SqlNum(mvarIDDISPARO) & vbNewLine
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
Public Property Get QryInsert(Optional pAutoId As Boolean = True) As String
   Dim Sql As String
 
   Sql = "Insert Into OMAQDISPAROS (IDLOJA, IDMAQUINA, IDDISPARO, IDTPMANIPULO" & vbNewLine
   Sql = Sql & ", IDMANIPULO, IDLAMP, NUMDISPARO, QTDDISPARO, DTDISPARO" & vbNewLine
   Sql = Sql & ", IDATENDIMENTO, IDSESSAO, IDTPTRATAMENTO, IDAREA, FLGDELETE" & vbNewLine
   Sql = Sql & ", IDLOJAD, ALTERSTAMP, TIMESTAMP) " & vbNewLine
   Sql = Sql & " Output Inserted.*" & vbNewLine
   Sql = Sql & " Select "
   Sql = Sql & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDMAQUINA) & vbNewLine
   If pAutoId Then
      Sql = Sql & ", " & "(Select isNull(Max(IDDISPARO),0)+1 From OMAQDISPAROS Where IDLOJA = " & SqlNum(mvarIDLOJA) & " And IDMAQUINA = " & SqlNum(mvarIDMAQUINA) & ")" & vbNewLine
   Else
      Sql = Sql & ", " & SqlNum(mvarIDDISPARO) & vbNewLine
   End If
   Sql = Sql & ", " & SqlNum(mvarIDTPMANIPULO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDMANIPULO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDLAMP) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarNUMDISPARO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarQTDDISPARO) & vbNewLine
   Sql = Sql & ", " & SqlDate(mvarDTDISPARO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDATENDIMENTO) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDSESSAO) & vbNewLine
   Sql = Sql & ", " & IIf(mvarIDTPTRATAMENTO = 0, "Null", SqlNum(mvarIDTPTRATAMENTO)) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDAREA) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarFLGDELETE) & vbNewLine
   Sql = Sql & ", " & SqlNum(mvarIDLOJAD) & vbNewLine
   Sql = Sql & ", " & "1" & vbNewLine
   Sql = Sql & ", " & "GetDate()" & vbNewLine

   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete(Optional Ch_IDLOJA, Optional Ch_IDMAQUINA, Optional Ch_IDDISPARO) As String
   Dim Sql As String
 
   Sql = "Delete From OMAQDISPAROS" & vbNewLine
   Sql = Sql & " Where" & vbNewLine
   If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
   If Not IsMissing(Ch_IDMAQUINA) Then Sql = Sql & " IDMAQUINA = " & SqlNum(CStr(Ch_IDMAQUINA)) & " AND " & vbNewLine
   If Not IsMissing(Ch_IDDISPARO) Then Sql = Sql & " IDDISPARO = " & SqlNum(CStr(Ch_IDDISPARO)) & " AND " & vbNewLine
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
 
   Sql = "Update OMAQDISPAROS Set " & vbNewLine
   Sql = Sql & " IDLOJA         = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " , IDMAQUINA      = " & SqlNum(mvarIDMAQUINA) & vbNewLine
   Sql = Sql & " , IDDISPARO      = " & SqlNum(mvarIDDISPARO) & vbNewLine
   Sql = Sql & " , IDTPMANIPULO   = " & SqlNum(mvarIDTPMANIPULO) & vbNewLine
   Sql = Sql & " , IDMANIPULO     = " & SqlNum(mvarIDMANIPULO) & vbNewLine
   Sql = Sql & " , IDLAMP         = " & SqlNum(mvarIDLAMP) & vbNewLine
   Sql = Sql & " , NUMDISPARO     = " & SqlNum(mvarNUMDISPARO) & vbNewLine
   Sql = Sql & " , QTDDISPARO     = " & SqlNum(mvarQTDDISPARO) & vbNewLine
   Sql = Sql & " , DTDISPARO      = " & SqlDate(mvarDTDISPARO, eSysDate.Data_Hora) & vbNewLine
   Sql = Sql & " , IDATENDIMENTO  = " & SqlNum(mvarIDATENDIMENTO) & vbNewLine
   Sql = Sql & " , IDSESSAO       = " & SqlNum(mvarIDSESSAO) & vbNewLine
   Sql = Sql & " , IDTPTRATAMENTO = " & IIf(mvarIDTPTRATAMENTO = 0, "Null", SqlNum(mvarIDTPTRATAMENTO)) & vbNewLine
   Sql = Sql & " , IDAREA         = " & SqlNum(mvarIDAREA) & vbNewLine
   Sql = Sql & " , FLGDELETE      = " & SqlNum(mvarFLGDELETE) & vbNewLine
   Sql = Sql & " , IDLOJAD        = " & SqlNum(mvarIDLOJAD) & vbNewLine
   Sql = Sql & " , ALTERSTAMP     = " & "1" & vbNewLine
   Sql = Sql & " , TIMESTAMP      = " & "GetDate()" & vbNewLine
   If Not mvarXDb.ExisteReg("Select S2.NAME, S1.* From SYSOBJECTS S1 Left Join SYSOBJECTS S2 On S1.PARENT_OBJ=S2.ID WHERE S1.TYPE= 'TR' And S2.NAME= 'OMAQDISPAROS'") Then
      Sql = Sql & " Output Inserted.*" & vbNewLine
   End If
   Sql = Sql & " Where " & vbNewLine
   Sql = Sql & " IDLOJA           = " & SqlNum(mvarIDLOJA) & vbNewLine
   Sql = Sql & " And IDMAQUINA    = " & SqlNum(mvarIDMAQUINA) & vbNewLine
   Sql = Sql & " And IDDISPARO    = " & SqlNum(mvarIDDISPARO) & vbNewLine

   mvarQryUpDate = Sql
   QryUpDate = mvarQryUpDate
End Property
Public Property Get QrySave(Optional pAutoId As Boolean = True) As String
   Dim Sql As String

   Sql = " If Exists(" & Me.QrySelect() & ") " & vbNewLine
   Sql = Sql & Me.QryUpDate
   Sql = Sql & " Else " & vbNewLine
   Sql = Sql & Me.QryInsert(pAutoId:=pAutoId)
 
   mvarQrySave = Sql
   QrySave = mvarQrySave
End Property
Public Property Get QrySelect(Optional Ch_IDLOJA, Optional Ch_IDMAQUINA, Optional Ch_IDDISPARO, Optional Ch_WHERE, Optional Ch_ORDERBY) As String
   Dim Sql As String
 
   Sql = "Select IDLOJA, IDMAQUINA, IDDISPARO, IDTPMANIPULO" & vbNewLine
   Sql = Sql & ", IDMANIPULO, IDLAMP, NUMDISPARO, QTDDISPARO, DTDISPARO" & vbNewLine
   Sql = Sql & ", IDATENDIMENTO, IDSESSAO, IDTPTRATAMENTO, IDAREA, FLGDELETE" & vbNewLine
   Sql = Sql & ", IDLOJAD, ALTERSTAMP, TIMESTAMP" & vbNewLine
   Sql = Sql & " From OMAQDISPAROS" & vbNewLine
   Sql = Sql & " Where " & vbNewLine
   If IsMissing(Ch_WHERE) And IsMissing(Ch_IDLOJA) And IsMissing(Ch_IDMAQUINA) And IsMissing(Ch_IDDISPARO) Then
      Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA) & " AND " & vbNewLine
      Sql = Sql & " IDMAQUINA = " & SqlNum(mvarIDMAQUINA) & " AND " & vbNewLine
      Sql = Sql & " IDDISPARO = " & SqlNum(mvarIDDISPARO) & " AND " & vbNewLine
   Else
      If Not IsMissing(Ch_IDLOJA) Then Sql = Sql & " IDLOJA = " & SqlNum(CStr(Ch_IDLOJA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDMAQUINA) Then Sql = Sql & " IDMAQUINA = " & SqlNum(CStr(Ch_IDMAQUINA)) & " AND " & vbNewLine
      If Not IsMissing(Ch_IDDISPARO) Then Sql = Sql & " IDDISPARO = " & SqlNum(CStr(Ch_IDDISPARO)) & " AND " & vbNewLine
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
Public Function Pesquisar(Optional Ch_IDLOJA, Optional Ch_IDMAQUINA, Optional Ch_IDDISPARO, Optional Ch_WHERE, Optional Ch_ORDERBY) As Boolean
   Dim Sql     As String
   Dim bExiste As Boolean

   Sql = QrySelect(Ch_IDLOJA, Ch_IDMAQUINA, Ch_IDDISPARO, Ch_WHERE, Ch_ORDERBY)
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
      mvarIDLOJA = xVal(!IDLOJA & "")
      mvarIDMAQUINA = xVal(!IDMAQUINA & "")
      mvarIDDISPARO = xVal(!IDDISPARO & "")
      mvarIDTPMANIPULO = xVal(!IDTPMANIPULO & "")
      mvarIDMANIPULO = xVal(!IDMANIPULO & "")
      mvarIDLAMP = xVal(!IDLAMP & "")
      mvarNUMDISPARO = xVal(!NUMDISPARO & "")
      mvarQTDDISPARO = xVal(!QTDDISPARO & "")
      mvarDTDISPARO = xDate(!DTDISPARO & "", True)
      mvarIDATENDIMENTO = xVal(!IDATENDIMENTO & "")
      mvarIDSESSAO = xVal(!IDSESSAO & "")
      mvarIDTPTRATAMENTO = xVal(!IDTPTRATAMENTO & "")
      mvarIDAREA = xVal(!IDAREA & "")
      mvarFLGDELETE = xVal(!FLGDELETE & "")
      mvarIDLOJAD = xVal(!IDLOJAD & "")
      mvarALTERSTAMP = xVal(!ALTERSTAMP & "")
      mvarTIMESTAMP = xDate(!TIMESTAMP & "", True)
   End With
   mvarisDirt = False
End Sub
Public Function Salvar(Optional ComCOMMIT = True, Optional pAutoId As Boolean = True) As Boolean
   Salvar = mvarXDb.Executa(Me.QrySave(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Incluir(Optional ComCOMMIT = False, Optional pAutoId As Boolean = True) As Boolean
   Incluir = mvarXDb.Executa(Me.QryInsert(pAutoId:=pAutoId), ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function Excluir(Optional ComCOMMIT = False) As Boolean
   Excluir = mvarXDb.Executa(Me.QryDelete(mvarIDLOJA, mvarIDMAQUINA, mvarIDDISPARO), ComCOMMIT)
End Function
Public Function Alterar(Optional ComCOMMIT = False) As Boolean
   Alterar = mvarXDb.Executa(Me.QryUpDate, ComCOMMIT)

   On Error Resume Next
   Call Popula(mvarXDb.ADORs)
End Function
Public Function AlterarChave(Ch_IDLOJA As String, Ch_IDMAQUINA As String, Ch_IDDISPARO As String, Optional ComCOMMIT = False) As Integer
   Dim Sql As String
 
   Sql = "Update OMAQDISPAROS Set "
   Sql = Sql & " IDLOJA = " & SqlNum(Ch_IDLOJA)
   Sql = Sql & " , IDMAQUINA = " & SqlNum(Ch_IDMAQUINA)
   Sql = Sql & " , IDDISPARO = " & SqlNum(Ch_IDDISPARO)
   Sql = Sql & " Where "
   Sql = Sql & " IDLOJA = " & SqlNum(mvarIDLOJA)
   Sql = Sql & " and IDMAQUINA = " & SqlNum(mvarIDMAQUINA)
   Sql = Sql & " and IDDISPARO = " & SqlNum(mvarIDDISPARO)
   AlterarChave = mvarXDb.Executa(Sql, ComCOMMIT)
End Function
Private Sub Class_Terminate()
   Set mvarXDb = Nothing
   Set mvarRS = Nothing
End Sub