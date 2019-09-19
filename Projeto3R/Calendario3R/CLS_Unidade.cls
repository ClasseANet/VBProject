VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLS_Unidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private mvarSalas As Collection
Private mvarSys      As Object
Private mvarIDLOJA   As Integer
Public Property Let IDLOJA(ByVal vData As Integer)
   Dim bSel As Boolean
   Dim Sql As String
   Dim MyRs As Object
   
   If mvarIDLOJA <> vData Then
      Dim TbSala As Object
      Dim cSala  As CLS_SALA

      Set mvarSalas = New Collection
      bSel = True
      mvarSys.Propriedades("ExibeSala") = mvarSys.GetParam("EXIBESALA", Default:=1)
      
      Sql = "Select * "
      Sql = Sql & " From OSALA"
      Sql = Sql & " Where IDLOJA=" & vData
      Sql = Sql & " And ATIVO=1"
      Sql = Sql & " Order By IDSALA"
      If mvarSys.xDb.AbreTabela(Sql, MyRs) Then
         While Not MyRs.EOF
            Set cSala = New CLS_SALA
            Set cSala.Sys = mvarSys
            cSala.IDLOJA = MyRs("IDLOJA")
            cSala.IDSALA = MyRs("IDSALA")
            cSala.NMSALA = IIf(mvarSys.Propriedades("ExibeSala") = 0, "", "Sala ") & MyRs("CODSALA")
            cSala.ScheduleID = (MyRs("IDLOJA") * 1000) + MyRs("IDSALA")
            cSala.Selecionada = bSel
            mvarSalas.Add cSala, CStr(MyRs("IDSALA"))
            
            MyRs.MoveNext
            Set cSala = Nothing
            bSel = False
         Wend
      Else
         If TbSala Is Nothing Then
            Set TbSala = CriarObjeto("BANCO_3R.TB_OSALA")
            Set TbSala.xDb = mvarSys.xDb
            'Call TbSala.Pesquisar(Ch_IDLOJA:=mvarIDLOJA, Ch_IDSALA:=vData)
         End If
         TbSala.IDLOJA = vData
         TbSala.IDSALA = 1
         TbSala.CODSALA = "01"
         TbSala.ATIVO = 1
         TbSala.DTOPERACAO = mvarSys.xDb.Sysdate(1)
         If TbSala.Salvar Then
            Set cSala = New CLS_SALA
            Set cSala.Sys = mvarSys
            cSala.IDLOJA = TbSala.IDLOJA
            cSala.IDSALA = TbSala.IDSALA
            cSala.NMSALA = IIf(mvarSys.Propriedades("ExibeSala") = 0, "", "Sala ") & TbSala.CODSALA
            cSala.ScheduleID = (TbSala.IDLOJA * 1000) + TbSala.IDSALA
            cSala.Selecionada = bSel
            mvarSalas.Add cSala, CStr(TbSala.IDSALA)
         End If
      End If
   End If
   mvarIDLOJA = vData
End Property
Public Property Get IDLOJA() As Integer
   IDLOJA = mvarIDLOJA
End Property
Public Property Set Sys(ByVal vData As Object)
   Set mvarSys = vData
End Property
Public Property Get Sys() As Object
    Set Sys = mvarSys
End Property
Public Property Set Salas(ByRef vData As Collection)
   Set mvarSalas = vData
End Property
Public Property Get Salas() As Collection
    Set Salas = mvarSalas
End Property
Private Sub Class_Initialize()
   Set mvarSalas = New Collection
End Sub
Private Sub Class_Terminate()
   Set mvarSalas = Nothing
End Sub
