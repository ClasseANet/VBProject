VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TB_PCG_IMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
Private mvardBase As Object
Private mvarEXISTE As Integer

Private mvarQryInsert As String
Private mvarQryUpDate As String
Private mvarQryDelete As String

Private mvarPCG_NINF As String
Private mvarHIS_PCG_SIT As String
Private mvarARM_COD As String
Private mvarBOX_NUM As String
Private mvarCRG_NUM As String
Private mvarHCRG_NUM As String
Private mvarHIS_PCG_DAT As String
Private mvarHIS_PCG_HOR As String
Private mvarMOT_REA_COD As String
Private mvarNUM_TERMO As String
Private mvarPCG_DES As String
Private mvarPCG_DST As String
Private mvarPCG_IDT_DSIC As String
Private mvarPCG_PES As String
Private mvarPCG_VOL As String
Private mvarSET_NUM As String
Private mvarPCG_TEMP As String
Private mvarAER_COD As String
Public Property Let AER_COD(ByVal vData As String)
   mvarAER_COD = vData
End Property
Public Property Get AER_COD() As String
   AER_COD = mvarAER_COD
End Property
Public Property Let PCG_NINF(ByVal vData As String)
   mvarPCG_NINF = vData
End Property
Public Property Get PCG_NINF() As String
   PCG_NINF = mvarPCG_NINF
End Property
Public Property Let HIS_PCG_SIT(ByVal vData As String)
   mvarHIS_PCG_SIT = vData
End Property
Public Property Get HIS_PCG_SIT() As String
   HIS_PCG_SIT = mvarHIS_PCG_SIT
End Property
Public Property Let ARM_COD(ByVal vData As String)
   mvarARM_COD = vData
End Property
Public Property Get ARM_COD() As String
   ARM_COD = mvarARM_COD
End Property
Public Property Let BOX_NUM(ByVal vData As String)
   mvarBOX_NUM = vData
End Property
Public Property Get BOX_NUM() As String
   BOX_NUM = mvarBOX_NUM
End Property
Public Property Let CRG_NUM(ByVal vData As String)
   mvarCRG_NUM = vData
End Property
Public Property Get CRG_NUM() As String
   CRG_NUM = mvarCRG_NUM
End Property
Public Property Let HCRG_NUM(ByVal vData As String)
   mvarHCRG_NUM = vData
End Property
Public Property Get HCRG_NUM() As String
   HCRG_NUM = mvarHCRG_NUM
End Property
Public Property Let HIS_PCG_DAT(ByVal vData As String)
   mvarHIS_PCG_DAT = vData
End Property
Public Property Get HIS_PCG_DAT() As String
   HIS_PCG_DAT = mvarHIS_PCG_DAT
End Property
Public Property Let HIS_PCG_HOR(ByVal vData As String)
   mvarHIS_PCG_HOR = vData
End Property
Public Property Get HIS_PCG_HOR() As String
   HIS_PCG_HOR = mvarHIS_PCG_HOR
End Property
Public Property Let MOT_REA_COD(ByVal vData As String)
   mvarMOT_REA_COD = vData
End Property
Public Property Get MOT_REA_COD() As String
   MOT_REA_COD = mvarMOT_REA_COD
End Property
Public Property Let NUM_TERMO(ByVal vData As String)
   mvarNUM_TERMO = vData
End Property
Public Property Get NUM_TERMO() As String
   NUM_TERMO = mvarNUM_TERMO
End Property
Public Property Let PCG_DES(ByVal vData As String)
   mvarPCG_DES = vData
End Property
Public Property Get PCG_DES() As String
   PCG_DES = mvarPCG_DES
End Property
Public Property Let PCG_DST(ByVal vData As String)
   mvarPCG_DST = vData
End Property
Public Property Get PCG_DST() As String
   PCG_DST = mvarPCG_DST
End Property
Public Property Let PCG_IDT_DSIC(ByVal vData As String)
   mvarPCG_IDT_DSIC = vData
End Property
Public Property Get PCG_IDT_DSIC() As String
   PCG_IDT_DSIC = mvarPCG_IDT_DSIC
End Property
Public Property Let PCG_PES(ByVal vData As String)
   mvarPCG_PES = vData
End Property
Public Property Get PCG_PES() As String
   PCG_PES = mvarPCG_PES
End Property
Public Property Let PCG_VOL(ByVal vData As String)
   mvarPCG_VOL = vData
End Property
Public Property Get PCG_VOL() As String
   PCG_VOL = mvarPCG_VOL
End Property
Public Property Let SET_NUM(ByVal vData As String)
   mvarSET_NUM = vData
End Property
Public Property Get SET_NUM() As String
   SET_NUM = mvarSET_NUM
End Property
Public Property Let PCG_TEMP(ByVal vData As String)
   mvarPCG_TEMP = vData
End Property
Public Property Get PCG_TEMP() As String
   PCG_TEMP = mvarPCG_TEMP
End Property
Public Property Get EXISTE() As Integer
   EXISTE = mvarEXISTE
End Property
Public Property Let dBase(ByVal vData As Object)
   Set mvardBase = vData
End Property
Public Property Get dBase() As Object
   Set dBase = mvardBase
End Property
Public Property Get QryInsert() As String
   Dim Sql$
   Sql = "insert into PCG_IMP (PCG_NINF, HIS_PCG_SIT, ARM_COD, BOX_NUM, CRG_NUM"
   Sql = Sql & " , HCRG_NUM, HIS_PCG_DAT, HIS_PCG_HOR, MOT_REA_COD,NUM_TERMO"
   Sql = Sql & ", PCG_DES, PCG_DST, PCG_IDT_DSIC, PCG_PES, PCG_VOL"
   Sql = Sql & ", SET_NUM, PCG_TEMP, AER_COD) "
   Sql = Sql & " Values "
   Sql = Sql & "("
   Sql = Sql & Aspas(mvarPCG_NINF)
   Sql = Sql & ", " & Aspas(mvarHIS_PCG_SIT)
   Sql = Sql & ", " & Aspas(mvarARM_COD)
   Sql = Sql & ", " & Aspas(mvarBOX_NUM)
   Sql = Sql & ", " & Aspas(mvarCRG_NUM)
   Sql = Sql & ", " & Aspas(mvarHCRG_NUM)
   Sql = Sql & ", " & SqlDate(mvarHIS_PCG_DAT)
   Sql = Sql & ", " & Aspas(mvarHIS_PCG_HOR)
   Sql = Sql & ", " & Aspas(mvarMOT_REA_COD)
   Sql = Sql & ", " & Aspas(mvarNUM_TERMO)
   Sql = Sql & ", " & Aspas(mvarPCG_DES)
   Sql = Sql & ", " & Aspas(mvarPCG_DST)
   Sql = Sql & ", " & Aspas(mvarPCG_IDT_DSIC)
   Sql = Sql & ", " & mvarPCG_PES
   Sql = Sql & ", " & mvarPCG_VOL
   Sql = Sql & ", " & Aspas(mvarSET_NUM)
   Sql = Sql & ", " & mvarPCG_TEMP
   Sql = Sql & ", " & Aspas(mvarAER_COD)
   Sql = Sql & ")"
   mvarQryInsert = Sql
   QryInsert = mvarQryInsert
End Property
Public Property Get QryDelete() As String
   Dim Sql$
   If Trim(mvarPCG_NINF) = "" Then Exit Property
   Sql$ = "Delete * From PCG_IMP"
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " PCG_NINF = " & Aspas(mvarPCG_NINF)
   Sql$ = Sql$ & " AER_COD = " & Aspas(mvarAER_COD)
   mvarQryDelete = Sql
   QryDelete = mvarQryDelete
End Property
Public Property Get QryUpDate() As String
   Dim Sql$
   Sql$ = "update PCG_IMP set "
   Sql$ = Sql$ & " PCG_NINF = " & Aspas(mvarPCG_NINF)
   Sql$ = Sql$ & " , HIS_PCG_SIT = " & Aspas(mvarHIS_PCG_SIT)
   Sql$ = Sql$ & " , ARM_COD = " & Aspas(mvarARM_COD)
   Sql$ = Sql$ & " , BOX_NUM = " & Aspas(mvarBOX_NUM)
   Sql$ = Sql$ & " , CRG_NUM = " & Aspas(mvarCRG_NUM)
   Sql$ = Sql$ & " , HCRG_NUM = " & Aspas(mvarHCRG_NUM)
   Sql$ = Sql$ & " , HIS_PCG_DAT = " & SqlDate(mvarHIS_PCG_DAT)
   Sql$ = Sql$ & " , HIS_PCG_HOR = " & Aspas(mvarHIS_PCG_HOR)
   Sql$ = Sql$ & " , MOT_REA_COD = " & Aspas(mvarMOT_REA_COD)
   Sql$ = Sql$ & " , NUM_TERMO = " & Aspas(mvarNUM_TERMO)
   Sql$ = Sql$ & " , PCG_DES = " & Aspas(mvarPCG_DES)
   Sql$ = Sql$ & " , PCG_DST = " & Aspas(mvarPCG_DST)
   Sql$ = Sql$ & " , PCG_IDT_DSIC = " & Aspas(mvarPCG_IDT_DSIC)
   Sql$ = Sql$ & " , PCG_PES = " & mvarPCG_PES
   Sql$ = Sql$ & " , PCG_VOL = " & mvarPCG_VOL
   Sql$ = Sql$ & " , SET_NUM = " & Aspas(mvarSET_NUM)
   Sql$ = Sql$ & " , PCG_TEMP = " & mvarPCG_TEMP
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " PCG_NINF = " & Aspas(mvarPCG_NINF)
   Sql$ = Sql$ & " AER_COD = " & Aspas(mvarAER_COD)
   mvarQryUpDate = Sql$
   QryUpDate = mvarQryUpDate
End Property
Public Function Incluir() As Integer
   Call mvardBase.Executa(Me.QryInsert)
   Incluir = mvardBase.CodeSql
End Function
Public Function Gravar(Optional ByVal ExibeResult = True) As Variant
   Dim Result
   Select Case mvarEXISTE
      Case ALTERACAO: Result = Alterar
      Case INCLUSAO: Result = Incluir()
   End Select
   If Not ExibeResult Then Exit Function
   If Result = FOUND Then
      Call ExibirAviso(LoadMsg(34), LoadMsg(57))
   Else
      Call ExibirAviso(LoadMsg(48), LoadMsg(57))
   End If
End Function
Public Function GetSelect(Ch_PCG_NINF$, Ch_AER_COD$) As Integer
   Dim Sql$
   mvarPCG_NINF = Ch_PCG_NINF
   mvarAER_COD = Ch_AER_COD

   Sql$ = "select distinct PCG_NINF, HIS_PCG_SIT, ARM_COD,BOX_NUM"
   Sql$ = Sql$ & ", CRG_NUM, HCRG_NUM, HIS_PCG_DAT, HIS_PCG_HOR"
   Sql$ = Sql$ & ", MOT_REA_COD, NUM_TERMO, PCG_DES, PCG_DST"
   Sql$ = Sql$ & ", PCG_IDT_DSIC, PCG_PES, PCG_VOL, SET_NUM, PCG_TEMP"
   Sql$ = Sql$ & ", AER_COD "
   Sql$ = Sql$ & " From PCG_IMP"
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " PCG_NINF = " & Aspas(mvarPCG_NINF)
   Sql$ = Sql$ & " and AER_COD = " & Aspas(mvarAER_COD)
   Call mvardBase.AbreTabela(Sql$)
   Select Case mvardBase.CodeSql
      Case FOUND
         GetSelect = ALTERACAO
         With mvardBase.Dys
            mvarPCG_NINF = !PCG_NINF
            mvarAER_COD = !AER_COD
            mvarHIS_PCG_SIT = !HIS_PCG_SIT & ""
            mvarARM_COD = !ARM_COD & ""
            mvarBOX_NUM = !BOX_NUM & ""
            mvarCRG_NUM = !CRG_NUM & ""
            mvarHCRG_NUM = !HCRG_NUM & ""
            mvarHIS_PCG_DAT = Format(!HIS_PCG_DAT & "", "DD/MM/YYYY")
            mvarHIS_PCG_HOR = !HIS_PCG_HOR & ""
            mvarMOT_REA_COD = !MOT_REA_COD & ""
            mvarNUM_TERMO = !NUM_TERMO & ""
            mvarPCG_DES = !PCG_DES & ""
            mvarPCG_DST = !PCG_DST & ""
            mvarPCG_IDT_DSIC = !PCG_IDT_DSIC & ""
            mvarPCG_PES = !PCG_PES & ""
            mvarPCG_VOL = !PCG_VOL & ""
            mvarSET_NUM = !SET_NUM & ""
            mvarPCG_TEMP = !PCG_TEMP & ""
            .Close
         End With
      Case NOTFOUND: GetSelect = INCLUSAO
      Case ERRO: GetSelect = ERRO
   End Select
   mvarEXISTE = GetSelect
End Function
Public Function Excluir() As Variant
   If mvardBase.GravarTabela(Me.QryDelete, EXCLUSAO) = vbYes Then
      Excluir = True
   End If
End Function
Public Function Alterar() As Integer
   Call mvardBase.Executa(Me.QryUpDate)
   Alterar = mvardBase.CodeSql
End Function
Public Function AlterarChave(Ch_PCG_NINF$) As Integer
   Dim Sql$
   Sql$ = "update PCG_IMP set "
   Sql$ = Sql$ & " PCG_NINF = " & Aspas(Ch_PCG_NINF)
   Sql$ = Sql$ & " Where "
   Sql$ = Sql$ & " PCG_NINF = " & Aspas(mvarPCG_NINF)
   Sql$ = Sql$ & " AER_COD = " & Aspas(mvarAER_COD)
   Call mvardBase.Executa(Sql)
   AlterarChave = mvardBase.CodeSql
End Function
Private Sub Class_Terminate()
   Set mvardBase = Nothing
End Sub
