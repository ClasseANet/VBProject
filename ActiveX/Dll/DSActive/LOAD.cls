VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DS_LOAD"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'local variable(s) to hold property value(s)
Private WithEvents mvarMe As FrmLoad
Attribute mvarMe.VB_VarHelpID = -1

Private mvarName     As String
Private mvarTitle    As String
Private mvarAnoDsvm  As String
Private mvarEmpresa  As String
Private mvarVersao   As String
Private mvarExeName  As String

Private mvarAplic    As Object

Private mvarAtiva    As Boolean
Private LerIni       As Boolean
Private IniFile      As String
'Public Property Let Ativa(ByVal vData As String)
'   mvarAtiva = vData
'End Property
Public Property Get Ativa() As Boolean
   mvarAtiva = ClsCtrl.AppAtiva(mvarAplic)
   Ativa = mvarAtiva
End Property
Public Property Let Aplic(ByVal vData As Object)
    Set mvarAplic = vData
    If TypeName(vData) <> "Nothing" Then
       IniFile = mvarAplic.Path & "\" & mvarAplic.ExeName & ".ini"
       LerIni = (ClsAutoIns.ReadIniFile(IniFile, "General", "Status") = "1")

       mvarName = Me.Name
       mvarTitle = Me.Title
       mvarVersao = Me.Versao
       mvarEmpresa = Me.Empresa
    End If
End Property
Public Property Get Aplic() As Object
    Set Aplic = mvarAplic
End Property
Public Property Let Versao(ByVal vData As String)
    mvarVersao = vData
End Property
Public Property Get Versao() As String
   If mvarVersao = "" Then mvarVersao = GetAppVersao
   Versao = mvarVersao
End Property
Public Property Let ExeName(ByVal vData As String)
    mvarExeName = vData
End Property
Public Property Get ExeName() As String
   If mvarExeName = "" Then mvarExeName = GetAppExeName
   ExeName = mvarExeName
End Property
Public Property Let Empresa(ByVal vData As String)
    mvarEmpresa = vData
End Property
Public Property Get Empresa() As String
   If mvarEmpresa = "" Then mvarEmpresa = GetAppLegalCopyright
   Empresa = mvarEmpresa
End Property
Public Property Let AnoDsvm(ByVal vData As String)
    mvarAnoDsvm = vData
End Property
Public Property Get AnoDsvm() As String
    AnoDsvm = mvarAnoDsvm
End Property
Public Property Let Title(ByVal vData As String)
    mvarTitle = vData
End Property
Public Property Get Title() As String
   If mvarTitle = "" Then mvarTitle = GetAppTitle
   Title = mvarTitle
End Property
Public Property Let Name(ByVal vData As String)
    mvarName = vData
End Property
Public Property Get Name() As String
   If mvarName = "" Then mvarName = GetAppName
   Name = mvarName
End Property
Private Sub Class_Initialize()
   Set mvarMe = New FrmLoad
End Sub
Private Sub Class_Terminate()
   Unload mvarMe
   Set mvarMe = Nothing
End Sub
Public Sub Show()
Attribute Show.VB_Description = "Exibe a Tela Inical Padrão de Programa"
   Call ClsDos.MakePath("C:\TMP\")
   mvarMe.Show vbModeless
   'mvarMe.Show vbModal
   mvarMe.ZOrder 0 '* A Frente
End Sub
Private Sub mvarMe_Activate()
   Do While mvarMe.Pnl(0).Width > mvarMe.Fre.Width + 220
      mvarMe.Pnl(0).Font.SIZE = mvarMe.Pnl(0).Font.SIZE - 1
      If mvarMe.Pnl(0).Font.SIZE <= 16 Then Exit Do
   Loop
   mvarMe.Refresh
End Sub
Private Sub mvarMe_Load()
   With mvarMe
      .Pnl(0).Font3D = 4
      .Pnl(0).Caption = " " & mvarName & " "
      .LblSistName = mvarTitle
      .lblVersao = mvarVersao
      .LblAno = mvarAnoDsvm
      .LblEmpresa = mvarEmpresa
      
      .Caption = mvarName & "  " & mvarVersao
      While .Pnl(0).Width > .Fre.Width + 220
         .Pnl(0).Font.SIZE = .Pnl(0).Font.SIZE - 1
         .Refresh
      Wend
      .Refresh
   End With
End Sub
Public Function GetAppEmpresa(Optional Aplicacao)
Attribute GetAppEmpresa.VB_Description = "Retorna o nome da Empresa que tem os direitos de execução do programa"
   If IsMissing(Aplicacao) Then Set Aplicacao = mvarAplic
   If Not Aplic Is Aplicacao Then Aplic = Aplicacao
   GetAppEmpresa = Aplicacao.CompanyName
End Function
Public Function GetAppName(Optional Aplicacao)
Attribute GetAppName.VB_Description = "Retorna o Nome do Programa"
   If IsMissing(Aplicacao) Then Set Aplicacao = mvarAplic
   If Not Aplic Is Aplicacao Then Aplic = Aplicacao
   GetAppName = Aplicacao.ProductName
End Function
Public Function GetAppTitle(Optional Aplicacao)
Attribute GetAppTitle.VB_Description = "Retorna o Título do Programa"
   If IsMissing(Aplicacao) Then Set Aplicacao = mvarAplic
   If Not Aplic Is Aplicacao Then Aplic = Aplicacao
   GetAppTitle = Aplicacao.Title
End Function
Public Function GetAppVersao(Optional Aplicacao)
Attribute GetAppVersao.VB_Description = "Retorna a Versão do Programa"
   On Error Resume Next
   If IsMissing(Aplicacao) Then Set Aplicacao = mvarAplic
   If Not Aplic Is Aplicacao Then Aplic = Aplicacao
   GetAppVersao = ClsMsg.LoadMsg(49) + Trim(CStr(Aplicacao.Major)) & "." & _
                                Trim(CStr(Aplicacao.Minor)) & "." & Trim(CStr(Aplicacao.Revision))
                                
End Function
Public Function GetAppExeName(Optional Aplicacao)
   If IsMissing(Aplicacao) Then Set Aplicacao = mvarAplic
   If Not Aplic Is Aplicacao Then Aplic = Aplicacao
   GetAppExeName = Aplicacao.ExeName
End Function
Public Function GetAppProductName(Optional Aplicacao)
   If IsMissing(Aplicacao) Then Set Aplicacao = mvarAplic
   If Not Aplic Is Aplicacao Then Aplic = Aplicacao
   GetAppProductName = Aplicacao.ProductName
End Function
Public Function GetAppDateStart()
   GetAppDateStart = Format$(Now, "dd/mm/yyy - hh:mm:ss")
End Function
Public Function GetAppLegalCopyright(Optional Aplicacao)
   If IsMissing(Aplicacao) Then Set Aplicacao = mvarAplic
   If Not Aplic Is Aplicacao Then Aplic = Aplicacao
   
   If LerIni Then
      GetAppLegalCopyright = ClsAutoIns.ReadIniFile(IniFile, "General", "LegalCopyright")
   End If

   GetAppLegalCopyright = IIf(GetAppLegalCopyright = "", Aplicacao.LegalCopyright, GetAppLegalCopyright)
End Function
Public Function SetFormat() As Boolean
'   If IsEmpty(ClsDsr.SetFormatDT_Number) Then
'      SetFormat = True
'   Else
      SetFormat = ClsDsr.SetFormatDT_Number
'   End If
End Function
