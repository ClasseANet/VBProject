VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Impressao"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Enum eDEST
  ToScreen = 0
  ToPrint = 1
  ToFile = 2
  ToMapi = 3
End Enum
Private mvarDBase       As Object
Private mvarCryRpt      As Object
Private mvarAplic       As Object
Private mvarRel_Mdb     As String
Private mvarRel_Drv     As String
Private mvarConexao     As String
Private mvarDSN         As String
Private mvarUID         As String
Private mvarPWD         As String
Private mvarDSQ         As String
Private mvarRpt_Name    As String
Private mvarRpt_Drive   As String
Private mvarRptCadName  As String
Private mvarRptCadMemo  As String
Private mvarTitulo      As String
Private mvarDESTINO     As eDEST
Private mvarIdioma      As Integer
Private wFactor         As Single
Private hFactor         As Single

Private mvarPrinterName   As String
Private mvarPrinterPort   As String
Private mvarPrinterDriver As String

Private mvarFormulas  As Collection
Public Property Let PrinterName(ByVal vData As String)
   mvarPrinterName = vData
End Property
Public Property Get PrinterName() As String
   PrinterName = mvarPrinterName
End Property
Public Property Let PrinterPort(ByVal vData As String)
   mvarPrinterPort = vData
End Property
Public Property Get PrinterPort() As String
   PrinterPort = mvarPrinterPort
End Property
Public Property Let PrinterDriver(ByVal vData As String)
   mvarPrinterDriver = vData
End Property
Public Property Get PrinterDriver() As String
   PrinterDriver = mvarPrinterDriver
End Property
Public Property Let Formulas(ByVal vData As Collection)
   Set mvarFormulas = vData
End Property
Public Property Set Formulas(ByVal vData As Collection)
   Set mvarFormulas = vData
End Property
Public Property Get Formulas() As Collection
   If mvarFormulas Is Nothing Then
      Set mvarFormulas = New Collection
   End If
   Set Formulas = mvarFormulas
End Property
Public Property Let CryRpt(ByVal vData As Object)
   Set mvarCryRpt = vData
End Property
Public Property Get CryRpt() As Object
   Set CryRpt = mvarCryRpt
End Property
Public Property Let Dbase(ByVal vData As Object)
   Set mvarDBase = vData
End Property
Public Property Set Dbase(ByVal vData As Object)
   Set mvarDBase = vData
End Property
Public Property Get Dbase() As Object
   Set Dbase = mvarDBase
End Property
Public Property Let Aplic(ByVal vData As Object)
   Set mvarAplic = vData
   If Trim(mvarRpt_Drive) = "" Then
      mvarRpt_Drive = App.Path & "\"
   End If
End Property
Public Property Get Aplic() As Object
   Set Aplic = mvarAplic
End Property
Public Property Let Idioma(ByVal vData As Integer)
   mvarIdioma = vData
End Property
Public Property Get Idioma() As Integer
   Idioma = mvarIdioma
End Property
Public Property Let Conexao(ByVal vData As String)
   mvarConexao = vData
End Property
Public Property Get Conexao() As String
   Conexao = mvarConexao
End Property
Public Property Let DSN(ByVal vData As String)
   mvarDSN = vData
End Property
Public Property Get DSN() As String
   DSN = mvarDSN
End Property
Public Property Let UID(ByVal vData As String)
   mvarUID = vData
End Property
Public Property Get UID() As String
   UID = mvarUID
End Property
Public Property Let PWD(ByVal vData As String)
   mvarPWD = vData
End Property
Public Property Get PWD() As String
   PWD = mvarPWD
End Property
Public Property Let DSQ(ByVal vData As String)
   mvarDSQ = vData
End Property
Public Property Get DSQ() As String
   DSQ = mvarDSQ
End Property
Public Property Let Rel_Mdb(ByVal vData As String)
   mvarRel_Mdb = vData
End Property
Public Property Get Rel_Mdb() As String
   Rel_Mdb = mvarRel_Mdb
End Property
Public Property Let Rel_Drv(ByVal vData As String)
   mvarRel_Drv = vData
End Property
Public Property Get Rel_Drv() As String
   Rel_Drv = mvarRel_Drv
End Property
Public Property Let Rpt_Name(ByVal vData As String)
   mvarRpt_Name = vData
End Property
Public Property Get Rpt_Name() As String
   Rpt_Name = mvarRpt_Name
End Property
Public Property Let Rpt_Drive(ByVal vData As String)
   If VBA.Right$(vData, 1) <> "\" Then vData = vData & "\"
   mvarRpt_Drive = vData
End Property
Public Property Get Rpt_Drive() As String
   Rpt_Drive = mvarRpt_Drive
End Property
Public Property Let RptCadMemo(ByVal vData As String)
   mvarRptCadMemo = vData
End Property
Public Property Get RptCadMemo() As String
   RptCadMemo = mvarRptCadMemo
End Property
Public Property Let RptCadName(ByVal vData As String)
   mvarRptCadName = vData
End Property
Public Property Get RptCadName() As String
   RptCadName = mvarRptCadName
End Property
Public Property Let Titulo(ByVal vData As String)
   mvarTitulo = vData
End Property
Public Property Get Titulo() As String
   Titulo = mvarTitulo
End Property
Public Property Let DESTINO(vData As eDEST)
   mvarDESTINO = vData
End Property
Public Property Get DESTINO() As eDEST
   DESTINO = mvarDESTINO
End Property
Public Function ImprimeGrid(Text As Object)
   Dim Atributos()
   Dim Indices()
   Dim DbRel      As DS_BANCO
   Dim CollCab    As Collection
   Dim n          As Variant
   Dim Sql        As String
   Dim cLIN_Det   As String
   Dim Cab        As String
   Dim Lin        As Integer
   Dim Col        As Integer
   Dim Tam        As Integer
   Dim TotTam     As Integer
   Dim nCols      As Integer
   Dim cTAB       As Integer
   Dim TamCol     As Integer
   Dim Escala     As Double
   
   Set DbRel = New DS_BANCO
   Set CollCab = New Collection
   
   Const HPg% = 93 '116  'Nº de Caracteres da Linha
'   On Error GoTo Fim
   
   Call ClsDos.Del(mvarRel_Drv$ & mvarRel_Mdb$)
   Call CriarBD(mvarRel_Drv$ & mvarRel_Mdb$)
   DbRel.dbTipo = Access
   DbRel.isADO = False
   Call DbRel.SrvConecta(mvarRel_Drv$, mvarRel_Mdb$)
   If Not DbRel.Conectado Then
      Exit Function
   End If
   
   '* Cria Tabela
   ReDim Atributos(1)
   'Atributos(0, 1) = "LINHA"
   'Atributos(0, 2) = "10" ' - Text
   'Atributos(0, 3) = CStr(HPg)
   Atributos(0) = Array("LINHA", dbText, CStr(HPg))
   

   ReDim Indices(0)
   Call DbRel.CriarTabela("TB_GRID", Atributos, Indices)
   '* Define Cabeçalho
   nCols = 1
   Col = 0
   Lin = 0
   While Col < Text.Cols
      If Text.ColWidth(Col) > 0 Then
         TamCol = Text.ColWidth(Col) / wFactor / 120
         TotTam = TotTam + TamCol + 1
         Tam = TamCol - Len(Trim(Text.TextMatrix(Lin, Col))) + 1
         Tam = IIf(Tam <= 0, 1, Tam)
'         TotTam = TotTam + Tam
         CollCab.Add UCase(Trim(Text.TextMatrix(Lin, Col))) & Space(Tam)
      End If
      Col = Col + 1
   Wend
   nCols = Col + 1
   cTAB = Int((HPg - TotTam) / nCols)
   cTAB = IIf(cTAB < 0, 0, cTAB)
'   cMOD = (HPg - TotTam) Mod nCols
   For Each n In CollCab
      i = i + 1
      Tam = Len(n) - Len(Trim(n))
      Tam = IIf(Tam < 0, 1, Tam)
      
      Cab$ = Cab$ & Space(cTAB) & Trim(n) & Space(Tam)
'      If i >= (nCols - cMOD - 1) Then Cab$ = Cab$ & Space(1)
'      If CollCab.Count <= 2 Then
'         Cab$ = Cab$ & Space(cTAB * 0.83) & Trim(n) & Space(Tam * 0.83)
'      Else
'         Cab$ = Cab$ & Space(cTAB * 0.8) & Trim(n) & Space(Tam * 0.8)
'      End If
   Next
   '* Define Linhas
   Lin = Lin + 1
   While Lin < Text.Rows
      Col = 0
      cLIN_Det = ""
      While Col < Text.Cols
         If Text.ColWidth(Col) > 0 Then
            TamCol = (Text.ColWidth(Col) / wFactor / 120)
            Tam = TamCol - Len(Trim(Text.TextMatrix(Lin, Col))) + 1
            Tam = IIf(Tam < 0, 1, Tam)
            cLIN_Det = cLIN_Det & Space(cTAB%) & Trim(Mid(Text.TextMatrix(Lin, Col), 1, TamCol)) & Space(Tam)
         End If
         Col = Col + 1
      Wend
      Sql = "insert into TB_GRID (LINHA) "
      Sql = Sql & " Values "
      Sql = Sql & "("
      Sql = Sql & SqlStr(cLIN_Det)
      Sql = Sql & ")"
      Call DbRel.Executa(Sql)
      Lin = Lin + 1
   Wend
   mvarRpt_Name = mvarRptCadName
   mvarDESTINO = ToScreen
   Call Me.LimpaFormulas
   mvarCryRpt.Formulas(0) = "CABECALHO=" & SqlStr(Cab)
   mvarCryRpt.Formulas(1) = "TITULO1=" & SqlStr(UCase(mvarAplic.Title) & " " & Trim(CStr(mvarAplic.Major)) & "." & Trim(CStr(mvarAplic.Minor))) 'LoadMsg(49)
   mvarCryRpt.Formulas(2) = "TITULO2=" & SqlStr(mvarTitulo)
   
   Call ClsCtrl.ExtractResData("CAD", "RPT", mvarRpt_Drive & mvarRpt_Name)
   
   Call Me.ExibeRelatorio
   DbRel.SrvDesconecta
   ImprimeGrid = True
   
   Set CollCab = Nothing
   Set DbRel = Nothing
   Call ClsDos.Del(mvarRpt_Drive & mvarRptCadName)
   Exit Function
Fim:
   If Err = 75 Then
      DbRel.SrvDesconecta
   Else
      ClsMsg.ShowError
  End If
End Function
Public Function ImprimeText(Text As Object, Optional CabDetalhe As String)
   Dim Atributos()
   Dim Indices()
   Dim DbRel      As DS_BANCO
   Dim CollCab    As Collection
   Dim n          As Variant
   Dim Sql        As String
   Dim cLIN_Det   As String
   Dim Cab        As String
   Dim Lin        As Integer
   Dim Col        As Integer
   Dim Tam        As Integer
   Dim TotTam     As Integer
   Dim nCols      As Integer
   Dim cTAB       As Integer
   Dim TamCol     As Integer
   Dim Escala     As Double
   
   Set DbRel = New DS_BANCO
   Set CollCab = New Collection
   
   Const HPg% = 93 '116  'Nº de Caracteres da Linha
'   On Error GoTo Fim
   
   Call ClsDos.Del(mvarRel_Drv$ & mvarRel_Mdb$)
   Call CriarBD(mvarRel_Drv$ & mvarRel_Mdb$)
   DbRel.dbTipo = Access
   DbRel.isADO = False
   Call DbRel.SrvConecta(mvarRel_Drv$, mvarRel_Mdb$)
   If Not DbRel.Conectado Then
      Exit Function
   End If
   
   '* Cria Tabela
   ReDim Atributos(1)
   'Atributos(0, 1) = "LINHA"
   'Atributos(0, 2) = "10" ' - Text
   'Atributos(0, 3) = CStr(HPg)
   Atributos(0) = Array("LINHA", dbMemo)
   

   ReDim Indices(0)
   Call DbRel.CriarTabela("TB_GRID", Atributos, Indices)
   '* Define Cabeçalho
   nCols = 1
   Col = 0
   Lin = 0
     
   Cab$ = CabDetalhe
   '* Incluir Texto
   cLIN_Det = Text.Text
   Sql = "insert into TB_GRID (LINHA) "
   Sql = Sql & " Values "
   Sql = Sql & "("
   Sql = Sql & SqlStr(cLIN_Det)
   Sql = Sql & ")"
   Call DbRel.Executa(Sql)
  
   
   mvarDESTINO = ToScreen
   Call Me.LimpaFormulas
   mvarCryRpt.Formulas(0) = "CABECALHO=" & SqlStr(Cab)
   mvarCryRpt.Formulas(1) = "TITULO1=" & SqlStr(UCase(mvarAplic.Title) & " " & Trim(CStr(mvarAplic.Major)) & "." & Trim(CStr(mvarAplic.Minor))) 'LoadMsg(49)
   mvarCryRpt.Formulas(2) = "TITULO2=" & SqlStr(mvarTitulo)
   
   Call ClsCtrl.ExtractResData("CADMEMO", "RPT", mvarRpt_Drive & mvarRptCadMemo)
   
   mvarRpt_Drive = mvarRpt_Drive
   mvarRpt_Name = mvarRptCadMemo
   Call Me.ExibeRelatorio
   DbRel.SrvDesconecta
   ImprimeText = True
   
   Set CollCab = Nothing
   Set DbRel = Nothing
   
   On Error Resume Next
   DoEvents
   Call ClsDos.Del(mvarRpt_Drive & mvarRptCadMemo, False)
   Exit Function
Fim:
   If Err = 75 Then
      DbRel.SrvDesconecta
   Else
      ClsMsg.ShowError
  End If
End Function
Public Function ExibeRelatorio(Optional Flag_RptFile = True, Optional Flag_Hora = True)
   Dim i%
   Dim sAux As String
   Dim sAux2 As String
   
   On Error GoTo Fim
   If Trim(mvarRpt_Drive) = "" Or Trim(mvarRpt_Name) = "" Then Exit Function
   
   With mvarCryRpt
      '** Área de Fórmulas *****************************************************
      If .Formulas(i) = "" Then
         i = -1
      Else
         While .Formulas(i) <> ""
            i = i + 1
         Wend
      End If
      If Flag_RptFile Then
         i = i + 1
         .Formulas(i) = "RPTFILE=" & SqlStr(UCase(mvarRpt_Name))
      End If
      If Flag_Hora Then
         i = i + 1
         .Formulas(i) = "HORA=" & SqlStr(Format(Now(), "hh:mm:ss"))
      End If
      
      If Not mvarFormulas Is Nothing Then
         For Each n In mvarFormulas
            i = i + 1
            If InStr(InStr(n, "="), n, "'") > 0 Then
               sAux = Trim(Mid(n, 1, InStr(InStr(n, "="), n, "'")))
               sAux2 = Mid(Trim(n), Len(sAux) + 1, Len(n) - Len(sAux) - 1)
               sAux2 = ClsDsr.StrReplace(sAux2, "'", "' + Chr(39) + '")
               sAux = sAux & sAux2 & "'"
            Else
               sAux = n
            End If
            .Formulas(i) = sAux
         Next
      End If
      
      '** Área de Localização *****************************************************
      If Not mvarDBase Is Nothing Then
         mvarDSN = mvarDBase.Server
         mvarUID = mvarDBase.UID
         mvarPWD = mvarDBase.PWD
         mvarDSQ = mvarDBase.dbName
      End If
      If mvarDSN <> "" And mvarDSQ <> "" Then
         mvarConexao = "DSN=" & mvarDSN
         mvarConexao = mvarConexao & ";UID=" & mvarUID
         mvarConexao = mvarConexao & ";PWD=" & mvarPWD
         mvarConexao = mvarConexao & ";DSQ=" & mvarDSQ
      End If

      If mvarConexao = "" Then
         .DataFiles(0) = mvarRel_Drv & mvarRel_Mdb
      Else
         .Connect = mvarConexao
      End If
      .ReportFileName = mvarRpt_Drive & mvarRpt_Name
      
      
      '** Configuração de Impressora **********************************************
      If Trim(mvarPrinterPort) = "" Then mvarPrinterPort = .PrinterPort
      If Trim(mvarPrinterDriver) = "" Then mvarPrinterDriver = .PrinterDriver
      If Trim(mvarPrinterName) = "" Then mvarPrinterName = .PrinterName
      
      If Trim(mvarPrinterPort) = "" Then mvarPrinterPort = Printer.Port
      If Trim(mvarPrinterDriver) = "" Then mvarPrinterDriver = Printer.DriverName
      If Trim(mvarPrinterName) = "" Then mvarPrinterName = Printer.DeviceName
      
      .PrinterPort = mvarPrinterPort
      .PrinterDriver = mvarPrinterDriver
      .PrinterName = mvarPrinterName

      
      '** Área de Formatação de Tela **********************************************
      
      .Destination = mvarDESTINO
      .WindowTitle = UCase(mvarTitulo)
      '.WindowLeft = 0
      '.WindowTop = 0
      '.WindowWidth = Screen.Width / Screen.TwipsPerPixelX
      '.WindowHeight = Screen.Height / Screen.TwipsPerPixelY
      
      .WindowState = crptMaximized
      .WindowState = crptMaximized
      .WindowControls = True
      .WindowShowCancelBtn = False
      .WindowShowCloseBtn = True
      .WindowShowExportBtn = True
      .WindowShowGroupTree = False
      .WindowShowNavigationCtls = True
      .WindowShowPrintBtn = True
      .WindowShowPrintSetupBtn = True
      .WindowShowProgressCtls = False
      .WindowShowRefreshBtn = False
      .WindowShowSearchBtn = True
      .WindowShowZoomCtl = True

      .Action = 1
      .PageCount
      .PageZoom 0
   End With
   ExibeRelatorio = True
   Exit Function
Fim:
   ClsMsg.ShowError
End Function
Public Sub LimpaFormulas(Optional ByRef pCryRpt)
   Dim i%
   If Not IsMissing(pCryRpt) Then
      Set mvarCryRpt = pCryRpt
   End If
   i% = 0
   While mvarCryRpt.Formulas(i%) <> ""
      mvarCryRpt.Formulas(i%) = ""
      i% = i% + 1
   Wend
   Set mvarFormulas = Nothing
End Sub
Private Sub Class_Initialize()
   mvarRel_Mdb$ = "REPORT.MDB"
   mvarRel_Drv$ = "C:\TMP\"
   mvarRpt_Drive = "C:\TMP\"
   mvarRptCadName = "CAD.RPT"
   mvarRptCadMemo = "CADMEMO.RPT"
   mvarDESTINO = ToScreen
   wFactor = (Screen.Width / Screen.TwipsPerPixelX) / 800
   hFactor = (Screen.Height / Screen.TwipsPerPixelY) / 600
   mvarIdioma = 5000 'Português
End Sub
