Attribute VB_Name = "Bl_Bio"
'Global ClFinger   As Object
Global Sys        As Object
Global TlCaBio    As TL_CABio
Public gCODSIS    As String
Public gTBNAME    As String
Public gIDNAME    As String
'Public Sub Main()
'   On Error Resume Next
'   '*********
'   '* Testa se já existe uma cópia da aplicação rodando e define formato Data e número.
'   Dim MyLoad As Object
'   Dim bAtiva As Boolean
'   bAtiva = False
'   Set MyLoad = CriarObjeto("DSACTIVE.DS_LOAD")
'   If Not MyLoad Is Nothing Then
'      MyLoad.Aplic = App
'      If MyLoad.Ativa Then
'         Call ExibirAviso("Já existe uma instância ativa.", "[CABio]")
'         bAtiva = True
'         End
'      End If
'   End If
'   Set MyLoad = Nothing
'
'   Dim sCommand As String
'
'   On Error GoTo TrataErro
'   Screen.MousePointer = vbHourglass
'
'   sCommand = Trim(UCase(Command$))
'   If sCommand = "" Then
'      sCommand = SetTag(sCommand, "EXEPATH", Environ("PROGRAMFILES") & "\ClasseA\Projeto3R\")
'      sCommand = SetTag(sCommand, "CODSIS", "P3R")
'      sCommand = SetTag(sCommand, "TBNAME", "RFUNCIONARIO")
'      sCommand = SetTag(sCommand, "IDNAME", "IDFUNC")
'      sCommand = SetTag(sCommand, "IDNAME", "IDFUNC")
'      sCommand = SetTag(sCommand, "SERVER", ComputerName & "\SQLEXPRESS")
'      sCommand = SetTag(sCommand, "DBNAME", "G3R")
'      sCommand = SetTag(sCommand, "UID", "USU_VERIF")
'      sCommand = SetTag(sCommand, "PWD", "MINOTAURO")
'   Else
'      sCommand = SetTag(sCommand, "EXEPATH", App.Path & "\")
'   End If
'   Set Sys = CriarObjeto("SysA.SetA")
'   If Sys Is Nothing Then
'      ExibirInformacao "Arquivo de configuração(Sys) não carregado."
'      End
'   Else
'      With Sys
'         .EXEPATH = GetTag(sCommand, "EXEPATH", "")
'         .CODSIS = GetTag(sCommand, "CODSIS", "")
'         With .xDb
'            .Server = GetTag(sCommand, "SERVER", "")
'            .dbName = GetTag(sCommand, "DBNAME", "")
'            .UID = GetTag(sCommand, "UID", "")
'            .PWD = GetTag(sCommand, "PWD", "")
'            .SrvConecta
'            If .Conectado Then
'               gCODSIS = Sys.CODSIS
'               gTBNAME = GetTag(sCommand, "TBNAME", "")
'               gIDNAME = GetTag(sCommand, "IDNAME", "")
'            Else
'               End
'            End If
'         End With
'         .LocalReg = .EXEPATH & .CODSIS & ".reg" 'App.Path & "\"
'         'If InStr(App.Path & "\", "C:\Sistemas\") <> 0 Then
'         '   .LocalReg = "C:\Sistemas\Dsr\Projeto3R\"
'         '   .CODSIS = "Projeto3R"
'         'End If
'         If .IDLOJA = 0 Then
'            .IDLOJA = ReadIniFile(.LocalReg, "Config", "LOJAPADRAO", "0")
'            If .IDLOJA = 0 Then
'               .IDLOJA = ReadIniFile(.LocalReg, "Config", "LOJA", "0")
'               If .IDLOJA <> 0 Then
'                  Call WriteIniFile(.LocalReg, "Config", "LOJAPADRAO", .IDLOJA)
'               End If
'            End If
'         End If
'      End With
'   End If
'
'   Set TlCaBio = New TL_CABio
'   Set TlCaBio.Sys = Sys
'   Call TlCaBio.Show
'
'   Exit Sub
'TrataErro:
'   MsgBox Err.Number & " - " & Err.Description
'   End
'End Sub
'
