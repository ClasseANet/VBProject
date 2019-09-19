Attribute VB_Name = "DSR"
'----------------------------------------------------------
'                                                         |
' Arquivo de Variáveis, Funções e Procedures Globais      |
' aplicável a  qualquer sistema desenvolvido.             |
'    Aqui também consiste variáveis utilizadas por outros |
' arquivos de Funcões e Procedures Globais como GRID.BAS  |
' SENHA.BAS,BD.BAS...                                     |
'                                                         |
'----------------------------------------------------------
Option Explicit

'****************************************************************************
'**********   Função e variáveis incluídas para utilização geral   **********
'****************************************************************************
Global Handle
    Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As Rect) As Long
    Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
    Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
    Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
    Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
    Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
'Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
'* Menu PopUp
'dll Declare Function TrackPopupMenu Lib "User32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hWnd As Long, lpReserved As Any) As Long
'dll Declare Function GetMenu Lib "User32" (ByVal hWnd As Long) As Long
'dll Declare Function GetSubMenu Lib "User32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'* Form Circular
Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
'* Hidi Task
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal Cx As Long, ByVal Cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_SHOWWINDOW = &H40

Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function FSyncShell Lib "STKIT432.DLL" Alias "SyncShell" (ByVal strCmdLine As String, ByVal intCmdShow As Long) As Long
   
Declare Function VerInstallFile Lib "VERSION.DLL" Alias "VerInstallFileA" (ByVal Flags&, ByVal SrcName$, ByVal DestName$, ByVal SrcDir$, ByVal DestDir$, ByVal CurrDir As Any, ByVal TmpName$, lpTmpFileLen&) As Long
Declare Function GetFileVersionInfoSize Lib "VERSION.DLL" Alias "GetFileVersionInfoSizeA" (ByVal strFileName As String, lVerHandle As Long) As Long
Declare Function GetFileVersionInfo Lib "VERSION.DLL" Alias "GetFileVersionInfoA" (ByVal strFileName As String, ByVal lVerHandle As Long, ByVal lcbSize As Long, lpvData As Byte) As Long
Declare Function VerQueryValue Lib "VERSION.DLL" Alias "VerQueryValueA" (lpvVerData As Byte, ByVal lpszSubBlock As String, lplpBuf As Long, lpcb As Long) As Long
Private Declare Function OSGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'****************************************************************************
'**************************      Tipos de Dados      ************************
'****************************************************************************
'Posição Gráfica do Cursor
Type POINTAPI
  x As Long
  y As Long
End Type
Type Rect
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

'****************************************************************************
'********************       Variáveis de Banco de Dados     *****************
'****************************************************************************

Global MsgTitulo$
Global flag_inicio_senha%
Global flag_inicio%
    '****************************************************************************
'******************       Variáveis de Controle de Grid      ****************
'****************************************************************************
Global MAX_LINHAS_GRID%
Global COL_PROCURA_GRID%

'****************************************************************************
'********************       Variáveis de Senha      *************************
'****************************************************************************
Global ClsUser As New User

Global DPH_dUsuarios_Acessos As Recordset
Global SysMdi As MDIForm
Global DPH_dSistemas_Acessos As Recordset
Global Retornou As Integer

Global USU_Master
Global Senha_Master
Global Nivel_Master


'****************************************************************************
'***********************    Variáveis de Abertura   *************************
'****************************************************************************
Global AppDate$
Global SYS_DataUsuario
Global SYS_FormatoData
Global SYS_Empresa%
Global SYS_Filial%
Global SYS_Local$
Global SYS_Camara%
Global SYS_Nome_Empresa$
Global SYS_Nome_Filial$
Global SYS_Nome_Local$

Global DataInicioSistema
'****************************************************************************
'***********************    Variaveis Padrão     ****************************
'****************************************************************************
Global Sys_DataHoje$
Global Sys_Dt_Mask$
Global Sys_Dt_Mask_Aux$
Global Sys_Sep_Dec$
Global Sys_Sep_Mil$
Global Sys_Sep_Data$
Global Sys_Idioma%

'Global DPH_i%
'Global DPH_msg$
'Global DPH_estilo%
'Global DPH_Territorio$

'diversos
'Global Lov_Cab
'Global GretOk%
'Global Gresp%
'Global Gmsg$
'Global Gestilo%
'Global GWindow$      'Indica a 'Window'chamadora
'Global GObjCursor#   'Indica o Objeto em que o cursor está posicionado
'Global GLastKey%     'Indica a última tecla digitada
'Global ArrScr() As Long 'Array de Posicão de Objetos para um possível Scrool
'Global GComCapa%

'****************************************************************************
'***********************    Constantes Padrão     ***************************
'****************************************************************************
'Identificção do Analista de Sistemas
Global Const ANALISTA$ = "DIO"                 'Código do Analista
Global Const NMANALISTA$ = "DIOGENES S. RAMOS" 'Nome do Analista
Global Const PWDANALISTA$ = "???"              'Senha do Analista
Global Const GRPANALISTA$ = "000"              'Grupo do Analista

Global Const REPORT_MDB = "REPORT.MDB"


'Banco de Dados
Global Const TEMPO_MAXIMO_CONEXAO% = 15

'booleano
Enum RESP
   SIM = 1
   NAO = 2
End Enum
'dias da semana
Enum SEMANA
   DOMINGO = 1
   SEGUNDA
   TERCA
   QUARTA
   QUINTA
   SEXTA
   SABADO
End Enum
'Operação de acesso a tabela
Global gOper%
'Public Enum Qry
'   INCLUSAO = 1
'   ALTERACAO
'   EXCLUSAO
'   LEITURA
'End Enum

'Niveis de Acesso
'Global Const GRPANALISTA$ = "000"
'Global Const MASTER = "001"
'Global Const GERENTE = "002"
'Global Const OPERADOR = "003"
'Global Const VISITA = "004"

'Fomato de Canexão
Global Const ACCESS = ";"
Global Const DBASEIII = "dBase III;"
Global Const DBASEIV = "dBase IV;"

'======================
'retirado de GLOBAL.BAS
'======================

' Show parameters
Global Const SHOWNOACTIVATE = 4


'cores
Global Const CINZA& = &HC0C0C0
Global Const CINZA_ESCURO& = &H808080
Global Const FUNDO& = &H80000008
Global Const AMARELO_CLARO& = &H80FFFF
Global Const TRALHA& = &HE0E0E0
Global Const PRETO& = &H0&
Global Const VERMELHO& = &HFF&
Global Const VERDE& = &HFF00&
Global Const VERDE_ESCURO& = &H8000&
Global Const AMARELO& = &HFFFF&
Global Const AZUL& = &HFF0000
Global Const MAGENTA& = &HFF00FF
Global Const CYAN& = &HFFFF00
Global Const BRANCO& = &HFFFFFF
#If Win16 Then
   '-----------------------------------------------------------
   ' FUNCTION: FSyncShell
   '
   ' Executa um programa externo e espera até seu término
   ' Retorna: True se o programa foi bem sucedido, se não False
   '-----------------------------------------------------------
   '
   Function FSyncShell(ByVal strExeName As String, intCmdShow As Integer) As Integer
      'vbHide              = 0
      'vbNormalFocus       = 1
      'vbMinimizedFocus    = 2
      'vbMaximizedFocus    = 3
      'vbNormalNoFocus     = 4
      'vbMinimizedNoFocus  = 6
      Const HINSTANCE_ERROR% = 32
    
      Dim hInstChild As Integer
      '
      'Inicia o programa e entra em "loop" até seu término
      '
      hInstChild = Shell(strExeName, intCmdShow)
      If hInstChild >= HINSTANCE_ERROR Then
         While GetModuleUsage(hInstChild)
            DoEvents
         Wend
     End If

     FSyncShell = IIf(hInstChild < HINSTANCE_ERROR, False, True)
   End Function
#End If
Public Sub Gradient(TheObject As Object, Redval&, Greenval&, Blueval&, TopToBottom As Boolean)
    'TheObject can be any object that supports the Line method (like forms and pictures).
    'Redval, Greenval, and Blueval are the Red, Green, and Blue starting values from 0 to 255.
    'TopToBottom determines whether the gradient will draw down or up.
    Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$
    'This will create 63 steps in the gradient. This looks smooth on 16-bit and 24-bit color.
    'You can change this, but be careful. You can do some strange-looking stuff with it...
    Step = (TheObject.Height / 63)
    'This tells it whether to start on the top or the bottom and adjusts variables accordingly.
    If TopToBottom = True Then
       FillTop = 0
    Else
      FillTop = TheObject.Height - Step
    End If
    FillLeft = 0
    FillRight = TheObject.Width
    FillBottom = FillTop + Step
    'If you changed the number of steps, change the number of reps to match it.
    'If you don't, the gradient will look all funny.
    For Reps = 1 To 63
        'This draws the colored bar.
        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
        'This decreases the RGB values to darken the color.
        'Lower the value for "squished" gradients. Raise it for incomplete gradients.
        'Also, if you change the number of steps, you will need to change this number.
        Redval = Redval - 4
        Greenval = Greenval - 4
        Blueval = Blueval - 4
        'This prevents the RGB values from becoming negative, which causes a runtime error.
        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0
        'More top or bottom stuff; Moves to next bar.
        If TopToBottom = True Then FillTop = FillBottom Else FillTop = FillTop - Step
        FillBottom = FillTop + Step
    Next
End Sub
Public Function Flood_Atualiza(ByVal n%, Optional ByVal total As Variant)
'================================================================
'= Última Alteração : 13/01/98                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Atualiza o flood de processamento.              **
'**                                                            **
'** Recebe: n%     - Número de tarefa executada              **
'**         Total% - Número Total de tarefa                  **
'**                                                            **
'** Retorna : Flood atualizado                                 **
'**                                                            **
'****************************************************************
    Dim perc%, prop%
    Dim Process As Object
    
    'Set Process = SysMdi.PnlProcessamento
    If IsMissing(total) Then
       total = 100
    End If
    perc% = Process.FloodPercent
    prop% = Int((n% + 1) / total * 100)
    
    If prop% > 100 Then
        Process.FloodPercent = 100
    Else
        If prop% > perc% Then
           Process.FloodPercent = prop%
        End If
    End If
    Flood_Atualiza = n% + 1
End Function


Public Sub Flood_Fim()
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Finaliza o flood de processamento.              **
'**                                                            **
'** Recebe:                                                    **
'**                                                            **
'** Retorna :                                                  **
'**                                                            **
'****************************************************************
    Dim i%
    Dim Process As Object
    
    'Set Process = SysMdi.PnlProcessamento

    For i% = Process.FloodPercent To 100
        Process.FloodPercent = i%
    Next
    
    Process.FloodType = 0
    Process.FloodPercent = 0
End Sub


Public Sub Flood_Interrompe()
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Interrompe flood de processamento.              **
'**                                                            **
'** Recebe:                                                    **
'**                                                            **
'** Retorna :                                                  **
'**                                                            **
'****************************************************************
   Dim Process As Object
    
   'Set Process = SysMdi.PnlProcessamento
   Process.FloodType = 0
   Process.FloodPercent = 0
End Sub



Public Sub Flood_Inicio()
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Inicializa flood de processamento.              **
'**                                                            **
'** Recebe:                                                    **
'**                                                            **
'** Retorna :                                                  **
'**                                                            **
'****************************************************************
    Dim Process As Object
    
    'Set Process = SysMdi.PnlProcessamento
    Process.FloodType = 1
    Process.FloodPercent = 0
End Sub




Public Sub DPH_Init()
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Definir formato de data e número                **
'**                                                            **
'** Recebe:                                                    **
'**                                                            **
'** Retorna: Mensagem com os formatos definidos se o sistema   **
'**          operacional estiver utilizando formatos diferntes **
'**                                                            **
'****************************************************************

'* Função deve ser revisada para uma utilização generalizada

   Dim i%, Txt$, Aux$, aux1$, DT$, tmp_data$, tmp_ano%
   Dim DIA$, Mes$, Ano$
   
   Aux$ = Format$(1000, "#,##0.00")
   Sys_Sep_Dec$ = Mid$(Aux$, 6, 1)
   Sys_Sep_Mil$ = Mid$(Aux$, 2, 1)
    
   Sys_Sep_Data$ = "/"
   Aux$ = CStr(Date)
   For i% = 2 To 5
      aux1$ = Mid$(Aux$, i%, 1)
      'procura o primeiro caracter que não seja um dígito
      If aux1 < "0" Or aux1 > "9" Then
         Sys_Sep_Data$ = aux1$
         Exit For
      End If
   Next
   DIA = Format$(Day(Aux$), "00")
   Mes = Format$(Month(Aux$), "00")
   Select Case Len(Aux$)
      Case 8: Ano = Right$(Year(Now), 2)
      Case 10: Ano = Year(Aux$)
   End Select
   Select Case Aux$
      Case DIA + Sys_Sep_Data$ + Mes + Sys_Sep_Data$ + Ano
         SYS_FormatoData = "DMA"
         Sys_Dt_Mask$ = "dd" + Sys_Sep_Data$ + "mm" + Sys_Sep_Data$ + "yyyy"
         Sys_Dt_Mask_Aux$ = "dd" + Sys_Sep_Data$ + "mm" + Sys_Sep_Data$ + "yy"
      Case Mes + Sys_Sep_Data$ + DIA + Sys_Sep_Data$ + Ano
         SYS_FormatoData = "MDA"
         Sys_Dt_Mask = "mm" + Sys_Sep_Data$ + "dd" + Sys_Sep_Data$ + "yyyy"
         Sys_Dt_Mask_Aux$ = "mm" + Sys_Sep_Data$ + "dd" + Sys_Sep_Data$ + "yy"
      Case Ano + Sys_Sep_Data$ + Mes + Sys_Sep_Data$ + DIA
         SYS_FormatoData = "AMD"
         Sys_Dt_Mask$ = "yyyy" + Sys_Sep_Data$ + "mm" + Sys_Sep_Data$ + "dd"
         Sys_Dt_Mask_Aux$ = "yy" + Sys_Sep_Data$ + "mm" + Sys_Sep_Data$ + "dd"
   End Select
   DT$ = Format$(Now, Sys_Dt_Mask)
   
   'testa formato data/número
   If Sys_Dt_Mask$ <> "dd" + Sys_Sep_Data$ + "mm" + Sys_Sep_Data$ + "yyyy" Then
      Txt$ = LoadMsg(17) + Chr$(10) + Chr$(10)
      Txt$ = Txt$ + LoadMsg(19)
      GoTo Erro_DPH_Init
   End If
   If Not (Sys_Sep_Data$ = "/" And Sys_Sep_Mil$ = "." And Sys_Sep_Dec$ = ",") Then
   ' And GetStringFromIni("Intl", "sShortDate", "win.ini") = UCase(Sys_Dt_Mask$)) Then
   'If Not (Sys_Sep_Data$ = "/" And Sys_Sep_Mil$ = "," And Sys_Sep_Dec$ = ".") Then
      If Not (Sys_Sep_Data$ = "." And Sys_Sep_Mil$ = "," And Sys_Sep_Dec$ = ".") Then
      Txt$ = LoadMsg(18) + Chr$(10) + Chr$(10)
      Txt$ = Txt$ + LoadMsg(19) + Chr$(10)
      Txt$ = Txt$ + LoadMsg(20) + Chr$(10)
'      Txt$ = Txt$ + "Número, utilize 9,999.99" + Chr$(10)
       GoTo Erro_DPH_Init
       End If
   End If
   Call WritePrivateProfileString("INTL", "SSHORTDATE", Sys_Dt_Mask$, "WIN.INI")


    'testa formato ano da data
    tmp_ano% = 0
    tmp_data$ = Trim$(CStr(CVDate(DT$)))
    For i% = Len(tmp_data$) To 1 Step -1
       If Mid$(tmp_data$, i%, 1) = Sys_Sep_Data$ Then
          tmp_ano% = Val(Trim$(Mid$(tmp_data, i% + 1)))
          Exit For
       End If
    Next i%
    
Exit Sub

Erro_DPH_Init:
    Screen.MousePointer = vbDefault
    MsgBox Txt$, vbCritical, SysMdi.Caption
    DoEvents
    End
End Sub
'****************************************************************
'*Author: Carl Slutter
'*
'*Description:
'*The higher the "Movement", the slower the window
'*"explosion".
'*
'*Creation Date: Thursday  23 January 1997  2:27 pm
'*Revision Date: Thursday  23 January 1997  2:27 pm
'*
'*Version Number: 1.00
'****************************************************************

Sub ExplodeForm(f As Form, Movement As Integer)
    Dim myRect As Rect
    Dim formWidth%, formHeight%, i%, x%, y%, Cx%, Cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
    For i = 1 To Movement
        Cx = formWidth * (i / Movement)
        Cy = formHeight * (i / Movement)
        x = myRect.Left + (formWidth - Cx) / 2
        y = myRect.Top + (formHeight - Cy) / 2
        Rectangle TheScreen, x, y, x + Cx, y + Cy
    Next i
    
    x = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
    
End Sub


Public Sub ImplodeForm(f As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
'****************************************************************
'*Author: Carl Slutter
'*
'*Description:
'*The larger the "Movement" value, the slower the "Implosion"
'*
'*Creation Date: Thursday  23 January 1997  2:42 pm
'*Revision Date: Thursday  23 January 1997  2:42 pm
'*
'*Version Number: 1.00
'****************************************************************
    
    Dim myRect As Rect
    Dim formWidth%, formHeight%, i%, x%, y%, Cx%, Cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
        For i = Movement To 1 Step -1
        Cx = formWidth * (i / Movement)
        Cy = formHeight * (i / Movement)
        x = myRect.Left + (formWidth - Cx) / 2
        y = myRect.Top + (formHeight - Cy) / 2
        Rectangle TheScreen, x, y, x + Cx, y + Cy
    Next i
    
    x = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
        
End Sub


