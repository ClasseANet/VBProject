VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DSControl"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarwFactor As Single 'local copy
Private mvarhFactor As Single 'local copy
Public Property Let hFactor(ByVal vData As Single)
    mvarhFactor = vData
End Property
Public Property Get hFactor() As Single
    hFactor = mvarhFactor
End Property
Public Property Let wFactor(ByVal vData As Single)
    mvarwFactor = vData
End Property
Public Property Get wFactor() As Single
    wFactor = mvarwFactor
End Property
Public Function AppAtiva(Aplic As Object) As Boolean
   Dim nERRO As Integer, cTxt As String
   
   If Aplic.PrevInstance Then
      On Error Resume Next
      '* Tenta ativar aplicação que já estava rodando (caption de MDIForm)
      AppActivate Aplic.ExeName
      nERRO = Err
      On Error GoTo 0
      '* Testa erro
      '* Aplicação pode estar com o foco em um form não MDIChild
      If nERRO <> 0 Then
         cTxt = "Já existe uma cópia da aplicação rodando."
         cTxt = cTxt & Chr$(10) & "Pressione Ok para fechar esta mensagem"
         cTxt = cTxt & Chr$(10) & "e Alt+Tab para localizar a aplicação."
         Screen.MousePointer = vbDefault
         
         MsgBox cTxt, vbSystemModal + vbExclamation, Aplic.Title
         DoEvents
      Else
         '* Maximiza aplicação que já estava rodando
         SendKeys "% X"
      End If
      '* Encerra aplicação cópia
      AppAtiva = True
   End If
End Function
Public Sub HabilitarBotao(ByRef Objeto As Object, ByVal Bool%, Optional Pct As Variant)
Attribute HabilitarBotao.VB_Description = "Determina a Propriedade Picture e MousePointer do objeto"
   On Error GoTo Fim
   If Not IsMissing(Pct) Then
      Objeto.Caption = ""
      Select Case UCase(TypeName(Pct))
         Case "STRING": Objeto.Picture = LoadPicture(Pct)
         Case "INTEGER": Objeto.Picture = LoadResPicture(Pct, vbResBitmap)
         Case "PICTURE", "LONG": Objeto.Picture = Pct
      End Select
   End If
   Objeto.MousePointer = IIf(Bool, vbDefault, vbNoDrop)
Exit Sub
Fim:
   ClsMsg.ShowError
End Sub
Public Sub HabilitarObj(ByRef Objeto As Object, ByVal Bool%, Optional ByRef Mnu)
Attribute HabilitarObj.VB_Description = "Abilita o objeto tornando as propriedades ""Visible"" e ""Enabled"" verdadeiras."
   Objeto.Enabled = Bool%
   Objeto.Visible = Bool%
   If Not IsMissing(Mnu) Then Mnu.Enabled = Bool
End Sub
'**************
'* Constantes de estilo do botão da barra de utilitários
'* 0 - tbrDefault
'* 1 - tbrCheck
'* 2 - tbrButtonGroup
'* 3 - tbrSeparator
'* 4 - tbrPlaceHolder
'**************
Public Sub HabilitarToolBar(ByRef Objeto As Object, ByRef ImgList As Object, Optional StyleDefault = 0, _
                          Optional Pct As Variant, Optional TipText = "", _
                          Optional isVisible = True, Optional isEnabled = True, Optional pSpace)
   Dim lTam As Long
   Dim Ind As Integer
   Dim i As Integer, lBotao As Variant
   
   On Error Resume Next
   
   If Objeto.Buttons.Count = 0 Then
      Set Objeto.ImageList = ImgList
      If GetTag(Objeto, "FLAT") <> "True" Then
         Call SetTag(Objeto, "FLAT", "True")
         Call Toolbar_MakeFlat(Objeto)
      End If
   End If
  
   If StyleDefault = True Then
      StyleDefault = ButtonStyleConstants.tbrDefault
   End If
   
   Ind = Objeto.Buttons.Count + 1
   If Not IsMissing(pSpace) Then
      For i = 1 To pSpace
         Objeto.Buttons.Add Ind, "'" & Format(Ind, "00") & "'", , StyleDefault, Empty
         Ind = Ind + 1
      Next
   Else
      Objeto.Buttons.Add Ind, "'" & Format(Ind, "00") & "'", , StyleDefault, Empty
      With Objeto.Buttons.Item(Ind)
         If Not IsMissing(Pct) Then
            .Image = ImgList.ListImages(Pct).Key
         End If
         .Style = StyleDefault
         .MixedState = Not (StyleDefault = 0)
         If isVisible Then
            .Visible = True
         Else
            .Image = LoadPicture()
            lTam = Objeto.Buttons.Item(Ind).Width
            .Style = 4 '* tbrPlaceHolder
            .Width = lTam
            .Visible = False
         End If
         .ToolTipText = TipText
         .Enabled = isEnabled
      End With
   End If
End Sub
Public Sub BotaoIcon(frm As Object, Optional Cursor = "PRESS", Optional CursorID = "LUPA", Optional PictureID = "LUPABTN")
Attribute BotaoIcon.VB_Description = "Define ícone do mouse para obejtos com eventos de ""CLICK"" obtendo um padrão Internet"
   Dim i          As Integer
   Dim sPicture   As String
   Dim vPicture   As IPictureDisp
   Dim bAux       As Boolean
   Dim TipoObj    As String
   
   On Error Resume Next
   For i = 0 To frm.Controls.Count - 1
      With frm.Controls(i)
         TipoObj = UCase(TypeName(frm.Controls(i)))
         Select Case TipoObj
            Case "SSCOMMAND", "COMMANDBUTTON", "TOOLBAR"
               .MousePointer = vbCustom
               .MouseIcon = LoadResPicture(Cursor, vbResCursor)
               If TipoObj = "SSCOMMAND" Or TipoObj = "COMMANDBUTTON" Then
                  
                  sPicture = UCase(Mid(.Name, 4))
                  
                  Select Case sPicture
                     Case "LOV":
                        If .Picture = 0 Or .Picture Is Nothing Then
                           If Len(.Caption) <= 1 Then
                              .Caption = ""
                              .Picture = LoadResPicture(PictureID, vbResIcon)
                           End If
                        End If
                        
                     Case Else
                        If .Picture = 0 Or .Picture Is Nothing Then
                           If Len(.Caption) <= 1 Then
                              bAux = .Enabled
                              .Enabled = True
                              Err = 0
                              Set .Picture = LoadResPicture(sPicture, vbResIcon)
                              If Err = 0 Then
                                 .Caption = ""
                                 If Len(.ToolTipText) <= 1 Then
                                    .ToolTipText = sPicture
                                 End If
      
                                 If TipoObj = "COMMANDBUTTON" Then
                                    .Style = 1
                                    If .DownPicture = 0 Or .DownPicture Is Nothing Then
                                       Set .DownPicture = LoadResPicture(sPicture & "_DOWN", vbResIcon)
                                       If .DownPicture = 0 Or .DownPicture Is Nothing Then
                                          Set .DownPicture = LoadResPicture(sPicture, vbResIcon)
                                       End If
                                    End If
                                    If .DisablePicture = 0 Or .DisablePicture Is Nothing Then
                                       Set .DisablePicture = LoadResPicture(sPicture & "_FALSE", vbResIcon)
                                    End If
                                    
                                 End If
                              End If
                              .Enabled = bAux
                           End If
                        End If
                  End Select
               End If
            Case "LABEL"
               If UCase(.Name) = "LBLID" And .MousePointer <> vbCustom Then
               'If UCase(.Name) = "LBLID" Then
                  .MousePointer = vbCustom
                  '.MouseIcon = LoadResPicture(CursorID, vbResCursor)
                  .MouseIcon = LoadResPicture(CursorID, vbResIcon)
                  .ForeColor = 16711680
                  
                  '.ForeColor = 160000
                  '.ForeColor = vbBlack
                  .BackStyle = 0
                  .BorderStyle = 0
               End If
         End Select
      End With
   Next
End Sub
Public Sub CarregarOption(ByRef Dbase As Object, ByRef Opt As Object, Tabela$, Campo$)
Attribute CarregarOption.VB_Description = "Carrega os valores de um Option Button "
   Dim Sql$, i%, Desc$
   If Opt.UBound > 0 Then
      For i = Opt.UBound To 1 Step -1
         Unload Opt(i)
      Next
      Opt(0).Visible = False
   End If
   i% = 0
   Sql$ = "select * from " + Tabela$
   If Not Dbase.AbreTabela(Sql$) Then
      If Dbase.CodeSql <> 0 Then Exit Sub
      Dbase.Dys.MoveFirst
      Opt(i).Visible = True
      Desc$ = UCase(Mid(Dbase.Dys(Campo$), 1, 1)) + _
              LCase(Mid(Dbase.Dys(Campo$), 2))
      Opt(i).Caption = Desc$
      Dbase.Dys.MoveNext
      While Not Dbase.Dys.EOF
         i = i + 1
         Load Opt(i)            ' Cria um novo botão.
         Opt(i).Visible = True  ' Mostra o novo botão.
         Desc$ = UCase(Mid(Dbase.Dys(Campo$), 1, 1)) + _
              LCase(Mid(Dbase.Dys(Campo$), 2))
         Opt(i).Caption = Desc$
         Dbase.Dys.MoveNext
      Wend
   Else
      If Dbase.CodeSql <> 0 Then Exit Sub
      Opt(0).Visible = False
   End If
   i% = IIf(Dbase.RowCount(Dbase.Dys) > Opt.UBound + 1, Dbase.RowCount(Dbase.Dys), Opt.UBound)
   ReDim ArrScr(i%)
   For i = 0 To UBound(ArrScr)
       ArrScr(i) = 360 + (i * 360)
   Next
   On Error Resume Next
   Opt(0).Value = True
   If Err <> 380 And Err <> 0 Then
      ClsMsg.ShowError
   End If
'   If DB.CodeSql = FOUND Then DB.Dys.Close
End Sub
Public Sub CentrarForm(MDI_name As Object, form_name As Object)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Centralizar um Form em outro form base.         **
'**                                                            **
'** Recebe:  MDI_name  - MDIForm base                          **
'**          form_name - Form a ser centralizado               **
'**                                                            **
'** Retorna: Form centralizado                                 **
'**                                                            **
'****************************************************************
   Dim hh!, LL!, tt!, ww!
   Dim Dif%, i%
   On Error Resume Next
   If form_name.WindowState = vbMaximized Then Exit Sub
   ww! = form_name.Width
   hh! = form_name.Height
    
   If form_name.MDIChild And TypeName(MDI_name) <> "Screen" Then
      tt! = (MDI_name.ScaleHeight - hh!) / 2
      LL! = (MDI_name.ScaleWidth - ww!) / 2
   Else
     tt! = (MDI_name.Height - hh!) / 2
     LL! = (MDI_name.Width - ww!) / 2
     If form_name.MDIChild Then
        tt! = tt! - 1000
        LL! = LL! + 60
     End If
   End If
   form_name.Move LL!, tt!
End Sub
Public Sub CentrarObj(ObjMain As Object, ObjChild As Object, Optional Tip_Hor_Ver, Optional PorEscala = False)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Centralizar um Form em outro form base.         **
'**                                                            **
'** Recebe:  MDI_name  - MDIForm base                          **
'**          form_name - Form a ser centralizado               **
'**                                                            **
'** Retorna: Form centralizado                                 **
'**                                                            **
'****************************************************************
   Dim hh!, LL!, tt!, ww!
   On Error Resume Next
   If IsMissing(Tip_Hor_Ver) Then Tip_Hor_Ver = ""
   
   '************
   '* Determina se objetos podem ser centralizados por sua escala
   If PorEscala Then
      ww! = ObjChild.ScaleWidth
      PorEscala = (Err = 0)
   End If
   
   If PorEscala Then
      ww! = ObjChild.ScaleWidth
      hh! = ObjChild.ScaleHeight
      tt! = ObjChild.ScaleTop
      LL! = ObjChild.ScaleLeft
      If Tip_Hor_Ver = "V" Or Tip_Hor_Ver = "" Then tt! = (ObjMain.ScaleHeight - hh!) / 2
      If Tip_Hor_Ver = "H" Or Tip_Hor_Ver = "" Then LL! = (ObjMain.ScaleWidth - ww!) / 2
   Else
      
      ww! = ObjChild.Width
      hh! = ObjChild.Height
      tt! = ObjChild.Top
      LL! = ObjChild.Left
      Err = 0
      If Tip_Hor_Ver = "V" Or Tip_Hor_Ver = "" Then
         tt! = (ObjMain.ScaleHeight - hh!) / 2
         If Err <> 0 Then
            tt! = (ObjMain.Height - hh!) / 2
         End If
      End If
      Err = 0
      If Tip_Hor_Ver = "H" Or Tip_Hor_Ver = "" Then
         LL! = (ObjMain.ScaleWidth - ww!) / 2
         If Err <> 0 Then
            LL! = (ObjMain.Width - ww!) / 2
        End If
      End If
   End If
    
   ObjChild.Move LL!, tt!
End Sub
Public Sub ConfigForm(frm As Object, Optional FrmIcone = "", Optional FundoTela = "FUNDO", Optional PintarTitulo = False, Optional Centrar = True, Optional Pintar = True, Optional BtnIcone = True, Optional Redimensiona = False, Optional AlturaTitulo As Integer = 460)
   On Error Resume Next
   If Redimensiona Then Call ResizeForm(frm)
   'Call ResizeForm(frm)
   If Centrar Then Call CentrarForm(Screen, frm)
   If Pintar And FundoTela <> "" Then
      Call ClsCtrl.PintarFundo(frm, FundoTela, PintarTitulo, , AlturaTitulo:=AlturaTitulo)
   Else
      If PintarTitulo Then
         Call ClsCtrl.PintarTitulo(frm, AlturaTitulo:=AlturaTitulo)
      End If
   End If
   If BtnIcone Then Call BotaoIcon(frm)
   If FrmIcone <> "" Then frm.ICON = FrmIcone
End Sub
Public Function ItemDataIndex(Cmb As Object, Num%)
   Dim i%
   For i = 0 To Cmb.ListCount - 1
      If Cmb.ItemData(i) = Num% Then Exit For
   Next
   ItemDataIndex = IIf(i >= Cmb.ListCount, 0, i)
End Function
Public Function LoadPct(Key, Optional Tipo = vbResBitmap)
   LoadPct = LoadResPicture(Key, Tipo)
End Function
Public Sub SetPicture(Controle As Object, Key$, Optional Tipo = vbResBitmap)
   On Error Resume Next
   Controle.Picture = LoadResPicture(Key, Tipo)
   If Err <> 0 Then
      Err = 0
      Set Controle = LoadResPicture(Key, Tipo)
   End If
End Sub
Public Sub ExtractResData(Id, Tipo, Arquivo As String)
   Dim nInt As Integer
   Dim byteFileBuf() As Byte 'This must be byte rather than String, so no Unicode conversion takes place
   
   On Error GoTo Fim
   
   Call ClsDos.Del(Arquivo)
   
   nInt = FreeFile
   Open Arquivo$ For Binary Access Write As nInt
      byteFileBuf = LoadResData(Id, Tipo)
      Put nInt, , byteFileBuf
   GoTo Saida
Fim:
   Resume
Saida:
    Close nInt
    Err = 0
    Exit Sub
End Sub

Public Function GetParents(ByRef Controle As Variant) As Collection
   Dim i       As Integer
   Dim j       As Integer
   Dim MyColl  As Collection
   
   Set MyColl = New Collection
   i = 1
   While GetTag(Controle, "PARENT" & i) <> ""
      Call SetTag(Controle, "PARENT" & i, "")
      i = i + 1
   Wend
      
   i = 1
   Call SetTag(Controle, "PARENT" & i, Controle.Parent.Name)
   MyColl.Add Controle.Parent.Name, Controle.Parent.Name
   
On Error Resume Next
   For j = 0 To Controle.Parent.Controls.Count - 1
      If Not ClsDsr.InArray(VarType(Controle.Parent.Controls(j)), Array(8)) Then 'And Not ClsDsr.InArray(VarType(Controle), Array(8)) Then '* Tipo String
         If ClsAPI.IsWindowChildOf(Controle.hWnd, Controle.Parent.Controls(j).hWnd) Then
            i = i + 1
            Call SetTag(Controle, "PARENT" & i, Controle.Parent.Controls(j).Name)
            If Not ClsDsr.ExisteItem(MyColl, Controle.Parent.Controls(j).Name) Then
               MyColl.Add Controle.Parent.Controls(j).Name, Controle.Parent.Controls(j).Name
            End If
         End If
      End If
   Next
   Set GetParents = MyColl
End Function
Public Function GetTag(ByRef Controle As Variant, ByVal VarName As String, Optional VarDefault As String) As String
   Dim PosIni As Long, posfim As Long
   Dim StrTAG As String
   Dim i%
   
   On Error GoTo Saida
   
   VarName = "|" & Trim(VarName) & "="
   
   If UCase(TypeName(Controle)) = "STRING" Then
      StrTAG = Controle
   Else
      StrTAG = Controle.Tag
   End If
   
   PosIni = InStr(StrTAG, Trim(VarName))
   If PosIni > 0 Then
      PosIni = PosIni + Len(Trim(VarName))
      posfim = InStr(PosIni, StrTAG$, "|")
      i = 0
      While Mid(StrTAG$, PosIni + i, 1) = "|"
         i = i + 1
      Wend
      If i > 0 Then
         posfim = InStr(PosIni + (i - 1), StrTAG$, "|")
      End If
      posfim = IIf(posfim = 0, Len(StrTAG$), posfim - 1)
      StrTAG$ = Mid(StrTAG$, PosIni, posfim - PosIni + 1)
   Else
      StrTAG$ = ""
   End If
   GetTag = StrTAG$
Saida:
   If StrTAG$ = "" Then
      GetTag = VarDefault
   End If
End Function
Public Function SetTag(ByRef Controle As Variant, ByVal VarName As String, ByVal VarValor As String) As String
   Dim StrTAG As String, StrAux As String
   Dim PosIni As Long, posfim As Long
   
   VarName = "|" & Trim(VarName) & "="
   
   If UCase(TypeName(Controle)) = "STRING" Then
      StrTAG = Controle
   Else
      StrTAG = Controle.Tag
   End If
   
   PosIni = InStr(StrTAG, Trim(VarName))
   If PosIni > 0 Then
      posfim = InStr(PosIni + 1, StrTAG$, "|")
      posfim = IIf(posfim = 0, Len(StrTAG) + 1, posfim)
      StrAux = Mid(StrTAG, 1, PosIni - 1) & Mid(StrTAG, PosIni, Len(VarName)) & Trim(VarValor)
      StrAux = StrAux & Mid(StrTAG, posfim, (Len(StrTAG) - posfim) + 1)
      StrTAG = StrAux
   Else
      If Trim(StrTAG) = "" Then
         StrTAG = VarName & VarValor
      Else
         If UCase(TypeName(Controle)) = "STRING" Then
            StrTAG = Controle & VarName & VarValor
         Else
            StrTAG = Controle.Tag & VarName & VarValor
         End If
      End If
   End If
   If UCase(TypeName(Controle)) = "STRING" Then
      Controle = StrTAG
   Else
      Controle.Tag = StrTAG
   End If
   SetTag = StrTAG
End Function
Public Sub LoadPctMouse(Key, Tipo, ByRef Ctrl As Object)
   Ctrl.MousePointer = vbCustom
   Ctrl.MouseIcon = LoadResPicture(Key, Tipo)
End Sub

Public Function LocalizarCombo(Cmb, chave As String, Optional SetCombo = True, Optional PorItemData As Boolean = False) As Integer
   Dim Ind  As Integer
   Dim i    As Integer
   If PorItemData Then
      Ind = -1
      For i = 0 To Cmb.ListCount - 1
         If Cmb.ItemData(i) = chave Then
            Ind = i
            Exit For
         End If
      Next
   Else
      Ind = SendMessageAny(Cmb.hWnd, CB_FINDSTRING, -1, ByVal chave)
   End If
   If SetCombo Then
      If Cmb.ListCount <> 0 Then
         Cmb.ListIndex = Ind
         'If Cmb.Style <> 2 Then
         '   Cmb = Cmb.List(Ind)
         'End If
      End If
      If Cmb.Style <> 2 And Ind = -1 Then
         Cmb.Text = ""
      End If
   End If
   
   LocalizarCombo = Ind
End Function
Public Function LocalizarImageCombo(Cmb, chave As String, Optional SetCombo = True, Optional PorItemData As Boolean = False) As Integer
   Dim Ind  As Integer
   Dim i    As Integer
   If PorItemData Then
      Ind = -1
      For i = 1 To Cmb.ComboItems.Count
         If Cmb.ComboItems(i).Key = chave Then
            Ind = i
            Exit For
         End If
      Next
   Else
      Ind = -1
      For i = 1 To Cmb.ComboItems.Count
      
         If UCase(Mid(Cmb.ComboItems(i).Text, 1, Len(chave))) = UCase(chave) Then
            Ind = i
            Exit For
         End If
      Next
   End If
   If SetCombo Then
      If Cmb.ComboItems.Count <> 0 And Ind > 0 And Ind <= Cmb.ComboItems.Count Then
         'Cmb.ListIndex = Ind
         Cmb.ComboItems(Ind).Selected = True
         'If Cmb.Style <> 2 Then
         '   Cmb = Cmb.List(Ind)
         'End If
      End If
      If Ind = -1 Then
         Cmb.Text = ""
      End If
   End If
   
   LocalizarImageCombo = Ind
End Function
Public Function ExisteItemLst(List As Object, pKey As String) As Boolean
   Dim xItem As Object
   
   On Error Resume Next
   If TypeName(List) = "ListView" Then
      Set xItem = List.ListItems(pKey)
      ExisteItemLst = (Err = 0)
   Else
   End If
End Function
Public Function LocalizarList(Lst As Object, chave$, Optional SetList = True, Optional ColListView) As Integer
   Dim i    As Integer
   Dim nPos As Integer
   
   If UCase(TypeName(Lst)) = "LISTVIEW" Then
      On Error Resume Next
      LocalizarList = -1
      If Trim(chave$) <> "" Then
         For i = 1 To Lst.ListItems.Count
            nPos = 0
            If IsMissing(ColListView) Then
               nPos = Lst.ListItems(chave$).Index
               If nPos = 0 Then
                  nPos = InStr(Lst.ListItems(i).Text, chave$)
               Else
                  LocalizarList = nPos
                  If SetList Then
                     Lst.ListItems(nPos).Selected = True
                  End If
                  Exit For
               End If
            
            ElseIf Val(ColListView) > 0 Then
               nPos = InStr(Lst.ListItems(i).SubItems(Val(ColListView)), chave$)
               nPos = IIf(nPos = 1, 1, 0)
            End If
            If nPos > 0 Then
               LocalizarList = i
               If SetList Then
                  Lst.ListItems(i).Selected = True
               End If
               Exit For
            End If
         Next
      End If
      If Err <> 0 Then
         LocalizarList = -1
      End If
   Else
      LocalizarList = SendMessageAny(Lst.hWnd, LB_FINDSTRING, -1, ByVal chave$)
      If SetList Then
         If Lst.ListCount <> 0 And LocalizarList >= 0 Then
            Lst.ListIndex = LocalizarList
            Lst.Selected(LocalizarList) = True
         End If
      End If
   End If
End Function

Public Sub MontarCombo(ByRef Combo As Object, Coll As Collection, Optional Propiedade = "") ' As Property)
   Dim n As Variant, Dsc$
   Dim mPointer%, i%
   mPointer% = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
   Combo.Clear
   i = 0
   For Each n In Coll
      Select Case UCase(Propiedade)
         Case "IDNF": Dsc = n.IDNF
         Case "DSCPIECE": Dsc = n.DSCPIECE
         Case "IDPIECE": Dsc = n.IDPIECE
         Case Else: Dsc = n
      End Select
      Combo.AddItem Dsc
'      If Combo.Sorted Then
'         Combo.ItemData(LocalizarCombo(Combo, Dsc)) = i
'      End If
      i = i + 1
   Next
   On Error Resume Next
   Combo.Enabled = True
   Combo.Visible = True
   Combo.ListIndex = 0
   Screen.MousePointer = vbDefault
   On Error GoTo 0
End Sub
Public Sub MontarDbCombo(ByRef Dbase As Object, ByRef Combo As Object, Sql As String, Dsc As String, Optional Id = "", Optional IdView = False, Optional ComboAux, Optional Limpa = True)
'===================================================================
'= Última Alteração : 28/11/97                                     =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)              =
'===================================================================
'*******************************************************************
'**                                                               **
'** OBJETIVO : Montar ComboBox                                    **
'**                                                               **
'** Recebe: Combo    - Nome do ComboBox.                          **
'**         SQL$     - Query que extrairá os dados para o Combo   **
'**         DSC$     - Decrição do dado do Combo                  **
'**         ID       - Código do dado do Combo (Opcional)         **
'**         IDVIEW   - Determina se o ID será visível ou não      **
'**         ComboAux - Se IDVIEW=True deve ser passado o Combo    **
'**                    Auxiliar que conterá os valores de ID      **
'**                                                               **
'** Retorna : Combo Montado                                       **
'**                                                               **
'*******************************************************************
   Dim mPointer%, DyCb As Object
   mPointer% = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   '* Limpa Combo
   If Limpa Then Combo.Clear
   If Trim(Sql) = "" Then
      Sql = Combo.Tag
   Else
      Combo.Tag = Sql$
   End If
   If Dbase.AbreTabela(Sql$, DyCb) Then
      'Carrega o Combo
      DyCb.MoveFirst
      Do While Not DyCb.EOF
         If Id = "" Then
            Combo.AddItem DyCb(Dsc$)
         Else
            If IdView Then
               Combo.AddItem DyCb(Id) & " - " & DyCb(Dsc$)
            Else
               Combo.AddItem DyCb(Dsc$)
               If IsMissing(ComboAux) Then
                  Combo.ItemData(Combo.ListCount - 1) = CInt(DyCb(Id))
               Else
                  ComboAux.AddItem CStr(DyCb(Id))
               End If
            End If
         End If
         DyCb.MoveNext
      Loop
      Combo.Tag = Combo.Tag & "LOAD"
      If Combo.ListCount > 0 Then Combo.ListIndex = 0 ' Combo.List(0)
      Combo.Tag = Mid(Combo.Tag, 1, Len(Combo.Tag) - 4)
      DyCb.Close
   End If
   Combo.Enabled = True
   Combo.Visible = True
   Screen.MousePointer = mPointer%
End Sub
Public Sub MontarDbList(ByRef ObjRecordSet As Object, ByRef Lst As Object, Arr_ID_DSC_PAI, Optional ByVal Cab, Optional ICON = "CLOSED", Optional SMALLICON = "CLOSED")
   Dim LV As Boolean, i%, Ord%
   Dim mItem As New ListItem
   Dim wFactor As Single, hFactor As Single
   
   Dim PosID As Integer
   Dim PosDSC As Integer
   Dim PosPAI As Integer
   
   
   wFactor = (Screen.Width / Screen.TwipsPerPixelX) / 800
   hFactor = (Screen.Height / Screen.TwipsPerPixelY) / 600
   
   On Error Resume Next
   
   PosID = Arr_ID_DSC_PAI(0)
   PosDSC = Arr_ID_DSC_PAI(1)
'   PosPAI = Arr_ID_DSC_PAI(2)
   

   LV = (TypeName(Lst) = "ListView")
   Select Case TypeName(Lst)
      Case "ListView"
         Lst.ListItems.Clear
         Ord = 4
         If IsMissing(Cab) Then
            If Lst.ColumnHeaders.Count = 1 Then
               Lst.ColumnHeaders(1).Width = (Lst.Width - 360) * wFactor
            End If
         Else
            For i = Lst.ColumnHeaders.Count To 1 Step -1
               Lst.ColumnHeaders.Remove i
            Next
            For i = 0 To ((UBound(Cab) + 1) / Ord) - 1
               Lst.ColumnHeaders.Add , Cab((Ord * i) + 1), Cab(Ord * i), Cab((Ord * i) + 2) * 100 * wFactor, Cab((Ord * i) + 3)
            Next
         End If
         
         If ObjRecordSet.RecordCount = 0 Then Exit Sub
         ObjRecordSet.MoveFirst
         While Not ObjRecordSet.EOF
            If IsMissing(Cab) Then
               Set mItem = Lst.ListItems.Add(, "k" & ObjRecordSet(PosID), ObjRecordSet(PosDSC) & "") ', "ICON", "SMALLICON")
            Else
               Set mItem = Lst.ListItems.Add(, "k" & ObjRecordSet(PosID), ObjRecordSet(Lst.ColumnHeaders(1).Key) & "", ICON, SMALLICON)
               For i = 2 To ((UBound(Cab) + 1) / Ord)
                  mItem.SubItems(i - 1) = ObjRecordSet(Lst.ColumnHeaders(i).Key) & ""
               Next
            End If
            ObjRecordSet.MoveNext
         Wend
      Case Else
         Lst.Clear
         Screen.MousePointer = vbDefault
         Lst.Visible = True
         If ObjRecordSet.RecordCount = 0 Then Exit Sub
         ObjRecordSet.MoveFirst
         On Error Resume Next
         While Not ObjRecordSet.EOF
            Lst.AddItem ObjRecordSet(PosDSC) & Space(99 - Len(ObjRecordSet(PosDSC))) & ObjRecordSet(PosID)
'            Lst.ItemData(Lst.NewIndex) = ObjRecordSet(0)
            If Err <> 0 Then
               Err = 0
            End If
            ObjRecordSet.MoveNext
         Wend
   End Select
   Lst.Refresh
   
   If UCase(TypeName(Lst)) = "LISTBOX" Then
      If Lst.ListCount > 1 Then Lst.ListIndex = 1
   Else
      Lst.ListItems(1).Selected = True
   End If
End Sub
Public Sub MontarTreeView(ByRef Dbase As Object, ByRef Tree As Object, Sql$, Root%, Optional IdRoot = "R", Optional DscRoot = "Sistema", Optional LenKEy = 3, Optional Imagem = "", Optional cTAG, Optional Flood = "", Optional LblFlood = "")
   Dim NodX As Object
   Dim Cod$, CodRoot$, Txt$, msg$
   Dim dyTre As Object
   Dim k%, Vet() As String
   Dim NvZero%, Tam%
   Dim Flood_Ini&, Flood_Fim&
   On Error GoTo Fim
   ReDim Vet(0)
   '* Configurações de Flood
   If Flood <> "" Then
      Flood_Ini = Flood.Value
      Flood_Fim = CDbl(Flood.Tag)
      Flood.Enabled = True
      Flood.Visible = True
      If TypeName(LblFlood) = "Label" Then
         LblFlood.Enabled = True
         LblFlood.Caption = Trim(CStr(Flood_Ini)) & " %"
         LblFlood.Visible = True
         LblFlood.Refresh
      End If
   End If
   If Trim(Sql$) = "" Then
      Sql$ = Tree.Tag
   Else
      Tree.Tag = Sql$
   End If
'   If Root Then If DscRoot = "Sistema" Then DscRoot = App.Title
   
   Set NodX = Tree.Nodes.Add(, , IdRoot, DscRoot)     ' Root
   Tree.Nodes(Tree.Nodes.Count).Sorted = Tree.Sorted
 
   If Imagem <> "" Then NodX.Image = Imagem
   
   If Not Dbase.AbreTabela(Sql$, dyTre) Then Exit Sub
   If dyTre.EOF Then Exit Sub
   
   dyTre.MoveFirst
   NvZero = IIf(Root, dyTre(2), 1)
   While Not dyTre.EOF
      Cod$ = "'" + dyTre(0) + "'"
      If dyTre(2) = NvZero Then    'Se foi definido Raíz e Nível
         CodRoot = IdRoot
      Else
         On Error Resume Next
         Tam% = dyTre(LenKEy)
         If Err = 3265 Then Tam% = 2 'Item não encontrado na coleção
         On Error GoTo 0
         On Error GoTo Fim
         CodRoot = "'" & Mid(dyTre(0), 1, IIf(Len(dyTre(0)) <= Tam%, Len(dyTre(0)), Len(dyTre(0)) - Tam%)) & "'"
      End If
      Txt$ = dyTre(1)
      On Error Resume Next
                                         'tvwChild =4
      Set NodX = Tree.Nodes.Add(CodRoot, 4, Cod$, Txt$)
      Tree.Nodes(Tree.Nodes.Count).Sorted = Tree.Sorted
      If Not IsMissing(cTAG) Then
         Tree.Nodes(Tree.Nodes.Count).Tag = Tree.Nodes(Tree.Nodes.Count).Tag & "|" & UCase(cTAG) & "=" & dyTre(cTAG)
         If Err = 3265 Then On Error GoTo 0 'Item not found in this collection.
      End If
      If Err <> 0 Then 'NÃO EXISTE O NÓ PAI ( CODROOT)
         k = k + 1
         ReDim Preserve Vet(3 * k)
         Vet(3 * k) = Cod$
         Vet((3 * k) - 2) = Txt$
         Select Case Err
            Case 35601: msg$ = "**Erro -> Item sem Nível Superior Correspondente"
         End Select
         Vet((3 * k) - 1) = msg$
          On Error GoTo 0
      End If
      On Error GoTo Fim
      '*** Expand tree to see all nodes. ***
      'nodX.EnsureVisible
      If Imagem <> "" Then NodX.Image = Imagem
      dyTre.MoveNext
      If Not dyTre.EOF Then
         If Flood <> "" Then
            Flood.Value = CInt(Flood_Ini + ((dyTre.PercentPosition / 100) * (Flood_Fim - Flood_Ini)))
            If TypeName(LblFlood) = "Label" Then
               If Trim(CStr(Flood.Value)) & " %" <> LblFlood.Caption Then
                  LblFlood.Caption = Trim(CStr(Flood.Value)) & " %"
                  LblFlood.Refresh
               End If
            End If
         End If
      End If
   Wend
   dyTre.Close
   If UBound(Vet) <> 0 Then
      Set NodX = Tree.Nodes.Add(, , "ERRO", "ERROS")
      For k = (LBound(Vet) + 1) To (UBound(Vet) / 3)
'         On Error Resume Next
                                           'tvwChild=4
         Set NodX = Tree.Nodes.Add("ERRO", 4, Vet(3 * k), Vet((3 * k) - 2) & "-" & Vet((3 * k) - 1))
         Tree.Nodes(Tree.Nodes.Count).Sorted = Tree.Sorted
      Next
   End If
   Tree.HideSelection = False
   If Root Then Tree.Nodes(2).EnsureVisible
   Tree.SelectedItem = Tree.Nodes(1)
   Tree.Refresh
Exit Sub
Fim:
   ReDim Vet(0)
   ClsMsg.ShowError
End Sub
Public Sub PintarFundo(Optional frm, Optional FundoTela = "FUNDO", Optional bPintarTitulo = False, Optional FundoTitulo = "FundoTit", Optional AlturaTitulo As Integer = 460)
   Dim i    As Integer
   Dim j    As Integer
   Dim Tam  As Double
   Dim Larg As Single
   Dim ww   As Long
   Dim hh   As Long
   Dim n    As Variant
   Dim Img As IPictureDisp
   
   On Error GoTo Fim
   If IsMissing(frm) Then
      Exit Sub
   End If
   
   If Trim(FundoTela) = "" Then FundoTela = "FUNDO"
   Set Img = LoadResPicture(FundoTela, vbResBitmap)
   
   '*****
   '* Definir Dimensão da Imagem
   '*****
   Tam = Img.Width
   Larg = Img.Height
   Larg = IIf(Larg < Tam, Larg, Tam)
   Tam = IIf(Larg < Tam, Larg, Tam)
   
   '*****
   '* Definir Dimensão do Objeto
   '*****
   On Error Resume Next
   ww = frm.ScaleWidth
   hh = frm.ScaleHeight
   If Err <> 0 Then ww = frm.Width
   If Err <> 0 Then hh = frm.Height
   Err = 0
   '*****
   '* Pintar Objeto
   '*****
   frm.Cls
   For i = 0 To Int(ww / Tam)
      For j = 0 To Int(hh / Tam)
         frm.PaintPicture Img, (i * Tam) - 50, (j * Larg) - 50, Tam, Larg, 0, 0
         frm.PaintPicture Img, (i * Tam), (j * Larg), Tam, Larg, 0, 0
      Next
   Next
   
   If bPintarTitulo Then
      Call PintarTitulo(frm, FundoTitulo, AlturaTitulo)
   End If
   Exit Sub
Fim:
   If Err = 3265 Then 'Resource with identifier 'qq' not found
     If ClsMsg.Idioma = 5000 Then
         Call ClsMsg.ExibirStop("Fundo de Tela """ & FundoTela & """ não existe.", ClsMsg.LoadMsg(1))
     Else
         Call ClsMsg.ExibirStop("Resource with identifier " & FundoTela & " not found.", ClsMsg.LoadMsg(1))
     End If
     Err = 0
   Else
      ClsMsg.ShowError
   End If
End Sub
Public Sub PintarTitulo(Optional frm, Optional FundoTitulo, Optional AlturaTitulo As Integer = 460)
   Dim i    As Integer
   Dim Tam  As Double
   Dim Larg As Single
   Dim ww   As Long
   Dim hh   As Long
   Dim n    As Variant
   Dim Img As IPictureDisp
   
   On Error GoTo Fim
   If IsMissing(frm) Then
      Exit Sub
   End If
   
   If IsMissing(FundoTitulo) Then FundoTitulo = "FundoTit"
   If Trim(FundoTitulo) = "" Then FundoTitulo = "FundoTit"
   
   '*****
   '* Definir Dimensão do Objeto
   '*****
   On Error Resume Next
   ww = frm.ScaleWidth
   hh = frm.ScaleHeight
   If Err <> 0 Then ww = frm.Width
   If Err <> 0 Then hh = frm.Height
   Err = 0
   
   '*****
   '* Redefinir Dimensão da Imagem para o Título
   '*****
   Set Img = LoadResPicture(FundoTitulo, vbResBitmap)
   
   '*****
   '* Definir Dimensão da Imagem
   '*****
   On Error Resume Next
   For Each n In frm.Controls
      If TypeName(n) <> "Menu" Then
         If n.Top < 60 Then
            n.Top = 60
         End If
      End If
   Next
   Tam = Img.Width
   Larg = Img.Height
   Larg = IIf(Larg < Tam, Larg, Tam)
   Tam = IIf(Larg < Tam, Larg, Tam)
   Larg = AlturaTitulo
   Tam = 2790
   For i = 0 To Int(ww / Tam)
      frm.PaintPicture Img, (i * Tam), (0 * Larg), Tam, Larg, 0, 0
   Next
   Exit Sub
Fim:
   If Err = 3265 Then 'Resource with identifier 'qq' not found
     If ClsMsg.Idioma = 5000 Then
         Call ClsMsg.ExibirStop("Imagem do título """ & FundoTitulo & """ não existe.", ClsMsg.LoadMsg(1))
     Else
         Call ClsMsg.ExibirStop("Resource with identifier " & FundoTitulo & " not found.", ClsMsg.LoadMsg(1))
     End If
     Err = 0
   Else
      ClsMsg.ShowError
   End If
End Sub
Public Sub PintarCab(Optional frm, Optional FundoTela = "FUNDO", Optional ObjImg)
   Dim i%, j%, Tam!, Larg!
   Dim ww&, hh&
   
   Dim Img As IPictureDisp
   
   On Error GoTo Fim
   If IsMissing(frm) Then
      Exit Sub
   End If
   
   Set Img = LoadResPicture(FundoTela, vbResBitmap)
   
   '*****
   '* Definir Dimensão da Imagem
   '*****
   Tam = Img.Width
   Larg = Img.Height
   Larg = 1164
   Tam = 1164
   Larg = IIf(Larg < Tam, Larg, Tam)
   Tam = IIf(Larg < Tam, Larg, Tam)
   
   '*****
   '* Definir Dimensão do Objeto
   '*****
   On Error Resume Next
   ww = frm.ScaleWidth
   hh = frm.ScaleHeight
   If Err <> 0 Then ww = frm.Width
   If Err <> 0 Then hh = frm.Height
   Err = 0
   '*****
   '* Pintar Objeto
   '*****
'   Frm.Cls
   j = 0
   For i = 0 To Int(ww / Tam)
      frm.PaintPicture Img, (i * Tam) - 50, (j * Larg) - 50, Tam, Larg, 0, 0
      frm.PaintPicture Img, (i * Tam), (j * Larg), Tam, Larg, 0, 0
   Next
   Exit Sub
Fim:
   If Err = 3265 Then 'Resource with identifier 'qq' not found
     If ClsMsg.Idioma = 5000 Then
         Call ClsMsg.ExibirStop("Fundo de Tela """ & FundoTela & """ não existe.", ClsMsg.LoadMsg(1))
     Else
         Call ClsMsg.ExibirStop("Resource with identifier " & FundoTela & " not found.", ClsMsg.LoadMsg(1))
     End If
     Err = 0
   Else
      ClsMsg.ShowError
   End If
End Sub
Public Sub PrintSetup(ByRef CmDialog As Object)
'Chama a procedure de "Cofigurar Impressão" no "Common Dialog Control".
   On Error Resume Next
   ' Somente a janela de "Configurar Impressão".
   CmDialog.Flags = &H40
   CmDialog.ShowPrinter
End Sub

'Public Sub PopMenu(Obj As Object, FrmMenu As Object, MnuIndex As Integer, x As Single, y As Single)
'=================================================================
'= Última Alteração : 04/03/99                                   =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)            =
'=================================================================
'*****************************************************************
'**                                                             **
'** OBJETIVO : Obter um Menu PopUp retirado de um form qualquer **
'**            ( FrmMenu )  que  contenha  um  menu  de  índice **
'**            especificado ( MnuIndex ).                       **
'**                                                             **
'** Recebe: Obj       - Objeto que comtém o foco.               **
'**         FrmMenu   - Form que comté o Menu requerido.        **
'**         MenuIndex - Indice do Menu no Form Especificado.    **
'**         X         - Posição Horizontal do cursor.           **
'**         Y         - Posição Vertical do Vetor.              **
'**                                                             **
'** Retorna : Menu na posição do mouse.                         **
'**                                                             **
'*****************************************************************
'   Dim IX  As Integer, IY As Integer
'   'hMenu and hSubMenu must be long to run on 32-bit
'   Dim hMenu As Long, hSubMenu As Long, R As Integer
'
'   Dim menRect As Rect
'   Obj.Parent.ScaleMode = vbTwips
'   menRect.Left = 0
'   menRect.Top = 0
'   menRect.Right = Screen.Width / Screen.TwipsPerPixelX
'   menRect.Bottom = Screen.Height / Screen.TwipsPerPixelY
'   IX = (x + Obj.Parent.Left + Obj.Left) \ Screen.TwipsPerPixelX
'   'IY = (y + Top + Obj.Top + Obj.Height) \ Screen.TwipsPerPixelY
'   IY = (y + Obj.Parent.Top + Obj.Top + ((Obj.Height - Obj.Top) / 2)) \ Screen.TwipsPerPixelY
'    hMenu = GetMenu(FrmMenu.hWnd)
'   hSubMenu = GetSubMenu(hMenu, MnuIndex)
'   R = TrackPopupMenu(hSubMenu, 2, IX, IY, 0, FrmMenu.hWnd, menRect)
'End Sub
Public Sub RefreshList(ByRef Dbase As Object, ByRef Lst As Object, ByVal Sql$, Optional ByVal Cab)
   Dim rc As Object
   Dim LV As Boolean, i%, Ord%
   Dim mItem As New ListItem
   Dim wFactor As Single, hFactor As Single
   wFactor = (Screen.Width / Screen.TwipsPerPixelX) / 800
   hFactor = (Screen.Height / Screen.TwipsPerPixelY) / 600
'   Lst.Visible = False
   Call Dbase.AbreTabela(Sql$, rc)
   LV = (TypeName(Lst) = "ListView")
   Select Case TypeName(Lst)
      Case "ListView"
         Lst.ListItems.Clear
         If Not Dbase.CodeSql Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         Ord = 4
         rc.MoveFirst
         While Not rc.EOF
            If IsMissing(Cab) Then
               Set mItem = Lst.ListItems.Add(, rc(0), rc(1)) ', "ICON", "SMALLICON")
            Else
               Set mItem = Lst.ListItems.Add(, rc(0), rc(Lst.ColumnHeaders(1).Key)) ', "ICON", "SMALLICON")
               For i = 2 To Lst.ColumnHeaders.Count
                  mItem.SubItems(i - 1) = rc(Lst.ColumnHeaders(i).Key)
               Next
            End If
            rc.MoveNext
         Wend
      Case Else
         Lst.Clear
         If Not Dbase.CodeSql Then
            Screen.MousePointer = vbDefault
            Lst.Visible = True
            Exit Sub
         End If
         rc.MoveFirst
         While Not rc.EOF
            Lst.AddItem rc(1)
            Lst.ItemData(Lst.NewIndex) = rc(0)
            rc.MoveNext
         Wend
   End Select
   rc.Close
   Lst.Visible = True
   'Lst.Refresh
End Sub

Public Sub ResizeCombo(Combo1 As Object, ItensToDisplay As Integer, Optional ObjTab)
   Dim pt As PointAPI
   Dim rc As RECT
   Dim cWidth As Long
   Dim newHeight As Long
   Dim oldScaleMode As Long
   Dim ItemHeight As Long
   Dim ObjParent As Object
   Dim DivAux As Integer
   
   'Save the current form scalemode, then
   'switch to pixels
   If IsMissing(ObjTab) Then
      Set ObjParent = Combo1.Parent
      oldScaleMode = ObjParent.ScaleMode
      ObjParent.ScaleMode = vbPixels
      DivAux = 1
   Else
      Set ObjParent = ObjTab
      DivAux = Screen.TwipsPerPixelX
   End If
   
   'the width of the combo, used below
   cWidth = Combo1.Width
   
   'get the system height of a single
   'combo box list Item
   ItemHeight = SendMessage(Combo1.hWnd, CB_GETItemHEIGHT, 0, ByVal 0)
   
   'Calculate the new height of the combo box. This
   'is the number of Items times the Item height
   'plus two. The 'plus two' is required to allow
   'the calculations to take into account the size
   'of the edit portion of the combo as it relates
   'to Item height. In other words, even if the
   'combo is only 21 px high (315 twips), if the
   'Item height is 13 px per Item (as it is with
   'small fonts), we need to use two Items to
   'achieve this height.
   newHeight = ItemHeight * (ItensToDisplay + 2)
   
   'Get the co-ordinates of the combo box
   'relative to the screen
   Call GetWindowRect(Combo1.hWnd, rc)
   pt.x = rc.Left
   pt.y = rc.Top
   
   'Then translate into co-ordinates
   'relative to the form.
   Call ScreenToClient(ObjParent.hWnd, pt)
   
   'Using the values returned and set above,
   'call MoveWindow to reposition the combo box
   Call MoveWindow(Combo1.hWnd, pt.x, pt.y, Combo1.Width / DivAux, newHeight, True)
   
   'Its done, so show the new combo height
'   Call SendMessage(Combo1.hWnd, CB_SHOWDROPDOWN, True, ByVal 0)
   
   'restore the original form scalemode
   'before leaving
   If IsMissing(ObjParent) Then
      ObjParent.ScaleMode = oldScaleMode
   End If
End Sub
Public Sub ResizeForm(frm As Object)
   Dim wStart     As Integer
   Dim hStart     As Integer
'   Dim wFactor    As Single
'   Dim hFactor    As Single
   Dim wAux       As Single
   Dim hAux       As Single
   Dim wAnterior  As Single
   Dim hAnterior  As Single
   Dim LoopIndex  As Integer
   Dim Ctrl       As Object
   Dim Ind        As Byte
   Dim n          As Variant
   Dim i          As Integer
   Dim CollTabs   As Collection
   Dim nLeft      As Long
   Dim nTop       As Long
   
   Set CollTabs = New Collection
   
   Const DesignX = 800 '640
   Const designY = 600 '480
   wStart = Screen.Width / Screen.TwipsPerPixelX
   hStart = Screen.Height / Screen.TwipsPerPixelY
   mvarwFactor = wStart / DesignX
   mvarhFactor = hStart / designY
   If wStart = DesignX Then Exit Sub
   On Error Resume Next
   
   frm.Width = frm.Width * wFactor
   frm.Height = frm.Height * hFactor
   
   If frm.Width > Screen.Width Then
      frm.Width = Screen.Width - 120
   End If
   If frm.Height > Screen.Height Then
      frm.Height = frm.Height - 120
   End If
   For LoopIndex = 0 To frm.Controls.Count - 1
      Set Ctrl = frm.Controls(LoopIndex)
      nLeft = Ctrl.Left
      nTop = Ctrl.Top
      
      For Each n In ClsCtrl.GetParents(Ctrl)
         If n <> frm.Name Then
            If TypeName(frm.Controls(n)) = "SSTab" Then
               If ClsDsr.ExisteItem(CollTabs, CStr(n)) Then
                  Call SetTag(frm.Controls(n), "TABINICIAL", frm.Controls(n).Tab)
                  CollTabs.Add CStr(n), CStr(n)
               End If
               If Ctrl.Left < 0 Then
                  For i = 0 To frm.Controls(n).Tabs
                     frm.Controls(n).Tab = i
                     If Ctrl.Left >= 0 Then
                        nLeft = Ctrl.Left
                        nTop = Ctrl.Top
                        Exit For
                     End If
                  Next
               End If
            End If
         End If
      Next
      
      wAnterior = Ctrl.Width
      hAnterior = Ctrl.Height
      'If nLeft > 0 Then
         Ctrl.Left = IIf(nLeft >= 0, nLeft * wFactor, ((75000 + nLeft) * wFactor) - 75000)
      'End If
      Ctrl.Width = Ctrl.Width * wFactor
      Ctrl.Top = IIf(nTop >= 0, nTop * hFactor, nTop)
      Ctrl.FontSize = Ctrl.FontSize * wFactor
      Ctrl.Height = Ctrl.Height * hFactor
      If TypeOf Ctrl Is Line Then
         Ctrl.X1 = Ctrl.X1 * hFactor
         Ctrl.x2 = Ctrl.x2 * hFactor
         Ctrl.Y1 = Ctrl.Y1 * hFactor
         Ctrl.y2 = Ctrl.y2 * hFactor
      ElseIf TypeOf Ctrl Is MSFlexGrid Then 'SpreadSheet Then
         Ctrl.Row = 1
         If nLeft >= 0 Then
            If Ctrl.Width + Ctrl.Left >= frm.Width Then
               Ctrl.Width = frm.Width - Ctrl.Left - 120
            End If
            If Ctrl.Height + Ctrl.Top >= frm.Height Then
               Ctrl.Height = frm.Height - Ctrl.Top - 120
            End If
         End If
         For Ind = 0 To Ctrl.Cols - 1
            Ctrl.Col = Ind
            wAux = Ctrl.ColWidth(Ind) / wAnterior
            'Ctrl.ColWidth(Ind) = Ctrl.ColWidth(Ind) * wFactor '+ 0.02
            Ctrl.ColWidth(Ind) = Ctrl.Width * wAux
         Next
         Ctrl.Col = 1
         For Ind = 0 To Ctrl.Rows - 1
            Ctrl.Row = Ind
            hAux = Ctrl.RowHeight(Ind) / hAnterior
            'Ctrl.RowHeight(Ind) = Ctrl.RowHeight(Ind) * hFactor '- 0.1)
            Ctrl.RowHeight(Ind) = Ctrl.Height * hAux
         Next
         If nLeft >= 0 Then
            If Ctrl.Width + Ctrl.Left > frm.Width Then
               Ctrl.Width = frm.Width - Ctrl.Left - 120
            End If
         End If
      End If
      Set Ctrl = Nothing
   Next
   For Each n In CollTabs
      frm.Controls(n).Tab = GetTag(frm.Controls(n), "TABINICIAL")
   Next
   Set CollTabs = Nothing
End Sub
Public Sub Set_Focus(ByVal Objeto As Object)
'================================================================
'= Última Alteração : 28/11/97                                  =
'= Por : DIOGENES SANTOS RAMOS (ANALISTA DE SISTEMAS)           =
'================================================================
'****************************************************************
'**                                                            **
'** OBJETIVO : Obter Foco para um determinado objeto, se ele   **
'**            estiver inativo se tornará ativo e visível.     **
'**                                                            **
'** Recebe: objeto - Objeto a receber o foco                   **
'**                                                            **
'** Retorna : Objeto focado                                    **
'**                                                            **
'****************************************************************
   On Error GoTo Fim
   DoEvents
   If Objeto.Enabled = True And Objeto.Visible = True Then
      Objeto.SetFocus
   Else
      Call HabilitarObj(Objeto.Parent, True)
      Call HabilitarObj(Objeto, True)
      Objeto.SetFocus
   End If
Exit Sub
Fim:
   ClsMsg.ShowError
End Sub
Public Sub ShowProgress(MDISts As Variant, Optional ByVal Value = 0, Optional MsgFim = False, Optional ByVal Panel = 4)
   Dim Visible As Boolean
   Dim rc As RECT
   Dim Prg As Object
   Dim Sts As Object
   

'   Value = CInt(Value)
   If Value = 0 Then Exit Sub
   If TypeOf MDISts Is MDIForm Then
      Set Prg = MDISts.ProgressBar
      Set Sts = MDISts.StatusBar
   ElseIf TypeOf MDISts Is StatusBar Then
      Set Sts = MDISts
      Set Prg = Sts.Parent.ProgressBar
   End If
'   If Value = Prg.Value Then Exit Sub
   Visible = (Value <= 100)
'   DoEvents
   If Visible Then
      If Value <> 0 And Prg.Visible Then

         Prg.Value = Value
         Sts.Panels(Panel).Alignment = 2
         Sts.Panels(Panel) = CStr(CInt(Prg.Value)) & " %"
         If Value >= 100 Then
            If MsgFim Then Call ClsMsg.ExibirAviso(LoadResString(ClsMsg.Idioma + 34), Prg.Parent.Caption)
            Sts.Panels(Panel) = ""
            SetParent Prg.hWnd, Sts.Parent.hWnd
            Prg.Visible = False
         End If
      Else
         '2 => Panel index (0 based)
         SendMessageAny Sts.hWnd, SB_GETRECT, Panel - 1, rc
         
         With rc
            .Top = .Top * Screen.TwipsPerPixelY
            .Left = .Left * Screen.TwipsPerPixelX
            .Bottom = .Bottom * Screen.TwipsPerPixelY - .Top
            .Right = .Right * Screen.TwipsPerPixelX - .Left
         End With
         With Prg
            SetParent .hWnd, Sts.hWnd
            .Move rc.Left, rc.Top, rc.Right - 540, rc.Bottom
            .Visible = True
            .Value = 0
         End With
      End If
   Else
      Sts.Panels(Panel) = ""
      SetParent Prg.hWnd, Sts.Parent.hWnd
      Prg.Visible = False
   End If
'   Prg.Refresh
End Sub
Public Function ValorOption(Opt As Object)
Attribute ValorOption.VB_Description = "Retorna o Valor Índice do ""OptionButton"" selecionado."
   Dim i%
   i% = 0
   For i = Opt.LBound To Opt.UBound
      If Opt(i).Value Then Exit For
   Next
   ValorOption = IIf(i% > Opt.UBound, Opt.LBound, i%)
End Function
Public Sub ListView_SpecialEffects(LV_hWnd As Long)
   Dim nStyle As Long
   Dim nStyleAux As Long
   
   Call ListView_TrackSelect(LV_hWnd)
   Call ListView_FullRowSelect(LV_hWnd)
   Call ListView_ActiveMove(LV_hWnd)
'   nStyle = LVS_EX_FULLROWSELECT Or LVS_EX_ONECLICKACTIVATE Or LVS_EX_TRACKSELECT
'   Do
'      Call SendMessage(LV_hWnd, LVM_SETEXTENDEDLISTVIEWStyle, 0&, ByVal nStyle)
'      Call ListView_TrackSelect(LV_hWnd)
'      nStyleAux = SendMessage(LV_hWnd, LVM_GETEXTENDEDLISTVIEWStyle, 0&, ByVal 0&)
'   Loop Until (nStyle = nStyleAux)
End Sub

Public Sub ListView_FullRowSelect(LV_hWnd As Long)
   Dim nStyle As Long
   nStyle = SendMessage(LV_hWnd, LVM_GETEXTENDEDLISTVIEWStyle, 0&, ByVal 0&)
   nStyle = nStyle Or LVS_EX_FULLROWSELECT
   
   Call SendMessage(LV_hWnd, LVM_SETEXTENDEDLISTVIEWStyle, 0&, ByVal nStyle)
End Sub

Public Sub ListView_ActiveMove(LV_hWnd As Long)
   Dim nStyle As Long
   nStyle = SendMessage(LV_hWnd, LVM_GETEXTENDEDLISTVIEWStyle, 0&, ByVal 0&)
   nStyle = nStyle Or LVS_EX_ONECLICKACTIVATE
   
   Call SendMessage(LV_hWnd, LVM_SETEXTENDEDLISTVIEWStyle, 0&, ByVal nStyle)
End Sub
Public Sub ListView_TrackSelect(LV_hWnd As Long)
   Dim nStyle As Long
   nStyle = SendMessage(LV_hWnd, LVM_GETEXTENDEDLISTVIEWStyle, 0&, ByVal 0&)
   nStyle = nStyle Or LVS_EX_TRACKSELECT
   
   Call SendMessage(LV_hWnd, LVM_SETEXTENDEDLISTVIEWStyle, 0&, ByVal nStyle)
End Sub
Public Sub ListView_GridLines(LV_hWnd As Long)
   Dim nStyle As Long
   nStyle = SendMessage(LV_hWnd, LVM_GETEXTENDEDLISTVIEWStyle, 0&, ByVal 0&)
   nStyle = nStyle Or LVS_EX_GRIDLINES
   
   Call SendMessage(LV_hWnd, LVM_SETEXTENDEDLISTVIEWStyle, 0&, ByVal nStyle)
End Sub
Public Sub ListView_CheckBoxes(LV_hWnd As Long)
   Dim nStyle As Long
   nStyle = SendMessage(LV_hWnd, LVM_GETEXTENDEDLISTVIEWStyle, 0&, ByVal 0&)
   nStyle = nStyle Or LVS_EX_CHECKBOXES
   
   Call SendMessage(LV_hWnd, LVM_SETEXTENDEDLISTVIEWStyle, 0&, ByVal nStyle)
End Sub
Public Sub ListView_HeaderDragDrop(LV_hWnd As Long)
   Dim nStyle As Long
   nStyle = SendMessage(LV_hWnd, LVM_GETEXTENDEDLISTVIEWStyle, 0&, ByVal 0&)
   nStyle = nStyle Or LVS_EX_HEADERDRAGDROP
   
   Call SendMessage(LV_hWnd, LVM_SETEXTENDEDLISTVIEWStyle, 0&, ByVal nStyle)
End Sub
Public Sub ListView_SetBackGroung(LV_hWnd As Long, Optional Imagem)
   Dim BKIMG As LVBKIMAGE
   With BKIMG
      .uFlags = LVBKIF_SOURCE_URL Or LVBKIF_Style_TILE 'use the file provided in pszImage and tile
      .pszImage = (Imagem) 'the imagefile to display]
      'h&v offsets for the pix (only if tiled not specified)
      '.xOffsetPercent = 0
      '.yOffsetPercent = 0
   End With
   
  'set the text background to none (transparent) to
  'allow image to show through.  A RGB() value can
  'also be passed to colour the text background
  '(to the edge of the control)
   Call SendMessage(LV_hWnd, LVM_SETTEXTBKCOLOR, 0&, ByVal CLR_NONE)
   Call SendMessage(LV_hWnd, LVM_SETBKIMAGE, 0&, BKIMG) 'set the image to the listview
End Sub
Public Sub Toolbar_MakeFlat(ByRef TB_Obj As Object) ' , TB_hWnd As Long)
   Dim Style As Long
   Dim hToolbar As Long
   Dim r As Long
   
   Dim TB_hWnd As Long
   TB_hWnd = TB_Obj.hWnd
  
  'get the handle of the toolbar
   hToolbar = FindWindowEx(TB_hWnd, 0&, "ToolbarWindow32", vbNullString)
   
  'retrieve the toolbar Styles
   Style = SendMessage(hToolbar, TB_GETStyle, 0&, ByVal 0&)
   
  'Set the new Style flag
   If Style And TBStyle_FLAT Then
      Style = Style Xor TBStyle_FLAT
   Else
      Style = Style Or TBStyle_FLAT
   End If
   
  'apply the new Style to the toolbar
   Call SendMessage(hToolbar, TB_SETStyle, 0, ByVal Style)
   TB_Obj.Refresh
   
End Sub


