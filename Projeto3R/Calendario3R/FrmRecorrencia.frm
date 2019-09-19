VERSION 5.00
Begin VB.Form FrmRecorrencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compromisso recorrente"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   ForeColor       =   &H8000000B&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frameRangeOfRecurrence 
      Caption         =   "Intervalo de recorrência"
      Height          =   1575
      Left            =   120
      TabIndex        =   31
      Top             =   3000
      Width           =   7575
      Begin VB.ComboBox ddPatternEndDate 
         Height          =   315
         Left            =   4680
         TabIndex        =   39
         Top             =   1080
         Width           =   1635
      End
      Begin VB.TextBox txtPatternEndAfter 
         Height          =   315
         Left            =   4680
         TabIndex        =   37
         Text            =   "10"
         Top             =   660
         Width           =   675
      End
      Begin VB.OptionButton optPatternEndByDate 
         Caption         =   "&Termina em:"
         Height          =   195
         Left            =   3360
         TabIndex        =   36
         Top             =   1140
         Width           =   1275
      End
      Begin VB.OptionButton optPatternEndAfter 
         Caption         =   "Te&rmina após:"
         Height          =   195
         Left            =   3360
         TabIndex        =   35
         Top             =   720
         Width           =   1395
      End
      Begin VB.OptionButton optPatternNoEnd 
         Caption         =   "Sem data de térmi&no"
         Height          =   255
         Left            =   3360
         TabIndex        =   34
         Top             =   300
         Width           =   1755
      End
      Begin VB.ComboBox ddPatternStartDate 
         Height          =   315
         Left            =   1200
         TabIndex        =   33
         Top             =   360
         Width           =   1875
      End
      Begin VB.Label Label7 
         Caption         =   "ocurrências"
         Height          =   195
         Left            =   5400
         TabIndex        =   38
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Começa &em:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   30
      Top             =   4800
      Width           =   1275
   End
   Begin VB.CommandButton btnCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   29
      Top             =   4800
      Width           =   1275
   End
   Begin VB.CommandButton btnRemoveRecurrence 
      Caption         =   "Remover Recorrência"
      Height          =   375
      Left            =   5760
      TabIndex        =   28
      Top             =   4800
      Width           =   1875
   End
   Begin VB.Frame frameRecurrencePatterm 
      Caption         =   "Padrão de recorrência"
      Height          =   1995
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   7575
      Begin VB.Frame pageYearly 
         Caption         =   "Anual"
         Height          =   1095
         Left            =   1200
         TabIndex        =   52
         Top             =   240
         Width           =   6135
         Begin VB.ComboBox cmbYearlyTheMonth 
            Height          =   315
            Left            =   4200
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   60
            Top             =   650
            Width           =   1815
         End
         Begin VB.ComboBox cmbYearlyDate 
            Height          =   315
            Left            =   3120
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   59
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optYearlyThe 
            Caption         =   "No(a)"
            Height          =   315
            Left            =   120
            TabIndex        =   57
            Top             =   660
            Width           =   795
         End
         Begin VB.OptionButton optYearlyDay 
            Caption         =   "Em:"
            Height          =   315
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Value           =   -1  'True
            Width           =   675
         End
         Begin VB.ComboBox cmbYearlyEveryDate 
            Height          =   315
            Left            =   960
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   240
            Width           =   2175
         End
         Begin VB.ComboBox cmbYearlyTheDay 
            Height          =   315
            Left            =   2160
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   650
            Width           =   1815
         End
         Begin VB.ComboBox cmdYearlyThePartOfWeek 
            Height          =   315
            Left            =   960
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   650
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "de"
            Height          =   255
            Left            =   3960
            TabIndex        =   58
            Top             =   720
            Width           =   255
         End
      End
      Begin VB.Frame pageMonthly 
         Caption         =   "Mensal"
         Height          =   1095
         Left            =   2400
         TabIndex        =   40
         Top             =   240
         Width           =   6015
         Begin VB.TextBox txtMonthlyOfEveryTheMonths 
            Height          =   315
            Left            =   4200
            TabIndex        =   49
            Text            =   "1"
            Top             =   650
            Width           =   615
         End
         Begin VB.ComboBox cmbMonthlyDay 
            Height          =   315
            Left            =   840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   650
            Width           =   1095
         End
         Begin VB.ComboBox cmbMonthlyDayOfWeek 
            Height          =   315
            Left            =   1920
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   47
            Top             =   650
            Width           =   1455
         End
         Begin VB.ComboBox cmbMonthlyDayOfMonth 
            Height          =   315
            Left            =   840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton optMonthlyDay 
            Caption         =   "Di&a"
            Height          =   315
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.TextBox txtMonthlyEveryMonth 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1049
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   2400
            MaxLength       =   3
            TabIndex        =   42
            Text            =   "1"
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optMonthlyThe 
            Caption         =   "N&o(a)"
            Height          =   315
            Left            =   120
            TabIndex        =   41
            Top             =   660
            Width           =   795
         End
         Begin VB.Label Label11 
            Caption         =   "de cada"
            Height          =   255
            Left            =   3480
            TabIndex        =   51
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "mês(meses)"
            Height          =   195
            Left            =   4920
            TabIndex        =   50
            Top             =   720
            Width           =   795
         End
         Begin VB.Label Label9 
            Caption         =   "de cada"
            Height          =   255
            Left            =   1680
            TabIndex        =   45
            Top             =   300
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "mês(meses)"
            Height          =   195
            Left            =   3120
            TabIndex        =   44
            Top             =   300
            Width           =   1155
         End
      End
      Begin VB.Frame pageWeekly 
         Caption         =   "Semanal"
         Height          =   1395
         Left            =   1560
         TabIndex        =   13
         Top             =   240
         Width           =   4935
         Begin VB.CheckBox chkWeeklySunday 
            Caption         =   "Domingo"
            Height          =   255
            Left            =   2400
            TabIndex        =   27
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklySaturday 
            Caption         =   "Sábado"
            Height          =   255
            Left            =   1260
            TabIndex        =   26
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklyFriday 
            Caption         =   "Sexta"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklyThursday 
            Caption         =   "Quinta"
            Height          =   255
            Left            =   3720
            TabIndex        =   24
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklyWednesday 
            Caption         =   "Quarta"
            Height          =   255
            Left            =   2400
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkWeeklyTusday 
            Caption         =   "Terça"
            Height          =   255
            Left            =   1260
            TabIndex        =   22
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox chkWeeklyMonday 
            Caption         =   "Segunda"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.TextBox txtWeeklyNWeeks 
            Height          =   315
            Left            =   1140
            TabIndex        =   19
            Text            =   "1"
            Top             =   180
            Width           =   795
         End
         Begin VB.Label Label6 
            Caption         =   "semana(s) no(a):"
            Height          =   255
            Left            =   1980
            TabIndex        =   20
            Top             =   240
            Width           =   1275
         End
         Begin VB.Label Recurr 
            Caption         =   "A &cada"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame pageDaily 
         Caption         =   "Diário"
         Height          =   1275
         Left            =   1200
         TabIndex        =   12
         Top             =   240
         Width           =   2355
         Begin VB.OptionButton optDailyEveryWorkDay 
            Caption         =   "Todos &os dias da semana"
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   600
            Width           =   2115
         End
         Begin VB.TextBox txtDailyEveryNdays 
            Height          =   285
            Left            =   960
            TabIndex        =   15
            Text            =   "1"
            Top             =   180
            Width           =   615
         End
         Begin VB.OptionButton optDailyEveryNdays 
            Caption         =   "A &cada"
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Top             =   180
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.Label Label5 
            Caption         =   "dia(s)"
            Height          =   195
            Left            =   1680
            TabIndex        =   16
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.OptionButton optRecYearly 
         Caption         =   "Anua&l"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1500
         Width           =   855
      End
      Begin VB.OptionButton optRecMonthly 
         Caption         =   "&Mensal"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1140
         Width           =   915
      End
      Begin VB.OptionButton optRecWeekly 
         Caption         =   "&Semanal"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   780
         Width           =   915
      End
      Begin VB.OptionButton optRecDaily 
         Caption         =   "&Diário"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   855
      End
      Begin VB.Line lnSeparator1 
         BorderColor     =   &H80000003&
         X1              =   1140
         X2              =   1140
         Y1              =   300
         Y2              =   1860
      End
   End
   Begin VB.Frame frameApointmentTime 
      Caption         =   "Hora do Compromisso"
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.ComboBox cmbEventDuration 
         Height          =   315
         Left            =   5160
         TabIndex        =   6
         Text            =   " 60 Minutos"
         Top             =   300
         Width           =   2235
      End
      Begin VB.ComboBox cmbEventStartTime 
         Height          =   315
         Left            =   780
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Duração:"
         Height          =   255
         Left            =   4260
         TabIndex        =   5
         Top             =   360
         Width           =   795
      End
      Begin VB.Label txtEventEndTime 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   3000
         TabIndex        =   4
         Top             =   300
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Fim:"
         Height          =   255
         Left            =   2100
         TabIndex        =   3
         Top             =   360
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Início:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "FrmRecorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Event Load()
Event Unload(Cancel As Integer)
Event BtnOKClick()
Event BtnCancelClick()
Event BtnRemoveRecurrenceClick()
Event CmbEventDurationChange()
Event OptRecDailyClick()
Event OptRecMonthlyClick()
Event OptRecWeeklyClick()
Event OptRecYearlyClick()
Private Sub btnCancel_Click()
   RaiseEvent BtnCancelClick
End Sub
Private Sub btnOK_Click()
   RaiseEvent BtnOKClick
End Sub
Private Sub btnRemoveRecurrence_Click()
   RaiseEvent BtnRemoveRecurrenceClick
End Sub
Private Sub cmbEventDuration_Change()
   RaiseEvent CmbEventDurationChange
End Sub
Private Sub cmbEventDuration_Click()
    RaiseEvent CmbEventDurationChange
End Sub
Private Sub cmbEventStartTime_Change()
'    UpdateEndTimeCombo
End Sub
Private Sub cmbEventStartTime_Click()
   RaiseEvent CmbEventDurationChange
End Sub
Private Sub cmbEventStartTime_LostFocus()
   RaiseEvent CmbEventDurationChange
End Sub
Private Sub Form_Load()
   RaiseEvent Load
End Sub
Private Sub optRecDaily_Click()
   RaiseEvent OptRecDailyClick
End Sub
Private Sub optRecMonthly_Click()
   RaiseEvent OptRecMonthlyClick
End Sub
Private Sub optRecWeekly_Click()
   RaiseEvent OptRecWeeklyClick
End Sub
Private Sub optRecYearly_Click()
   RaiseEvent OptRecYearlyClick
End Sub
