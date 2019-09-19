VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#11.2#0"; "Codejock.Controls.v11.2.2.ocx"
Begin VB.Form MDI 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3120
   ClientLeft      =   17595
   ClientTop       =   3675
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton CmdRestore 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      _Version        =   720898
      _ExtentX        =   2778
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Restore Database"
      UseVisualStyle  =   -1  'True
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdRestore_Click()
   Dim Sql As String

   Sql = "USE [master];"
   Sql = Sql & "RESTORE DATABASE [G3R] FROM  [G3R] WITH  FILE = 1,  NOUNLOAD,  REPLACE,  STATS = 10"
   Sql = Sql & "USE [G3R];"
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'USU_VERIF') DROP USER [USU_VERIF];"
   Sql = Sql & "IF  EXISTS (SELECT * FROM sys.database_principals WHERE name = N'DBA') DROP USER [DBA];"
   Sql = Sql & "CREATE USER [USU_VERIF] FOR LOGIN [USU_VERIF] WITH DEFAULT_SCHEMA=[dbo];"
   Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_datareader', @membername = 'USU_VERIF';"
   Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_datawriter', @membername = 'USU_VERIF';"
   Sql = Sql & "EXEC [sp_addrolemember] @rolename = 'db_backupoperator', @membername = 'USU_VERIF';"
   Sql = Sql & "EXEC sys.sp_addsrvrolemember @loginame = N'USU_VERIF', @rolename = N'sysadmin';"
   Call XDb.executa(Sql)

   
End Sub

Private Sub Form_Activate()
   If Me.Tag = "" Then
      Call Main
      Me.Tag = "1"
   End If
End Sub
