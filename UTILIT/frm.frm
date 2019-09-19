VERSION 5.00
Begin VB.Form Frm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Materiais / Serviços"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   2100
   ClientWidth     =   7170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5760
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   Begin VB.Image ImgFundo 
      BorderStyle     =   1  'Fixed Single
      Height          =   990
      Left            =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Acesso$
