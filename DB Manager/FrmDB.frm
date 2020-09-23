VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmDB 
   Caption         =   "Database Window"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.TreeView TV 
      Height          =   2535
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   4471
      _Version        =   327682
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "FrmDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Sub Form_Load()
   '
   With Me
      .Top = 0
      .Left = 0
      .Height = MDIManager.ScaleHeight
      .Width = MDIManager.ScaleWidth \ 2
   End With
   '
   With TV
      .Top = 0
      .Left = 0
      .Height = Me.ScaleHeight
      .Width = Me.ScaleWidth
   End With
   '
End Sub

Private Sub Form_Resize()
   '
   With TV
      .Top = 0
      .Left = 0
      .Height = Me.ScaleHeight
      .Width = Me.ScaleWidth
   End With
   '
End Sub

