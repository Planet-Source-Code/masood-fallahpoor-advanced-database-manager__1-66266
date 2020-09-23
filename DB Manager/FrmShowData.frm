VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmShowData 
   ClientHeight    =   4065
   ClientLeft      =   2430
   ClientTop       =   3180
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3690
      Visible         =   0   'False
      Width           =   9765
   End
   Begin MSFlexGridLib.MSFlexGrid FlexGrid 
      Bindings        =   "FrmShowData.frx":0000
      Height          =   1815
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   3201
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "FrmShowData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
   '
   With FlexGrid
      .Top = 0
      .Left = 0
      .Height = Me.ScaleHeight
      .Width = Me.ScaleWidth
   End With
   '
End Sub
