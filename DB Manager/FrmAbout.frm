VERSION 5.00
Begin VB.Form FrmAbout 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Database Manager"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3120
      Width           =   1260
   End
   Begin VB.Line Line3 
      BorderColor     =   &H008080FF&
      BorderWidth     =   3
      X1              =   563.431
      X2              =   563.431
      Y1              =   165.652
      Y2              =   662.609
   End
   Begin VB.Line Line2 
      BorderColor     =   &H008080FF&
      BorderWidth     =   3
      X1              =   563.431
      X2              =   4732.821
      Y1              =   662.609
      Y2              =   662.609
   End
   Begin VB.Line Line4 
      BorderColor     =   &H008080FF&
      BorderWidth     =   3
      X1              =   4732.821
      X2              =   4732.821
      Y1              =   165.652
      Y2              =   662.609
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Height          =   615
      Left            =   480
      TabIndex        =   5
      Top             =   360
      Width           =   120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H008080FF&
      BorderWidth     =   3
      X1              =   563.431
      X2              =   4732.821
      Y1              =   165.652
      Y2              =   165.652
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Date Created :  3/2/2006"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Programed by : Masoud Fallahpour"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1573.697
      Y2              =   1573.697
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1573.697
      Y2              =   1573.697
   End
   Begin VB.Label lblVersion 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Version"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   3885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Database Manager"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF8080&
      Height          =   135
      Left            =   480
      TabIndex        =   6
      Top             =   975
      Width           =   4455
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    '
    With Me
       .Top = 2500
       .Left = (MDIManager.ScaleWidth - .Width) \ 2
    End With
    '
    ' Show program version
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    '
End Sub

Private Sub cmdOK_Click()
  '
  Unload Me
  '
End Sub
