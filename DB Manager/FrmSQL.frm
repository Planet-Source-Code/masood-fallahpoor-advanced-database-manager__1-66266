VERSION 5.00
Begin VB.Form FrmSQL 
   Caption         =   "SQL Statment"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdExe 
      Caption         =   "Execute"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox TxtStatment 
      Height          =   2415
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   6495
   End
End
Attribute VB_Name = "FrmSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'
Private Sub Form_Load()
   '
   With Me
      .Top = 0
      .Left = MDIManager.ScaleWidth \ 2
      .Width = MDIManager.ScaleWidth \ 2
      .Height = MDIManager.ScaleHeight
   End With

End Sub

Private Sub Form_Resize()
   '
   With TxtStatment
      .Top = CmdSave.Height + 200
      .Left = 0
      .Height = Me.ScaleHeight
      .Width = Me.ScaleWidth
   End With
   '
End Sub

Private Sub TxtStatment_Change()
   '
   If Len(TxtStatment.Text) <> 0 Then
      CmdExe.Enabled = True
      CmdSave.Enabled = True
   Else
      CmdExe.Enabled = False
      CmdSave.Enabled = False
   End If
   '
End Sub

Private Sub CmdClear_Click()
   '
   TxtStatment.Text = Empty
   TxtStatment.SetFocus
   '
End Sub

Private Sub CmdExe_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   With FrmShowData
      .Data1.DatabaseName = MDIManager.strFileName
      .Data1.RecordSource = TxtStatment.Text
      .Data1.Refresh
      .FlexGrid.Refresh
      .Caption = TxtStatment.Text
   End With
   '
   FrmShowData.Show (vbModal)
   '
   Exit Sub
   '
ErrHandler:
   '
   ' Error handler
   MsgBox "Number : " & Err.Number & vbCrLf & _
          "Description : " & Err.Description, vbCritical, "Database Manager"
   '
End Sub

Private Sub CmdSave_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   Dim db As Database
   Dim qd As QueryDef
   Dim QueryName As String
   '
   If TxtStatment.Text = Empty Then
      MsgBox "There isn't Any Query", vbInformation, "Database Manager"
      Exit Sub
   End If
   '
   QueryName = InputBox("Enter Query Name: ", "Save Query", "DefultName")
   '
   If QueryName = Empty Then
      Exit Sub
   End If
   '
   Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
   Set qd = db.CreateQueryDef(QueryName, TxtStatment)
   MsgBox "Query Saved", vbInformation, "Database Manager"
   '
   db.Close
   Set db = Nothing
   Set qd = Nothing
   '
   Exit Sub
   '
ErrHandler:
   '
   ' Error handler
   MsgBox "Number : " & Err.Number & vbCrLf & _
          "Description : " & Err.Description, vbCritical, "Database Manager"
   '
End Sub

