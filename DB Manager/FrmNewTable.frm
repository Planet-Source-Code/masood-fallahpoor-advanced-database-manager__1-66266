VERSION 5.00
Begin VB.Form FrmNewTable 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table Structure"
   ClientHeight    =   5775
   ClientLeft      =   2205
   ClientTop       =   1800
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdBuild 
      Caption         =   "Bulid the Table"
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdRemove 
      Caption         =   "Remove Field"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add Field"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.ListBox LstFields 
      Height          =   3180
      ItemData        =   "FrmNewTable.frx":0000
      Left            =   360
      List            =   "FrmNewTable.frx":0002
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox TxtTBName 
      Height          =   285
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label LblFieldType 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   9
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Type :"
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label LblFieldName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Name :"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Field List :"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Table Name  :"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   135
      Width           =   1095
   End
End
Attribute VB_Name = "FrmNewTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Sub Form_Load()
   '
   ' Module variable
   Index = 1
   '
End Sub

Private Sub LstFields_Click()
   '
   If LstFields.ListCount <> 0 Then
      LblFieldName.Caption = strFName(LstFields.ListIndex + 1)
      LblFieldType.Caption = strFType(LstFields.ListIndex + 1)
   End If
   '
End Sub

Private Sub CmdAdd_Click()
   '
   FrmAddField.Show (vbModal)
   '
End Sub

Private Sub CmdRemove_Click()
   '
   If MDIManager.strFileName <> Empty And LstFields.ListCount <> 0 Then
      strFName(LstFields.ListIndex) = Empty
      IntFType(LstFields.ListIndex) = 0
      strFType(LstFields.ListIndex) = Empty
      blnReq(LstFields.ListIndex) = False
      LstFields.RemoveItem (LstFields.ListIndex)
   End If
   '
End Sub

Private Sub CmdBuild_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   Dim db As Database
   Dim tb As TableDef
   Dim fl As Field
   Dim CurrNode As Integer
   Dim Cnt As Byte
   '
   If TxtTBName.Text = Empty Then
      MsgBox "You Must Enter Table Name", , "DB Builder"
      Exit Sub
   End If
   '
   If LstFields.ListCount = 0 Then
      MsgBox "There isn't Any Field For Create Database", , "Db Builder"
      Exit Sub
   End If
   '
   Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
   Set tb = db.CreateTableDef(Me.TxtTBName.Text)
   '
   For Cnt = 1 To Index - 1
      If strFName(Cnt) <> Empty Then
         ' Create field
         Set fl = tb.CreateField(strFName(Cnt), IntFType(Cnt))
         fl.Required = blnReq(Cnt)
         '
         ' Append field to table
         tb.Fields.Append fl
      End If
   Next Cnt
   '
   ' Append table to database
   db.TableDefs.Append tb
   Set fl = Nothing
   '
   With FrmDB
      .TV.Nodes.Add = TxtTBName.Text
      '
      CurrNode = .TV.Nodes.Count
      For Each fl In tb.Fields
         .TV.Nodes.Add CurrNode, tvwChild, , fl.Name
      Next
      '
      .TV.Nodes.Add CurrNode, tvwChild, , "Properties"
      CurrNode = .TV.Nodes.Count
   End With
   '
   ' Show New Table Properties
   Call GetProperties(tb, CurrNode)
   FrmDB.TV.Refresh
   '
   Erase strFName
   Erase IntFType
   Erase strFType
   Erase blnReq
   db.Close
   Set db = Nothing
   Set tb = Nothing
   '
   MDIManager.MnuUtilityDesign.Enabled = True
   Unload FrmNewTable
   '
   Exit Sub
   '
ErrHandler:
   '
   ' Error handler
   MsgBox "Number : " & Err.Number & vbCrLf & _
          "Description : " & Err.Description, vbInformation, "Database Manager"
   '
End Sub

Private Sub CmdClose_Click()
   '
   Unload FrmNewTable
   '
End Sub
