VERSION 5.00
Begin VB.Form FrmQuery 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Query Builder"
   ClientHeight    =   6510
   ClientLeft      =   2415
   ClientTop       =   1995
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOr 
      Caption         =   "Or into Criteria"
      Height          =   375
      Left            =   2760
      TabIndex        =   21
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton CmdAnd 
      Caption         =   "And  into Criteria"
      Height          =   375
      Left            =   480
      TabIndex        =   20
      Top             =   840
      Width           =   1935
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5280
      TabIndex        =   19
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   2400
      TabIndex        =   17
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton CmdRun 
      Caption         =   "Run"
      Height          =   375
      Left            =   960
      TabIndex        =   16
      Top             =   6000
      Width           =   1335
   End
   Begin VB.TextBox TxtStatment 
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   4440
      Width           =   7335
   End
   Begin VB.OptionButton OptDesc 
      Caption         =   "Desc"
      Height          =   255
      Left            =   5880
      TabIndex        =   12
      Top             =   2280
      Width           =   735
   End
   Begin VB.OptionButton OptAsc 
      Caption         =   "Asc"
      Height          =   255
      Left            =   5160
      TabIndex        =   11
      Top             =   2280
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.ComboBox CmbOrderBy 
      Height          =   315
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1800
      Width           =   2175
   End
   Begin VB.ListBox LstFields 
      Height          =   2205
      ItemData        =   "FrmQuery.frx":0000
      Left            =   2640
      List            =   "FrmQuery.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ListBox LstTables 
      Height          =   2205
      ItemData        =   "FrmQuery.frx":0004
      Left            =   360
      List            =   "FrmQuery.frx":0006
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox TxtValue 
      Height          =   300
      Left            =   5040
      MaxLength       =   15
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.ComboBox CmbOperation 
      Height          =   315
      ItemData        =   "FrmQuery.frx":0008
      Left            =   2760
      List            =   "FrmQuery.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.ComboBox CmbFieldsName 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "Criteria :"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Order By :"
      Height          =   255
      Left            =   5040
      TabIndex        =   13
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label5 
      Caption         =   "Fields to Show :"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "Tables :"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Value :"
      Height          =   255
      Left            =   4800
      TabIndex        =   5
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Operation :"
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Fields Name :"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim bln As Boolean
Dim strSQL As String
Dim strFields As String
Dim SQLToSave As String
'
'
Private Sub Form_Load()
   '
   Dim db As Database
   Dim tb As TableDef
   '
   Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
   '
   For Each tb In db.TableDefs
      If Left$(tb.Name, 4) <> "MSys" Then
         LstTables.AddItem (tb.Name)
      End If
   Next
   '
   db.Close
   Set db = Nothing
   '
   CmbOperation.ListIndex = 0
   TxtStatment.Text = Empty
   TxtValue.Text = Empty
   
   '
End Sub

Private Sub LstTables_Click()
   '
   Dim db As Database
   Dim tb As TableDef
   Dim fl As Field
   '
   LstFields.Clear
   CmbFieldsName.Clear
   CmbOrderBy.Clear
   '
   If LstTables.ListIndex <> -1 Then
      Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
      Set tb = db.TableDefs(LstTables.Text)
      For Each fl In tb.Fields
         LstFields.AddItem (tb.Name & "." & fl.Name)
         CmbFieldsName.AddItem (tb.Name & "." & fl.Name)
         CmbOrderBy.AddItem (tb.Name & "." & fl.Name)
      Next
      '
      db.Close
      Set db = Nothing
      Set tb = Nothing
      '
      CmbFieldsName.ListIndex = 0
      CmbOrderBy.ListIndex = 0
   End If
   '
End Sub

Private Sub CmdAnd_Click()
   '
   If bln = False Then
      TxtStatment.Text = CmbFieldsName.Text & CmbOperation.Text & "'" & TxtValue.Text & "'"
      bln = True
   Else
      TxtStatment.Text = TxtStatment.Text & " AND " & CmbFieldsName.Text & CmbOperation.Text & "'" & TxtValue.Text & "'"
   End If
   '
End Sub

Private Sub CmdOr_Click()
   '
   If bln = False Then
      TxtStatment.Text = CmbFieldsName.Text & CmbOperation.Text & "'" & TxtValue.Text & "'"
      bln = True
   Else
      TxtStatment.Text = TxtStatment.Text & " OR " & CmbFieldsName.Text & CmbOperation.Text & "'" & TxtValue.Text & "'"
   End If
   '
End Sub

Private Sub CmdRun_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   Dim I As Integer
   Dim Cnt As Byte
   Dim strOrderBy As String
   '
   Cnt = 0
   strFields = Empty
   '
   If TxtStatment.Text = Empty Then
      MsgBox "Please Enter Criteria", vbInformation, "Database Manager"
      Exit Sub
   End If
   '
   For I = 0 To LstFields.ListCount - 1
      If LstFields.Selected(I) = True Then
         Cnt = Cnt + 1
         If Cnt > 1 Then
            strFields = strFields & "," & LstFields.List(I)
         Else
            strFields = LstFields.List(I)
         End If
      End If
   Next
   '
   If Cnt = 0 Then
      strFields = "*"
   End If
   '
   If OptAsc.Value = True Then
      strOrderBy = " ASC"
   Else
      strOrderBy = " DESC"
   End If
   '
   strSQL = "SELECT " & strFields & " FROM " & LstTables.Text & " WHERE " & TxtStatment & " ORDER BY " & CmbOrderBy.Text & strOrderBy
   '
   With FrmShowData
      .Data1.DatabaseName = MDIManager.strFileName
      .Data1.RecordSource = strSQL
      .Data1.Refresh
      .FlexGrid.Refresh
      .Caption = strSQL
   End With
   '
   FrmShowData.Show (vbModal)
   '
   SQLToSave = strSQL
   strSQL = Empty
   '
   Exit Sub
   '
ErrHandler:
   '
   ' Error handler
   MsgBox "Number : " & Err.Number & vbCrLf & _
          "Description : " & Err.Description, vbCritical, "Database Manager"
   strSQL = Empty
   '
End Sub

Private Sub CmdSave_Click()
   '
   Dim db As Database
   Dim qd As QueryDef
   '
   Dim QueryName As String
   '
   If SQLToSave = Empty Then
      MsgBox "There isn't Any Query", vbInformation, "Database Manager"
      Exit Sub
   End If
   '
   QueryName = InputBox("Enter Query Name : ", "Save Query", "DefultName")
   '
   If QueryName = Empty Then
      Exit Sub
   Else
      Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
      Set qd = db.CreateQueryDef(QueryName, SQLToSave)
      MsgBox "Query Saved", vbInformation, "Database Manager"
      '
      db.Close
      Set db = Nothing
      Set qd = Nothing
   End If
   '
End Sub

Private Sub CmdClear_Click()
   '
   ' Reset TextBoxes and general variables
   TxtStatment.Text = Empty
   TxtValue.Text = Empty
   strSQL = Empty
   SQLToSave = Empty
   strFields = Empty
   bln = False
   '
End Sub

Private Sub CmdClose_Click()
   '
   Unload Me
   '
End Sub

