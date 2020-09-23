VERSION 5.00
Begin VB.Form FrmDesignTables 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Table Structure"
   ClientHeight    =   5745
   ClientLeft      =   2205
   ClientTop       =   1605
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdRemove 
      Caption         =   "Remove Field"
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add Field"
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.ListBox LstFields 
      Height          =   2985
      ItemData        =   "FrmDesignTables.frx":0000
      Left            =   480
      List            =   "FrmDesignTables.frx":0002
      TabIndex        =   2
      Top             =   1320
      Width           =   2535
   End
   Begin VB.ComboBox CmbTables 
      Height          =   315
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "Field Type :"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label LblFieldType 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Field Name :"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label LblFieldName 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Fields :"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Tables :"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FrmDesignTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Dim db As Database
Dim tb As TableDef
Dim fl As Field
'
Private Sub Form_Load()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   CmbTables.Clear
   '
   Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
   '
   For Each tb In db.TableDefs
      If Left$(tb.Name, 4) <> "MSys" Then
         CmbTables.AddItem (tb.Name)
      End If
   Next
   '
   db.Close
   Set db = Nothing
   CmbTables.ListIndex = 0
   '
   Exit Sub
   '
ErrHandler:
   '
   ' Error handler
   MsgBox "Number : " & Err.Number & vbCrLf & _
          "Description : " & "There isn't Any Table to Design", vbInformation, "Database Manager"
   '
End Sub

Private Sub CmbTables_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   
   LstFields.Clear
   '
   Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
   Set tb = db.TableDefs(CmbTables.Text)
   '
   For Each fl In tb.Fields
      LstFields.AddItem (fl.Name)
   Next
   '
   db.Close
   Set db = Nothing
   Set tb = Nothing
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

Private Sub LstFields_Click()
   '
   If LstFields.Text <> Empty Then
      Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
      Set tb = db.TableDefs(CmbTables.Text)
      Set fl = tb.Fields(LstFields.Text)
      LblFieldName.Caption = fl.Name
      LblFieldType.Caption = ShowType(fl.Type)
      '
      db.Close
      Set db = Nothing
      Set tb = Nothing
      Set fl = Nothing
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
   ' Handle error
   On Error GoTo ErrHandler
   '
   If LstFields.ListIndex <> -1 Then
      Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
      Set tb = db.TableDefs(CmbTables.Text)
      '
      ' Delete selected filed
      tb.Fields.Delete LstFields.Text
      '
      LstFields.RemoveItem (LstFields.ListIndex)
      '
      db.Close
      Set db = Nothing
      Set tb = Nothing
   End If
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
   Dim CurrNode As Integer
   '
   Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
   '
   FrmDB.TV.Nodes.Clear
   FrmDB.TV.Nodes.Add = "Properties"
   '
   ' Call sub GetProperties
   Call GetProperties(db, 1)
   '
   For Each tb In db.TableDefs
      If Left$(tb.Name, 4) <> "MSys" Then
         FrmDB.TV.Nodes.Add = tb.Name
         CurrNode = FrmDB.TV.Nodes.Count
         For Each fl In tb.Fields
            FrmDB.TV.Nodes.Add CurrNode, tvwChild, , fl.Name
         Next
         FrmDB.TV.Nodes.Add CurrNode, tvwChild, , "Properties"
         CurrNode = FrmDB.TV.Nodes.Count
         '
         ' Call sub GetProperties
         Call GetProperties(tb, CurrNode)
       End If
   Next
   '
   db.Close
   Set db = Nothing
   '
   Unload FrmDesignTables
   '
End Sub

'---------------------------------------------------------------------'
' This function get an integer value and return user-friendly string  '
'---------------------------------------------------------------------'
Private Function ShowType(TypeCode As Integer)
   '
   Dim StrType As String
   '
   Select Case TypeCode
      Case dbBoolean
           StrType = "Boolean"
      Case dbByte
           StrType = "Byte"
      Case dbInteger
           StrType = "Integer"
      Case dbLong
           StrType = "Long"
      Case dbSingle
           StrType = "Single"
      Case dbDouble
           StrType = "Double"
      Case dbCurrency
           StrType = "Currency"
      Case dbDate
           StrType = "Date\Time"
      Case dbText
           StrType = "Text"
      Case dbDecimal
           StrType = "Decimal"
      Case dbBinary
           StrType = "Binary"
      Case dbMemo
           StrType = "Memo"
   End Select
   '
   ShowType = StrType
   '
End Function
