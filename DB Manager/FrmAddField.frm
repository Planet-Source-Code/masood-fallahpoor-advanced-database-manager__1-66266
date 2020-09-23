VERSION 5.00
Begin VB.Form FrmAddField 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Field"
   ClientHeight    =   3090
   ClientLeft      =   2415
   ClientTop       =   1995
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CheckBox ChkRequired 
      Caption         =   "Required"
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   1080
      Width           =   1935
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "FrmAddField.frx":0000
      Left            =   1320
      List            =   "FrmAddField.frx":0028
      TabIndex        =   3
      Top             =   600
      Width           =   1935
   End
   Begin VB.TextBox TxtFieldName 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "Field Type :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Field Name :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   140
      Width           =   975
   End
End
Attribute VB_Name = "FrmAddField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Private Sub Form_Load()
   '
   CmbType.ListIndex = 7
   TxtFieldName.Text = Empty
   ChkRequired.Value = Unchecked
   '
End Sub

Private Sub TxtFieldName_Change()
   '
   If Len(TxtFieldName.Text) <> Empty Then
      cmdOK.Enabled = True
   Else
      cmdOK.Enabled = False
   End If
   '
End Sub

Private Sub cmdOK_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   Dim db As Database
   Dim tb As TableDef
   Dim fl As Field
   '
   If TxtFieldName.Text = Empty Then
      MsgBox "Please Enter Field Name", , "Database Manager"
      Exit Sub
   End If
   '
   If blnDesign = False Then
      strFName(Index) = TxtFieldName.Text
      IntFType(Index) = GetType(CmbType.Text)
      strFType(Index) = CmbType.Text
      blnReq(Index) = ChkRequired.Value
      FrmNewTable.LstFields.AddItem (TxtFieldName)
      Index = Index + 1
   Else
      Set db = DBEngine.OpenDatabase(MDIManager.strFileName)
      Set tb = db.TableDefs(FrmDesignTables.CmbTables.Text)
      Set fl = tb.CreateField(TxtFieldName.Text, GetType(CmbType.Text))
      FrmDesignTables.LstFields.AddItem (TxtFieldName.Text)
      '
      ' Append new field to selected table
      tb.Fields.Append fl
      '
      db.Close
      Set db = Nothing
      Set tb = Nothing
      Set fl = Nothing
   End If
   '
   TxtFieldName.Text = Empty
   CmbType.ListIndex = 9
   ChkRequired.Value = Unchecked
   cmdOK.Enabled = False
   TxtFieldName.SetFocus
   '
   '
   Exit Sub
   '
ErrHandler:
   '
   ' Error handler
   MsgBox "Number : " & Err.Number & vbCrLf & _
          "Description : " & "There isn't Any Table to Append Field", vbInformation, "Database Manager"
   '
End Sub

Private Sub CmdClose_Click()
   '
   Unload FrmAddField
   '
End Sub


Private Function GetType(S As String) As Integer
   '
   Dim IntType As String
   '
   Select Case S
      Case "Boolean"
           IntType = dbBoolean
      Case "Byte"
           IntType = dbByte
      Case "Integer"
           IntType = dbInteger
      Case "Long"
           IntType = dbLong
      Case "Single"
           IntType = dbSingle
      Case "Double"
           IntType = dbDouble
      Case "Currency"
           IntType = dbCurrency
      Case "Date\Time"
           IntType = dbDate
      Case "Text"
           IntType = dbText
      Case "Decimal"
           IntType = dbDecimal
      Case "Binary"
           IntType = dbBinary
      Case "Memo"
           IntType = dbMemo
   End Select
   '
   GetType = IntType
   '
End Function
