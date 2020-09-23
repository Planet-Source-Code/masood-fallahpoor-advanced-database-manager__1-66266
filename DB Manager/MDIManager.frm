VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIManager 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Database Manager"
   ClientHeight    =   5700
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9315
   LinkTopic       =   "Database Manager"
   Moveable        =   0   'False
   NegotiateToolbars=   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   0
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.MDB|*.MDB"
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuFileNew 
         Caption         =   "New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu MnuFileOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu MnuFileClose 
         Caption         =   "Close"
      End
      Begin VB.Menu MnuFileSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu MnuUtility 
      Caption         =   "&Utility"
      Enabled         =   0   'False
      Begin VB.Menu MnuUtilityRepair 
         Caption         =   "Repair Database..."
         Shortcut        =   ^R
      End
      Begin VB.Menu MnuUtilityCompact 
         Caption         =   "Compact Database..."
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuUtilitySeprator 
         Caption         =   "-"
      End
      Begin VB.Menu MnuUtilityQuery 
         Caption         =   "Query Builder..."
         Shortcut        =   ^Q
      End
      Begin VB.Menu MnuUtilityDesign 
         Caption         =   "Design Table..."
      End
      Begin VB.Menu MnuUtilityNewTable 
         Caption         =   "New Table..."
         Shortcut        =   ^T
      End
      Begin VB.Menu MnuUtitilyDeleteTable 
         Caption         =   "Delete Table"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu MnuWin 
      Caption         =   "&Windows"
      Begin VB.Menu MnuWinArrange 
         Caption         =   "Tile Horizontaly"
         Index           =   1
      End
      Begin VB.Menu MnuWinArrange 
         Caption         =   "Tile Verticaly"
         Index           =   2
      End
      Begin VB.Menu MnuWinSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu MnuWinShowDB 
         Caption         =   "Show Database Window"
      End
      Begin VB.Menu MnuWinSQL 
         Caption         =   "Show SQL Statment Window"
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "&About"
      Begin VB.Menu MnuAboutManager 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "MDIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
Public strFileName As String


Private Sub MnuFileNew_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   FrmDB.TV.Nodes.Clear
   MnuUtilityDesign.Enabled = False
   '
   Dim db As Database
   Dim obj As Object
   Dim strProperty As String
   '
   ComDlg.DialogTitle = "Create Database"
   ComDlg.ShowSave
   '
   If ComDlg.FileName <> Empty Then
      strFileName = ComDlg.FileName
      Set db = DBEngine.CreateDatabase(strFileName, dbLangGeneral, dbVersion30)
      FrmDB.TV.Nodes.Add = "Properties"
      '
      ' Call sub GetProperties
      Call GetProperties(db, 1)
      '
      FrmSQL.Show
      FrmDB.Show
      MnuUtility.Enabled = True
      db.Close
      Set db = Nothing
   End If
   '
   Exit Sub
ErrHandler:
   '
   ' Handle error
   MsgBox "Number : " & Err.Number & vbCrLf & "Description : " & _
          Err.Description, vbCritical, "Database Manager"
   '
End Sub

Private Sub MnuFileOpen_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   Dim db As Database
   Dim tb As TableDef
   Dim fl As Field
   Dim CurrNode As Integer
   '
   FrmDB.TV.Nodes.Clear
   '
   ComDlg.DialogTitle = "Open Database"
   ComDlg.ShowOpen
   '
   If ComDlg.FileName <> Empty Then
      strFileName = ComDlg.FileName
      Set db = DBEngine.OpenDatabase(strFileName)
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
            Call GetProperties(tb, CurrNode)
         End If
      Next
      '
      FrmSQL.Show
      FrmDB.Show
      MnuUtility.Enabled = True
      db.Close
      Set db = Nothing
   End If
   '
   Exit Sub
   '
ErrHandler:
   '
   ' Error handler
   MsgBox "Number : " & Err.Number & vbCrLf & "Description : " & _
           Err.Description, vbCritical, "Database Manager"
   '
End Sub

Private Sub MnuFileClose_Click()
   '
   Unload FrmDB
   Unload FrmSQL
   MnuUtility.Enabled = False
   '
End Sub

Private Sub MnuFileExit_Click()
   '
   Dim frm As Form
   '
   ' Unload all forms
   For Each frm In Forms
      Unload frm
   Next
   '
End Sub


Private Sub MnuUtilityCompact_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   Dim OldFileName As String
   Dim NewFileName As String
   Dim StrVersion As String
   Dim IntEncrypt As Integer
   Dim IntVersion As Integer
   '
   ComDlg.DialogTitle = "Select Database to Compact to"
   ComDlg.ShowSave
   '
   If ComDlg.FileName = Empty Then
      Exit Sub
   End If
   '
   OldFileName = strFileName
   NewFileName = ComDlg.FileName
   '
SelectVersion:
   StrVersion = Empty
   StrVersion = InputBox("Select Target Version" & vbCrLf & "1.x,2.x,3.x", _
                       "Select Version", "3.x")
   Select Case LCase(StrVersion)
      Case "1.x": IntVersion = dbVersion11
      Case "2.x": IntVersion = dbVersion20
      Case "3.x": IntVersion = dbVersion30
      Case "": Exit Sub
      Case Else
           MsgBox "Invalid Version", vbCritical, "Error"
           GoTo SelectVersion
   End Select
   '
   IntEncrypt = MsgBox("Encrypt Target?", vbYesNo + vbInformation, "Compact Database")
   '
   If IntEncrypt = vbYes Then
      IntEncrypt = dbEncrypt
   Else
      IntEncrypt = dbDecrypt
   End If
   '
   ' Compact selected database
   Call DBEngine.CompactDatabase(OldFileName, NewFileName, dbLangGeneral, _
                                 IntVersion + IntEncrypt)
   MsgBox "Database Compacted" & vbCrLf & "Location: " & NewFileName, vbInformation, "Database Manager"
   '
   Exit Sub
   '
ErrHandler:
   '
   ' Handle error
   MsgBox "Error Number : " & Err.Number & vbCrLf & _
          "Description : " & Err.Description, vbCritical, "Error"
   '
End Sub

Private Sub MnuUtilityRepair_Click()
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   ' Repair selected database
   DBEngine.RepairDatabase (strFileName)
   MsgBox "Repair Completed" & vbCrLf & "Location: " & strFileName, vbInformation, "Database Manager"
   '
   Exit Sub
   '
ErrHandler:
   '
   'Handle error
   MsgBox "Error Number : " & Err.Number & vbCrLf & _
          "Description : " & Err.Description, vbCritical, "Error"
   '
End Sub

Private Sub MnuUtilityQuery_Click()
   '
   FrmQuery.Show (vbModal)
   '
End Sub

Private Sub MnuUtilityDesign_Click()
   '
   blnDesign = True
   FrmDesignTables.Show (vbModal)
   '
End Sub

Private Sub MnuUtilityNewTable_Click()
   '
   blnDesign = False
   FrmNewTable.Show (vbModal)
   '
End Sub

Private Sub MnuUtitilyDeleteTable_Click()
   '
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   Dim db As Database
   Dim intResponse As Byte
   Dim strTBName As String
   '
   Set db = DBEngine.OpenDatabase(strFileName)
   strTBName = FrmDB.TV.Nodes(FrmDB.TV.SelectedItem.Index).Text
   '
   If strTBName = Empty Or strTBName = "Properties" Then
      MsgBox "Please Select a Table", vbInformation, "Database Manager"
      Exit Sub
   End If
   '
   intResponse = MsgBox("Are You Sure For Delete Selected Table?", vbYesNo + vbInformation, "Database Manager")
   If intResponse = vbNo Then
      Exit Sub
   End If
   '
   ' Delete selected table
   db.TableDefs.Delete strTBName
   '
   ' Delete node from TreeView control
   FrmDB.TV.Nodes.Remove (FrmDB.TV.SelectedItem.Index)
   '
   db.Close
   Set db = Nothing
   '
   Exit Sub
   '
ErrHandler:
   '
   'Handle error
   MsgBox "Error Number : " & Err.Number & vbCrLf & _
          "Description : " & "You Don't Select a Table", vbCritical, "Error"
   '
End Sub

Private Sub MnuWinArrange_Click(Index As Integer)
   '
   MDIManager.Arrange (Index)
   '
End Sub

Private Sub MnuWinShowDB_Click()
   '
   FrmDB.Show
   '
End Sub

Private Sub MnuWinSQL_Click()
   '
   FrmSQL.Show
   '
End Sub

Private Sub MnuAboutManager_Click()
   '
   frmAbout.Show (vbModal)
   '
End Sub
