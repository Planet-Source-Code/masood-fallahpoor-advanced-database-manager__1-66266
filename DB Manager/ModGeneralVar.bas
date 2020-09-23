Attribute VB_Name = "ModGeneralVar"
Option Explicit
Option Base 1

Public strFName(20) As String
Public IntFType(20) As Integer
Public strFType(20) As String
Public blnReq(20) As Boolean
Public blnDesign As Boolean
Public Index As Integer


Public Sub GetProperties(DAOObj As Object, N As Integer)
   '
   Dim obj As Object
   Dim strProperty As String
   '
   ' Handle error
   On Error GoTo ErrHandler
   '
   For Each obj In DAOObj.Properties
      strProperty = obj.Name & "="
      '
      If obj.Name = "BookMark" Then
         strProperty = strProperty & "?" ' skip bookmark value
      Else
         strProperty = strProperty & obj.Value
      End If
      '
      ' Add Node to TreeView Control
      FrmDB.TV.Nodes.Add N, tvwChild, , strProperty
   Next
   '
   ' Error handler
ErrHandler:
      strProperty = strProperty & "<err>"
      Resume Next
   '
End Sub

