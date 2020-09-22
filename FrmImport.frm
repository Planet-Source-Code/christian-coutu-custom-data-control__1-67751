VERSION 5.00
Object = "*\ACustomDataControl.vbp"
Begin VB.Form FrmImport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Import Database"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5310
   StartUpPosition =   1  'CenterOwner
   Begin CustomDataControl.DataControl DataControl1 
      Height          =   285
      Left            =   2460
      TabIndex        =   5
      Top             =   1590
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   503
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convert database"
      Height          =   500
      Left            =   60
      TabIndex        =   3
      Top             =   1140
      Width           =   1500
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2850
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   150
      Width           =   2235
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Select database"
      Height          =   500
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   1500
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   90
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   690
      Width           =   5025
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   1620
      TabIndex        =   4
      Top             =   1260
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Select Table :"
      Height          =   225
      Left            =   1680
      TabIndex        =   1
      Top             =   210
      Width           =   1005
   End
End
Attribute VB_Name = "FrmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Combo1_Click()
If Combo1.ListIndex > -1 Then
Data1.RecordSource = "Select * From [" & Combo1.Text & "]"
Data1.Refresh
Data1.Recordset.MoveLast
Data1.Recordset.MoveFirst

Data1.Caption = Data1.Recordset.RecordCount & " Records"
End If
End Sub


Private Sub Command1_Click()
On Error Resume Next
Dim tStr1 As String
Dim tStr3 As String
Dim I As Integer
Dim J As Long

tStr1 = SaveDialog(Me, "Database file|*.dat", "Save Database", App.Path, Combo1.Text & ".dat")
tStr3 = InputBox("Enter the password for the imported database (leave blank if you don't want a password)", , "")

DataControl1.CloseDatabase
DataControl1.Password = tStr3
DataControl1.CreateDatabase tStr1, Data1.Recordset.Fields.Count
DoEvents

Data1.Recordset.MoveFirst
DataControl1.DatabaseName = tStr1
DataControl1.OpenDatabase
For J = 1 To Data1.Recordset.RecordCount
DataControl1.AddNew
    For I = 0 To Data1.Recordset.Fields.Count - 1
    DataControl1.Field(I) = Data1.Recordset(I)
    Next I
DataControl1.Update
Data1.Recordset.MoveNext
LblStatus.Caption = J & " of " & Data1.Recordset.RecordCount
DoEvents
Next J
DataControl1.Save
DoEvents
DataControl1.CloseDatabase

End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim tStr1 As String
Dim I As Integer

tStr1 = OpenDialog(Me, "Database file|*.mdb", "Open Database", App.Path)

Data1.DatabaseName = tStr1
Data1.RecordSource = ""
Data1.Refresh

For I = 0 To Data1.Database.TableDefs.Count - 1
    If InStr(1, Data1.Database.TableDefs(I).Name, "MSys") = 0 Then
    Combo1.AddItem Data1.Database.TableDefs(I).Name
    End If
Next I
Combo1.ListIndex = 0
End Sub

Private Sub DataControl1_AfterSave()
MsgBox "database converted successfully!"
End Sub

