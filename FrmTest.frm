VERSION 5.00
Object = "*\ACustomDataControl.vbp"
Begin VB.Form FrmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DataControl test"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command8 
      Caption         =   "Add 1000 Items"
      Height          =   285
      Left            =   6360
      TabIndex        =   22
      Top             =   1710
      Width           =   1365
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Modify record"
      Height          =   500
      Left            =   6240
      TabIndex        =   21
      Top             =   630
      Width           =   1500
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Import Database"
      Height          =   500
      Left            =   7770
      TabIndex        =   20
      Top             =   630
      Width           =   1500
   End
   Begin VB.TextBox Textb 
      Height          =   285
      Index           =   4
      Left            =   5040
      TabIndex        =   19
      Top             =   1680
      Width           =   1200
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Add picture"
      Height          =   285
      Left            =   7800
      TabIndex        =   18
      Top             =   1710
      Width           =   1455
   End
   Begin CustomDataControl.DataControl DataControl1 
      Height          =   495
      Left            =   90
      TabIndex        =   17
      Top             =   1140
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   873
      Caption         =   "No database open"
      CaptionMultiLines=   -1  'True
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Find String"
      Height          =   500
      Left            =   3180
      TabIndex        =   16
      Top             =   630
      Width           =   1500
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Filter"
      Height          =   500
      Left            =   1650
      TabIndex        =   15
      Top             =   630
      Width           =   1500
   End
   Begin VB.CommandButton Command14 
      Caption         =   "Change Password"
      Height          =   500
      Left            =   4710
      TabIndex        =   14
      Top             =   630
      Width           =   1500
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Sort Records"
      Height          =   500
      Left            =   120
      TabIndex        =   13
      Top             =   630
      Width           =   1000
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Asc"
      Height          =   255
      Index           =   0
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   630
      Value           =   -1  'True
      Width           =   500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Desc"
      Height          =   255
      Index           =   1
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   870
      Width           =   500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save database"
      Height          =   500
      Left            =   3180
      TabIndex        =   10
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open database"
      Height          =   500
      Left            =   1650
      TabIndex        =   9
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create new database"
      Height          =   500
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add new record"
      Height          =   500
      Left            =   6240
      TabIndex        =   7
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Delete record"
      Height          =   500
      Left            =   7770
      TabIndex        =   6
      Top             =   120
      Width           =   1500
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Close database"
      Height          =   500
      Left            =   4710
      TabIndex        =   5
      Top             =   120
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   6150
      Left            =   120
      ScaleHeight     =   6090
      ScaleWidth      =   9030
      TabIndex        =   4
      Top             =   2070
      Width           =   9090
   End
   Begin VB.TextBox Textb 
      Height          =   285
      Index           =   3
      Left            =   3810
      TabIndex        =   3
      Top             =   1680
      Width           =   1200
   End
   Begin VB.TextBox Textb 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   1200
   End
   Begin VB.TextBox Textb 
      Height          =   285
      Index           =   1
      Left            =   1350
      TabIndex        =   1
      Top             =   1680
      Width           =   1200
   End
   Begin VB.TextBox Textb 
      Height          =   285
      Index           =   2
      Left            =   2580
      TabIndex        =   0
      Top             =   1680
      Width           =   1200
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim toFill As Boolean

Private Sub Command1_Click()
DataControl1.Save
End Sub

Private Sub Command11_Click()
DataControl1.CloseDatabase
End Sub

Private Sub Command13_Click()
DataControl1.Modify
End Sub

Private Sub Command7_Click()
Dim tStr1 As String

tStr1 = OpenDialog(Me, "All Picture files|*.jpg;*.bmp;*.gif;*.ico;*.cur;*.dib;*.wmf;*.emf", "Insert Picture", App.Path)

DataControl1.InsertFile tStr1, 5
Picture1.Picture = LoadPicture(tStr1)

End Sub

Private Sub DataControl1_AfterDelete()
DataControl1_Reposition
End Sub

Private Sub DataControl1_AfterFilter()
MsgBox "filter applied"
DataControl1.Caption = "Database name: " & DataControl1.DatabaseName & vbCrLf & _
"Record #" & DataControl1.AbsolutePosition & " of " & DataControl1.RecordCount & _
" (" & DataControl1.FieldCount & " fields)"
End Sub

Private Sub DataControl1_AfterFind(StringFound As Boolean)
MsgBox "Find completed"
DataControl1_Reposition
End Sub



Private Sub DataControl1_Error(ErrorNumber As Integer, ErrorDescription As String, ErrorProcedure As String)
MsgBox "Error #: " & ErrorNumber & " Description: " & ErrorDescription & " in " & ErrorProcedure & " module"

End Sub


Private Sub DataControl1_Reposition()
On Error Resume Next
Dim I As Integer
Dim tStr1 As String

toFill = True
For I = 0 To 4
Textb(I).Text = DataControl1.Field(I)
Next I
toFill = False

DataControl1.Caption = "Database name: " & DataControl1.DatabaseName & vbCrLf & _
"Record #" & DataControl1.AbsolutePosition & " of " & DataControl1.RecordCount & _
" (" & DataControl1.FieldCount & " fields)"

tStr1 = App.Path & "\Tmp"

Set Picture1.Picture = Nothing
If DataControl1.FieldCount > 5 Then
DataControl1.ExtractFile tStr1, 5
End If

Picture1.Picture = LoadPicture(tStr1)
Picture1.PaintPicture Picture1.Picture, 0, 0, Picture1.Width, Picture1.Height

Kill tStr1
End Sub

Private Sub DataControl1_AfterClose()
Dim I As Integer

For I = 0 To 4
Textb(I).Text = ""
Next I

Set Picture1.Picture = Nothing

DataControl1.Caption = "No database open"
End Sub

Private Sub DataControl1_AfterOpen()
DataControl1.Caption = "Database name: " & DataControl1.DatabaseName & vbCrLf & _
"Record #" & DataControl1.AbsolutePosition & " of " & DataControl1.RecordCount & _
" (" & DataControl1.FieldCount & " fields)"
End Sub

Private Sub DataControl1_AfterSave()
MsgBox "Saving Done!!!"

DataControl1_Reposition
End Sub

Private Sub DataControl1_PasswordChanged()
MsgBox "password changed successfully"
End Sub

Private Sub DataControl1_PasswordError()
MsgBox "Wrong password"
End Sub
Private Sub Command10_Click()
Dim mSort As eSortDirection
If Option1(0).Value = True Then mSort = sortAsc Else mSort = sortDesc

DataControl1.Sort mSort, TextCompare
End Sub

Private Sub Command12_Click()
Dim tStr1 As String

tStr1 = InputBox("Enter the filter to apply" & vbCrLf & _
"Format: Field Index=Pattern (ex:0=*a?,2=111)" & vbCrLf & _
"For more help see ""Like"" operator in VB Help", , "")

DataControl1.Filter tStr1, TextCompare
End Sub

Private Sub Command14_Click()
Dim tStr1 As String
Dim tStr2 As String


tStr1 = InputBox("Enter the old password", , "")
tStr2 = InputBox("Enter the new password", , "")

DataControl1.ChangePassword tStr1, tStr2
End Sub


Private Sub Command2_Click()
Dim tStr1 As String
Dim tStr3 As String

DataControl1.CloseDatabase

tStr1 = OpenDialog(Me, "Database file|*.dat", "Open Database", App.Path)
tStr3 = InputBox("Enter the password of this database (leave blank if not)", , "")

DataControl1.Password = tStr3
DataControl1.DatabaseName = tStr1
DataControl1.OpenDatabase
End Sub

Private Sub Command3_Click()
Dim tStr1 As String
Dim tStr2 As String
Dim tStr3 As String

tStr1 = SaveDialog(Me, "Database file|*.dat", "Create new Database", App.Path, "New Database.dat")
tStr2 = InputBox("How many field you want in this new database ?", , 3)
tStr3 = InputBox("Enter the password of this database (leave blank if you don't want a password)", , "")

DataControl1.Password = tStr3
DataControl1.CreateDatabase tStr1, Int(tStr2)
End Sub

Private Sub Command4_Click()
DataControl1.AddNew
Set Picture1.Picture = Nothing
End Sub

Private Sub Command5_Click()
DataControl1.Delete
End Sub


Private Sub Command6_Click()
Dim tStr1 As String
Dim tStr2 As String

tStr1 = InputBox("Enter the string to find", , "")
tStr2 = InputBox("Enter in which field you want to find-it (leave blank for all)", , "")

DataControl1.Find tStr1, , tStr2, FindFirst, TextCompare, True
End Sub

Private Sub Command8_Click()
Dim I As Integer
Dim J As Integer
Dim tQ As Long

tQ = DataControl1.RecordCount

For I = 0 To 999
DataControl1.AddNew
    For J = 0 To DataControl1.FieldCount - 1
    DataControl1.Field(J) = "Item " & tQ + I + 1 & " - Field " & J + 1
    Next J
DataControl1.Update
Next I
DataControl1.Save
DataControl1.MoveFirst
End Sub

Private Sub Command9_Click()
FrmImport.Show vbModal, Me
End Sub

Private Sub Textb_Change(Index As Integer)
Dim I As Integer

If toFill = True Then Exit Sub

DataControl1.Modify

For I = 0 To 4
DataControl1.Field(I) = Textb(I).Text
Next I

End Sub
