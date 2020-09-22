VERSION 5.00
Object = "*\ACustomDataControl.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option1 
      Caption         =   "Desc"
      Height          =   255
      Index           =   1
      Left            =   7620
      TabIndex        =   24
      Top             =   390
      Width           =   825
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Asc"
      Height          =   255
      Index           =   0
      Left            =   7620
      TabIndex        =   23
      Top             =   90
      Value           =   -1  'True
      Width           =   825
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Fill with 1000 more items"
      Height          =   435
      Left            =   3060
      TabIndex        =   22
      Top             =   3180
      Width           =   1875
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      LargeChange     =   10
      Left            =   3030
      Max             =   1
      Min             =   1
      TabIndex        =   19
      Top             =   1650
      Value           =   1
      Width           =   2025
   End
   Begin VB.ComboBox CboField 
      Height          =   315
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   4740
      Width           =   855
   End
   Begin CustomDataControl.DataControl DataControl1 
      Height          =   345
      Left            =   2730
      TabIndex        =   14
      Top             =   180
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   609
      EOFAction       =   1
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Filter"
      Height          =   400
      Left            =   5970
      TabIndex        =   13
      Top             =   1620
      Width           =   2000
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   5700
      TabIndex        =   12
      Top             =   1260
      Width           =   2415
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Close database"
      Height          =   400
      Left            =   120
      TabIndex        =   11
      Top             =   2250
      Width           =   2000
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Sort Records"
      Height          =   400
      Left            =   5460
      TabIndex        =   10
      Top             =   150
      Width           =   2000
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   390
      TabIndex        =   9
      Top             =   3450
      Width           =   1305
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Find String"
      Height          =   400
      Left            =   120
      TabIndex        =   8
      Top             =   3810
      Width           =   2000
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Delete record"
      Height          =   400
      Left            =   120
      TabIndex        =   7
      Top             =   1410
      Width           =   2000
   End
   Begin VB.TextBox Textb 
      Height          =   285
      Index           =   2
      Left            =   3000
      TabIndex        =   6
      Top             =   1260
      Width           =   2055
   End
   Begin VB.TextBox Textb 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   5
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox Textb 
      Height          =   285
      Index           =   0
      Left            =   3000
      TabIndex        =   4
      Top             =   660
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add new record"
      Height          =   400
      Left            =   120
      TabIndex        =   3
      Top             =   990
      Width           =   2000
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create new database"
      Height          =   400
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   2000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open database"
      Height          =   400
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save database"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   1830
      Width           =   2000
   End
   Begin VB.Label Label5 
      Caption         =   "Field Count:"
      Height          =   225
      Left            =   3090
      TabIndex        =   21
      Top             =   2520
      Width           =   2085
   End
   Begin VB.Label Label4 
      Caption         =   "Record Count:"
      Height          =   225
      Left            =   3090
      TabIndex        =   20
      Top             =   2250
      Width           =   2085
   End
   Begin VB.Label Label3 
      Caption         =   "Filter to apply (Field index=String) ex: 0=a*,2=200?"
      Height          =   405
      Left            =   5700
      TabIndex        =   18
      Top             =   810
      Width           =   2445
   End
   Begin VB.Label Label2 
      Caption         =   "String to find (you can use ? or *)"
      Height          =   405
      Left            =   360
      TabIndex        =   17
      Top             =   3000
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "In Field (Leave blank for all fields)"
      Height          =   405
      Left            =   120
      TabIndex        =   16
      Top             =   4260
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim toFill As Boolean



Private Sub Command1_Click()
DataControl1.Save
End Sub


Private Sub Command10_Click()
Dim mSort As eSortDirection
If Option1(0).Value = True Then mSort = sortAsc Else mSort = sortDesc

DataControl1.Sort mSort, TextCompare
End Sub

Private Sub Command11_Click()
DataControl1.CloseDatabase
End Sub

Private Sub Command12_Click()
DataControl1.Filter Text2.Text, TextCompare
End Sub

Private Sub Command2_Click()
Dim tStr1 As String
Dim tStr3 As String

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
Dim I As Integer

DataControl1.AddNew
For I = 0 To 2
DataControl1.Field(I) = "TEST " & I
Next I

End Sub


Private Sub Command5_Click()
DataControl1.Delete
End Sub

Private Sub Command6_Click()
DataControl1.Find Text1.Text, , CboField.ListIndex, FindNext, TextCompare, True
End Sub


Private Sub Command8_Click()
On Error Resume Next
Dim I As Long

Combo1.Clear

For I = 1 To DataControl1.RecordCount
Combo1.AddItem UCase(DataControl1.Field(0) & " " & DataControl1.Field(1) & " " & DataControl1.Field(2))

'Combo1.AddItem UCase(DataControl1.Field(1) & " " & DataControl1.Field(2) & " " & DataControl1.Field(3))
DataControl1.MoveNext
Next I
Combo1.ListIndex = 0

MsgBox "Done!!!"
End Sub

Private Sub Command9_Click()
Command2_Click
Command8_Click
End Sub

Private Sub Command7_Click()
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

Private Sub DataControl1_AfterClose()
CboField.Clear
CboField.Clear

HScroll1.Max = 1
Label4.Caption = "Record Count: "
Label5.Caption = "Field Count: "
End Sub

Private Sub DataControl1_AfterFilter()
DataControl1_AfterOpen
End Sub

Private Sub DataControl1_AfterOpen()
On Error Resume Next
Dim I As Integer

CboField.Clear
For I = 0 To DataControl1.FieldCount - 1
CboField.AddItem I
Next I

HScroll1.Max = DataControl1.RecordCount
Label4.Caption = "Record Count: " & DataControl1.RecordCount
Label5.Caption = "Field Count: " & DataControl1.FieldCount
End Sub

Private Sub DataControl1_AfterSave()
MsgBox "Saving Done!!!"
HScroll1.Max = DataControl1.RecordCount
Label4.Caption = "Record Count: " & DataControl1.RecordCount
Label5.Caption = "Field Count: " & DataControl1.FieldCount
End Sub

Private Sub DataControl1_PasswordError()
MsgBox "Wrong password"
End Sub


Private Sub DataControl1_Reposition()
Dim I As Integer

toFill = True
For I = 0 To 2
Textb(I).Text = DataControl1.Field(I) '+1
Next I
toFill = False

DataControl1.Caption = "Record #" & DataControl1.AbsolutePosition
HScroll1.Value = DataControl1.AbsolutePosition
End Sub


Private Sub HScroll1_Change()
DataControl1.AbsolutePosition = HScroll1.Value
End Sub

Private Sub Textb_Change(Index As Integer)
Dim I As Integer

If toFill = True Then Exit Sub

DataControl1.Modify

For I = 0 To 2
DataControl1.Field(I) = Textb(I).Text
Next I

End Sub

