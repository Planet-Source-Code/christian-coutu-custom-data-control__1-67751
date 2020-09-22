VERSION 5.00
Begin VB.UserControl DataControl 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   ScaleHeight     =   375
   ScaleWidth      =   2190
   ToolboxBitmap   =   "DataControl.ctx":0000
   Begin VB.CommandButton BtnFirst 
      Enabled         =   0   'False
      Height          =   375
      Left            =   0
      Picture         =   "DataControl.ctx":00FA
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton BtnNext 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      Picture         =   "DataControl.ctx":0484
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton BtnLast 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      Picture         =   "DataControl.ctx":080E
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton BtnPrev 
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      Picture         =   "DataControl.ctx":0B98
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "DataControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Dim myDatabaseName As String
Dim myCaption As String
Dim myCapAlign As sTxtPosition
Dim myFieldCount As Integer
Dim myRecordCount As Long
Dim myField() As String
Dim myRecord() As String
Dim myAbsolutePos As Long
Dim myBOFAction As eEndAction
Dim myEOFAction As eEndAction
Dim myPassword As String
Dim myOpenMode As eOpenMode
Dim myMaxRecord As Long
Dim myBOF As Boolean
Dim myEOF As Boolean
Dim myMaxCryptLen As Integer
Dim myCapML As Boolean

Dim bAddNew As Boolean
Dim bModify As Boolean
Dim DbOpen As Boolean

Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Public Enum eFind
    FindFirst = 0
    FindPrevious = 1
    FindNext = 2
    FindLast = 3
End Enum

Public Enum eCompare
    BinaryCompare = 0
    TextCompare = 1
End Enum

Public Enum eEndAction
    StayHere = 0
    MoveOpposite = 1
End Enum

Public Enum eState
    isClosed = 0
    isOpen = 1
    isBusy = 2
End Enum

Public Enum eSortDirection
    sortAsc = 1
    sortDesc = -1
End Enum

Public Enum eOpenMode
    ReadWrite = 0
    ReadOnly = 1
    ReadWriteDenyAddNew = 2
End Enum

Dim FileIdent As String
Dim FieldSep As String
Dim RecordSep As String

Public Event Reposition()
Public Event BeforeSave()
Public Event AfterSave()
Public Event BeforeDelete()
Public Event AfterDelete()
Public Event BeforeClose()
Public Event AfterClose()
Public Event BeforeModify()
Public Event BeforeAddNew()
Public Event BeforeOpen()
Public Event AfterOpen()
Public Event BeforeFind()
Public Event AfterFind(StringFound As Boolean)
Public Event BeforeFilter()
Public Event AfterFilter()
Public Event StateChanged(DbState As eState)
Public Event Error(ErrorNumber As Integer, ErrorDescription As String, ErrorProcedure As String)
Public Event PasswordError()
Public Event BeforeChangePassword()
Public Event PasswordChanged()
Public Event MaxRecordReached()
Public Event Progress(Position As Long)

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Const DT_BOTTOM = &H8
Private Const DT_CALCRECT = &H400
Private Const DT_CENTER = &H1
Private Const DT_EXPANDTABS = &H40
Private Const DT_EXTERNALLEADING = &H200
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TABSTOP = &H80
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DT_EDITCONTROL As Long = &H2000

Public Enum sTxtPosition
    TopLeft = 0
    TopCenter = 1
    TopRight = 2
    MiddleLeft = 3
    MiddleCenter = 4
    MiddleRight = 5
    BottomLeft = 6
    BottomCenter = 7
    BottomRight = 8
End Enum

Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lByteLen As Long)
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lByteLen As Long)

Private Const NEG1 = -1&, n0 = 0&, n1 = 1&, n2 = 2&, n3 = 3&, n4 = 4&, n5 = 5&
Private Const n6 = 6&, n7 = 7&, n8 = 8&, n12 = 12&, n16 = 16&, n32 = 32&
Private Sub DrawTxt(ObjHdc As Long, oText As String, oLeft As Long, oRight As Long, oTop As Long, oBottom As Long, mPosition As sTxtPosition, Optional MultiLine As Boolean = False, Optional WordWrap As Boolean = False)
Dim Rct As RECT
Dim tFormat As Long

Rct.Left = oLeft
Rct.Right = oRight
Rct.Top = oTop
Rct.Bottom = oBottom

Select Case mPosition
    Case TopLeft
    tFormat = DT_TOP + DT_LEFT
    Case TopCenter
    tFormat = DT_TOP + DT_CENTER
    Case TopRight
    tFormat = DT_TOP + DT_RIGHT
    Case MiddleLeft
    tFormat = DT_VCENTER + DT_LEFT
    Case MiddleCenter
    tFormat = DT_VCENTER + DT_CENTER
    Case MiddleRight
    tFormat = DT_VCENTER + DT_RIGHT
    Case BottomLeft
    tFormat = DT_BOTTOM + DT_LEFT
    Case BottomCenter
    tFormat = DT_BOTTOM + DT_CENTER
    Case BottomRight
    tFormat = DT_BOTTOM + DT_RIGHT
End Select

If MultiLine = False Then tFormat = tFormat + DT_SINGLELINE

If WordWrap = True And MultiLine = True Then tFormat = tFormat + DT_WORDBREAK

tFormat = tFormat + DT_NOCLIP + DT_NOPREFIX

DrawText ObjHdc, oText, Len(oText), Rct, tFormat

End Sub

Private Sub BtnFirst_Click()
MoveFirst
End Sub

Private Sub BtnLast_Click()
MoveLast
End Sub


Private Sub BtnNext_Click()
MoveNext
End Sub


Private Sub BtnPrev_Click()
MovePrevious
End Sub


Private Sub UserControl_Initialize()
FileIdent = "DataControlFile"
FieldSep = Chr(0) & Chr(&HFE) & Chr(&HFE) & Chr(&HFE) & Chr(&HFE) & Chr(0)
RecordSep = Chr(0) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(&HFF) & Chr(0)
End Sub

Private Sub UserControl_InitProperties()
myCaption = Ambient.DisplayName
myCapAlign = MiddleCenter
myMaxCryptLen = 1000
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    myCaption = .ReadProperty("Caption", Ambient.DisplayName)
    myCapAlign = .ReadProperty("Alignment", 4)
    myDatabaseName = .ReadProperty("DatabaseName", "")
    myBOFAction = .ReadProperty("BOFAction", 0)
    myEOFAction = .ReadProperty("EOFAction", 0)
    myPassword = .ReadProperty("Password", "")
    myOpenMode = .ReadProperty("OpenMode", 0)
    myMaxRecord = .ReadProperty("MaxRecord", 0)
    myMaxCryptLen = .ReadProperty("MaxCryptLen", 1000)
    myCapML = .ReadProperty("CaptionMultiLines", False)
End With
End Sub


Private Sub UserControl_Resize()
BtnFirst.Move 0, 0, 250, Height - 55
BtnPrev.Move 250, 0, 250, Height - 55
BtnLast.Move Width - 250 - 55, 0, 250, Height - 55
BtnNext.Move Width - 500 - 55, 0, 250, Height - 55

UserControl.Cls

DrawTxt hDc, myCaption, ScaleX(500, vbTwips, vbPixels), ScaleX(Width - 500, vbTwips, vbPixels), _
0, ScaleY(Height - 55, vbTwips, vbPixels), myCapAlign, myCapML, True

End Sub


Public Property Get Caption() As String
Caption = myCaption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
myCaption = vNewCaption
UserControl_Resize
PropertyChanged "Caption"
End Property

Public Property Get Alignment() As sTxtPosition
Alignment = myCapAlign
End Property

Public Property Let Alignment(ByVal vNewCapAlign As sTxtPosition)
myCapAlign = vNewCapAlign
UserControl_Resize
PropertyChanged "Alignment"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Caption", myCaption, Ambient.DisplayName
    .WriteProperty "Alignment", myCapAlign, 4
    .WriteProperty "DatabaseName", myDatabaseName, ""
    .WriteProperty "BOFAction", myBOFAction, 0
    .WriteProperty "EOFAction", myEOFAction, 0
    .WriteProperty "Password", myPassword, ""
    .WriteProperty "OpenMode", myOpenMode, 0
    .WriteProperty "MaxRecord", myMaxRecord, 0
    .WriteProperty "MaxCryptLen", myMaxCryptLen, 1000
    .WriteProperty "CaptionMultiLines", myCapML, False
End With
End Sub

Public Sub CreateDatabase(DataFileName As String, DataFieldCount As Integer)
Dim FF As Integer
Dim mData As String
Dim mFieldQty As String

On Error GoTo CreateDatabase_Error

If Len(Dir(DataFileName)) Then Kill DataFileName
DoEvents

If DataFieldCount > 0 And DataFieldCount <= 255 Then
mFieldQty = Chr(DataFieldCount)
Else
Exit Sub
End If

mData = FileIdent & mFieldQty & Crypt("DB") & String(14, Chr(0))

FF = FreeFile
Open DataFileName For Binary Access Write As #FF
    Put #FF, , mData
Close #FF

On Error GoTo 0
Exit Sub

CreateDatabase_Error:

RaiseEvent Error(Err.Number, Err.Description, "CreateDatabase")
End Sub

Public Sub OpenDatabase()
Dim FF As Long
Dim tmpFile As String
Dim bContent() As Byte
Dim FileLenght As Long
Dim Result As Long
Dim sDB As String

On Error GoTo OpenDatabase_Error

bAddNew = False
bModify = False

RaiseEvent BeforeOpen

FF = FreeFile

tmpFile = Space(FileLen(myDatabaseName))
Open myDatabaseName For Binary As #FF
Get #FF, , tmpFile
Close #FF

If Left(tmpFile, Len(FileIdent)) = FileIdent Then
myFieldCount = Asc(Mid(tmpFile, Len(FileIdent) + 1, 1))
ReDim myField(myFieldCount)
sDB = Mid(tmpFile, Len(FileIdent) + 2, 2)

If Crypt(sDB) <> "DB" Then
RaiseEvent PasswordError
Exit Sub
End If


myRecord = Split(tmpFile, RecordSep)

myRecordCount = UBound(myRecord)
DbOpen = True
RaiseEvent StateChanged(isOpen)
RaiseEvent AfterOpen
End If

If myRecordCount > 0 Then
MoveFirst
BtnEnabled True
End If

On Error GoTo 0
Exit Sub

OpenDatabase_Error:

RaiseEvent Error(Err.Number, Err.Description, "OpenDatabase")
End Sub

Public Sub Save(Optional noEvents As Boolean = False)
Dim M As Integer
Dim N As Integer
Dim FF As Integer
Dim mData As String
Dim mFieldQty As String

On Error GoTo Save_Error

If DbOpen = False Or myOpenMode = ReadOnly Then Exit Sub

If noEvents = False Then
RaiseEvent BeforeSave
RaiseEvent StateChanged(isBusy)
End If

If myFieldCount > 0 And myFieldCount <= 255 Then
mFieldQty = Chr(myFieldCount)
Else
Exit Sub
End If

If bAddNew = True Then
ReDim Preserve myRecord(myRecordCount)
myRecord(myRecordCount) = Join(myField, FieldSep)
bAddNew = False
ElseIf bModify = True Then
myRecord(myAbsolutePos) = Join(myField, FieldSep)
End If

mData = Join(myRecord, RecordSep)

If Len(Dir(myDatabaseName)) Then Kill myDatabaseName
DoEvents

FF = FreeFile
Open myDatabaseName For Binary Access Write As #FF
    Put #FF, , mData
Close #FF

If noEvents = False Then
RaiseEvent AfterSave
RaiseEvent StateChanged(isOpen)
End If

On Error GoTo 0
Exit Sub

Save_Error:

RaiseEvent Error(Err.Number, Err.Description, "Save")
End Sub

Public Sub AddNew()
On Error GoTo AddNew_Error

If bAddNew = True Or DbOpen = False Or myOpenMode = ReadWriteDenyAddNew Or _
myOpenMode = ReadOnly Then Exit Sub

RaiseEvent BeforeAddNew
bAddNew = True

If myMaxRecord > 0 And myRecordCount >= myMaxRecord Then
RaiseEvent MaxRecordReached
Exit Sub
End If

myRecordCount = myRecordCount + 1
ReDim Preserve myRecord(myRecordCount)
myAbsolutePos = myRecordCount

BtnEnabled True

On Error GoTo 0
Exit Sub

AddNew_Error:

RaiseEvent Error(Err.Number, Err.Description, "AddNew")
End Sub

Public Property Get Field(ByVal Index As Integer) As String
If DbOpen = False Then Exit Property
Field = Crypt(myField(Index))
End Property

Public Property Let Field(ByVal Index As Integer, ByVal vNewValue As String)
If DbOpen = False Then Exit Property
myField(Index) = Crypt(vNewValue)
End Property

Public Property Get DatabaseName() As String
DatabaseName = myDatabaseName
PropertyChanged "DatabaseName"
End Property

Public Property Let DatabaseName(ByVal vNewValue As String)
myDatabaseName = vNewValue
PropertyChanged "DatabaseName"
End Property

Public Property Get AbsolutePosition() As Long
Attribute AbsolutePosition.VB_MemberFlags = "400"
If DbOpen = False Then Exit Property
AbsolutePosition = myAbsolutePos
End Property

Public Property Let AbsolutePosition(ByVal vNewValue As Long)
If vNewValue <= 0 Or vNewValue > myRecordCount Or DbOpen = False Then Exit Property

myAbsolutePos = vNewValue
myField = Split(myRecord(myAbsolutePos), FieldSep)

If myAbsolutePos = 1 Then myBOF = True Else myBOF = False
If myAbsolutePos = myRecordCount Then myEOF = True Else myEOF = False

RaiseEvent Reposition
End Property

Public Sub MoveFirst()
On Error GoTo MoveFirst_Error

If DbOpen = False Then Exit Sub
myAbsolutePos = 1
myField = Split(myRecord(1), FieldSep)

If myAbsolutePos <= 1 Then myBOF = True Else myBOF = False
myEOF = False

RaiseEvent Reposition

On Error GoTo 0
Exit Sub

MoveFirst_Error:

RaiseEvent Error(Err.Number, Err.Description, "MoveFirst")
End Sub

Public Sub MoveLast()
On Error GoTo MoveLast_Error

If DbOpen = False Then Exit Sub
myAbsolutePos = myRecordCount
myField = Split(myRecord(myRecordCount), FieldSep)

If myAbsolutePos >= myRecordCount Then myEOF = True Else myEOF = False
myBOF = False

RaiseEvent Reposition

On Error GoTo 0
Exit Sub

MoveLast_Error:

RaiseEvent Error(Err.Number, Err.Description, "MoveLast")
End Sub

Public Sub MoveNext()
On Error GoTo MoveNext_Error

If DbOpen = False Then Exit Sub

myAbsolutePos = myAbsolutePos + 1

If myAbsolutePos > myRecordCount Then
    If myEOFAction = MoveOpposite Then
    myAbsolutePos = 1
    Else
    myEOF = True
    myAbsolutePos = myRecordCount
    End If
Exit Sub
Else
myEOF = False
End If

myField = Split(myRecord(myAbsolutePos), FieldSep)

RaiseEvent Reposition

On Error GoTo 0
Exit Sub

MoveNext_Error:

RaiseEvent Error(Err.Number, Err.Description, "MoveNext")
End Sub

Public Sub MovePrevious()
On Error GoTo MovePrevious_Error

If DbOpen = False Then Exit Sub

myAbsolutePos = myAbsolutePos - 1

If myAbsolutePos <= 0 Then
    If myEOFAction = MoveOpposite Then
    myAbsolutePos = myRecordCount
    Else
    myBOF = True
    myAbsolutePos = 1
    End If
Exit Sub
Else
myEOF = False
End If

myField = Split(myRecord(myAbsolutePos), FieldSep)

RaiseEvent Reposition

On Error GoTo 0
Exit Sub

MovePrevious_Error:

RaiseEvent Error(Err.Number, Err.Description, "MovePrevious")
End Sub

Public Sub Modify()
On Error GoTo Modify_Error

If DbOpen = False Or bAddNew = True Or bModify = True Or myOpenMode = ReadOnly Then Exit Sub
RaiseEvent BeforeModify
bModify = True

On Error GoTo 0
Exit Sub

Modify_Error:

RaiseEvent Error(Err.Number, Err.Description, "Modify")
End Sub

Public Sub CloseDatabase()
On Error GoTo CloseDatabase_Error

RaiseEvent BeforeClose
ReDim myField(0)
ReDim myRecord(0)
myAbsolutePos = 0
myRecordCount = 0
DbOpen = False
RaiseEvent AfterClose
RaiseEvent StateChanged(isClosed)

BtnEnabled False
On Error GoTo 0
Exit Sub

CloseDatabase_Error:

RaiseEvent Error(Err.Number, Err.Description, "CloseDatabase")
End Sub

Public Sub Delete(Optional Position As Long = -1)
Dim M As Integer

On Error GoTo Delete_Error

If DbOpen = False Then Exit Sub

If myOpenMode = ReadOnly Then Exit Sub

RaiseEvent BeforeDelete
If Position = -1 Then Position = myAbsolutePos

If Position < myRecordCount Then
    For M = Position To myRecordCount - 1
    myRecord(M) = myRecord(M + 1)
    Next M
End If
myRecordCount = myRecordCount - 1
ReDim Preserve myRecord(myRecordCount)
Save
MoveFirst
RaiseEvent AfterDelete

On Error GoTo 0
Exit Sub

Delete_Error:

RaiseEvent Error(Err.Number, Err.Description, "Delete")
End Sub

Public Property Get FieldCount() As Integer
FieldCount = myFieldCount
End Property


Public Property Get RecordCount() As Long
RecordCount = myRecordCount
End Property


Public Function Find(FindWhat As String, Optional StartRec As Long = -1, Optional WhichField As Integer = -1, _
Optional FindMethod As eFind = FindNext, Optional Compare As eCompare, Optional GotoFoundRecord As Boolean = True) As Long

Dim M As Long
Dim tField As String
On Error GoTo Find_Error

Find = 0

RaiseEvent BeforeFind
Select Case FindMethod
    Case FindFirst
        StartRec = 1
    Case FindPrevious
        If StartRec = -1 Then StartRec = myAbsolutePos - 1
    Case FindNext
        If StartRec = -1 Then StartRec = myAbsolutePos + 1
    Case FindLast
        StartRec = myRecordCount
End Select

If Compare = TextCompare Then FindWhat = LCase(FindWhat)

If WhichField = -1 Then
FindWhat = "*" & FindWhat & "*"
    If FindMethod = FindFirst Or FindMethod = FindNext Then
        For M = StartRec To myRecordCount
        RaiseEvent Progress(M)
            If Compare = TextCompare Then
            tField = LCase(Crypt(myRecord(M)))
            Else
            tField = Crypt(myRecord(M))
            End If
            
            If tField Like FindWhat Then
            Find = M
            Exit For
            End If
        Next M
    ElseIf FindMethod = FindPrevious Or FindMethod = FindLast Then
        For M = StartRec To 1 Step -1
        RaiseEvent Progress(M)
            If Compare = TextCompare Then
            tField = LCase(Crypt(myRecord(M)))
            Else
            tField = Crypt(myRecord(M))
            End If
            
            If tField Like FindWhat Then
            Find = M
            Exit For
            End If
        Next M
    End If
ElseIf WhichField <= myFieldCount Then
    If FindMethod = FindFirst Or FindMethod = FindNext Then
        For M = StartRec To myRecordCount
        RaiseEvent Progress(M)
        tField = Split(myRecord(M), FieldSep)(WhichField)
        tField = Crypt(tField)
        If Compare = TextCompare Then tField = LCase(tField)
            If tField Like FindWhat Then
            Find = M
            Exit For
            End If
        Next M
    ElseIf FindMethod = FindPrevious Or FindMethod = FindLast Then
        For M = StartRec To 1 Step -1
        RaiseEvent Progress(M)
        tField = Split(myRecord(M), FieldSep)(WhichField)
        tField = Crypt(tField)
        If Compare = TextCompare Then tField = LCase(tField)
            If tField Like FindWhat Then
            Find = M
            Exit For
            End If
        Next M
    End If
End If

RaiseEvent AfterFind(CBool(Find))
If Find > 0 And GotoFoundRecord = True Then
myAbsolutePos = Find
myField = Split(myRecord(myAbsolutePos), FieldSep)

If myAbsolutePos = 1 Then myBOF = True Else myBOF = False
If myAbsolutePos = myRecordCount Then myEOF = True Else myEOF = False

RaiseEvent Reposition
End If

On Error GoTo 0
Exit Function

Find_Error:

RaiseEvent Error(Err.Number, Err.Description, "Find")
End Function

Public Property Get BOFAction() As eEndAction
BOFAction = myBOFAction
End Property

Public Property Let BOFAction(ByVal vNewValue As eEndAction)
myBOFAction = vNewValue
PropertyChanged "BOFAction"
End Property
Public Property Get EOFAction() As eEndAction
EOFAction = myEOFAction
End Property

Public Property Let EOFAction(ByVal vNewValue As eEndAction)
myEOFAction = vNewValue
PropertyChanged "EOFAction"
End Property

Public Sub Update() 'Not for saving -> for adding record faster
bAddNew = False
bModify = False

If myOpenMode = ReadOnly Then Exit Sub

myRecord(myAbsolutePos) = Join(myField, FieldSep)
End Sub

Private Function Crypt(myString As String, Optional NewPW As String = "-1") As String
Dim iI As Integer
Dim jJ As Integer
Dim cTemp As Byte
Dim Temp As String
Dim eTemp As String
Dim P As Integer
Dim tPW As String
Dim tLenC As Integer
P = 0

If NewPW = "-1" Then
tPW = myPassword
Else
tPW = NewPW
End If

If tPW = "" Then
Crypt = myString
Exit Function
End If

If Len(myString) > myMaxCryptLen Then ' To prevent to slow down the crypt process
tLenC = MaxCryptLen
Else
tLenC = Len(myString)
End If

For iI = 1 To tLenC
    Temp = Mid$(myString, iI, 1)
    cTemp = Asc(Temp)
        For jJ = 1 To Len(tPW)
        cTemp = cTemp Xor (Asc(Mid$(tPW, jJ, 1)))
        Next jJ
    P = P + 1
    If P = Len(tPW) Then P = 1
    cTemp = cTemp Xor Asc(Mid$(tPW, P, 1))
    eTemp = eTemp + Chr$(cTemp)
Next iI

If Len(myString) > MaxCryptLen Then
Crypt = eTemp$ & Right(myString, Len(myString) - MaxCryptLen)
Else
Crypt = eTemp$
End If

End Function


Public Sub Sort(Optional Direction As eSortDirection, Optional Compare As eCompare = TextCompare)
On Error GoTo Sort_Error

If myPassword <> "" Then
MsgBox "Sorry, can't sort a database protected by password"
Exit Sub
End If

If DbOpen = False Or bAddNew = True Or bModify = True Then Exit Sub

strSort myRecord, 1, myRecordCount, Compare, Direction

MoveFirst

On Error GoTo 0
Exit Sub

Sort_Error:

RaiseEvent Error(Err.Number, Err.Description, "Sort")
End Sub

Private Sub strSort(sA() As String, ByVal lbA As Long, ByVal ubA As Long, _
Optional Compare As eCompare = TextCompare, Optional ByVal Direction As eSortDirection = sortAsc)
Dim ptr1 As Long, ptr2 As Long, cnt As Long
Dim lpS As Long, idx As Long, pvt As Long
Dim inter1 As Long, inter2 As Long
Dim Item As String, lpStr As Long
Dim lA_1() As Long, lpL_1 As Long
Dim lA_2() As Long, lpL_2 As Long
Dim base As Long, optimal As Long
Dim run As Long, cast As Long

idx = ubA - lbA ' cnt-1
ReDim lA_1(n0 To idx) As Long
ReDim lA_2(n0 To idx) As Long
lpL_1 = VarPtr(lA_1(n0))
lpL_2 = VarPtr(lA_2(n0))

pvt = (idx \ n8) + n16             ' Allow for worst case senario + some
ReDim lbs(n1 To pvt) As Long       ' Stack to hold pending lower boundries
ReDim ubs(n1 To pvt) As Long       ' Stack to hold pending upper boundries
lpStr = VarPtr(Item)               ' Cache pointer to the string variable
lpS = VarPtr(sA(lbA)) - (lbA * n4) ' Cache pointer to the string array

Do: ptr1 = n0: ptr2 = n0
    pvt = ((ubA - lbA) \ n2) + lbA         ' Get pivot index position
    CopyMemByV lpStr, lpS + (pvt * n4), n4 ' Grab current value into item

    For idx = lbA To pvt - n1
        If (StrComp(sA(idx), Item, Compare) = Direction) Then ' (idx > item)
            CopyMemByV lpL_2 + (ptr2 * n4), lpS + (idx * n4), n4  '3
            ptr2 = ptr2 + n1
        Else
            CopyMemByV lpL_1 + (ptr1 * n4), lpS + (idx * n4), n4  '1
            ptr1 = ptr1 + n1
        End If
    Next
    inter1 = ptr1: inter2 = ptr2
    For idx = pvt + n1 To ubA
        If (StrComp(Item, sA(idx), Compare) = Direction) Then ' (idx < item)
            CopyMemByV lpL_1 + (ptr1 * n4), lpS + (idx * n4), n4  '2
            ptr1 = ptr1 + n1
        Else
            CopyMemByV lpL_2 + (ptr2 * n4), lpS + (idx * n4), n4  '4
            ptr2 = ptr2 + n1
        End If
    Next '-Avalanche v2 ©Rd-
    CopyMemByV lpS + (lbA * n4), lpL_1, ptr1 * n4
    CopyMemByV lpS + ((lbA + ptr1) * n4), lpStr, n4 ' Re-assign current
    CopyMemByV lpS + ((lbA + ptr1 + n1) * n4), lpL_2, ptr2 * n4

    If (ptr1 - inter1 <> n0) Then
        CopyMemByV lpStr, lpS + ((lbA + ptr1 - n1) * n4), n4
        optimal = lbA + (inter1 \ n2)
        run = lbA + inter1
        Do While run > optimal                      ' Runner do loop
            If StrComp(sA(run - n1), Item, Compare) <> Direction Then Exit Do
            run = run - n1
        Loop: cast = lbA + inter1 - run
        If cast <> n0 Then
            CopyMemByV lpL_1, lpS + (run * n4), cast * n4                      ' Grab current value(s)
            CopyMemByV lpS + (run * n4), lpS + ((lbA + inter1) * n4), (ptr1 - inter1) * n4 ' Move up items
            CopyMemByV lpS + ((lbA + ptr1 - cast - n1) * n4), lpL_1, cast * n4 ' Re-assign current value(s) into found pos
        End If
    End If '1 2 i 3 4
    If (ptr2 - inter2 <> n0) Then
        base = lbA + ptr1 + n1
        CopyMemByV lpStr, lpS + (base * n4), n4
        pvt = lbA + ptr1 + inter2
        optimal = pvt + ((ptr2 - inter2) \ n2)
        run = pvt
        Do While run < optimal                      ' Runner do loop
            If StrComp(sA(run + n1), Item, Compare) <> Direction Then Exit Do
            run = run + n1
        Loop: cast = run - pvt
        If cast <> n0 Then
            CopyMemByV lpL_1, lpS + ((pvt + n1) * n4), cast * n4 ' Grab current value(s)
            CopyMemByV lpS + ((base + cast) * n4), lpS + (base * n4), inter2 * n4 ' Move up items
            CopyMemByV lpS + (base * n4), lpL_1, cast * n4       ' Re-assign current value(s) into found pos
    End If: End If

    If (ptr1 > n1) Then
        If (ptr2 > n1) Then cnt = cnt + n1: lbs(cnt) = lbA + ptr1 + n1: ubs(cnt) = ubA
        ubA = lbA + ptr1 - n1
    ElseIf (ptr2 > n1) Then
        lbA = lbA + ptr1 + n1
    Else
        If cnt = n0 Then Exit Do
        lbA = lbs(cnt): ubA = ubs(cnt): cnt = cnt - n1
    End If
Loop: Erase lbs: Erase ubs: CopyMem ByVal lpStr, 0&, n4
End Sub



Private Sub BtnEnabled(isEnabled As Boolean)
BtnPrev.Enabled = isEnabled
BtnNext.Enabled = isEnabled
BtnFirst.Enabled = isEnabled
BtnLast.Enabled = isEnabled
End Sub

Public Sub Filter(FilterStr As String, Optional Compare As eCompare = TextCompare)
' FieldIndex1=Text1,FieldIndex2=Text2
Dim M As Long
Dim N As Long
Dim tInd() As Integer
Dim tFilt() As String
Dim tStr() As String
Dim fQty As Integer
Dim newAr() As String
Dim aQty As Long
Dim tField As String

On Error GoTo Filter_Error

OpenDatabase

RaiseEvent BeforeFilter
If FilterStr <> "" Then
If Compare = TextCompare Then FilterStr = LCase(FilterStr)
If InStr(FilterStr, "=") = 0 Then Exit Sub
    If InStr(FilterStr, ",") = 0 Then
    ReDim tFilt(1)
    tFilt(0) = FilterStr
    fQty = 1
    Else
    tFilt = Split(FilterStr, ",")
    fQty = UBound(tFilt) + 1
    End If


ReDim tInd(fQty)
ReDim tStr(fQty)

    For M = 0 To fQty - 1
    tInd(M) = Trim(Split(tFilt(M), "=")(0))
    tStr(M) = Trim(Split(tFilt(M), "=")(1))
    Next M
    
    For M = 0 To fQty - 1
        For N = 1 To myRecordCount
            tField = Split(myRecord(N), FieldSep)(tInd(M))
            tField = Crypt(tField)
            If Compare = TextCompare Then tField = LCase(tField)
                If tField Like tStr(M) Then
                aQty = aQty + 1
                ReDim Preserve newAr(1 To aQty)
                newAr(aQty) = myRecord(N)
                End If
        Next N
    If aQty > 0 Then
    myRecord = newAr
    myRecordCount = UBound(myRecord)
    Else
    Exit For
    End If
    aQty = 0
    Next M
RaiseEvent AfterFilter
End If
MoveFirst

On Error GoTo 0
Exit Sub

Filter_Error:

RaiseEvent Error(Err.Number, Err.Description, "Filter")
End Sub

Public Property Get Password() As String
Password = myPassword
End Property

Public Property Let Password(ByVal vNewValue As String)
myPassword = vNewValue
PropertyChanged "Password"
End Property

Public Property Get OpenMode() As eOpenMode
OpenMode = myOpenMode
End Property

Public Property Let OpenMode(ByVal vNewValue As eOpenMode)
myOpenMode = vNewValue
PropertyChanged "OpenMode"
End Property

Public Property Get MaxRecord() As Long
MaxRecord = myMaxRecord
End Property

Public Property Let MaxRecord(ByVal vNewValue As Long)
myMaxRecord = vNewValue
PropertyChanged "MaxRecord"
End Property

Public Property Get BOF() As Boolean
BOF = myBOF
End Property


Public Property Get EOF() As Boolean
EOF = myEOF
End Property


Public Sub ChangePassword(OldPassword As String, NewPassword As String)
On Error GoTo ChangePassword_Error

If DbOpen = False Then Exit Sub
Dim M As Long
Dim N As Long
Dim tField() As String

If OldPassword <> myPassword Then
RaiseEvent PasswordError
Exit Sub
End If

RaiseEvent BeforeChangePassword

For M = 1 To myRecordCount
RaiseEvent Progress(M)
tField = Split(myRecord(M), FieldSep)
    For N = 0 To myFieldCount - 1
    tField(N) = Crypt(tField(N))
    tField(N) = Crypt(tField(N), NewPassword)
    Next N
myRecord(M) = Join(tField, FieldSep)
Next M

myRecord(0) = FileIdent & Chr(myFieldCount) & Crypt("DB", NewPassword) & String(14, Chr(0))

myPassword = NewPassword
Save True

RaiseEvent PasswordChanged

On Error GoTo 0
Exit Sub

ChangePassword_Error:

RaiseEvent Error(Err.Number, Err.Description, "ChangePassword")
End Sub

Public Sub InsertFile(FileName As String, FieldIndex As Integer)
Dim FF As Long
Dim tmpFile As String
Dim FileLenght As Long

On Error GoTo InsertFile_Error

If FieldIndex > myFieldCount - 1 Then Err.Raise 9

FF = FreeFile

tmpFile = Space(FileLen(FileName))
Open FileName For Binary As #FF
Get #FF, , tmpFile
Close #FF

myField(FieldIndex) = Crypt(tmpFile)

On Error GoTo 0
Exit Sub

InsertFile_Error:

RaiseEvent Error(Err.Number, Err.Description, "InsertFile")

End Sub
Public Sub ExtractFile(FileName As String, FieldIndex As Integer)
Dim FF As Long
Dim tmpFile As String
Dim FileLenght As Long

On Error GoTo ExtractFile_Error

If FieldIndex > myFieldCount - 1 Then Err.Raise 9

tmpFile = Crypt(myField(FieldIndex))

FF = FreeFile

Open FileName For Binary Access Write As #FF
    Put #FF, , tmpFile
Close #FF

On Error GoTo 0
Exit Sub

ExtractFile_Error:

RaiseEvent Error(Err.Number, Err.Description, "ExtractFile")

End Sub

Public Property Get MaxCryptLen() As Integer
MaxCryptLen = myMaxCryptLen
End Property

Public Property Let MaxCryptLen(ByVal vNewValue As Integer)
If vNewValue > 0 Then
myMaxCryptLen = vNewValue
PropertyChanged "MaxCryptLen"
End If
End Property

Public Property Get CaptionMultiLine() As Boolean
CaptionMultiLine = myCapML
End Property

Public Property Let CaptionMultiLine(ByVal vNewValue As Boolean)
myCapML = vNewValue
UserControl_Resize
PropertyChanged "CaptionMultiLines"
End Property
