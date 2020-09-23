VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form MainForm 
   Caption         =   "by_ www.MixofTix.net (cabs, drivers, insurance... databank)"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   10560
      TabIndex        =   22
      Text            =   "130.00"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Return All Records (Fail-Filtering)"
      Height          =   975
      Left            =   5160
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox tx 
      Height          =   285
      Index           =   8
      Left            =   7560
      TabIndex        =   16
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox tx 
      Height          =   285
      Index           =   7
      Left            =   7560
      TabIndex        =   15
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox tx 
      Height          =   285
      Index           =   5
      Left            =   3360
      TabIndex        =   11
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox tx 
      Height          =   285
      Index           =   4
      Left            =   3360
      TabIndex        =   10
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox tx 
      Height          =   285
      Index           =   3
      Left            =   3360
      TabIndex        =   9
      Top             =   840
      Width           =   1695
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "Form1.frx":0000
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "Form1.frx":0014
      TabIndex        =   8
      Top             =   6000
      Width           =   11655
   End
   Begin VB.Data Data3 
      Caption         =   "insuranceinfo"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Desktop\Amir\amir.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "insuranceinfo"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "Form1.frx":24B3
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "Form1.frx":24C7
      TabIndex        =   7
      Top             =   1200
      Width           =   11655
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "Form1.frx":3F26
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "Form1.frx":3F3A
      TabIndex        =   6
      Top             =   3600
      Width           =   11655
   End
   Begin VB.Data Data2 
      Caption         =   "driverinfo"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Desktop\Amir\amir.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "driverinfo"
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox tx 
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox tx 
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox tx 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Main Information"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\Desktop\Amir\amir.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "main"
      Top             =   840
      Width           =   5400
   End
   Begin VB.Label Label6 
      Caption         =   "Current:"
      Height          =   255
      Left            =   9480
      TabIndex        =   23
      Top             =   480
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   9360
      X2              =   9360
      Y1              =   120
      Y2              =   720
   End
   Begin VB.Label Label5 
      Caption         =   "LOU Rate($):"
      Height          =   255
      Left            =   9480
      TabIndex        =   21
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "filenum"
      DataSource      =   "Data1"
      Height          =   255
      Left            =   10320
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "FH#:"
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   18
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Insurance Co.:"
      Height          =   255
      Index           =   2
      Left            =   6360
      TabIndex        =   17
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "DOA:"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   14
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "CCIC#:"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   13
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Status:"
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   12
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Policy#:"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Cab#:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "File#:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'By_ Shahin Noursalehi - copy 2000-2001
'Contact: admin@MixofTix.net
'Terms of Agreement:
'By using this code, you agree to the following terms...
'1) You may use this code in your own programs (and may compile it into a program and distribute it in compiled format for languages that allow it) freely and with no charge.
'2) You MAY NOT redistribute this code (for example to a web site) without written permission from the original author. Failure to do so is a violation of copyright laws.
'3) You may link to this code from another website, but ONLY if it is not wrapped in a frame.
'4) You will abide by any additional copyright restrictions which the author may have placed in the code or code's description.
'5) Not for commercial use!


Private Sub Command1_Click()
strSql = "select * from main"
Data1.RecordSource = strSql
Data1.Refresh
DBGrid2.ReBind
strSql = "select * from driverinfo"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
strSql = "select * from insuranceinfo"
Data3.RecordSource = strSql
Data3.Refresh
DBGrid3.ReBind
End Sub

Private Sub Data1_Reposition()
On Error GoTo errhand
Dim strSql As String
Dim StrLink As String
If Data1.Recordset.EOF = True And Data1.Recordset.BOF = True Then
'MsgBox "Amir-Joon, No Records matched!", vbInformation, "Alert to Amir-Joon..."
Else
'ident = DBGrid2.Text
ident = Label4.Caption
If ident <> "" Then
strSql = "select * from driverinfo where filenum='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
strSql = "select * from insuranceinfo where filenum='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
DBGrid3.ReBind
End If
End If
Exit Sub
errhand:
MsgBox Err.Description
End Sub

Private Sub DBGrid1_OnAddNew()
DBGrid1.Col = 0   'Current Row
DBGrid1.Text = Label4.Caption
End Sub

Private Sub DBGrid2_AfterDelete()
ddd = Label4.Caption
'Data2.RecordSource = "delete * from driverinfo where filenum='" & ddd & "'"
'Data3.RecordSource = "delete * from insuranceinfo where filenum='" & ddd & "'"

    Dim ws As Workspace
    Dim db As Database
    '
    Dim strDBName As String
    Dim strSql As String
    Dim IngRecords As Long
    '
    strDBName = App.Path & "\amir.mdb"
    strSql = "delete * from driverinfo where filenum='" & ddd & "'"
    IngRecords = 0
    '
    Set ws = DBEngine.Workspaces(0)
    Set db = ws.OpenDatabase(strDBName)
    '
    With db
        .Execute strSql, dbFailOnError
        IngRecords = .RecordsAffected
    End With
    '
    MsgBox "Drivers info Affected = " & CStr(IngRecords), vbInformation, "Amir-Joon Info"
    strSql = "delete * from insuranceinfo where filenum='" & ddd & "'"
    IngRecords = 0
    '
    'Set ws = DBEngine.Workspaces(0)
    'Set db = ws.OpenDatabase(strDBName)
    '
    With db
        .Execute strSql, dbFailOnError
        IngRecords = .RecordsAffected
    End With
    '
    MsgBox "Insurance Info Affected = " & CStr(IngRecords), vbInformation, "Amir-Joon Info"
'    db.Close
'    ws.Close
'    Set db = Nothing
'    Set ws = Nothing

End Sub

Private Sub DBGrid2_AfterInsert()
DBGrid2.Col = 1    'Cab#
DBGrid2.Row = fixedrow   'Current Row
CabText = DBGrid2.Text
DBGrid2.Col = 2    'DOA
DBGrid2.Row = fixedrow   'Current Row
DOAText = CDate(DBGrid2.Text)
mmDOAText = Month(DOAText)
ddDOAText = Day(DOAText)
yyDOAText = Year(DOAText)
totaltext = CabText & "DA" & mmDOAText & ddDOAText & yyDOAText
DBGrid2.Col = 0    'File#
DBGrid2.Row = fixedrow   'Current Row
DBGrid2.Text = totaltext
'Data1.Recordset.MoveLast
'MsgBox "ff"
DBGrid2.ReBind
End Sub

Private Sub DBGrid2_BeforeInsert(Cancel As Integer)
'MsgBox DBGrid2.Row
'fixedrow = DBGrid2.Row 'Current Row
End Sub

Private Sub DBGrid2_BeforeUpdate(Cancel As Integer)
ident = Label4.Caption
'\\\\\\\\
fixedrow = DBGrid2.Row 'Current Row
DBGrid2.Col = 1    'Cab#
DBGrid2.Row = fixedrow   'Current Row
CabText = DBGrid2.Text
DBGrid2.Col = 2    'DOA
DBGrid2.Row = fixedrow   'Current Row
DOAText = CDate(DBGrid2.Text)
mmDOAText = Month(DOAText)
ddDOAText = Day(DOAText)
yyDOAText = Year(DOAText)
totaltext = CabText & "DA" & mmDOAText & ddDOAText & yyDOAText
DBGrid2.Col = 0    'File#
DBGrid2.Row = fixedrow   'Current Row
DBGrid2.Text = totaltext
'Data1.Recordset.MoveLast
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
strSql = "select * from driverinfo where filenum='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
If Data2.Recordset.BOF = True And Data2.Recordset.EOF = True Then
'
Else
Data2.Recordset.MoveLast
i = Data2.Recordset.RecordCount
'""""""""""""""""""""
Data2.Recordset.MoveFirst
  For j = 0 To i - 1
    Data2.Recordset.Edit
    Data2.Recordset("filenum").Value = totaltext
    Data2.Recordset.Update
    Data2.Recordset.MoveNext
  Next j
'""""""""""""""""""""
End If
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
strSql = "select * from insuranceinfo where filenum='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
If Data3.Recordset.BOF = True And Data3.Recordset.EOF = True Then
'
Else
Data3.Recordset.MoveLast
i = Data3.Recordset.RecordCount
'""""""""""""""""""""
Data3.Recordset.MoveFirst
  For j = 0 To i - 1
    Data3.Recordset.Edit
    Data3.Recordset("filenum").Value = totaltext
    Data3.Recordset.Update
    Data3.Recordset.MoveNext
  Next j
'""""""""""""""""""""
End If
'/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\
End Sub

Private Sub DBGrid3_AfterInsert()
DBGrid3.Col = 12    'NODays
DBGrid3.Row = fixedrow   'Current Row
NODText = DBGrid3.Text
'MsgBox NODText
totaltext = Val(NODText) * Val(Text1.Text)
DBGrid3.Col = 11    'LOU
DBGrid3.Row = fixedrow   'Current Row
DBGrid3.Text = CStr(totaltext)
'Data1.Recordset.MoveLast
'MsgBox "ff"
'/\/\/\/\/\/\/\//\/\/\/\/\/\/\/\/\/\\/\/\/\/\/\
DBGrid3.Col = 9    '////////////
DBGrid3.Row = fixedrow   'Current Row
r1 = DBGrid3.Text
DBGrid3.Col = 13    '////////////
DBGrid3.Row = fixedrow   'Current Row
r2 = DBGrid3.Text
DBGrid3.Col = 15   '////////////
DBGrid3.Row = fixedrow   'Current Row
r3 = DBGrid3.Text
totalrec = Val(r1) + Val(r2) + Val(r3)
DBGrid3.Col = 16    '///////////
DBGrid3.Row = fixedrow   'Current Row
DBGrid3.Text = CStr(totalrec)
DBGrid3.Col = 17    '///////////
DBGrid3.Row = fixedrow   'Current Row
DBGrid3.Text = CStr(totalrec * 1.15)

DBGrid3.ReBind
'MsgBox totaltext
End Sub

Private Sub DBGrid3_BeforeUpdate(Cancel As Integer)
fixedrow = DBGrid3.Row 'Current Row
DBGrid3.Col = 12    'NODays
DBGrid3.Row = fixedrow   'Current Row
NODText = DBGrid3.Text
'MsgBox NODText
totaltext = Val(NODText) * Val(Text1.Text)
DBGrid3.Col = 11    'LOU
DBGrid3.Row = fixedrow   'Current Row
DBGrid3.Text = CStr(totaltext)
'/\/\/\/\/\/\/\//\/\/\/\/\/\/\/\/\/\\/\/\/\/\/\
DBGrid3.Col = 9    '////////////
DBGrid3.Row = fixedrow   'Current Row
r1 = DBGrid3.Text
DBGrid3.Col = 13    '////////////
DBGrid3.Row = fixedrow   'Current Row
r2 = DBGrid3.Text
DBGrid3.Col = 15   '////////////
DBGrid3.Row = fixedrow   'Current Row
r3 = DBGrid3.Text
totalrec = Val(r1) + Val(r2) + Val(r3)
DBGrid3.Col = 16    '///////////
DBGrid3.Row = fixedrow   'Current Row
DBGrid3.Text = CStr(totalrec)
DBGrid3.Col = 17    '///////////
DBGrid3.Row = fixedrow   'Current Row
DBGrid3.Text = CStr(totalrec * 1.15)

End Sub

Private Sub DBGrid3_OnAddNew()
DBGrid3.Col = 0   'Current Row
DBGrid3.Text = Label4.Caption

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\amir.mdb"
Data2.DatabaseName = App.Path & "\amir.mdb"
Data3.DatabaseName = App.Path & "\amir.mdb"
 DBGrid2.Col = 0
 DBGrid2.Row = 0
End Sub

Private Sub tx_Change(Index As Integer)
ident = tx(Index).Text
Select Case Index
Case 0
internalVar = "filenum"
strSql = "select * from main where " & internalVar & "='" & ident & "'"
Data1.RecordSource = strSql
Data1.Refresh
If Not Data1.Recordset.BOF = True And Not Data1.Recordset.EOF = True Then
Data1.Recordset.MoveFirst
End If
DBGrid2.ReBind
If ident <> "" Then
strSql = "select * from driverinfo where filenum='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
strSql = "select * from insuranceinfo where filenum='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
DBGrid3.ReBind
End If


Case 1
internalVar = "cabnum"
strSql = "select * from main where " & internalVar & "='" & ident & "'"
Data1.RecordSource = strSql
Data1.Refresh
If Not Data1.Recordset.BOF = True And Not Data1.Recordset.EOF = True Then
Data1.Recordset.MoveFirst
End If
DBGrid2.ReBind
If ident <> "" Then
strSql = "select * from driverinfo where filenum='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
strSql = "select * from insuranceinfo where filenum='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
DBGrid3.ReBind
End If


Case 2
internalVar = "policy"
strSql = "select * from main where " & internalVar & "='" & ident & "'"
Data1.RecordSource = strSql
Data1.Refresh
If Not Data1.Recordset.BOF = True And Not Data1.Recordset.EOF = True Then
Data1.Recordset.MoveFirst
End If
DBGrid2.ReBind
If ident <> "" Then
strSql = "select * from driverinfo where filenum='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
strSql = "select * from insuranceinfo where filenum='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
DBGrid3.ReBind
End If


Case 3
internalVar = "status"
strSql = "select * from main where " & internalVar & "='" & ident & "'"
Data1.RecordSource = strSql
Data1.Refresh
If Not Data1.Recordset.BOF = True And Not Data1.Recordset.EOF = True Then
Data1.Recordset.MoveFirst
End If
DBGrid2.ReBind
If ident <> "" Then
strSql = "select * from driverinfo where filenum='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
strSql = "select * from insuranceinfo where filenum='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
DBGrid3.ReBind
End If


Case 4
internalVar = "ccic"
strSql = "select * from main where " & internalVar & "='" & ident & "'"
Data1.RecordSource = strSql
Data1.Refresh
If Not Data1.Recordset.BOF = True And Not Data1.Recordset.EOF = True Then
Data1.Recordset.MoveFirst
End If
DBGrid2.ReBind
If ident <> "" Then
strSql = "select * from driverinfo where filenum='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
strSql = "select * from insuranceinfo where filenum='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
DBGrid3.ReBind
End If


Case 5
internalVar = "doa"
strSql = "select * from main where " & internalVar & "='" & ident & "'"
Data1.RecordSource = strSql
Data1.Refresh
If Not Data1.Recordset.BOF = True And Not Data1.Recordset.EOF = True Then
Data1.Recordset.MoveFirst
End If
DBGrid2.ReBind
If ident <> "" Then
strSql = "select * from driverinfo where filenum='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
strSql = "select * from insuranceinfo where filenum='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
DBGrid3.ReBind
End If


Case 7
internalVar = "company"
strSql = "select * from insuranceinfo where " & internalVar & "='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
If Not Data3.Recordset.BOF = True And Not Data3.Recordset.EOF = True Then
Data3.Recordset.MoveFirst
ident = Data3.Recordset.Fields("filenum").Value
End If
DBGrid3.ReBind
If ident <> "" Then
strSql = "select * from main where filenum='" & ident & "'"
Data1.RecordSource = strSql
Data1.Refresh
DBGrid2.ReBind
strSql = "select * from driverinfo where filenum='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
DBGrid1.ReBind
End If

Case 8
internalVar = "fhnum"
strSql = "select * from driverinfo where " & internalVar & "='" & ident & "'"
Data2.RecordSource = strSql
Data2.Refresh
If Not Data2.Recordset.BOF = True And Not Data2.Recordset.EOF = True Then
Data2.Recordset.MoveFirst
ident = Data2.Recordset.Fields("filenum").Value
End If
DBGrid1.ReBind
If ident <> "" Then
strSql = "select * from main where filenum='" & ident & "'"
Data1.RecordSource = strSql
Data1.Refresh
DBGrid2.ReBind
strSql = "select * from insuranceinfo where filenum='" & ident & "'"
Data3.RecordSource = strSql
Data3.Refresh
DBGrid3.ReBind
End If
End Select
Call Data1_Reposition
End Sub
