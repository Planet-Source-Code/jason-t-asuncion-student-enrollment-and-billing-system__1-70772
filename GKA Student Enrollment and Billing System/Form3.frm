VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form Form3 
   Caption         =   "Main Form"
   ClientHeight    =   7470
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7275
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7470
   ScaleWidth      =   7275
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   600
      TabIndex        =   19
      Text            =   "Text10"
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   960
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   2760
      TabIndex        =   14
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1440
      TabIndex        =   13
      Text            =   "Text7"
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   360
      TabIndex        =   12
      Text            =   "Text6"
      Top             =   960
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   720
      TabIndex        =   11
      Text            =   "Text5"
      Top             =   6000
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   6360
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4320
      TabIndex        =   9
      Text            =   "Text3"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin lvButton.lvButtons_H lvButtons_H2 
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Print"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14864790
      cGradient       =   14864790
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   13279782
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4680
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   5
      Top             =   6240
      Visible         =   0   'False
      Width           =   375
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      Caption         =   "Generate"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cBhover         =   14864790
      cGradient       =   14864790
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   13279782
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   255
      Left            =   2520
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39178
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39178
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5880
      Top             =   6600
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   4800
      TabIndex        =   16
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image Image3 
      Height          =   1140
      Left            =   3060
      Picture         =   "Form3.frx":2BEB3
      Top             =   240
      Width           =   1140
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   360
      TabIndex        =   18
      Top             =   1560
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   7515
      Left            =   0
      Picture         =   "Form3.frx":302A5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8130
   End
   Begin VB.Image Image2 
      Height          =   1140
      Left            =   3502
      Picture         =   "Form3.frx":49CF7
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   645
      TabIndex        =   17
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Menu View1 
      Caption         =   "&View"
      WindowList      =   -1  'True
      Begin VB.Menu Income1 
         Caption         =   "Income"
         Shortcut        =   ^I
      End
      Begin VB.Menu Expenses1 
         Caption         =   "Expenses"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Form1 
      Caption         =   "&Forms"
      Begin VB.Menu IncomeF 
         Caption         =   "Income Form"
         Shortcut        =   ^A
      End
      Begin VB.Menu ExpensesF 
         Caption         =   "Expenses Form"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu Print1 
      Caption         =   "&Print"
      Begin VB.Menu IncomeR 
         Caption         =   "Income Report"
         Shortcut        =   ^C
      End
      Begin VB.Menu ExpensesR 
         Caption         =   "Expenses Report"
         Shortcut        =   ^D
      End
      Begin VB.Menu Net 
         Caption         =   "Net Income Report"
         Shortcut        =   ^N
      End
   End
   Begin VB.Menu About1 
      Caption         =   "Abou&t"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim c As String
Dim d As String
Dim e As String
Dim h As Integer
Dim a As Double
Dim jason1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Private rs As New ADODB.Recordset
Private db As New ADODB.Connection




Private Sub About1_Click()
Form4.Show vbModal
End Sub

Private Sub Expenses1_Click()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = False

DTPicker1.Visible = True
DTPicker2.Visible = True
Text2.Visible = False
Text1.Text = "2"
Set DataGrid1.DataSource = Nothing
lvButtons_H2.Visible = False
lvButtons_H1.Visible = True
End Sub

Private Sub ExpensesF_Click()
Form2.Show vbModal
Set DataGrid1.DataSource = Nothing
End Sub

Private Sub ExpensesR_Click()
Label1.Visible = True
Label2.Visible = True
DTPicker1.Visible = True
DTPicker2.Visible = True
Text2.Visible = False
Label3.Visible = False
lvButtons_H2.Visible = True
lvButtons_H1.Visible = False
Text1.Text = "4"
Text4.Text = "Expenses"
Text5.Text = "Expenses Report"
Set DataGrid1.DataSource = Nothing
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
db.ConnectionString = "Provider=Microsoft.Jet.Oledb.4.0;Data Source=" & App.Path & "\GKADB.mdb;"
db.Open
db.CursorLocation = adUseClient
h = 0
End Sub

Private Sub Income1_Click()
Label1.Visible = True
Label2.Visible = True
Label3.Visible = False
Text2.Visible = False
DTPicker1.Visible = True
DTPicker2.Visible = True
Text1.Text = "1"
Set DataGrid1.DataSource = Nothing
lvButtons_H2.Visible = False
lvButtons_H1.Visible = True

End Sub

Private Sub IncomeF_Click()
Form5.Show vbModal
Set DataGrid1.DataSource = Nothing
End Sub

Private Sub IncomeR_Click()
Label1.Visible = True
Label2.Visible = True
Text2.Visible = False
Label3.Visible = False
DTPicker1.Visible = True
DTPicker2.Visible = True
Set DataGrid1.DataSource = Nothing
lvButtons_H2.Visible = True
lvButtons_H1.Visible = False
Text1.Text = "3"
Text4.Text = "Income"
Text5.Text = "Income Report"

End Sub

Private Sub lvButtons_H1_Click()

If Text1.Text = "1" Then
If DTPicker1.Value = DTPicker2.Value Then
Set rs = New ADODB.Recordset
rs.Open "select * from Income where Date_of_Receipt= #" & DTPicker1 & "# ", db, 3, 3
Set DataGrid1.DataSource = rs
Set jason1 = New ADODB.Recordset
jason1.Open "select sum(Amount) as money1 from Income where Date_of_Receipt= #" & DTPicker1 & "#", db, 3, 3
Text2.Text = "" & jason1!money1
Else
Set rs67 = New ADODB.Recordset
rs67.Open "select * from Income where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "#", db, 3, 3
Set DataGrid1.DataSource = rs67
Set jason1 = New ADODB.Recordset
jason1.Open "select sum(Amount) as money1 from Income where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "#", db, 3, 3
Text2.Text = "" & jason1!money1
End If
ElseIf Text1.Text = "2" Then
If DTPicker1.Value = DTPicker2.Value Then
Set rs45 = New ADODB.Recordset
rs45.Open "select * from Expenses where Date_of_Receipt=#" & DTPicker1 & "#", db, 3, 3
Set DataGrid1.DataSource = rs45
Set jason13 = New ADODB.Recordset
jason13.Open "select sum(Amount) as money14 from Expenses where Date_of_Receipt= #" & DTPicker1 & "#", db, 3, 3
Text2.Text = "" & jason13!money14
Else
Set rs34 = New ADODB.Recordset
rs34.Open "select * from Expenses where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "#", db, 3, 3
Set DataGrid1.DataSource = rs34
Set jason15 = New ADODB.Recordset
jason15.Open "select sum(Amount) as money16 from Expenses where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "#", db, 3, 3
Text2.Text = "" & jason15!money16
End If
End If
Text2.Visible = True
Label3.Visible = True



End Sub

Private Sub lvButtons_H2_Click()


Set rs1 = New ADODB.Recordset
If Text1.Text = "3" Then

If DTPicker1.Value = DTPicker2.Value Then
Set rs = New ADODB.Recordset
rs.Open "select * from Income where Date_of_Receipt= #" & DTPicker1 & "# ", db, 3, 3
Set DataGrid1.DataSource = rs
Set DataReport1.DataSource = rs
Set jason1 = New ADODB.Recordset
jason1.Open "select sum(Amount) as money1 from Income where Date_of_Receipt= #" & DTPicker1 & "#", db, 3, 3
Text2.Text = "" & jason1!money1
DataReport1.Sections("Section4").Controls("j1").Caption = DTPicker1.Value
DataReport1.Sections("Section4").Controls("j2").Caption = DTPicker2.Value
DataReport1.Sections("Section4").Controls("j3").Caption = Text4.Text
DataReport1.Sections("Section4").Controls("j4").Caption = Text2.Text
DataReport1.Sections("Section4").Controls("j5").Caption = Text5.Text
DataReport1.Show vbModal
Else
Set rs67 = New ADODB.Recordset
rs67.Open "select * from Income where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "#", db, 3, 3
Set DataGrid1.DataSource = rs67
Set DataReport1.DataSource = rs67
Set jason1 = New ADODB.Recordset
jason1.Open "select sum(Amount) as money1 from Income where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "#", db, 3, 3
Text2.Text = "" & jason1!money1
DataReport1.Sections("Section4").Controls("j1").Caption = DTPicker1.Value
DataReport1.Sections("Section4").Controls("j2").Caption = DTPicker2.Value
DataReport1.Sections("Section4").Controls("j3").Caption = Text4.Text
DataReport1.Sections("Section4").Controls("j4").Caption = Text2.Text
DataReport1.Sections("Section4").Controls("j5").Caption = Text5.Text
DataReport1.Show vbModal
End If
ElseIf Text1.Text = "4" Then
Set rs1 = New ADODB.Recordset
If DTPicker1.Value = DTPicker2.Value Then
Set rs45 = New ADODB.Recordset
rs45.Open "select * from Expenses where Date_of_Receipt=#" & DTPicker1 & "#", db, 3, 3
Set DataGrid1.DataSource = rs45
Set DataReport1.DataSource = rs45
Set jason13 = New ADODB.Recordset
jason13.Open "select sum(Amount) as money14 from Expenses where Date_of_Receipt= #" & DTPicker1 & "#", db, 3, 3
Text2.Text = "" & jason13!money14
DataReport1.Sections("Section4").Controls("j1").Caption = DTPicker1.Value
DataReport1.Sections("Section4").Controls("j2").Caption = DTPicker2.Value
DataReport1.Sections("Section4").Controls("j3").Caption = Text4.Text
DataReport1.Sections("Section4").Controls("j4").Caption = Text2.Text
DataReport1.Sections("Section4").Controls("j5").Caption = Text5.Text
DataReport1.Show vbModal
Else
Set rs34 = New ADODB.Recordset
rs34.Open "select * from Expenses where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "#", db, 3, 3
Set DataGrid1.DataSource = rs34
Set DataReport1.DataSource = rs34
Set jason15 = New ADODB.Recordset
jason15.Open "select sum(Amount) as money16 from Expenses where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "#", db, 3, 3
Text2.Text = "" & jason15!money16
DataReport1.Sections("Section4").Controls("j1").Caption = DTPicker1.Value
DataReport1.Sections("Section4").Controls("j2").Caption = DTPicker2.Value
DataReport1.Sections("Section4").Controls("j3").Caption = Text4.Text
DataReport1.Sections("Section4").Controls("j4").Caption = Text2.Text
DataReport1.Sections("Section4").Controls("j5").Caption = Text5.Text
DataReport1.Show vbModal
End If
ElseIf Text1.Text = "5" Then
If DTPicker1.Value = DTPicker2.Value Then
rs1.Open "Select * from income where Date_of_Receipt=#" & DTPicker1 & "#", db, 3, 3
Set DataReport2.DataSource = rs1
Set rs2 = New ADODB.Recordset
rs2.Open "select sum(Amount) as jason2 from income where Date_of_Receipt=#" & DTPicker1 & "#", db, 3, 3
Text6.Text = "" & rs2!jason2
Set rs3 = New ADODB.Recordset
rs3.Open "Select * from expenses where Date_of_Receipt=#" & DTPicker1 & "#", db, 3, 3
Set rs4 = New ADODB.Recordset
rs4.Open "select sum(Amount) as jason2 from Expenses where Date_of_Receipt=#" & DTPicker1 & "#", db, 3, 3
Text7.Text = "" & rs4!jason2
a = Val(Text6.Text & "") - Val(Text7.Text & "")
Text8.Text = a
DataReport2.Sections("Section4").Controls("j4").Caption = DTPicker1.Value
DataReport2.Sections("Section4").Controls("j5").Caption = DTPicker2.Value
DataReport2.Sections("Section4").Controls("j1").Caption = Text6.Text
DataReport2.Sections("Section4").Controls("j2").Caption = Text7.Text
DataReport2.Sections("Section4").Controls("j3").Caption = Text8.Text
DataReport2.Show vbModal
Else
rs1.Open "Select * from income where Date_of_Receipt between #" & DTPicker1 & "# and  #" & DTPicker2 & "# ", db, 3, 3
Set DataReport2.DataSource = rs1
Set rs2 = New ADODB.Recordset
rs2.Open "select sum(Amount) as jason2 from income where Date_of_Receipt between #" & DTPicker1 & "# and  #" & DTPicker2 & "#", db, 3, 3
Text6.Text = "" & rs2!jason2
Set rs3 = New ADODB.Recordset
rs3.Open "Select * from expenses where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "# ", db, 3, 3
Set rs4 = New ADODB.Recordset
rs4.Open "select sum(Amount) as jason2 from Expenses where Date_of_Receipt between #" & DTPicker1 & "# and #" & DTPicker2 & "# ", db, 3, 3
Text7.Text = "" & rs4!jason2
a = Val(Text6.Text & "") - Val(Text7.Text & "")
Text8.Text = a
DataReport2.Sections("Section4").Controls("j4").Caption = DTPicker1.Value
DataReport2.Sections("Section4").Controls("j5").Caption = DTPicker2.Value
DataReport2.Sections("Section4").Controls("j1").Caption = Text6.Text
DataReport2.Sections("Section4").Controls("j2").Caption = Text7.Text
DataReport2.Sections("Section4").Controls("j3").Caption = Text8.Text
DataReport2.Show vbModal
End If
End If
Text2.Visible = True
Label3.Visible = True

End Sub

Private Sub Net_Click()
Label1.Visible = True
Label2.Visible = True
DTPicker1.Visible = True
DTPicker2.Visible = True
Text2.Visible = False
Label3.Visible = False
lvButtons_H2.Visible = True
lvButtons_H1.Visible = False
Text1.Text = "5"
Text4.Text = "Net Income"
Text5.Text = "Net Income Report"
Set DataGrid1.DataSource = Nothing

End Sub

Private Sub Timer1_Timer()

c = "Global"
d = "Knowledge"
e = "Academy"

h = 1 + h
Text10.Text = h

If Text10.Text = "1" Then
Label4.Caption = c
ElseIf Text10.Text = "2" Then
Label4.Caption = d
ElseIf Text10.Text = "3" Then
Label4.Caption = e
ElseIf Text10.Text = "4" Then
Label4.Caption = "Global Knowledge Academy"
h = 0
End If
End Sub

Private Sub View1_Click()
Label1.Visible = False
Label2.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Text2.Visible = False
Text2.Text = ""
End Sub
Private Sub Form1_Click()
Label1.Visible = False
Label2.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Text2.Visible = False
Label3.Visible = False
Set DataGrid1.DataSource = Nothing
Text2.Text = ""

End Sub
Private Sub Print1_Click()
Label1.Visible = False
Label2.Visible = False
DTPicker1.Visible = False
DTPicker2.Visible = False
Text2.Visible = False
Label3.Visible = False
Text2.Text = ""

Set DataGrid1.DataSource = Nothing
End Sub


