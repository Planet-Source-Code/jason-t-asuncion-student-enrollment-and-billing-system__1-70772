VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form Form2 
   Caption         =   "Expenses Form"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7155
   ScaleWidth      =   8145
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   4080
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   2535
      Left            =   600
      TabIndex        =   3
      Top             =   4080
      Width           =   3735
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5880
      Top             =   6600
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4800
      TabIndex        =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   855
      Left            =   5280
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1508
      Caption         =   "Save Expenses Stament"
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   4194304
      CalendarTitleBackColor=   49344
      CalendarTrailingForeColor=   -2147483630
      Format          =   20709377
      CurrentDate     =   39175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2520
      Top             =   6720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OR Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   10
      Top             =   3720
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      Top             =   3720
      Width           =   2175
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
      TabIndex        =   7
      Top             =   1440
      Width           =   6855
   End
   Begin VB.Image Image2 
      Height          =   1140
      Left            =   3502
      Picture         =   "Form2.frx":2BEB3
      Top             =   120
      Width           =   1140
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Expenses Form"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1425
      TabIndex        =   6
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   7185
      Left            =   0
      Picture         =   "Form2.frx":302A5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8130
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Dim b As Integer

Private cn As New ADODB.Connection



Private Sub Form_Load()
Set cn = New ADODB.Connection
cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\GKADB.mdb"
cn.Open

a = 0
End Sub

Private Sub lvButtons_H1_Click()
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
If Val(Text2.Text) = 0 Then
MsgBox "Amount is 0 or invalid type of Data", vbInformation, "GKA Database"
ElseIf Val(Text2.Text) > 0 Then
cn.Execute "insert into expenses (Or_number,Date_of_Receipt,Particulars,Amount) Values ('" & Text1.Text & "','" & DTPicker1 & "','" & Text3.Text & "','" & Text2.Text & "')"
MsgBox "Data Income Statement has been saved!!", vbInformation, "GKA Database"
Unload Me
End If
Else
MsgBox "All Fields are required!!", vbInformation, "GKA Database"
End If
End Sub

Private Sub Timer1_Timer()
Dim c As String
Dim d As String
Dim e As String
c = "Global"
d = "Knowledge"
e = "Academy"

a = 1 + a
Text4.Text = a

If Text4.Text = "1" Then

Label5.Caption = c
ElseIf Text4.Text = "2" Then

Label5.Caption = d
ElseIf Text4.Text = "3" Then

Label5.Caption = e
ElseIf Text4.Text = "4" Then

Label5.Caption = "Global Knowledge Academy"
a = 0
End If

End Sub

