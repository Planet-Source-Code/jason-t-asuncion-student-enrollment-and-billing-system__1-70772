VERSION 5.00
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   5805
   LinkTopic       =   "Form4"
   ScaleHeight     =   5070
   ScaleWidth      =   5805
   StartUpPosition =   1  'CenterOwner
   Begin lvButton.lvButtons_H lvButtons_H1 
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "OK"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© Copyright 2008 ASoft"
      Height          =   255
      Left            =   975
      TabIndex        =   13
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Developer: Jason T. Asuncion"
      Height          =   255
      Left            =   1455
      TabIndex        =   12
      Top             =   4680
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form4.frx":0000
      ForeColor       =   &H00000080&
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   5520
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5520
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Just contact me at 09165546245 or email me at jadsummer99@yahoo.com."
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3000
      Width           =   5535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MIS"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "POS"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2400
      Width           =   3375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Biling System"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Accounting System"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory System"
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   1680
      Width           =   3495
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "If you want more program like the following:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form4.frx":00B5
      ForeColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   5295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "And Reports Generating System"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Computerize Accounts Filing "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   5115
      Left            =   0
      Picture         =   "Form4.frx":0142
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lvButtons_H1_Click()
Unload Me
End Sub
