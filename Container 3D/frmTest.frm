VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "3D Container Test Form"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   4125
   ScaleWidth      =   5355
   StartUpPosition =   3  'Windows Default
   Begin prjContainer3D.Container3D Container3D8 
      Height          =   1410
      Left            =   2565
      Top             =   2625
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   2487
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   7
         Left            =   1335
         TabIndex        =   8
         Top             =   1050
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   6
         Left            =   1335
         TabIndex        =   7
         Top             =   740
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   5
         Left            =   1335
         TabIndex        =   6
         Top             =   430
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   4
         Left            =   1335
         TabIndex        =   5
         Top             =   120
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1050
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   740
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   430
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   870
      End
   End
   Begin prjContainer3D.Container3D Container3D7 
      Height          =   810
      Left            =   2565
      Top             =   1650
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   1429
      BackColor       =   65535
   End
   Begin prjContainer3D.Container3D Container3D6 
      Height          =   1380
      Left            =   2550
      Top             =   150
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   2434
      AutoSize        =   1
      Begin VB.TextBox Text1 
         Height          =   1320
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   0
         Text            =   "frmTest.frx":0000
         Top             =   30
         Width           =   2610
      End
   End
   Begin prjContainer3D.Container3D Container3D5 
      Height          =   615
      Left            =   135
      Top             =   3405
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1085
      CtrlEffect      =   5
   End
   Begin prjContainer3D.Container3D Container3D4 
      Height          =   615
      Left            =   135
      Top             =   2595
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1085
      CtrlEffect      =   4
   End
   Begin prjContainer3D.Container3D Container3D2 
      Height          =   615
      Left            =   135
      Top             =   975
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1085
      CtrlEffect      =   2
   End
   Begin prjContainer3D.Container3D Container3D1 
      Height          =   615
      Left            =   135
      Top             =   165
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1085
   End
   Begin prjContainer3D.Container3D Container3D3 
      Height          =   615
      Left            =   135
      Top             =   1785
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   1085
      CtrlEffect      =   3
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Container3D7_MouseEnter()
  'Example of Mouse Enter Event
  UpdateColor Container3D7
End Sub

Private Sub Container3D7_MouseExit()
  'Example of Mouse Exit Event
  UpdateColor Container3D7
End Sub

Private Sub UpdateColor(objCnt3D As Container3D)
  'Just a silly thing I did to see the Mouse Enter & Exit Events work!
  Randomize
  Dim myColor As Integer
  
NoGood:
  myColor = Int(15 * Rnd)
  If QBColor(myColor) = objCnt3D.BackColor Then GoTo NoGood
  objCnt3D.BackColor = QBColor(myColor)
End Sub
