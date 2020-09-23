VERSION 5.00
Object = "{FEA88812-E6DE-4A75-A946-124DDC4C195A}#1.0#0"; "LED7Seg.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LED7Seg Control - You set the color and the character it uses, and voila! This is a control array to illustrate the capabilities."
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   0
      Left            =   10860
      TabIndex        =   0
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   1
      Left            =   10140
      TabIndex        =   1
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   2
      Left            =   9420
      TabIndex        =   2
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   3
      Left            =   8700
      TabIndex        =   3
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   4
      Left            =   7980
      TabIndex        =   4
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   5
      Left            =   7260
      TabIndex        =   5
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   6
      Left            =   6540
      TabIndex        =   6
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   7
      Left            =   5820
      TabIndex        =   7
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   8
      Left            =   5100
      TabIndex        =   8
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   9
      Left            =   4380
      TabIndex        =   9
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   10
      Left            =   3660
      TabIndex        =   10
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   11
      Left            =   2940
      TabIndex        =   11
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   12
      Left            =   2220
      TabIndex        =   12
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   13
      Left            =   1500
      TabIndex        =   13
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   14
      Left            =   780
      TabIndex        =   14
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
   Begin LED7Segctl.LED7Seg LED7Seg1 
      Height          =   1020
      Index           =   15
      Left            =   120
      TabIndex        =   15
      Top             =   180
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1799
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim i As Integer
    Const Alpha As String = "fedcba9876543210"
    For i = 0 To LED7Seg1.UBound
        LED7Seg1(i).DrawLED RGB(95, 0, 0), RGB(255, 0, 0), Mid$(Alpha, i + 1, 1)
        If i = 0 Then
            LED7Seg1(i).DrawLED RGB(95, 0, 0), RGB(255, 0, 0), Mid$(Alpha, i + 1, 1) & "."
        End If
    Next i
End Sub
