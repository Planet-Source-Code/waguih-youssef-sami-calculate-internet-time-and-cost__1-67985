VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDetails 
   BackColor       =   &H8000000A&
   Caption         =   "Internet Connection Time By Minutes"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Back"
      Height          =   615
      Left            =   480
      Picture         =   "frmDetails.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox RTBox3 
      Height          =   1335
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      _Version        =   393217
      TextRTF         =   $"frmDetails.frx":0442
   End
   Begin RichTextLib.RichTextBox RTBox2 
      Height          =   4575
      Left            =   3480
      TabIndex        =   1
      Top             =   480
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   8070
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmDetails.frx":0573
   End
   Begin RichTextLib.RichTextBox RTBox1 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   2355
      _Version        =   393217
      TextRTF         =   $"frmDetails.frx":06A4
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost uptill Now"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Details By Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Internet Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   2775
   End
End
Attribute VB_Name = "frmDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
frmTime.Show
Me.Hide
End Sub

Private Sub Form_Load()
frmDetails.BackColor = RGB(0, 175, 200)

Dim MyFile1, MyFile2, MyFile3
MyFile1 = App.Path & "\MyTime.txt"
MyFile2 = App.Path & "\TimeDetail.txt"
MyFile3 = App.Path & "\Cost.txt"
RTBox1.LoadFile MyFile1
RTBox2.LoadFile MyFile2
RTBox3.LoadFile MyFile3
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
