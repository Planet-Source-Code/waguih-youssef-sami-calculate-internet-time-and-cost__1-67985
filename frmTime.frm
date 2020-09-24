VERSION 5.00
Begin VB.Form frmTime 
   Caption         =   "Calculate Time"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3165
   Icon            =   "frmTime.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Status"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmdTray 
      Caption         =   "To Tray"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   855
   End
   Begin VB.Timer TimerStart 
      Interval        =   1000
      Left            =   2640
      Top             =   1320
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Timer TimerEnd 
      Interval        =   1000
      Left            =   2160
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Time"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "End"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyTime2, MyTime1

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Dim sConnType As String * 255

Private Type NOTIFYICONDATA
    cbSize As Long              'Size of the structure
    hwnd As Long                'Window handle of the icon's owner
    uId As Long                 'Unique identifier, use for multiple icons
    uFlags As Long              'Flags
    uCallBackMessage As Long    'Window message (WM) sent to the icon's owner
    hIcon As Long               'Handle of the icon to use (use VB's Form.Icon property)
    szTip As String * 64        'ToolTip textType
End Type

'Shell_NotifyIcon messages
Private Const NIM_ADD = &H0         'Add to tray
Private Const NIM_MODIFY = &H1      'Change Icon
Private Const NIM_DELETE = &H2      'Delete Icon
 
'NotifyIconData uFlags parameters, specify which members of
'NOTIFYICONDATA are valid and should be used by Shell_NotifyIcon
'AND these together in uFlags member
Private Const NIF_MESSAGE = &H1     'Honor uCallbackMessage member
Private Const NIF_ICON = &H2        'Honor hIcon member
Private Const NIF_TIP = &H4         'Honor szInfo member
 

'Window messages Shell_NotifyIcon sends to your app
Private Const WM_MOUSEMOVE = &H200       'MouseMove message
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
 
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
 
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
 
Private m_IconData As NOTIFYICONDATA
Private m_lngLastMessage As Long

Private Sub cmdTray_Click()
frmTime.WindowState = 1
End Sub

Private Sub Command1_Click()
frmDetails.Show
Me.Hide
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim lRet As Long
    Dim lMessage As Long
   
    If Me.ScaleMode = vbPixels Then
        'No conversion needed
        lMessage = X
    Else
        'VB mangled X to convert it from Pixels to Twips
        lMessage = X / Screen.TwipsPerPixelX
    End If
   
    
    Select Case lMessage
          
    Case WM_LBUTTONDBLCLK
        'Double-click, restore the form
        Result = SetForegroundWindow(Me.hwnd)
        'MsgBox "Double-Click"
        Me.WindowState = vbNormal
        Me.Show
        
                  
    End Select
   
    m_lngLastMessage = lMessage
End Sub
Private Sub Form_Resize()
    
    If Me.WindowState = vbMinimized Then
    Me.Hide
    End If
End Sub
 
Private Sub Form_Unload(Cancel As Integer)
    'Remove the icon from the tray
    Shell_NotifyIcon NIM_DELETE, m_IconData
    End
End Sub

Private Sub Form_Load()
frmTime.Icon = LoadPicture(App.Path & "\Clock06.ico")
Text1.Text = "Start"
TimerStart.Enabled = True   'enabled but its functions will not untill connection is true
    
With m_IconData
        .cbSize = Len(m_IconData)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE    'Use MouseMove message
        .hIcon = Me.Icon
        .szTip = "Waguih Alarm" & vbNullChar
              
    End With
Shell_NotifyIcon NIM_ADD, m_IconData
End Sub


Private Sub TimerStart_Timer()
Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret = 1 Then
        'MsgBox "You are connected to Internet via a " & sConnType, vbInformation
   Text1.Text = Time
    

    If Hour(Time) > 12 Then
    MyTime1 = CDate(Time - CDate("12:00:00"))
    Else
    MyTime1 = Time
    End If

TimerStart.Enabled = False
 Exit Sub
    'Else
        'MsgBox "You are not connected to internet", vbInformation
       End If
End Sub

Private Sub TimerEnd_Timer()
Text2.Text = Time
If Hour(Time) = 0 Then
MyTime2 = CDate(Time + CDate("12:00:00"))
Else
    If Hour(Time) > 12 Then
    MyTime2 = CDate(Time - CDate("12:00:00"))
    Else
    MyTime2 = Time
    End If
End If

Dim Ret As Long
    Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret = 0 Then
        If Text1.Text <> "Start" Then
       
        MyStop
        End If
    End If

End Sub


Public Sub MyStop()
TimerEnd.Enabled = False

MyTime1 = Hour(MyTime1) * 60 + Minute(MyTime1)
MyTime2 = Hour(MyTime2) * 60 + Minute(MyTime2)
Text3.Text = MyTime2 - MyTime1

Dim NewTime
Dim MyFile, MyFile2, MyFile3
Dim MyString
MyFile = App.Path & "\MyTime.txt"
MyFile2 = App.Path & "\TimeDetail.txt"
MyFile3 = App.Path & "\Cost.txt"

Open MyFile For Input As #1
Line Input #1, MyString
NewTime = CSng(MyString) + CSng(Text3.Text)
Close #1

Open MyFile For Output As #1
Print #1, NewTime
Close #1


Open MyFile2 For Append As #1
'Print #1, vbCrLf
Print #1, Date & " = " & CSng(Text3.Text) & " minutes"
Close #1

Dim MyCost
MyCost = 1.24 * NewTime / 60
MyCost = Format(MyCost, "##.00")

Open MyFile3 For Output As #1
Print #1, MyCost & " LE"
Close #1

Dim Z
Z = MsgBox("Total Internet Time= " & NewTime & " Minutes", vbOKCancel)
If Z = vbOK Then
frmDetails.Show
Me.Hide
End If

End Sub
