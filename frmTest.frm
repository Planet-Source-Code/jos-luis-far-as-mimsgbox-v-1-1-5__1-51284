VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00E0E0E0&
   Caption         =   "MiMsgBox Test Form"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   4950
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSeconds 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   20
      Top             =   3720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Killme in xx seconds"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   19
      Top             =   3720
      Width           =   1815
   End
   Begin Proyecto_MiMsgBox.ChameleonButton cmdView 
      Default         =   -1  'True
      Height          =   495
      Left            =   1800
      TabIndex        =   18
      Top             =   4200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   "&View Example"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   13160664
      BCOLO           =   13160664
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmTest.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox Btn3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Text            =   "&Nice Code =)"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Btn2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Text            =   "&Cancel"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox Btn1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   9
      Text            =   "&Ok"
      Top             =   3120
      Width           =   1455
   End
   Begin VB.ComboBox cmbButtons 
      Height          =   315
      ItemData        =   "frmTest.frx":001C
      Left            =   2040
      List            =   "frmTest.frx":0029
      TabIndex        =   8
      Text            =   "3"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.ComboBox cmbButtonStyle 
      Height          =   315
      ItemData        =   "frmTest.frx":0036
      Left            =   3600
      List            =   "frmTest.frx":0061
      TabIndex        =   7
      Text            =   "Office XP"
      Top             =   2400
      Width           =   1575
   End
   Begin VB.ComboBox cmbIcon 
      Height          =   315
      ItemData        =   "frmTest.frx":0100
      Left            =   480
      List            =   "frmTest.frx":0134
      TabIndex        =   6
      Text            =   "Security"
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox txtText 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Text            =   "Incorrect Password. Please Try Again."
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox txtHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Text            =   "Incorrect Password"
      Top             =   960
      Width           =   4695
   End
   Begin VB.TextBox txtTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   480
      TabIndex        =   0
      Text            =   "Security Problem"
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buttons Style"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   17
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "# of Buttons"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Icon"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button 3 Text"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button 2 Text"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label lblbtn1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Button 1 Text"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   10
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MsgBox Text"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MsgBox Header"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   720
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MsgBox Title"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Check1_Click()
    If txtSeconds.Visible = False Then
        txtSeconds.Visible = True
        txtSeconds.SetFocus
    Else
        txtSeconds.Visible = False
        txtSeconds = ""
    End If
End Sub
Private Sub cmdView_Click()
    Dim Icon As Byte, ButtonStyle As Byte
    Dim Text As String
    Select Case cmbIcon
        Case "Computer"
            Icon = 1
        Case "Critical"
            Icon = 2
        Case "Exclamation"
            Icon = 3
        Case "FontFolder"
            Icon = 4
        Case "Hardware"
            Icon = 5
        Case "Information"
            Icon = 6
        Case "Internet"
            Icon = 7
        Case "Power"
            Icon = 8
        Case "Question"
            Icon = 9
        Case "Search"
            Icon = 10
        Case "Security"
            Icon = 11
        Case "Sound"
            Icon = 12
        Case "Users"
            Icon = 13
        Case "Devil"
            Icon = 14
        Case "DragonBall"
            Icon = 15
        Case "Face"
            Icon = 16
    End Select
    
    Select Case cmbButtonStyle
        Case "Windows 16-bit"
            ButtonStyle = 1
        Case "Windows 32-bit"
            ButtonStyle = 2
        Case "Windows XP"
            ButtonStyle = 3
        Case "Mac"
            ButtonStyle = 4
        Case "Java metal"
            ButtonStyle = 5
        Case "Netscape 6"
            ButtonStyle = 6
        Case "Simple Flat"
            ButtonStyle = 7
        Case "Flat Highlight"
            ButtonStyle = 8
        Case "Office XP"
            ButtonStyle = 9
        Case "Transparent"
            ButtonStyle = 11
        Case "3D Hover"
            ButtonStyle = 12
        Case "Oval Flat"
            ButtonStyle = 13
        Case "KDE 2"
            ButtonStyle = 14
    End Select
    If Val(txtSeconds) <> 0 Then
        Text = txtText & vbCrLf & "AutoHide in " & txtSeconds & " seconds"
    Else
        Text = txtText
    End If
    ShowMsg Text, Val(cmbButtons), Btn1, Btn2, Btn3, txtTitle, txtHeader, Icon, Val(txtSeconds), ButtonStyle
    If AutoHide = False Then
        MsgBox "Click on Button # " & ButtonClicked
    End If
End Sub
' Example of buttons click programming
'    Select Case ButtonClicked
'        Case 1
'            blah , blah, blah
'        Case 2
'            blah , blah, blah
'        Case 3
'            blah , blah, blah
'    End Select

