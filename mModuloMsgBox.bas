Attribute VB_Name = "ModuloMsgBox"
' MiMsgBox v. 1.1.5
' Visual Design inspired in Kevin Figg's custom MsgBox
' but all the code as new, by me.
' Super easy to use Custom MsgBox.
' Just add one form and one module in your Proyect and call one Function.
' If you like, change the form design with your preferences... and i'ts all.
' Easy, fast, nice.
' Proyect Start - Dec/15/2003
' Actual Revision - Jan/24/2004
' Comments, sugestions, etc. are welcome.
' · Use the gonchuki ChameleonButton and 13 button styles.
' · You can set the number of buttons (1, 2, 3 or none)
' · Button(s) AutoCentering
' · Any text for any button
' · Self hiding MsgBox in x seconds
' Written by José Luis Farías.
' Chile 1446 - Salto - Uruguay - CP 50.000
' JoseloFarias[at]adinet.com.uy
' ¡¡¡Vamo' arriba Uruguay, carajo!!!
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
' ¡PLEASE!, if you use this Code sendme your Name and Country
' And if you like, emailme a program copy (source code if better)
'*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*
Option Explicit 'is better
Public Seconds As Byte
Public ButtonClicked As Byte 'Remember the clicked button number
Public AutoHide As Boolean 'AutoHide On
Public Enum enIcons 'Enumerate Icons
    Computer = 1
    Critical = 2
    Exclamation = 3
    FontFolder = 4
    Hardware = 5
    Information = 6
    Internet = 7
    Power = 8
    Question = 9
    Search = 10
    Security = 11
    Sound = 12
    Users = 13
    Devil = 14
    DragonBall = 15
    Face = 16
End Enum
Public Enum enButtons 'Enumerate number of buttons
    One = 1
    Two = 2
    Three = 3
End Enum
Public Sub ShowMsg(Text As String, Buttons As enButtons, Optional Btn1Text As String, Optional Btn2Text As String, Optional Btn3Text As String, Optional Title As String, Optional Header As String, Optional ByVal Icon As enIcons = 11, Optional SecondsToKillme As Byte = 0, Optional ByVal ButtonSkin As ButtonTypes = 9)
    On Error Resume Next 'I'm sorry...
    Load frmMensaje 'Load the Form
    frmMensaje.lblText = Text 'Set the text
    frmMensaje.lblTitle = Title 'Set the title
    frmMensaje.lblHeader = Header 'Set the header
    'Set the Buttons Style
    frmMensaje.cmdCommand1(0).ButtonType = ButtonSkin
    frmMensaje.cmdCommand1(1).ButtonType = ButtonSkin
    frmMensaje.cmdCommand1(2).ButtonType = ButtonSkin
    'Only valid numbers for Icon setting
    If Icon > frmMensaje.ImageList1.ListImages.Count Then Exit Sub
    'Set the Icon
        frmMensaje.Image1.Picture = frmMensaje.ImageList1.ListImages(Icon).Picture
    'Show only 1 button
    If Buttons = 1 Then
        frmMensaje.cmdCommand1(0).Visible = True
        frmMensaje.cmdCommand1(1).Visible = False
        frmMensaje.cmdCommand1(2).Visible = False
    'Center the button in Form
        frmMensaje.cmdCommand1(0).Move (frmMensaje.ScaleWidth / 2) - (frmMensaje.cmdCommand1(0).Width / 2)
    'Show 2 buttons
    ElseIf Buttons = 2 Then
        frmMensaje.cmdCommand1(0).Visible = True
        frmMensaje.cmdCommand1(1).Visible = True
        frmMensaje.cmdCommand1(2).Visible = False
    'Center in the Form
        frmMensaje.cmdCommand1(0).Move (frmMensaje.ScaleWidth / 2) - frmMensaje.cmdCommand1(0).Width - 50
        frmMensaje.cmdCommand1(1).Move (frmMensaje.ScaleWidth / 2) + 75
    'Show 3 buttons
    ElseIf Buttons = 3 Then
        frmMensaje.cmdCommand1(0).Visible = True
        frmMensaje.cmdCommand1(1).Visible = True
        frmMensaje.cmdCommand1(2).Visible = True
    'and center...
        frmMensaje.cmdCommand1(0).Move (frmMensaje.ScaleWidth / 2) - (frmMensaje.cmdCommand1(0).Width / 2) - frmMensaje.cmdCommand1(0).Width - 100
        frmMensaje.cmdCommand1(1).Move (frmMensaje.ScaleWidth / 2) - (frmMensaje.cmdCommand1(0).Width / 2)
        frmMensaje.cmdCommand1(2).Move (frmMensaje.ScaleWidth / 2) + (frmMensaje.cmdCommand1(0).Width / 2) + 100
    Else
    'Hide all buttons
        frmMensaje.cmdCommand1(0).Visible = False
        frmMensaje.cmdCommand1(1).Visible = False
        frmMensaje.cmdCommand1(2).Visible = False
    End If
    'Set the buttons Captions
    frmMensaje.cmdCommand1(0).Caption = Btn1Text
    frmMensaje.cmdCommand1(1).Caption = Btn2Text
    frmMensaje.cmdCommand1(2).Caption = Btn3Text
    'AutoHide Form
    If SecondsToKillme > 0 Then
        frmMensaje.Timer.Enabled = True
        frmMensaje.Timer.Interval = SecondsToKillme * 1000 'milliseconds to seconds
        AutoHide = True 'set the flag
    End If
    'No buttons and no AutoHiding... how to close the form?
    If Buttons = 0 Then
        frmMensaje.Timer.Interval = 2500 'With Autohide
    End If
    frmMensaje.Show vbModal 'Now, show the MsgBox
End Sub
