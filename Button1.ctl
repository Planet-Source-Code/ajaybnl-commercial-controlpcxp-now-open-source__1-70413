VERSION 5.00
Begin VB.UserControl button 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1260
   ScaleHeight     =   420
   ScaleWidth      =   1260
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit ':( Line inserted by Formatter
Public Event Click()
Public Event Mouseup()
Private Enab As Boolean ':( Missing Scope
Private Skinned As Boolean
Public Property Get Caption() As String
    Caption = Label1.Caption
End Property
Public Property Let Caption(Text As String)
    Label1.Caption = Text
    Label1.Left = (UserControl.Width - Label1.Width) / 2
    Label1.Top = (UserControl.Height - Label1.Height) / 2
    ReDraw False
End Property
Public Property Let FontSize(Size As Long)
    Label1.FontSize = Size
End Property
Public Property Let FontColor(Color As Long)
    Label1.ForeColor = Color
End Property
Public Property Let Backcolor(Color As Long)
    UserControl.Backcolor = Color
End Property
Public Property Let Enabled(En As Boolean)
    Enab = En
End Property
Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseDown 1, 0, 0, 0
End Sub
Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp 1, 0, 0, 0
End Sub
Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enab = True Then UserControl_MouseDown Button, 1, 1, 1
':( Expand Structure
End Sub
Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent Mouseup
End Sub
Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseUp Button, 1, 1, 1
End Sub
Private Sub UserControl_Initialize()
    Label1.Left = (UserControl.Width - Label1.Width) / 2
    Label1.Top = (UserControl.Height - Label1.Height) / 2
    UserControl.Backcolor = &H8000000F
    Label1.Backcolor = &H8000000F
    ReDraw False
End Sub
Private Sub UserControl_InitProperties()
    Enab = True
    Skinned = True
    ReDraw False
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Enab = True Then ReDraw True ':( Expand Structure
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent Mouseup
End Sub
Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReDraw False
    If Enab = True Then RaiseEvent Click   ':( Expand Structure
End Sub
Private Sub UserControl_Paint()
    ReDraw False
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        Label1.Caption = .ReadProperty("Caption", "Button")
        UserControl.Height = .ReadProperty("Height", UserControl.Height)
        UserControl.Width = .ReadProperty("Width", UserControl.Width)
        Enab = .ReadProperty("Enabled", "True")
    End With
    Label1.Left = (UserControl.Width - Label1.Width) / 2
    Label1.Top = (UserControl.Height - Label1.Height) / 2
End Sub
Private Sub UserControl_Resize()
    Label1.Left = (UserControl.Width - Label1.Width) / 2
    Label1.Top = (UserControl.Height - Label1.Height) / 2
    ReDraw False
End Sub
Private Sub UserControl_IdeOK()
    ReDraw False
End Sub
Private Sub ReDraw(Press As Boolean)
    UserControl.Cls
    If Press = True Then ':( Remove Pleonasm
        UserControl.ForeColor = &H808080
        UserControl.Line (0, 0)-(UserControl.Width, 0)
        UserControl.Line (0, 0)-(0, UserControl.Height)
        UserControl.ForeColor = &HE0E0E0
        UserControl.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height)
        UserControl.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15)
    Else
        UserControl.ForeColor = &HE0E0E0
        UserControl.Line (0, 0)-(UserControl.Width, 0)
        UserControl.Line (0, 0)-(0, UserControl.Height)
        UserControl.ForeColor = &H808080
        UserControl.Line (UserControl.Width - 15, 0)-(UserControl.Width - 15, UserControl.Height)
        UserControl.Line (0, UserControl.Height - 15)-(UserControl.Width, UserControl.Height - 15)
    End If
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "Enabled", Enab, "True"
    PropBag.WriteProperty "Caption", Label1.Caption, "Button"
    PropBag.WriteProperty "Height", UserControl.Height, UserControl.Height
    PropBag.WriteProperty "Width", UserControl.Width, UserControl.Width
End Sub
':)Code Fixer V3.0.9 (11/15/2006 12:12:21 PM) 5 + 173 = 178 Lines Thanks Ulli for inspiration and lots of code.
