VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMixer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NexIRC - Mixer"
   ClientHeight    =   2415
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9855
   DrawWidth       =   4
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMixer.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   9855
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   9270
      Top             =   225
   End
   Begin VB.Frame Frame1 
      Caption         =   "VU Meter"
      Height          =   2310
      Left            =   9900
      TabIndex        =   92
      Top             =   45
      Visible         =   0   'False
      Width           =   960
      Begin VB.TextBox VuLeftLevel 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   94
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   285
      End
      Begin VB.TextBox VuRightLevel 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   540
         TabIndex        =   93
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   285
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   1590
         Left            =   180
         TabIndex        =   1
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2805
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Max             =   3275
         Orientation     =   1
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   1590
         Left            =   540
         TabIndex        =   0
         Top             =   630
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   2805
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   1
         Max             =   3275
         Orientation     =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   " BASS"
      Height          =   2325
      Index           =   11
      Left            =   45
      TabIndex        =   88
      Top             =   45
      Width           =   705
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1770
         Index           =   11
         Left            =   45
         TabIndex        =   108
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   3122
         _Version        =   327682
         Orientation     =   1
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   1
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   106
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   11
         Left            =   90
         TabIndex        =   89
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   345
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   11
         Left            =   90
         TabIndex        =   91
         Top             =   2340
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "TREB"
      Height          =   2325
      Index           =   10
      Left            =   765
      TabIndex        =   86
      Top             =   45
      Width           =   705
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   10
         Left            =   45
         TabIndex        =   90
         Top             =   2340
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   105
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   10
         Left            =   90
         TabIndex        =   87
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   525
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1770
         Index           =   10
         Left            =   45
         TabIndex        =   109
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   3122
         _Version        =   327682
         Orientation     =   1
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   1
         TickFrequency   =   8000
         Value           =   32768
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "PCSPK"
      Height          =   2325
      Index           =   9
      Left            =   8370
      TabIndex        =   78
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   104
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   9
         Left            =   90
         TabIndex        =   83
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   80
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   9
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   82
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   9
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   81
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   9
         Left            =   45
         TabIndex        =   84
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   9
         Left            =   45
         TabIndex        =   85
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   9
         Left            =   90
         TabIndex        =   79
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Visible         =   0   'False
         Width           =   555
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "12SIN"
      Height          =   2325
      Index           =   8
      Left            =   7605
      TabIndex        =   70
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   103
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   8
         Left            =   90
         TabIndex        =   75
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   72
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   8
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   8
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   8
         Left            =   90
         TabIndex        =   71
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Visible         =   0   'False
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   8
         Left            =   30
         TabIndex        =   76
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   8
         Left            =   45
         TabIndex        =   77
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "MIDI"
      Height          =   2325
      Index           =   7
      Left            =   6840
      TabIndex        =   62
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   102
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   90
         TabIndex        =   67
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   64
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   7
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   7
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   65
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   7
         Left            =   90
         TabIndex        =   63
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   7
         Left            =   30
         TabIndex        =   68
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   7
         Left            =   45
         TabIndex        =   69
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "TAD"
      Height          =   2325
      Index           =   6
      Left            =   6075
      TabIndex        =   54
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   101
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   6
         Left            =   90
         TabIndex        =   59
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   56
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   6
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   58
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   6
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   57
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   6
         Left            =   90
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   6
         Left            =   30
         TabIndex        =   60
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   6
         Left            =   45
         TabIndex        =   61
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "WAVE"
      Height          =   2325
      Index           =   5
      Left            =   5310
      TabIndex        =   46
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   100
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   51
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   48
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   5
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   50
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   5
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   5
         Left            =   60
         TabIndex        =   52
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   5
         Left            =   45
         TabIndex        =   53
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "AUX"
      Height          =   2325
      Index           =   4
      Left            =   4560
      TabIndex        =   38
      Top             =   60
      Width           =   705
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   99
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   43
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   40
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   4
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   4
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   41
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   4
         Left            =   90
         TabIndex        =   39
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   4
         Left            =   30
         TabIndex        =   44
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   4
         Left            =   45
         TabIndex        =   45
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "MIC"
      Height          =   2325
      Index           =   3
      Left            =   3780
      TabIndex        =   30
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   98
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   35
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   32
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   3
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   3
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   31
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   3
         Left            =   30
         TabIndex        =   36
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   3
         Left            =   45
         TabIndex        =   37
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "LINEIN"
      Height          =   2325
      Index           =   2
      Left            =   3015
      TabIndex        =   22
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   97
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   90
         TabIndex        =   27
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   24
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   2
         Left            =   45
         TabIndex        =   28
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   2
         Left            =   45
         TabIndex        =   29
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "CDVOL"
      Height          =   2325
      Index           =   1
      Left            =   2250
      TabIndex        =   14
      Top             =   45
      Width           =   705
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   1800
         Width           =   240
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   16
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   1
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   1
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   1
         Left            =   30
         TabIndex        =   20
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   21
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame VolFrame 
      Caption         =   "MAST"
      Height          =   2325
      Index           =   0
      Left            =   1500
      TabIndex        =   6
      Top             =   45
      Width           =   705
      Begin VB.TextBox TXTVol 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   90
         TabIndex        =   95
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox TXTVolumeControl 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "32768"
         Top             =   210
         Width           =   555
      End
      Begin VB.PictureBox Picture11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   390
         ScaleHeight     =   270
         ScaleWidth      =   180
         TabIndex        =   8
         ToolTipText     =   "Mute"
         Top             =   1680
         Width           =   240
         Begin VB.OptionButton MuteOff 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Mute"
            Top             =   135
            Width           =   195
         End
         Begin VB.OptionButton MuteOn 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Unmute"
            Top             =   0
            Value           =   -1  'True
            Width           =   195
         End
      End
      Begin VB.CheckBox SBMLink 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   7
         Top             =   1800
         Width           =   240
      End
      Begin ComctlLib.Slider VolumeControl 
         Height          =   1230
         Index           =   0
         Left            =   30
         TabIndex        =   12
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2170
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   2
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin MSComctlLib.Slider BalanceControl 
         Height          =   225
         Index           =   0
         Left            =   45
         TabIndex        =   13
         Top             =   2070
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   397
         _Version        =   393216
         LargeChange     =   10
         SmallChange     =   5
         Min             =   -100
         Max             =   100
         TickStyle       =   3
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "ALL"
      Height          =   2325
      Left            =   9135
      TabIndex        =   2
      Top             =   45
      Width           =   705
      Begin VB.CommandButton Command1 
         Caption         =   "----I----"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   150
         Left            =   90
         TabIndex        =   107
         ToolTipText     =   "Balance = 0"
         Top             =   2115
         Width           =   555
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   3
         ToolTipText     =   "Links all selected faders to SBM fader"
         Top             =   1830
         Width           =   225
      End
      Begin ComctlLib.Slider Slider6 
         Height          =   1260
         Left            =   45
         TabIndex        =   4
         Top             =   510
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   2223
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   200
         Max             =   65535
         SelStart        =   32768
         TickStyle       =   1
         TickFrequency   =   8000
         Value           =   32768
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000B&
         BackStyle       =   0  'Transparent
         Caption         =   "ALL"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   345
         TabIndex        =   5
         Top             =   1830
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmMixer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim volR As Long
Dim volL As Long
Dim volume As Long
Dim mute As MIXERCONTROL
Dim unmute As MIXERCONTROL
Dim ONOFF As MIXERCONTROL
Dim hmixer As Long
Dim VolCtrl As MIXERCONTROL
Dim WavCtrl As MIXERCONTROL
Dim CDVol As MIXERCONTROL
Dim LineVol As MIXERCONTROL
Dim MICROPHONE As MIXERCONTROL
Dim PCSPEAKER As MIXERCONTROL
Dim AUXVol As MIXERCONTROL
Dim TELEPHONE As MIXERCONTROL
Dim MIDIVol As MIXERCONTROL
Dim I25InVol As MIXERCONTROL
Dim Treble As MIXERCONTROL
Dim Bass As MIXERCONTROL
Dim rc As Long
Dim ok As Boolean

Private Sub Check11_Click()
Dim i As Integer
For i = 0 To 9
    If VolumeControl(i).Enabled = True Then
        If Check11.Value = 1 Then
            SBMLink(i).Value = 1
        End If
        If Check11.Value = 0 Then
            SBMLink(i).Value = 0
        End If
    End If
Next i
End Sub

Private Sub Command1_Click()
On Local Error Resume Next
Dim A As Integer
For A = 0 To 9
    BalanceControl(A).Value = 0
Next A
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Command2_Click
Width = 9975
Height = 2775
End Sub

Function Run(strFilePath As String, Optional strParms As String, Optional strDir As String) As String
On Local Error Resume Next
Const SW_SHOW = 5
Select Case ShellExecute(0, "Open", strFilePath, strParms, strDir, SW_SHOW)
Case 0
    Run = "Insufficent system memory or corrupt program file"
Case Else
    Run = ""
End Select
End Function

Private Sub Timer1_Timer()
On Local Error Resume Next
ProgressBar1.Value = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_SIGNEDMETER, VolCtrl) + 200
ProgressBar2.Value = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_SIGNEDMETER, VolCtrl) + 200
End Sub

Private Sub VolumeControl_Scroll(i As Integer)
On Local Error Resume Next
volume = 65535 - CLng(VolumeControl(i).Value)
TXTVolumeControl(i).text = volume
BalanceControl_Scroll (i)
End Sub

Private Sub BalanceControl_Scroll(i As Integer)
On Local Error Resume Next
volR = TXTVolumeControl(i).text * (IIf(BalanceControl(i) >= 0, 1, (100 + BalanceControl(i)) / 100))
volL = TXTVolumeControl(i).text * (IIf(BalanceControl(i) <= 0, 1, (100 - BalanceControl(i)) / 100))
Select Case (i)
Case 0
    SetPANControl hmixer, VolCtrl, volL, volR
    TXTVol = (volume / 6553)
Case 1
    SetPANControl hmixer, CDVol, volL, volR
    Text1 = (volume / 6553)
Case 2
    SetPANControl hmixer, LineVol, volL, volR
    Text2 = (volume / 6553)
Case 3
    SetVolumeControl hmixer, MICROPHONE, volume
    Text3 = (volume / 6553)
Case 4
    SetPANControl hmixer, AUXVol, volL, volR
    Text4 = (volume / 6553)
Case 5
    SetPANControl hmixer, WavCtrl, volL, volR
    Text5 = (volume / 6553)
Case 6
    SetVolumeControl hmixer, TELEPHONE, volume
    Text6 = (volume / 6553)
Case 7
    SetPANControl hmixer, MIDIVol, volL, volR
    Text7 = (volume / 6553)
Case 8
    SetPANControl hmixer, I25InVol, volL, volR
    Text8 = (volume / 6553)
Case 9
    SetVolumeControl hmixer, PCSPEAKER, volume
    Text9 = (volume / 6553)
Case 10
    SetVolumeControl hmixer, Treble, volume
    Text10 = (volume / 6553)
Case 11
    SetVolumeControl hmixer, Bass, volume
    Text11 = (volume / 6553)
End Select
End Sub

Function GetTheDevice(TheTextBox As TextBox, TheControl As String)
On Local Error Resume Next
Select Case TheControl
Case "VolCtrl"
    volume = GetVolumeControlValue(hmixer, VolCtrl)
Case "WavCtrl"
    volume = GetVolumeControlValue(hmixer, WavCtrl)
Case "MICROPHONE"
    volume = GetVolumeControlValue(hmixer, MICROPHONE)
Case "CDVol"
    volume = GetVolumeControlValue(hmixer, CDVol)
Case "AUXVol"
    volume = GetVolumeControlValue(hmixer, AUXVol)
Case "TELEPHONE"
    volume = GetVolumeControlValue(hmixer, TELEPHONE)
Case "MIDIVol"
    volume = GetVolumeControlValue(hmixer, MIDIVol)
Case "PCSPEAKER"
    volume = GetVolumeControlValue(hmixer, PCSPEAKER)
Case "I25InVol"
    volume = GetVolumeControlValue(hmixer, I25InVol)
Case "LineVol"
    volume = GetVolumeControlValue(hmixer, LineVol)
Case "Bass"
    volume = GetVolumeControlValue(hmixer, Bass)
Case "Treble"
    volume = GetVolumeControlValue(hmixer, Treble)
End Select
TheTextBox.text = (volume / 6553)
End Function

Function NoDevice(TheTextBox As TextBox, VolFrameControl As Frame, VControl As Slider)
On Local Error Resume Next
TheTextBox.text = ("N/A")
VolFrameControl.Enabled = False
VControl.Value = 0
VControl.Visible = False
If VControl.Index > 9 Then Exit Function
Picture11(VControl.Index).Visible = False
SBMLink(VControl.Index).Visible = False
BalanceControl(VControl.Index).Visible = False
End Function

Private Sub Command2_Click()
On Local Error Resume Next
Dim i As Integer
rc = mixerOpen(hmixer, 0, 0, 0, 0)
If ((MMSYSERR_NOERROR <> rc)) Then
    DoColor lSettings.sActiveServerForm.txtStatus, "Couldn't open the mixer please check if a audio mixer is installed then retry."
    Exit Sub
End If
For i = 0 To 11
Select Case i
Case 0
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_VOLUME, VolCtrl)
    If (ok = True) Then
        GetTheDevice TXTVol, "VolCtrl"
    Else
        NoDevice TXTVol, VolFrame(0), VolumeControl(0)
    End If
Case 1
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, MIXERCONTROL_CONTROLTYPE_VOLUME, WavCtrl)
    If (ok = True) Then
        GetTheDevice Text5, "WavCtrl"
    Else
        NoDevice Text5, VolFrame(5), VolumeControl(5)
    End If
Case 2
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_VOLUME, MICROPHONE)
    If (ok = True) Then
        GetTheDevice Text3, "MICROPHONE"
    Else
        NoDevice Text3, VolFrame(3), VolumeControl(3)
    End If
Case 3
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, MIXERCONTROL_CONTROLTYPE_VOLUME, CDVol)
    If (ok = True) Then
        GetTheDevice Text1, "CDVol"
    Else
        NoDevice Text1, VolFrame(1), VolumeControl(1)
    End If
Case 4
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_AUXVol, MIXERCONTROL_CONTROLTYPE_VOLUME, AUXVol)
    If (ok = True) Then
        GetTheDevice Text4, "AUXVol"
    Else
        NoDevice Text4, VolFrame(4), VolumeControl(4)
    End If
Case 5
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE, MIXERCONTROL_CONTROLTYPE_VOLUME, TELEPHONE)
    If (ok = True) Then
        GetTheDevice Text6, "TELEPHONE"
    Else
        NoDevice Text6, VolFrame(6), VolumeControl(6)
    End If
Case 6
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, MIXERCONTROL_CONTROLTYPE_VOLUME, MIDIVol)
    If (ok = True) Then
        GetTheDevice Text7, "MIDIVol"
    Else
        NoDevice Text7, VolFrame(7), VolumeControl(7)
    End If
Case 7
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER, MIXERCONTROL_CONTROLTYPE_VOLUME, PCSPEAKER)
    If (ok = True) Then
        GetTheDevice Text9, "PCSPEAKER"
    Else
        
        
        Text9.text = ("N/A")
        VolFrame(0).Enabled = False
        VolumeControl(9).Value = 0
        VolumeControl(9).Visible = False
        If VolumeControl(9).Index > 9 Then Exit Sub
        Picture11(VolumeControl(9).Index).Visible = False
        SBMLink(VolumeControl(9).Index).Visible = False
        BalanceControl(VolumeControl(9).Index).Visible = False
        
        
        'NoDevice Text9, VolFrame(0), VolumeControl(9)
    End If
Case 8
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, MIXERCONTROL_CONTROLTYPE_VOLUME, I25InVol)
    If (ok = True) Then
        GetTheDevice Text8, "I25InVol"
    Else
        NoDevice Text8, VolFrame(8), VolumeControl(8)
    End If
Case 9
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, MIXERCONTROL_CONTROLTYPE_VOLUME, LineVol)
    If (ok = True) Then
        GetTheDevice Text2, "LineVol"
    Else
        NoDevice Text2, VolFrame(2), VolumeControl(2)
    End If
Case 10
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_BASS, Bass)
    If (ok = True) Then
        GetTheDevice Text11, "Bass"
    Else
        NoDevice Text11, VolFrame(11), VolumeControl(11)
    End If
Case 11
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_TREBLE, Treble)
    If (ok = True) Then
        GetTheDevice Text10, "Treble"
    Else
        NoDevice Text10, VolFrame(10), VolumeControl(10)
    End If
End Select
    If volume <> -1 Then
        TXTVolumeControl(i) = volume
        VolumeControl(i) = 65535 - volume
    End If
Next i
End Sub

Private Sub MuteOn_Click(m As Integer)
On Local Error Resume Next
Select Case m
Case 0
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 1
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 2
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 3
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 4
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_AUXVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 5
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 6
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 7
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 8
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
Case 9
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER, MIXERCONTROL_CONTROLTYPE_MUTE, mute)
End Select
SetMuteControl hmixer, mute, 1
End Sub

Private Sub MuteOff_Click(MOff As Integer)
Select Case MOff
Case 0
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 1
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_CDVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 2
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_LINEVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 3
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 4
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_AUXVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 5
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEDSVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 6
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_TELEPHONE, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 7
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MIDIVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 8
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_I25InVol, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
Case 9
    ok = GetMixerControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_PCSPEAKER, MIXERCONTROL_CONTROLTYPE_MUTE, unmute)
End Select
unSetMuteControl hmixer, unmute, 1
End Sub

Private Sub Slider6_scroll()
Dim i As Integer
For i = 0 To 9
If SBMLink(i).Value = 1 Then
If VolumeControl(i).Enabled = True Then
     SetPANControl hmixer, VolCtrl, volL, volR
     VolumeControl(i) = Slider6.Value
End If
VolumeControl_Scroll (i)
End If
Next i
End Sub
