VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmAIMFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " AIM Filter - Beta"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAIMFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   4080
      Top             =   4080
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   5
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   41
      Top             =   600
      Width           =   3375
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "coded by robbie saunders"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   48
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " build 106 - beta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1560
         TabIndex        =   47
         Top             =   1850
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   4755
         Left            =   -1920
         Picture         =   "frmAIMFilter.frx":0442
         Top             =   -360
         Width           =   6330
      End
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   3360
      TabIndex        =   37
      Text            =   "8"
      Top             =   4800
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   4
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   28
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
      Begin VB.OptionButton Option2 
         Caption         =   "Weird One  (Works Sometimes)"
         Height          =   255
         Left            =   360
         TabIndex        =   39
         Top             =   1920
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "No Capabilities (heh)"
         Height          =   255
         Left            =   360
         TabIndex        =   38
         Top             =   1680
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Enable Special Capabilites"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1560
         TabIndex        =   31
         Text            =   "aim.b.t"
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Text            =   "aim.b.f"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1200
         TabIndex        =   29
         Text            =   "aim.warn.user"
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "Block Command:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "UnBlock Command:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Warn User:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   3
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox Text13 
         Height          =   285
         Left            =   1320
         TabIndex        =   53
         Text            =   "1"
         Top             =   1410
         Width           =   375
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Enable Rate Limiting"
         Height          =   255
         Left            =   120
         TabIndex        =   51
         Top             =   1080
         Width           =   3135
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Block Incoming Errors (Crashes Old)"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1920
         Width           =   3135
      End
      Begin VB.TextBox SBYTE 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   3
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   46
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox SBYTE 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox SBYTE 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox SBYTE 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label17 
         Caption         =   "Rate Limit By:             Second(s)"
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   3240
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line17 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   3240
         Y1              =   1815
         Y2              =   1815
      End
      Begin VB.Line Line16 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   3240
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   3240
         Y1              =   975
         Y2              =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Sequence Bytes:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   360
      TabIndex        =   23
      Text            =   "6"
      Top             =   4680
      Width           =   615
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   1560
      TabIndex        =   21
      Text            =   "4"
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   840
      TabIndex        =   20
      Text            =   "0"
      Top             =   5880
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   2
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         ItemData        =   "frmAIMFilter.frx":17864
         Left            =   120
         List            =   "frmAIMFilter.frx":17877
         Style           =   2  'Dropdown List
         TabIndex        =   78
         Top             =   960
         Width           =   3135
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Index           =   4
         Left            =   120
         TabIndex        =   73
         Top             =   1200
         Width           =   3135
         Begin VB.TextBox Text21 
            Height          =   285
            Left            =   1080
            TabIndex        =   77
            Text            =   "aim.blank.icon"
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox Text20 
            Height          =   285
            Left            =   960
            TabIndex        =   75
            Text            =   "aim.file.error"
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label25 
            Caption         =   "Blank Icon: "
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label24 
            Caption         =   "File Error:"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Index           =   3
         Left            =   120
         TabIndex        =   68
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox Text19 
            Height          =   285
            Left            =   1680
            TabIndex        =   72
            Text            =   "aim.im.c"
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox Text18 
            Height          =   285
            Left            =   1680
            TabIndex        =   70
            Text            =   "aim.talk.c"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label23 
            Caption         =   "Fake Connect (IM):"
            Height          =   255
            Left            =   120
            TabIndex        =   71
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label22 
            Caption         =   "Fake Connect (Talk):"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Index           =   2
         Left            =   120
         TabIndex        =   64
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox Text16 
            Height          =   285
            Left            =   1200
            TabIndex        =   65
            Text            =   "aim.funny.buddy"
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Caption         =   "note: user must have aim 4.7 or better"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label18 
            Caption         =   "Funny Buddy:"
            Height          =   255
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   1200
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox Text15 
            Height          =   285
            Left            =   1320
            TabIndex        =   61
            Text            =   "aim.file.send"
            Top             =   225
            Width           =   1695
         End
         Begin VB.TextBox Text17 
            Height          =   285
            Left            =   1320
            TabIndex        =   60
            Text            =   "AIM FIlter.exe"
            Top             =   550
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "Fake File Send:"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Fake File Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   62
            Top             =   555
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   1200
         Width           =   3135
         Begin VB.TextBox Text14 
            Height          =   285
            Left            =   1320
            TabIndex        =   56
            Text            =   "aim.crash.user"
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label20 
            Alignment       =   2  'Center
            Caption         =   "note: user must support buddy icons"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label12 
            Caption         =   "Killer Command:"
            Height          =   255
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmAIMFilter.frx":178CB
         Left            =   2760
         List            =   "frmAIMFilter.frx":178ED
         TabIndex        =   54
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Block Incoming Killer .gif's =)"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Value           =   1  'Checked
         Width           =   3135
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Change Incoming Font Size To:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   3255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Change All IM's to `Auto Response`"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   3135
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   1
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1200
         MaxLength       =   7000
         TabIndex        =   17
         Text            =   "away test"
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Text            =   "aim.away.off"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1200
         TabIndex        =   13
         Text            =   "aim.away.on"
         Top             =   120
         Width           =   2055
      End
      Begin VB.Label Label8 
         Caption         =   "The Message:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Off Command:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "On Command:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   1095
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   3240
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   3240
         Y1              =   1455
         Y2              =   1455
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   $"frmAIMFilter.frx":1790F
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   120
         TabIndex        =   11
         Top             =   1560
         Width           =   3135
      End
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Text            =   "0"
      Top             =   5880
      Width           =   615
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check2"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   6240
      Width           =   2415
   End
   Begin MSWinsockLib.Winsock Winsock4 
      Left            =   2760
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   3
      Left            =   1800
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   2
      Left            =   1320
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   1
      Left            =   1800
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock3 
      Left            =   2280
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2295
      Index           =   0
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   3375
      TabIndex        =   1
      Top             =   600
      Width           =   3375
      Begin VB.ListBox List1 
         Height          =   1425
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Listen"
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   120
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "5190"
         Top             =   120
         Width           =   615
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   3240
         Y1              =   620
         Y2              =   620
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   3240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label3 
         Caption         =   "Listen on port:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   1575
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   1320
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "click here to visit aim filter's official website"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   49
      ToolTipText     =   "http://www.ssnbc.com/wiz/"
      Top             =   3050
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "about"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   27
      Top             =   225
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "extras"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   26
      Top             =   225
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "rates"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1800
      TabIndex        =   24
      Top             =   225
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "im toys"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   1200
      TabIndex        =   18
      Top             =   225
      Width           =   615
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "away"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   9
      Top             =   225
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "main"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   225
      Width           =   495
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00E0E0E0&
      X1              =   3480
      X2              =   3480
      Y1              =   480
      Y2              =   120
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   3480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line6 
      X1              =   120
      X2              =   3480
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   120
      Y1              =   480
      Y2              =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      X1              =   3480
      X2              =   3480
      Y1              =   600
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      BorderWidth     =   3
      X1              =   120
      X2              =   3480
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   120
      X2              =   3480
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   120
      X2              =   120
      Y1              =   600
      Y2              =   2880
   End
   Begin VB.Menu ListMenu 
      Caption         =   "ListMenu"
      Visible         =   0   'False
      Begin VB.Menu copylistitem 
         Caption         =   "Copy Item"
      End
      Begin VB.Menu removelistitem 
         Caption         =   "Remove Item"
      End
      Begin VB.Menu listline1 
         Caption         =   "-"
      End
      Begin VB.Menu clearthelist 
         Caption         =   "Clear List"
      End
   End
End
Attribute VB_Name = "frmAIMFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TheData(3) As String, IncomingStuff(3) As String, PacketLen(3), LeftOver(3) As Boolean
Dim A1, A2, A3, A4, A5, A6, A7, A8, A9, R1 As String
Dim TheServer, ThePort, Capa As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then
    BeginListen
Else
    Winsock3.Close
    Winsock4.Close
End If
End Sub

Private Sub clearthelist_Click()
List1.Clear
End Sub

Private Sub Combo2_Click()
On Error Resume Next
For i = 1 To 10
    Frame1(i).Visible = False
    DoEvents
Next i
Frame1(Combo2.ListIndex).Visible = True
End Sub

Private Sub copylistitem_Click()
Clipboard.Clear
Clipboard.SetText List1.List(List1.ListIndex)
End Sub

Private Sub Form_Load()
Combo2.ListIndex = 0
BeginListen
End Sub

Private Sub Label16_Click()
OpenURL "http://www.ssnbc.com/wiz/"
End Sub

Private Sub Label4_Click(Index As Integer)
On Error Resume Next
For i = 0 To 6
    Picture1(i).Visible = False
    Label4(i).FontBold = False
    DoEvents
Next i
Picture1(Index).Visible = True
Label4(Index).FontBold = True
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 And List1.ListIndex >= 0 Then PopupMenu ListMenu
End Sub

Private Sub removelistitem_Click()
List1.RemoveItem List1.ListIndex
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
'Incoming Packet Splitter/Joiner
Winsock1(Index).GetData IncomingStuff(Index) 'grab new info
TheData(Index) = TheData(Index) & IncomingStuff(Index) 'add it to cache
If Left(TheData(Index), 1) <> "*" Then 'if this messes up, reconnect
    ListAdd "Protocol Error (1)"
    BeginListen
End If
NextOne:
If LeftOver(Index) = False Then PacketLen(Index) = GetLength(Chr(0) & Mid(TheData(Index), 5, 2)) + 6 'grab the instruction length
If PacketLen(Index) > Len(TheData(Index)) Then
    LeftOver(Index) = True 'make sure we don't regrab the length bytes and slow us down
    Exit Sub
End If
ProcessData Index, Left(TheData(Index), PacketLen(Index)) 'process the instruction
TheData(Index) = Right(TheData(Index), Len(TheData(Index)) - PacketLen(Index)) 'remove it from the cache
PacketLen(Index) = 0
LeftOver(Index) = False 're-enable len grabbing
If Len(TheData(Index)) >= 6 Then GoTo NextOne: 'if we have the complete header of the next packet then keep going
End Sub

Private Sub Winsock3_ConnectionRequest(ByVal requestID As Long)
For i = 0 To 3
    Winsock1(i).Close
    DoEvents
Next i
Winsock1(0).Accept requestID
ListAdd "Connection(Login) " & Winsock1(0).RemoteHostIP
Winsock1(1).Connect "login.oscar.aol.com", "5190"
End Sub

Private Sub Winsock4_ConnectionRequest(ByVal requestID As Long)
For i = 0 To 3
    Winsock1(i).Close
    DoEvents
Next i
Winsock1(2).Accept requestID
ListAdd "Connection(AIM) " & Winsock1(0).RemoteHostIP
Winsock1(3).Connect TheServer, ThePort
End Sub

Sub ProcessData(Index As Integer, TheStuff As String)
Select Case Index
    Case 0 'login (client)
        TheStuff = Replace(TheStuff, Chr(14) & "sobbieraunders", Chr(15) & "sobbie raunders")
        SendPacket 1, TheStuff 'send to server
    Case 1 'login (server)
        If Mid(TheStuff, 8, 1) = Chr(23) And Mid(TheStuff, 10, 1) = Chr(3) Then 'username grab/server edit
            A1 = GetLength(Chr(0) & Mid(TheStuff, 19, 2))
            ListAdd "Welcome " & Mid(TheStuff, 21, A1) & "."
            A2 = GetLength(Chr(0) & Mid(TheStuff, 23 + A1, 2))
            A3 = Split(Mid(TheStuff, 25 + A1, A2), ":")
            TheServer = A3(0): ThePort = A3(1)
            TheStuff = Left(TheStuff, 22 + A1) & TwoByteLen("localhost:" & (Text1 + 1)) & Right(TheStuff, Len(TheStuff) - (24 + A1 + A2))
        End If
        SendPacket 0, TheStuff 'send to client
    Case 2 'aim (client)
        If Mid(TheStuff, 8, 1) = Chr(4) And Mid(TheStuff, 10, 1) = Chr(6) Then 'instant message sending
            A1 = Asc(Mid(TheStuff, 27, 1))
            A2 = Mid(TheStuff, 28, A1)
            ListAdd "IM To " & A2
            If InStr(1, TheStuff, Text3) <> 0 Then 'away on
                TheStuff = Left(TheStuff, 6) & AwayMessage("text/aolrtf; charset=" & Chr(34) & "us-ascii" & Chr(34), Text5)
                ListAdd "AIM Filter [away on]"
            ElseIf InStr(1, TheStuff, Text4) <> 0 Then 'away off
                TheStuff = Left(TheStuff, 6) & AwayMessage("", "")
                ListAdd "AIM Filter [away off]"
            ElseIf InStr(1, TheStuff, Text8) <> 0 Then 'user block
                TheStuff = Left(TheStuff, 6) & BlockUser(Mid(TheStuff, 28, A1), 8)
                ListAdd "AIM Filter [blocked]"
            ElseIf InStr(1, TheStuff, Text10) <> 0 Then 'unblock user
                TheStuff = Left(TheStuff, 6) & BlockUser(Mid(TheStuff, 28, A1), 10)
                ListAdd "AIM Filter [unblocked]"
            ElseIf InStr(1, TheStuff, Text11) <> 0 Then 'warn the user
                TheStuff = Left(TheStuff, 6) & UserWarning(Chr(0) & Chr(9) & Chr(0) & Chr(8) & Chr(0) & Chr(0), Mid(TheStuff, 28, A1))
                ListAdd "AIM Filter [warned]"
            ElseIf InStr(1, TheStuff, Text14) <> 0 Then 'add a corrupt buddy icon to the end of an im
                R1 = text_read(App.Path & "\image1.gif")
                Mid(R1, 7, 4) = ChrA("0 50 0 50")
                TheStuff = EditAttach(TheStuff, BuddyIconEdit(Mid(TheStuff, 17, 8), R1, ChrA("59 78 26 209")), "AIM Filter [user crashed]")
            ElseIf InStr(1, TheStuff, Text15) <> 0 Then 'send file request
                TheStuff = EditAttach(TheStuff, FileSendEdit(Mid(TheStuff, 17, 8), Text17), "AIM Filter [file send]")
            ElseIf InStr(1, TheStuff, Text18) <> 0 Then 'send talk connection request
                TheStuff = EditAttach(TheStuff, BlankAttach(Mid(TheStuff, 17, 8), ChrA("9 70 19 65 76 127 17 209 130 34 68 69 83 84 0 0")), "AIM Filter [connect t]")
            ElseIf InStr(1, TheStuff, Text19) <> 0 Then 'send im image connection request
                TheStuff = EditAttach(TheStuff, IMConnect(Mid(TheStuff, 17, 8)), "AIM Filter [connect i]")
            ElseIf InStr(1, TheStuff, Text21) <> 0 Then 'send blank buddy icon
                TheStuff = EditAttach(TheStuff, BlankAttach(Mid(TheStuff, 17, 8), ChrA("9 70 19 70 76 127 17 209 130 34 68 69 83 84 0 0")), "AIM Filter [connect t]")
            ElseIf InStr(1, TheStuff, Text20) <> 0 Then 'send file error
                TheStuff = EditAttach(TheStuff, BlankAttach(Mid(TheStuff, 17, 8), ChrA("9 70 19 67 76 127 17 209 130 34 68 69 83 84 0 0")), "AIM Filter [connect t]")
            ElseIf InStr(1, TheStuff, Text16) <> 0 Then 'send funny buddy list thing
                R1 = ""
                R2 = BuddyListForm(Mid(TheStuff, 28, A1) & " eats poo")
                For i = 1 To 100
                    R1 = R1 & R2
                    DoEvents
                Next i
                TheStuff = EditAttach(TheStuff, BuddyListEdit(Mid(TheStuff, 17, 8), R1), "AIM Filter [buddy list]")
            End If
            If Check3.Value = 1 And Right(TheStuff, 6) = "/HTML" & Chr(62) Then
                TheStuff = TheStuff & Chr(0) & Chr(4) & Chr(0) & Chr(0)
            End If
        End If
        If Mid(TheStuff, 8, 1) = Chr(2) And Mid(TheStuff, 10, 1) = Chr(4) Then 'member info set
            If Check6.Value = 1 Then
                A1 = InStr(1, TheStuff, "/HTML" & Chr(62) & Chr(0) & Chr(5))
                If A1 <> 0 Then
                    TheStuff = Left(TheStuff, A1 + 5)
                End If
                If Option2.Value Then
                    TheStuff = TheStuff & Chr(0) & Chr(5) & TwoByteLen(ChrA("9 70 19 70 76 127 17 209 130 34 68 69 83 84 0 0 9 70 19 70 76 127 17 209 130 34 68 69 83 84 0 0"))
                End If
            End If
        End If
        If Mid(TheStuff, 8, 1) = Chr(14) And Mid(TheStuff, 10, 1) = Chr(6) Then 'chatsend
            If Check3.Value = 1 And Right(TheStuff, 6) = "/HTML" & Chr(62) Then
                TheStuff = TheStuff & Chr(0) & Chr(4) & Chr(0) & Chr(0)
            End If
        End If
        If Mid(TheStuff, 8, 1) = Chr(5) And Mid(TheStuff, 10, 2) = Chr(6) Then 'ad request
            ListAdd "AIM Filter [ad blocked]"
            Exit Sub
        End If
        SendPacket 3, TheStuff 'send to server
    Case 3 'aim (server)
        If Mid(TheStuff, 8, 1) = Chr(4) And Mid(TheStuff, 10, 1) = Chr(7) Then 'instant message receiving
            A1 = Asc(Mid(TheStuff, 27, 1))
            A2 = Mid(TheStuff, 28, A1)
            ListAdd "IM From " & A2
            If Check5.Value = 1 Then
                For i = 0 To 9
                    If InStr(1, TheStuff, "SIZE=" & i) <> 0 Then
                        TheStuff = Replace(TheStuff, "SIZE=" & i, "SIZE=" & Combo1.text)
                    End If
                    DoEvents
                Next i
            End If
            If Check8.Value = 1 Then
                A1 = InStr(1, TheStuff, "GIF89a")
                If A1 <> 0 Then
                    If Mid(TheStuff, A1 + 7, 1) <> Chr(0) Or Mid(TheStuff, A1 + 9, 1) <> Chr(0) Then
                        ListAdd "AIM Filter [crash blocked]"
                        Exit Sub
                    End If
                End If
            End If
        End If
        If Mid(TheStuff, 8, 1) = Chr(4) And Mid(TheStuff, 10, 1) = Chr(1) Then 'im error
            If Check4.Value = 1 Then ListAdd "AIM Filter [error blocked]": Exit Sub
        End If
        SendPacket 2, TheStuff 'send to client
End Select
End Sub

Sub SendPacket(Index As Integer, ThePacket As String)
If Winsock1(Index).State = sckConnected Then
    If Check7.Value = 1 And Index = 3 Then Pause Text13
    SBYTE(Index) = SBYTE(Index) + 1
    If SBYTE(Index) > 65535 Then SBYTE(Index) = 0
    Mid(ThePacket, 3, 2) = IntegerToBase256(SBYTE(Index))
    Mid(ThePacket, 5, 2) = IntegerToBase256(Len(ThePacket) - 6) 'correct any length jackups
    Winsock1(Index).SendData ThePacket
Else
    ListAdd "Protocol Error (2)"
    BeginListen
End If
End Sub

Sub BeginListen()
On Error Resume Next
For i = 0 To 3
    Winsock1(i).Close
    DoEvents
Next i
Winsock3.Close
Winsock4.Close
Winsock3.LocalPort = Text1
Winsock4.LocalPort = Text1 + 1
Winsock3.Close
Winsock4.Close
Winsock3.Listen
Winsock4.Listen
End Sub

Sub ListAdd(TheEvent)
Dim B1, B2, B3, B4, B5
For i = 0 To List1.ListCount - 1
    If Right(List1.List(i), Len(TheEvent) + 2) = "x " & TheEvent Then
        B1 = List1.List(i)
        B2 = InStr(1, B1, "x")
        B3 = Left(B1, B2 - 2)
        B4 = Right(B1, Len(B1) - (B2 + 1))
        B3 = B3 + 1
        List1.List(i) = B3 & " x " & B4
        Exit Sub
    End If
    DoEvents
Next i
List1.AddItem "1 x " & TheEvent
End Sub

Function EditAttach(TheIM As String, TheAtt As String, TheInfo As String)
TheIM = Left(TheIM, 27 + A1) & TheAtt
Mid(TheIM, 26, 1) = Chr(2)
Mid(TheIM, 23, 2) = ChrA("0 0")
ListAdd TheInfo
EditAttach = TheIM
End Function
