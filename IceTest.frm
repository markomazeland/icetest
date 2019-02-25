VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "Test"
   ClientHeight    =   9810
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   13440
   Icon            =   "IceTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9810
   ScaleWidth      =   13440
   Begin VB.CommandButton cmdIceSort 
      Caption         =   "IceSort ..."
      Height          =   375
      Left            =   10080
      MaskColor       =   &H80000000&
      Picture         =   "IceTest.frx":67E2
      TabIndex        =   41
      ToolTipText     =   "Compose start groups using IceSort Tool"
      Top             =   7200
      Width           =   1575
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6480
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   6480
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   80
      LocalPort       =   80
   End
   Begin VB.Frame fraDivider2 
      BackColor       =   &H80000000&
      Height          =   5895
      Left            =   6000
      MousePointer    =   9  'Size W E
      TabIndex        =   1
      ToolTipText     =   "Click here to resize the the left and right part of the window"
      Top             =   1080
      Width           =   135
   End
   Begin VB.Timer Timer2 
      Interval        =   60000
      Left            =   3360
      Top             =   7800
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   60
      Top             =   9120
      Visible         =   0   'False
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame fraOther 
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   6240
      TabIndex        =   59
      Top             =   960
      Width           =   6855
      Begin VB.Data dtaTestInfo 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   5400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame fraJudges 
         Caption         =   "Judges"
         Height          =   735
         Left            =   3720
         TabIndex        =   45
         Top             =   6960
         Width           =   3135
         Begin VB.ComboBox cmbNumJudges 
            DataField       =   "Num_j_3"
            DataSource      =   "dtaTestInfo"
            Height          =   315
            Index           =   3
            ItemData        =   "IceTest.frx":6AEC
            Left            =   720
            List            =   "IceTest.frx":6AFC
            Sorted          =   -1  'True
            TabIndex        =   67
            ToolTipText     =   "How many judges will judge this test"
            Top             =   -120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbNumJudges 
            DataField       =   "Num_j_2"
            DataSource      =   "dtaTestInfo"
            Height          =   315
            Index           =   2
            ItemData        =   "IceTest.frx":6B0C
            Left            =   720
            List            =   "IceTest.frx":6B1C
            Sorted          =   -1  'True
            TabIndex        =   70
            ToolTipText     =   "How many judges will judge this test"
            Top             =   -120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbNumJudges 
            DataField       =   "Num_j_1"
            DataSource      =   "dtaTestInfo"
            Height          =   315
            Index           =   1
            ItemData        =   "IceTest.frx":6B2C
            Left            =   0
            List            =   "IceTest.frx":6B3C
            Sorted          =   -1  'True
            TabIndex        =   66
            ToolTipText     =   "How many judges will judge this test"
            Top             =   0
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.CommandButton cmdMarks 
            Enabled         =   0   'False
            Height          =   495
            Left            =   1920
            Picture         =   "IceTest.frx":6B4C
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Individual marks per participant"
            Top             =   120
            Width           =   495
         End
         Begin VB.CommandButton cmdTestInfo 
            Height          =   495
            Left            =   2520
            Picture         =   "IceTest.frx":6F8E
            Style           =   1  'Graphical
            TabIndex        =   49
            ToolTipText     =   "Extra information about this test (judges, sponsor, etc.)"
            Top             =   120
            Width           =   495
         End
         Begin VB.ComboBox cmbNumJudges 
            DataField       =   "Num_j_0"
            DataSource      =   "dtaTestInfo"
            Height          =   360
            Index           =   0
            ItemData        =   "IceTest.frx":7118
            Left            =   120
            List            =   "IceTest.frx":7128
            Sorted          =   -1  'True
            TabIndex        =   46
            ToolTipText     =   "How many judges will judge this test"
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblNumJudges 
            Caption         =   "Judges"
            Height          =   360
            Left            =   1080
            TabIndex        =   47
            ToolTipText     =   "How many judges will judge this test"
            Top             =   240
            Width           =   1155
         End
      End
      Begin VB.Frame fraGroups 
         Caption         =   "Groups / Colors / Classes"
         Height          =   1215
         Left            =   1800
         TabIndex        =   38
         Top             =   4920
         Width           =   4215
         Begin VB.CheckBox chkSplitResultLists 
            Caption         =   "Split result lists"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            ToolTipText     =   "Split result lists (and finals) in different (age) classes"
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton cmdComposeGroups 
            Caption         =   "Compose"
            Height          =   375
            Left            =   2400
            MaskColor       =   &H80000000&
            Picture         =   "IceTest.frx":7138
            TabIndex        =   42
            ToolTipText     =   "Compose start groups (again)"
            Top             =   240
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.ComboBox cmbGroupSize 
            DataField       =   "Groupsize"
            DataSource      =   "dtaTest"
            Height          =   360
            ItemData        =   "IceTest.frx":7442
            Left            =   120
            List            =   "IceTest.frx":7455
            Sorted          =   -1  'True
            TabIndex        =   39
            ToolTipText     =   "What is the size of groups"
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Frame fraFinals 
         Caption         =   "Final"
         Enabled         =   0   'False
         Height          =   855
         Left            =   0
         TabIndex        =   43
         Top             =   6000
         Visible         =   0   'False
         Width           =   3255
         Begin VB.CheckBox chkSplitFinals 
            Caption         =   "Split finals"
            Height          =   375
            Left            =   120
            TabIndex        =   63
            ToolTipText     =   "Split finals in different (age) classes"
            Top             =   240
            Width           =   500
         End
         Begin VB.CommandButton cmdComposeFinals 
            Caption         =   "Compose"
            Height          =   375
            Left            =   1680
            TabIndex        =   44
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.Frame fraCurrent 
         Caption         =   "Current Participant"
         Height          =   5775
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   6735
         Begin VB.Frame fraMarks 
            Enabled         =   0   'False
            Height          =   2055
            Left            =   -2640
            TabIndex        =   26
            Top             =   3120
            Width           =   6375
            Begin VB.TextBox txtMarks 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   4
               Left            =   3600
               MaxLength       =   5
               TabIndex        =   20
               ToolTipText     =   "Mark of the respective judge"
               Top             =   1680
               Width           =   855
            End
            Begin VB.TextBox txtMarks 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   3
               Left            =   3600
               MaxLength       =   5
               TabIndex        =   19
               ToolTipText     =   "Mark of the respective judge"
               Top             =   1440
               Width           =   855
            End
            Begin VB.TextBox txtMarks 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   2
               Left            =   3600
               MaxLength       =   5
               TabIndex        =   18
               ToolTipText     =   "Mark of the respective judge"
               Top             =   1080
               Width           =   855
            End
            Begin VB.TextBox txtMarks 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   1
               Left            =   3600
               MaxLength       =   5
               TabIndex        =   17
               ToolTipText     =   "Mark of the respective judge"
               Top             =   720
               Width           =   855
            End
            Begin VB.TextBox txtMarks 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   3600
               MaxLength       =   5
               TabIndex        =   16
               ToolTipText     =   "Mark of the respective judge"
               Top             =   360
               Width           =   855
            End
            Begin VB.Label lblMarks 
               Alignment       =   1  'Right Justify
               Caption         =   "Judge E:"
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   31
               Top             =   1680
               Width           =   3375
            End
            Begin VB.Label lblMarks 
               Alignment       =   1  'Right Justify
               Caption         =   "Judge D:"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   30
               Top             =   1440
               Width           =   3375
            End
            Begin VB.Label lblMarks 
               Alignment       =   1  'Right Justify
               Caption         =   "Judge C:"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   29
               Top             =   1080
               Width           =   3375
            End
            Begin VB.Label lblMarks 
               Alignment       =   1  'Right Justify
               Caption         =   "Judge B:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   28
               Top             =   720
               Width           =   3375
            End
            Begin VB.Label lblMarks 
               Alignment       =   1  'Right Justify
               Caption         =   "Judge A:"
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   27
               Top             =   360
               Width           =   3375
            End
         End
         Begin VB.Frame fraParticipant 
            Height          =   735
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   6495
            Begin VB.TextBox txtParticipant 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               MaxLength       =   3
               TabIndex        =   15
               ToolTipText     =   "Enter a startnumber of a participant"
               Top             =   240
               Width           =   735
            End
            Begin VB.CommandButton cmdInfo 
               Enabled         =   0   'False
               Height          =   375
               Left            =   5880
               Picture         =   "IceTest.frx":7468
               Style           =   1  'Graphical
               TabIndex        =   25
               ToolTipText     =   "Background information of current participant"
               Top             =   240
               Width           =   375
            End
            Begin VB.Label lblParticipant 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   1200
               TabIndex        =   24
               Top             =   360
               UseMnemonic     =   0   'False
               Width           =   3495
               WordWrap        =   -1  'True
            End
         End
         Begin VB.Frame fraResults 
            Enabled         =   0   'False
            Height          =   1410
            Left            =   120
            TabIndex        =   34
            Top             =   4320
            Width           =   6375
            Begin VB.CheckBox chkNoStart 
               Caption         =   "&No (further) start"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   1320
               TabIndex        =   68
               ToolTipText     =   "Red flag has been shown"
               Top             =   960
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.CheckBox chkWithdrawn 
               Caption         =   "&Withdrawn"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   120
               TabIndex        =   36
               ToolTipText     =   "Withdraw participant from this test"
               Top             =   600
               Width           =   2415
            End
            Begin VB.CheckBox chkFlag 
               Caption         =   "&Red flag"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   120
               TabIndex        =   37
               ToolTipText     =   "Red flag has been shown"
               Top             =   960
               Visible         =   0   'False
               Width           =   2655
            End
            Begin VB.CommandButton cmdOK 
               Cancel          =   -1  'True
               Caption         =   "&OK"
               Height          =   375
               Left            =   5040
               TabIndex        =   23
               ToolTipText     =   "Save marks of current participant"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtScore 
               Height          =   285
               Left            =   5280
               Locked          =   -1  'True
               TabIndex        =   22
               ToolTipText     =   "Score in the current section"
               Top             =   240
               Width           =   855
            End
            Begin VB.CheckBox chkDisqualified 
               Caption         =   "&Eliminated"
               ForeColor       =   &H000000C0&
               Height          =   255
               Left            =   120
               TabIndex        =   35
               ToolTipText     =   "Eliminate participant for this test"
               Top             =   240
               Width           =   2655
            End
            Begin VB.Label lblScore 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Score:"
               Height          =   195
               Left            =   3720
               TabIndex        =   51
               Top             =   240
               Width           =   1305
            End
         End
         Begin VB.Frame fraTime 
            Height          =   855
            Left            =   120
            TabIndex        =   32
            Top             =   3240
            Visible         =   0   'False
            Width           =   6255
            Begin VB.TextBox txtTime 
               Height          =   375
               Left            =   5160
               TabIndex        =   21
               Top             =   240
               Width           =   975
            End
            Begin VB.Label lblTime 
               Alignment       =   1  'Right Justify
               Caption         =   "Time:"
               Height          =   255
               Left            =   3000
               TabIndex        =   33
               Top             =   240
               Width           =   1215
            End
         End
      End
   End
   Begin VB.Frame fraLists 
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   120
      TabIndex        =   58
      Top             =   960
      Width           =   5535
      Begin MSWinsockLib.Winsock winsock 
         Left            =   3240
         Top             =   7320
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Frame fraPrevious 
         Caption         =   "Previous participant"
         Height          =   855
         Left            =   0
         TabIndex        =   61
         Top             =   360
         Width           =   5535
         Begin VB.TextBox txtPrevious 
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   62
            ToolTipText     =   "Previously entered participant"
            Top             =   240
            Width           =   5175
         End
      End
      Begin VB.Frame fraDivider1 
         BackColor       =   &H80000000&
         Height          =   100
         Left            =   0
         MousePointer    =   7  'Size N S
         TabIndex        =   2
         ToolTipText     =   "Click here to resize the upper and lower part of the window"
         Top             =   3120
         Width           =   5535
      End
      Begin VB.Frame fraShow 
         Caption         =   "Show"
         Height          =   855
         Left            =   0
         TabIndex        =   10
         Top             =   5760
         Width           =   5535
         Begin VB.CheckBox chkFeifId 
            Caption         =   "&FEIFId"
            Height          =   255
            Left            =   3720
            TabIndex        =   65
            ToolTipText     =   "Show the FEIFId of horses"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkTeam 
            Caption         =   "&Team/Club"
            Height          =   255
            Left            =   2400
            TabIndex        =   64
            ToolTipText     =   "Show the team/club of riders"
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chkRein 
            Caption         =   "&rein"
            Height          =   255
            Left            =   1440
            TabIndex        =   12
            ToolTipText     =   "Show left/right rein for participants not yet started"
            Top             =   360
            Width           =   1515
         End
         Begin VB.CheckBox chkColor 
            Caption         =   "&groups / colors"
            Height          =   195
            Left            =   480
            TabIndex        =   11
            ToolTipText     =   "Show group and color for participants not yet started"
            Top             =   360
            Width           =   1755
         End
      End
      Begin VB.Frame fraNotYet 
         Caption         =   "Participants not yet started"
         Height          =   2535
         Left            =   0
         TabIndex        =   7
         Top             =   3240
         Width           =   5535
         Begin VB.TextBox txtMove 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDBCtls.DBList dblstNotYet 
            Bindings        =   "IceTest.frx":799A
            DataSource      =   "dtaNotYet"
            Height          =   285
            Left            =   120
            TabIndex        =   8
            ToolTipText     =   "Select one of the participants not yet started"
            Top             =   360
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   503
            _Version        =   393216
            ListField       =   ""
            BoundColumn     =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame fraAlready 
         Caption         =   "Participants already started "
         Height          =   1695
         Left            =   0
         TabIndex        =   3
         Top             =   1320
         Width           =   5500
         Begin VB.ListBox lstAlready 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            ItemData        =   "IceTest.frx":79B2
            Left            =   1320
            List            =   "IceTest.frx":79B4
            TabIndex        =   6
            ToolTipText     =   "Select one of the participants already started"
            Top             =   960
            Width           =   4575
         End
         Begin VB.TextBox txtAlready 
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MSDBCtls.DBList dblstAlready 
            Bindings        =   "IceTest.frx":79B6
            DataSource      =   "dtaAlready"
            Height          =   960
            Left            =   -120
            TabIndex        =   4
            Top             =   1080
            Visible         =   0   'False
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   1693
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Data dtaTestSection 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data dtaMarks 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data dtaTest 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   700
      Left            =   480
      Top             =   6120
   End
   Begin VB.Data dtaParticipant 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data dtaAlready 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data dtaNotYet 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   54
      Top             =   9435
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   3519
            MinWidth        =   3529
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20108
            Key             =   "StatusMessage"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsSection 
      Height          =   9855
      Index           =   0
      Left            =   0
      TabIndex        =   50
      Top             =   600
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   17383
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedHeight  =   529
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   1764
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Test"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsSection 
      Height          =   3495
      Index           =   2
      Left            =   360
      TabIndex        =   52
      Top             =   720
      Visible         =   0   'False
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6165
      TabWidthStyle   =   2
      TabFixedHeight  =   529
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   1764
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsSection 
      Height          =   2775
      Index           =   1
      Left            =   360
      TabIndex        =   53
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4895
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedHeight  =   529
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   1764
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsSection 
      Height          =   2775
      Index           =   3
      Left            =   360
      TabIndex        =   69
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4895
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedHeight  =   529
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   1764
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip tbsSelFin 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      TabWidthStyle   =   2
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   1764
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Preliminary Round"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   630
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   13440
      _ExtentX        =   23707
      _ExtentY        =   1111
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      _Version        =   393216
   End
   Begin RichTextLib.RichTextBox rtfResult 
      Height          =   855
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      _Version        =   393217
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"IceTest.frx":79CF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblJudge 
      Caption         =   "Judge 1:"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   55
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New Event"
      End
      Begin VB.Menu mnuFileChange 
         Caption         =   "&Open Event..."
      End
      Begin VB.Menu mnuFileEven 
         Caption         =   "&Event Properties"
         Begin VB.Menu mnuFileEvenName 
            Caption         =   "&Event name, date, code"
         End
         Begin VB.Menu mnuFileEvenCode 
            Caption         =   "&FEIF WorldRanking"
         End
         Begin VB.Menu mnuFileEvenSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileEvenTest 
            Caption         =   "&Tests"
         End
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBackup 
         Caption         =   "&Backup"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuFileFEIFWR 
         Caption         =   "FEIF &WorldRanking File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Begin VB.Menu mnuFilePrintResultFinal 
            Caption         =   "&Final Result List"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnuFilePrintResultInterim 
            Caption         =   "&Provisional Result List"
         End
         Begin VB.Menu mnuFilePrintResultRevised 
            Caption         =   "&Revised Result List"
         End
         Begin VB.Menu mnuFilePrintSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFilePrintFormJudges 
            Caption         =   "&Judges' Form (list format)"
         End
         Begin VB.Menu mnuFilePrintFormJudges3 
            Caption         =   "J&udges' Form (individual format)"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFilePrintFormJudgesLand 
            Caption         =   "Judges' Form &Landscape"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFilePrintFormTime 
            Caption         =   "&Time Keepers Form"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuFilePrintFormStart 
            Caption         =   "&Starting Order"
         End
         Begin VB.Menu mnuFilePrintFormEquipment 
            Caption         =   "&Equipment check"
            Begin VB.Menu mnuFilePrintFormEquipmentComplete 
               Caption         =   "&Complete starting order"
            End
            Begin VB.Menu mnuFilePrintFormEquipmentOnly 
               Caption         =   "&Participants to be checked only"
            End
         End
         Begin VB.Menu mnuFilePrintLog 
            Caption         =   "&Log File"
         End
         Begin VB.Menu mnuFilePrintSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFilePrintAll 
            Caption         =   "&All participants"
            Begin VB.Menu mnuFilePrintAllPrinter 
               Caption         =   "to &printer"
            End
            Begin VB.Menu mnuFilePrintAllMerge 
               Caption         =   "as &Merge File for MS Word/Excell"
            End
         End
         Begin VB.Menu mnuFilePrintOverview 
            Caption         =   "&Overview"
            Begin VB.Menu mnuFilePrintOverviewMarks 
               Caption         =   "&including marks per judge"
            End
            Begin VB.Menu mnuFilePrintOverviewNoMarks 
               Caption         =   "&without marks per judge"
            End
            Begin VB.Menu mnuFilePrintOverviewFinals 
               Caption         =   "&finals/top 10 only"
            End
         End
         Begin VB.Menu mnuFilePrintComb 
            Caption         =   "&Combination"
            Begin VB.Menu mnuFilePrintCombComb 
               Caption         =   "*"
               Index           =   0
            End
         End
         Begin VB.Menu mnuFilePrintForms 
            Caption         =   "&Forms..."
         End
         Begin VB.Menu mnuFilePrintRider 
            Caption         =   "&Results per participant..."
         End
         Begin VB.Menu mnuFilePrintFormEntrance 
            Caption         =   "&Entrance Checks"
         End
         Begin VB.Menu mnuFilePrintWarnings 
            Caption         =   "Warnings"
         End
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "&Reports ..."
         Visible         =   0   'False
         Begin VB.Menu mnuFileSaveAsItem 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFileResults 
         Caption         =   "&Archive"
         Begin VB.Menu mnuFileResultsRtf 
            Caption         =   "&Result Lists and Forms..."
         End
         Begin VB.Menu mnuFileResultsHtml 
            Caption         =   "Results in &Browser..."
         End
         Begin VB.Menu mnuFileResultsExcel 
            Caption         =   "Results in &Excel"
         End
      End
      Begin VB.Menu mnuFileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Participants"
      Begin VB.Menu mnuEditSelect 
         Caption         =   "&Select Participant in this test"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "&Add Participant to this test"
      End
      Begin VB.Menu mnuEditRemove 
         Caption         =   "&Remove Participant-Marks from this test"
      End
      Begin VB.Menu mnuEditChangeRein 
         Caption         =   "&Change Rein"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditNoStart 
         Caption         =   "&No start in next heat(s)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditMove 
         Caption         =   "&Move Participant in Startlist"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStartOrder 
         Caption         =   "&Recompose Starting Order"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditTieBreak 
         Caption         =   "&Tiebreak"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditWarning 
         Caption         =   "&Warnings "
      End
      Begin VB.Menu mnuEditSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "&Find Participant"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditNew 
         Caption         =   "&New Participant"
      End
      Begin VB.Menu mnuEditEdit 
         Caption         =   "&Edit Participant"
      End
      Begin VB.Menu mnuEditPart 
         Caption         =   "&View Participants"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Tests"
      Begin VB.Menu mnutestQual1 
         Caption         =   "1"
         Begin VB.Menu mnuTestQual1Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual2 
         Caption         =   "2"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual2Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual3 
         Caption         =   "3"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual3Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual4 
         Caption         =   "4"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual4Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual5 
         Caption         =   "5"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual5Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual6 
         Caption         =   "6"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual6Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual7 
         Caption         =   "7"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual7Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual8 
         Caption         =   "8"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual8Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual9 
         Caption         =   "9"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual9Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual10 
         Caption         =   "10"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual10Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnutestQual11 
         Caption         =   "11"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual11Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual12 
         Caption         =   "12"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual12Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual13 
         Caption         =   "13"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual13Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual14 
         Caption         =   "14"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual14Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual15 
         Caption         =   "15"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual15Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual16 
         Caption         =   "16"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual16Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual17 
         Caption         =   "17"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual17Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual18 
         Caption         =   "18"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual18Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual19 
         Caption         =   "19"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual19Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual20 
         Caption         =   "20"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual20Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnutestQual21 
         Caption         =   "21"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual21Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual22 
         Caption         =   "22"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual22Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual23 
         Caption         =   "23"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual23Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual24 
         Caption         =   "24"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual24Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual25 
         Caption         =   "25"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual25Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual26 
         Caption         =   "26"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual26Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual27 
         Caption         =   "27"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual27Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual28 
         Caption         =   "28"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual28Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestQual29 
         Caption         =   "29"
         Visible         =   0   'False
         Begin VB.Menu mnuTestQual29Test 
            Caption         =   "*"
            Index           =   0
         End
      End
      Begin VB.Menu mnuTestSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTestAll 
         Caption         =   "&Show all tests"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuTestSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTestAddNew 
         Caption         =   "&New Test"
      End
      Begin VB.Menu mnuTestEdit 
         Caption         =   "&Edit Test..."
      End
      Begin VB.Menu mnuTestRemove 
         Caption         =   "&Remove Test..."
      End
   End
   Begin VB.Menu mnuComb 
      Caption         =   "&Combinations"
      Begin VB.Menu mnuCombAddNew 
         Caption         =   "&New Combination"
      End
      Begin VB.Menu mnuCombEdit 
         Caption         =   "&Edit Combination..."
      End
      Begin VB.Menu mnuCombRemove 
         Caption         =   "&Remove Combination..."
      End
   End
   Begin VB.Menu mnuTool 
      Caption         =   "Too&ls"
      Begin VB.Menu mnuToolCompress 
         Caption         =   "&Compact and Repair Database"
      End
      Begin VB.Menu mnuToolImport 
         Caption         =   "&Import"
         Begin VB.Menu mnuToolImportFipo 
            Caption         =   "&Rules"
         End
         Begin VB.Menu mnuToolImportSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuToolImportParticipants 
            Caption         =   "&Participants"
            Begin VB.Menu mnuToolImportAccess 
               Caption         =   "MS &Access (FEIF compatible)"
            End
            Begin VB.Menu mnuToolImportExcel 
               Caption         =   "MS E&xcel (Xls)"
            End
            Begin VB.Menu mnuToolImportTab 
               Caption         =   "&Tab delimited (Txt)"
            End
            Begin VB.Menu mnuToolImportCsv 
               Caption         =   "&Comma delimited (Csv)"
            End
            Begin VB.Menu mnuToolImportNSIJP 
               Caption         =   "&NSIJP (Mdb)"
               Visible         =   0   'False
            End
            Begin VB.Menu mnuToolImportDI 
               Caption         =   "&DI (Mdb)"
               Visible         =   0   'False
            End
         End
         Begin VB.Menu mnuToolImportBackup 
            Caption         =   "&Backup"
         End
         Begin VB.Menu mnuToolImportMarks 
            Caption         =   "&Marks"
            Begin VB.Menu mnuToolImportMarksExcel 
               Caption         =   "MS Excel"
            End
         End
      End
      Begin VB.Menu mnuToolOutput 
         Caption         =   "&Output"
         Begin VB.Menu mnuToolOutputParticipants 
            Caption         =   "&Participants"
            Begin VB.Menu mnuToolOutputExcel 
               Caption         =   "MS E&xcel (Xls)"
            End
            Begin VB.Menu mnuToolOutputCsv 
               Caption         =   "&Comma Delimited (Csv)"
            End
         End
         Begin VB.Menu mnuToolOutputMarks 
            Caption         =   "&Marks"
            Begin VB.Menu mnuToolOutputMarksExcel 
               Caption         =   "MS Excel (Xls)"
            End
         End
         Begin VB.Menu mnuToolOutputJudges 
            Caption         =   "Marks for &Judges"
            Begin VB.Menu mnuToolOutputJudgesExcel 
               Caption         =   "MS Excel (Xls)"
            End
         End
         Begin VB.Menu mnuToolOutputOntrackdata 
            Caption         =   "On-Track data"
            Shortcut        =   {F8}
         End
      End
      Begin VB.Menu mnuToolOutputFolders 
         Caption         =   "Output &Folders"
         Begin VB.Menu mnuToolOutputRtf 
            Caption         =   "&Result Lists and Forms"
         End
         Begin VB.Menu mnuToolOutputHtml 
            Caption         =   "&Browser"
         End
         Begin VB.Menu mnuToolOutputExternal 
            Caption         =   "&Extra files"
         End
      End
      Begin VB.Menu mnuToolForm 
         Caption         =   "&Forms"
         Begin VB.Menu mnuToolFormNew 
            Caption         =   "&Add new form"
         End
         Begin VB.Menu mnuToolFormEdit 
            Caption         =   "&Edit form"
         End
         Begin VB.Menu mnuToolFormDel 
            Caption         =   "&Delete form"
         End
      End
      Begin VB.Menu mnuToolReset 
         Caption         =   "&Reset screen "
      End
      Begin VB.Menu mnuToolSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupPopUp 
         Caption         =   "*"
         Index           =   0
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpTests 
         Caption         =   "Tests"
         Begin VB.Menu mnuPopUpTestsTest 
            Caption         =   "*"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "Help &Index"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpFeif 
         Caption         =   "&Support on Internet"
      End
      Begin VB.Menu mnuHelpFeifWR 
         Caption         =   "&WorldRanking"
      End
      Begin VB.Menu mnuHelpIcetest 
         Caption         =   "&IceTest News"
      End
      Begin VB.Menu mnuHelpCheckForUpdate 
         Caption         =   "&Check for update"
      End
      Begin VB.Menu mnuHelpFEIFTech 
         Caption         =   "&Technical support"
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Main functions, needs cleaning up and splitting into more modules

' Copyright (C) Marko Mazeland 2003
'
' This program is free software; you can redistribute it and/or modify it under the terms of the
' GNU General Public License as published by the Free Software Foundation; either version 2 of the License,
' or (at your option) any later version.
' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the
' implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License
' for more details (http://www.opensource.org/licenses/gpl-license.php).
'
' You should have received a copy of the GNU General Public License along with this program; if not, write to the
' Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA

Option Explicit
Option Compare Text

Public Tempvar As Variant
Public TestCode As String
Public TestName As String
Public TestStatus As Integer
Public TestSection As Integer
Public TestTable As Integer
Public TestJudges As Integer
Public TestColors As String

Public TestMarkDecimals As Integer
Public TestTimeDecimals As Integer
Public TestMarkFormat As String
Public TestTimeFormat As String
Public TestTotalFormat As String
Public TestInfoMessage As String

Public EventName As String
Public LastUsedSta As String

Public miTabMinWidth As Integer
Public miTabMinHeight As Integer
Public miInvalidMark As Integer
Public mctlActive As Control
Public miAlreadyHeight As Integer
Public miParticipantLeft As Integer
Public miChangeCaption As Integer
Public miDoNotCheckTieBreakAgain As Integer
Public miHorseAgeLimit As Integer
Public miMaxTestSection As Integer
Public miBlockOutputToExcel As Integer

'Constant VersionSwitch, can be used to handle special conditions for national compilations.
Const VersionSwitch = "feif"

Const ConnectorApplication = "feifconnector.exe"

Private blnConnected As Boolean         'flag: internet connection (winsock) active?
Private winsockresponse As String       'storage for winsock's response
Public ontrack As Boolean               'flag: ontrack function (window) active?
Dim hwndOldOwner As Long


Private Sub chkColor_Click()
    ChangeCaption True
End Sub

Private Sub chkColor_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If

End Sub

Private Sub chkDisqualified_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   Dim iKey As Integer
   
   If Button = vbRightButton Then
        StartMenuPopUp
   Else
        Select Case chkDisqualified.Value
        Case 0
           iKey = MsgBox(Translate("Eliminate this participant?", mcLanguage), vbYesNo + vbQuestion)
           If iKey = vbYes Then
              chkDisqualified.Value = 1
              chkWithdrawn.Value = 0
              chkWithdrawn.Enabled = False
            
              ParticipantDisqWith Me.txtParticipant, TestCode, TestStatus, -1
              Call cmdOK_Click
           
           Else
              chkDisqualified.Value = 0
              chkWithdrawn.Enabled = True
           End If
        Case Else
           iKey = MsgBox(Translate("Remove elimination for this participant?", mcLanguage), vbYesNo + vbQuestion)
           If iKey = vbYes Then
              chkDisqualified.Value = 0
              chkWithdrawn.Enabled = True
              ParticipantDisqWith Me.txtParticipant, TestCode, TestStatus, 0
              Call cmdOK_Click

           Else
              chkDisqualified.Value = 1
              chkWithdrawn.Value = 0
              chkWithdrawn.Enabled = False
           End If
        End Select
        lblParticipant.Tag = "*"
    End If
End Sub
Private Sub chkNoStart_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   Dim iKey As Integer
   
   If Button = vbRightButton Then
        StartMenuPopUp
   Else
        Select Case chkNoStart.Value
        Case 0
           iKey = MsgBox(Translate("Will this participant not start in the next run?", mcLanguage), vbYesNo + vbQuestion)
           If iKey = vbYes Then
              chkNoStart.Value = 1
            
              ParticipantNoStart Me.txtParticipant, TestCode, TestStatus, -1
              
              Call cmdOK_Click
            Else
              chkNoStart.Value = 0
           End If
        Case Else
           iKey = MsgBox(Translate("Will this participant not start in the next run?", mcLanguage), vbYesNo + vbQuestion)
           If iKey = vbYes Then
              chkNoStart.Value = 0
              ParticipantNoStart Me.txtParticipant, TestCode, TestStatus, 0
              Call cmdOK_Click
           Else
              chkNoStart.Value = 1
           End If
        End Select
        lblParticipant.Tag = "*"
    End If
End Sub
Private Sub chkFeifId_Click()
    ChangeCaption True
End Sub

Private Sub chkFeifId_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub chkFlag_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
    Else
     lblParticipant.Tag = "*"
   End If
End Sub

Private Sub chkFlag_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If dtaTest.Recordset.Fields("Type_Special") = 2 Then
        ValidateScore
    Else
        ValidateTimeScore
    End If
    CalculateResult Format$(Val(txtParticipant.Text), "000")
End Sub


Private Sub chkSplitResultLists_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim iKey As Integer
    Dim cQry As String
    Dim rst As DAO.Recordset
    Dim cTemp As String
    
    If Button = vbRightButton Then
        StartMenuPopUp
    Else
        '* check if there are participants from different (age) classes
        '*
        cQry = "SELECT Count(Participants.Sta) AS Expr1,Participants.Class"
        cQry = cQry & " FROM Participants INNER JOIN Results ON Participants.STA = Results.STA"
        cQry = cQry & " WHERE Results.Code Like '" & TestCode & "' AND Results.Status = 0 AND Results.Disq > -1"
        cQry = cQry & " GROUP BY Participants.Class"
        cQry = cQry & " ORDER BY Participants.Class;"

        Set rst = mdbMain.OpenRecordset(cQry)
        If rst.RecordCount > 0 Then
            '* if so, show the classes and the number of participants
            '*
            Do While Not rst.EOF
                If rst.Fields(1) & "" = "" Then
                    If cTemp = "" Then
                        cTemp = TestCode & " (" & rst.Fields(0) & ")"
                    Else
                        cTemp = cTemp & ", " & TestCode & " (" & rst.Fields(0) & ")"
                    End If
                Else
                    If cTemp = "" Then
                        cTemp = TestCode & "-" & rst.Fields(1) & " (" & rst.Fields(0) & ")"
                    Else
                        cTemp = cTemp & ", " & TestCode & "-" & rst.Fields(1) & " (" & rst.Fields(0) & ")"
                    End If
                End If
                rst.MoveNext
            Loop
        End If
        
        rst.Close
        
        Select Case chkSplitResultLists.Value
        Case 0
            iKey = MsgBox(Translate("Split result lists into " & cTemp & "?", mcLanguage), vbQuestion + vbYesNo)
            If iKey = vbYes Then
                chkSplitFinals.Value = 1
                chkSplitResultLists.Value = 1
            Else
                '* If finals should not be split, check if participants have already been distributed over
                '* different finals. If so, remove entries and unblock preliminary rounds for those tests
                chkSplitFinals.Value = 0
                chkSplitResultLists.Value = 0
                mdbMain.Execute "DELETE * FROM Entries WHERE Status>0 AND Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "')"
                If mdbMain.RecordsAffected > 0 Then
                    MsgBox Translate("Please reprint result lists when needed.", mcLanguage), vbExclamation
                    Set rst = mdbMain.OpenRecordset("SELECT Handling FROM TestInfo WHERE Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "')")
                    If rst.RecordCount > 0 Then
                        Do While Not rst.EOF
                            With rst
                                If .Fields(0) > 2 And .Fields(0) < 5 Then
                                    .Edit
                                    .Fields(0) = .Fields(0) - 2
                                    .Update
                                ElseIf .Fields(0) = 6 Then
                                    .Edit
                                    .Fields(0) = 5
                                    .Update
                                End If
                                .MoveNext
                            End With
                        Loop
                    End If
                    rst.Close
                    ChangeCaption True
                End If
            End If
        Case Else
            iKey = MsgBox(Translate("Undo splitting of result lists into " & cTemp & "?", mcLanguage), vbQuestion + vbYesNo)
            If iKey = vbYes Then
                '* If finals should not be split, check if participants have already been distributed over
                '* different finals. If so, remove entries and unblock preliminary rounds for those tests
                chkSplitFinals.Value = 0
                chkSplitResultLists.Value = 0
                mdbMain.Execute "DELETE * FROM Entries WHERE Status>0 AND Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "')"
                If mdbMain.RecordsAffected > 0 Then
                    MsgBox Translate("Please reprint result lists when needed.", mcLanguage), vbExclamation
                    Set rst = mdbMain.OpenRecordset("SELECT Handling FROM TestInfo WHERE Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "')")
                    If rst.RecordCount > 0 Then
                        Do While Not rst.EOF
                            With rst
                                If .Fields(0) > 2 And .Fields(0) < 5 Then
                                    .Edit
                                    .Fields(0) = .Fields(0) - 2
                                    .Update
                                ElseIf .Fields(0) = 6 Then
                                    .Edit
                                    .Fields(0) = 5
                                    .Update
                                End If
                                .MoveNext
                            End With
                        Loop
                    End If
                    rst.Close
                    ChangeCaption True
                End If
           Else
                chkSplitFinals.Value = 1
                chkSplitResultLists.Value = 1
            End If
        End Select
        '* store value in TestInfo
        With dtaTestInfo.Recordset
            .Edit
            .Fields("SplitFinals") = chkSplitFinals.Value
            .Update
        End With
        Set rst = Nothing
        
        If chkSplitResultLists.Value = 1 Then
            '* let user decide how participants will be split
            '*
            frmSplitFinals.fcTestCode = TestCode
            frmSplitFinals.Show 1, Me
        End If
        
    End If
End Sub

Private Sub chkTeam_click()
    ChangeCaption True
End Sub
Private Sub chkRein_Click()
    ChangeCaption True
End Sub

Private Sub chkRein_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub chkSplitFinals_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim iKey As Integer
    Dim cQry As String
    Dim rst As DAO.Recordset
    Dim cTemp As String
    
    If Button = vbRightButton Then
        StartMenuPopUp
    Else
        '* check if there are participants from different (age) classes
        '*
        cQry = "SELECT Count(Participants.Sta) AS Expr1,Participants.Class"
        cQry = cQry & " FROM Participants INNER JOIN Results ON Participants.STA = Results.STA"
        cQry = cQry & " WHERE Results.Code Like '" & TestCode & "' AND Results.Status = 0 AND Results.Disq > -1"
        cQry = cQry & " GROUP BY Participants.Class"
        cQry = cQry & " ORDER BY Participants.Class;"

        Set rst = mdbMain.OpenRecordset(cQry)
        If rst.RecordCount > 0 Then
            '* if so, show the classes and the number of participants
            '*
            Do While Not rst.EOF
                If rst.Fields(1) & "" = "" Then
                    If cTemp = "" Then
                        cTemp = TestCode & " (" & rst.Fields(0) & ")"
                    Else
                        cTemp = cTemp & ", " & TestCode & " (" & rst.Fields(0) & ")"
                    End If
                Else
                    If cTemp = "" Then
                        cTemp = TestCode & "-" & rst.Fields(1) & " (" & rst.Fields(0) & ")"
                    Else
                        cTemp = cTemp & ", " & TestCode & "-" & rst.Fields(1) & " (" & rst.Fields(0) & ")"
                    End If
                End If
                rst.MoveNext
            Loop
        End If
        
        rst.Close
        
        Select Case chkSplitFinals.Value
        Case 0
            iKey = MsgBox(Translate("Split finals into " & cTemp & "?", mcLanguage), vbQuestion + vbYesNo)
            If iKey = vbYes Then
                chkSplitFinals.Value = 1
                chkSplitResultLists.Value = 1
            Else
                '* If finals should not be split, check if participants have already been distributed over
                '* different finals. If so, remove entries and unblock preliminary rounds for those tests
                chkSplitFinals.Value = 0
                chkSplitResultLists.Value = 0
                mdbMain.Execute "DELETE * FROM Entries WHERE Status>0 AND Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "')"
                If mdbMain.RecordsAffected > 0 Then
                    MsgBox Translate("Please re-compose finals.", mcLanguage), vbExclamation
                    Set rst = mdbMain.OpenRecordset("SELECT Handling FROM TestInfo WHERE Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "')")
                    If rst.RecordCount > 0 Then
                        Do While Not rst.EOF
                            With rst
                                If .Fields(0) > 2 And .Fields(0) < 5 Then
                                    .Edit
                                    .Fields(0) = .Fields(0) - 2
                                    .Update
                                ElseIf .Fields(0) = 6 Then
                                    .Edit
                                    .Fields(0) = 5
                                    .Update
                                End If
                                .MoveNext
                            End With
                        Loop
                    End If
                    rst.Close
                    ChangeCaption True
                End If
            End If
        Case Else
            iKey = MsgBox(Translate("Undo splitting of finals into " & cTemp & "?", mcLanguage), vbQuestion + vbYesNo)
            If iKey = vbYes Then
                '* If finals should not be split, check if participants have already been distributed over
                '* different finals. If so, remove entries and unblock preliminary rounds for those tests
                chkSplitFinals.Value = 0
                chkSplitResultLists.Value = 0
                mdbMain.Execute "DELETE * FROM Entries WHERE Status>0 AND Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "')"
                If mdbMain.RecordsAffected > 0 Then
                    MsgBox Translate("Please re-compose finals.", mcLanguage), vbExclamation
                    Set rst = mdbMain.OpenRecordset("SELECT Handling FROM TestInfo WHERE Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "')")
                    If rst.RecordCount > 0 Then
                        Do While Not rst.EOF
                            With rst
                                If .Fields(0) > 2 And .Fields(0) < 5 Then
                                    .Edit
                                    .Fields(0) = .Fields(0) - 2
                                    .Update
                                ElseIf .Fields(0) = 6 Then
                                    .Edit
                                    .Fields(0) = 5
                                    .Update
                                End If
                                .MoveNext
                            End With
                        Loop
                    End If
                    rst.Close
                    ChangeCaption True
                End If
           Else
                chkSplitFinals.Value = 1
                chkSplitResultLists.Value = 1
            End If
        End Select
        '* store value in TestInfo
        With dtaTestInfo.Recordset
            .Edit
            .Fields("SplitFinals") = chkSplitFinals.Value
            .Update
        End With
        
        Set rst = Nothing
    End If
End Sub
Private Sub chkTeam_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub chkWithdrawn_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   Dim iKey As Integer
   
   If Button = vbRightButton Then
        StartMenuPopUp
   Else
        Select Case chkWithdrawn.Value
        Case 0
           iKey = MsgBox(Translate("Does this participant withdraw?", mcLanguage), vbYesNo + vbQuestion)
           If iKey = vbYes Then
              chkWithdrawn.Value = 1
              ParticipantDisqWith Me.txtParticipant, TestCode, TestStatus, -2
              Call cmdOK_Click
              If TestStatus > 0 Then
                    If ComposeFinals(TestStatus) = True Then
                        If TestStatus = 3 Then
                            iKey = MsgBox(Translate("(Re-)Compose A-Final and B-Final as well?", mcLanguage), vbQuestion + vbYesNo)
                            If iKey = vbYes Then
                                ComposeFinals 2, vbYes
                                ComposeFinals 1, vbYes
                            End If
                        ElseIf TestStatus = 2 Then
                            iKey = MsgBox(Translate("(Re-)Compose A-Final as well?", mcLanguage), vbQuestion + vbYesNo)
                            If iKey = vbYes Then
                                ComposeFinals 1, vbYes
                            End If
                        Else
                            iKey = MsgBox(Translate("(Re-)Compose B-Final as well?", mcLanguage), vbQuestion + vbYesNo)
                            If iKey = vbYes Then
                                ComposeFinals 2, vbYes
                            End If
                        End If
                    End If
               End If
           Else
              chkWithdrawn.Value = 0
           End If
        Case Else
           iKey = MsgBox(Translate("Remove withdrawal for this participant?", mcLanguage), vbYesNo + vbQuestion)
           If iKey = vbYes Then
              chkWithdrawn.Value = 0
              ParticipantDisqWith Me.txtParticipant, TestCode, TestStatus, 0
              Call cmdOK_Click
              mdbMain.Execute ("DELETE * FROM Marks WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND STA='" & Me.txtParticipant & "'")
              mdbMain.Execute ("DELETE * FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND STA='" & Me.txtParticipant & "'")
              ClearMarks True
              txtParticipant = ""
              ChangeCaption True
              If TestStatus > 0 Then
                    If ComposeFinals(TestStatus) = True Then
                        If TestStatus = 3 Then
                            iKey = MsgBox(Translate("(Re-)Compose A-Final and B-Final as well?", mcLanguage), vbQuestion + vbYesNo)
                            If iKey = vbYes Then
                                ComposeFinals 2, vbYes
                                ComposeFinals 1, vbYes
                            End If
                        ElseIf TestStatus = 2 Then
                            iKey = MsgBox(Translate("(Re-)Compose A-Final as well?", mcLanguage), vbQuestion + vbYesNo)
                            If iKey = vbYes Then
                                ComposeFinals 1, vbYes
                            End If
                        Else
                            iKey = MsgBox(Translate("(Re-)Compose B-Final as well?", mcLanguage), vbQuestion + vbYesNo)
                            If iKey = vbYes Then
                                ComposeFinals 2, vbYes
                            End If
                        End If
                    End If
               End If
           Else
              chkWithdrawn.Value = 1
           End If
        End Select
        lblParticipant.Tag = "*"
    End If
End Sub


Private Sub cmbGroupSize_Change()
    If Val(cmbGroupSize.Text) > 100 Then
        cmbGroupSize = 100
    End If
End Sub

Private Sub cmbNumJudges_Change(Index As Integer)
   If Val(cmbNumJudges(Index).Text) > 5 Then
        cmbNumJudges(Index).Text = "5"
   ElseIf Index > 0 And Val(cmbNumJudges(Index).Text) < Val(cmbNumJudges(0).Text) Then
        cmbNumJudges(Index).Text = cmbNumJudges(0).Text
   End If
   cmbNumJudges(Index).Refresh
End Sub
Private Sub cmbNumJudges_Click(Index As Integer)
   Dim cNumJudges As String
   
   If Index > 0 Then
        If Val(cmbNumJudges(Index).Text) < Val(cmbNumJudges(0).Text) Then
            With dtaTestInfo.Recordset
                .Edit
                .Fields("Num_j_" & Format$(Index)).Value = .Fields("Num_j_0").Value
                .Update
            End With
            cmbNumJudges(Index).Refresh
        End If
   End If
   cNumJudges = cmbNumJudges(Index).Text
   
   
   If Index = 1 Then
        cmbNumJudges(2).Text = Trim$(cmbNumJudges(1).Text)
   ElseIf Index = 2 Then
        cmbNumJudges(1).Text = Trim$(cmbNumJudges(2).Text)
   ElseIf Index = 3 Then
        cmbNumJudges(1).Text = Trim$(cmbNumJudges(3).Text)
   End If
   
   If Index = 0 Then
        With dtaTestInfo.Recordset
            .Edit
            .Fields("Num_j_1").Value = .Fields("Num_j_0").Value
            .Fields("Num_j_2").Value = .Fields("Num_j_0").Value
            .Fields("Num_j_3").Value = .Fields("Num_j_0").Value
            .Update
        End With
        cmbNumJudges(1).Text = cmbNumJudges(0).Text
        cmbNumJudges(2).Text = cmbNumJudges(0).Text
        cmbNumJudges(3).Text = cmbNumJudges(0).Text
   End If
   
   StoreCurrentMarks
   
   LookUpTest
   
   ChangeCaption True
   
   cmbNumJudges(Index).Text = cNumJudges
   
   Select Case Index
   Case Is = 0
        tbsSelFin.SelectedItem = frmMain.tbsSelFin.Tabs(1)
        SetFocusTo frmMain.tbsSelFin.Tabs(1)
   Case Is = 1
        tbsSelFin.SelectedItem = frmMain.tbsSelFin.Tabs(frmMain.tbsSelFin.Tabs.Count - 1)
        SetFocusTo frmMain.tbsSelFin.Tabs(frmMain.tbsSelFin.Tabs.Count - 1)
   Case Is = 2
        tbsSelFin.SelectedItem = frmMain.tbsSelFin.Tabs(frmMain.tbsSelFin.Tabs.Count - 2)
        SetFocusTo frmMain.tbsSelFin.Tabs(frmMain.tbsSelFin.Tabs.Count - 2)
   Case Is = 2
        tbsSelFin.SelectedItem = frmMain.tbsSelFin.Tabs(frmMain.tbsSelFin.Tabs.Count - 3)
        SetFocusTo frmMain.tbsSelFin.Tabs(frmMain.tbsSelFin.Tabs.Count - 3)
   End Select
   
   DoEvents
   
   

End Sub

Private Sub cmdComposeFinals_Click()
    Dim iKey As Integer
    
    If chkSplitFinals.Value = 1 And chkSplitFinals.Enabled = True Then
        '* let user decide how participants will be split
        '*
        frmSplitFinals.fcTestCode = TestCode
        frmSplitFinals.Show 1, Me
    End If

    If ComposeFinals(TestStatus) = True Then
        If TestStatus = 3 Then
            iKey = MsgBox(Translate("(Re-)Compose A-Final and B-Final as well?", mcLanguage), vbQuestion + vbYesNo)
            If iKey = vbYes Then
                ComposeFinals 2, vbYes
                ComposeFinals 1, vbYes
            End If
        ElseIf TestStatus = 2 Then
            iKey = MsgBox(Translate("(Re-)Compose A-Final as well?", mcLanguage), vbQuestion + vbYesNo)
            If iKey = vbYes Then
                ComposeFinals 1, vbYes
            End If
        Else
            iKey = MsgBox(Translate("(Re-)Compose B-Final as well?", mcLanguage), vbQuestion + vbYesNo)
            If iKey = vbYes Then
                ComposeFinals 2, vbYes
            End If
        End If
    End If

End Sub

Private Sub cmdComposeFinals_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub cmdComposeGroups_Click()
    Dim iKey As Integer
    Dim iTemp As Integer
    Dim iGroupsize As Integer
    Dim iRRGroupsize As Integer
    Dim iLRGroupsize As Integer
    Dim iRR As Integer
    Dim iRRMax As Integer
    Dim iRRLast As Integer
    Dim iRRGroupNum As Integer
    Dim iLR As Integer
    Dim iLRMax As Integer
    Dim iLRLast As Integer
    Dim iLRGroupNum As Integer
    Dim iNextGroupNum As Integer
    Dim iGroupcount As Integer
    
    Dim cColor() As String
    Dim cQry As String
    
    Dim iColor As String
    Dim iGroup As Integer
    Dim iOldGroup As Integer
    
    Dim rstEntry As DAO.Recordset
    
    iGroupsize = Val(cmbGroupSize)
        
    If Val(iGroupsize) < 0 Then
        MsgBox Translate("Set group size first!", mcLanguage), vbExclamation
        Exit Sub
    ElseIf dtaNotYet.Recordset.RecordCount = 0 Then
        MsgBox Translate("Enter participants first!", mcLanguage), vbExclamation
        Exit Sub
    Else
        If cmbGroupSize = "" Then
            'make sure that the basic groups are set (to avoid null values)
            iKey = vbYes
            cmbGroupSize.Text = "1"
            iGroupsize = 1
        ElseIf iGroupsize <= 1 Then
            iGroupsize = 1
            iKey = MsgBox(Translate("Remove groups and/or colors?", mcLanguage), vbQuestion + vbYesNo)
        Else
            If iGroupsize > 10 Then
                iGroupsize = 10
            End If
            chkColor.Value = 1
            iKey = MsgBox(Translate("Compose groups with a size of", mcLanguage) & " " & cmbGroupSize.Text & "?", vbQuestion + vbYesNo)
        End If
        
        If iKey = vbYes Then
            SetMouseHourGlass
            cmbGroupSize.Enabled = False
            
            'count participants
            If fraTime.Visible = True Then
                cQry = "SELECT Entries.*,Alltimes,Disq "
                cQry = cQry & " FROM Entries "
                cQry = cQry & " LEFT JOIN Results ON (Entries.STA = Results.STA) AND (Entries.Code = Results.Code)"
                cQry = cQry & " WHERE Entries.Code='" & TestCode & "' "
                cQry = cQry & " AND (IsNull(Disq) Or Disq = 0)"
                cQry = cQry & " AND (IsNull(NoStart) Or NoStart = 0)"
                cQry = cQry & " ORDER BY ISNULL(Results.AllTimes), Results.AllTimes DESC"
                
                Set rstEntry = mdbMain.OpenRecordset(cQry)
                If rstEntry.RecordCount > 0 Then
                    iLRMax = rstEntry.RecordCount
                    iRRMax = 0
                End If
                dtaNotYet.Refresh
            Else
                Set rstEntry = mdbMain.OpenRecordset("SELECT * From Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND NOT RR=0 AND NOT STA IN (SELECT STA FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ")")
                If rstEntry.RecordCount > 0 Then
                    rstEntry.MoveLast
                    iRRMax = rstEntry.RecordCount
                End If
                Set rstEntry = mdbMain.OpenRecordset("SELECT * From Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND RR=0 AND NOT STA IN (SELECT STA FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ")")
                If rstEntry.RecordCount > 0 Then
                    rstEntry.MoveLast
                    iLRMax = rstEntry.RecordCount
                End If
            End If
            
            iRRLast = iGroupsize
            iRRGroupsize = iGroupsize
            If iRRMax Mod iRRGroupsize <> 0 Then
                If iRRMax < iRRGroupsize Then
                    iRRGroupsize = iRRMax
                    iRRLast = iRRMax
                ElseIf iRRMax < iRRGroupsize * 2 And iRRMax > iRRGroupsize Then
                    iRRGroupsize = (iRRMax + 1) \ 2
                    iRRLast = iRRMax - iRRGroupsize
                ElseIf iRRMax < iRRGroupsize * 3 And iRRMax > iRRGroupsize * 2 Then
                    iRRGroupsize = (iRRMax + 2) \ 3
                    iRRLast = iRRMax \ 3
                Else
                    iTemp = iRRMax \ iRRGroupsize + 1
                    If iTemp <= 3 Then
                        iRRGroupsize = (iRRMax + 0.5) / iTemp
                    End If
                    iRRLast = (iRRMax Mod iRRGroupsize + iRRGroupsize) \ 2
                End If
                iNextGroupNum = 1 + iNextGroupNum + iRRMax \ iRRGroupsize
            Else
                If iRRMax > 0 Then
                    iNextGroupNum = iNextGroupNum + iRRMax \ iRRGroupsize
                End If
            End If
            
            iLRLast = iGroupsize
            iLRGroupsize = iGroupsize
            If iLRMax Mod iLRGroupsize <> 0 Then
                If iLRMax < iLRGroupsize Then
                    iLRGroupsize = iLRMax
                    iLRLast = iLRMax
                ElseIf iLRMax < iLRGroupsize * 2 - 1 And iLRMax > iLRGroupsize Then
                    iLRGroupsize = (iLRMax + 1) \ 2
                    iLRLast = iLRMax \ 2
                ElseIf iLRMax < iLRGroupsize * 3 - 1 And iLRMax > iLRGroupsize * 2 Then
                    iLRGroupsize = (iLRMax + 2) \ 3
                    iLRLast = iLRMax \ 3
                Else
                    iTemp = iLRMax \ iLRGroupsize + 1
                    If iTemp <= 3 Then
                        iLRGroupsize = (iLRMax + 0.5) / iTemp
                    End If
                    iLRLast = (iLRMax Mod iLRGroupsize + iLRGroupsize) \ 2
                End If
                iNextGroupNum = 1 + iNextGroupNum + iLRMax \ iLRGroupsize
            Else
                iNextGroupNum = iNextGroupNum + iLRMax \ iLRGroupsize
            End If
            cColor = Split(TestColors, ",")
            
            
            dtaNotYet.Recordset.MoveLast
            Do While Not dtaNotYet.Recordset.BOF
                cQry = "SELECT * From Entries "
                cQry = cQry & " WHERE STA='" & dtaNotYet.Recordset.Fields("STA") & "' "
                cQry = cQry & " AND Code='" & TestCode & "' "
                cQry = cQry & " AND Status=" & TestStatus
                cQry = cQry & " AND (IsNull(NoStart) Or NoStart = 0)"
                Set rstEntry = mdbMain.OpenRecordset(cQry)
                If rstEntry.RecordCount > 0 Then
                    rstEntry.Edit
                    If rstEntry.Fields("RR") = True Then
                        iRR = iRR + 1
                        If iRR Mod iRRGroupsize = 1 Or iRRMax - iRR = iRRLast - 1 Then
                            If iRRMax - iRR >= iRRLast - 1 Then
                                iRRGroupNum = iNextGroupNum
                                iNextGroupNum = iNextGroupNum - 1
                            End If
                        End If
                        rstEntry.Fields("Group") = iRRGroupNum
                    Else
                        iLR = iLR + 1
                        If iLR Mod iLRGroupsize = 1 Or iLRMax - iLR = iLRLast - 1 Then
                            If iLRMax - iLR >= iLRLast - 1 Then
                                iLRGroupNum = iNextGroupNum
                                iNextGroupNum = iNextGroupNum - 1
                            End If
                        End If
                        rstEntry.Fields("Group") = iLRGroupNum
                    End If
                    rstEntry.Update
                End If
                dtaNotYet.Recordset.MovePrevious
            Loop
            
            If fraTime.Visible = True Then
                Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND NOT STA IN (SELECT STA FROM RESULTS WHERE Code='" & TestCode & "' AND Disq=-1) ORDER BY Group,Position")
            Else
                Set rstEntry = mdbMain.OpenRecordset("SELECT * From Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND NOT STA IN (SELECT STA FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ") ORDER BY Group,Position")
            End If
            If rstEntry.RecordCount > 0 Then
                iColor = 0
                iGroup = 0
                iOldGroup = 0
                Do While Not rstEntry.EOF
                    rstEntry.Edit
                    If iGroupsize > 1 Then
                        If rstEntry.Fields("Group") <> iOldGroup Then
                            iGroup = iGroup + 1
                            iColor = 0
                            iOldGroup = rstEntry.Fields("Group")
                        End If
                        rstEntry.Fields("Group") = iGroup
                        If iColor <= UBound(cColor) Then
                            rstEntry.Fields("Color") = Left$(cColor(iColor), rstEntry.Fields("Color").Size)
                        Else
                            rstEntry.Fields("Color") = mcNoColor
                        End If
                    Else
                        rstEntry.Fields("Group") = 0
                        rstEntry.Fields("Color") = ""
                    End If
                    rstEntry.Update
                    rstEntry.MoveNext
                    iColor = iColor + 1
                Loop
            End If
            
            dtaNotYet.Refresh
            rstEntry.Close
            Set rstEntry = Nothing
        End If
    End If
    SetMouseNormal
    cmbGroupSize.Enabled = True
End Sub

Private Sub cmdComposeGroups_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub cmdIceSort_Click()
    Dim sortApp As String
    Dim sortCmdLine As String
    
    On Local Error GoTo cmdIceSort_ClickError
    If TestCode <> "" Then
        sortCmdLine = "/TESTCODE:" & TestCode
    End If
    If Not IsNull(TestStatus) Then
      sortCmdLine = sortCmdLine & " " & "/TESTSTATUS:" & TestStatus
    End If
    sortCmdLine = sortCmdLine & "  " & "/USEDB:" & mcDatabaseName
    sortApp = App.Path & "\" & SortingApplication
    
    Call Shell(sortApp & " " & sortCmdLine, vbNormalFocus)

    Exit Sub
cmdIceSort_ClickError:
    If Err > 0 Then
        If (sortApp = "") Then
                LogLine Err.Source & ": " & Err.Number & ": " & Err.Description
                MsgBox App.EXEName & ". Could not find " & SortingApplication & " in '" & App.Path & "'.", vbCritical
        End If
    End If
End Sub

Private Sub cmdInfo_Click()
   Dim cTemp As String
   Dim ctlActive As Control
   
   If mnuEditEdit.Visible = True And mnuEditEdit.Enabled = True Then
        mnuEditEdit_Click
   Else
   
        With dtaParticipant.Recordset
           cTemp = cTemp & .Fields("Name_First")
           If .Fields("Name_Middle") & "" <> "" Then
              cTemp = cTemp & " " & .Fields("Name_Middle")
           End If
           cTemp = cTemp & " " & .Fields("Name_Last")
           If .Fields("Class") & "" <> "" Then
               cTemp = cTemp & " [" & .Fields("Class") & "]"
           End If
           If .Fields("Club") & "" <> "" Then
               cTemp = cTemp & " / " & .Fields("Club")
           End If
           If .Fields("Team") & "" <> "" Then
               cTemp = cTemp & " / " & .Fields("Team")
           End If
           
           cTemp = cTemp & vbCrLf & "-----"
           cTemp = cTemp & vbCrLf & .Fields("Name_Horse")
           Select Case .Fields("Sex_Horse")
              Case 1
                 cTemp = cTemp & ", " & Translate("Stallion", mcLanguage)
              Case 2
                 cTemp = cTemp & ", " & Translate("Mare", mcLanguage)
              Case 3
                 cTemp = cTemp & ", " & Translate("Gelding", mcLanguage)
              Case Else
                 cTemp = cTemp & ", --"
              End Select
           cTemp = cTemp & ", " & Format$(.Fields("Birthday_horse"), "YYYY")
           cTemp = cTemp & " (" & .Fields("Country_Horse") & ")"
           cTemp = cTemp & ", " & .Fields("Color")
           If .Fields("FEIFID") & "" <> "" Then
               cTemp = cTemp & ", " & .Fields("FEIFID")
           Else
               cTemp = cTemp & ", " & .Fields("HorseID")
           End If
           cTemp = cTemp & vbCrLf & "F: " & .Fields("F")
           cTemp = cTemp & vbCrLf & vbTab & "-FF: " & .Fields("FF")
           cTemp = cTemp & vbCrLf & vbTab & "-FM: " & .Fields("FM")
           cTemp = cTemp & vbCrLf & "M: " & .Fields("M")
           cTemp = cTemp & vbCrLf & vbTab & "-MF: " & .Fields("MF")
           cTemp = cTemp & vbCrLf & vbTab & "-MM: " & .Fields("MM")
           cTemp = cTemp & vbCrLf & Translate("Breeder", mcLanguage) & ": " & .Fields("Breeder")
           cTemp = cTemp & vbCrLf & Translate("Owner", mcLanguage) & ": " & .Fields("Owner")
           
        End With
        MsgBox cTemp, , ClipAmp(fraCurrent.Caption)
    End If
    SetFocusTo txtParticipant
    txtParticipant_Change
End Sub

Private Sub cmdInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub cmdMarks_Click()
    cmdMarks.Enabled = False
    If Dir$(App.Path & "\icemarks.exe") <> "" Then
        Shell App.Path & "\icemarks.exe Test=" & TestCode & " Status=" & TestStatus & " Judges=" & TestJudges, vbNormalFocus
    ElseIf Dir$(Replace(App.Path, "IceTest", "IceHorseTools") & "\icemarks.exe") <> "" Then
        Shell Replace(App.Path, "IceTest", "IceHorseTools") & "\icemarks.exe Test=" & TestCode & " Status=" & TestStatus & " Judges=" & TestJudges, vbNormalFocus
    End If
    cmdMarks.Enabled = True
End Sub

Private Sub cmdOK_Click()
    cmdOkClick
End Sub

Private Sub cmdOk_GotFocus()
   If txtScore.BackColor <> mlAlertColor Then
        txtScore.BackColor = mlAlertColor
        miNoBackupNow = True
   End If
End Sub

Private Sub cmdOK_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
       If fraMarks.Visible = True Then
            SetFocusTo txtMarks(TestJudges - 1)
        ElseIf fraTime.Visible = True Then
            SetFocusTo txtTime
        End If
    End If

End Sub

Private Sub cmdOK_LostFocus()
    txtScore.BackColor = QBColor(15)
    miNoBackupNow = False
End Sub

Private Sub cmdOK_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub cmdTestInfo_Click()
    frmTestInfo.Show 1, Me
End Sub


Private Sub dblstAlready_Click()
   Dim cTemp As String
   cTemp = dblstAlready.BoundText
   dblstNotYet.BoundText = ""
   dblstAlready.BoundText = cTemp
End Sub

Private Sub dblstAlready_DblClick()
   Dim cTemp As String
   cTemp = dblstAlready.BoundText
   StoreCurrentMarks
   txtParticipant.Text = cTemp
   dblstNotYet.BoundText = ""
   dblstAlready.BoundText = cTemp
   LookUpParticipant
End Sub

Private Sub dblstAlready_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      dblstAlready_DblClick
   End If
End Sub

Private Sub dblstAlready_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub dblstNotYet_Click()
    Dim cTemp As String
    Dim ontrackexists As Boolean
    Dim icountontrack As Integer
    
    cTemp = dblstNotYet.BoundText
    
    If miSelectBySingleClick = True Then
       StoreCurrentMarks
       dblstAlready.Text = ""
       txtParticipant.Text = cTemp
    End If
    
    txtMove.Text = ""
    lstAlready.ListIndex = -1
    If dblstNotYet.BoundText <> cTemp Then
        dblstNotYet.BoundText = cTemp
    End If
    
    'If ontrack info is running, make data available to it:
    If ontrack = True Then
        ontrackexists = False
        For icountontrack = 0 To frmOntrack.lstOntrack.ListCount
            If frmOntrack.lstOntrack.List(icountontrack) = cTemp Then
                ontrackexists = True
            End If
        Next icountontrack
        If ontrackexists = False Then
            frmOntrack.lstOntrack.AddItem cTemp
        End If
    End If
    
    If miSelectBySingleClick = True Then
       LookUpParticipant
    End If
End Sub

Private Sub dblstNotYet_DblClick()
    Dim cTemp As String
    
    cTemp = dblstNotYet.BoundText
    
    StoreCurrentMarks
    dblstAlready.Text = ""
    txtParticipant.Text = cTemp
    
    txtMove.Text = ""
    lstAlready.ListIndex = -1
    If dblstNotYet.BoundText <> cTemp Then
        dblstNotYet.BoundText = cTemp
    End If
    
    LookUpParticipant
End Sub

Private Sub dblstNotYet_GotFocus()
    mnuEditMove.Enabled = True
    Me.mnuEditChangeRein.Enabled = Me.chkRein.Enabled
End Sub

Private Sub dblstNotYet_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then
      dblstNotYet_DblClick
   End If
End Sub

Private Sub dblstNotYet_LostFocus()
    txtMove.Text = ""
    dblstNotYet.Tag = ""
    txtMove.Visible = False
    mnuEditMove.Enabled = False
    Me.mnuEditChangeRein.Enabled = False
End Sub

Private Sub dblstNotYet_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   ElseIf Button = vbLeftButton Then
      If dblstNotYet.Tag = "" Then
         dblstNotYet.Tag = y
        DoEvents
      End If
   End If
End Sub

Private Sub dblstNotYet_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then
        If dblstNotYet.Tag <> "" Then
            If txtMove.Text = "" And (dblstNotYet.Tag < y - 50 Or dblstNotYet.Tag > y + 50) Then
                txtMove.Text = dblstNotYet.BoundText
                txtMove.Visible = True
            End If
            txtMove.Top = y + 100
            DoEvents
        End If
    End If
End Sub

Private Sub dblstNotYet_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    Dim iKey As Integer
    Dim cMsg As String
    Dim iNewPosition As Integer
    Dim iPosition As Integer
    Dim iColor As Integer
    Dim cColor() As String
    Dim rstEntry As DAO.Recordset
    Dim iNewGroup As Integer
    Dim iOldGroup As Integer
    Dim iOldHand As Integer
    
    iPosition = 0
    If dblstNotYet.Tag <> "" And txtMove.Text <> "" Then
        If txtMove.Text <> dblstNotYet.BoundText Then
            cMsg = UCase$(Translate("Move", mcLanguage)) & " '" & txtMove.Text & "' " & UCase$(Translate("ahead of", mcLanguage)) & " '" & dblstNotYet.BoundText & "'?" & vbCrLf
            cMsg = cMsg & vbCrLf & "- " & Translate("Select 'Yes' to move the participant AHEAD OF the other participant.", mcLanguage)
            cMsg = cMsg & vbCrLf & "- " & Translate("Select 'No' to move the participant BEHIND the other participant.", mcLanguage)
            cMsg = cMsg & vbCrLf & "- " & Translate("Select 'Cancel' to CANCEL this action.", mcLanguage)
            iKey = MsgBox(cMsg, vbYesNoCancel + vbQuestion + vbDefaultButton1)
            iNewPosition = -1
            Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND NOT Sta IN (SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ") ORDER BY Position")
            If rstEntry.RecordCount > 0 Then
                With rstEntry
                    Do While Not .EOF
                        .Edit
                        .Fields("Position") = (.AbsolutePosition + 1) * 2
                        .Update
                        .MoveNext
                    Loop
                End With
            End If
            If iKey = vbYes Then
                Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Sta='" & Left$(dblstNotYet.BoundText, 3) & "'")
                If rstEntry.RecordCount > 0 Then
                    iNewPosition = rstEntry.Fields("Position") - 1
                    iNewGroup = rstEntry.Fields("Group")
                End If
            ElseIf iKey = vbNo Then
                Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Sta='" & Left$(dblstNotYet.BoundText, 3) & "'")
                If rstEntry.RecordCount > 0 Then
                    iNewPosition = rstEntry.Fields("Position") + 1
                    iNewGroup = rstEntry.Fields("Group")
                End If
            End If
            If iNewPosition <> -1 Then
                Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Sta='" & Left$(txtMove.Text, 3) & "'")
                If rstEntry.RecordCount > 0 Then
                    With rstEntry
                        .Edit
                        .Fields("Position") = iNewPosition
                        .Fields("Group") = iNewGroup
                        .Update
                    End With
                End If
                Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND NOT Sta IN (SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ") ORDER BY Position")
                If rstEntry.RecordCount > 0 Then
                    With rstEntry
                        Do While Not .EOF
                            .Edit
                            .Fields("Position") = (.AbsolutePosition + 1) * 2
                            .Update
                            .MoveNext
                        Loop
                    End With
                End If
                If TestStatus > 0 Then
                    AddColorsToFinals
                Else
                    CorrectColorsAndGroups
                End If
            End If
            rstEntry.Close
            Set rstEntry = Nothing
            dtaNotYet.Refresh
        End If
    End If
    txtMove.Text = ""
    dblstNotYet.Tag = ""
    txtMove.Visible = False
End Sub

Private Sub Form_Load()
    Dim cTemp As String
    Dim dFirst As Date
    Dim cFIPO As String
    Dim cLang As String
    Dim iItem As Integer
    Dim iTemp As Integer
    Dim iKey As Integer
    Dim cMsg As String
    Dim cPrevQual As String
    Dim cDriveList As String
    Dim cDrive As String
    Dim rstTest As DAO.Recordset
    
    mcVersionSwitch = VersionSwitch
    
    Set fSplash = New frmSplash
    If Command$ = "" Then
       fSplash.Show
       DoEvents
    End If
    
    SetMouseHourGlass
    
    'what inifile is going to be used
    gcIniFile = TestIniFile(App.EXEName)
    gcIniHorseFile = TestIniFile("ICEHORSE")
    
    'is there a command line switch to set the test?
    If Me.TestCode = "" Then
         cTemp = ParseCommand("Test=")
         Me.TestCode = cTemp
    End If
    If Me.TestCode = "" Then
       ReadIniFile gcIniFile, Me.Name, "LastTest", cTemp
       Me.TestCode = cTemp
    End If
    
    'is there a command line switch to set the status?
    cTemp = ParseCommand("Status=")
    If cTemp <> "" Then
         Me.TestStatus = Val(cTemp)
    End If
        
    On Local Error Resume Next
    
    'what is the language used
    ReadIniFile gcIniFile, Me.Name, "Language", cTemp
    If cTemp = "" Then
        cTemp = "English"
        WriteIniFile gcIniFile, Me.Name, "Language", cTemp
    End If
    mcLanguage = cTemp
    
    If Dir$(App.Path) & "\Languages.Mdb" <> "" And Dir$(Environ$("APPDATA") & "\IceHorse\" & "Languages.Mdb") = "" Then
        FileCopy Dir$(App.Path) & "\Languages.Mdb", Dir$(Environ$("APPDATA") & "\IceHorse\" & "Languages.Mdb")
    End If

    ReadIniFile gcIniHorseFile, "Database", "Language", cTemp
    If cTemp <> "" And Right$(cTemp, 1) <> "\" Then
        cTemp = cTemp & "\"
    End If
    
    If cTemp <> "" And Dir$(cTemp & "Languages.Mdb") = "" And Dir$(App.Path & "\Languages.mdb") <> "" Then
        FileCopy App.Path & "\Languages.mdb", cTemp & "Languages.Mdb"
    End If
    
    If cTemp <> "" And Dir$(cTemp & "Languages.Mdb") <> "" Then
        InitializeLanguageDB mcLanguage, cTemp
    ElseIf Dir$(Environ$("APPDATA") & "\IceHorse\" & "Languages.Mdb") <> "" Then
        InitializeLanguageDB mcLanguage, Environ$("APPDATA") & "\IceHorse\"
        WriteIniFile gcIniHorseFile, "Database", "Language", Environ$("APPDATA") & "\IceHorse\"
    ElseIf Dir$(Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Languages.Mdb") <> "" Then
        InitializeLanguageDB mcLanguage, Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\"))
        WriteIniFile gcIniHorseFile, "Database", "Language", Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\"))
    Else
        MsgBox Translate("Please be sure to copy 'Languages.Mdb' first to", mcLanguage) & ": '" & Environ$("APPDATA") & "\IceHorse\" & "'."
        Unload Me
        End
    End If
        
    'find the database to use
    ReadIniFile gcIniHorseFile, "Database", "Folder", mcDatabaseName
    If mcDatabaseName = "" Or Dir$(mcDatabaseName) = "" Then
        'create folder
                
        If mcDatabaseName <> "" Then
            cTemp = Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\"))
        Else
            cTemp = "C:\IceHorse\"
        End If
        If Dir$(cTemp, vbDirectory) = "" Then
            If Err = 0 Then
                MkDir cTemp
                If Err > 0 Then
                    cTemp = Environ$("APPDATA") & "\IceHorse\"
                    MkDir cTemp
                    If Err > 0 Then
                       cTemp = ""
                    End If
                End If
            Else
                cTemp = Environ$("APPDATA") & "\IceHorse\"
                MkDir cTemp
                If Err > 0 Then
                   cTemp = ""
                End If
            End If
        End If
        
        'instruct user
        If cTemp = "" Then
            cMsg = Translate("needs a specific folder to store the databases for different events. Each event is stored in a separate database. You are advised to create a folder first (like 'C:\ICEHORSE\').", mcLanguage)
            cMsg = cMsg & vbCrLf & Translate("After you have selected or created a specific folder, you should give the database for this specific event its own name, like:", mcLanguage) & " " & UserName & "1.Mdb"
            MsgBox App.EXEName & " " & cMsg
        Else
            cMsg = Translate("will install the database for each event in", mcLanguage) & ": " & cTemp & "."
            cMsg = cMsg & vbCrLf & Translate("After you have selected or created a specific folder, you should give the database for this specific event its own name, like:", mcLanguage) & " " & UserName & "1.Mdb"
            MsgBox App.EXEName & " " & cMsg
        End If
        
        'see if a file is available on Cd-Rom
        Call GetDrives(cDriveList, "5")
        If cDriveList <> "" Then
            Parse cDrive, cDriveList, " "
            If Dir$(cDrive & "*.Mdb") <> "" Then
                If Err = 0 Then
                    ChangeDatabase cTemp, Dir$(cDrive & "*.Mdb")
                    cTemp = cDrive & Dir$(cDrive & "*.Mdb")
                    If cTemp <> "" Then
                        iKey = MsgBox(Translate("Do you want to read data from " + cDrive, mcLanguage) & " (" & cTemp & ")?", vbYesNo + vbQuestion)
                        If iKey = vbYes Then
                            mdbMain.Close
                            DoEvents
                            
                            FileCopy cTemp, mcDatabaseName
                            If Err > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical
                            End If
                            AttribNormal mcDatabaseName
                            Me.RestartApp
                        Else
                            ChangeDatabase cTemp
                        End If
                    Else
                        ChangeDatabase cTemp
                    End If
                Else
                    ChangeDatabase cTemp
                End If
            Else
                ChangeDatabase cTemp
            End If
        Else
            ChangeDatabase cTemp
        End If
    End If
    
    OpenDatabase mcDatabaseName
    
    'check if FIPO needs update
    If Dir$(App.Path & "\Fipo.Mdb") <> "" Then
        cFIPO = App.Path & "\Fipo.Mdb"
    Else
        ReadIniFile gcIniFile, "Import", "FIPO", cFIPO
    End If
    If cFIPO <> "" Then
        If Dir$(cFIPO) <> "" Then
            cTemp = GetVariable("FIPO")
            If cTemp <> "" Then
                If CDate(cTemp) < FileDateTime(cFIPO) Then
                    iKey = MsgBox(Translate("A new set of Sport Rules is available", mcLanguage) & " (" & Format$(FileDateTime(cFIPO), "dd-mm-yyyy") & "). " & Translate("Update now?", mcLanguage), vbQuestion + vbYesNo)
                    If iKey = vbYes Then
                        WriteIniFile gcIniFile, "Import", "FIPO", cFIPO
                        ProcessFipo cFIPO
                    End If
                ElseIf CDate(cTemp) <= Now - 365 Then
                    MsgBox Translate("Don't forget to download a new set of Sport Rules every year!", mcLanguage) & vbCrLf & Translate("(This should be done around April 1.)", mcLanguage)
                End If
            End If
        End If
    End If
    
    'check if database needs update
    cTemp = GetVariable("ProgramVersion")
    If cTemp <> App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") Or InStr(Command$, "/SYS") Then
        mdbMain.Close
        DoEvents
        If CheckDatabase(mcDatabaseName) = True Then
            OpenDatabase mcDatabaseName
            cTemp = App.Major & "." & App.Minor & "." & Format$(App.Revision, "000")
            SetVariable "ProgramVersion", cTemp
        Else
            OpenDatabase mcDatabaseName
        End If
        CheckOnIndexes
    End If
    
    'check if the variable number of judges has been implemented
    CheckNumJ
    
    CheckLanguageDb
    
    'MM-check if this is a beta version
    If App.Revision Mod 5 <> 0 Then
        cTemp = App.Major & "." & App.Minor & "." & Format$(App.Revision, "000")
        MsgBox "This is a test version." & vbCrLf & vbCrLf & "Please report any remarks to icetest@feif.org ," & vbCrLf & "including version number " & cTemp & " .", vbExclamation
    End If
    
    'selected country
    mcCountry = GetVariable("Country")
    If mcCountry = "" Then
        Dim LCID As Long
        LCID = GetSystemDefaultLCID()
        mcCountry = GetUserLocaleInfo(LCID, LOCALE_SISO3166CTRYNAME)
        SetVariable "Country", mcCountry
    End If
       
    'check if combinations are defined
    If TableExist(mdbMain, "Combinations") = False Then
        If TableExist(mdbMain, "CombinationSections") = False Then
            Me.mnuComb.Enabled = False
            Me.mnuFilePrintComb.Enabled = False
        End If
    End If
        
    'check if this is the IPZV build of IceTest
    If VersionSwitch = "ipzv" Or GetVariable("VersionSwitch") = "ipzv" Then
        mcVersionSwitch = "ipzv"
        Me.mnuEditNew.Visible = False
        Me.mnuEditEdit.Visible = False
        Me.mnuFileNew.Visible = False
    End If
    
    'check if this is the ÖIV build of IceTest
    If VersionSwitch = "öiv" Or GetVariable("VersionSwitch") = "öiv" Then
        mcVersionSwitch = "öiv"
    End If
    
    'MM FIPO rules are now valid from April 1 to April 1, a new variable FIPO Version is used for this
    cTemp = GetVariable("FIPO Version")
    If cTemp = "" Then
        cTemp = GetVariable("FIPO Year")
    End If
    If cTemp <> "" Then
        '*** MM: check is competition is running before April 1, if yes don't update FIPO yet
        If GetVariable("Event_First") <> "" Then
            If CDate(GetVariable("Event_First")) >= CDate("1-4-" & Year(Now)) And Val(cTemp) < Year(Now) And Month(Now) >= 4 Then
                MsgBox Translate("Please download a new set of Sport Rules first!", mcLanguage)
            End If
        ElseIf Val(cTemp) < Year(Now) And Month(Now) >= 4 Then
            MsgBox Translate("Please download a new set of Sport Rules first!", mcLanguage)
        End If
    Else
        iKey = MsgBox(Translate("Please download a new set of Sport Rules first!", mcLanguage) & vbCrLf & Translate("Select 'Yes' to download now (requires access to Internet), select 'No' to download later.", mcLanguage), vbYesNo + vbExclamation)
        If iKey = vbYes Then
            ShowDocument "https://www.feif.org/software", Me
            End
        End If
    End If
    
    'read the dir for html files
    ReadIniFile gcIniFile, "Html Files", "Folder", mcHtmlDir
    If mcHtmlDir = "" Or Dir$(mcHtmlDir, vbDirectory) = "" Then
        mcHtmlDir = Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Html\"
        WriteIniFile gcIniFile, "Html Files", "Folder", mcHtmlDir
    End If
    If Dir$(mcHtmlDir, vbDirectory) = "" Then
        If Err > 0 Then
            MsgBox Translate("Cannot access", mcLanguage) & " " & mcHtmlDir, vbCritical
            mcHtmlDir = ""
        Else
            MkDir mcHtmlDir
        End If
    End If
    
    'read the dir for Excel files
    ReadIniFile gcIniFile, "Excel Files", "Folder", mcExcelDir
    If mcExcelDir = "" Or Dir$(mcExcelDir, vbDirectory) = "" Then
        mcExcelDir = Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Excel\"
        WriteIniFile gcIniFile, "Excel Files", "Folder", mcExcelDir
    End If
    If Dir$(mcExcelDir, vbDirectory) = "" Then
        If Err > 0 Then
            MsgBox Translate("Cannot access", mcLanguage) & " " & mcExcelDir, vbCritical
            mcExcelDir = ""
        Else
            MkDir mcExcelDir
        End If
    End If
    
    
    'read the dir for rtf files
    ReadIniFile gcIniFile, "Rtf Files", "Folder", mcRtfDir
    If mcRtfDir = "" Or Dir$(mcRtfDir, vbDirectory) = "" Then
        mcRtfDir = Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Rtf\"
        WriteIniFile gcIniFile, "Rtf Files", "Folder", mcRtfDir
    End If
    If Dir$(mcRtfDir, vbDirectory) = "" Then
        If Err > 0 Then
            MsgBox Translate("Cannot access", mcLanguage) & " " & mcRtfDir, vbCritical
            mcRtfDir = ""
        Else
            MkDir mcRtfDir
        End If
    End If
    
    'is there a special name for the LogDB set?
    ReadIniFile gcIniHorseFile, "Database", "LogDB", cTemp
    If cTemp <> "" Or Dir$(cTemp) = "" Then
        strLogDBName = cTemp
    End If
    
    Dim boolTest As Boolean
    If OpenLogDB = False Then
        miWriteLogDB = 0
    End If
    
    'LL 2007-8-1: change back to using only one HTML directory:
    mcTempHtmlDir = mcHtmlDir
    
    'define the dir for temporary html files...
    'generate a temporary directory name derived from the database name to ensure consistency
    'of HTML files when different databases are used on the same machine
    'cTemp = Replace(Mid(mcDatabaseName, (InStrRev(mcDatabaseName, "\") + 1), (InStrRev(mcDatabaseName, ".") - (InStrRev(mcDatabaseName, "\") + 1))), Chr$(32), "_")
    'mcTempHtmlDir = Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "TempHtml_" & cTemp & "\"
    '...and check if it exists already:
    ' If Dir$(mcTempHtmlDir, vbDirectory) = "" Then
    '    If Err > 0 Then
    '        MsgBox Translate("Cannot access", mcLanguage) & " " & mcTempHtmlDir, vbCritical
    '        mcTempHtmlDir = ""
    '    Else
    '        MkDir mcTempHtmlDir
    '    End If
    'End If
   
    'Assign database to data controls:
    Me.dtaNotYet.DatabaseName = mcDatabaseName
    Me.dtaAlready.DatabaseName = mcDatabaseName
    Me.dtaParticipant.DatabaseName = mcDatabaseName
    Me.dtaTest.DatabaseName = mcDatabaseName
    Me.dtaTestSection.DatabaseName = mcDatabaseName
    Me.dtaTestInfo.DatabaseName = mcDatabaseName
    Me.dtaMarks.DatabaseName = mcDatabaseName
        
    'Make language- or country-specific menu entries visible:
    If mcCountry = "NL" Then
        Me.mnuToolImportNSIJP.Visible = True
    Else
        Me.mnuToolImportNSIJP.Visible = False
    End If
    If mcCountry = "DK" Then
        Me.mnuToolImportDI.Visible = True
    Else
        Me.mnuToolImportDI.Visible = False
    End If
        
    'set help file according to language:
    App.HelpFile = App.Path & "\" & App.EXEName & "_" & mcLanguage & ".Hlp"
    If Dir$(App.HelpFile) = "" And App.Path & "\" & App.EXEName & "_English.Hlp" <> "" Then
        App.HelpFile = App.Path & "\" & App.EXEName & "_English.Hlp"
    End If
    
    'check if database is valid
    If TableExist(mdbMain, "Tests") = False Then
        MsgBox "'" & mcDatabaseName & "' " & Translate("is not a valid database. Select another one or update Sport Rules first!", mcLanguage), vbExclamation
        ChangeDatabase
    End If
            
    cTemp = GetVariable("UseIceSort")
    If cTemp = "" Then
        cTemp = "0"
    End If
    miUseIceSort = Val(cTemp)
    
    'check if SortingApplication is available
    If (Dir$(App.Path & "\" & SortingApplication) > "") And miUseIceSort = 1 Then
        frmMain.cmdComposeGroups.Visible = False
        frmMain.cmbGroupSize.Visible = True
        frmMain.cmbGroupSize.Enabled = False
        frmMain.cmdIceSort.Visible = True
        frmMain.mnuEditStartOrder.Enabled = False
    Else
        miUseIceSort = 0
        frmMain.cmdComposeGroups.Visible = True
        frmMain.cmbGroupSize.Visible = True
        frmMain.cmbGroupSize.Enabled = True
        frmMain.cmdIceSort.Visible = False
        If frmMain.dtaAlready.Recordset.RecordCount = 0 Then
            frmMain.mnuEditStartOrder.Enabled = True
        End If
    End If
    
    'restore old settings
    ReadFormPosition Me, Me.Name
    
    ReadIniFile gcIniFile, Me.Name, "FontSize", cTemp
    msFontSize = MakeStringValue(cTemp)
    If msFontSize < 6 Then
        msFontSize = 0
    ElseIf msFontSize > 14 Then
        msFontSize = 14
    End If
    If msFontSize = 0 Then
        msFontSize = 8.25
        WriteIniFile gcIniFile, Me.Name, "FontSize", Format$(msFontSize)
    End If
    
    ChangeFontSize Me, msFontSize
    
    Form_Resize
    
    DoEvents
    
    StatusBar1.Panels(1).Text = App.EXEName & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") & " "
    
    ReadIniFile gcIniFile, Me.Name, "SelectBySingleClick", cTemp
    If cTemp = "1" Then
      miSelectBySingleClick = True
    Else
      miSelectBySingleClick = False
    End If
    
    ReadIniFile gcIniFile, Me.Name, "BackUpInterval", cTemp
    If cTemp = "" Then cTemp = -1
    miBackupInterval = Val(cTemp)
    
    Me.TestColors = GetColorList
    
    'set minimum tab width for tabstrips
    ReadIniFile gcIniFile, Me.Name, "MinTabWidth", cTemp
    miTabMinWidth = Val(cTemp)
    If miTabMinWidth < 500 Then
        miTabMinWidth = 500
        WriteIniFile gcIniFile, Me.Name, "MinTabWidth", Format$(miTabMinWidth)
    End If
    Me.tbsSelFin.TabMinWidth = miTabMinWidth
    
    'set minimum tab height for tabstrips
    ReadIniFile gcIniFile, Me.Name, "MinTabHeight", cTemp
    miTabMinHeight = Val(cTemp)
    If miTabMinHeight <> 350 Then
        miTabMinHeight = 350
        WriteIniFile gcIniFile, Me.Name, "MinTabHeight", Format$(miTabMinHeight)
    End If
    Me.tbsSelFin.TabFixedHeight = miTabMinHeight
    For iItem = 0 To tbsSection.Count - 1
        Me.tbsSection(iItem).TabFixedHeight = miTabMinHeight
    Next iItem
    
    'Read size of list of participants not yet started
    ReadIniFile gcIniFile, Me.Name, "NotYetHeight", cTemp
    miAlreadyHeight = Val(cTemp)
    
    'Read size of list of participants not yet started
    ReadIniFile gcIniFile, Me.Name, "ParticipantLeft", cTemp
    Me.miParticipantLeft = Val(cTemp)
    
    ' Check if all tests should be shown or only the relevant ones for this event
    ReadIniFile gcIniFile, Me.Name, "ShowAllTests", cTemp
    If cTemp = "" Then cTemp = -1
    frmMain.mnuTestAll.Checked = Val(cTemp)
    
    'build a list of available tests
    CreateTestMenu
       
    'build a list of available tests
    CreateCombinationMenu
   
    ReadIniFile gcIniFile, Me.Name, "ShowColor", cTemp
    Me.chkColor.Value = Val(cTemp)
   
    ReadIniFile gcIniFile, Me.Name, "ShowRein", cTemp
    Me.chkRein.Value = Val(cTemp)
    
    ReadIniFile gcIniFile, Me.Name, "ShowTeam", cTemp
    If cTemp = "" Then cTemp = "1"
    Me.chkTeam.Value = Val(cTemp)
    
    ReadIniFile gcIniFile, Me.Name, "ShowFEIFId", cTemp
    If cTemp = "" Then cTemp = "1"
    Me.chkFeifId.Value = Val(cTemp)
    
    TranslateControls Me
    
    'disable menu items for popup menu
    ChangeTagItem frmMain.mnuFileSep1, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileSep2, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileSep3, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileExit, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileNew, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileChange, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileEven, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileEvenCode, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileEvenName, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileEvenTest, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileResults, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileResultsRtf, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileResultsHtml, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileResultsExcel, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintOverview, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintOverviewFinals, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintOverviewMarks, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintOverviewNoMarks, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintAll, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintAllPrinter, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintAllMerge, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintRider, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrint, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintAllPrinter, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileSaveAs, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintForms, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileBackup, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileFEIFWR, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintComb, "PopUp", "No"
    ChangeTagItem frmMain.mnuComb, "PopUp", "No"
    ChangeTagItem frmMain.mnuCombEdit, "PopUp", "No"
    ChangeTagItem frmMain.mnuCombAddNew, "PopUp", "No"
    ChangeTagItem frmMain.mnuCombRemove, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestEdit, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestAll, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestAddNew, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestRemove, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestSep1, "PopUp", "No"
    ChangeTagItem frmMain.mnutestQual1, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual2, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual3, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual4, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual5, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual6, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual7, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual8, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual9, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual10, "PopUp", "No"
    ChangeTagItem frmMain.mnutestQual11, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual12, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual13, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual14, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual15, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual16, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual17, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual18, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual19, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual20, "PopUp", "No"
    ChangeTagItem frmMain.mnutestQual21, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual22, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual23, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual24, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual25, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual26, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual27, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual28, "PopUp", "No"
    ChangeTagItem frmMain.mnuTestQual29, "PopUp", "No"
    ChangeTagItem frmMain.mnuFileSaveAs, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintFormEntrance, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintFormEquipment, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintFormEquipmentComplete, "PopUp", "No"
    ChangeTagItem frmMain.mnuFilePrintFormEquipmentOnly, "PopUp", "No"
        
    'is this a FEIF WorldRanking (WR) Event?
    cTemp = GetVariable("WR_code")
    If cTemp <> "" Then
         mnuFileFEIFWR.Enabled = True
    End If
        
    ' are there any queries to show in special menu?
    If mdbMain.QueryDefs.Count > 0 Then
        Dim Qdf As DAO.QueryDef
        iItem = 0
        For Each Qdf In mdbMain.QueryDefs
            If Left$(Qdf.Name, 5) = "__" & mcCountry & "_" Then
                frmMain.mnuFileSaveAs.Visible = True
                iItem = iItem + 1
                If iItem > frmMain.mnuFileSaveAsItem.Count Then
                    Load frmMain.mnuFileSaveAsItem(frmMain.mnuFileSaveAsItem.Count)
                End If
                frmMain.mnuFileSaveAsItem(frmMain.mnuFileSaveAsItem.Count - 1).Caption = StrConv(Translate(Mid$(Qdf.Name, 6), mcLanguage), vbProperCase)
                ChangeTagItem frmMain.mnuFileSaveAsItem(frmMain.mnuFileSaveAsItem.Count - 1), "Sql", Qdf.Name
                ChangeTagItem frmMain.mnuFileSaveAsItem(frmMain.mnuFileSaveAsItem.Count - 1), "PopUp", "No"
            ElseIf Left$(Qdf.Name, 1) = "_" And Left$(Qdf.Name, 2) <> "__" Then
                frmMain.mnuFileSaveAs.Visible = True
                iItem = iItem + 1
                If iItem > frmMain.mnuFileSaveAsItem.Count Then
                    Load frmMain.mnuFileSaveAsItem(frmMain.mnuFileSaveAsItem.Count)
                End If
                frmMain.mnuFileSaveAsItem(frmMain.mnuFileSaveAsItem.Count - 1).Caption = StrConv(Translate(Mid$(Qdf.Name, 2), mcLanguage), vbProperCase)
                ChangeTagItem frmMain.mnuFileSaveAsItem(frmMain.mnuFileSaveAsItem.Count - 1), "Sql", Qdf.Name
                ChangeTagItem frmMain.mnuFileSaveAsItem(frmMain.mnuFileSaveAsItem.Count - 1), "PopUp", "No"
            End If
        Next
    End If
    
    'are there any add on's to include?
    If Dir$(App.Path & "\iceform.exe") <> "" Or Dir$(Replace(App.Path, "IceTest", "IceHorseTools") & "\iceform.exe") <> "" Or Dir$(Replace(App.Path, "IceHorseTools", "IceTest") & "\iceform.exe") <> "" Then
        Me.mnuFilePrintFormJudgesLand.Visible = True
    Else
        Me.mnuFilePrintFormJudgesLand.Visible = False
    End If
    
    'LL: Is the check for icemarks.exe still needed?
    If Dir$(App.Path & "\icemarks.exe") <> "" Or Dir$(Replace(App.Path, "IceTest", "IceHorseTools") & "\icemarks.exe") <> "" Or Dir$(Replace(App.Path, "IceHorseTools", "IceTest") & "\icemarks.exe") <> "" Then
        Me.cmdMarks.Visible = True
    Else
        Me.cmdMarks.Visible = False
    End If
    
     'LL: Check if FEIFconnector is present and running
     If Dir$(App.Path & "\feifconnector.exe") <> "" Then
         'ok, FEIFconnector is installed, but is it also running?
         If WindowIsOpen("FEIFconnector") = False Then
             Dim RetVal
             RetVal = Shell(App.Path & "\feifconnector.exe", vbMinimizedNoFocus)
         End If
     End If
    
    cTemp = GetVariable("ShowHorseId")
    miShowHorseId = Val(cTemp)

    
    cTemp = GetVariable("ShowRidersClub")
    miShowRidersClub = Val(cTemp)
    
    cTemp = GetVariable("ShowRidersTeam")
    miShowRidersTeam = Val(cTemp)
    
    cTemp = GetVariable("ShowJudgesRanking")
    If cTemp = "" Then cTemp = "1"
    miShowJudgesRanking = Val(cTemp)
    
    'LL 2010-3-8: new variable to know if database log file is desired
    cTemp = GetVariable("WriteLogDB")
    If cTemp = "" Then cTemp = "0"
    miWriteLogDB = Val(cTemp)
        
    'LL 2010-4-4: new variable to know if LKs should be printed (IPZV only)
    cTemp = GetVariable("ShowRidersLK")
    If cTemp = "" Then cTemp = "1"
    miShowRidersLK = Val(cTemp)
    
    'LL 2007-8-1: changed default for Excel file creation to YES:
    cTemp = GetVariable("CreateExcelFiles")
    If cTemp = "" Then
        cTemp = "1"
        SetVariable "CreateExcelFiles", 1
    End If
    miExcelFiles = Val(cTemp)
    
    'LL 2007-8-1: changed default for HTML file creation to YES:
    cTemp = GetVariable("CreateHtmlFiles")
    If cTemp = "" Then
        cTemp = "1"
        SetVariable "CreateHtmlFiles", 1
    End If
    miHtmlFiles = Val(cTemp)
    
    cTemp = GetVariable("UseHighLights")
    If cTemp = "" Then cTemp = "0"
    miUseHighLights = Val(cTemp)
    
    cTemp = GetVariable("FinalsSequence")
    If cTemp = "" Then cTemp = "0"
    miFinalsSequence = Val(cTemp)
    
    cTemp = GetVariable("MarkFinals")
    If cTemp = "" Then cTemp = "1"
    miMarkFinalsInResultLists = Val(cTemp)
    
    cTemp = GetVariable("BFinalLevel")
    If cTemp = "" Then cTemp = "20"
    miBFinalLevel = Val(cTemp)
    
    cTemp = GetVariable("CFinalLevel")
    If cTemp = "" Then cTemp = "30"
    miCFinalLevel = Val(cTemp)
    
    cTemp = GetVariable("UseColors")
    If cTemp = "" Then cTemp = "1"
    miUseColors = Val(cTemp)
    If miUseColors = 1 Then
        Me.chkColor.Caption = TranslateCaption("&Groups / Colors", 0, False)
    Else
        Me.chkColor.Caption = TranslateCaption("&Groups", 0, False)
    End If
    
    cTemp = GetVariable("NoColor")
    If cTemp = "" Then
        cTemp = "xx"
        SetVariable "NoColor", cTemp
    End If
    mcNoColor = cTemp
    
    cTemp = GetVariable("ShowHorseAge")
    If cTemp = "" Then
        cTemp = "1"
        SetVariable "ShowHorseAge", cTemp
    End If
    miShowHorseAge = Val(cTemp)
   
    'what is the horseage to be checked
    cTemp = GetVariable("HorseAgeLimit")
    If cTemp = "" Then
        cTemp = "7"
        SetVariable "HorseAgeLimit", "7"
    End If
    miHorseAgeLimit = Val(cTemp)

    cTemp = GetVariable("ExcelSeparator")
    If cTemp = "" Then
        cTemp = ";"
        SetVariable "ExcelSeparator", cTemp
    End If
    mcExcelSeparator = cTemp
    
    'fill combos
    With Me.cmbGroupSize
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
    End With
    
    For iTemp = 0 To 3
        With Me.cmbNumJudges(iTemp)
            .Clear
            .AddItem "1"
            .AddItem "2"
            .AddItem "3"
            .AddItem "5"
        End With
    Next
    
    
    'read test data
    Set rstTest = mdbMain.OpenRecordset("SELECT Code FROM Tests WHERE Code='" & Me.TestCode & "' AND NOT Removed=True")
    If rstTest.RecordCount = 0 Then
        Me.TestCode = ""
    End If
    If Me.TestCode = "" Then
        Set rstTest = mdbMain.OpenRecordset("SELECT Code FROM Tests WHERE NOT Removed=True ORDER BY Code ASC")
        If rstTest.RecordCount > 0 Then
            Me.TestCode = rstTest.Fields(0)
        Else
            MsgBox Translate("No valid list of tests available. Download Sport Rules first.", mcLanguage)
        End If
    End If
    rstTest.Close
    Set rstTest = Nothing
   
    DoEvents
   
    'generate participants' detail data for HTML
    CreateParticipantsFile
   
    Me.Show
    fSplash.Hide
    
    LookUpTest

    On Local Error GoTo 0
         
    SetMouseNormal
   
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If

End Sub

Private Sub Form_Resize()
    Dim iItem As Integer
    Dim iTemp As Integer
    Static iBusy As Integer
    Dim iSwap As Integer
    
    On Local Error Resume Next
    
    If iBusy = False Then
        iBusy = True
        'iSwap = True
        
        Me.tbsSelFin.Width = ScaleWidth
        Me.tbsSelFin.Height = ScaleHeight - Me.StatusBar1.Height
        Me.tbsSelFin.TabMinWidth = miTabMinWidth
        Me.tbsSelFin.TabFixedWidth = (Me.tbsSelFin.Width - 300) \ (Me.tbsSelFin.Tabs.Count + 1)
        Me.tbsSelFin.TabFixedHeight = miTabMinHeight
        
        For iItem = 0 To 3
            Me.tbsSection(iItem).Top = Me.tbsSelFin.ClientTop + 100
            Me.tbsSection(iItem).Left = Me.tbsSelFin.ClientLeft
            Me.tbsSection(iItem).Width = Me.tbsSelFin.ClientWidth
            Me.tbsSection(iItem).Height = Me.tbsSelFin.ClientHeight - 100
            Me.tbsSection(iItem).TabMinWidth = miTabMinWidth
            Me.tbsSection(iItem).TabFixedHeight = miTabMinHeight
            Me.tbsSection(iItem).TabFixedWidth = (Me.tbsSection(iItem).Width - 100) \ Max(Me.tbsSection(iItem).Tabs.Count + 1, 2)
        Next iItem
                
        With Me.fraLists
             .Top = Me.tbsSection(0).ClientTop + 100
             .Left = Me.tbsSection(0).ClientLeft
             .Height = Me.tbsSection(0).ClientHeight - 200
             If Me.miParticipantLeft = 0 Then
                 .Width = Me.tbsSection(0).ClientWidth / 2 - 100
             Else
                .Width = miParticipantLeft
             End If
        End With
        
        With Me.fraDivider2
             .Top = Me.tbsSection(0).ClientTop + 100
             .Left = fraLists.Left + fraLists.Width
             .Width = 150
             .Height = Me.tbsSection(0).ClientHeight - 200
        End With
        
        With Me.fraOther
             .Top = Me.tbsSection(0).ClientTop + 100
             .Left = fraDivider2.Left + fraDivider2.Width
             .Width = Me.tbsSection(0).ClientWidth - fraDivider2.Width - fraLists.Width
             .Height = Me.tbsSection(0).ClientHeight - 200
        End With
        
        With Me.fraShow
            .Container = Me.fraLists
            .Top = .Container.Height - .Height
            .Left = 0
            .Width = .Container.Width
        End With
            
        With Me.fraPrevious
             .Container = fraLists
             .Top = 0
             .Left = 0
             .Width = .Container.Width
             .Height = Me.Font.Size * 97
        End With
        
        With txtPrevious
            .Container = fraPrevious
            .Top = 250
            .Left = 50
            .Width = fraPrevious.Width - 150
            .Height = fraPrevious.Height - 300
        End With
        
        With Me.fraAlready
             .Container = fraLists
             .Left = 0
             .Width = .Container.Width
             If fraPrevious.Visible = True Then
                .Top = Me.fraPrevious.Top + Me.fraPrevious.Height + 150
                If miAlreadyHeight = 0 Then
                   .Height = (.Container.Height - fraShow.Height - fraPrevious.Height - 100) / 2 - 200
                Else
                   .Height = miAlreadyHeight - fraPrevious.Height - 100
                End If
             Else
                .Top = 0
                If miAlreadyHeight = 0 Then
                   .Height = (.Container.Height - fraShow.Height) / 2 - 200
                Else
                   .Height = miAlreadyHeight
                End If
             End If
        End With
        
        With Me.fraDivider1
            .Container = fraLists
            .Left = 0
            .Top = fraAlready.Top + fraAlready.Height
            .Height = 300
            .Width = .Container.Width
        End With
        
        With Me.fraNotYet
            .Container = fraLists
            .Left = 0
            .Width = .Container.Width
            .Top = fraDivider1.Top + fraDivider1.Height
            .Height = .Container.Height - fraShow.Height - .Top
        End With
                
        With Me.chkColor
          .Container = fraShow
          .Width = fraShow.Width \ 4 - 50
          .Left = 150
        End With
        
        With Me.chkRein
          .Top = chkColor.Top
          .Container = fraShow
          .Width = fraShow.Width \ 4 - 50
          .Left = (fraShow.Width - 100) \ 4
        End With
        
        With Me.chkTeam
          .Top = chkColor.Top
          .Container = fraShow
          .Width = fraShow.Width \ 4 - 50
          .Left = (fraShow.Width - 100) * 2 \ 4
        End With
        
        With Me.chkFeifId
          .Top = chkColor.Top
          .Container = fraShow
          .Width = fraShow.Width \ 4 - 50
          .Left = (fraShow.Width - 100) * 3 \ 4
        End With
        
        With Me.fraCurrent
            .Container = fraOther
            .Left = 0
            .Top = 0
            .Width = .Container.Width
        End With
        
        With fraParticipant
            .Container = fraCurrent
            .Top = 250
            .Left = 100
            .Width = Me.fraCurrent.Width - 200
            .Height = txtParticipant.Top + 2.5 * txtParticipant.Height + 100
            fraCurrent.Height = .Height + 350
        End With
        
        With fraMarks
            .Container = fraCurrent
            .Top = fraParticipant.Top + fraParticipant.Height + 50
            .Left = 100
            .Width = Me.fraCurrent.Width - 200
            .Height = (txtMarks(0).Height + 50) * 5 + 600
            fraCurrent.Height = fraCurrent.Height + .Height + 50
        End With
        
        With fraTime
            .Container = fraCurrent
            .Top = fraMarks.Top
            .Left = fraMarks.Left
            .Width = fraMarks.Width
            .Height = fraMarks.Height
        End With
        
        With fraResults
             .Container = fraCurrent
             .Top = fraMarks.Top + fraMarks.Height + 50
             .Left = 100
             .Width = Me.fraCurrent.Width - 200
             .Height = txtScore.Top + txtScore.Height + cmdOK.Height + 600
            fraCurrent.Height = fraCurrent.Height + .Height + 50
        End With
        
        With fraGroups
             .Container = fraOther
             .Left = 0
             .Top = fraCurrent.Top + fraCurrent.Height + 50
             .Width = .Container.Width / 2 - 50
             .Height = Me.chkSplitResultLists.Top + Me.chkSplitResultLists.Height + 100
        End With
        
        With fraJudges
             .Container = fraOther
             .Left = .Container.Width / 2 + 50
            .Top = fraCurrent.Top + fraCurrent.Height + 50
             .Width = .Container.Width / 2 - 50
             .Height = Me.cmbNumJudges(0).Top + Me.cmbNumJudges(0).Height + 100
        End With
        
        With fraFinals
             .Container = fraOther
             .Left = fraGroups.Left
             .Width = fraGroups.Width
             .Top = fraGroups.Top
             .Height = Me.chkSplitFinals.Top + Me.chkSplitFinals.Height + 100
        End With
        
        With Me.cmbGroupSize
            .Container = Me.fraGroups
            .Left = 100
            .Top = 350
        End With
        
        With Me.chkSplitResultLists
            .Container = Me.fraGroups
            .Left = 100
            .Top = Me.cmbGroupSize.Top + Me.cmbGroupSize.Height + 100
            If Me.dtaTestInfo.Recordset.Fields("Handling") > 4 Then
              .Enabled = False
            Else
              .Enabled = True
            End If
        End With
        
        With Me.cmdIceSort
            Set .Container = Me.fraGroups
            .Left = .Container.Width - .Width - 100
            .Top = 250
            .Height = .Container.Height - .Top - 100
        End With
        
        With Me.cmdComposeGroups
            .Container = Me.fraGroups
            .Left = .Container.Width - .Width - 100
            .Top = 250
            .Height = cmbGroupSize.Height
        End With
        
        For iTemp = 0 To 3
            With Me.cmbNumJudges(iTemp)
                .Container = Me.fraJudges
                .Left = 100
                .Top = 350
            End With
        Next
        
        With Me.lblNumJudges
            .Container = Me.fraJudges
            .Left = cmbNumJudges(0).Left + cmbNumJudges(0).Width + 50
            .Top = cmbNumJudges(0).Top
        End With
        
        With Me.cmdTestInfo
            .Container = Me.fraJudges
            .Top = 250
            .Height = .Container.Height - .Top - 100
            .Width = .Height
            .Left = .Container.Width - .Width - 100
        End With
        
        With Me.cmdMarks
            .Container = Me.fraJudges
            .Top = 250
            .Height = .Container.Height - .Top - 100
            .Width = .Height
            .Left = cmdTestInfo.Left - .Width - 50
        End With
        
        With Me.lblParticipant
             .Container = fraParticipant
             .Top = txtParticipant.Top
             .Left = txtParticipant.Left + txtParticipant.Width + 100
             .Height = fraParticipant.Height - .Top - 100
             .Width = fraParticipant.Width - .Left - 100
        End With
        
        With cmdInfo
             .Container = fraParticipant
             .Top = txtParticipant.Top
             .Height = txtParticipant.Height
             .Width = txtParticipant.Height
             .Left = fraParticipant.Width - cmdInfo.Width - 100
        End With
       
        With Me.dblstNotYet
          .Container = fraNotYet
          .Top = 350
          .Left = 50
          .Width = fraNotYet.Width - 150
          .Height = fraNotYet.Height - 400
        End With
        
        With txtMove
          .Container = fraNotYet
          .Left = 50
          .Width = fraNotYet.Width - 150
          .Font = dblstNotYet.Font
        End With
        
        With Me.lstAlready
          .Container = fraAlready
          .Top = 350
          .Left = 50
          .Width = fraAlready.Width - 150
          .Height = fraAlready.Height - 400
        End With
        
        With Me.txtAlready
          .Container = fraAlready
          .Top = 350
          .Left = 50
          .Width = fraAlready.Width - 150
          .Height = fraAlready.Height - 400
        End With
        
        For iItem = 0 To 4
           With lblMarks(iItem)
                .Container = fraMarks
                .Left = 50
                If iItem = 0 Then
                   .Top = 350
                Else
                   .Top = txtMarks(iItem - 1).Top + txtMarks(iItem - 1).Height + 100
                End If
                .Width = fraMarks.Width - txtMarks(0).Width - 200
           End With
           
           With txtMarks(iItem)
                .Container = fraMarks
                .Left = lblMarks(iItem).Left + lblMarks(iItem).Width + 50
                .Top = lblMarks(iItem).Top
           End With
        Next iItem
        
        With lblTime
            .Container = fraTime
            .Left = 50
            .Top = 350
            .Width = fraMarks.Width - txtTime.Width - 200
        End With
        
        With txtTime
            .Container = fraTime
            .Top = lblTime.Top
            .Left = lblTime.Left + lblTime.Width + 50
        End With
        
        With txtScore
          .Container = Me.fraResults
          .Top = 350
          .Left = Me.fraResults.Width - txtScore.Width - 100
        End With
        
        With lblScore
          .Container = Me.fraResults
          .Top = txtScore.Top
          .Width = Me.fraResults.Width / 2 - txtScore.Width - 100
          .Left = Me.txtScore.Left - lblScore.Width - 50
        End With
        
        With Me.chkDisqualified
          .Container = Me.fraResults
          .Top = txtScore.Top
          .Width = Me.fraResults.Width / 2 - 100
        End With
        
        With Me.chkWithdrawn
          .Container = Me.fraResults
          .Top = chkDisqualified.Top + chkDisqualified.Height + 50
          .Left = chkDisqualified.Left
          .Width = Me.fraResults.Width / 2 - 100
        End With
        
        With Me.chkFlag
          .Container = Me.fraResults
          .Top = chkWithdrawn.Top + chkWithdrawn.Height + 50
          .Left = chkDisqualified.Left
          .Width = Me.fraResults.Width / 2 - 100
        End With
                
        With Me.chkNoStart
          .Container = Me.fraResults
          .Top = chkFlag.Top + chkFlag.Height + 50
          .Left = chkDisqualified.Left
          .Width = Me.fraResults.Width / 2 - 100
        End With
        
        With cmdOK
          .Container = Me.fraResults
          .Top = Me.fraResults.Height - .Height - 100
          .Left = Me.fraResults.Width - .Width - 100
        End With

        With Me.cmdComposeFinals
            .Container = Me.fraFinals
            .Left = .Container.Width - .Width - 100
            .Top = 250
            .Height = Me.chkSplitFinals.Height
        End With
        
        With Me.chkSplitFinals
          .Container = fraFinals
          .Width = fraFinals.Width - Me.cmdComposeFinals.Width - 250
          .Left = 150
          If Me.dtaTestInfo.Recordset.Fields("Handling") > 4 Then
            .Enabled = False
          Else
            .Enabled = True
          End If
        End With
        
             
        ShowProgressbar Me, 2, 0

        iBusy = False
        
    End If
    
    On Local Error GoTo 0
    
End Sub
    
Private Sub Form_Unload(Cancel As Integer)
    Dim iKey As Integer
    
    If miBackupInterval >= 0 Then
        iKey = MsgBox(Translate("Create backup first?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton1)
        If iKey = vbYes Then
            CreateBackup mdbMain
        End If
    End If
    
    Unload fSplash
    
    WriteIniFile gcIniFile, Me.Name, "LastTest", Me.TestCode
    WriteIniFile gcIniFile, Me.Name, "ShowColor", Me.chkColor.Value
    WriteIniFile gcIniFile, Me.Name, "ShowRein", Me.chkRein.Value
    WriteIniFile gcIniFile, Me.Name, "ShowTeam", Me.chkTeam.Value
    WriteIniFile gcIniFile, Me.Name, "ShowFEIFId", Me.chkFeifId.Value
    WriteIniFile gcIniFile, Me.Name, "NotYetHeight", Format$(miAlreadyHeight)
    WriteIniFile gcIniFile, Me.Name, "ParticipantLeft", Format$(Me.miParticipantLeft)
    WriteIniFile gcIniFile, Me.Name, "FontSize", Format$(msFontSize)
    WriteIniFile gcIniFile, Me.Name, "SelectBySingleClick", IIf(miSelectBySingleClick, "1", "0")
    WriteIniFile gcIniFile, Me.Name, "BackupInterval", Format$(miBackupInterval)
    
    WriteFormPosition Me, Me.Name
        
End Sub

Private Sub fraAlready_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub fraDivider2_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then
        Me.fraDivider2.Left = Me.fraDivider2.Left + X
    End If
End Sub

Private Sub fraDivider2_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Me.fraDivider2.Left <> Me.miParticipantLeft Then
        miParticipantLeft = Me.fraDivider2.Left
        If miParticipantLeft < ScaleWidth \ 5 Then
            miParticipantLeft = ScaleWidth \ 5
        ElseIf miParticipantLeft > 0.8 * ScaleWidth Then
            miParticipantLeft = 0.8 * ScaleWidth
        End If
        Form_Resize
    End If
End Sub

Private Sub fraGroups_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub fraJudges_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub fraLists_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub fraOther_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub fraParticipant_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub


Private Sub fraCurrent_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub
Private Sub fraFinals_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub fraMarks_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub fraNotYet_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub fraResults_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub
Private Sub fraDivider1_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbLeftButton Then
        Me.fraDivider1.Top = Me.fraDivider1.Top + y
    End If
End Sub
Private Sub fraDivider1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Me.fraDivider1.Top <> Me.miAlreadyHeight Then
        miAlreadyHeight = Me.fraDivider1.Top
        If miAlreadyHeight < 1 * (Me.fraLists.Height - Me.fraShow.Height) \ 10 Then
            miAlreadyHeight = 1 * (Me.fraLists.Height - Me.fraShow.Height) \ 10
        ElseIf miAlreadyHeight > 9 * (Me.fraLists.Height - Me.fraShow.Height) \ 10 Then
            miAlreadyHeight = 9 * (Me.fraLists.Height - Me.fraShow.Height) \ 10
        End If
        Form_Resize
    End If
End Sub

Private Sub fraShow_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub fraTime_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub lblMarks_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub lblParticipant_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub lblScore_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub lblTime_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub lstAlready_Click()
   Dim iTemp As Integer
   Dim cTemp As Integer
   
   If Left$(lstAlready.Text, 1) = vbTab Then
      lstAlready.ListIndex = lstAlready.ListIndex - 1
   End If
   
   If miSelectBySingleClick = True And miChangeCaption = False Then
      If lstAlready.SelCount > 0 Then
         cTemp = Format$(lstAlready.ItemData(lstAlready.ListIndex), "000")
      End If
      StoreCurrentMarks
   End If
   
   iTemp = lstAlready.ListIndex
   dblstNotYet.BoundText = ""
   lstAlready.ListIndex = iTemp
   
   If miSelectBySingleClick = True And miChangeCaption = False Then
      If lstAlready.SelCount > 0 Then
         txtParticipant.Text = cTemp
      End If
      LookUpParticipant
   End If
End Sub
Private Sub lstAlready_DblClick()
   Dim iTemp As Integer
   Dim cTemp As String
   
   If lstAlready <> "" Then
    If Left$(lstAlready.Text, 1) = vbTab Then
       lstAlready.ListIndex = lstAlready.ListIndex - 1
    End If
    
    cTemp = Format$(lstAlready.ItemData(lstAlready.ListIndex), "000")
    StoreCurrentMarks
    
    iTemp = lstAlready.ListIndex
    dblstNotYet.BoundText = ""
    lstAlready.ListIndex = iTemp
    
    txtParticipant.Text = cTemp
    LookUpParticipant
   End If
End Sub
Private Sub lstAlready_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
      ChangeCaption True
   ElseIf KeyCode = vbKeyReturn Then
      lstAlready_DblClick
      KeyCode = 0
   ElseIf KeyCode = vbKeyDown Then
      If lstAlready.ListIndex < lstAlready.ListCount - 2 Then
         lstAlready.ListIndex = lstAlready.ListIndex + 2
      Else
         lstAlready.ListIndex = lstAlready.ListIndex + 1
      End If
      KeyCode = 0
   End If
End Sub
Private Sub lstAlready_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If

End Sub


Private Sub mnuCombAddNew_Click()
    CreateNewCombination
    CreateCombinationMenu
End Sub


Private Sub mnuEditMove_Click()
    Dim iKey As Integer
    Dim cMsg As String
    Dim cTemp As String
    Dim iNewPosition As Integer
    Dim iPosition As Integer
    Dim iColor As Integer
    Dim cColor() As String
    Dim rstEntry As DAO.Recordset
    Dim iNewGroup As Integer
    
    iPosition = 0
    If dblstNotYet.BoundText <> "" Then
        iNewPosition = -1
        cTemp = InputBox(UCase$(Translate("Move", mcLanguage)) & " " & Translate("Startnumber", mcLanguage) & " " & Left$(dblstNotYet.BoundText, 3) & " " & UCase$(Translate("behind", mcLanguage)) & " " & Translate("which other rider", mcLanguage) & " (" & Translate("Startnumber", mcLanguage) & "; " & "0 = " & Translate("top of the list", mcLanguage) & ")?")
        If Val(cTemp) > 0 Then
            Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Sta='" & Format$(Val(cTemp), "000") & "'")
            If rstEntry.RecordCount > 0 Then
                iNewPosition = rstEntry.Fields("Position") + 1
                iNewGroup = rstEntry.Fields("Group") + 0
            End If
            rstEntry.Close
        ElseIf cTemp <> "" Then
            Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " ORDER BY Group,Position")
            If rstEntry.RecordCount > 0 Then
                iNewPosition = rstEntry.Fields("Position") - 1
                iNewGroup = rstEntry.Fields("Group") + 0
            End If
            rstEntry.Close
        End If
        If iNewPosition <> -1 Then
            Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Sta='" & Left$(dblstNotYet.BoundText, 3) & "'")
            If rstEntry.RecordCount > 0 Then
                With rstEntry
                    .Edit
                    .Fields("Position") = iNewPosition
                    .Fields("Group") = iNewGroup
                    .Update
                End With
            End If
            Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND NOT Sta IN (SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ") ORDER BY Position")
            If rstEntry.RecordCount > 0 Then
                With rstEntry
                    Do While Not .EOF
                        .Edit
                        .Fields("Position") = (.AbsolutePosition + 1) * 2
                        .Update
                        .MoveNext
                    Loop
                End With
            End If
            If TestStatus > 0 Then
                AddColorsToFinals
            Else
                Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND NOT Sta IN (SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ") ORDER BY Position")
                If rstEntry.RecordCount > 0 Then
                    iColor = 0
                    cColor = Split(TestColors, ",")
                    With rstEntry
                        Do While Not .EOF
                            .Edit
                            If iColor <= UBound(cColor) Then
                                .Fields("Color") = cColor(iColor)
                            Else
                                .Fields("Color") = mcNoColor
                            End If
                            .Update
                            .MoveNext
                            iColor = iColor + 1
                        Loop
                    End With
                End If
            End If
            rstEntry.Close
        End If
        Set rstEntry = Nothing
        dtaNotYet.Refresh
    Else
        MsgBox Translate("Select a participant first.", mcLanguage), vbExclamation
    End If
    txtMove.Text = ""
    dblstNotYet.Tag = ""
    txtMove.Visible = False
End Sub

Private Sub mnuEditStartOrder_Click()
    Dim iKey As Integer
    Dim lNewPosition As Long
    Dim rstEntry As DAO.Recordset
    
    iKey = MsgBox(Translate("Do you want to recompose the starting order (random)?", mcLanguage), vbQuestion + vbYesNo)
    If iKey = vbYes Then
        Randomize Timer
        Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND NOT Sta IN (SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ") ORDER BY Rnd(Position);")
        If rstEntry.RecordCount > 0 Then
            With rstEntry
                Do While Not .EOF
                    .Edit
                    lNewPosition = Rnd(Timer) * 10000
                    .Fields("Position") = lNewPosition
                    .Fields("Group") = 0
                    .Update
                    .MoveNext
                Loop
            End With
        End If
        rstEntry.Close
        DoEvents
        Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND NOT Sta IN (SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ") ORDER BY Rnd(Position);")
        If rstEntry.RecordCount > 0 Then
            With rstEntry
                Do While Not .EOF
                    .Edit
                    lNewPosition = (.AbsolutePosition + 1) * 2
                    .Fields("Position") = lNewPosition
                    .Update
                    .MoveNext
                Loop
            End With
        End If
        rstEntry.Close
        DoEvents
        Set rstEntry = Nothing
        If Val(frmMain.cmbGroupSize) > 1 And frmMain.chkColor.Value <> 0 Then
            dtaNotYet.Refresh
            cmdComposeGroups_Click
            MsgBox Translate("Starting order has been recomposed (random).", mcLanguage), vbInformation
        Else
            MsgBox Translate("Starting order has been recomposed (random). Recompose groups when needed.", mcLanguage), vbInformation
        End If
        dtaNotYet.Refresh
    End If
End Sub

Private Sub mnuEditTieBreak_Click()
    CheckTieBreak True
End Sub

Private Sub mnuEditWarning_Click()
    
    'the database check can be removed once the changed database layout of table penalties
    'is present in all databases (2008-1-1 at the latest):
    If CheckDatabase(mcDatabaseName) = True Then
        frmWarnings.Show 1
    Else
        NotYet
    End If
End Sub

Private Sub mnuFileEvenJudg_Click()
    NotYet
End Sub

Private Sub mnuFileEvenTest_Click()
    With frmPick
        .QryAll = "SELECT Tests.Code & ' - ' & Test AS cList FROM Tests INNER JOIN TestInfo ON Tests.Code=TestInfo.Code WHERE (Removed=False or ISNULL(Removed)) AND (TestInfo.Nr<1 AND NOT Tests.Code IN (SELECT DISTINCT Code FROM Results)) ORDER BY Tests.Code"
        .QryPicked = "SELECT Tests.Code & ' - ' & Test AS cList FROM Tests INNER JOIN TestInfo ON Tests.Code=TestInfo.Code WHERE (Removed=False or ISNULL(Removed)) AND (TestInfo.Nr>0 OR Tests.Code IN (SELECT DISTINCT Code FROM Results)) ORDER BY TestInfo.Nr"
        .TableName = "TestInfo"
        .FieldKey = "Code"
        .FieldSeq = "Nr"
        .Caption = Translate("Tests for this event", mcLanguage)
        .cmdAdd.Visible = False
        .Show 1, Me
    End With
    CreateTestMenu
End Sub

Private Sub mnuFilePrintForms_Click()
    PrintRtfForms
End Sub

Private Sub mnuFilePrintAllMerge_Click()
    PrintAllMerge
End Sub

Private Sub mnuFilePrintAllPrinter_Click()
    PrintAllPrinter
End Sub

Private Sub mnufileprintCombComb_Click(Index As Integer)
    If ReadTagItem(mnuFilePrintCombComb(Index), "Comb") = "Team" Then
        CalculateTeamCombination ReadTagItem(mnuFilePrintCombComb(Index), "Comb")
    ElseIf ReadTagItem(mnuFilePrintCombComb(Index), "Comb") = "Club" Then
        CalculateTeamCombination ReadTagItem(mnuFilePrintCombComb(Index), "Comb")
    Else
        CalculateCombination ReadTagItem(mnuFilePrintCombComb(Index), "Comb")
    End If
End Sub

Private Sub mnuCombEdit_Click()
    With frmToolBox
        .strQry = "SELECT DISTINCT Combination FROM Combinations WHERE USERLEVEL>=0 ORDER BY Combination"
        .Caption = ClipAmp(mnuCombEdit.Caption)
        .Show 1, Me
    End With
    If Me.Tempvar <> "" Then
        With frmCombination
            .fcCombination = Tempvar
            Tempvar = ""
            .Show 1, Me
        End With
        CreateCombinationMenu
    End If
End Sub

Private Sub mnuCombRemove_Click()
    RemoveCombination
    CreateCombinationMenu
End Sub

Private Sub mnuEditAdd_Click()
   Dim cTemp As String
   Dim iTemp As Integer
   Dim cSta() As String
   Dim iPosition As Integer
   Dim rstEntry As Recordset
   Dim cQry As String
   
   cTemp = InputBox$(Translate("Search for", mcLanguage), Translate("Add a new participant to this test.", mcLanguage), mnuEditAdd.Tag)
   
   cQry = "SELECT Participants.Sta "
   cQry = cQry & " & '  -  ' & Persons.Name_First"
   cQry = cQry & " & ' ' & Persons.Name_Last"
   cQry = cQry & " & ' - ' & Horses.Name_Horse as cList"
   cQry = cQry & " FROM (Participants"
   cQry = cQry & " INNER JOIN Persons"
   cQry = cQry & " ON Participants.PersonId=Persons.PersonId)"
   cQry = cQry & " INNER JOIN Horses"
   cQry = cQry & " ON Participants.HorseId=Horses.HorseId"
   cQry = cQry & " WHERE Participants.Sta & ' ' & Persons.Name_First & ' ' & Persons.Name_Last & ' - ' & Horses.Name_Horse LIKE " & Chr$(34) & "*" & cTemp & "*" & Chr$(34)
   cQry = cQry & " ORDER BY Participants.Sta"
   
   With frmToolBox
        .intChecked = True
        .strQry = cQry
        .Caption = Translate("Searching", mcLanguage) & " '" & cTemp & "' "
        .Show 1, Me
    End With
   
   If Me.Tempvar <> "" Then
      cSta = Split(Me.Tempvar, "|")
      For iTemp = 0 To UBound(cSta)
        If cSta(iTemp) <> "" Then
            Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code LIKE '" & TestCode & "' AND STATUS=" & Me.TestStatus & " AND Sta LIKE '" & Left$(cSta(iTemp), 3) & "'")
            If rstEntry.RecordCount = 0 Then
                  Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code LIKE '" & TestCode & "' AND STATUS=" & Me.TestStatus & " ORDER BY Position")
                  If rstEntry.RecordCount = 0 Then
                      With rstEntry
                          .AddNew
                          .Fields("Code") = TestCode
                          .Fields("Sta") = Left$(cSta(iTemp), 3)
                          .Fields("Position") = 1
                          .Fields("Group") = 1
                          .Fields("RR") = False
                          .Fields("Deleted") = 0
                          .Fields("Status") = TestStatus
                          .Fields("Timestamp") = Now
                          .Update
                      End With
                  Else
                      iPosition = 0
                      With rstEntry
                          .AddNew
                          .Fields("Code") = TestCode
                          .Fields("Sta") = Left$(cSta(iTemp), 3)
                          .Fields("Position") = 1 '0
                          .Fields("Group") = 1
                          .Fields("RR") = False
                          .Fields("Deleted") = 0
                          .Fields("Status") = TestStatus
                          .Fields("Timestamp") = Now
                          .Update
                      End With
                  End If
                  If TestStatus > 0 Then
                      AddColorsToFinals
                  Else
                      AddOneToPosition TestCode, TestStatus
                  End If
                  ChangeCaption True
            Else
                  MsgBox cSta(iTemp) & " " & Translate("already entered for this test.", mcLanguage)
            End If
            rstEntry.Close
            Set rstEntry = Nothing
        End If
        Next iTemp
        Me.Tempvar = ""
    End If
    
End Sub
Private Sub mnuEditChangeRein_Click()
   Dim iKey As Integer
   Dim rstEntry As DAO.Recordset
   Dim cSta As String
   
   If TestStatus > 0 Then
        MsgBox Translate("This has no use in finals!", mcLanguage)
   Else
        'what participant is selected?
        If Me.txtParticipant <> "" Then
            cSta = Me.txtParticipant
            Me.dblstNotYet.Text = ""
        ElseIf Me.dblstNotYet.BoundText <> "" Then
            cSta = Left$(Me.dblstNotYet.BoundText, 3)
            Me.txtParticipant.Text = cSta
        End If
        If cSta <> "" Then
             Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Results WHERE Sta LIKE '" & cSta & "' AND Code LIKE '" & TestCode & "' AND STATUS=" & Me.TestStatus)
             If chkRein.Enabled = True And rstEntry.RecordCount = 0 Then
                 iKey = MsgBox(Translate("Change rein for startnumber", mcLanguage) & " " & cSta & "?", vbQuestion + vbYesNo + vbDefaultButton1)
                 If iKey = vbYes Then
                       Set rstEntry = mdbMain.OpenRecordset("SELECT RR FROM Entries WHERE Sta='" & cSta & "' AND Code='" & TestCode & "' AND STATUS=" & Me.TestStatus)
                       With rstEntry
                           If .RecordCount > 0 Then
                               .Edit
                               If IsNull(.Fields("RR")) Then
                                    .Fields("RR") = True
                               Else
                                    .Fields("RR") = Not .Fields("RR")
                               End If
                               .Update
                           Else
                               MsgBox Translate("This participant is not entered for this test yet. No rein to change!", mcLanguage), vbInformation
                           End If
                       End With
                 End If
                 If chkRein.Value = 0 Then
                      chkRein.Value = 1
                 End If
                 ChangeCaption True
                 txtParticipant_Change
             Else
                MsgBox Translate("This has no use for this participant!", mcLanguage)
             End If
             rstEntry.Close
             Set rstEntry = Nothing
         Else
            MsgBox Translate("Select a participant first.", mcLanguage), vbExclamation
         End If
   End If

End Sub


Private Sub mnuEditNoStart_Click()
   Dim iKey As Integer
   Dim rstEntry As DAO.Recordset
   Dim cSta As String
   Dim iNoStart As Integer
   
   If TestStatus > 0 Then
        MsgBox Translate("This has no use in finals!", mcLanguage)
   Else
        'what participant is selected?
        If Me.txtParticipant <> "" Then
            cSta = Me.txtParticipant
            Me.dblstNotYet.Text = ""
        ElseIf Me.dblstNotYet.BoundText <> "" Then
            cSta = Left$(Me.dblstNotYet.BoundText, 3)
            Me.txtParticipant.Text = cSta
        End If
        If cSta <> "" Then
            iNoStart = 1
            Set rstEntry = mdbMain.OpenRecordset("SELECT NoStart FROM Entries WHERE Sta='" & cSta & "' AND Code='" & TestCode & "' AND STATUS=" & Me.TestStatus)
            With rstEntry
                If .RecordCount > 0 Then
                    iNoStart = IIf(IsNull(.Fields("NoStart")), 0, .Fields("NoStart"))
                    If .Fields("NoStart") = 0 Or IsNull(.Fields("NoStart")) = True Then
                          iKey = MsgBox(cSta & " " & Translate("will not start in the next heat(s)", mcLanguage) & "?", vbQuestion + vbYesNo + vbDefaultButton1)
                          If iKey = vbYes Then
                              iNoStart = -1
                          End If
                  Else
                          iKey = MsgBox(cSta & " " & Translate("will start (again) in the next heat(s)", mcLanguage) & "?", vbQuestion + vbYesNo + vbDefaultButton1)
                          If iKey = vbYes Then
                              iNoStart = 0
                          End If
                  End If

                    .Edit
                    .Fields("NoStart") = iNoStart
                    .Update
                Else
                    MsgBox Translate("This participant is not entered for this test yet. No start planned!", mcLanguage), vbInformation
                End If
            End With
            ChangeCaption True
            txtParticipant_Change
             rstEntry.Close
             Set rstEntry = Nothing
         Else
            MsgBox Translate("Select a participant first.", mcLanguage), vbExclamation
         End If
   End If
Exit Sub

End Sub

Private Sub mnuEditEdit_Click()
    Dim fPart As New frmParticipant
    'what participant is selected?
    If Me.txtParticipant <> "" Then
        fPart.cSta = Me.txtParticipant
        Me.dblstNotYet.Text = ""
    Else
        fPart.cSta = Left$(Me.dblstNotYet.BoundText, 3)
        Me.txtParticipant = fPart.cSta
    End If
    
    fPart.Show 1, Me
    
    ChangeCaption True
End Sub
Private Sub mnuEditFind_Click()
    Dim cTemp As String
    Dim cQry As String
    
    cTemp = InputBox$(Translate("Enter text to search for", mcLanguage), Translate("Find a participant", mcLanguage), mnuEditFind.Tag)
   
    mnuEditFind.Tag = cTemp
    
    cQry = "SELECT Participants.Sta "
    cQry = cQry & " & '  -  ' & Persons.Name_First"
    cQry = cQry & " & ' ' & Persons.Name_Last"
    cQry = cQry & " & IIF(Participants.Class<>'',' [' & Participants.Class & ']','')"
    cQry = cQry & " & ' - ' & Horses.Name_Horse as cList"
    cQry = cQry & " FROM (Participants"
    cQry = cQry & " INNER JOIN Persons"
    cQry = cQry & " ON Participants.PersonId=Persons.PersonId)"
    cQry = cQry & " INNER JOIN Horses"
    cQry = cQry & " ON Participants.HorseId=Horses.HorseId"
    cQry = cQry & " WHERE Participants.Sta & ' ' & Persons.Name_First & ' ' & Persons.Name_Last & ' - ' & Horses.Name_Horse & ' [' & Participants.Class & ']'  LIKE " & Chr$(34) & "*" & cTemp & "*" & Chr$(34)
    cQry = cQry & " ORDER BY Participants.Sta"
    
    frmToolBox.strQry = cQry
    frmToolBox.Caption = Translate("Searching", mcLanguage) & " '" & cTemp & "' "
    frmToolBox.Show 1, Me
    
    If Me.Tempvar <> "" Then
       txtParticipant.Text = Me.Tempvar
       Me.Tempvar = ""
    End If
End Sub

Public Sub mnuEditNew_Click()
    Dim cTemp As String
    Dim rstSta As Recordset
    Dim fPart As New frmParticipant
    
    Set rstSta = mdbMain.OpenRecordset("SELECT Sta FROM Participants ORDER BY Sta DESC")
    If rstSta.RecordCount > 0 Then
        cTemp = Format$(rstSta.Fields("Sta") + 1, "000")
    Else
        cTemp = "001"
    End If
    rstSta.Close
    Set rstSta = Nothing
    
    cTemp = InputBox$(Translate("Enter a new start number for this new participant.", mcLanguage), , cTemp)
    
    If cTemp <> "" Then
        MsgBox cTemp & " " & Translate("will be added as new participant; select the name of the rider and the horse from the relevant lists.", mcLanguage) & vbCrLf & Translate("You may add a new rider or a new horse by using the respective 'New' buttons.", mcLanguage), vbInformation
        fPart.cNewSta = cTemp
        fPart.Show 1, Me
        ChangeCaption True
    End If
End Sub

Private Sub mnuEditRemove_Click()
   Dim iKey As Integer
   Dim rstOk As Recordset
   
   If lblParticipant.Caption <> "" Then
      iKey = MsgBox(Translate("You may remove all data for the selected participant." & vbCrLf & "- select 'Yes' to remove the participant completely" & vbCrLf & "- select 'No' to remove marks only" & vbCrLf & "- select 'Cancel' to keep this participant", mcLanguage), vbExclamation + vbYesNoCancel + vbDefaultButton3, Translate("Remove from", mcLanguage) & ": " & ClipAmp(tbsSelFin.SelectedItem))
      If iKey = vbYes Then
         GoSub RemoveMarks
         GoSub RemoveEntry
      ElseIf iKey = vbNo Then
         GoSub RemoveMarks
      End If
      ChangeCaption True
   Else
      MsgBox Translate("Select a participant first.", mcLanguage), vbExclamation
   End If
   Set rstOk = Nothing
Exit Sub

RemoveEntry:
   'is this participant already an entry?
   Set rstOk = mdbMain.OpenRecordset("SELECT * FROM ENTRIES WHERE Sta='" & txtParticipant.Text & "' AND Status=" & Me.TestStatus & " AND Code='" & Me.TestCode & "'")
   With rstOk
      If .RecordCount > 0 Then
         .Delete
      End If
      .Close
   End With
   
Return

RemoveMarks:
   'are there already marks?
   Set rstOk = mdbMain.OpenRecordset("SELECT * FROM MARKS WHERE Sta='" & txtParticipant.Text & "' AND Status=" & Me.TestStatus & " AND Code='" & Me.TestCode & "'")
   With rstOk
      If .RecordCount > 0 Then
        Do While Not .EOF
         .Delete
         .Requery
        Loop
      End If
      .Close
   End With
   Set rstOk = mdbMain.OpenRecordset("SELECT * FROM Results WHERE Sta='" & txtParticipant.Text & "' AND Status=" & Me.TestStatus & " AND Code='" & Me.TestCode & "'")
   With rstOk
      If .RecordCount > 0 Then
        Do While Not .EOF
         .Delete
         .Requery
        Loop
      End If
      .Close
   End With
   ClearMarks
   LookUpRelevantParticipants
Return
End Sub

Private Sub mnuEditSelect_Click()
   If dblstNotYet.BoundText <> "" Then
      dblstNotYet_DblClick
   ElseIf lstAlready.Text <> "" Then
      lstAlready_DblClick
   End If
End Sub

Private Sub mnuFileEvenCode_Click()
    Dim cTemp As String
    
    frmWR.Show 1, Me
    
    cTemp = GetVariable("WR_code")
    If cTemp <> "" And cTemp <> Chr$(27) Then
        mnuFileFEIFWR.Enabled = True
    Else
        mnuFileFEIFWR.Enabled = False
    End If
End Sub

Private Sub mnuFileEvenName_Click()
    frmEvent.Show 1, Me
    ChangeCaption True
End Sub

Private Sub mnuFileBackup_Click()
    Me.Enabled = False
    If CreateBackup(mdbMain) = True Then
        StatusMessage Translate("Backup created successfully.", mcLanguage)
    End If
    Me.Enabled = True
End Sub

Private Sub mnuFileExit_Click()
   Unload Me
   End
End Sub

Private Sub mnuFileFEIFWR_Click()
    Me.Enabled = False
    If CreateWRFile = True Then
        MsgBox Translate("FEIF WorldRanking File created successfully. Please send this file to your national FEIF WorldRanking Registrar.", mcLanguage)
    End If
    Me.Enabled = True
End Sub

Private Sub mnuFilePrintFormEntrance_Click()
    Dim iKey As Integer
    Dim cQry As String
    Dim cTemp As String
    Dim cStaList As String
    
    iKey = MsgBox(Translate("Sort Entrance Check Form by Name of Rider (Yes) or by Starting Number (No).", mcLanguage), vbYesNoCancel + vbQuestion)
    If iKey = vbYes Then
        PrintEntranceForm 0
    ElseIf iKey = vbNo Then
        PrintEntranceForm 1
    End If

End Sub

Private Sub mnuFilePrintFormEquipmentComplete_Click()
    PrintStartForm 1
End Sub

Private Sub mnuFilePrintFormEquipmentOnly_Click()
    PrintStartForm 2
End Sub

Private Sub mnuFilePrintFormJudges_Click()
    PrintJudgesForm
End Sub

Private Sub mnuFilePrintFormJudges3_Click()
    PrintJudgesForm3
End Sub

Private Sub mnuFilePrintFormJudgesLand_Click()
    If Dir$(App.Path & "\iceform.exe") <> "" Then
        Shell App.Path & "\iceform.exe Test=" & TestCode & " Status=" & TestStatus, vbNormalFocus
    ElseIf Dir$(Replace(App.Path, "IceTest", "IceHorseTools") & "\iceform.exe") <> "" Then
        Shell Replace(App.Path, "IceTest", "IceHorseTools") & "\iceform.exe Test=" & TestCode & " Status=" & TestStatus, vbNormalFocus
    End If
End Sub

Private Sub mnuFilePrintFormStart_Click()
    PrintStartForm
End Sub

Private Sub mnuFilePrintFormTime_Click()
    PrintTimeForm
End Sub

Private Sub mnuFilePrintFormVet_Click()
    Dim iKey As Integer
    Dim cQry As String
    Dim cTemp As String
    Dim cStaList As String
    
    iKey = MsgBox(Translate("Print Veterinary Checks Form for all participants (Yes) or for selected participants (No).", mcLanguage), vbYesNoCancel + vbQuestion)
    If iKey = vbYes Then
        PrintVetForm
    ElseIf iKey = vbNo Then
    
        cTemp = InputBox$(Translate("Search for", mcLanguage), Translate("Participants", mcLanguage), mnuEditAdd.Tag)
   

        cQry = "SELECT Participants.Sta "
        cQry = cQry & " & '  -  ' & Persons.Name_First"
        cQry = cQry & " & ' ' & Persons.Name_Last"
        cQry = cQry & " & IIF(Participants.Class<>'',' ['  & participants.Class & ']','') "
        cQry = cQry & " & ' - ' & Horses.Name_Horse as cList"
        cQry = cQry & " FROM (Participants"
        cQry = cQry & " INNER JOIN Persons"
        cQry = cQry & " ON Participants.PersonId=Persons.PersonId)"
        cQry = cQry & " INNER JOIN Horses"
        cQry = cQry & " ON Participants.HorseId=Horses.HorseId"
        cQry = cQry & " WHERE Participants.Sta & ' ' & Persons.Name_First & ' ' & Persons.Name_Last & ' - ' & Horses.Name_Horse LIKE " & Chr$(34) & "*" & cTemp & "*" & Chr$(34)
        cQry = cQry & " ORDER BY Participants.Sta"
        
        frmToolBox.strQry = cQry
        frmToolBox.intChecked = True
        frmToolBox.intReturnLen = 3
        frmToolBox.Caption = Translate("Searching", mcLanguage) & " '" & cTemp & "' "
        frmToolBox.Show 1, Me
        
        cStaList = "'" & Replace(Me.Tempvar, "|", "','") & "'"
        Me.Tempvar = ""
        PrintVetForm cStaList
    End If
End Sub

Private Sub mnuFilePrintLog_Click()
    PrintLogFile
End Sub

Private Sub mnuFilePrintOverviewFinals_Click()
    PrintOverview 2
End Sub

Private Sub mnuFilePrintOverviewMarks_Click()
    PrintOverview 1
End Sub

Private Sub mnuFilePrintOverviewNoMarks_Click()
    PrintOverview 0
End Sub

Private Sub mnuFilePrintResultFinal_Click()
    Select Case TestStatus
    Case 0
        SetTestStatus TestCode, 1
    Case 1
        SetTestStatus TestCode, 3
    Case 2
        SetTestStatus TestCode, 2
    Case 3
        SetTestStatus TestCode, 4
    End Select
    PrintResultList ClipAmp(mnuFilePrintResultFinal.Caption), False, IIf(TestStatus = 0, True, False)
End Sub

Private Sub mnuFilePrintResultInterim_Click()
    mdbMain.Execute "UPDATE Results SET Position=0 WHERE Code='" & TestCode & "' And Status=" & TestStatus & ";"
    PrintResultList ClipAmp(mnuFilePrintResultInterim.Caption)
    If GetTestStatus(TestCode) = 1 Then
        SetTestStatus TestCode, 0, True
    End If
End Sub

Private Sub mnuFilePrintResultRevised_Click()
    PrintResultList ClipAmp(mnuFilePrintResultRevised.Caption) & " [" & Format$(Now, "D-M-YYYY HH:MM") & "]", False, IIf(TestStatus = 0, True, False)
End Sub

Private Sub mnuFilePrintRider_Click()
    PrintParticipant
End Sub

Private Sub mnuFilePrintWarnings_Click()
    PrintWarnings
End Sub

Private Sub mnuFileResultsExcel_Click()
    Dim cTemp As String
    If Dir$(mcExcelDir & "*.csv") <> "" Then
        On Local Error Resume Next
        With CommonDialog1
            .InitDir = mcExcelDir
            .CancelError = True
            .DefaultExt = "Csv"
            .DialogTitle = Translate("Select results", mcLanguage)
            .Filter = Translate("Results", mcLanguage) & "|*.Csv|" & Translate("All files", mcLanguage) & "|*.*"
            .FilterIndex = 1
            .Flags = cdlOFNCreatePrompt Or cdlOFNHideReadOnly
            .ShowOpen
            If Err = cdlCancel Then
                Exit Sub
            End If
            If .FileName <> "" Then
                'Show in default spread sheet
                If ShowDocument(.FileName, Me) = 3 Then
                    MsgBox Translate("Please install MS Excel first.", mcLanguage), vbExclamation
                End If
            End If
        End With
        On Local Error GoTo 0
    Else
        MsgBox Translate("No results available", mcLanguage) & "!"
    End If
End Sub

Private Sub mnuFileResultsHtml_Click()
    Dim X As Integer
    
    'if the automated HTML option is switched off, copy the temp files first:
    If miHtmlFiles = 0 Then
        X = CopyHTML(mcTempHtmlDir, mcHtmlDir)
    End If
    
    If Dir$(mcHtmlDir & "index.html") <> "" Then
        If ShowDocument(mcHtmlDir & "index.html", Me) = 3 Then
            MsgBox Translate("Please install a web browser (like MS Internet Explorer) first.", mcLanguage), vbExclamation
        End If
    Else
        MsgBox Translate("No results available", mcLanguage) & "!"
    End If
End Sub

Private Sub mnuFileResultsRtf_Click()
    Dim cTemp As String
    On Local Error Resume Next
    With CommonDialog1
        .InitDir = mcRtfDir
        .CancelError = True
        .DefaultExt = "Rtf"
        .DialogTitle = Translate("Select a result list", mcLanguage)
        .Filter = Translate("Result Lists", mcLanguage) & "|*.Rtf|" & Translate("All files", mcLanguage) & "|*.*"
        .FilterIndex = 1
        .Flags = cdlOFNCreatePrompt Or cdlOFNHideReadOnly
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        If .FileName <> "" Then
            'Show in default editor
            If ShowDocument(.FileName, Me) = 3 Then
                MsgBox Translate("Please install editor for RTF-files first (like MS Word).", mcLanguage), vbExclamation
            End If
        End If
    End With
    On Local Error GoTo 0

End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show 1, Me
End Sub

' Checks FEIF website for program updates
' Update server can be stored in database variable "webupdates". If this variable should
' not exist, we will check this URL for update information:
' http://feif.glitnir.nl/icehorse/icehorsetools_version.txt
' icehorsetools_version.txt can hold info on several applications, one line per app:
' [app.EXEName],[app.major],[app.minor],[app.revision],[URL for Download]
' example:
' IceTest,1,0,416,http://www.feif.org/software/
Private Sub mnuHelpCheckForUpdate_Click()
    Dim X As Integer
    Dim entry() As String
    Dim pos As Integer
    Dim segment() As String
    Dim spos As Integer
    
    X = MsgBox(Translate("Check FEIF server for updates now?", mcLanguage) & vbCrLf & Translate("Choose YES if this computer is connected to the Internet right now, otherwise select NO.", mcLanguage), vbYesNo + vbQuestion, Translate("Check for update", mcLanguage))
    
    If X = vbYes Then
        Call CheckWebForUpdate
        
        'wait until connection is closed
        While blnConnected = True
            DoEvents
        Wend
        
        If winsockresponse <> "" Then
            'parse the server's response:
            entry = Split(winsockresponse, vbCrLf, , vbTextCompare)
            pos = 0

            Do While pos < UBound(entry)
                If Trim$(entry(pos)) <> "" Then
                    If Left$(entry(pos), Len(App.EXEName)) = App.EXEName Then
                        'split the current line into its five segments:
                        segment = Split(entry(pos), ",", , vbTextCompare)
                        'compare the version with the retrieved info:
                        If (CInt(segment(1)) > App.Major) Or _
                            (CInt(segment(1)) = App.Major And CInt(segment(2)) > App.Minor) Or _
                            (CInt(segment(1)) = App.Major And CInt(segment(2)) = App.Minor And CInt(segment(3)) > App.Revision) Then
                            
                                X = MsgBox(Translate("A new version is available on the server", mcLanguage) & ": " & segment(1) & "." & segment(2) & "." & segment(3) & vbCrLf & Translate("Would you like to download the update now?", mcLanguage), vbExclamation + vbYesNo, Translate("Found update", mcLanguage))
                                If X = vbYes Then
                                    ShowDocument segment(4), Me
                                End If
                                Exit Sub
                        Else
                            X = MsgBox(Translate("You already have the latest version of IceTest.", mcLanguage), vbOKOnly + vbInformation)
                            Exit Sub
                        End If
                        
                    End If
                End If
                pos = pos + 1
            Loop
            X = MsgBox(Translate("Found no info about this program on the update server", mcLanguage), vbOKOnly + vbExclamation)
        End If
        
    End If
End Sub

Private Sub mnuHelpFEIF_Click()
    ShowDocument "https://www.feif.org/icetest", Me
End Sub

Private Sub mnuHelpFEIFTech_Click()
    NotYet
End Sub

Private Sub mnuHelpFEIFWR_Click()
    Dim cWRCode As String
    
    cWRCode = WrTest(Me.TestCode)
    
    ShowDocument "https://www.feif.org/worldranking/" & LCase$(cWRCode), Me
End Sub

Private Sub mnuHelpIcetest_Click()
    ShowDocument "https://www.feif.org/icetest", Me
End Sub

Private Sub mnuHelpIndex_Click()
    With CommonDialog1
        .HelpFile = App.HelpFile
        .HelpCommand = cdlHelpContents
        .HelpContext = 0
        .HelpKey = ""
        .ShowHelp
    End With
End Sub

Private Sub mnuPopupPopup_Click(Index As Integer)
   Dim cTemp As String
   cTemp = ReadTagItem(mnuPopupPopUp(Index), "Control")
   Select Case cTemp
   Case "mnuEditStartOrder"
        mnuEditStartOrder_Click
   Case "mnuEditMove"
        mnuEditMove_Click
   Case "mnuFilePrintFormJudgesLand"
        mnuFilePrintFormJudgesLand_Click
   Case "mnuFilePrintFormTime"
        mnuFilePrintFormTime_Click
   Case "mnuFilePrintFormVet"
        mnuFilePrintFormVet_Click
    Case "mnuFilePrintOverviewFinals"
      mnuFilePrintOverviewFinals_Click
    Case "mnuFilePrintOverviewMarks"
      mnuFilePrintOverviewMarks_Click
    Case "mnuFilePrintOverviewNoMarks"
      mnuFilePrintOverviewNoMarks_Click
    Case "mnuEditChangeRein"
      mnuEditChangeRein_Click
    Case "mnuEditNoStart"
      mnuEditNoStart_Click
    Case "mnuFileResultsHtml"
      mnuFileResultsHtml_Click
    Case "mnuFileResultsExcel"
      mnuFileResultsExcel_Click
    Case "mnuFileResultsRtf"
      mnuFileResultsRtf_Click
    Case "mnuEditFind"
      mnuEditFind_Click
   Case "mnuFilePrintFormStart"
      mnuFilePrintFormStart_Click
   Case "mnuFilePrintFormJudges"
      mnuFilePrintFormJudges_Click
   Case "mnuFilePrintFormJudges3"
      mnuFilePrintFormJudges3_Click
   Case "mnuFilePrintLog"
      mnuFilePrintLog_Click
   Case "mnuFilePrintResultFinal"
      mnuFilePrintResultFinal_Click
   Case "mnuFilePrintResultInterim"
      mnuFilePrintResultInterim_Click
   Case "mnuFilePrintResultRevised"
      mnuFilePrintResultRevised_Click
   Case "mnuFileExit"
      mnuFileExit_Click
   Case "mnuTestQual"
      StoreCurrentMarks
      Me.TestCode = ReadTagItem(mnuPopupPopUp(Index), "Tag")
      LookUpTest
   Case "mnuEditSelect"
      mnuEditSelect_Click
   Case "mnuEditRemove"
      mnuEditRemove_Click
   Case "mnuEditNew"
      mnuEditNew_Click
   Case "mnuEditAdd"
      mnuEditAdd_Click
   Case "mnuEditEdit"
      mnuEditEdit_Click
   Case "mnuFilePrintRider"
      mnuFilePrintRider_Click
   Case Else
      NotYet
   End Select
End Sub

Private Sub mnuPopUpTestsTest_Click(Index As Integer)
   Dim cTemp As String
   cTemp = ReadTagItem(Me.mnuPopUpTestsTest(Index), "Control")
   Select Case cTemp
   Case "mnuTestQual"
      StoreCurrentMarks
      Me.TestCode = ReadTagItem(Me.mnuPopUpTestsTest(Index), "Tag")
      LookUpTest
   Case Else
      NotYet
   End Select
End Sub

Private Sub mnuTestAddNew_Click()
   CreateNewTest
   CreateTestMenu
End Sub

Private Sub mnuTestAll_Click()
    Dim rstTests As DAO.Recordset
    Set rstTests = mdbMain.OpenRecordset("SELECT Nr FROM TestInfo WHERE Nr>0")
    If rstTests.RecordCount > 0 Then
        mnuTestAll.Checked = Not mnuTestAll.Checked
        CreateTestMenu
    Else
        mnuTestAll.Checked = True
        MsgBox Translate("Select a list of tests for this event first", mcLanguage) & "!", vbCritical
        mnuFileEvenTest_Click
    End If
    rstTests.Close
    Set rstTests = Nothing
    WriteIniFile gcIniFile, Me.Name, "ShowAllTests", IIf(mnuTestAll.Checked = True, "-1", "0")
End Sub

Private Sub mnuTestEdit_Click()
    With frmToolBox
        .strQry = "SELECT Code FROM Tests WHERE (Removed=False or ISNULL(Removed)) ORDER BY Code"
        '.strQry = "SELECT Code FROM Tests WHERE Code IN (SELECT Code FROM TestSections WHERE Status=1) AND (Removed=False or ISNULL(Removed)) ORDER BY Code"
        .Caption = ClipAmp(mnuTestEdit.Caption)
        .Show 1, Me
    End With
    If Me.Tempvar <> "" Then
        With frmTests
            .fcInitCode = Tempvar
            Tempvar = ""
           .Show 1, Me
        End With
        CreateTestMenu
        LookUpTest
    End If
End Sub

Private Sub mnuTestQual10Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual10Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual11Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual11Test(Index).Tag
   LookUpTest
End Sub

Private Sub mnuTestQual12Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual12Test(Index).Tag
   LookUpTest
End Sub

Private Sub mnuTestQual13Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual13Test(Index).Tag
   LookUpTest
End Sub

Private Sub mnuTestQual14Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual14Test(Index).Tag
   LookUpTest
End Sub

Private Sub mnuTestQual15Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual15Test(Index).Tag
   LookUpTest
End Sub

Private Sub mnuTestQual16Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual16Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual17Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual17Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual18Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual18Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual19Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual19Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual1Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual1Test(Index).Tag
   LookUpTest
End Sub

Private Sub mnuTestQual20Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual20Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual21Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual21Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual22Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual22Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual23Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual23Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual24Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual24Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual25Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual25Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual26Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual26Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual27Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual27Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual28Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual28Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual29Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual29Test(Index).Tag
   LookUpTest

End Sub

Private Sub mnuTestQual2Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual2Test(Index).Tag
   LookUpTest
End Sub
Private Sub mnuTestQual3Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual3Test(Index).Tag
   LookUpTest

End Sub
Private Sub mnuTestQual4Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual4Test(Index).Tag
   LookUpTest

End Sub
Private Sub mnuTestQual5Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual5Test(Index).Tag
   LookUpTest
End Sub


Private Sub mnuTestQual6Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual6Test(Index).Tag
   LookUpTest
End Sub


Private Sub mnuTestQual7Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual7Test(Index).Tag
   LookUpTest
End Sub


Private Sub mnuTestQual8Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual8Test(Index).Tag
   LookUpTest
End Sub

Private Sub mnuTestQual9Test_Click(Index As Integer)
   StoreCurrentMarks
   Me.TestCode = mnuTestQual9Test(Index).Tag
   LookUpTest
End Sub

Private Sub mnuFileChange_Click()
    Dim iOldBackupInterval
    Dim iKey As Integer
    
    Me.Enabled = False
    
    If miBackupInterval = 0 Then
        iOldBackupInterval = miBackupInterval
        miBackupInterval = -1
        iKey = MsgBox(Translate("Create backup first?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton1)
        If iKey = vbYes Then
            CreateBackup mdbMain
        End If
    End If
    ChangeDatabase
    miBackupInterval = iOldBackupInterval
    Me.Enabled = True
End Sub

Private Sub mnuTestRemove_Click()
    RemoveTest
    CreateTestMenu
End Sub

Private Sub mnuToolCompress_Click()
    Dim iKey As Integer
    Dim cTemp As String
    Dim iOldBackupInterval As Integer
    Dim tdf As DAO.TableDef
    
    On Local Error Resume Next
    If InStr(Command$, "/C") Then
        iKey = vbYes
    Else
        iKey = MsgBox(Translate("Do you want to compact (and repair) the database (this might take some time)?", mcLanguage), vbYesNo + vbQuestion)
    End If
    If iKey = vbYes Then
        miBackupInterval = -1
        If InStr(Command$, "/C") Then
            iKey = vbNo
        Else
            iKey = MsgBox(Translate("Create backup first?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton1)
        End If
        If iKey = vbYes Then
            CreateBackup mdbMain
        End If
        
        Me.Enabled = False
        'Clean up some tables first
        mdbMain.Execute ("DELETE * FROM Entries WHERE STA NOT IN (SELECT STA FROM Participants) OR STA='' OR ISNULL(Sta)")
        If mdbMain.RecordsAffected > 0 Then
            LogLine mdbMain.RecordsAffected & " records removed from Entries"
        End If
        mdbMain.Execute ("DELETE * FROM Marks WHERE STA NOT IN (SELECT STA FROM Entries) OR STA='' OR ISNULL(Sta)")
        If mdbMain.RecordsAffected > 0 Then
            LogLine mdbMain.RecordsAffected & " records removed from Marks"
        End If
        mdbMain.Execute ("DELETE * FROM Results WHERE STA NOT IN (SELECT STA FROM Marks) OR STA='' OR ISNULL(Sta)")
        If mdbMain.RecordsAffected > 0 Then
            LogLine mdbMain.RecordsAffected & " records removed from Results"
        End If
        
        'Delete _Temp tables
        For Each tdf In mdbMain.TableDefs
            If Left$(tdf.Name, 5) = "_Temp" Then
                cTemp = tdf.Name
                mdbMain.Execute ("Drop Table [" & cTemp & "]")
                LogLine "Table " & cTemp & " removed"
            End If
        Next
        
        SetVariable "ProgramVersion", ""
        mdbMain.Close
        DoEvents
        
        If CheckDatabase(mcDatabaseName) = True Then
            OpenDatabase mcDatabaseName
            cTemp = App.Major & "." & App.Minor & "." & Format$(App.Revision, "000")
            SetVariable "ProgramVersion", cTemp
        Else
            OpenDatabase mcDatabaseName
        End If
        CheckOnIndexes
        
        mdbMain.Close
        DoEvents
        
        If CompressDatabase = True Then
            MsgBox Translate("Database successfully compressed (and repaired).", mcLanguage)
            LogLine "Database successfully compressed."
        End If
        
        
        Me.Enabled = True
    End If
    On Local Error GoTo 0
End Sub

Private Sub mnuToolFormDel_Click()
    DeleteRTF
End Sub

Private Sub mnuToolFormEdit_Click()
    EditRtf
End Sub

Private Sub mnuToolFormNew_Click()
    NewRTF
End Sub

Private Sub mnuToolImportAccess_Click()
    Me.Enabled = False
    ImportFeif
    Me.Enabled = True
    ChangeCaption True
End Sub

Private Sub mnuToolImportBackup_Click()
    ReadBackUp mdbMain
    ChangeCaption True
End Sub

Private Sub mnuToolImportCsv_Click()
    ImportCsv
    ChangeCaption True
End Sub

Private Sub mnuToolImportExcel_Click()
    ImportExcel
    ChangeCaption True
End Sub

Private Sub mnuToolImportFipo_Click()
    ImportFeif True
    ChangeCaption True
End Sub

Private Sub mnuToolImportIPZV_Click()
End Sub

Private Sub mnuToolImportDI_Click()
    Me.Enabled = False
    If ImportDI = True Then
        MsgBox Translate("Participants successfully imported.", mcLanguage)
    End If
    Me.Enabled = True
    ChangeCaption True

End Sub


Private Sub mnuToolImportMarksExcel_Click()
    ImportTestFromExcel
End Sub

Private Sub mnuToolImportNSIJP_Click()
    Me.Enabled = False
    If ImportNSIJP = True Then
        MsgBox Translate("Participants successfully imported.", mcLanguage)
    End If
    Me.Enabled = True
    ChangeCaption True
End Sub

Private Sub mnuFileNew_Click()
    Dim iOldBackupInterval
    Dim iKey As Integer
    
    Me.Enabled = False
    
    If miBackupInterval = 0 Then
        iOldBackupInterval = miBackupInterval
        miBackupInterval = -1
        iKey = MsgBox(Translate("Create backup first?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton1)
        If iKey = vbYes Then
            CreateBackup mdbMain
        End If
    End If
    CreateNewDatabase
    miBackupInterval = iOldBackupInterval
    Me.Enabled = True
End Sub

Private Sub mnuToolImportTab_Click()
    ImportTab
    ChangeCaption True
End Sub

Private Sub mnuToolOptions_Click()
    frmOptions.Show 1, Me
    If Tempvar <> "" Then
        mdbMain.Close
        RestartApp
    End If
End Sub

Private Sub mnuToolOutputExcel_Click()
    ExportExcel
End Sub

Private Sub mnuToolOutputCsv_Click()
    ExportCsv
End Sub


Private Sub mnuToolOutputHtml_Click()
    Dim cTemp As String
    Dim iKey As String
    
    On Local Error Resume Next
    cTemp = PickDirFromTree(Me.hwnd, Translate("Select a folder for files for a default browser (Html).", mcLanguage))
    If cTemp <> "" Then
        If Right$(cTemp, 1) <> "\" Then
            cTemp = cTemp & "\"
        End If
        mcHtmlDir = cTemp
        If Dir$(mcHtmlDir, vbDirectory) <> "" Then
            WriteIniFile gcIniFile, "Html Files", "Folder", mcHtmlDir
            If Dir$(mcHtmlDir & "*.html") <> "" Then
                iKey = MsgBox(Translate("Delete files in current folder?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
                If iKey = vbYes Then
                    KillFile mcHtmlDir & "*.htm?"
                End If
            End If
        Else
            MsgBox cTemp & ": " & Translate("folder not found!", mcLanguage), vbExclamation
        End If
    End If
    On Local Error GoTo 0
End Sub


Private Sub mnuToolOutputJudgesExcel_Click()
    ExportTestToExcelForJudges
End Sub

Private Sub mnuToolOutputMarksExcel_Click()
    ExportTestToExcel
End Sub

Public Sub mnuToolOutputOntrackdata_Click()
          
    If ontrack = False Then
        hwndOldOwner = GetWindow(frmOntrack.hwnd, GW_OWNER)
        SetWindowLong frmOntrack.hwnd, GWL_HWNDPARENT, Me.hwnd
        frmOntrack.Show
        'Call EinblendungEinzelFinanz
        ontrack = True
    Else
        frmOntrack.Hide
        ontrack = False
    End If

End Sub

Private Sub mnuToolOutputRtf_Click()
    Dim cTemp As String
    Dim iKey As String
    
    On Local Error Resume Next
    cTemp = PickDirFromTree(Me.hwnd, Translate("Select a folder for result lists (Rtf)", mcLanguage))
    If cTemp <> "" Then
        If Right$(cTemp, 1) <> "\" Then
            cTemp = cTemp & "\"
        End If
        mcRtfDir = cTemp
        If Dir$(mcRtfDir, vbDirectory) <> "" Then
            WriteIniFile gcIniFile, "Rtf Files", "Folder", mcRtfDir
            If Dir$(mcRtfDir & "*.Rtf") <> "" Then
                iKey = MsgBox(Translate("Delete files in current folder?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
                If iKey = vbYes Then
                    KillFile mcRtfDir & "*.Rtf"
                End If
            End If
        Else
            MsgBox cTemp & ": " & Translate("folder not found!", mcLanguage), vbExclamation
        End If
    End If
    On Local Error GoTo 0

End Sub

Private Sub mnuToolOutputExternal_Click()
    Dim cTemp As String
    Dim iKey As String
    
    On Local Error Resume Next
    cTemp = PickDirFromTree(Me.hwnd, Translate("Select a folder for Extra files for external functions.", mcLanguage))
    If cTemp <> "" Then
        If Right$(cTemp, 1) <> "\" Then
            cTemp = cTemp & "\"
        End If
        mcExcelDir = cTemp
        If Dir$(mcExcelDir, vbDirectory) <> "" Then
            WriteIniFile gcIniFile, "Excel Files", "Folder", mcExcelDir
            If Dir$(mcExcelDir & "*.*") <> "" Then
                iKey = MsgBox(Translate("Delete files in current folder?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
                If iKey = vbYes Then
                    KillFile mcExcelDir & "*.*"
                End If
            End If
        Else
            MsgBox cTemp & ": " & Translate("folder not found!", mcLanguage), vbExclamation
        End If
    End If
    On Local Error GoTo 0

End Sub

Private Sub mnuFileSaveAsItem_Click(Index As Integer)
    Dim rstXls As DAO.Recordset
    Dim cExcelFile As String
    Dim iExcelFile As Integer
    Dim iKey As Integer
    
    Dim xlObj As Object
    Dim iTemp As Integer
    Dim cTemp As String
    Dim iTemp2 As Integer
    Dim iRow As Integer
    Dim iColumn As Integer
    Dim iColumnCount As Integer
    Dim cColumn() As String
    
    On Local Error Resume Next
    
    cExcelFile = mcExcelDir & Replace(mnuFileSaveAsItem(Index).Caption, " ", "_") & "-" & Replace(Me.EventName, " ", "_") & ".Xls"
    
    With frmMain.CommonDialog1
        .DefaultExt = "Xls"
        .DialogTitle = Translate("Enter a file name", mcLanguage)
        .FileName = cExcelFile
        .Filter = "Excel (*.Xls)|*.Xls"
        .InitDir = mcDatabaseName
        .CancelError = True
        .ShowSave
        If Err = cdlCancel Then
            Exit Sub
        End If
        cExcelFile = .FileName
    End With
    
    SetMouseHourGlass
    If Dir$(cExcelFile) = "" Then
        iKey = vbYes
    Else
        iKey = MsgBox(Translate("Overwrite file", mcLanguage) & " '" & cExcelFile & "'?", vbQuestion + vbYesNo)
    End If
    If iKey = vbYes Then
    
        StatusMessage Translate("Exporting ", mcLanguage) & mnuFileSaveAsItem(Index).Caption
        
        KillFile cExcelFile
        
        Set xlObj = CreateObject("Excel.Sheet")
        
        xlObj.Application.Visible = False
        
        Set rstXls = mdbMain.OpenRecordset(mdbMain.QueryDefs(ReadTagItem(mnuFileSaveAsItem(Index), "Sql")).SQL)
        If rstXls.RecordCount > 0 Then
            iRow = 1
            For iColumn = 1 To rstXls.Fields.Count
               xlObj.Application.Cells(iRow, iColumn).Value = rstXls.Fields(iColumn - 1).Name
            Next iColumn
            Do While Not rstXls.EOF
                If iRow Mod 10 = 0 Then
                    StatusMessage Translate("Exporting ", mcLanguage) & mnuFileSaveAsItem(Index).Caption & " (" & iRow & ")"
                End If
                iRow = iRow + 1
                For iColumn = 1 To rstXls.Fields.Count
                    xlObj.Application.Cells(iRow, iColumn).Value = rstXls.Fields(iColumn - 1)
                Next iColumn
                rstXls.MoveNext
            Loop
        End If
        
        xlObj.SaveAs cExcelFile
        ' Close Excel with the Quit method on the Application object.
        xlObj.Application.Quit
        ' Release the object variable.
        Set xlObj = Nothing

        
        WriteIniFile gcIniFile, "Export", "Excel", cExcelFile
        iKey = MsgBox(rstXls.RecordCount & " " & Translate("items have been exported to", mcLanguage) & " '" & cExcelFile & "'." & vbCrLf & Translate("Open", mcLanguage) & " '" & cExcelFile & "'?", vbYesNo + vbQuestion)
        rstXls.Close
        If iKey = vbYes Then
            ShowDocument cExcelFile, frmMain
        End If
        StatusMessage
    Else
        MsgBox Translate("No file created.", mcLanguage)
    End If
    Set rstXls = Nothing
    
    On Local Error GoTo 0
    
    StatusMessage
    
    SetMouseNormal

End Sub

Private Sub mnuToolReset_Click()
    Me.miAlreadyHeight = 0
    Me.miParticipantLeft = 0
    Form_Resize
End Sub

Private Sub StatusBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)
   Select Case Panel.Index
   Case 1
      If fSplash.Visible = True Then
        fSplash.Hide
        StatusBar1.Panels(Panel.Index).Bevel = sbrRaised
      Else
        fSplash.Show
        StatusBar1.Panels(Panel.Index).Bevel = sbrInset
      End If
   End Select
End Sub

Public Sub tbsSection_Click(Index As Integer)
   Dim cTemp As String
   Dim iTemp As Integer
   
   On Local Error Resume Next
   
   For iTemp = 1 To tbsSection(Index).Tabs.Count
        tbsSection(Index).Tabs(iTemp).HighLighted = False
   Next iTemp
   tbsSection(Index).SelectedItem.HighLighted = miUseHighLights
   StoreCurrentMarks
   cTemp = Me.txtParticipant.Text
   Me.TestSection = tbsSection(Index).SelectedItem.Index
   If Me.TestSection = tbsSection(Index).Tabs.Count And Me.TestStatus > 0 Then
        miMaxTestSection = 1
   Else
        miMaxTestSection = 0
   End If
   
   ChangeCaption
       
   If dblstNotYet.Visible = True And dtaNotYet.Recordset.RecordCount > 0 Then
        SetFocusTo dblstNotYet
        dblstNotYet.BoundText = dtaNotYet.Recordset.Fields("cList") & ""
    Else
        SetFocusTo txtParticipant
   End If

    On Local Error GoTo 0
End Sub

Private Sub tbsSection_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub tbsSelFin_Click()
    Dim iTemp As Integer
    
    miDoNotCheckTieBreakAgain = False
   
    For iTemp = 1 To tbsSelFin.Tabs.Count
        tbsSelFin.Tabs(iTemp).HighLighted = False
    Next iTemp
    tbsSelFin.SelectedItem.HighLighted = miUseHighLights

    StoreCurrentMarks
   
    fraPrevious.Visible = True
    
    ChangeCaption True
    
    LookUpJudges
    
    For iTemp = 1 To tbsSection(tbsSelFin.SelectedItem.Index - 1).Tabs.Count
        If tbsSection(tbsSelFin.SelectedItem.Index - 1).Tabs(iTemp).Selected = True Then
            tbsSection(tbsSelFin.SelectedItem.Index - 1).Tabs(iTemp).HighLighted = miUseHighLights
        Else
            tbsSection(tbsSelFin.SelectedItem.Index - 1).Tabs(iTemp).HighLighted = False
        End If
    Next iTemp
    
    If tbsSelFin.SelectedItem.Index = 1 And Me.dtaTestInfo.Recordset.Fields("Handling") > 2 And Me.dtaTestInfo.Recordset.Fields("Handling") <> 5 Then
        Me.fraLists.Enabled = False
        Me.fraOther.Enabled = False
        Me.fraDivider2.Enabled = False
        StatusMessage Translate("No preliminary rounds in this test.", mcLanguage), 2
        Me.tbsSelFin.SelectedItem = Me.tbsSelFin.Tabs.Item(2)
        If tbsSection(1).Visible = True And tbsSection(1).Tabs.Count > 0 Then
            Me.tbsSection(1).SelectedItem = Me.tbsSection(1).Tabs.Item(1)
            tbsSelFin_Click
            tbsSection_Click 1
        End If

    Else
        Me.fraLists.Enabled = True
        Me.fraOther.Enabled = True
        Me.fraDivider2.Enabled = True
    End If
     
End Sub
' LL: Since I tend to forget: This is where the SQL queries for retrieving participants are generated after changing the current test...
Sub ChangeCaption(Optional iForced As Integer = 0)
    Dim cQry As String
    Dim iItem As Integer
    Dim cTemp As String
    Dim iTemp As Integer
    Dim cInfoMsg As String
    
    Dim rstMarks As Recordset
    
    On Local Error Resume Next
    
    miChangeCaption = True
    
    If EventName = "" Then
        EventName = GetVariable("Event_name")
        
        If GetVariable("Event_date") <> "" Then
            EventName = EventName & " - " & GetVariable("Event_date")
        End If
    End If
        
    If (Me.TestCode <> "" And iForced <> 0) Or Me.Caption <> CreateCaption Then
    
        TestInfoMessage = ""
        
        SetMouseHourGlass
   
        Me.Caption = CreateCaption
                
        'determine test status
        Select Case tbsSelFin.SelectedItem.Index
        Case 1
            TestStatus = 0
        Case tbsSelFin.Tabs.Count - 2
            TestStatus = 3
        Case tbsSelFin.Tabs.Count - 1
            TestStatus = 2
        Case tbsSelFin.Tabs.Count
            TestStatus = 1
        End Select
        
        LookUpJudges
        
        'which frame to show to enter marks/times
        fraMarks.Visible = False
        fraTime.Visible = False
        chkFlag.Visible = False
        chkNoStart.Visible = False
        TestTable = False
        mnuFilePrintFormJudges3.Enabled = False
        mnuFilePrintFormTime.Enabled = False
        cmdMarks.Enabled = False
        mnuEditNoStart.Enabled = False
        Me.mnuToolOutputJudges.Enabled = True
        
        For iTemp = 0 To 3
            If iTemp = TestStatus Then
                cmbNumJudges(iTemp).Visible = True
            Else
                cmbNumJudges(iTemp).Visible = False
            End If
            cmbNumJudges(iTemp).Enabled = True
        Next iTemp
        If TestStatus = 0 Then
            Select Case dtaTest.Recordset.Fields("Type_Pre")
            Case Is <= 2 'marks or placemarks
                fraMarks.Visible = True
                If dtaTest.Recordset.Fields("Type_Time") = 3 Then  'PP1
                    cmbNumJudges(TestStatus).Text = "5"
                    cmbNumJudges(TestStatus).Enabled = False
                    Me.lblMarks(4).Caption = TranslateCaption("&Time", lblMarks(4).Width, True) & ":"
                    chkFlag.Visible = True
                    If TestMarkDecimals = 1 Then
                        cInfoMsg = Translate("Enter marks and time with 1 decimal.", mcLanguage)
                    Else
                        cInfoMsg = Translate("Enter marks and time with # decimals.", mcLanguage)
                    End If
                ElseIf dtaTest.Recordset.Fields("Type_Pre") = 2 Then
                    cInfoMsg = Translate("Enter place marks only!", mcLanguage)
                Else
                    If TestMarkDecimals = 1 Then
                        cInfoMsg = Translate("Enter marks with 1 decimal.", mcLanguage)
                    Else
                        cInfoMsg = Translate("Enter marks with # decimals.", mcLanguage)
                    End If
                End If
                TestInfoMessage = Replace(cInfoMsg, "#", TestMarkDecimals)
            Case Is = 3  'time
                If dtaTest.Recordset.Fields("Type_Time") > 0 Then
                    TestTable = True
                End If
                fraTime.Visible = True
                chkFlag.Visible = True
                chkNoStart.Visible = True
                mnuEditNoStart.Enabled = True
                If TestTimeDecimals = 1 Then
                    cInfoMsg = Translate("Enter times in seconds with 1 decimal.", mcLanguage)
                Else
                    cInfoMsg = Translate("Enter times in seconds with # decimals.", mcLanguage)
                End If
                TestInfoMessage = Replace(cInfoMsg, "#", TestTimeDecimals)
                Me.mnuToolOutputJudges.Enabled = False
            Case Else
            End Select
        Else
            If dtaTest.Recordset.Fields("Type_Final") <= 2 Then
                fraMarks.Visible = True
                If dtaTest.Recordset.Fields("Type_Final") = 2 Then
                    cInfoMsg = Translate("Enter place marks only!", mcLanguage)
                Else
                    If TestMarkDecimals = 1 Then
                        cInfoMsg = Translate("Enter marks with 1 decimal.", mcLanguage)
                    Else
                        cInfoMsg = Translate("Enter marks with # decimals.", mcLanguage)
                    End If
                End If
                TestInfoMessage = Replace(cInfoMsg, "#", TestMarkDecimals)
            End If
        End If
        
        'determine queries for lists
        For iItem = 0 To Me.tbsSection.Count - 1
            If iItem = tbsSelFin.SelectedItem.Index - 1 Then
                Select Case TestStatus
                Case 0
                   Me.fraFinals.Visible = False
                   Me.chkRein.Visible = True
                   If fraTime.Visible = False Then
                        If dtaTest.Recordset.Fields("Type_special") <> 3 Then
                            Me.mnuFilePrintFormJudges3.Enabled = True
                        End If
                   Else
                        Me.mnuFilePrintFormTime.Enabled = True
                   End If
                Case Is > 0
                   Me.fraFinals.Visible = True
                   Me.chkRein.Visible = False
                End Select
                
                Me.fraGroups.Visible = Not Me.fraFinals.Visible
                Me.chkColor.Visible = True
                Me.fraGroups.Enabled = Me.chkColor.Visible
                Me.TestSection = tbsSection(iItem).SelectedItem.Index
                
                'how to format marks
                If ((TestStatus = 0 And dtaTest.Recordset.Fields("Type_pre") = 2) Or (TestStatus <> 0 And dtaTest.Recordset.Fields("Type_Final") = 2)) Then
                    TestMarkFormat = "0"
                Else
                    TestMarkFormat = "0." & String$(TestMarkDecimals, "0")
                End If
                TestTimeFormat = "0." & String$(TestTimeDecimals, "0")
                
                If dtaTest.Recordset.Fields("Type_Special") = 3 Then 'gaedingakeppni
                    TestTotalFormat = "0.000"
                Else
                    TestTotalFormat = "0.00"
                End If
                
                'select entered riders that haven't started first
                tbsSection.Item(iItem).Visible = True
                cQry = "SELECT Entries.Sta "
                If chkColor.Value <> 0 Then
                    If miUseColors = 1 Then
                        cQry = cQry & " & ' - ' & LCASE(LEFT(Entries.Color & '   ',2)) "
                    End If
                    If TestStatus = 0 Then
                        cQry = cQry & " & ' ' & IIF(Entries.Nostart=0 OR ISNULL(Entries.Nostart),IIF(Entries.group>0,FORMAT(Entries.group,'00'),''),'---') "
                    End If
                End If
                If TestStatus = 0 Then
                    If chkRein.Value <> 0 And chkRein.Enabled = True Then
                        cQry = cQry & " & ' ' & IIF(Entries.RR=TRUE,'" & Left$(Translate("Right", mcLanguage), 1) & "','" & Left$(Translate("Left", mcLanguage), 1) & "')"
                    End If
                End If
                cQry = cQry & " & ' - ' & Persons.Name_First"
                cQry = cQry & " & ' ' & Persons.Name_Last"
                cQry = cQry & " & IIF(Participants.Class<>'',' [' & Participants.Class & ']','')"
                If frmMain.chkTeam.Value = 1 Then
                    cQry = cQry & " & IIF(Participants.Team<>'',' [' & Participants.Team & ']','')"
                    cQry = cQry & " & IIF(Participants.Club<>'',' [' & Participants.Club & ']','')"
                End If
                cQry = cQry & " & ' - ' & Horses.Name_Horse"
                If frmMain.chkFeifId.Value = 1 Then
                    cQry = cQry & " & IIF(Horses.FEIFId<>'',' [' & Horses.FEIFId & ']','')"
                End If
                If miShowHorseAge <> 0 Then
                    cQry = cQry & " & IIF(Horses.FEIFId<>'', IIF(LEN(Horses.FEIFID)>8 and VAL(MID(Horses.FEIFId,3,4))>1900 AND VAL(MID(Horses.FEIFId,3,4))<2030,IIF(DATEDIFF('y',YEAR(CDATE(MID(Horses.FEIFId,3,4) & '-01-01')),year(date()))< " & Format$(miHorseAgeLimit) & ",' **' & DATEDIFF('y',YEAR(CDATE(MID(Horses.FEIFId,3,4) & '-01-01')),year(date())) & '**' ,''),''),'') "
                End If
                If TestStatus = 0 Then
                    cQry = cQry & " & ' ' & IIF(Entries.Check>0,' [" & Left$(Translate("Equipment check", mcLanguage), 1) & "]','') "
                    cQry = cQry & " & ' ' & IIF(Entries.NoStart=-1,' [" & (Translate("No start", mcLanguage)) & "]',' ') "
                End If
                cQry = cQry & " as cList"
                cQry = cQry & ",Entries.Sta"
                cQry = cQry & ",Entries.Position"
                cQry = cQry & ",Entries.Group"
                cQry = cQry & ",Entries.Color"
                cQry = cQry & ",Entries.RR"
                cQry = cQry & ",Entries.NoStart"
                cQry = cQry & ",Entries.Check"
                cQry = cQry & ",Entries.Deleted"
                cQry = cQry & ",Persons.Name_First"
                cQry = cQry & ",Persons.Name_Last"
                cQry = cQry & ",Horses.Name_Horse"
                If miShowHorseId <> 0 Then
                    cQry = cQry & ",Horses.HorseID"
                    cQry = cQry & ",Horses.FEIFID"
                End If
                If miShowRidersClub <> 0 Then
                    cQry = cQry & ",Participants.Club"
                End If
                If miShowRidersTeam <> 0 Then
                    cQry = cQry & ",Participants.Team"
                End If
                cQry = cQry & ",Participants.Status AS pStatus"
                cQry = cQry & ",Participants.Class"
                cQry = cQry & " FROM (((Entries "
                cQry = cQry & " INNER JOIN Participants"
                cQry = cQry & " ON Entries.STA = Participants.STA) "
                cQry = cQry & " INNER JOIN Persons "
                cQry = cQry & " ON Participants.PersonID = Persons.PersonID) "
                cQry = cQry & " INNER JOIN Horses "
                cQry = cQry & " ON Participants.HorseID = Horses.HorseID) "
                cQry = cQry & " LEFT JOIN Results "
                cQry = cQry & " ON (Entries.STA = Results.STA) "
                cQry = cQry & " AND (Entries.Code = Results.Code) "
                cQry = cQry & " AND (Entries.Status = Results.Status)"
                cQry = cQry & " WHERE (((Entries.STA) Not In "
                cQry = cQry & " (SELECT Marks.Sta FROM Marks  "
                cQry = cQry & " Where Marks.Status = " & Me.TestStatus
                cQry = cQry & " AND Marks.Section=" & Me.TestSection
                cQry = cQry & " AND Marks.Code='" & Me.TestCode & "') "
                
                'don't show participants who withdrew completely or got eliminated:
                'but include participants with a null status !
                cQry = cQry & " AND (Participants.Status < 2 OR ISNULL(Participants.Status))"
                
                cQry = cQry & " And (Entries.STA) Not In "
                cQry = cQry & " (SELECT Results.Sta "
                cQry = cQry & " FROM Results  "
                cQry = cQry & " WHERE Results.Sta=Entries.Sta  "
                cQry = cQry & " AND Results.Code='" & Me.TestCode & "' "
                cQry = cQry & " AND (Results.Disq<0 OR Entries.Deleted<0))) "
                cQry = cQry & " AND ((Entries.Code)='" & Me.TestCode & "')"
                cQry = cQry & " AND ((Entries.Status)=" & Me.TestStatus & ")"
                cQry = cQry & " AND ((Entries.Deleted)<>-1)"
                cQry = cQry & " AND ((Entries.Deleted)<>1))"
                
                If fraMarks.Visible = True Then
                    If TestStatus > 0 And miFinalsSequence <> 0 Then
                        cQry = cQry & " ORDER BY Entries.Group, Entries.Position DESC;"
                    Else
                        cQry = cQry & " ORDER BY Entries.Group, Entries.Position;"
                    End If
                Else
                    cQry = cQry & " ORDER BY (ISNULL(Entries.Nostart) OR NoStart=0), ISNULL(Results.AllTimes), Results.AllTimes DESC, Entries.Group, Entries.Position;"
                End If

                dtaNotYet.RecordSource = cQry
                dtaNotYet.Refresh
                
                fraNotYet.Caption = Translate("Participants not yet started", mcLanguage) & " [" & dtaNotYet.Recordset.RecordCount & "]"
                dblstNotYet.ListField = "cList"
                
                dtaAlready.RecordSource = StartedParticipants(Me.TestStatus)
                dtaAlready.Refresh
                fraAlready.Caption = Translate("Participants already started", mcLanguage) & " [" & dtaAlready.Recordset.RecordCount & "]"
                
                cmdComposeFinals.Enabled = True
                If Me.dtaTestInfo.Recordset.Fields("Handling") > 4 Then
                    chkSplitFinals.Enabled = False
                    chkSplitResultLists.Enabled = False
                Else
                    chkSplitFinals.Enabled = True
                    chkSplitResultLists.Enabled = True
                End If
                
            Else
                tbsSection.Item(iItem).Visible = False
            End If
        Next iItem
        
        StatusMessage
        
        fraJudges.Visible = fraMarks.Visible
        
        'add label to score field
        lblScore.Caption = FitString(Me, " " & ClipAmp(tbsSection(Me.TestStatus).Tabs(Me.TestSection).Caption) & ":", Me.fraResults.Width / 2 - txtScore.Width - 100, 1)
        txtParticipant_Change
       
        cQry = "SELECT * FROM TestSections "
        cQry = cQry & " WHERE Code='" & Me.TestCode & "'"
        cQry = cQry & " AND Section=" & Me.TestSection
        cQry = cQry & " AND Status=" & IIf(Me.TestStatus = 3, 1, IIf(Me.TestStatus = 2, 1, Me.TestStatus))
        dtaTestSection.RecordSource = cQry
        dtaTestSection.Refresh
       
        'check if there are already any valid marks
        
        Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Code Like '" & TestCode & "' AND Status=0")
        If rstMarks.RecordCount = 0 And dtaTest.Recordset.Fields("Type_Time") <> 3 Then
            cmbNumJudges(TestStatus).Enabled = True
            Form_Resize
        Else
            cmbNumJudges(TestStatus).Enabled = False
            If TestStatus <> 0 Then
                Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Code Like '" & TestCode & "' AND Status<>0")
                If rstMarks.RecordCount = 0 And dtaTest.Recordset.Fields("Type_Time") <> 3 And dtaTestInfo.Recordset.Fields("num_j_0") < 5 Then
                    cmbNumJudges(TestStatus).Enabled = True
                    Form_Resize
                End If
            End If
        End If
        
        rstMarks.Close
        Set rstMarks = Nothing
   
        LookUpRelevantParticipants
       
        If dtaTest.Recordset.Fields("RR") = True Then
            Me.chkRein.Enabled = True
        Else
            Me.chkRein.Enabled = False
        End If
        
        txtPrevious.Text = ""
        txtPrevious.Tag = ""
        txtPrevious.BackColor = vbWhite
        
        SetMouseNormal
    End If
    
    miChangeCaption = False
    
    On Local Error GoTo 0

End Sub

Private Sub tbsSelFin_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub Timer1_Timer()
   LookUpParticipant
End Sub

Private Sub Timer2_Timer()
    If miBackupInterval > 0 Then
        miBackupTicker = (miBackupTicker + 1) Mod miBackupInterval
        If miBackupTicker = 0 Then
            If miNoBackupNow = True Then
                miBackupTicker = miBackupInterval - 1
                StatusMessage Translate("Backup pending ...", mcLanguage)
            Else
                CreateBackup mdbMain, GetVariable("Backup")
                StatusMessage
            End If
        End If
    End If
End Sub

Private Sub txtAlready_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub txtMarks_KeyPress(Index As Integer, KeyAscii As Integer)
    'suppress annoying beep when using <Enter> in stead of <Tab>
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtMove_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub txtParticipant_Change()
   Timer1.Enabled = False
   If Trim$(txtParticipant.Text) = "" Then
      Me.fraMarks.Enabled = False
      Me.fraResults.Enabled = False
      Me.cmdInfo.Enabled = False
      Me.fraResults.Enabled = False
   Else
      Me.fraMarks.Enabled = True
      Me.fraResults.Enabled = True
      Me.cmdInfo.Enabled = True
      Me.fraResults.Enabled = True
      Me.txtPrevious.BackColor = vbWhite
      Timer1.Enabled = True
   End If
   Me.lblParticipant.Caption = ""
End Sub

Private Sub txtParticipant_DblClick()
    If txtParticipant.Text <> "" Then
       LookUpParticipant
       If fraMarks.Visible = True Then
            SetFocusTo txtMarks(0)
        ElseIf fraTime.Visible = True Then
            SetFocusTo txtTime
        End If
        
    End If
End Sub

Private Sub txtParticipant_GotFocus()
    txtParticipant.BackColor = mlAlertColor
    mnuEditMove.Enabled = True
    Me.mnuEditChangeRein.Enabled = Me.chkRein.Enabled
End Sub

Private Sub txtParticipant_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn
      If txtParticipant.Text <> "" Then
         LookUpParticipant
         If fraMarks.Visible = True Then
             SetFocusTo txtMarks(0)
         ElseIf fraTime.Visible = True Then
             SetFocusTo txtTime
         End If
      End If
   End Select
End Sub

Private Sub txtParticipant_KeyPress(KeyAscii As Integer)
    'suppress annoying beep when using <Enter> in stead of <Tab>
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtParticipant_LostFocus()
    Set mctlActive = txtParticipant
    txtParticipant.BackColor = QBColor(15)
    miNoBackupNow = False
End Sub

Private Sub txtParticipant_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If

End Sub

Private Sub txtParticipant_Validate(Cancel As Boolean)
   LookUpParticipant
End Sub
Public Sub LookUpParticipant()
   Dim cTemp As String
   Dim iTemp As Integer
   Dim cSta As String
   Dim cSex As String
   Dim iKey As Integer
   
   Timer1.Enabled = False
   txtParticipant.Text = Trim$(txtParticipant.Text)
   If txtParticipant.Text <> "" Then
      cSta = Format$(Val(txtParticipant.Text), "000")
      dtaParticipant.RecordSource = "SELECT Persons.*, Horses.*, Participants.Class, Participants.Club, Participants.Team FROM Horses INNER JOIN (Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) ON Horses.HorseID = Participants.HorseID WHERE Participants.STA='" & cSta & "'"
      dtaParticipant.Refresh
      If dtaParticipant.Recordset.RecordCount > 0 Then
         Timer1.Enabled = False
         Select Case dtaParticipant.Recordset.Fields("Sex_Horse")
         Case 1
            cSex = Translate("Stallion", mcLanguage)
         Case 2
            cSex = Translate("Mare", mcLanguage)
         Case 3
            cSex = Translate("Gelding", mcLanguage)
         Case Else
            cSex = "--"
         End Select
         cTemp = dtaParticipant.Recordset.Fields("Name_First") & " " & dtaParticipant.Recordset.Fields("Name_Last")
         
         cTemp = cTemp & " - " & dtaParticipant.Recordset.Fields("Name_horse")
         If dtaParticipant.Recordset.Fields("Class") & "" <> "" Then
            cTemp = cTemp & " [" & dtaParticipant.Recordset.Fields("Class") & "]"
         End If
         If miShowRidersClub <> 0 And dtaParticipant.Recordset.Fields("Club") & "" <> "" Then
            cTemp = cTemp & " [" & dtaParticipant.Recordset.Fields("Club") & "]"
         End If
         If miShowRidersTeam <> 0 And dtaParticipant.Recordset.Fields("Team") & "" <> "" Then
            cTemp = cTemp & " [" & dtaParticipant.Recordset.Fields("Team") & "]"
         End If
         cTemp = cTemp & vbCrLf & cSex & ", " & Format$(dtaParticipant.Recordset.Fields("Birthday_Horse"), "yyyy") & ", " & dtaParticipant.Recordset.Fields("Country") & ", " & dtaParticipant.Recordset.Fields("Color")
         
         lblParticipant.Caption = cTemp
         
         LookUpMarks
         
         cTemp = cSta & ": " & cTemp
         
         If miExcelFiles <> 0 Then
            CreateExcelParticipant
         End If
         
         SetFocusTo txtParticipant
      Else
         lblParticipant = cTemp
         iKey = MsgBox(Translate("Participant not found; add participant on a temporary basis?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
         If iKey = vbYes Then
            dtaParticipant.RecordSource = "SELECT Persons.*, Horses.*, Participants.Class, Participants.Club, Participants.Team FROM Horses INNER JOIN (Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) ON Horses.HorseID = Participants.HorseID"
            dtaParticipant.Refresh
            If dtaParticipant.Recordset.RecordCount > 0 Then
                frmParticipant.FindAddParticipant Format$(Val(txtParticipant), "000")
                LookUpParticipant
            Else
                MsgBox (Translate("No participants found at all. Enter the participants first using the relevant function in the menu.", mcLanguage))
            End If
         Else
            txtParticipant.Text = ""
         End If
      End If
   End If
   txtParticipant.SelStart = Len(txtParticipant.Text)
   
   StatusMessage
   
End Sub

Private Sub LookUpMarks(Optional iClear As Integer = False)
   Dim iItem As Integer
   Dim cQry As String
   Dim rstMarks As Recordset
   Dim rstDisq As Recordset
   Dim rstNoStart As Recordset
   
   txtScore.Text = ""
   
   If iClear = True Then
        ClearMarks True
   Else
        'is there already a result record?
        cQry = "SELECT * FROM Marks "
        cQry = cQry & " WHERE Sta='" & Format$(Val(txtParticipant.Text), "000") & "'"
        cQry = cQry & " AND Status=" & Me.TestStatus
        cQry = cQry & " AND Section=" & Me.TestSection
        cQry = cQry & " AND Code='" & Me.TestCode & "'"
        Set rstMarks = mdbMain.OpenRecordset(cQry)
        
        If rstMarks.RecordCount > 0 Then
            If fraJudges.Visible = True Then
               For iItem = 0 To TestJudges - 1
                  If rstMarks.Fields("Mark" & Format$(iItem + 1)) & "" <> "" Then
                     txtMarks(iItem).Text = Format$(rstMarks.Fields("Mark" & Format$(iItem + 1)), TestMarkFormat)
                     ValidateMark iItem
                  Else
                     txtMarks(iItem).Text = ""
                  End If
               Next iItem
            ElseIf fraTime.Visible = True Then
               txtTime = Format$(rstMarks.Fields("Mark1"), TestTimeFormat)
            End If
            chkFlag.Value = IIf(rstMarks.Fields("Flag") = True, 1, 0)
        ElseIf TestSection <> 0 And dtaTestSection.Recordset.Fields("Recycle") <> 0 Then
            'are there results from prelim?
            cQry = "SELECT * FROM Marks "
            cQry = cQry & " WHERE Sta='" & Format$(Val(txtParticipant.Text), "000") & "'"
            cQry = cQry & " AND Status=0"
            cQry = cQry & " AND Section=" & Me.TestSection
            cQry = cQry & " AND Code='" & Me.TestCode & "'"
            Set rstMarks = mdbMain.OpenRecordset(cQry)
            
            If rstMarks.RecordCount > 0 Then
                If fraJudges.Visible = True Then
                   For iItem = 0 To TestJudges - 1
                      If rstMarks.Fields("Mark" & Format$(iItem + 1)) & "" <> "" Then
                         txtMarks(iItem).Text = Format$(rstMarks.Fields("Mark" & Format$(iItem + 1)), TestMarkFormat)
                         ValidateMark iItem
                      Else
                         txtMarks(iItem).Text = ""
                      End If
                   Next iItem
                ElseIf fraTime.Visible = True Then
                   txtTime = Format$(rstMarks.Fields("Mark1"), TestTimeFormat)
                End If
                chkFlag.Value = IIf(rstMarks.Fields("Flag") = True, 1, 0)
            End If
        Else
            ClearMarks True
        End If
        rstMarks.Close
        Set rstMarks = Nothing
        
        Set rstDisq = mdbMain.OpenRecordset("SELECT Disq FROM Results WHERE Sta='" & Format$(Val(txtParticipant.Text), "000") & "' AND Status=" & Me.TestStatus & " AND Code='" & Me.TestCode & "'")
        With rstDisq
           If .RecordCount > 0 Then
              chkDisqualified.Value = IIf(.Fields("Disq") = -1, 1, 0)
              chkWithdrawn.Value = IIf(.Fields("Disq") = -2, 1, 0)
              If chkDisqualified.Value = 1 Then
                  chkWithdrawn.Value = 0
                  chkWithdrawn.Enabled = False
              End If
           End If
        End With
        rstDisq.Close
        
        If chkDisqualified.Value = 1 Or chkWithdrawn.Value = 1 Then
            chkNoStart.Enabled = False
            chkFlag.Enabled = False
        Else
            chkNoStart.Enabled = True
            Set rstNoStart = mdbMain.OpenRecordset("SELECT NoStart FROM Entries WHERE Sta='" & Format$(Val(txtParticipant.Text), "000") & "' AND Status=" & Me.TestStatus & " AND Code='" & Me.TestCode & "'")
            With rstNoStart
               If .RecordCount > 0 Then
                  chkNoStart.Value = IIf(.Fields("Nostart") = -1, 1, 0)
               End If
            End With
            rstNoStart.Close
            Set rstNoStart = Nothing
        End If
        
      lblParticipant.Tag = ""
      If fraJudges.Visible = True Then
            ValidateScore
      ElseIf fraTime.Visible = True Then
            ValidateTimeScore
      End If
        mnuEditNoStart.Enabled = chkNoStart.Visible

   End If
End Sub

Private Sub txtMarks_Change(Index As Integer)
   lblParticipant.Tag = "*"
   If txtMarks(Index).BackColor <> mlAlertColor And txtMarks(Index).Text = "" Then
       txtMarks(Index).BackColor = mlAlertColor
       miNoBackupNow = True
   End If
End Sub

Private Sub txtMarks_Click(Index As Integer)
    If txtMarks(Index).BackColor <> mlAlertColor Then
       txtMarks(Index).BackColor = mlAlertColor
       miNoBackupNow = True
    End If
End Sub

Private Sub txtMarks_GotFocus(Index As Integer)
   txtMarks(Index).SelStart = 0
   txtMarks(Index).SelLength = Len(txtMarks(Index).Text)
   If txtMarks(Index).BackColor <> mlAlertColor Then
        txtMarks(Index).BackColor = mlAlertColor
        miNoBackupNow = True
    End If
End Sub

Private Sub txtMarks_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyDown, vbKeyReturn
      If ValidateMark(Index) = True Then
         If Index < TestJudges - 1 Then
            If txtMarks(Index + 1).BackColor <> mlAlertColor Then
               txtMarks(Index + 1).BackColor = mlAlertColor
               miNoBackupNow = True
            End If
            SetFocusTo txtMarks(Index + 1)
         ElseIf KeyCode = vbKeyReturn Then
            SetFocusTo cmdOK
         End If
      End If
      KeyCode = 0
   Case vbKeyUp
      If ValidateMark(Index) = True Then
         If Index > 0 Then
            If txtMarks(Index - 1).BackColor <> mlAlertColor Then
               txtMarks(Index - 1).BackColor = mlAlertColor
               miNoBackupNow = True
            End If
            SetFocusTo txtMarks(Index - 1)
         End If
      End If
      KeyCode = 0
   End Select
End Sub

Private Sub txtMarks_LostFocus(Index As Integer)
    Set mctlActive = txtMarks(Index)
    txtMarks(Index).BackColor = QBColor(15)
    miNoBackupNow = False
End Sub

Private Sub txtMarks_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub txtMarks_Validate(Index As Integer, Cancel As Boolean)
   ValidateMark Index
End Sub
Public Sub LookUpTest()
   Dim iTabIndex As Integer
   Dim iTabMinWidth As Integer
   Dim cAmpList() As String
   Dim cTemp As String
   
   ReDim cAmpList(0 To 2)
   Dim iKey As Integer
   
   SetMouseHourGlass
           
   DoEvents
   
   txtAlready.Text = Translate("Wait", mcLanguage) & "..."
   txtAlready.Visible = True
   lstAlready.Visible = False
   miDoNotCheckTieBreakAgain = False
   
   'read test data
   dtaTest.RecordSource = "SELECT * FROM Tests WHERE Code LIKE '" & Me.TestCode & "'"
   dtaTest.Refresh
   If dtaTest.Recordset.RecordCount = 0 Then
       dtaTest.RecordSource = "SELECT * FROM Tests"
       dtaTest.Refresh
       If dtaTest.Recordset.RecordCount = 0 Then
           MsgBox Translate("No valid list of tests available. Download Sport Rules first. ", mcLanguage), vbCritical
           Unload Me
           End
       End If
       Me.TestCode = dtaTest.Recordset.Fields("Code") & ""
       Me.TestName = dtaTest.Recordset.Fields("Test") & ""
   Else
       Me.TestName = dtaTest.Recordset.Fields("Test") & ""
   End If
      
    'check for additional info
    CreateTestInfoAll
    With dtaTestInfo
        .RecordSource = "SELECT * FROM TestInfo WHERE Code='" & Me.TestCode & "'"
        .Refresh
    End With
        
    Me.TestMarkDecimals = 1
    Me.TestTimeDecimals = 1
    Select Case dtaTest.Recordset.Fields("Type_Pre")
    Case Is <= 2 'marks or placemarks
        fraMarks.Visible = True
        With dtaTest.Recordset
            If IsNull(.Fields("Mark_Decimals")) Then
                Me.TestMarkDecimals = 1
            ElseIf .Fields("Mark_Decimals") = 0 And dtaTest.Recordset.Fields("Type_Pre") = 1 Then
                Me.TestMarkDecimals = 1
            Else
                Me.TestMarkDecimals = .Fields("Mark_Decimals")
            End If
        End With
    Case Is = 3  'time
        fraTime.Visible = True
        With dtaTest.Recordset
            If IsNull(.Fields("Time_Decimals")) Or .Fields("Time_Decimals") = 0 Then
                Me.TestTimeDecimals = 1
            Else
                Me.TestTimeDecimals = .Fields("Time_Decimals")
            End If
        End With
    Case Else
    End Select
    chkFlag.Visible = fraTime.Visible
    chkNoStart.Visible = fraTime.Visible
    
   
   If dtaTest.Recordset.Fields("Status").Value & "" = "" Then
        dtaTest.Recordset.Edit
        dtaTest.Recordset.Fields("Status").Value = 0
        dtaTest.Recordset.Update
   End If
   If dtaTest.Recordset.Fields("Type_Time").Value & "" = "" Then
        dtaTest.Recordset.Edit
        dtaTest.Recordset.Fields("Type_Time").Value = 0
        dtaTest.Recordset.Update
   End If
   
   If dtaTest.Recordset.Fields("Type_Time") > 0 Then
        Me.fraFinals.Enabled = False
    Else
        Me.fraFinals.Enabled = True
   End If
   
   If dtaTestInfo.Recordset.Fields("Handling").Value & "" = "" Then
        dtaTestInfo.Recordset.Edit
        If dtaTest.Recordset.Fields("Type_Time").Value = 0 Then
            dtaTestInfo.Recordset.Fields("Handling").Value = 2
        Else
            dtaTestInfo.Recordset.Fields("Handling").Value = 0
        End If
        dtaTestInfo.Recordset.Update
   End If
   
   cmdOK.TabIndex = iTabIndex

   'read the sections
   dtaTestSection.RecordSource = "SELECT * FROM TestSections WHERE Code LIKE '" & Me.TestCode & "' ORDER BY Status,Section"
   dtaTestSection.Refresh
   If dtaTestSection.Recordset.RecordCount = 0 Then
        MsgBox Me.TestName & ": " & Translate("no valid test data found; " & App.EXEName & " will be aborted.", mcLanguage) & ".", vbCritical
        Me.TestCode = ""
        Unload Me
        End
   Else
        Me.tbsSelFin.Tabs.Clear
        Me.tbsSection(0).Tabs.Clear
        Me.tbsSection(1).Tabs.Clear
        Me.tbsSection(2).Tabs.Clear
        Me.tbsSection(3).Tabs.Clear
        Do While Not dtaTestSection.Recordset.EOF
            If dtaTestSection.Recordset.Fields("Status") = 0 Then
                Me.tbsSection(0).TabMinWidth = iTabMinWidth
                If Me.tbsSelFin.Tabs.Count = 0 Then
                   Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&" & Me.TestName, 0, True)
                End If
                Me.tbsSection(0).Tabs.Add pvcaption:=Me.TestName
                If dtaTestSection.Recordset.Fields("Name") <> "" Then
                    Me.tbsSection(0).Tabs.Item(Me.tbsSection(0).Tabs.Count).Caption = Translate(dtaTestSection.Recordset.Fields("Name"), mcLanguage)
                Else
                    Me.tbsSection(0).Tabs.Item(Me.tbsSection(0).Tabs.Count).Caption = Me.TestName
                End If
            ElseIf dtaTestSection.Recordset.Fields("Status") = 1 Then  'Finals
                If Me.tbsSelFin.Tabs.Count = 0 Then
                   Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&Preliminary Round", 0, True)
                Else
                   Me.tbsSelFin.Tabs(1).Caption = TranslateCaption("&Preliminary Round", 0, True)
                End If
                
                If dtaTestInfo.Recordset.Fields("Handling").Value > 0 Then
                    Me.tbsSection(1).Tabs.Add
                    Me.tbsSection(1).Tabs.Item(Me.tbsSection(1).Tabs.Count).Caption = Me.tbsSection(1).Tabs.Count & ". " & Translate(dtaTestSection.Recordset.Fields("Name"), mcLanguage)
                    Me.tbsSection(1).TabMinWidth = iTabMinWidth
                    Select Case dtaTestInfo.Recordset.Fields("Handling").Value
                    Case 1, 3 '1 prelim+A-final, 3 A-final only
                        If Me.tbsSelFin.Tabs.Count = 0 Then
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&Preliminary Round", 0, True)
                        Else
                           Me.tbsSelFin.Tabs(1).Caption = TranslateCaption("&Preliminary Round", 0, True)
                        End If
                        If tbsSelFin.Tabs.Count = 1 Then
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&A-Final", 0, True)
                        Else
                           Me.tbsSelFin.Tabs(2).Caption = TranslateCaption("&A-Final", 0, True)
                        End If
                        '* check if finals should be split
                        '*
                        If dtaTestInfo.Recordset.Fields("SplitFinals") = 1 Then
                            Me.chkSplitFinals.Value = 1
                            Me.chkSplitResultLists.Value = 1
                        Else
                            Me.chkSplitFinals.Value = 0
                            Me.chkSplitResultLists.Value = 0
                        End If
                        If dtaTestInfo.Recordset.Fields("Handling").Value = 3 Then
                        End If
                    Case 2, 4 'prelim+A+B, 4 without prelim
                        
                        If Me.tbsSelFin.Tabs.Count = 0 Then
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&Preliminary Round", 0, True)
                        Else
                           Me.tbsSelFin.Tabs(1).Caption = TranslateCaption("&Preliminary Round", 0, True)
                        End If
                        If tbsSelFin.Tabs.Count = 1 Then
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&B-Final", 0, True)
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&A-Final", 0, True)
                        ElseIf tbsSelFin.Tabs.Count = 2 Then
                           Me.tbsSelFin.Tabs(2).Caption = TranslateCaption("&B-Final", 0, True)
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&A-Final", 0, True)
                        End If
                        '* check if finals should be split
                        '*
                        If dtaTestInfo.Recordset.Fields("SplitFinals") = 1 Then
                            Me.chkSplitFinals.Value = 1
                            Me.chkSplitResultLists.Value = 1
                        Else
                            Me.chkSplitFinals.Value = 0
                            Me.chkSplitResultLists.Value = 0
                        End If
                        '--- Added by MM 8-3-2017
                    Case 5, 6 'prelim+A+B+C, 6 without prelim
                        If Me.tbsSelFin.Tabs.Count = 0 Then
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&Preliminary Round", 0, True)
                        Else
                           Me.tbsSelFin.Tabs(1).Caption = TranslateCaption("&Preliminary Round", 0, True)
                        End If
                        If tbsSelFin.Tabs.Count = 1 Then
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&C-Final", 0, True)
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&B-Final", 0, True)
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&A-Final", 0, True)
                        ElseIf tbsSelFin.Tabs.Count = 2 Then
                           Me.tbsSelFin.Tabs(2).Caption = TranslateCaption("&C-Final", 0, True)
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&B-Final", 0, True)
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&A-Final", 0, True)
                        ElseIf tbsSelFin.Tabs.Count = 3 Then
                           Me.tbsSelFin.Tabs(2).Caption = TranslateCaption("&C-Final", 0, True)
                           Me.tbsSelFin.Tabs(3).Caption = TranslateCaption("&B-Final", 0, True)
                           Me.tbsSelFin.Tabs.Add pvcaption:=TranslateCaption("&A-Final", 0, True)
                        End If
                        '* check if finals should be split: would be ludacrous, no way
                        '*
                        Me.chkSplitFinals.Value = 0
                        Me.chkSplitResultLists.Value = 0
                    End Select
                    If dtaTestInfo.Recordset.Fields("Handling").Value = 2 Or dtaTestInfo.Recordset.Fields("Handling").Value = 4 Then
                        Me.tbsSection(Me.tbsSelFin.Tabs.Count - 1).Tabs.Add
                        Me.tbsSection(Me.tbsSelFin.Tabs.Count - 1).Tabs.Item(Me.tbsSection(2).Tabs.Count).Caption = Me.tbsSection(1).Tabs.Count & ". " & Translate(dtaTestSection.Recordset.Fields("Name"), mcLanguage)
                        Me.tbsSection(Me.tbsSelFin.Tabs.Count - 1).TabMinWidth = iTabMinWidth
                    End If
                    If dtaTestInfo.Recordset.Fields("Handling").Value = 5 Or dtaTestInfo.Recordset.Fields("Handling").Value = 6 Then
                        Me.tbsSection(Me.tbsSelFin.Tabs.Count - 2).Tabs.Add
                        Me.tbsSection(Me.tbsSelFin.Tabs.Count - 2).Tabs.Item(Me.tbsSection(2).Tabs.Count).Caption = Me.tbsSection(1).Tabs.Count & ". " & Translate(dtaTestSection.Recordset.Fields("Name"), mcLanguage)
                        Me.tbsSection(Me.tbsSelFin.Tabs.Count - 2).TabMinWidth = iTabMinWidth
                        Me.tbsSection(Me.tbsSelFin.Tabs.Count - 1).Tabs.Add
                        Me.tbsSection(Me.tbsSelFin.Tabs.Count - 1).Tabs.Item(Me.tbsSection(3).Tabs.Count).Caption = Me.tbsSection(1).Tabs.Count & ". " & Translate(dtaTestSection.Recordset.Fields("Name"), mcLanguage)
                        Me.tbsSection(Me.tbsSelFin.Tabs.Count - 1).TabMinWidth = iTabMinWidth
                    End If
                End If
            End If
            dtaTestSection.Recordset.MoveNext
        Loop
   End If
   
   ChangeCaption True
          
   LookUpJudges
    
    'LL: This might cause problems when starting list has been pre-sorted with IceSort!
    '
   If IsNull(dtaTest.Recordset.Fields("Groupsize")) And dtaNotYet.Recordset.RecordCount > 0 Then
        cmdComposeGroups_Click
   End If
   Form_Resize
   
   If dtaTestInfo.Recordset.Fields("Handling").Value = 3 Or dtaTestInfo.Recordset.Fields("Handling").Value = 4 Or dtaTestInfo.Recordset.Fields("Handling").Value = 6 Then
      Me.tbsSelFin.Tabs(1).Caption = ""
      Me.tbsSelFin.SelectedItem = Me.tbsSelFin.Tabs.Item(2)
      If tbsSection(1).Visible = True And tbsSection(1).Tabs.Count > 0 Then
         Me.tbsSection(1).SelectedItem = Me.tbsSection(1).Tabs.Item(1)
         tbsSelFin_Click
         tbsSection_Click 1
      End If
   Else
      Me.tbsSelFin.SelectedItem = Me.tbsSelFin.Tabs.Item(1)
      If tbsSection(0).Visible = True And tbsSection(0).Tabs.Count > 0 Then
         Me.tbsSection(0).SelectedItem = Me.tbsSection(0).Tabs.Item(1)
         tbsSelFin_Click
         tbsSection_Click 0
      End If
   End If
   lstAlready.Visible = True
   txtAlready.Visible = False
   
   SetMouseNormal

End Sub
Public Sub LookUpJudges()
    Dim iItem As Integer
    Dim cJudge As String
    
    DoEvents
    
    If fraMarks.Visible = True Then
        Me.TestJudges = IIf(IsNull(dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus))), 5, dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus)))
        If IsNull(dtaTestInfo.Recordset.Fields("Num_J_0")) Then
            With dtaTestInfo.Recordset
                .Edit
                .Fields("Num_J_0") = Me.TestJudges
                .Update
            End With
        End If
        If IsNull(dtaTestInfo.Recordset.Fields("Num_J_1")) Then
            With dtaTestInfo.Recordset
                .Edit
                .Fields("Num_J_1") = Me.TestJudges
                .Update
            End With
        End If
        If IsNull(dtaTestInfo.Recordset.Fields("Num_J_2")) Then
            With dtaTestInfo.Recordset
                .Edit
                .Fields("Num_J_2") = Me.TestJudges
                .Update
            End With
        End If
        If IsNull(dtaTestInfo.Recordset.Fields("Num_J_3")) Then
            With dtaTestInfo.Recordset
                .Edit
                .Fields("Num_J_3") = Me.TestJudges
                .Update
            End With
        End If
        For iItem = 1 To 5
             If iItem <= TestJudges Then
                 cJudge = GetJudgeId(TestCode, TestStatus, iItem) & ""
                 If cJudge <> "" Then
                    If dtaTest.Recordset.Fields("Type_Time") = 3 And iItem = 5 Then
                    Else
                        lblMarks(iItem - 1).Caption = GetPersonsName(cJudge) & " - " & Translate("Judge", mcLanguage) & " " & Chr$(64 + iItem) & ":"
                    End If
                 Else
                    lblMarks(iItem - 1).Caption = Translate("Judge", mcLanguage) & " " & Chr$(64 + iItem) & ":"
                 End If
                 lblMarks(iItem - 1).Visible = True
                 txtMarks(iItem - 1) = ""
                 txtMarks(iItem - 1).Visible = True
             Else
                 lblMarks(iItem - 1).Visible = False
                 txtMarks(iItem - 1).Visible = False
             End If
        Next iItem
        If dtaTest.Recordset.Fields("Type_Time") = 3 Then 'PP1
            cmbNumJudges(TestStatus).Text = "5"
            cmbNumJudges(TestStatus).Enabled = False
            Me.lblMarks(4).Caption = TranslateCaption("&Time", lblMarks(4).Width, True) & ":"
            chkFlag.Visible = True
        End If
   ElseIf fraTime.Visible = True Then
        Me.TestJudges = 1
   End If

End Sub
Public Sub LookUpRelevantParticipants()
   Dim iItem As Integer
   Dim iJudge As Integer
   Dim iSections As Integer
   Dim iPosition As Integer
   Dim iFactor As Integer
   Dim iLastUsedSta As Integer
   Dim iEmptyline As Integer
   Dim iTemp As Integer
   Dim iNotFinal As Integer
   Dim iCheckTieBreak As Integer
   Dim curOldResult As Currency
   Dim cOldAllTimes As String
   Dim cOldParticipant As String
   Dim curJudge() As Currency
   Dim cPosition As String
   Dim cOldPosition As String
   Dim cOldList As String
   Dim cScore As String
   Dim cResult As String
   Dim cHtmlResult As String
   Dim cHtml As String
   Dim cHtmlCurrent As String
   Dim cHtmlTest As String
   Dim cOut As String
   Dim cSum As String
   Dim cTemp As String
   Dim cPrevious As String
   Dim rstMarks As DAO.Recordset
   Dim rstLeft As DAO.Recordset
   Dim curHighestScore As Currency
   Dim cHtmlDetails As String
   Dim cQry As String
   Dim cList As String
   
   
   If dtaAlready.RecordSource <> "" Then
      iLastUsedSta = -1
      SetMouseHourGlass
      If TestStatus > 1 Then
            iPosition = GetHighestPosition(TestCode, TestStatus)
      Else
            iPosition = 1
      End If
      
      dtaAlready.Recordset.Requery
    
      lstAlready.Clear
      
      txtAlready.Text = Translate("Wait", mcLanguage) & "..."
      txtAlready.Visible = True
      DoEvents
      lstAlready.Visible = False
      mnuEditTieBreak.Enabled = False
      If dtaAlready.Recordset.RecordCount > 0 Or miUseIceSort = 1 Then
        mnuEditStartOrder.Enabled = False
      ElseIf TestStatus = 0 And miUseIceSort = 0 Then
        mnuEditStartOrder.Enabled = True
      End If
      
      cHtml = "<!-- CreateHTMLBody start -->" & vbCrLf
      
      iNotFinal = True
      Select Case GetTestStatus(dtaTest.Recordset.Fields("Code"))
      Case 1
          If TestStatus = 0 Then
              iNotFinal = False
          End If
      Case 2, 3
          If TestStatus <> 1 Then
              iNotFinal = False
          End If
      End Select
      
      cHtml = cHtml & "<h3>"
      If iNotFinal = True Then
            If fraTime.Visible = True Then
                cHtml = cHtml & Translate("Temporary results - not all participants have started yet", mcLanguage) & " - " & Translate("times are not based upon approved times", mcLanguage) & " (" & Format$(Now, "HH:MM") & " h)"
            Else
                cHtml = cHtml & Translate("Temporary results - not all participants have started yet", mcLanguage) & " (" & Format$(Now, "HH:MM") & " h)"
            End If
      Else
            cHtml = cHtml & Translate("Result list", mcLanguage)
      End If
      cHtml = cHtml & "</h3><p>" & vbCrLf
      
      cHtml = cHtml & "<table>" & vbCrLf
      cHtml = cHtml & "<thead>" & vbCrLf
      cHtml = cHtml & "<tr><th>" & UCase$(Translate("Pos", mcLanguage)) & "</th>"
      cHtml = cHtml & "<th>#</th>"
      cHtml = cHtml & "<th>" & UCase$(Translate("Rider", mcLanguage)) & " / " & UCase$(Translate("Horse", mcLanguage)) & "</th>"
      cHtml = cHtml & "<th>" & UCase$(Translate("Tot", mcLanguage)) & "</th></tr>"
      cHtml = cHtml & "</thead>" & vbCrLf
      cHtml = cHtml & "<tbody>" & vbCrLf
      
      ReDim curJudge(TestJudges)
      cOldAllTimes = ""
      
      If dtaAlready.Recordset.RecordCount > 0 Then
            Do While Not dtaAlready.Recordset.EOF
                iSections = 0
                If dtaAlready.Recordset.AbsolutePosition = 0 Then
                    curHighestScore = dtaAlready.Recordset.Fields("Results.Score")
                End If
               If dtaAlready.Recordset.Fields("Disq") < 0 Then
                    cPosition = mcNoPosition & " "
               ElseIf dtaAlready.Recordset.Fields("Results.Score") = 0 Then
                    cPosition = mcNoPosition & " "
               ElseIf dtaAlready.Recordset.Fields("Results.Score") <> curOldResult Then
                    cPosition = Right$(" " & Format$(iPosition, "00"), 2) & ": "
               ElseIf dtaAlready.Recordset.Fields("AllTimes") & "" <> cOldAllTimes Then
                    cPosition = Right$(" " & Format$(iPosition, "00"), 2) & ": "
               End If
               cOldAllTimes = dtaAlready.Recordset.Fields("AllTimes") & ""
               
               If dtaAlready.Recordset.Fields("Position") & "" <> Val(cPosition) And Val(cPosition) > 0 Then
                    Set rstMarks = mdbMain.OpenRecordset("SELECT Position,Timestamp,Alltimes,Score FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND STA='" & Left$(dtaAlready.Recordset.Fields("cList"), 3) & "'")
                    If rstMarks.RecordCount > 0 Then
                        With rstMarks
                            .Edit
                            .Fields("Position") = Val(cPosition)
                            If .Fields("Score") <> curHighestScore And fraTime.Visible = False And TestStatus = 1 Then
                                'check on tie break
                                .Fields("Alltimes") = ""
                            End If
                            .Fields("TimeStamp") = Now
                            .Update
                        End With
                    End If
                    rstMarks.Close
                    Set rstMarks = Nothing
               End If
               If dtaAlready.Recordset.Fields("Results.Score") = curHighestScore And dtaAlready.Recordset.AbsolutePosition > 0 And TestStatus = 1 And fraMarks.Visible = True Then
                    mnuEditTieBreak.Enabled = True
               End If
               
               If dtaAlready.Recordset.Fields("DISQ") < 0 Then
                    iEmptyline = True
               End If
               
               If dtaAlready.Recordset.Fields("cList") & "" <> cOldList Then
                    If iEmptyline = True And cPosition <> cOldPosition Then
                       iEmptyline = False
                       cHtml = cHtml & "<tr>" & vbCrLf
                       cHtml = cHtml & "<td>&nbsp;</td>"
                       cHtml = cHtml & "<td>&nbsp;</td>"
                       cHtml = cHtml & "<td>&nbsp;</td>"
                       cHtml = cHtml & "<td>&nbsp;</td>"
                       cHtml = cHtml & "</tr>" & vbCrLf
                    End If
                    If iPosition Mod 5 = 0 Then
                       iEmptyline = True
                    End If
               End If
               
               If fraTime.Visible = True Then
                  cScore = ""
               ElseIf TestStatus <> 0 Then
                  cScore = " = " & Format$(dtaAlready.Recordset.Fields("Marks.Score"), TestTotalFormat)
               End If
               
               If dtaAlready.Recordset.Fields("DISQ") = -1 Then
                    cResult = " >> " & UCase$(Translate("ELIMINATED", mcLanguage))
                    cHtmlResult = UCase$(Translate("ELIMINATED", mcLanguage))
               ElseIf dtaAlready.Recordset.Fields("DISQ") = -2 Then
                    cResult = " >> " & Translate("Withdrawn", mcLanguage)
                    cHtmlResult = Translate("Withdrawn", mcLanguage)
               ElseIf fraTime.Visible = True And TestTable = True Then
                    cResult = " >> " & Format$(dtaAlready.Recordset.Fields("Results.Score"), TestTimeFormat)
                    cHtmlResult = Format$(dtaAlready.Recordset.Fields("Results.Score"), TestTimeFormat) & Chr$(34)
               Else
                    cResult = " >> " & Format$(dtaAlready.Recordset.Fields("Results.Score"), TestTotalFormat)
                    cHtmlResult = Format$(dtaAlready.Recordset.Fields("Results.Score"), TestTotalFormat)
               End If
            
               If cOldParticipant <> dtaAlready.Recordset.Fields("cList") & "" Then
                    If iSections > 1 And TestStatus = 0 And dtaAlready.Recordset.AbsolutePosition > 0 And fraMarks.Visible = True Then
                        cSum = ">>>>> "
                        For iJudge = 1 To TestJudges
                            cSum = cSum & Format$(curJudge(iJudge) / iFactor, TestMarkFormat)
                            If iJudge < TestJudges Then
                                cSum = cSum & " - "
                            End If
                            curJudge(iJudge) = 0
                        Next iJudge
                        lstAlready.AddItem vbTab & cSum
                        iFactor = 0
                        iSections = 0
                    End If
                    
                    lstAlready.AddItem cPosition & dtaAlready.Recordset.Fields("cList") & cResult & IIf(dtaAlready.Recordset.Fields("Alltimes") = "1" And dtaAlready.Recordset.Fields("Position") = 1, " [T]", "")
                    lstAlready.ItemData(lstAlready.NewIndex) = dtaAlready.Recordset.Fields("Participants.Sta")
                    cOldParticipant = dtaAlready.Recordset.Fields("cList") & ""
                    iPosition = iPosition + 1
                    If Left$(dtaAlready.Recordset.Fields("cList"), 3) = LastUsedSta Then
                         iLastUsedSta = lstAlready.NewIndex
                         cPrevious = cPosition & dtaAlready.Recordset.Fields("cList") & cResult & IIf(dtaAlready.Recordset.Fields("Alltimes") = "1" And dtaAlready.Recordset.Fields("Position") = 1, " [T]", "")
                    End If
                End If
                
                For iJudge = 1 To TestJudges
                    If Not IsNull(dtaAlready.Recordset.Fields("Mark" & Format$(iJudge))) Then
                        curJudge(iJudge) = curJudge(iJudge) + (dtaAlready.Recordset.Fields("Mark" & Format$(iJudge)) & "")
                    End If
                Next iJudge
                iSections = iSections + 1
                iFactor = iFactor + dtaAlready.Recordset.Fields("Factor")
                If dtaAlready.Recordset.Fields("Out") <> 0 Then
                    cOut = " <<"
                Else
                    cOut = ""
                End If
                If dtaAlready.Recordset.Fields("DISQ") <> -2 Or IsNull(dtaAlready.Recordset.Fields("DISQ")) Then
                    lstAlready.AddItem vbTab & UCase$(Left$(Translate(dtaAlready.Recordset.Fields("Name") & "", mcLanguage), 4)) & " " & dtaAlready.Recordset.Fields("cMarks") & cScore & cOut
                    lstAlready.ItemData(lstAlready.NewIndex) = dtaAlready.Recordset.Fields("Marks.Section")
                    If Left$(dtaAlready.Recordset.Fields("cList"), 3) = LastUsedSta And dtaAlready.Recordset.Fields("Marks.Section") = TestSection Then
                         txtPrevious.Text = FitString(Me, cPrevious, txtPrevious.Width * 0.83, 4) & vbCrLf & vbTab & UCase$(Left$(Translate(dtaAlready.Recordset.Fields("Name") & "", mcLanguage), 4)) & " " & dtaAlready.Recordset.Fields("cMarks") & cScore & cOut
                         txtPrevious.BackColor = vbYellow
                         txtPrevious.Tag = LastUsedSta
                    End If
                End If
               
               If dtaAlready.Recordset.Fields("cList") & "" <> cOldList Then
                    cHtml = cHtml & "<tr>" & vbCrLf
                    cHtml = cHtml & "<td><b>" & cPosition & "</b></td>"
                    
                    'add link to participants details:
                    cHtml = cHtml & "<td><a href=""participant_" & LCase$(Left$(dtaAlready.Recordset.Fields("cList"), 3)) & ".html"">" & Left$(dtaAlready.Recordset.Fields("cList"), 3) & "</a></td>"
                    
                    'Create HTML page with participants details:
                    CreateHTMLDetails Left$(dtaAlready.Recordset.Fields("cList"), 3)
                    
                    cHtml = cHtml & "<td><b>" & Mid$(dtaAlready.Recordset.Fields("cList"), 7) & IIf(dtaAlready.Recordset.Fields("Alltimes") = "1" And dtaAlready.Recordset.Fields("Position") = 1, "&nbsp;[" & Translate("Tiebreak", mcLanguage) & "]", "")
                    cHtml = cHtml & "</b></td>"
                    cHtml = cHtml & "<td><b>" & cHtmlResult & "</b></td>"
                    cHtml = cHtml & "</tr>" & vbCrLf
                End If
                
                If dtaAlready.Recordset.Fields("DISQ") <> -2 Or IsNull(dtaAlready.Recordset.Fields("DISQ")) Then
                    cHtml = cHtml & "<tr>" & vbCrLf
                    cHtml = cHtml & "<td>&nbsp;</td>"
                    cHtml = cHtml & "<td>&nbsp;</td>"
                    cHtml = cHtml & "<td><a title=""" & Translate(dtaAlready.Recordset.Fields("Name") & "", mcLanguage) & """>" & UCase$(Left$(Translate(dtaAlready.Recordset.Fields("Name") & "", mcLanguage), 4)) & "</a> " & dtaAlready.Recordset.Fields("cMarks") & cScore & cOut
                    cHtml = cHtml & "</td>"
                    cHtml = cHtml & "<td>&nbsp;</td>"
                    cHtml = cHtml & "</tr>" & vbCrLf
               End If
               
               cOldList = dtaAlready.Recordset.Fields("cList") & ""
               curOldResult = dtaAlready.Recordset.Fields("Results.Score")
               dtaAlready.Recordset.MoveNext
               cOldPosition = cPosition
            Loop
            If iSections > 1 And TestStatus = 0 And dtaAlready.Recordset.AbsolutePosition <> 0 And fraMarks.Visible = True Then
                cSum = ">>>>> "
                For iJudge = 1 To TestJudges
                    cSum = cSum & Format$(curJudge(iJudge) / iFactor, TestMarkFormat)
                    If iJudge < TestJudges Then
                        cSum = cSum & " - "
                    End If
                    curJudge(iJudge) = 0
                Next iJudge
                lstAlready.AddItem vbTab & cSum
                iFactor = 0
            End If
            cHtml = cHtml & "</tbody>" & vbCrLf
            cHtml = cHtml & "</table>" & vbCrLf
            
            dtaNotYet.Recordset.Requery
            
            'test results without not yest started
            cHtmlTest = cHtml
            cHtmlTest = cHtmlTest & "<!-- CreateHTMLBody end -->" & vbCrLf
            
            'Add participants not yet started to current.html:
            If dtaNotYet.Recordset.RecordCount > 0 And TestStatus = 0 Then
                cHtml = cHtml & "<p><HR>"
                cHtml = cHtml & "<p><b>" & Translate("Participants not yet started", mcLanguage)
                If TestSection > 1 Then
                    cHtml = cHtml & " (" & Translate("Heat", mcLanguage) & " " & TestSection & ")"
                End If
                cHtml = cHtml & " [" & dtaNotYet.Recordset.RecordCount & "]"
                cHtml = cHtml & "<b>"
                cHtml = cHtml & "<p><table>" & vbCrLf
                Do While Not dtaNotYet.Recordset.EOF

                    With dtaNotYet.Recordset
                        cHtml = cHtml & "<tr><td>"
                        cTemp = Replace(.Fields("cList"), "[" & Left$(Translate("Equipment check", mcLanguage), 1) & "]", "")
                        cTemp = "<a href=""participant_" & LCase$(Left$(cTemp, 3)) & ".html"">" & Left$(cTemp, 3) & "</a>" & Mid$(cTemp, 4)
                        cHtml = cHtml & cTemp
                        cHtml = cHtml & "</td></tr>" & vbCrLf
                        .MoveNext
                    End With
                Loop
                dtaNotYet.Recordset.MoveFirst
                cHtml = cHtml & "</table>" & vbCrLf
            End If
                        
            cHtmlCurrent = cHtml
            cHtmlCurrent = cHtmlCurrent & "<!-- CreateHTMLBody end -->" & vbCrLf
            
            fraNotYet.Caption = Translate("Participants not yet started", mcLanguage) & " [" & dtaNotYet.Recordset.RecordCount & "]"
            fraAlready.Caption = Translate("Participants already started", mcLanguage) & " [" & dtaAlready.Recordset.RecordCount & "]"
            SetFocusTo txtParticipant
            
            'Create temporary HTML:
            CreateHTML cHtmlCurrent, cHtmlTest
        Else
            cHtml = cHtml & "<tr><td><b>-</b></td>"
            cHtml = cHtml & "<td> </td>"
            cHtml = cHtml & "<td><b>- " & Translate("No results available", mcLanguage) & " -</b></td>"
            cHtml = cHtml & "<td> </td>"
            cHtml = cHtml & "</tr>" & vbCrLf
            cHtml = cHtml & "</table>"
            
            'Add participants not yet started to current.html:
            If dtaNotYet.Recordset.RecordCount > 0 And TestStatus = 0 Then
                cHtml = cHtml & "<p><HR>"
                cHtml = cHtml & "<p><b>" & Translate("Participants not yet started", mcLanguage)
                If TestSection > 1 Then
                    cHtml = cHtml & " (" & Translate("Heat", mcLanguage) & " " & TestSection & ")"
                End If
                cHtml = cHtml & " [" & dtaNotYet.Recordset.RecordCount & "]"
                cHtml = cHtml & "<b>"
                cHtml = cHtml & "<p><table>" & vbCrLf
                Do While Not dtaNotYet.Recordset.EOF
                    With dtaNotYet.Recordset
                        cHtml = cHtml & "<tr><td>"
                        cTemp = Replace(.Fields("cList"), "[" & Left$(Translate("Equipment check", mcLanguage), 1) & "]", "")
                        cTemp = "<a href=""participant_" & LCase$(Left$(cTemp, 3)) & ".html"">" & Left$(cTemp, 3) & "</a>" & Mid$(cTemp, 4)
                        cHtml = cHtml & cTemp
                        cHtml = cHtml & "</td></tr>" & vbCrLf
                        .MoveNext
                    End With
                Loop
                dtaNotYet.Recordset.MoveFirst
                cHtml = cHtml & "</table>" & vbCrLf
            End If
                        
            cHtmlCurrent = cHtml
            cHtmlCurrent = cHtmlCurrent & "<!-- CreateHTMLBody end -->" & vbCrLf
            CreateHTML cHtmlCurrent, cHtmlTest
        
        End If
        
        'LL 2007-8-1: Switching back to only one HtmlDir
        'If miHtmlFiles <> 0 Then
        '    iTemp = CopyHTML(mcTempHtmlDir, mcHtmlDir)
        'End If
                
        'show the top of the list again
        If lstAlready.ListCount > 0 Then
            lstAlready.ListIndex = 0
        End If
        lstAlready.Visible = True
        
        txtAlready.Visible = False
        chkSplitFinals.Enabled = False
        chkSplitResultLists.Enabled = False
        'check if it is possible to re-compose finals
        Select Case TestStatus
            'This is a C-Final:
            Case 3
                If ComposeFinalsAllowed = True Then
                    Me.cmdComposeFinals.Enabled = True
                    '* MM
                    '* split finals in different classes not available when C-Finals are used.
                    '*
                    chkSplitFinals.Enabled = False
                    chkSplitResultLists.Enabled = False
                Else
                    Me.cmdComposeFinals.Enabled = False
                End If
            'This is a B-Final:
            Case 2
                If ComposeFinalsAllowed = True Then
                    Me.cmdComposeFinals.Enabled = True
                    '* check if it is still possible to split finals in different classes
                    '* split finals in different classes not available when C-Finals are used.
                    '*
                    If SplitFinalsAllowed = True And Me.dtaTestInfo.Recordset.Fields("Handling") <= 4 Then
                        chkSplitFinals.Enabled = True
                        chkSplitResultLists.Enabled = True
                    Else
                        chkSplitFinals.Enabled = False
                        chkSplitResultLists.Enabled = False
                    End If
                Else
                    Me.cmdComposeFinals.Enabled = False
                End If
            'This is an A-Final:
            Case 1
                If ComposeFinalsAllowed = True Then
                    Me.cmdComposeFinals.Enabled = True
                    '* check if it is still possible to split finals in different classes
                    '* split finals in different classes not available when C-Finals are used.
                    '*
                    If SplitFinalsAllowed = True And Me.dtaTestInfo.Recordset.Fields("Handling") <= 4 Then
                        chkSplitFinals.Enabled = True
                        chkSplitResultLists.Enabled = True
                    Else
                        chkSplitFinals.Enabled = False
                        chkSplitResultLists.Enabled = False
                    End If
                Else
                    Me.cmdComposeFinals.Enabled = False
                End If
                CheckTieBreak
            Case 0
                If SplitFinalsAllowed = True Then
                    Me.chkSplitResultLists.Value = Me.chkSplitFinals.Value
                    chkSplitFinals.Enabled = True
                    chkSplitResultLists.Enabled = True
                End If
        End Select
        
        Me.Tag = "*"
        If miExcelFiles <> 0 Then
            CreateExcelRanking
        End If
        
    End If
    
    SetMouseNormal
   
    Timer1.Enabled = False
End Sub
Private Sub StoreCurrentMarks()
   Dim iKey As Integer
   Dim iItem As Integer
   
   If fraMarks.Visible = True Then
        ValidateScore
   ElseIf fraTime.Visible = True Then
        ValidateTimeScore
   End If
   
   StatusMessage
   
   'store current marks first
   If lblParticipant.Caption <> "" Then
      If lblParticipant.Tag <> "" Then
         iKey = MsgBox(Translate("Save results for", mcLanguage) & ": " & txtParticipant.Text & vbCrLf & lblParticipant.Caption, vbQuestion + vbYesNo + vbDefaultButton1)
         If iKey = vbYes Then
            cmdOK_Click
         End If
         lblParticipant.Tag = ""
      Else
         ClearMarks
         Timer1.Enabled = False
      End If
   End If
End Sub
Private Function ValidateMark(Index As Integer) As Integer
   Dim curMark As Currency
   Dim rstSectionMark As DAO.Recordset
   Dim iKey As Integer
   
   On Local Error Resume Next
   
   ValidateMark = True
   
   If Trim$(txtMarks(Index).Text) <> "" Then
      
      If (TestStatus = 0 And dtaTest.Recordset.Fields("Type_pre") = 2) Or (TestStatus <> 0 And dtaTest.Recordset.Fields("Type_Final") = 2) Then
            txtMarks(Index).Text = Val(txtMarks(Index).Text)
      ElseIf InStr(txtMarks(Index).Text, ",") = 0 And InStr(txtMarks(Index).Text, ".") = 0 And Val(txtMarks(Index).Text) > 10 Then
          If Err > 0 Then
            txtMarks(Index).Text = ""
          End If
          If txtMarks(Index).Text <= 10 ^ (TestMarkDecimals) Then
             txtMarks(Index).Text = Val(txtMarks(Index).Text) / 10 ^ (TestMarkDecimals - 1)
          ElseIf txtMarks(Index).Text <= 10 ^ (TestMarkDecimals + 1) Then
             txtMarks(Index).Text = Val(txtMarks(Index).Text) / 10 ^ TestMarkDecimals
          ElseIf txtMarks(Index).Text <= 10 ^ (TestMarkDecimals + 2) Then
             txtMarks(Index).Text = Val(txtMarks(Index).Text) / 10 ^ (TestMarkDecimals + 1)
          End If
      End If
      curMark = MakeStringValue(txtMarks(Index).Text)
      If dtaTest.Recordset.Fields("Type_Time") = 3 And Index = 4 Then
         txtMarks(Index).Text = Format$(curMark, TestMarkFormat)
         If txtMarks(Index).BackColor <> QBColor(15) Then
            txtMarks(Index).BackColor = QBColor(15)
            miNoBackupNow = False
         End If
      ElseIf ((TestStatus = 0 And dtaTest.Recordset.Fields("Type_pre") = 1) Or (TestStatus <> 0 And dtaTest.Recordset.Fields("Type_Final") = 1)) And curMark < dtaTestSection.Recordset.Fields("Mark_low") Or curMark > dtaTestSection.Recordset.Fields("Mark_hi") Then
         ValidateMark = False
         StatusMessage Translate("Only marks between", mcLanguage) & " " & dtaTestSection.Recordset.Fields("Mark_low") & " - " & dtaTestSection.Recordset.Fields("Mark_hi") & " " & Translate("are accepted", mcLanguage) & " !", 2
         txtMarks(Index).SelStart = 0
         txtMarks(Index).SelLength = 5
         SetFocusTo txtMarks(Index)
         txtMarks(Index).BackColor = QBColor(12)
         miNoBackupNow = True
         PlaySound "SYSTEMEXCLAMATION", 0, 1
         fraCurrent.Enabled = True
         fraMarks.Enabled = True
         miInvalidMark = True
      Else
         txtMarks(Index).Text = Format$(curMark, TestMarkFormat)
         If txtMarks(Index).BackColor <> QBColor(15) Then
            txtMarks(Index).BackColor = QBColor(15)
            miNoBackupNow = False
         End If
      End If

      ValidateScore
   Else
        If txtMarks(Index).BackColor <> mlAlertColor Then
           txtMarks(Index).BackColor = mlAlertColor
           miNoBackupNow = True
        End If
   End If
End Function
Private Function ValidateTime() As Integer
    Dim curTime As Currency
    ValidateTime = True
    If txtTime.Text <> "" Then
        If InStr(txtTime.Text, ",") = 0 And InStr(txtTime.Text, ".") = 0 Then
            If dtaTest.Recordset.Fields("Type_time") = 3 Then 'Pace test
                txtTime.Text = Val(txtTime.Text) / 10 ^ TestMarkDecimals
            ElseIf Val(txtTime.Text) >= 10 ^ TestTimeDecimals And Val(txtTime.Text) < 10000 Then
                txtTime.Text = Val(txtTime.Text) / 10 ^ TestTimeDecimals
            End If
        End If
        curTime = MakeStringValue(txtTime.Text)
        txtTime.Text = Format$(curTime, TestTimeFormat)
        If txtTime.BackColor <> QBColor(15) Then
           txtTime.BackColor = QBColor(15)
            miNoBackupNow = False
        End If
        ValidateTimeScore
    End If
End Function
Private Sub ValidateTimeScore()
   Dim curScore As Currency
   curScore = MakeStringValue(txtTime.Text)
   If chkFlag.Value <> 0 Then
       txtScore.FontStrikethru = True
       txtScore.Text = TestTimeFormat
   Else
       txtScore.FontStrikethru = False
       txtScore.Text = Format$(curScore, TestTimeFormat)
   End If
End Sub
Private Sub ValidateScore()
   Dim curMark As Currency
   Dim curTimeValue
   Dim curHi As Currency
   Dim curLo As Currency
   Dim curScore As Currency
   Dim iTimeIsZero As Integer
   Dim iCountZero As Integer
   Dim iItem As Integer
   Dim iIsMark As Integer
   Dim iKey As Integer
   Dim iPPDivider As Integer
   
   On Local Error Resume Next
   
   iPPDivider = 6
    
    If Not IsNull(dtaTest.Recordset.Fields("Div_Pre")) Then
        iPPDivider = dtaTest.Recordset.Fields("Div_Pre")
    End If
   
   If lblParticipant.Caption <> "" Then
      iIsMark = False
      miInvalidMark = False
      For iItem = 0 To TestJudges - 1
         If txtMarks(iItem).Text <> "" Then
            iIsMark = True
         End If
         curMark = MakeStringValue(txtMarks(iItem).Text)
         txtMarks(iItem).FontStrikethru = False
         If dtaTest.Recordset.Fields("Type_time") = 3 Then 'Pace test
            If iItem = 4 Then
                If iTimeIsZero = 1 Or Me.chkFlag.Value <> 0 Then
                    txtMarks(iItem).FontStrikethru = True
                Else
                    txtMarks(iItem).FontStrikethru = False
                    curScore = curScore + Time2Mark(curMark, "")
                End If
            Else
                If curMark = 0 Then
                    iCountZero = iCountZero + 1
                    If iItem = 1 Or iItem = 2 Then
                        iTimeIsZero = 1
                    End If
                End If
                curScore = curScore + curMark
            End If
         ElseIf curMark >= dtaTestSection.Recordset.Fields("Mark_low") And curMark <= dtaTestSection.Recordset.Fields("Mark_hi") Then
            If iItem = 0 Then
               curLo = curMark
            End If
            If curMark < curLo Then
               curLo = curMark
            ElseIf curMark > curHi Then
               curHi = curMark
            End If
            curScore = curScore + curMark
         ElseIf (TestStatus = 0 And dtaTest.Recordset.Fields("Type_pre") = 2) Or (TestStatus <> 0 And dtaTest.Recordset.Fields("Type_Final") = 2) Then
            If iItem = 0 Then
               curLo = curMark
            End If
            If curMark < curLo Then
               curLo = curMark
            ElseIf curMark > curHi Then
               curHi = curMark
            End If
            curScore = curScore + curMark
         Else
            miInvalidMark = True
            SetFocusTo txtMarks(iItem)
            curMark = 0
         End If
      Next iItem
      
      If dtaTest.Recordset.Fields("Type_Special") = 1 Or dtaTest.Recordset.Fields("Type_Special") = 3 Then
         curScore = curScore / TestJudges
      ElseIf dtaTest.Recordset.Fields("Type_time") = 3 Then 'Pace Test
         If iCountZero >= 3 Then
            If dtaTest.Recordset.Fields("Type_special") = 2 Then
                curScore = 0
            End If
         Else
            curScore = curScore / iPPDivider
         End If
      ElseIf TestJudges = 5 Then
         curScore = (curScore - curLo - curHi) / 3
      Else
         curScore = curScore / TestJudges
      End If
      If iIsMark And Not miInvalidMark Then
         txtScore.Text = Format$(curScore, TestTotalFormat)
      Else
         txtScore.Text = ""
      End If
   Else
      txtScore.Text = ""
   End If
End Sub
Public Sub ClearMarks(Optional intKeepNumber As Integer = False)
   Dim iItem As Integer
   For iItem = 0 To TestJudges - 1
      txtMarks(iItem).Text = ""
      If txtMarks(iItem).BackColor <> QBColor(15) Then
         txtMarks(iItem).BackColor = QBColor(15)
         miNoBackupNow = False
      End If
   Next iItem
   txtTime.Text = ""
   If intKeepNumber = False Then
        txtParticipant.Text = ""
        lblParticipant.Caption = ""
        lblParticipant.Tag = ""
   End If
   txtScore.Text = ""
   chkDisqualified.Value = 0
   chkWithdrawn.Value = 0
   chkWithdrawn.Enabled = True
   chkNoStart.Value = 0
   chkNoStart.Enabled = True
   chkFlag.Value = 0
   chkFlag.Enabled = True

End Sub
Public Sub StartMenuPopUp()
    Dim ctl As Control
    Dim iTeller As Integer
    Dim iTestTeller As Integer
    Dim iSorted As Integer
    Dim cOldMenu As String
    Dim cOldTestMenu As String
    Dim cTmpCaption As String
    Dim cTmpTag As String
    
    Do While mnuPopupPopUp.Count > 1
        Unload mnuPopupPopUp.Item(mnuPopupPopUp.Count - 1)
    Loop
    DoEvents

    mnuPopupPopUp.Item(0).Visible = True
    iTeller = 1
    iTestTeller = 0
    cOldMenu = ""
    For Each ctl In Controls
        Select Case Left$(ctl.Name, 7)
        Case "mnuFile", "mnuEdit"
            If ctl.Visible = True And ctl.Enabled = True And Len(ctl.Name) > 7 And ReadTagItem(ctl, "PopUp") <> "No" And ctl.Caption <> "*" Then
                If ctl.Caption <> mnuPopupPopUp.Item(mnuPopupPopUp.Count - 1).Caption Then
                    If cOldMenu <> Left$(ctl.Name, 7) And cOldMenu <> "" Then
                        iTeller = iTeller + 1
                        If iTeller > mnuPopupPopUp.Count Then
                            Load mnuPopupPopUp.Item(mnuPopupPopUp.Count)
                        End If
                        mnuPopupPopUp.Item(mnuPopupPopUp.Count - 1).Caption = "-"
                    End If
                    iTeller = iTeller + 1
                    If iTeller > mnuPopupPopUp.Count Then
                        Load mnuPopupPopUp.Item(mnuPopupPopUp.Count)
                    End If
                    mnuPopupPopUp.Item(mnuPopupPopUp.Count - 1).Caption = ctl.Caption
                    If Left$(ctl.Name, 11) = "mnuTestQual" Then
                        ChangeTagItem mnuPopupPopUp.Item(mnuPopupPopUp.Count - 1), "Control", "mnuTestQual"
                    Else
                        ChangeTagItem mnuPopupPopUp.Item(mnuPopupPopUp.Count - 1), "Control", ctl.Name
                    End If
                    ChangeTagItem mnuPopupPopUp.Item(mnuPopupPopUp.Count - 1), "Tag", ctl.Tag
                    cOldMenu = Left$(ctl.Name, 7)
                End If
            End If
        Case "mnuTest"
            If ctl.Visible = True And ctl.Enabled = True And Len(ctl.Name) > 7 And ReadTagItem(ctl, "PopUp") <> "No" And ctl.Caption <> "*" Then
                If ctl.Caption <> Me.mnuPopUpTestsTest.Item(mnuPopUpTestsTest.Count - 1).Caption Then
                    If cOldTestMenu <> Left$(ctl.Name, 7) And cOldTestMenu <> "" Then
                        iTestTeller = iTestTeller + 1
                        If iTestTeller > mnuPopUpTestsTest.Count Then
                            Load mnuPopUpTestsTest.Item(mnuPopUpTestsTest.Count)
                        End If
                        mnuPopUpTestsTest.Item(mnuPopUpTestsTest.Count - 1).Caption = "-"
                    End If
                    Me.mnuPopupSep1.Visible = True
                    Me.mnuPopUpTests.Visible = True
                    iTestTeller = iTestTeller + 1
                    If iTestTeller > mnuPopUpTestsTest.Count Then
                        Load mnuPopUpTestsTest.Item(mnuPopUpTestsTest.Count)
                    End If
                    mnuPopUpTestsTest.Item(mnuPopUpTestsTest.Count - 1).Caption = ctl.Caption
                    If Left$(ctl.Name, 11) = "mnuTestQual" Then
                        ChangeTagItem mnuPopUpTestsTest.Item(mnuPopUpTestsTest.Count - 1), "Control", "mnuTestQual"
                    Else
                        ChangeTagItem mnuPopUpTestsTest.Item(mnuPopUpTestsTest.Count - 1), "Control", ctl.Name
                    End If
                    ChangeTagItem mnuPopUpTestsTest.Item(mnuPopUpTestsTest.Count - 1), "Tag", ctl.Tag
                    cOldTestMenu = Left$(ctl.Name, 7)
                End If
            End If
        End Select
    Next
    mnuPopupPopUp.Item(0).Visible = False
    mnuPopUpTests.Caption = Translate("Tests", mcLanguage)
    
    If mnuPopUpTestsTest.Count > 1 Then
        Do
            iSorted = False
            For iTeller = 0 To mnuPopUpTestsTest.Count - 2
                If mnuPopUpTestsTest(iTeller).Caption > mnuPopUpTestsTest(iTeller + 1).Caption Then
                    iSorted = True
                    cTmpCaption = mnuPopUpTestsTest(iTeller + 1).Caption
                    cTmpTag = mnuPopUpTestsTest(iTeller + 1).Tag
                    mnuPopUpTestsTest(iTeller + 1).Caption = mnuPopUpTestsTest(iTeller).Caption
                    mnuPopUpTestsTest(iTeller + 1).Tag = mnuPopUpTestsTest(iTeller).Tag
                    mnuPopUpTestsTest(iTeller).Caption = cTmpCaption
                    mnuPopUpTestsTest(iTeller).Tag = cTmpTag
                End If
            Next iTeller
        Loop While iSorted = True
    End If
    
    PopupMenu mnuPopup
   
End Sub
Public Function CalculateResult(cSta As String) As Currency
    Dim rstMarks As Recordset
    Dim rstTest As Recordset
    Dim cQry As String
    Dim iFactor As Integer
    Dim curResult As Currency
    Dim curDeduct As Currency
    Dim curJudge() As Currency
    Dim curLo As Currency
    Dim curHi As Currency
    Dim iJudge As Integer
    Dim iBit As Integer
   
    If TestStatus > 0 Or dtaTest.Recordset.Fields("Type_time") = 3 Then 'finals and pace test
        '* mark sections to be taken out
        Set rstTest = mdbMain.OpenRecordset("SELECT Out_Fin FROM Tests WHERE Code LIKE '" & Me.TestCode & "'")
        If rstTest.Fields("Out_Fin") > 0 Then
            cQry = "SELECT * FROM Marks "
            cQry = cQry & " WHERE Code='" & Me.TestCode & "' "
            cQry = cQry & " AND Status=" & Me.TestStatus
            cQry = cQry & " AND STA='" & cSta & "'"
            cQry = cQry & " AND Section "
            cQry = cQry & " IN (SELECT Section "
            cQry = cQry & " FROM TestSections "
            cQry = cQry & " WHERE Code='" & Me.TestCode & "'"
            If TestStatus = 3 Then
                 cQry = cQry & " AND Status = 1"
            ElseIf TestStatus = 2 Then
                 cQry = cQry & " AND Status = 1"
            Else
                 cQry = cQry & " AND Status = " & Me.TestStatus
            End If
            cQry = cQry & " AND Out=-1)"
            cQry = cQry & " ORDER BY Score"
            Set rstMarks = mdbMain.OpenRecordset(cQry)
            If rstMarks.RecordCount > 0 Then
                Do While Not rstMarks.EOF
                    With rstMarks
                        .Edit
                        If .AbsolutePosition < rstTest.Fields("Out_Fin") Then
                            .Fields("Out") = 1
                        Else
                            .Fields("Out") = 0
                        End If
                        .Update
                        .MoveNext
                    End With
                Loop
            End If
            rstMarks.Close
        End If
        rstTest.Close
           
        cQry = "SELECT Marks.*,Testsections.Factor,Testsections.Out"
        cQry = cQry & " FROM Marks"
        cQry = cQry & " INNER JOIN Testsections"
        If TestStatus = 3 Then
             cQry = cQry & " ON Testsections.Status = Marks.Status-2"
        ElseIf TestStatus = 2 Then
             cQry = cQry & " ON Testsections.Status = Marks.Status-1"
        Else
             cQry = cQry & " ON Testsections.Status = Marks.Status"
        End If
        cQry = cQry & " AND Testsections.Section = Marks.Section"
        cQry = cQry & " AND Marks.Code = Testsections.Code"
        cQry = cQry & " WHERE Marks.Sta='" & cSta & "'"
        cQry = cQry & " AND Marks.Code='" & Me.TestCode & "'"
        cQry = cQry & " AND Marks.Status=" & Me.TestStatus
        cQry = cQry & " ORDER BY Marks.Score"
        
        Set rstMarks = mdbMain.OpenRecordset(cQry)
        iFactor = 0
        curResult = 0
        Do While Not rstMarks.EOF
             If rstMarks.Fields("Marks.Out") = 0 Or IsNull(rstMarks.Fields("Marks.Out")) Then
                 curResult = curResult + (rstMarks.Fields("Score") * rstMarks.Fields("Factor"))
                 iFactor = iFactor + rstMarks.Fields("Factor")
             End If
             rstMarks.MoveNext
        Loop
        rstMarks.Close
        
        Set rstMarks = Nothing
        Set rstTest = Nothing
        
        If iFactor <> 0 Then
           CalculateResult = MakeStringValue(Format$(curResult / iFactor, TestTotalFormat))
        Else
           CalculateResult = 0
        End If
    Else
        ReDim curJudge(Me.TestJudges)
        curLo = 9999
        curHi = 0
        '* mark sections to be taken out
        For iJudge = 1 To TestJudges
            Set rstTest = mdbMain.OpenRecordset("SELECT Out_Fin FROM Tests WHERE Code LIKE '" & Me.TestCode & "'")
            If rstTest.Fields("Out_Fin") > 0 Then
                 cQry = "SELECT * FROM Marks "
                 cQry = cQry & " WHERE Code='" & Me.TestCode & "' "
                 cQry = cQry & " AND Status=" & Me.TestStatus
                 cQry = cQry & " AND STA='" & cSta & "'"
                 cQry = cQry & " AND Section "
                 cQry = cQry & " IN (SELECT Section "
                 cQry = cQry & " FROM TestSections "
                 cQry = cQry & " WHERE Code='" & Me.TestCode & "'"
                If TestStatus = 3 Then
                     cQry = cQry & " AND Status = 1"
                ElseIf TestStatus = 2 Then
                     cQry = cQry & " AND Status = 1"
                Else
                     cQry = cQry & " AND Status = " & Me.TestStatus
                End If
                cQry = cQry & " AND Out=-1)"
                cQry = cQry & " ORDER BY Score"
                Set rstMarks = mdbMain.OpenRecordset(cQry)
                If rstMarks.RecordCount > 0 Then
                    Do While Not rstMarks.EOF
                        With rstMarks
                            .Edit
                            If .AbsolutePosition < rstTest.Fields("Out_Fin") Then
                                .Fields("Out") = 1
                            Else
                                .Fields("Out") = 0
                            End If
                            .Update
                            .MoveNext
                        End With
                    Loop
                End If
                rstMarks.Close
            End If
            rstTest.Close
               
            cQry = "SELECT Marks.*,Testsections.Factor,Testsections.Out"
            cQry = cQry & " FROM Marks"
            cQry = cQry & " INNER JOIN Testsections"
            If TestStatus = 3 Then
                 cQry = cQry & " ON Testsections.Status = Marks.Status-2"
            ElseIf TestStatus = 2 Then
                 cQry = cQry & " ON Testsections.Status = Marks.Status-1"
            Else
                 cQry = cQry & " ON Testsections.Status = Marks.Status"
            End If
            cQry = cQry & " AND Testsections.Section = Marks.Section"
            cQry = cQry & " AND Marks.Code = Testsections.Code"
            cQry = cQry & " WHERE Marks.Sta='" & cSta & "'"
            cQry = cQry & " AND Marks.Code='" & Me.TestCode & "'"
            cQry = cQry & " AND Marks.Status=" & Me.TestStatus
            cQry = cQry & " ORDER BY Marks.Score"
            
            Set rstMarks = mdbMain.OpenRecordset(cQry)
            iFactor = 0
            curResult = 0
            curJudge(iJudge) = 0
            Do While Not rstMarks.EOF
                If dtaTest.Recordset.Fields("Type_Time") = 3 And iJudge = 4 Then
                    curJudge(iJudge) = curJudge(iJudge) + Time2Mark(rstMarks.Fields("Mark" & Format$(iJudge)), dtaTest.Recordset.Fields("Code")) * rstMarks.Fields("factor")
                    iFactor = iFactor + rstMarks.Fields("Factor")
                Else
                    curJudge(iJudge) = curJudge(iJudge) + rstMarks.Fields("Mark" & Format$(iJudge)) * rstMarks.Fields("factor")
                    iFactor = iFactor + rstMarks.Fields("Factor")
                End If
                rstMarks.MoveNext
            Loop
            rstMarks.Close
            
            Set rstMarks = Nothing
            Set rstTest = Nothing
            
            If iFactor <> 0 Then
               curJudge(iJudge) = curJudge(iJudge) / iFactor
            Else
               curJudge(iJudge) = 0
            End If
            If curJudge(iJudge) < curLo Then
                curLo = curJudge(iJudge)
            End If
            If curJudge(iJudge) > curHi Then
                curHi = curJudge(iJudge)
            End If
            CalculateResult = CalculateResult + curJudge(iJudge)

        Next iJudge
        
        If dtaTest.Recordset.Fields("Type_Special") = 1 Or dtaTest.Recordset.Fields("Type_Special") = 3 Then
            CalculateResult = CalculateResult / TestJudges
        ElseIf Me.TestJudges = 5 Then
            CalculateResult = (CalculateResult - curHi - curLo) / 3
        Else
            CalculateResult = CalculateResult / TestJudges
        End If
    End If
End Function
Private Function CalculateTime(cSta As String) As Currency
   Dim rstMarks As Recordset
   
   Dim cQry As String
   Dim curResult As Currency
   Dim cAllTimes As String
   Dim iTemp As Integer
   
   For iTemp = 1 To 4
        cAllTimes = cAllTimes & Format$(99.99, "00.00 ")
   Next iTemp
   
   cQry = "SELECT Score,Section"
   cQry = cQry & " FROM Marks"
   cQry = cQry & " WHERE Sta='" & cSta & "'"
   cQry = cQry & " AND Code='" & Me.TestCode & "'"
   cQry = cQry & " AND Status=" & Me.TestStatus
   cQry = cQry & " ORDER BY Score=0 DESC,Score"
   
   Set rstMarks = mdbMain.OpenRecordset(cQry)
   CalculateTime = 0
   curResult = 9999
   Do While Not rstMarks.EOF
        If rstMarks.Fields("Score") > 0 Then
            Mid$(cAllTimes, rstMarks.AbsolutePosition * 6 + 1, 5) = Format$(rstMarks.Fields("Score"), "00.00")
            If rstMarks.Fields("Score") < curResult Then
                curResult = rstMarks.Fields("Score")
            End If
        End If
        rstMarks.MoveNext
   Loop
   
   rstMarks.Close
   Set rstMarks = Nothing
   
   If curResult <> 9999 Then
      CalculateTime = MakeStringValue(Format$(curResult, TestTimeFormat))
   End If
    
End Function
Public Function CalculatePP1Times(cSta As String) As String
   Dim rstMarks As Recordset
   
   Dim cQry As String
   Dim curResult As Currency
   Dim cAllTimes As String
   Dim iTemp As Integer
   Dim iSections As Integer
   
    iSections = 2
    'Make sure we're dealing with 3 heats in PP2:
    If dtaTest.Recordset.Fields("out_fin") = 1 Then
        iSections = 3
    End If
   
   For iTemp = 1 To iSections
        cAllTimes = cAllTimes & Format$(99.99, "00.00 ")
   Next iTemp
   
   cQry = "SELECT Mark5"
   cQry = cQry & " FROM Marks"
   cQry = cQry & " WHERE Sta='" & cSta & "'"
   cQry = cQry & " AND Code='" & Me.TestCode & "'"
   cQry = cQry & " AND Status=" & Me.TestStatus
   cQry = cQry & " ORDER BY Mark5=0 DESC,Mark5"
   
   Set rstMarks = mdbMain.OpenRecordset(cQry)
   curResult = 9999
   Do While Not rstMarks.EOF
        If rstMarks.Fields("Mark5") > 0 Then
            Mid$(cAllTimes, rstMarks.AbsolutePosition * 6 + 1, 5) = Format$(rstMarks.Fields("Mark5"), "00.00")
        End If
        rstMarks.MoveNext
   Loop
   
   rstMarks.Close
   Set rstMarks = Nothing
   
   CalculatePP1Times = cAllTimes
   

End Function
Public Function CalculateAllTimes(cSta As String) As String
   Dim rstMarks As Recordset
   
   Dim cQry As String
   Dim curResult As Currency
   Dim cAllTimes As String
   Dim iTemp As Integer
   
   For iTemp = 1 To 4
        cAllTimes = cAllTimes & Format$(99.99, "00.00 ")
   Next iTemp
   
   cQry = "SELECT Score,Section"
   cQry = cQry & " FROM Marks"
   cQry = cQry & " WHERE Sta='" & cSta & "'"
   cQry = cQry & " AND Code='" & Me.TestCode & "'"
   cQry = cQry & " AND Status=" & Me.TestStatus
   cQry = cQry & " ORDER BY Score=0 DESC,Score"
   
   Set rstMarks = mdbMain.OpenRecordset(cQry)
   curResult = 9999
   Do While Not rstMarks.EOF
        If rstMarks.Fields("Score") > 0 Then
            Mid$(cAllTimes, rstMarks.AbsolutePosition * 6 + 1, 5) = Format$(rstMarks.Fields("Score"), "00.00")
            If rstMarks.Fields("Score") < curResult Then
                curResult = rstMarks.Fields("Score")
            End If
        End If
        rstMarks.MoveNext
   Loop
   
   rstMarks.Close
   Set rstMarks = Nothing
   
   CalculateAllTimes = cAllTimes
   

End Function

Sub ChangeDatabase(Optional NewDatabasePath As String, Optional NewDatabaseName As String = "")
    
    On Local Error Resume Next
    
    If NewDatabasePath <> "" Then
        If NewDatabaseName <> "" Then
            mcDatabaseName = NewDatabasePath & NewDatabaseName
        Else
            mcDatabaseName = NewDatabasePath & UserName & "1.Mdb"
        End If
    End If
    With Me.CommonDialog1
        .CancelError = True
        .DefaultExt = ".Mdb"
        .DialogTitle = "Select a folder"
        .Filter = "Database|*.Mdb|"
        .FilterIndex = 1
        .Flags = cdlOFNHideReadOnly
        .FileName = mcDatabaseName
        .ShowOpen
        
        If Err = cdlCancel Then
            Exit Sub
        End If
        
        If .FileName <> "" Then
            mcDatabaseName = .FileName
        End If
    End With
    On Local Error GoTo 0
    
    'No (valid) name?  Then exit
    If mcDatabaseName = "" Then
        MsgBox Translate("No (valid) database name.", mcLanguage), vbCritical
        Unload Me
        End
    Else
        WriteIniFile gcIniHorseFile, "Database", "Folder", mcDatabaseName
        WriteIniFile gcIniFile, "Html Files", "Folder", Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Html\"
        WriteIniFile gcIniFile, "Rtf Files", "Folder", Left$(mcDatabaseName, InStrRev(mcDatabaseName, "\")) & "Rtf\"
        
        OpenDatabase mcDatabaseName
        
        SetVariable "Programversion", ""
        If NewDatabaseName = "" Then
            mdbMain.Close
            RestartApp
        End If
    End If
    
End Sub

Function CompressDatabase() As Integer

    Dim rst As DAO.Recordset
    Dim cOldId As String
    
    SetMouseHourGlass
    
    StatusMessage Translate("Compressing database", mcLanguage)
    
    On Local Error GoTo CompressDatabaseError
    
    CompressDatabase = True
    
    ShowProgressbar Me, 2, 4
    
    OpenDatabase mcDatabaseName
    
    'remove participants with double ID
    cOldId = ""
    Set rst = mdbMain.OpenRecordset("SELECT * FROM Participants ORDER BY Sta")
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            If rst.Fields("Sta") = cOldId Or IsNull(rst.Fields("Sta")) Then
                rst.Delete
            Else
                cOldId = rst.Fields("Sta")
            End If
            rst.MoveNext
        Loop
    End If
    
    IncreaseProgressbarValue Me.ProgressBar1
    
    'remove persons with double ID
    cOldId = ""
    Set rst = mdbMain.OpenRecordset("SELECT * FROM Persons ORDER BY PersonID")
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            If rst.Fields("PersonId") = cOldId Or IsNull(rst.Fields("PersonId")) Then
                rst.Delete
            Else
                cOldId = rst.Fields("PersonId")
            End If
            rst.MoveNext
        Loop
    End If

    IncreaseProgressbarValue Me.ProgressBar1
    
    'remove horses with double ID
    cOldId = ""
    Set rst = mdbMain.OpenRecordset("SELECT * FROM Horses ORDER BY HorseID")
    If rst.RecordCount > 0 Then
        Do While Not rst.EOF
            If rst.Fields("HorseId") = cOldId Or IsNull(rst.Fields("HorseId")) Then
                rst.Delete
            Else
                cOldId = rst.Fields("HorseId") & ""
            End If
            rst.MoveNext
        Loop
    End If
    rst.Close
    Set rst = Nothing
    
    mdbMain.Close
    Set mdbMain = Nothing
    
    DoEvents
    
    IncreaseProgressbarValue Me.ProgressBar1
    
    DoEvents
    
    '+++ MM: 2018-04-02: skip this one as it doesn't work properly
    '+++ CompressDatabase = DatabaseCompact(mcDatabaseName)
    
    IncreaseProgressbarValue Me.ProgressBar1

CompressDatabaseError:
    If Err > 0 Then
        CompressDatabase = False
        MsgBox Err.Description, vbCritical
    End If
    
    ShowProgressbar Me, 2, 0
    
    SetMouseNormal
    
    StatusMessage
    
    RestartApp
    
End Function
Sub RestartApp()
      MsgBox Translate("It is required to re-start this program in order to adapt the program to the required changes.", mcLanguage)
        Unload Me
        If Dir$(App.Path & "\" & App.EXEName & ".Exe") <> "" Then
           Shell App.Path & "\" & App.EXEName & ".Exe Test=" & Me.TestCode & " Status=" & Me.TestStatus, vbNormalNoFocus
        End If
        End
End Sub
Sub PrintRtfHeader(Optional cVersion As String = "", Optional iEmpty As Integer = True, Optional iSuppressTest As Integer = False, Optional cClass As String = "")
    Dim cTemp As String
    Dim cSponsor As String
    
    With rtfResult
        If iEmpty = True Then
            .Text = ""
        End If
        
        'name of event
        .SelBold = True
        .SelFontSize = 18
        .SelText = EventName & vbCrLf
        
        'test code-name-selfin
        If iSuppressTest = False Then
            .SelText = vbCrLf
            .SelFontSize = 14
            .SelBold = True
            .SelText = dtaTest.Recordset.Fields("Code") & " " & Translate(dtaTest.Recordset.Fields("Test"), mcLanguage)
            cSponsor = GetSponsor(dtaTest.Recordset.Fields("Code"))
            If cSponsor <> "" Then
                .SelText = " " & GetVariable("Sponsors") & " " & cSponsor & vbCrLf
            Else
                .SelText = " - "
            End If
            If dtaTest.Recordset.Fields("Type_pre") <= 2 And Translate(dtaTest.Recordset.Fields("Test"), mcLanguage) <> ClipAmp(tbsSelFin.SelectedItem.Caption) Then
                .SelFontSize = 14
                .SelBold = True
                .SelText = ClipAmp(tbsSelFin.SelectedItem.Caption)
            End If
            If cClass <> "" Then
                .SelFontSize = 14
                .SelBold = True
                .SelText = " - " & cClass
            End If
            .SelBold = False
        End If
        If cVersion <> "" Then
            .SelItalic = True
            .SelFontSize = 12
            .SelText = vbCrLf & cVersion
            .SelItalic = False
        End If
        
        .SelText = vbCrLf
        If iSuppressTest = False Then
            .SelFontSize = 10
            .SelBold = True
            .SelAlignment = rtfRight
            .SelText = Translate("Qualification", mcLanguage) & ": " & Translate(dtaTest.Recordset.Fields("Qualification") & "", mcLanguage) & vbCrLf
            .SelAlignment = rtfLeft
            .SelBold = False
            .SelItalic = False
        End If
        .SelBold = False
        .SelItalic = False
        .SelUnderline = False
        .SelBold = False
        .SelFontSize = 12
        .SelText = vbCrLf
    End With
End Sub

Sub PrintRtfFooter(Optional cVersion As String = "", Optional cFormName As String = "", Optional iPagenum As Integer = 0, Optional iEmpty As Integer = 0)
    Dim cTemp As String
    Dim cTemp2 As String
    Dim lngTemp As Long
    Dim iKey As Integer
    Dim iFileNum As Integer
    Dim iTemp As Integer
    Dim strPar As String
    
    On Local Error Resume Next
    
    If iEmpty = 0 Then
        MakeRtfFooter
    End If
    
    If Left$(cFormName, 1) <> "_" Then
        Me.Tempvar = True
    End If
    
    If InStr(cVersion, "[") > 0 Then
        cTemp2 = RTrim$(Left$(cVersion, InStr(cVersion, "[") - 1))
    Else
        cTemp2 = cVersion
    End If

    ReadIniFile gcIniFile, "Print", cTemp2, cTemp
    If Val(cTemp) < 1 Then
        cTemp = "1"
        WriteIniFile gcIniFile, "Print", cTemp2, cTemp
    End If
    frmPrint.txtCounter.Text = cTemp
    If cVersion <> "" Then
        frmPrint.Caption = cVersion
    End If
    
    For iTemp = 1 To 50
        strPar = strPar & "\par "
    Next iTemp
    
    frmPrint.rtfPrint.TextRTF = Replace(rtfResult.TextRTF, "$#@!", strPar)
    frmPrint.Show 1, Me
    
    If Me.Tempvar = "Print" Or Me.Tempvar = "Preview" Then
        If Dir$(mcRtfDir, vbDirectory) = "" Then
            MkDir mcRtfDir
        End If
        
        If cFormName <> "" Then
            cTemp = mcRtfDir & cFormName
        Else
            cTemp = mcRtfDir & dtaTest.Recordset.Fields("Code") & IIf(dtaTest.Recordset.Fields("Type_pre") <= 2, "-" & ClipAmp(tbsSelFin.SelectedItem.Caption), "") & IIf(cVersion <> "", "-" & cVersion, "")
        End If
        
        cTemp = cTemp & ".Rtf"
        lngTemp = KillFile(cTemp)
        Printer.PaperSize = vbPRPSA4
        rtfResult.SaveFile cTemp
        
        iFileNum = FreeFile
        Open cTemp For Binary Access Read Write Shared As #iFileNum
        cTemp2 = Space$(LOF(iFileNum))
        Get #iFileNum, 1, cTemp2
        iTemp = InStr(cTemp2, ";}}")
        If iTemp > 0 Then
            cTemp2 = Left$(cTemp2, iTemp + 2) & "\paperw11907\paperh16840\margl1418\margr704\margt1418\margb1418" & Mid$(RTrim$(cTemp2 & " "), iTemp + 3)
        End If
        cTemp2 = Replace(cTemp2, "$#@!", "{\f1\fs20\page}")
        Put #iFileNum, 1, cTemp2
        Close #iFileNum
        
        DoEvents
        
        If Me.Tempvar = "Preview" Then
            If ShowDocument(cTemp, Me) = 3 Then
                MsgBox Translate("Please install editor for RTF-files first (like MS Word).", mcLanguage), vbExclamation
            End If
        End If
    End If
    
    Me.Tempvar = ""
    SetMouseNormal
    
End Sub

Private Sub txtPrevious_DblClick()
    If txtParticipant.Text = "" Then
        Me.txtParticipant.Text = txtPrevious.Tag
    End If
End Sub

Private Sub txtPrevious_GotFocus()
    txtPrevious.BackColor = vbYellow
End Sub

Private Sub txtPrevious_LostFocus()
    txtPrevious.BackColor = vbWhite
End Sub
Private Sub txtScore_GotFocus()
    If GetKeyState(vbKeyTab) < 0 Then
        cmdOkClick
        dblstNotYet_DblClick
    End If

End Sub

Private Sub txtScore_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyUp Then
       KeyCode = 0
       If fraMarks.Visible = True Then
            SetFocusTo txtMarks(TestJudges - 1)
        ElseIf fraTime.Visible = True Then
            SetFocusTo txtTime
        End If
    ElseIf KeyCode = vbKeyTab Then
        cmdOkClick
    End If

End Sub

Private Sub txtScore_LostFocus()
    Set mctlActive = txtScore
End Sub
Sub PrintOverview(Optional iFullList As Integer = 0)
    Dim rstOverview As DAO.Recordset
    Dim rstDisq As DAO.Recordset
    Dim rstComb As DAO.Recordset
    Dim rstCombList As DAO.Recordset
    Dim rstTestSections As DAO.Recordset
    Dim rstMarks As DAO.Recordset
    
    Dim cQry As String
    Dim cOldCode As String
    Dim iOldStatus As Integer
    Dim cTemp As String
    Dim iPosition As Integer
    Dim cPosition As String
    Dim curOldResult As Currency
    Dim cOldAllTimes As String
    Dim cTotalFormat As String
    Dim cMarkFormat As String
    Dim cTimeFormat As String
    Dim cTableName As String
    Dim iDoNotPrintYet As Integer
    Dim iDoNotPrintNow As Integer
    Dim iSectioncount As Integer
    
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    
    '*** iFulllist = 0 --> no marks
    '*** iFulllist = 1 --> judges' marks
    '*** iFulllist = 2 --> finals only (for prize giving)
    
    SetMouseHourGlass
    
    With rtfResult
        .Text = ""
        .SelBold = True
        .SelFontSize = 18
        .SelText = EventName & vbCrLf
        .SelBold = True
        .SelFontSize = 18
        .SelText = Translate("Overview", mcLanguage) & vbCrLf
        .SelBold = False
        .SelFontSize = 11
    End With
    
    'non-races
    iOldStatus = -1
    
    cQry = "SELECT Tests.Code,Tests.Test,Tests.Type_Final,Tests.Type_Special,"
    cQry = cQry & " Results.Status,Results.Position,"
    cQry = cQry & " Participants.STA,"
    cQry = cQry & " Persons.Name_First & ' ' & Persons.Name_Last AS Name_Rider,"
    cQry = cQry & " Participants.Club,Participants.Team,Participants.Class,"
    cQry = cQry & " Horses.Name_Horse,Horses.HorseID,Horses.FEIFID,"
    cQry = cQry & " Results.Score,Results.AllTimes "
    cQry = cQry & " FROM ((((Results "
    cQry = cQry & " INNER JOIN Tests ON Results.Code = Tests.Code) "
    cQry = cQry & " INNER JOIN TestInfo ON Results.Code = TestInfo.Code) "
    cQry = cQry & " INNER JOIN Participants ON Results.STA = Participants.STA) "
    cQry = cQry & " INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) "
    cQry = cQry & " INNER JOIN Horses ON Participants.HorseID = Horses.HorseID "
    cQry = cQry & " Where Tests.Type_Time <> 1 "
    cQry = cQry & " AND Tests.Type_Time <> 2 "
    cQry = cQry & " AND Results.Score>0 "
    If iFullList = 2 Then
        cQry = cQry & " AND (Results.Status<>0 OR Tests.Type_Special=2 OR NOT Tests.Code IN (SELECT Code FROM Results WHERE STATUS>0)) "
    End If
    cQry = cQry & " AND Results.Disq>-1 "
    cQry = cQry & " ORDER BY Tests.Type_Special,"
    If mnuTestAll.Checked = False Then
        cQry = cQry & " TestInfo.Nr,"
    End If
    cQry = cQry & " Tests.Code,"
    cQry = cQry & " Results.Status=1,Results.Status=2,Results.Status=0,Results.Position;"
   
    Set rstOverview = mdbMain.OpenRecordset(cQry)
    If rstOverview.RecordCount > 0 Then
        GoSub SetTabsRider
        curOldResult = 0
        cOldAllTimes = ""
        Do While Not rstOverview.EOF
            With rtfResult
                If rstOverview.Fields("Code") <> cOldCode Or rstOverview.Fields("Status") <> iOldStatus Then
                    iDoNotPrintYet = True
                    iDoNotPrintNow = False
                    If rstOverview.Fields("Type_Special") = 3 Then 'gaedingakeppni
                        cTotalFormat = "#0.000"
                        cMarkFormat = "#0.00"
                        cTimeFormat = "#0.00"
                    Else
                        cTotalFormat = "#0.00"
                        cMarkFormat = "#0.0"
                        cTimeFormat = "#0.00"
                    End If
                    Select Case GetTestStatus(rstOverview.Fields("Code"))
                    Case 1
                        If rstOverview.Fields("Status") = 0 Then
                            iDoNotPrintYet = False
                        End If
                    Case 2
                        If rstOverview.Fields("Status") <> 1 Then
                            iDoNotPrintYet = False
                        End If
                    Case 3
                        iDoNotPrintYet = False
                    End Select
                    cOldCode = rstOverview.Fields("Code")
                    iOldStatus = rstOverview.Fields("Status")
                    If iDoNotPrintYet = False Then
                        .SelText = vbCrLf & vbCrLf
                        .SelBold = True
                        .SelFontSize = 12
                        .SelText = cOldCode & " - " & Translate(rstOverview.Fields("Test"), mcLanguage) & IIf(rstOverview.Fields("Type_final") <> 0, " - " & IIf(iOldStatus = 0, Translate("Preliminary Round", mcLanguage), IIf(iOldStatus = 1, Translate("A-Final", mcLanguage), IIf(iOldStatus = 2, Translate("B-Final", mcLanguage), Translate("C-Final", mcLanguage)))), "") & vbCrLf & vbCrLf
                        .SelBold = False
                        If iOldStatus >= 1 Then
                            iPosition = GetHighestPosition(cOldCode, iOldStatus) - 1
                        Else
                            iPosition = 0
                        End If
                    End If
                    If iFullList = 1 Then
                        Set rstTestSections = mdbMain.OpenRecordset("SELECT * FROM TestSections WHERE Code='" & cOldCode & "' AND Status=" & IIf(iOldStatus = 2, 1, Format$(iOldStatus)) & " ORDER BY Section")
                        If rstTestSections.RecordCount > 0 Then
                            rstTestSections.MoveLast
                            iSectioncount = rstTestSections.RecordCount
                        Else
                            iSectioncount = 0
                        End If
                    End If
                End If
                If iDoNotPrintYet = False Then
                    Set rstDisq = mdbMain.OpenRecordset("SELECT * FROM Results WHERE STA='" & rstOverview.Fields("STA") & "' AND Code='" & cOldCode & "' And Disq = -1")
                    If rstDisq.RecordCount = 0 Then
                        GoSub SetTabsRider
                        iPosition = iPosition + 1
                        If rstOverview.Fields("Score") <> curOldResult Or rstOverview.Fields("AllTimes") & "" <> cOldAllTimes Then
                            cPosition = Format$(rstOverview.Fields("Position"), "00")
                            If iFullList = 2 And Val(cPosition) > 10 And (rstOverview.Fields("Type_Special") = 2 Or rstOverview.Fields("Status") = 0) Then
                                iDoNotPrintNow = True
                            End If
                        End If
                        If iDoNotPrintNow = False Then
                            .SelText = cPosition & vbTab & rstOverview.Fields("STA") & vbTab & rstOverview.Fields("Name_rider")
                            If rstOverview.Fields("Class") & "" <> "" Then
                                .SelText = " [" & rstOverview.Fields("Class") & "]"
                            End If
                            .SelText = " / " & rstOverview.Fields("Name_Horse")
                            If miShowHorseId <> 0 Then
                                .SelText = " [" & GetHorseId(rstOverview) & "]"
                            End If
                            If miShowRidersClub <> 0 Then
                                .SelText = " / " & GetRidersClub(rstOverview)
                            End If
                            If miShowRidersTeam <> 0 Then
                                .SelText = " / " & Left$(GetRidersTeam(rstOverview), 2)
                            End If
                            'IPZV LK output:
                            If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                                .SelText = " / " & Left$(GetRidersLk(rstOverview), 2)
                            End If
                            
                            .SelText = vbTab & Format$(rstOverview.Fields("Score"), cTotalFormat) & vbCrLf
                            curOldResult = rstOverview.Fields("Score")
                            cOldAllTimes = rstOverview.Fields("AllTimes") & ""
                            If iSectioncount > 0 Then
                                GoSub SetTabsMarks
                                rstTestSections.MoveFirst
                                Do While Not rstTestSections.EOF
                                    Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE STA='" & rstOverview.Fields("STA") & "' AND Code='" & cOldCode & "' AND Status=" & rstOverview.Fields("Status") & " AND Section=" & rstTestSections.AbsolutePosition + 1)
                                    If rstMarks.RecordCount > 0 Then
                                        .SelItalic = True
                                        .SelText = vbTab & vbTab
                                        If iSectioncount > 1 Then
                                            iTemp = InStr(rstTestSections.Fields("Name"), ",")
                                            If iTemp > 0 Then
                                                .SelText = Left$(rstTestSections.Fields("Name"), iTemp - 1)
                                            Else
                                                .SelText = rstTestSections.Fields("Name")
                                            End If
                                        End If
                                        .SelText = vbTab
                                        .SelText = Format$(rstMarks.Fields("Mark1"), cMarkFormat)
                                        .SelText = " - " & Format$(rstMarks.Fields("Mark2"), cMarkFormat)
                                        .SelText = " - " & Format$(rstMarks.Fields("Mark3"), cMarkFormat)
                                        .SelText = " - " & Format$(rstMarks.Fields("Mark4"), cMarkFormat)
                                        .SelText = " - " & Format$(rstMarks.Fields("Mark5"), cMarkFormat)
                                        If rstOverview.Fields("Type_Special") = 2 Then
                                            .SelText = """"
                                        End If
                                        .SelText = vbTab & Format$(rstMarks.Fields("Score"), cTotalFormat)
                                        .SelText = vbCrLf
                                        .SelItalic = False
                                    End If
                                    rstMarks.Close
                                    rstTestSections.MoveNext
                                Loop
                            End If
                        End If
                    End If
                    rstDisq.Close
                End If
                rstOverview.MoveNext
            End With
        Loop
        If iFullList = 1 Then
            rstTestSections.Close
        End If
    End If
    rstOverview.Close
    
    'races
    cQry = "SELECT Tests.Code,Tests.Test,Tests.Type_time,Tests.Status,"
    cQry = cQry & " Participants.STA,Participants.Club,Participants.Team,Participants.Class,"
    cQry = cQry & " Persons.Name_First & ' ' &  Persons.Name_Last AS Name_Rider,"
    cQry = cQry & " Horses.Name_Horse,Horses.HorseID,Horses.FEIFID,"
    cQry = cQry & " Results.Score,Results.Alltimes,Results.Position"
    cQry = cQry & " FROM (((Results INNER JOIN Tests ON (Results.Code = Tests.Code) AND (Results.Status = Tests.Status)) "
    cQry = cQry & " INNER JOIN Participants ON Results.STA = Participants.STA) "
    cQry = cQry & " INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) "
    cQry = cQry & " INNER JOIN Horses ON Participants.HorseID = Horses.HorseID "
    cQry = cQry & " WHERE (Tests.Type_time=1 OR Tests.Type_time=2) "
    cQry = cQry & " AND Results.Score>0 AND Results.Disq>-1 "
    cQry = cQry & " ORDER BY Tests.Code,Tests.Status,Results.Position;"
    
    Set rstOverview = mdbMain.OpenRecordset(cQry)
    If rstOverview.RecordCount > 0 Then
        Do While Not rstOverview.EOF
            With rtfResult
                curOldResult = 0
                cOldAllTimes = ""
                If rstOverview.Fields("Code") <> cOldCode Or rstOverview.Fields("Status") <> iOldStatus Then
                    iDoNotPrintYet = True
                    Select Case GetTestStatus(rstOverview.Fields("Code"))
                    Case 1
                        If TestStatus = 0 Then
                            iDoNotPrintYet = False
                        End If
                    Case 2
                        If TestStatus <> 1 Then
                            iDoNotPrintYet = False
                        End If
                    Case 3
                        iDoNotPrintYet = False
                    End Select
                    cOldCode = rstOverview.Fields("Code")
                    iOldStatus = rstOverview.Fields("Status")
                    If iDoNotPrintYet = False Then
                        iDoNotPrintNow = False
                        iDoNotPrintYet = False
                        .SelText = vbCrLf & vbCrLf
                        .SelBold = True
                        .SelFontSize = 12
                        .SelText = rstOverview.Fields("Code") & " - " & rstOverview.Fields("Test") & vbCrLf & vbCrLf
                        .SelBold = False
                        If iOldStatus > 1 Then
                            iPosition = GetHighestPosition(cOldCode, iOldStatus) - 1
                        Else
                            iPosition = 0
                        End If
                    End If
                    If iFullList = 1 Then
                        Set rstTestSections = mdbMain.OpenRecordset("SELECT * FROM TestSections WHERE Code='" & cOldCode & "' AND Status=" & IIf(iOldStatus = 2, 1, Format$(iOldStatus)))
                        If rstTestSections.RecordCount > 0 Then
                            rstTestSections.MoveLast
                            iSectioncount = rstTestSections.RecordCount
                        Else
                            iSectioncount = 0
                        End If
                    End If
                End If
                If iDoNotPrintYet = False Then
                    iPosition = iPosition + 1
                    GoSub SetTabsRider
                    If rstOverview.Fields("Score") <> curOldResult Or rstOverview.Fields("AllTimes") & "" <> cOldAllTimes Then
                        cPosition = Format$(rstOverview.Fields("Position"), "00")
                        If iFullList = 2 And Val(cPosition) > 10 Then
                            iDoNotPrintNow = True
                        End If
                    End If
                    If iDoNotPrintNow = False Then
                        .SelText = cPosition & vbTab & rstOverview.Fields("STA") & vbTab & rstOverview.Fields("Name_rider")
                        If rstOverview.Fields("Class") & "" <> "" Then
                            .SelText = " [" & rstOverview.Fields("Class") & "]"
                        End If
                        .SelText = " / " & rstOverview.Fields("Name_Horse")
                        If miShowHorseId <> 0 Then
                            .SelText = " [" & GetHorseId(rstOverview) & "]"
                        End If
                        If miShowRidersClub <> 0 Then
                            .SelText = " / " & GetRidersClub(rstOverview)
                        End If
                        If miShowRidersTeam <> 0 Then
                            .SelText = " / " & Left$(GetRidersTeam(rstOverview), 2)
                        End If
                        'IPZV LK output:
                        If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                            .SelText = " / " & Left$(GetRidersLk(rstOverview), 2)
                        End If
                        If cPosition = mcNoPosition Then
                            .SelText = vbTab & mcNoPosition & vbCrLf
                        Else
                            .SelText = vbTab & Format$(rstOverview.Fields("Score"), cTotalFormat) & Chr$(34) & vbCrLf
                        End If
                        curOldResult = rstOverview.Fields("Score")
                        cOldAllTimes = rstOverview.Fields("AllTimes") & ""
                        If iSectioncount > 0 Then
                            GoSub SetTabsMarks
                            rstTestSections.MoveFirst
                            Do While Not rstTestSections.EOF
                                Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE STA='" & rstOverview.Fields("STA") & "' AND Code='" & cOldCode & "' AND Status=" & rstOverview.Fields("Status") & " AND Section=" & rstTestSections.AbsolutePosition + 1)
                                If rstMarks.RecordCount > 0 Then
                                    .SelItalic = True
                                    .SelText = vbTab & vbTab
                                    If iSectioncount > 1 Then
                                        iTemp = InStr(rstTestSections.Fields("Name"), ",")
                                        If iTemp > 0 Then
                                            .SelText = Left$(rstTestSections.Fields("Name"), iTemp - 1)
                                        Else
                                            .SelText = rstTestSections.Fields("Name")
                                        End If
                                    End If
                                    .SelText = vbTab
                                    .SelText = Format$(rstMarks.Fields("Mark1"), cTimeFormat)
                                    .SelText = """"
                                    .SelText = vbCrLf
                                    .SelItalic = False
                                End If
                                rstMarks.Close
                                rstTestSections.MoveNext
                            Loop
                        End If
                    End If
                End If
                rstOverview.MoveNext
            End With
        Loop
        If iFullList = 1 Then
            rstTestSections.Close
        End If
    End If
    rstOverview.Close
            
    'combination winners
    If TableExist(mdbMain, "Combinations") = True Then
        Set rstCombList = mdbMain.OpenRecordset("SELECT DISTINCT Combination,Code FROM Combinations ORDER BY Combination")
        If rstCombList.RecordCount > 0 Then
            cTableName = "_Temp-" & MachineName
            Do While Not rstCombList.EOF
                If rstCombList.Fields("Code") & "" <> "" And rstCombList.Fields("Code") & "" <> "Club" And rstCombList.Fields("Code") & "" <> "Team" Then
                    If GetCombinationStatus(rstCombList.Fields("Code")) = 1 Then
                        CalculateCombination rstCombList.Fields("Code"), False
                        iDoNotPrintNow = False
                        cQry = "SELECT Participants.STA, [" & cTableName & "].*,"
                        cQry = cQry & " Horses.Name_Horse,Horses.HorseID,Horses.FEIFID,"
                        cQry = cQry & " Persons.Name_First & ' ' & Persons.Name_Last AS Name_rider,"
                        cQry = cQry & " Participants.Club,Participants.Team,Participants.Class "
                        cQry = cQry & " FROM (([" & cTableName & "] "
                        cQry = cQry & " INNER JOIN Participants ON [" & cTableName & "].STA = Participants.STA) "
                        cQry = cQry & " INNER JOIN Horses ON Participants.HorseID = Horses.HorseID) "
                        cQry = cQry & " INNER JOIN Persons ON Participants.PersonID = Persons.PersonID "
                        cQry = cQry & " WHERE Score>0 "
                        cQry = cQry & " ORDER BY Score DESC"
                        Set rstOverview = mdbMain.OpenRecordset(cQry)
                        If rstOverview.RecordCount > 0 Then
                            curOldResult = 0
                            With rtfResult
                                .SelText = vbCrLf & vbCrLf
                                .SelBold = True
                                .SelFontSize = 12
                                .SelText = Translate("Combination", mcLanguage) & ": " & Translate(rstCombList.Fields("Combination"), mcLanguage) & vbCrLf & vbCrLf
                                .SelBold = False
                                iPosition = 0
                                Do While Not rstOverview.EOF
                                    GoSub SetTabsRider
                                    iPosition = iPosition + 1
                                    If rstOverview.Fields("Score") <> curOldResult Then
                                        cPosition = Format$(iPosition, "00")
                                        If iFullList = 2 And Val(cPosition) > 10 Then
                                            iDoNotPrintNow = True
                                        End If
                                    End If
                                    If iDoNotPrintNow = False Then
                                        .SelText = cPosition & vbTab & rstOverview.Fields("Participants.STA") & vbTab & rstOverview.Fields("Name_rider")
                                        If rstOverview.Fields("Class") & "" <> "" Then
                                            .SelText = " [" & rstOverview.Fields("Class") & "]"
                                        End If
                                        .SelText = " / " & rstOverview.Fields("Name_Horse")
                                        If miShowHorseId <> 0 Then
                                            .SelText = " [" & GetHorseId(rstOverview) & "]"
                                        End If
                                        If miShowRidersClub <> 0 Then
                                            .SelText = " / " & GetRidersClub(rstOverview)
                                        End If
                                        If miShowRidersTeam <> 0 Then
                                            .SelText = " / " & Left$(GetRidersTeam(rstOverview), 2)
                                        End If
                                        'IPZV LK output:
                                        If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                                            .SelText = " / " & Left$(GetRidersLk(rstOverview), 2)
                                        End If
                                        .SelText = vbTab & Format$(rstOverview.Fields("Score"), cTotalFormat) & vbCrLf
                                    End If
                                    curOldResult = rstOverview.Fields("Score")
                                    rstOverview.MoveNext
                                Loop
                            End With
                        End If
                        rstOverview.Close
                    End If
                End If
                rstCombList.MoveNext
            Loop
        End If
        rstCombList.Close
        Set rstCombList = Nothing
    End If
    
    Set rstOverview = Nothing
    
    If iFullList = 1 Then
        Set rstTestSections = Nothing
    End If
    
    SetMouseNormal
    
    PrintRtfFooter Translate("Overview", mcLanguage), NameOfFile(Dir$(mcDatabaseName))
    
Exit Sub

SetTabsRider:
    With rtfResult
        .SelFontSize = 11
        .SelBold = True
        .SelTabCount = 4
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 2.5 * 567
        .SelTabs(2) = 15 * 567
        .SelTabs(3) = 16 * 567
    End With
Return

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelBold = False
        .SelTabCount = 4
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 3 * 567
        .SelTabs(2) = 8 * 567
        .SelTabs(3) = 13 * 567
    End With
Return

End Sub
Sub PrintParticipant()
    Dim rstMarks As DAO.Recordset
    Dim rstTest As DAO.Recordset
    Dim rstParticipant As DAO.Recordset
    
    Dim cParticipant As String
    Dim cHorse As String
    Dim cQry As String
    Dim cSta As String
    
    Dim cOldCode As String
    Dim iOldStatus As Integer
    Dim iOldSection As Integer
    Dim iTemp As Integer
    Dim iPage  As Integer
    Dim cTemp As String
    Dim curOldResult As Currency
    Dim cOldAllTimes As String
    Dim cStaList As String
    
    Dim cMarkFormat As String
    Dim cTimeFormat As String
    Dim cTotalFormat As String
    Dim iTimeDecimals As Integer
    Dim iMarkDecimals As Integer
    
    iOldStatus = -1
    iOldSection = -1
    
   cTemp = InputBox$(Translate("Search for", mcLanguage), Translate("Results per participant", mcLanguage), mnuEditAdd.Tag)
   
   cQry = "SELECT Participants.Sta "
   cQry = cQry & " & '  -  ' & Persons.Name_First"
   cQry = cQry & " & ' ' & Persons.Name_Last"
   cQry = cQry & " & IIF(Participants.Class<>'',' ['  & participants.Class & ']','') "
   cQry = cQry & " & ' - ' & Horses.Name_Horse as cList"
   cQry = cQry & " FROM (Participants"
   cQry = cQry & " INNER JOIN Persons"
   cQry = cQry & " ON Participants.PersonId=Persons.PersonId)"
   cQry = cQry & " INNER JOIN Horses"
   cQry = cQry & " ON Participants.HorseId=Horses.HorseId"
   cQry = cQry & " WHERE Participants.Sta & ' ' & Persons.Name_First & ' ' & Persons.Name_Last & ' - ' & Horses.Name_Horse LIKE " & Chr$(34) & "*" & cTemp & "*" & Chr$(34)
   cQry = cQry & " ORDER BY Participants.Sta"
   
   frmToolBox.strQry = cQry
   frmToolBox.intChecked = True
   frmToolBox.intReturnLen = 3
   frmToolBox.Caption = Translate("Searching", mcLanguage) & " '" & cTemp & "' "
   frmToolBox.Show 1, Me
   
   cStaList = Me.Tempvar
   Me.Tempvar = ""
   
   rtfResult.Text = ""
   
   Do While cStaList <> ""
        Parse cSta, cStaList, "|"
        cSta = Format$(Val(Left$(cSta, 3)), "000")
        
        cQry = "SELECT Results.STA,Results.Code, Results.Status,Results.Disq,Results.Score,Results.Position,Marks.Section, Marks.Mark1, Marks.Mark2, Marks.Mark3, Marks.Mark4, Marks.Mark5, Marks.Score,Tests.Test,TestInfo.Num_j_0,TestInfo.Num_j_1,TestInfo.Num_j_2,TestInfo.Num_j_3"
        cQry = cQry & " FROM ((Results INNER JOIN Marks ON (Results.Status = Marks.Status) AND (Results.STA = Marks.STA) AND (Results.Code = Marks.Code)) INNER JOIN Tests ON Marks.Code = Tests.Code) INNER JOIN TestInfo ON Marks.Code=TestInfo.Code"
        cQry = cQry & " WHERE Results.STA='" & cSta & "'"
        cQry = cQry & " ORDER BY Results.STA, Results.Code, Results.Status=0, Results.Status DESC , Marks.Section;"
    
        Set rstMarks = mdbMain.OpenRecordset(cQry)
        If rstMarks.RecordCount > 0 Then
            Set rstParticipant = mdbMain.OpenRecordset("SELECT Participants.STA, Participants.Class, Persons.Name_First,Persons.Name_Last, Horses.Name_Horse, Horses.FEIFID FROM (Participants INNER JOIN Horses ON Participants.HorseID = Horses.HorseID) INNER JOIN Persons ON Participants.PersonID = Persons.PersonID WHERE STA='" & cSta & "';")
    
            cParticipant = rstParticipant.Fields("Name_first") & " " & rstParticipant.Fields("Name_Last")
            If rstParticipant.Fields("Class") & "" <> "" Then
                cParticipant = cParticipant & " [" & rstParticipant.Fields("Class") & "]"
            End If
            cHorse = rstParticipant.Fields("Name_horse") & " [" & GetHorseId(rstParticipant) & "]"
            
            If iPage = True Then
                MakeRtfFooter
                rtfResult.SelText = "$#@!"
            End If
            
            With rtfResult
                .SelBold = True
                .SelFontSize = 18
                .SelText = EventName & vbCrLf
                .SelBold = True
                .SelFontSize = 14
                .SelText = vbCrLf & cSta & vbCrLf
                .SelBold = True
                .SelFontSize = 14
                .SelText = cParticipant & vbCrLf
                .SelBold = True
                .SelFontSize = 14
                .SelText = cHorse & vbCrLf
                .SelBold = False
            End With
            
            PrintRtfLine
            
            GoSub SetTabsRider
            
            Do While Not rstMarks.EOF
                With rtfResult
                    If rstMarks.Fields("Code") <> cOldCode Or rstMarks.Fields("Status") <> iOldStatus Or iOldSection <> rstMarks.Fields("Section") Then
                        .SelFontSize = 11
                        cOldCode = rstMarks.Fields("Code")
                        iOldStatus = rstMarks.Fields("Status")
                        If rstMarks.Fields("Section") = 1 Then
                            .SelText = vbCrLf
                            .SelBold = True
                            .SelText = cOldCode & " - " & Translate(rstMarks.Fields("Test"), mcLanguage) & IIf(iOldStatus = 0, "", IIf(iOldStatus = 1, " - " & Translate("A-Final", mcLanguage), IIf(iOldStatus = 2, " - " & Translate("B-Final", mcLanguage), Translate("C-Final", mcLanguage)))) & vbCrLf
                            If Val(rstMarks("Position") & "") > 0 Then
                                .SelText = Translate("Position", mcLanguage) & ": " & rstMarks("Position") & vbCrLf
                            End If
                            .SelText = vbCrLf
                            .SelBold = False
                        End If
                        iOldSection = rstMarks.Fields("Section")
                        Set rstTest = mdbMain.OpenRecordset("SELECT * FROM Tests INNER JOIN Testsections ON Tests.Code=Testsections.Code WHERE Tests.Code='" & cOldCode & "' AND TestSections.Section=" & iOldSection & " AND TestSections.Status=" & IIf(iOldStatus = 0, 0, 1))
                    End If
                    .SelFontSize = 11
                    .SelText = vbTab & Left$(Translate(rstTest.Fields("Name"), mcLanguage), 20) & vbTab
                    
                    iMarkDecimals = 1
                    iTimeDecimals = 1
                    Select Case rstTest.Fields("Type_Pre")
                    Case Is <= 2 'marks or placemarks
                        With rstTest
                            If IsNull(.Fields("Mark_Decimals")) Then
                                iMarkDecimals = 1
                            Else
                                iMarkDecimals = .Fields("Mark_Decimals")
                            End If
                        End With
                    Case Is = 3  'time
                        With rstTest
                            If IsNull(.Fields("Time_Decimals")) Then
                                iTimeDecimals = 1
                            Else
                                iTimeDecimals = .Fields("Time_Decimals")
                            End If
                        End With
                    Case Else
                    End Select
                    
                    'how to format marks
                    If ((iOldStatus = 0 And rstTest.Fields("Type_pre") = 2) Or (iOldStatus <> 0 And rstTest.Fields("Type_Final") = 2)) Then
                        cMarkFormat = "0"
                    Else
                        cMarkFormat = "0." & String$(iMarkDecimals, "0")
                    End If
                    cTimeFormat = "0." & String$(iTimeDecimals, "0")
                    If rstTest.Fields("Type_Special") = 3 Then 'gaedingakeppni
                        cTotalFormat = "0.000"
                    Else
                        cTotalFormat = "0.00"
                    End If
                    
                    If rstMarks.Fields("Disq") = -2 Then
                            .SelText = mcNoPosition
                    ElseIf rstTest.Fields("Type_pre") = 3 Then
                        If rstMarks.Fields("Mark1") = 0 Then
                            .SelText = mcNoPosition
                        Else
                            .SelText = Format$(rstMarks.Fields("Mark1"), cTimeFormat) & Chr$(34)
                            .SelText = " (= " & Format$(Time2Mark(rstMarks.Fields("Mark1"), cOldCode), cTotalFormat) & ")"
                        End If
                    Else
                        For iTemp = 1 To rstMarks.Fields("num_j_" & Format$(rstMarks.Fields("Status")))
                            .SelText = Format$(rstMarks.Fields("Mark" & iTemp), cMarkFormat)
                            If iTemp < rstMarks.Fields("num_j_" & Format$(rstMarks.Fields("Status"))) Then
                                .SelText = " - "
                            End If
                        Next iTemp
                        .SelText = " = " & Format$(rstMarks.Fields("Marks.Score"), cTotalFormat)
                    End If
                    If iOldSection = 1 Then
                        If rstMarks.Fields("Disq") = -1 Then
                            .SelText = vbTab & Translate("Eliminated", mcLanguage)
                        ElseIf rstMarks.Fields("Disq") = -2 Then
                            .SelText = vbTab & Translate("Withdrawn", mcLanguage)
                        Else
                            .SelBold = True
                            .SelText = vbTab & Translate("Total", mcLanguage) & ": " & Format$(rstMarks.Fields("Results.Score"), cTotalFormat)
                            If rstTest.Fields("Type_pre") = 3 Then
                                .SelText = Chr$(34)
                            End If
                            .SelBold = False
                        End If
                    End If
                    .SelText = vbCrLf
                    rstMarks.MoveNext
                End With
            Loop
            If cStaList <> "" Then
                iPage = True
            End If
            rstParticipant.Close
        End If
        rstMarks.Close
        If cStaList = "" Then
            PrintRtfFooter Translate("Overview ", mcLanguage), "_"
            Exit Do
        End If
    Loop
    Set rstMarks = Nothing
    Set rstParticipant = Nothing
Exit Sub

SetTabsRider:
    With rtfResult
        .SelFontSize = 11
        .SelBold = True
        .SelTabCount = 4
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 5 * 567
        .SelTabs(2) = 13 * 567
        .SelTabs(3) = 16 * 567
    End With
Return

End Sub
Sub PrintResultList(cVersion As String, Optional iAllTests As Integer = False, Optional iMarkFinals As Integer = False)
    Dim iKey As Integer
    Dim curOldResult As Currency
    Dim cOldAllTimes As String
    Dim iPosition As Integer
    Dim iHighestPosition As Integer
    Dim cPosition As String
    Dim cScore As String
    Dim cResult As String
    Dim iEmptyline As Integer
    Dim iInWr As Integer
    
    Dim cRider As String
    Dim cHorse As String
    Dim cParticipant As String
    Dim cClass As String
    Dim cClub As String
    Dim cTeam As String
    Dim iTemp As Integer
    Dim iA_Line As Integer
    Dim iB_line As Integer
    Dim iC_line As Integer
    Dim iBold As Integer
    Dim iPrintResult As Integer
    Dim cTemp As String
    Dim curWRLimit As Currency
    Dim cRidersList As String
    Dim cRidersDouble As String
    Dim cSplitCode As String
    Dim cSplitClass As String
    Dim cOldSplitCode As String
    
    Dim rstJ(5) As DAO.Recordset
    Dim rstPos As DAO.Recordset
    Dim rstSplit As DAO.Recordset
    Dim iJ As Integer
    
    Dim cOldParticipant As String
    
    Dim bFlag As Boolean
    Dim writeLogDB As Boolean
    Dim logDisq As Integer, logPos As Integer, logResult As Double
    
    SetMouseHourGlass
    
    LookUpRelevantParticipants
    
    CheckTieBreak True
    
    cOldSplitCode = "|"
    
    If frmMain.chkSplitResultLists = 1 And TestStatus = 0 Then
        '* prepare result list for all classes involved
        '*
        
        Set rstSplit = mdbMain.OpenRecordset("SELECT Class FROM TestSplits WHERE SplitToTest NOT LIKE '" & TestCode & "'")
        If rstSplit.RecordCount > 0 Then
            Do While Not rstSplit.EOF
                cSplitCode = rstSplit.Fields(0)
                GoSub PrintResultList
                cOldSplitCode = cOldSplitCode & cSplitCode & "|"
                rstSplit.MoveNext
            Loop
            cSplitCode = ""
            GoSub PrintResultList
        Else
            cSplitCode = ""
            GoSub PrintResultList
        End If
        rstSplit.Close
    Else
        cSplitCode = ""
        GoSub PrintResultList
    End If
    
    SetMouseNormal
    
Exit Sub

PrintResultList:

    If dtaAlready.Recordset.RecordCount > 0 Then
                
        PrintRtfHeader cVersion, True, False, cSplitCode
        
        If miWriteLogDB Then
            bFlag = DelLogDBConfMarks(EventName, dtaTest.Recordset("code"), TestStatus)
        End If
        
        dtaAlready.Recordset.MoveLast
        dtaAlready.Recordset.MoveFirst
    
        With rtfResult
            .SelTabCount = 1
            .SelTabs(0) = 16 * 567
            .SelUnderline = True
            .SelText = vbTab & vbCrLf
            .SelUnderline = False
            .SelBold = False
            .SelFontSize = 8
            GoSub SetTabsRider
            .SelItalic = True
            .SelBold = False
            .SelFontSize = 8
            
            If miShowRidersClub <> 0 Or miShowRidersTeam <> 0 Then
                If fraTime.Visible = True Then
                    .SelText = "POS" & vbTab & "#" & vbTab & UCase$(Translate("Rider", mcLanguage)) & IIf(miShowRidersClub <> 0, "/" & UCase$(Translate("Club", mcLanguage)), "") & IIf(miShowRidersTeam <> 0, "/" & UCase$(Translate("Team", mcLanguage)), "") & vbTab & vbCrLf
                    .SelItalic = True
                    .SelBold = False
                    .SelFontSize = 8
                    .SelText = "" & vbTab & "" & vbTab & UCase$(Translate("Horse", mcLanguage)) & vbTab & UCase$(Translate("MARK", mcLanguage)) & vbTab & UCase$(Translate("TIME", mcLanguage)) & vbCrLf
                Else
                    .SelText = "POS" & vbTab & "#" & vbTab & UCase$(Translate("Rider", mcLanguage)) & IIf(miShowRidersClub <> 0, "/" & UCase$(Translate("Club", mcLanguage)), "") & IIf(miShowRidersTeam <> 0, "/" & UCase$(Translate("Team", mcLanguage)), "") & vbTab & vbTab & vbCrLf
                    .SelItalic = True
                    .SelBold = False
                    .SelFontSize = 8
                    .SelText = "" & vbTab & "" & vbTab & UCase$(Translate("Horse", mcLanguage)) & vbTab & vbTab & "TOT" & vbCrLf
                End If
            Else
                If fraTime.Visible = True Then
                    .SelText = "POS" & vbTab & "#" & vbTab & UCase$(Translate("Rider", mcLanguage)) & "/" & UCase$(Translate("Horse", mcLanguage)) & vbTab & UCase$(Translate("MARK", mcLanguage)) & vbTab & UCase$(Translate("TIME", mcLanguage)) & vbCrLf
                Else
                    .SelText = "POS" & vbTab & "#" & vbTab & UCase$(Translate("Rider", mcLanguage)) & "/" & UCase$(Translate("Horse", mcLanguage)) & vbTab & vbTab & "TOT" & vbCrLf
                End If
            End If
            GoSub SetTabsMarks
            .SelFontSize = 8
            .SelUnderline = True
            
            If fraMarks.Visible = True Then
                .SelText = vbTab & UCase$(Translate("Judge", mcLanguage)) & vbTab & IIf(dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus)) > 0, "A", "") & vbTab & IIf(dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus)) > 1, "B", "") & vbTab & IIf(dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus)) > 2, "C", "") & vbTab & IIf(dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus)) > 3, "D", "") & vbTab & IIf(dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus)) > 4, IIf(dtaTest.Recordset.Fields("Type_special") = 2, UCase$(Translate("TIME", mcLanguage)), "E"), "") & vbTab & "SUB" & vbTab & vbTab & vbCrLf & vbCrLf
            ElseIf fraTime.Visible = True Then
                .SelText = vbTab & vbTab & UCase$(Left$(Translate("Seconds", mcLanguage), 3)) & vbTab & vbTab & vbCrLf & vbCrLf
            End If
            .SelUnderline = False
            If iMarkFinals = True And miMarkFinalsInResultLists <> 0 Then
                If dtaTestInfo.Recordset.Fields("Handling") = 1 Or dtaTestInfo.Recordset.Fields("Handling") = 2 Or dtaTestInfo.Recordset.Fields("Handling") = 5 Then
                    cRidersList = "|"
                    If tbsSelFin.Tabs.Count > 1 Then
                        .SelFontSize = 11
                        .SelBold = True
                        .SelText = ClipAmp(tbsSelFin.Tabs(tbsSelFin.Tabs.Count).Caption & vbCrLf & vbCrLf)
                    End If
                End If
            End If
        End With
        If TestStatus > 1 Then
            iPosition = GetHighestPosition(TestCode, TestStatus) - 1
            iHighestPosition = iPosition
        Else
            iPosition = 0
            iHighestPosition = 0
        End If
            
        Do While Not dtaAlready.Recordset.EOF
           iPrintResult = 0
           cSplitClass = dtaAlready.Recordset.Fields("Class") & ""
           If cSplitCode <> "" And cSplitClass <> "" And InStr(cSplitCode, cSplitClass) > 0 Then
                iPrintResult = 1
            ElseIf cSplitCode = "" And (InStr(cOldSplitCode, cSplitClass) = 0 Or cSplitClass = "") Then
                iPrintResult = 2
            End If
            
           If iPrintResult > 0 Then
                
                If fraTime.Visible = True Then
                     If dtaAlready.Recordset.Fields("Marks.Score") = 0 Then
                         cScore = mcNoPosition
                     Else
                         cScore = Format$(dtaAlready.Recordset.Fields("Marks.Score"), TestTotalFormat) & Chr$(34)
                     End If
                Else
                     cScore = Format$(dtaAlready.Recordset.Fields("Marks.Score"), TestTotalFormat)
                End If
                
                If dtaAlready.Recordset.Fields("DISQ") = -1 Then
                     cResult = UCase$(Translate("Eliminated", mcLanguage))
                ElseIf dtaAlready.Recordset.Fields("DISQ") = -2 Then
                     cResult = Translate("Withdrawn", mcLanguage)
                Else
                     If fraTime.Visible = True Then
                         If dtaAlready.Recordset.Fields("Results.Score") = 0 Then
                             cResult = mcNoPosition
                         Else
                             cResult = Format$(dtaAlready.Recordset.Fields("Results.Score"), TestTotalFormat) & Chr$(34)
                         End If
                     Else
                         cResult = Format$(dtaAlready.Recordset.Fields("Results.Score"), TestTotalFormat)
                         If dtaAlready.Recordset.Fields("Alltimes") = "1" And dtaAlready.Recordset.Fields("Position") = 1 Then
                             cResult = cResult & " [T]"
                         End If
                     End If
                End If
                
                 If cOldParticipant <> dtaAlready.Recordset.Fields("cList") & "" Then
                      GoSub SetTabsRider
                      If dtaAlready.Recordset.Fields("DISQ") < 0 Then
                          rtfResult.SelBold = False
                      End If
                      rtfResult.SelText = cPosition & vbTab & dtaAlready.Recordset.Fields("Participants.Sta") & vbTab & cParticipant & vbTab
                      If cResult <> Translate("ELIMINATED", mcLanguage) And cResult <> Translate("Withdrawn", mcLanguage) And cResult <> mcNoPosition Then
                         If fraTime.Visible = True Then
                             rtfResult.SelBold = False
                             rtfResult.SelFontSize = rtfResult.SelFontSize - 2
                             rtfResult.SelText = Format$(Time2Mark(dtaAlready.Recordset.Fields("Results.Score"), TestCode), TestTotalFormat)
                             rtfResult.SelFontSize = rtfResult.SelFontSize + 2
                             rtfResult.SelBold = True
                         End If
                          rtfResult.SelText = vbTab
                      End If
                      rtfResult.SelText = cResult & vbCrLf
                      cOldParticipant = dtaAlready.Recordset.Fields("cList") & ""
                 End If
                  
                If dtaAlready.Recordset.Fields("DISQ") <> -2 Or IsNull(dtaAlready.Recordset.Fields("DISQ")) Then
                     GoSub SetTabsMarks
                     
                     If fraMarks.Visible = True Then
                         rtfResult.SelText = vbTab
                         If dtaAlready.Recordset.Fields("Out") <> 0 Then
                             rtfResult.SelStrikeThru = True
                         End If
                         rtfResult.SelText = UCase$(Left$(Translate(dtaAlready.Recordset.Fields("Name"), mcLanguage), 4))
                     ElseIf fraTime.Visible = True And cScore <> mcNoPosition Then
                         rtfResult.SelText = vbTab
                         rtfResult.SelText = UCase$(Translate(dtaAlready.Recordset.Fields("Name"), mcLanguage))
                     End If
                     If fraMarks.Visible = True Then
                         For iTemp = 1 To 5
                             If iTemp <= dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus)) Then
                                 If dtaAlready.Recordset.Fields("Mark" & Format$(iTemp)) & "" <> "" Then
                                     rtfResult.SelText = vbTab & Format$(dtaAlready.Recordset.Fields("Mark" & Format$(iTemp)), Me.TestMarkFormat)
                                     If miShowJudgesRanking <> 0 Then
                                        If frmMain.chkSplitResultLists = 1 And TestStatus = 0 Then
                                            If iPrintResult = 1 Then
                                                Set rstJ(iTemp) = mdbMain.OpenRecordset("SELECT Mark" & Format$(iTemp) & " FROM Marks Where Code='" & TestCode & "' AND Status=" & TestStatus & " AND Section=" & dtaAlready.Recordset.Fields("Marks.Section") & " AND STA IN (SELECT STA FROM Participants WHERE INSTR('|" & cSplitCode & "|','|' & CLASS & '|')>0) ORDER BY Mark" & Format$(iTemp) & " DESC")
                                            ElseIf iPrintResult = 2 Then
                                                Set rstJ(iTemp) = mdbMain.OpenRecordset("SELECT Mark" & Format$(iTemp) & " FROM Marks Where Code='" & TestCode & "' AND Status=" & TestStatus & " AND Section=" & dtaAlready.Recordset.Fields("Marks.Section") & " AND STA IN (SELECT STA FROM Participants WHERE INSTR('" & cOldSplitCode & "','|' + CLASS + '|')=0) ORDER BY Mark" & Format$(iTemp) & " DESC")
                                            Else
                                                Set rstJ(iTemp) = mdbMain.OpenRecordset("SELECT Mark" & Format$(iTemp) & " FROM Marks Where Code='" & TestCode & "' AND Status=" & TestStatus & " AND Section=" & dtaAlready.Recordset.Fields("Marks.Section") & " ORDER BY Mark" & Format$(iTemp) & " DESC")
                                            End If
                                        Else
                                            Set rstJ(iTemp) = mdbMain.OpenRecordset("SELECT Mark" & Format$(iTemp) & " FROM Marks Where Code='" & TestCode & "' AND Status=" & TestStatus & " AND Section=" & dtaAlready.Recordset.Fields("Marks.Section") & " ORDER BY Mark" & Format$(iTemp) & " DESC")
                                        End If
                                         If rstJ(iTemp).RecordCount > 0 Then
                                             rstJ(iTemp).FindFirst "Mark" & Format$(iTemp) & "=" & Replace(dtaAlready.Recordset.Fields("Mark" & Format$(iTemp)), ",", ".")
                                             rtfResult.SelFontSize = rtfResult.SelFontSize - 2
                                             If rstJ(iTemp).NoMatch = False Then
                                                 rtfResult.SelText = Format$(rstJ(iTemp).AbsolutePosition + 1 + iHighestPosition, " (##)")
                                             Else
                                                 rtfResult.SelText = " (-)"
                                             End If
                                             rtfResult.SelFontSize = rtfResult.SelFontSize + 2
                                         End If
                                         rstJ(iTemp).Close
                                     End If
                                 Else
                                     rtfResult.SelText = vbTab
                                 End If
                             Else
                                 rtfResult.SelText = vbTab
                             End If
                         Next iTemp
                     End If
                     If fraMarks.Visible = True Then
                         If TestStatus <> 0 Then
                             rtfResult.SelText = vbTab & cScore & vbCrLf
                         Else
                             rtfResult.SelText = vbTab & " " & vbCrLf
                         End If
                     ElseIf fraTime.Visible = True And cScore <> mcNoPosition Then
                         rtfResult.SelText = vbTab & cScore & vbCrLf
                     End If
                     rtfResult.SelStrikeThru = False
                     curOldResult = dtaAlready.Recordset.Fields("Results.Score")
                     cOldAllTimes = dtaAlready.Recordset.Fields("AllTimes") & ""
                     If iPosition Mod 5 = 0 Then
                         iEmptyline = True
                     End If
                 End If
                 
                 'Send info to LogDB if desired
                 If miWriteLogDB Then
                     On Error Resume Next
                     If dtaAlready.Recordset.Fields("DISQ") < 0 Then
                         logDisq = Abs(dtaAlready.Recordset.Fields("DISQ"))
                         logPos = 0
                         logResult = 0
                     Else
                         logDisq = 0
                         If cPosition = "---" Then
                             logPos = 0
                         Else
                             logPos = CInt(cPosition)
                         End If
                         If cResult = "---" Then
                             logResult = 0
                         Else
                             logResult = CDbl(Replace(cResult, Chr$(34), ""))
                         End If
                     End If
                     
                      'Find judges names:
                      Dim cLogJudges As String
                      If GetJudgeId(TestCode, TestStatus, 1) <> "" Then
                          For iTemp = 1 To dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus))
                              If dtaTest.Recordset.Fields("Type_special") = 2 And iTemp = 5 Then
                                  cLogJudges = cLogJudges & "Time: "
                              Else
                                  cLogJudges = cLogJudges & Chr$(64 + iTemp) & ": "
                          End If
                          cLogJudges = cLogJudges & GetPersonsName(GetJudgeId(TestCode, TestStatus, iTemp))
                          cLogJudges = cLogJudges & "; "
                      Next iTemp
                      End If
                      
                 End If
            End If
            dtaAlready.Recordset.MoveNext
        Loop
        
        If miWriteLogDB Then
            writeLogDB = WriteLogDBConfMarks2(EventName, dtaTest.Recordset("code"), IIf(IsNull(dtaTest.Recordset("wrtest")), "", dtaTest.Recordset("wrtest")), TestStatus, TranslateBack(cVersion, mcLanguage))
        End If
        
        If miShowJudgesRanking <> 0 Then
            For iTemp = 1 To 5
                Set rstJ(iTemp) = Nothing
            Next iTemp
        End If
        
        If GetJudgeId(TestCode, TestStatus, 1) <> "" Then
            With rtfResult
                .SelTabCount = 1
                .SelTabs(0) = 16 * 567
                .SelUnderline = True
                .SelText = vbTab & vbCrLf
                .SelUnderline = False
                .SelBold = True
                .SelText = Translate("Judges", mcLanguage) & ":" & vbCrLf
                .SelBold = False
                .SelFontSize = 9
                For iTemp = 1 To dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus))
                    If dtaTest.Recordset.Fields("Type_special") = 2 And iTemp = 5 Then
                        .SelText = UCase$(Translate("Time", mcLanguage)) & ": "
                    Else
                        .SelText = Chr$(64 + iTemp) & ": "
                    End If
                    .SelText = GetPersonsName(GetJudgeId(TestCode, TestStatus, iTemp))
                    .SelText = "; "
                Next iTemp
                .SelText = vbCrLf
            End With
        End If
        
        If mnuFileFEIFWR.Enabled = True Then
            With rtfResult
                .SelText = vbCrLf
                .SelItalic = True
                .SelFontSize = 9
                .SelText = Translate("This is a FEIF WorldRanking Event", mcLanguage) & ". "
                curWRLimit = WRLimit(Me.TestCode)
                                
                If curWRLimit > 0 And TestStatus = 0 Then
                    .SelItalic = True
                    If fraTime.Visible = True Then
                        .SelText = Translate("Times of", mcLanguage) & " " & Format$(curWRLimit, TestTotalFormat) & Chr$(34) & " " & Translate("or faster", mcLanguage) & " " & Translate("will count for the FEIF WorldRanking", mcLanguage) & " (" & WrTest(Me.TestCode) & ")."
                    Else
                        .SelText = Translate("Marks of", mcLanguage) & " " & Format$(curWRLimit, TestTotalFormat) & " " & Translate("or higher", mcLanguage) & " " & Translate("will count for the FEIF WorldRanking", mcLanguage) & " (" & WrTest(Me.TestCode) & ")."
                    End If
                End If
                .SelText = vbCrLf
                .SelItalic = False
            End With
        End If
                
        cTemp = cVersion
        If cSplitCode <> "" Then
            cTemp = cTemp & "-" & cSplitCode
        End If
        
        PrintRtfFooter cTemp
    Else
        SetMouseNormal
        MsgBox Translate("No results to print yet!", mcLanguage), vbExclamation
    End If
    
Return
        
SetTabsRider:
    iPosition = iPosition + 1
    If dtaAlready.Recordset.Fields("Disq") < 0 Or cResult = Translate("Eliminated", mcLanguage) Or cResult = Translate("Withdrawn", mcLanguage) Or cResult = mcNoPosition Then
        If cPosition <> mcNoPosition Then
            rtfResult.SelText = vbCrLf
        End If
        cPosition = mcNoPosition
    ElseIf dtaAlready.Recordset.Fields("Results.Score") <> curOldResult Or dtaAlready.Recordset.Fields("AllTimes") & "" <> cOldAllTimes Then
        cPosition = Format$(iPosition, "00")
    End If
    
    cRider = Replace(dtaAlready.Recordset.Fields("Name_First") & " " & dtaAlready.Recordset.Fields("Name_Middle") & " " & dtaAlready.Recordset.Fields("Name_Last"), "  ", " ")
    If dtaAlready.Recordset.Fields("Class") & "" <> "" Then
        cRider = cRider & " [" & dtaAlready.Recordset.Fields("Class") & "]"
    End If
    
    cHorse = Replace(dtaAlready.Recordset.Fields("Name_Horse") & "", "  ", " ")
    If miShowRidersClub <> 0 Or miShowRidersTeam <> 0 Or miShowRidersLK = 1 Then
        If miShowRidersClub <> 0 Then
            cRider = Trim$(cRider & " / " & GetRidersClub(dtaAlready.Recordset))
        Else
            cClub = ""
        End If
        If miShowRidersTeam <> 0 Then
            cRider = Trim$(cRider & " / " & GetRidersTeam(dtaAlready.Recordset))
        Else
            cTeam = ""
        End If
        'IPZV LK output:
        If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
            cRider = Trim$(cRider & " / " & GetRidersLk(dtaAlready.Recordset))
        End If
        cParticipant = cRider & vbCrLf & vbTab & vbTab & cHorse
        If miShowHorseId <> 0 Then
            cParticipant = cParticipant & " [" & GetHorseId(dtaAlready.Recordset) & "]"
        End If
    Else
        If miShowHorseId <> 0 Then
            cParticipant = Left$(cRider & " / " & cHorse, 35) & " [" & GetHorseId(dtaAlready.Recordset) & "]"
        Else
            cParticipant = Left$(cRider & " / " & cHorse, 45)
        End If
    End If
    cParticipant = Replace(cParticipant, "  ", " ")
    If iPosition = Val(cPosition) And iEmptyline = True And TestStatus = 0 Then
        iEmptyline = False
        If iMarkFinals = True And miMarkFinalsInResultLists <> 0 Then
            If dtaTestInfo.Recordset.Fields("Handling") = 1 Then
                Do While cRidersDouble <> ""
                    Parse cTemp, cRidersDouble, "|"
                    If cTemp <> "" Then
                        rtfResult.SelText = cTemp & " " & Translate("has more than one horse in this final.", mcLanguage) & vbCrLf
                    End If
                Loop
                cRidersList = "|"
                If iA_Line = False Then
                    iA_Line = True
                    PrintRtfLine
                End If
            ElseIf dtaTestInfo.Recordset.Fields("Handling") = 2 Then
                Do While cRidersDouble <> ""
                    Parse cTemp, cRidersDouble, "|"
                    If cTemp <> "" Then
                        rtfResult.SelText = cTemp & " " & Translate("has more than one horse in this final.", mcLanguage) & vbCrLf
                    End If
                Loop
                cRidersList = "|"
                If iA_Line = False Then
                    iA_Line = True
                    PrintRtfLine
                    If tbsSelFin.Tabs.Count > 1 And dtaAlready.Recordset.RecordCount > miBFinalLevel Then
                        With rtfResult
                            .SelText = vbCrLf
                            .SelFontSize = 11
                            .SelBold = True
                            .SelText = ClipAmp(tbsSelFin.Tabs(2).Caption & vbCrLf)
                        End With
                    End If
                ElseIf iB_line = False Then
                    iB_line = True
                    If tbsSelFin.Tabs.Count > 1 And dtaAlready.Recordset.RecordCount > miBFinalLevel Then
                        PrintRtfLine
                    End If
                End If
            ElseIf dtaTestInfo.Recordset.Fields("Handling") = 5 Then
                Do While cRidersDouble <> ""
                    Parse cTemp, cRidersDouble, "|"
                    If cTemp <> "" Then
                        rtfResult.SelText = cTemp & " " & Translate("has more than one horse in this final.", mcLanguage) & vbCrLf
                    End If
                Loop
                cRidersList = "|"
                If iA_Line = False Then
                    iA_Line = True
                    PrintRtfLine
                    If tbsSelFin.Tabs.Count > 1 And dtaAlready.Recordset.RecordCount > miBFinalLevel Then
                        With rtfResult
                            .SelText = vbCrLf
                            .SelFontSize = 11
                            .SelBold = True
                            .SelText = ClipAmp(tbsSelFin.Tabs(3).Caption & vbCrLf)
                        End With
                    End If
                ElseIf iB_line = False Then
                    iB_line = True
                    PrintRtfLine
                    If tbsSelFin.Tabs.Count > 1 And dtaAlready.Recordset.RecordCount > miCFinalLevel Then
                        With rtfResult
                            .SelText = vbCrLf
                            .SelFontSize = 11
                            .SelBold = True
                            .SelText = ClipAmp(tbsSelFin.Tabs(2).Caption & vbCrLf)
                        End With
                    End If
                ElseIf iC_line = False Then
                    iC_line = True
                    If tbsSelFin.Tabs.Count > 1 And dtaAlready.Recordset.RecordCount > miCFinalLevel Then
                        PrintRtfLine
                    End If
                End If
            End If
        End If
        rtfResult.SelText = vbCrLf
    End If
    
    'check on double rider
    If cVersion = ClipAmp(mnuFilePrintResultFinal.Caption) Then
        If InStr(cRidersList, "|" & cRider & "|") > 0 Then
            cRidersDouble = cRidersDouble & cRider & "|"
        Else
            cRidersList = cRidersList & cRider & "|"
        End If
    End If
    
    With rtfResult
        If cPosition = mcNoPosition Then
            .SelFontSize = 10
            .SelBold = False
        Else
            .SelFontSize = 11
            .SelBold = True
        End If
        .SelItalic = False
        .SelTabCount = 5
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 2.5 * 567
        .SelTabs(2) = 13.5 * 567
        .SelTabs(3) = 15 * 567
        .SelTabs(4) = 16 * 567
    End With
Return

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        If fraMarks.Visible = True Then
            .SelTabCount = 9
            .SelTabs(0) = 3 * 567
            .SelTabs(1) = 5 * 567
            .SelTabs(2) = 6.5 * 567
            .SelTabs(3) = 8 * 567
            .SelTabs(4) = 9.5 * 567
            .SelTabs(5) = 11 * 567
            .SelTabs(6) = 13# * 567
            .SelTabs(7) = 15 * 567
            .SelTabs(8) = 16 * 567
        ElseIf fraTime.Visible = True Then
            .SelTabCount = 4
            .SelTabs(0) = 3 * 567
            .SelTabs(1) = 11.5 * 567
            .SelTabs(2) = 13 * 567
            .SelTabs(3) = 16 * 567
        End If
    End With
Return


End Sub
Sub PrintJudgesForm()
    Dim iKey As Integer
    Dim iColumn As Integer
    Dim iRow As Integer
    Dim iMaxSections As Integer
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iPagenum As Integer
    Dim iMaxPageNum As Integer
    Dim iMaxRows As Integer
    Dim iMaxUsedRows As Integer
    Dim iRowCounter As Integer
    Dim iRiderWidth As Integer
    
    Dim rstTest As DAO.Recordset
    Dim cNr As String
    Dim cGroup As String
    Dim cColor As String
    Dim cRider As String
    Dim cHorse As String
    Dim cClub As String
    Dim cClass As String
    Dim cTeam As String
    Dim cSql As String
    
    iMaxRows = 10
    iMaxUsedRows = 9
    
    ' use seperate forms for finals (gaedingakeppni)
    If TestStatus > 0 Then
        iTemp2 = 4
    Else
        iTemp2 = 3
    End If
    
    For iTemp = iTemp2 To 0 Step -1
        cSql = "SELECT * FROM TestSections WHERE Code LIKE '" & Me.TestCode & "' AND Status=" & iTemp & " ORDER BY Section"
        Set rstTest = mdbMain.OpenRecordset(cSql)
        If rstTest.RecordCount > 0 Then Exit For
    Next iTemp
    
    If rstTest.RecordCount = 0 Then
        MsgBox Translate("No form available.", mcLanguage), vbExclamation
    Else
        rstTest.MoveLast
        rstTest.MoveFirst
    End If
    iMaxSections = rstTest.RecordCount
    
    If dtaNotYet.Recordset.RecordCount > 0 Then
                
        With rtfResult
            .Text = ""
        End With
        
        GoSub PrintFormHeader
        
        dtaNotYet.Recordset.MoveLast
        iMaxPageNum = dtaNotYet.Recordset.RecordCount \ iMaxUsedRows
        If dtaNotYet.Recordset.RecordCount Mod iMaxUsedRows > 0 Then
            iMaxPageNum = iMaxPageNum + 1
        End If
        
        dtaNotYet.Recordset.MoveFirst
        Do While Not dtaNotYet.Recordset.EOF
            cNr = dtaNotYet.Recordset.Fields("Sta")
            cGroup = ""
            If miUseColors = 1 Then
                cColor = UCase$(Left$(dtaNotYet.Recordset.Fields("Color") & "  ", 2))
            End If
            cRider = dtaNotYet.Recordset.Fields("Name_First") & " " & dtaNotYet.Recordset.Fields("Name_Last")
            If dtaNotYet.Recordset.Fields("Class") & "" <> "" Then
                cRider = cRider & " [" & dtaNotYet.Recordset.Fields("Class") & "]"
            End If
            cHorse = dtaNotYet.Recordset.Fields("Name_Horse")
            If miShowHorseId <> 0 Then
                cHorse = cHorse & " [" & GetHorseId(dtaNotYet.Recordset) & "]"
            End If
            If miShowRidersClub <> 0 Then
                cClub = GetRidersClub(dtaNotYet.Recordset)
            Else
                cClub = ""
            End If
            If miShowRidersTeam <> 0 Then
                cTeam = GetRidersTeam(dtaNotYet.Recordset)
            Else
                cTeam = ""
            End If
            'IPZV LK output:
            If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                cTeam = Trim$(cTeam & " / " & GetRidersLk(dtaNotYet.Recordset))
            End If
            GoSub PrintTableRow
            
            If (dtaNotYet.Recordset.AbsolutePosition + 1) Mod iMaxUsedRows = 0 And dtaNotYet.Recordset.AbsolutePosition < dtaNotYet.Recordset.RecordCount - 1 Then
                If iMaxRows > iMaxUsedRows Then
                    For iTemp = iMaxUsedRows + 1 To iMaxRows
                        GoSub PrintEmptyRow
                    Next iTemp
                End If
                If dtaNotYet.Recordset.RecordCount > iMaxRows Then
                    GoSub PrintPageNumber
                    rtfResult.SelText = "$#@!"
                    GoSub PrintFormHeader
                End If
            End If
            iRowCounter = (dtaNotYet.Recordset.AbsolutePosition + 1) Mod iMaxUsedRows
            dtaNotYet.Recordset.MoveNext
        Loop
        
        Do While iRowCounter < iMaxRows
            GoSub PrintEmptyRow
        Loop
        GoSub PrintPageNumber
        
        PrintRtfFooter Translate("Judges' Form", mcLanguage)
    Else
        iKey = MsgBox(Translate("No list of relevant participants available. Print empty form?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton2)
        If iKey = vbYes Then
        
            With rtfResult
                .Text = ""
            End With
            
            GoSub PrintFormHeader
            
            For iRow = 1 To iMaxRows
                GoSub PrintTableRow
            Next iRow
            
            PrintRtfFooter Translate("Judges' Form", mcLanguage)
            
        End If
        
    End If
    rstTest.Close
    Set rstTest = Nothing
    
Exit Sub


PrintEmptyRow:
    iRowCounter = iRowCounter + 1
    cNr = ""
    cGroup = ""
    cColor = ""
    cRider = ""
    cHorse = ""
    cClub = ""
    cClass = ""
    cTeam = ""
    
    GoSub PrintTableRow
Return

PrintPageNumber:
    With rtfResult
        .SelFontSize = 10
        .SelItalic = True
        .SelText = vbCrLf
        .SelAlignment = rtfRight
        .SelText = Translate("Page", mcLanguage) & ": "
        .SelText = iPagenum & "/" & iMaxPageNum
        .SelText = vbCrLf
        .SelAlignment = rtfLeft
        .SelItalic = False
    End With
Return

PrintFormHeader:
    PrintRtfHeader Translate("Judges' Form", mcLanguage), False
    iPagenum = iPagenum + 1
    
    With rtfResult
        .SelFontSize = 12
        .SelBold = True
        .SelText = Translate("Judge", mcLanguage) & ": "
        .SelUnderline = True
        .SelText = Space$(50)
        .SelUnderline = False
        .SelFontSize = 10
        .SelText = vbCrLf
        .SelBold = False
    End With
    
    GoSub SetTabsMarks
    
    GoSub PrintTableHeader
Return

SetTabsMarks:
    With rtfResult
        .SelItalic = True
        .SelTabCount = 3 + iMaxSections + 1
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 2 * 567
        For iColumn = .SelTabCount - 1 To 2 Step -1
            .SelTabs(iColumn) = (16 - (.SelTabCount - 1 - iColumn) * 1.5) * 567
        Next iColumn
        iRiderWidth = .SelTabs(2) - .SelTabs(1) - 350
        If iRiderWidth < 200 Then
            iRiderWidth = 200
        End If
    End With
Return

PrintTableHeader:
    With rtfResult
        .SelFontSize = 9
        .SelBold = True
        .SelUnderline = True
        .SelItalic = True
        .SelText = vbTab & vbTab & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = vbTab
        Next iColumn
        .SelText = vbCrLf
        .SelUnderline = False
        .SelBold = True
        .SelText = "| " & UCase$(Translate("Nr", mcLanguage)) & vbTab & "| " & IIf(miUseColors = 1, Left(UCase$(Translate("Col", mcLanguage)), 3), "") & vbTab & "| " & Translate("participant", mcLanguage) & vbTab
        rstTest.MoveFirst
        For iColumn = 1 To iMaxSections
            .SelText = "|" & Left$(Translate(rstTest.Fields("Name"), mcLanguage), 6) & vbTab
            rstTest.MoveNext
        Next iColumn
        .SelText = "|" & UCase$(Left$(Translate("Total", mcLanguage), 3)) & vbTab
        .SelText = "|" & vbCrLf
        .SelBold = True
        .SelItalic = True
        .SelUnderline = True
        .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab
        rstTest.MoveFirst
        For iColumn = 1 To iMaxSections
            If rstTest.Fields("Factor") <> 1 Then
                .SelText = "| x" & Format$(rstTest.Fields("Factor"), "0") & vbTab
            Else
                .SelText = "| " & vbTab
            End If
            rstTest.MoveNext
        Next iColumn
        .SelText = "|" & vbTab
        .SelText = "|" & vbCrLf
        .SelBold = False
        .SelItalic = False
        .SelUnderline = False
        .SelFontSize = 10
    End With
Return

PrintTableRow:
    With rtfResult
        Me.FontSize = .SelFontSize
        .SelUnderline = False
        .SelText = "| " & cNr & vbTab
        .SelText = "| " & cColor & vbTab & "| " & FitString(Me, cRider, iRiderWidth, 2) & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & cGroup & vbTab & "| " & vbTab & "| " & FitString(Me, cHorse, iRiderWidth, 2) & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "|" & vbTab & "|" & vbTab & "|" & FitString(Me, cTeam, iRiderWidth, 2) & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = True
        .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
    End With
Return

End Sub
Sub PrintTimeForm()
    Dim iKey As Integer
    Dim iColumn As Integer
    Dim iRow As Integer
    Dim iMaxSections As Integer
    Dim iTemp As Integer
    Dim iPagenum As Integer
    Dim iMaxRows As Integer
    Dim iMaxPageNum As Integer
    Dim iRiderWidth As Integer
    
    Dim rstTest As DAO.Recordset
    Dim cNr As String
    Dim cGroup As String
    Dim cColor As String
    Dim cRider As String
    Dim cHorse As String
    Dim cClub As String
    Dim cClass As String
    Dim cTeam As String
    
    iMaxRows = 8
    
    For iTemp = 3 To 0 Step -1
        Set rstTest = mdbMain.OpenRecordset("SELECT * FROM TestSections WHERE Code LIKE '" & Me.TestCode & "' AND Status=" & iTemp & " ORDER BY Section")
        If rstTest.RecordCount > 0 Then Exit For
    Next iTemp
    If rstTest.RecordCount = 0 Then
        MsgBox Translate("No form available.", mcLanguage), vbExclamation
    Else
        rstTest.MoveLast
        rstTest.MoveFirst
    End If
    iMaxSections = 3
    
    If dtaNotYet.Recordset.RecordCount > 0 Then
                
        With rtfResult
            .Text = ""
        End With
        
        GoSub PrintFormHeader
        
        dtaNotYet.Recordset.MoveLast
        iMaxPageNum = dtaNotYet.Recordset.RecordCount \ iMaxRows
        If dtaNotYet.Recordset.RecordCount Mod iMaxRows > 0 Then
            iMaxPageNum = iMaxPageNum + 1
        End If
        
        dtaNotYet.Recordset.MoveFirst
        Do While Not dtaNotYet.Recordset.EOF
            cNr = dtaNotYet.Recordset.Fields("Sta")
            cGroup = dtaNotYet.Recordset.Fields("Group")
            If miUseColors = 1 Then
                cColor = UCase$(Left$(dtaNotYet.Recordset.Fields("Color") & "  ", 2))
            End If
            cRider = dtaNotYet.Recordset.Fields("Name_First") & " " & dtaNotYet.Recordset.Fields("Name_Last")
            If dtaNotYet.Recordset.Fields("Class") & "" <> "" Then
                cRider = cRider & " [" & dtaNotYet.Recordset.Fields("Class") & "]"
            End If
            cHorse = dtaNotYet.Recordset.Fields("Name_Horse")
            If miShowHorseId <> 0 Then
                cHorse = cHorse & " [" & GetHorseId(dtaNotYet.Recordset) & "]"
            End If
            If miShowRidersClub <> 0 Then
                cClub = GetRidersClub(dtaNotYet.Recordset)
            Else
                cClub = ""
            End If
            If miShowRidersTeam <> 0 Then
                cTeam = GetRidersTeam(dtaNotYet.Recordset)
            Else
                cTeam = ""
            End If
            'IPZV LK output:
            If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                cTeam = Trim$(cTeam & " / " & GetRidersLk(dtaNotYet.Recordset))
            End If
            GoSub PrintTableRow
            If (dtaNotYet.Recordset.AbsolutePosition + 1) Mod iMaxRows = 0 And dtaNotYet.Recordset.AbsolutePosition < dtaNotYet.Recordset.RecordCount - 1 Then
                If dtaNotYet.Recordset.RecordCount > iMaxRows Then
                    GoSub PrintPageNumber
                    rtfResult.SelText = "$#@!"
                    GoSub PrintFormHeader
                End If
            End If
            iTemp = (dtaNotYet.Recordset.AbsolutePosition + 1) Mod iMaxRows
            dtaNotYet.Recordset.MoveNext
        Loop
        
        Do While iTemp < iMaxRows
            iTemp = iTemp + 1
            cNr = ""
            cGroup = ""
            cColor = ""
            cRider = ""
            cHorse = ""
            cClub = ""
            cTeam = ""
            GoSub PrintTableRow
        Loop
        GoSub PrintPageNumber
        
        PrintRtfFooter Translate("Judges' Form", mcLanguage)
    Else
        iKey = MsgBox(Translate("No list of relevant participants available. Print empty form?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton2)
        If iKey = vbYes Then
        
            With rtfResult
                .Text = ""
            End With
            
            GoSub PrintFormHeader
            
            For iRow = 1 To iMaxRows
                GoSub PrintTableRow
            Next iRow
            
            PrintRtfFooter Translate("Judges' Form", mcLanguage)
            
        End If
        
    End If
    rstTest.Close
    Set rstTest = Nothing
    
Exit Sub

PrintPageNumber:
    With rtfResult
        .SelFontSize = 10
        .SelItalic = True
        .SelText = vbCrLf
        .SelAlignment = rtfRight
        .SelText = Translate("Page", mcLanguage) & ": "
        .SelText = iPagenum & "/" & iMaxPageNum
        .SelText = vbCrLf
        .SelAlignment = rtfLeft
        .SelItalic = False
    End With
Return

PrintFormHeader:
    PrintRtfHeader Translate("Time Keepers Form", mcLanguage), False
    iPagenum = iPagenum + 1
    
    With rtfResult
        .SelFontSize = 12
        .SelBold = True
        .SelText = Translate("Judge", mcLanguage) & ": "
        .SelUnderline = True
        .SelText = Space$(50)
        .SelUnderline = False
        .SelText = "  "
        .SelText = Translate("Heat", mcLanguage) & ": "
        .SelUnderline = True
        .SelText = Space$(10)
        .SelUnderline = False
        .SelText = vbCrLf
        .SelBold = False
    End With
    
    GoSub SetTabsMarks
    
    GoSub PrintTableHeader
Return

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        .SelTabCount = 3 + iMaxSections + 1
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 2 * 567
        For iColumn = .SelTabCount - 1 To 2 Step -1
            .SelTabs(iColumn) = (16 - (.SelTabCount - 1 - iColumn) * 1.5) * 567
        Next iColumn
        iRiderWidth = .SelTabs(2) - .SelTabs(1) - 350
        If iRiderWidth < 200 Then
            iRiderWidth = 200
        End If
    End With
Return

PrintTableHeader:
    With rtfResult
        .SelFontSize = 9
        .SelBold = True
        .SelUnderline = True
        .SelItalic = True
        .SelText = vbTab & vbTab & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = vbTab
        Next iColumn
        .SelText = vbCrLf
        .SelUnderline = False
        .SelBold = True
        .SelText = "| " & UCase$(Translate("Nr", mcLanguage)) & vbTab & "| " & IIf(miUseColors = 1, UCase$(Translate("Col", mcLanguage)), "") & vbTab & "| " & Translate("participant", mcLanguage) & vbTab
        For iColumn = 1 To iMaxSections
            .SelText = "|" & Left$(Translate("Time", mcLanguage), 4) & " " & Format$(iColumn) & vbTab
        Next iColumn
        .SelText = "|" & UCase$(Left$(Translate("Total", mcLanguage), 3)) & vbTab
        .SelText = "|" & vbCrLf
        .SelBold = True
        .SelItalic = True
        .SelUnderline = True
        .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab
        For iColumn = 1 To iMaxSections
            .SelText = "| " & vbTab
        Next iColumn
        .SelText = "|" & vbTab
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelBold = False
        .SelItalic = False
    End With
Return

PrintTableRow:
    With rtfResult
        .SelFontSize = 10
        Me.FontSize = .SelFontSize
        .SelUnderline = False
        .SelFontSize = 12
        .SelText = "| " & cNr & vbTab
        .SelFontSize = 10
        .SelText = "| " & cColor & vbTab & "| " & FitString(Me, cRider, iRiderWidth, 2) & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & cGroup & vbTab & "| " & vbTab & "| " & FitString(Me, cHorse, iRiderWidth, 2) & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & vbTab & "| " & vbTab & "| " & FitString(Me, cTeam, iRiderWidth, 2) & vbTab
        .SelUnderline = True
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = True
        .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab
        For iColumn = 1 To iMaxSections + 1
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
    End With
Return


End Sub
Sub PrintVetForm(Optional cSta As String)
    Dim iKey As Integer
    Dim iColumn As Integer
    Dim iRow As Integer
    Dim iMaxSections As Integer
    Dim iTemp As Integer
    Dim iPagenum As Integer
    Dim iMaxRows As Integer
    Dim iMaxPageNum As Integer
    Dim iRiderWidth As Integer
    Dim iLine As Integer
    
    Dim rstAll As DAO.Recordset
    Dim cNr As String
    Dim cGroup As String
    Dim cColor As String
    Dim cRider As String
    Dim cHorse As String
    Dim cHorseId As String
    Dim cClub As String
    Dim cClass As String
    Dim cTeam As String
    Dim cQry As String
    
    
    iMaxRows = 5
    
    cQry = "SELECT Participants.*, Persons.*, Horses.*"
    cQry = cQry & " FROM (Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
    If cSta <> "" Then
        cQry = cQry & " WHERE STA IN (" & cSta & " )"
    End If
    cQry = cQry & " ORDER BY Persons.Name_First & ' ' & Persons.Name_Last,Horses.Name_Horse;"

    Set rstAll = mdbMain.OpenRecordset(cQry)
    If rstAll.RecordCount > 0 Then
        
        SetMouseHourGlass
        
        rstAll.MoveLast
        rstAll.MoveFirst
                
        With rtfResult
            .Text = ""
        End With
        
        GoSub PrintFormHeader
        
        iMaxPageNum = rstAll.RecordCount \ iMaxRows
        If rstAll.RecordCount Mod iMaxRows > 0 Then
            iMaxPageNum = iMaxPageNum + 1
        End If
        
        Do While Not rstAll.EOF
            cNr = rstAll.Fields("Sta")
            cRider = rstAll.Fields("Name_First") & " " & rstAll.Fields("Name_Last")
            If rstAll.Fields("Class") & "" <> "" Then
                cRider = cRider & " [" & rstAll.Fields("Class") & "]"
            End If
            cHorse = rstAll.Fields("Name_Horse")
            cHorseId = GetHorseId(rstAll)
            If miShowRidersClub <> 0 Then
                cClub = GetRidersClub(rstAll)
            Else
                cClub = ""
            End If
            If miShowRidersTeam <> 0 Then
                cTeam = GetRidersTeam(rstAll)
            Else
                cTeam = ""
            End If
             'IPZV LK output:
            If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                cTeam = Trim$(cTeam & " / " & GetRidersLk(rstAll))
            End If
            
            GoSub PrintTableRow
            If (rstAll.AbsolutePosition + 1) Mod iMaxRows = 0 And rstAll.AbsolutePosition < rstAll.RecordCount - 1 Then
                If rstAll.RecordCount > iMaxRows Then
                    GoSub PrintPageNumber
                    rtfResult.SelText = "$#@!"
                    GoSub PrintFormHeader
                End If
            End If
            iTemp = (rstAll.AbsolutePosition + 1) Mod iMaxRows
            rstAll.MoveNext
        Loop
        
        Do While iTemp < iMaxRows And iTemp <> 0
            iTemp = iTemp + 1
            cNr = ""
            cGroup = ""
            cColor = ""
            cRider = ""
            cHorse = ""
            cHorseId = ""
            cClub = ""
            cTeam = ""
            GoSub PrintTableRow
        Loop
        
        GoSub PrintPageNumber
        
        PrintRtfFooter Translate("Judges' Form", mcLanguage)
        
        SetMouseNormal
    Else
        iKey = MsgBox(Translate("No list of relevant participants available. Print empty form?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton2)
        If iKey = vbYes Then
        
            With rtfResult
                .Text = ""
            End With
            
            GoSub PrintFormHeader
            
            For iRow = 1 To iMaxRows
                GoSub PrintTableRow
            Next iRow
            
            PrintRtfFooter Translate("Veterinary Check Form", mcLanguage)
            
        End If
        
    End If
    rstAll.Close
    Set rstAll = Nothing
    
Exit Sub

PrintPageNumber:
    With rtfResult
        .SelFontSize = 10
        .SelItalic = True
        .SelText = vbCrLf
        .SelAlignment = rtfRight
        .SelText = Translate("Page", mcLanguage) & ": "
        .SelText = iPagenum & "/" & iMaxPageNum
        .SelText = vbCrLf
        .SelAlignment = rtfLeft
        .SelItalic = False
    End With
Return

PrintFormHeader:
    PrintRtfHeader Translate("Veterinary Checks", mcLanguage), False, True
    
    iPagenum = iPagenum + 1
    
    With rtfResult
        .SelFontSize = 12
        .SelBold = True
        .SelText = Translate("Veterinarian", mcLanguage) & ": "
        .SelBold = False
        .SelText = String$(40, "_")
        .SelText = vbCrLf
    End With
    
    GoSub SetTabsMarks
    
    GoSub PrintTableHeader
Return

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        .SelTabCount = 6
        .SelTabs(0) = 1 * 567
        For iColumn = .SelTabCount - 1 To 2 Step -1
            .SelTabs(iColumn) = (16 - (.SelTabCount - 1 - iColumn) * 1.5) * 567
        Next iColumn
        iRiderWidth = .SelTabs(2) - .SelTabs(1) - 400
        If iRiderWidth < 200 Then
            iRiderWidth = 200
        End If
    End With
Return

PrintTableHeader:
    With rtfResult
        .SelFontSize = 9
        .SelBold = True
        .SelUnderline = True
        .SelItalic = True
        .SelText = vbTab & vbTab
        For iColumn = 1 To 3
            .SelText = vbTab
        Next iColumn
        .SelText = vbCrLf
        .SelUnderline = False
        .SelBold = True
        .SelText = "| " & UCase$(Translate("Nr", mcLanguage)) & vbTab & "| " & Translate("Participant", mcLanguage) & vbTab
        .SelText = "|" & Left$(Translate("Passport", mcLanguage), 6) & vbTab
        .SelText = "|" & Left$(Translate("Vaccinations", mcLanguage), 6) & vbTab
        .SelText = "|" & Left$(Translate("Health", mcLanguage), 6) & vbTab
        .SelText = "|" & vbCrLf
        .SelBold = True
        .SelItalic = True
        .SelUnderline = True
        .SelText = "|" & vbTab & "|" & vbTab
        For iColumn = 1 To 3
            .SelText = "| " & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelBold = False
        .SelItalic = False
    End With
Return

PrintTableRow:
    With rtfResult
        .SelFontSize = 10
        Me.FontSize = .SelFontSize
        .SelUnderline = False
        .SelFontSize = 12
        .SelText = "| " & cNr & vbTab
        .SelFontSize = 10
        .SelText = "| " & FitString(Me, cRider, iRiderWidth, 2) & vbTab
        For iColumn = 1 To 3
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelFontSize = 10
        .SelText = "| " & vbTab & "| " & FitString(Me, cHorse, iRiderWidth, 2) & vbTab
        For iColumn = 1 To 3
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelFontSize = 10
        .SelText = "| " & vbTab & "| " & FitString(Me, cHorseId, iRiderWidth, 2) & vbTab
        For iColumn = 1 To 3
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & vbTab & "| " & FitString(Me, cTeam, iRiderWidth, 2) & vbTab
        .SelUnderline = True
        For iColumn = 1 To 3
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "|" & vbTab & "| " & Translate("Comment", mcLanguage) & ":" & vbTab
        For iColumn = 1 To 3
            .SelText = " " & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        
        For iLine = 1 To 2
            .SelText = "|" & vbTab & "|" & vbTab
            For iColumn = 1 To 3
                .SelText = " " & vbTab
            Next iColumn
            .SelText = "|" & vbCrLf
        Next iLine
        
        .SelUnderline = True
        .SelText = "|" & vbTab & "| " & Translate("Date/time", mcLanguage) & ":" & vbTab
        For iColumn = 1 To 3
            .SelText = " " & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
    End With
Return


End Sub
Sub PrintEntranceForm(iOrder As Integer)
    Dim iKey As Integer
    Dim iColumn As Integer
    Dim iRow As Integer
    Dim iMaxSections As Integer
    Dim iTemp As Integer
    Dim iPagenum As Integer
    Dim iMaxRows As Integer
    Dim iMaxPageNum As Integer
    Dim iRiderWidth As Integer
    Dim iLine As Integer
    
    Dim rstAll As DAO.Recordset
    Dim cNr As String
    Dim cGroup As String
    Dim cColor As String
    Dim cRider As String
    Dim cHorse As String
    Dim cHorseId As String
    Dim cClub As String
    Dim cClass As String
    Dim cTeam As String
    Dim cQry As String
    
    
    iMaxRows = 20
    
    cQry = "SELECT Participants.*, Persons.*, Horses.*"
    cQry = cQry & " FROM (Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
    If iOrder = 0 Then
        cQry = cQry & " ORDER BY Persons.Name_First & ' ' & Persons.Name_Last,Horses.Name_Horse;"
    Else
        cQry = cQry & " ORDER BY Participants.Sta, Persons.Name_First & ' ' & Persons.Name_Last,Horses.Name_Horse;"
    End If

    Set rstAll = mdbMain.OpenRecordset(cQry)
    If rstAll.RecordCount > 0 Then
        
        SetMouseHourGlass
        
        rstAll.MoveLast
        rstAll.MoveFirst
                
        With rtfResult
            .Text = ""
        End With
        
        GoSub PrintFormHeader
        
        iMaxPageNum = rstAll.RecordCount \ iMaxRows
        If rstAll.RecordCount Mod iMaxRows > 0 Then
            iMaxPageNum = iMaxPageNum + 1
        End If
        
        Do While Not rstAll.EOF
        
            cNr = rstAll.Fields("Sta")
            cRider = rstAll.Fields("Name_First") & " " & rstAll.Fields("Name_Last")
            cHorse = rstAll.Fields("Name_Horse")
            cHorseId = GetHorseId(rstAll)
            
            
            GoSub PrintTableRow
            If (rstAll.AbsolutePosition + 1) Mod iMaxRows = 0 And rstAll.AbsolutePosition < rstAll.RecordCount - 1 Then
                If rstAll.RecordCount > iMaxRows Then
                    GoSub PrintPageNumber
                    rtfResult.SelText = "$#@!"
                    GoSub PrintFormHeader
                End If
            End If
            iTemp = (rstAll.AbsolutePosition + 1) Mod iMaxRows
            rstAll.MoveNext
        Loop
        
        Do While iTemp < iMaxRows And iTemp <> 0
            iTemp = iTemp + 1
            cNr = ""
            cGroup = ""
            cColor = ""
            cRider = ""
            cHorse = ""
            cHorseId = ""
            cClub = ""
            cTeam = ""
            GoSub PrintTableRow
        Loop
        
        GoSub PrintPageNumber
        
        PrintRtfFooter Translate("Entrance Check Form", mcLanguage)
        
        SetMouseNormal
    Else
        iKey = MsgBox(Translate("No list of relevant participants available. Print empty form?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton2)
        If iKey = vbYes Then
        
            With rtfResult
                .Text = ""
            End With
            
            GoSub PrintFormHeader
            
            For iRow = 1 To iMaxRows
                GoSub PrintTableRow
            Next iRow
            
            PrintRtfFooter Translate("Judges' Form", mcLanguage)
            
        End If
        
    End If
    rstAll.Close
    Set rstAll = Nothing
    
Exit Sub

PrintPageNumber:
    With rtfResult
        .SelFontSize = 10
        .SelItalic = True
        .SelText = vbCrLf
        .SelAlignment = rtfRight
        .SelText = Translate("Page", mcLanguage) & ": "
        .SelText = iPagenum & "/" & iMaxPageNum
        .SelText = vbCrLf
        .SelAlignment = rtfLeft
        .SelItalic = False
    End With
Return

PrintFormHeader:
    PrintRtfHeader Translate("Entrance Checks", mcLanguage), False, True
    
    iPagenum = iPagenum + 1
    
    With rtfResult
        .SelFontSize = 12
        .SelBold = True
        .SelText = Translate("Responsible person", mcLanguage) & ": "
        .SelBold = False
        .SelText = String$(40, "_")
        .SelText = vbCrLf
    End With
    
    GoSub SetTabsMarks
    
    GoSub PrintTableHeader
Return

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        .SelTabCount = 5
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 6 * 567
        .SelTabs(2) = 11 * 567
        .SelTabs(3) = 12 * 567
        .SelTabs(4) = 15 * 567
        iRiderWidth = .SelTabs(1) - .SelTabs(0) - 400
        If iRiderWidth < 200 Then
            iRiderWidth = 200
        End If
    End With
Return

PrintTableHeader:
    With rtfResult
        .SelFontSize = 9
        .SelBold = True
        .SelUnderline = True
        .SelItalic = True
        .SelText = vbTab & vbTab
        For iColumn = 1 To 3
            .SelText = vbTab
        Next iColumn
        .SelText = vbCrLf
        .SelUnderline = False
        .SelBold = True
        .SelText = "| " & UCase$(Translate("Nr", mcLanguage)) & vbTab
        .SelText = "| " & Translate("Rider", mcLanguage) & vbTab
        .SelText = "| " & Translate("Horse", mcLanguage) & vbTab
        .SelText = "| " & Translate("OK", mcLanguage) & vbTab
        .SelText = "| " & Translate("Comment", mcLanguage) & vbTab
        .SelText = "| " & vbCrLf
        .SelBold = True
        .SelItalic = True
        .SelUnderline = True
        For iColumn = 1 To 5
            .SelText = "| " & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelBold = False
        .SelItalic = False
    End With
Return

PrintTableRow:
    With rtfResult
        .SelFontSize = 10
        Me.FontSize = .SelFontSize
        .SelUnderline = False
        .SelFontSize = 12
        .SelText = "| " & cNr & vbTab
        .SelFontSize = 10
        .SelText = "| " & FitString(Me, cRider, iRiderWidth, 2) & vbTab
        .SelText = "| " & FitString(Me, cHorse, iRiderWidth, 2) & vbTab
        For iColumn = 1 To 2
            .SelText = "|" & vbTab
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelItalic = True
        .SelUnderline = True
        For iColumn = 1 To 5
            If iColumn = 3 Then
                .SelText = "| " & FitString(Me, cHorseId, iRiderWidth, 2) & vbTab
            Else
                .SelText = "| " & vbTab
            End If
        Next iColumn
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelItalic = False
    End With
Return


End Sub

Sub PrintWarnings()
    Dim iKey As Integer
    Dim iColumn As Integer
    Dim iRow As Integer
    Dim iMaxSections As Integer
    Dim iTemp As Integer
    Dim iPagenum As Integer
    Dim iMaxRows As Integer
    Dim iMaxPageNum As Integer
    Dim iRiderWidth As Integer
    Dim iLine As Integer
    
    Dim rstAll As DAO.Recordset
    Dim cNr As String
    Dim cGroup As String
    Dim cColor As String
    Dim cRider As String
    Dim cHorse As String
    Dim cHorseId As String
    Dim cClub As String
    Dim cClass As String
    Dim cTeam As String
    Dim cMeasure As String
    Dim cCause As String
    Dim cComments As String
    Dim cJudge As String
    Dim cTest As String
    Dim cQry As String
    
    
    iMaxRows = 5
    
    cQry = "SELECT Penalties.Measure, Penalties.Cause, Penalties.Comments, Penalties.STA, Penalties.Test, Penalties.Status, [Persons].[Name_Last] & ', ' & [Persons].[Name_First] AS Culprit, [Persons_1].[Name_Last] & ', ' & [Persons_1].[Name_First] AS Judge, Penalties.PersonID, Penalties.Responsible_ID "
    cQry = cQry & "FROM (Persons INNER JOIN Penalties ON Persons.PersonID = Penalties.PersonID) INNER JOIN Persons AS Persons_1 ON Penalties.Responsible_ID = Persons_1.PersonID "
    cQry = cQry & "ORDER BY [Persons].[Name_Last] & ', ' & [Persons].[Name_First];"

    Set rstAll = mdbMain.OpenRecordset(cQry)
    If rstAll.RecordCount > 0 Then
        
        SetMouseHourGlass
        
        rstAll.MoveLast
        rstAll.MoveFirst
                
        With rtfResult
            .Text = ""
        End With
        
        GoSub PrintFormHeader
        
        iMaxPageNum = rstAll.RecordCount \ iMaxRows
        If rstAll.RecordCount Mod iMaxRows > 0 Then
            iMaxPageNum = iMaxPageNum + 1
        End If
        
        Do While Not rstAll.EOF
            cNr = rstAll.Fields("Sta")
            cRider = rstAll.Fields("culprit")
            cMeasure = rstAll.Fields("Measure") & ""
            cCause = Translate("Cause:", mcLanguage) & " " & rstAll.Fields("Cause")
            cComments = rstAll.Fields("Comments") & ""
            cJudge = Translate("Responsible:", mcLanguage) & " " & rstAll.Fields("Judge")
            cTest = Translate("Test:", mcLanguage) & " " & rstAll.Fields("Test")
            
            GoSub PrintTableRow
            If (rstAll.AbsolutePosition + 1) Mod iMaxRows = 0 And rstAll.AbsolutePosition < rstAll.RecordCount - 1 Then
                If rstAll.RecordCount > iMaxRows Then
                    GoSub PrintPageNumber
                    rtfResult.SelText = "$#@!"
                    GoSub PrintFormHeader
                End If
            End If
            iTemp = (rstAll.AbsolutePosition + 1) Mod iMaxRows
            rstAll.MoveNext
        Loop
        
        Do While iTemp < iMaxRows And iTemp <> 0
            iTemp = iTemp + 1
            cNr = ""
            cRider = ""
            cMeasure = ""
            cCause = ""
            cComments = ""
            cJudge = ""
            cTest = ""
            GoSub PrintTableRow
        Loop
        
        GoSub PrintPageNumber
        
        PrintRtfFooter Translate("Disciplinary Report", mcLanguage)
        
        SetMouseNormal
    Else
        iKey = MsgBox(Translate("No list of relevant participants available. Print empty form?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton2)
        If iKey = vbYes Then
        
            With rtfResult
                .Text = ""
            End With
            
            GoSub PrintFormHeader
            
            For iRow = 1 To iMaxRows
                GoSub PrintTableRow
            Next iRow
            
            PrintRtfFooter Translate("Disciplinary Report", mcLanguage)
            
        End If
        
    End If
    rstAll.Close
    Set rstAll = Nothing
    
Exit Sub

PrintPageNumber:
    With rtfResult
        .SelFontSize = 10
        .SelItalic = True
        .SelText = vbCrLf
        .SelAlignment = rtfRight
        .SelText = Translate("Page", mcLanguage) & ": "
        .SelText = iPagenum & "/" & iMaxPageNum
        .SelText = vbCrLf
        .SelAlignment = rtfLeft
        .SelItalic = False
    End With
Return

PrintFormHeader:
    PrintRtfHeader Translate("Disciplinary Measures", mcLanguage), False, True
    
    iPagenum = iPagenum + 1
    
    With rtfResult
        .SelFontSize = 12
        .SelBold = True
        .SelText = Translate("Head judge", mcLanguage) & ": "
        .SelUnderline = True
        .SelText = Space$(50)
        .SelUnderline = False
        .SelText = vbCrLf
        .SelBold = False
    End With
    
    GoSub SetTabsMarks
    
    GoSub PrintTableHeader
Return

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        .SelTabCount = 6
        .SelTabs(0) = 1 * 567
        .SelTabs(2) = 9072
        iRiderWidth = .SelTabs(2) - .SelTabs(1) - 400
        If iRiderWidth < 200 Then
            iRiderWidth = 200
        End If
    End With
Return

PrintTableHeader:
    With rtfResult
        .SelFontSize = 9
        .SelBold = True
        .SelUnderline = True
        .SelItalic = True
        .SelText = vbTab & vbTab
        .SelText = vbCrLf
        .SelUnderline = False
        .SelBold = True
        .SelText = "| " & UCase$(Translate("Nr", mcLanguage)) & vbTab & "| " & Translate("Participant", mcLanguage) & vbTab
        .SelText = "|" & vbCrLf
        .SelBold = True
        .SelItalic = True
        .SelUnderline = True
        .SelText = "|" & vbTab & "|" & vbTab
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelBold = False
        .SelItalic = False
    End With
Return


PrintTableRow:
    With rtfResult
        .SelFontSize = 10
        Me.FontSize = .SelFontSize
        .SelUnderline = False
        .SelFontSize = 12
        .SelText = "| " & cNr & vbTab
        .SelFontSize = 10
        
        .SelText = "| " & FitString(Me, cRider, iRiderWidth, 2) & vbTab
        
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & vbTab & "| " & FitString(Me, cMeasure, iRiderWidth, 2) & vbTab
        
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & vbTab & "| " & FitString(Me, cCause, iRiderWidth, 2) & vbTab
        
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & vbTab & "| " & FitString(Me, cJudge, iRiderWidth, 2) & vbTab
        
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & vbTab & "| " & FitString(Me, cTest, iRiderWidth, 2) & vbTab
        .SelUnderline = True
        
        
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "|" & vbTab & "| " & Translate("Comment", mcLanguage) & ":" & vbTab
        '.SelText = "|" & vbCrLf
        
        .SelText = "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & vbTab & "| " & FitString(Me, cComments, iRiderWidth, 2) & vbTab
        
        .SelText = "|" & vbCrLf
        For iLine = 1 To 2
            .SelText = "|" & vbTab & "|" & vbTab
            .SelText = "|" & vbCrLf
        Next iLine
        
        .SelUnderline = True
        .SelText = "|" & vbTab & "| " & Translate("Date/time", mcLanguage) & ":" & vbTab
        .SelText = "|" & vbCrLf
        .SelUnderline = False
    End With
Return



End Sub
Sub PrintJudgesForm3()
    Dim iKey As Integer
    Dim iColumn As Integer
    Dim iRow As Integer
    Dim iMaxSections As Integer
    Dim iMaxRows As Integer
    Dim iTemp As Integer
    Dim iPagenum As Integer
    Dim iMaxPageNum As Integer
    Dim iRiderWidth As Integer
    Dim iMaxfactor As Integer
    
    Dim rstTestSections As DAO.Recordset
    Dim cNr As String
    Dim cGroup As String
    Dim cColor As String
    Dim cRider As String
    Dim cHorse As String
    Dim cClub As String
    Dim cClass As String
    Dim cTeam As String
    Dim cRein As String
    
    SetMouseHourGlass
       
    For iTemp = 3 To 0 Step -1
        Set rstTestSections = mdbMain.OpenRecordset("SELECT * FROM TestSections WHERE Code LIKE '" & Me.TestCode & "' AND Status=" & iTemp & " ORDER BY Section")
        If rstTestSections.RecordCount > 0 Then Exit For
    Next iTemp
    If rstTestSections.RecordCount = 0 Then
        MsgBox Translate("No form available.", mcLanguage), vbExclamation
    Else
        rstTestSections.MoveLast
        rstTestSections.MoveFirst
    End If
    
    iMaxSections = rstTestSections.RecordCount
    
    'check number of test sections to determine number of participants per sheet:
    Select Case iMaxSections
        Case Is < 4
            iMaxRows = 3
        Case 4 To 9
            iMaxRows = 2
        Case Is > 9
            iMaxRows = 1
    End Select
    
    If dtaNotYet.Recordset.RecordCount > 0 Then
                
        With rtfResult
            .Text = ""
            .Font.Size = 10
        End With
        
        dtaNotYet.Recordset.MoveLast
        iMaxPageNum = dtaNotYet.Recordset.RecordCount \ iMaxRows
        If dtaNotYet.Recordset.RecordCount Mod iMaxRows > 0 Then
            iMaxPageNum = iMaxPageNum + 1
        End If
        
        dtaNotYet.Recordset.MoveFirst
        Do While Not dtaNotYet.Recordset.EOF
            cNr = dtaNotYet.Recordset.Fields("Sta")
            cGroup = ""
            If miUseColors = 1 Then
                cColor = UCase$(Left$(dtaNotYet.Recordset.Fields("Color") & "  ", 2))
            End If
            cRider = dtaNotYet.Recordset.Fields("Name_First") & " " & dtaNotYet.Recordset.Fields("Name_Last")
            If dtaNotYet.Recordset.Fields("Class") & "" <> "" Then
                cRider = cRider & " [" & dtaNotYet.Recordset.Fields("Class") & "]"
            End If
            cHorse = dtaNotYet.Recordset.Fields("Name_Horse")
            If miShowHorseId <> 0 Then
                cHorse = cHorse & " [" & GetHorseId(dtaNotYet.Recordset) & "]"
            End If
            If miShowRidersClub <> 0 Then
                cClub = GetRidersClub(dtaNotYet.Recordset)
            Else
                cClub = ""
            End If
            If miShowRidersTeam <> 0 Then
                cTeam = GetRidersTeam(dtaNotYet.Recordset)
            Else
                cTeam = ""
            End If
            'IPZV LK output:
            If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                cTeam = Trim$(cTeam & " " & GetRidersLk(dtaNotYet.Recordset))
            End If
            
            cRein = IIf(dtaNotYet.Recordset.Fields("RR") = True, "R", "L")
            GoSub PrintTable
            If (dtaNotYet.Recordset.AbsolutePosition + 1) Mod iMaxRows = 0 And dtaNotYet.Recordset.AbsolutePosition < dtaNotYet.Recordset.RecordCount - 1 Then
                If dtaNotYet.Recordset.RecordCount > iMaxRows Then
                    rtfResult.SelText = "$#@!"
                End If
            End If
            iTemp = (dtaNotYet.Recordset.AbsolutePosition + 1) Mod iMaxRows
            dtaNotYet.Recordset.MoveNext
        Loop
        
        Do While iTemp < iMaxRows And iTemp <> 0
            iTemp = iTemp + 1
            cNr = ""
            cGroup = ""
            cColor = ""
            cRider = ""
            cHorse = ""
            cClub = ""
            cClass = ""
            cTeam = ""
            cRein = ""
            GoSub PrintTable
        Loop
        PrintRtfFooter Translate("Judges' Form", mcLanguage), "", 0, True
    Else
        iKey = MsgBox(Translate("No list of relevant participants available. Print empty form?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton2)
        If iKey = vbYes Then
        
            With rtfResult
                .Text = ""
            End With
            
            For iRow = 1 To iMaxRows
                GoSub PrintTable
            Next iRow
            
        End If
            
        PrintRtfFooter Translate("Judges' Form", mcLanguage), "", 0, True
        
    End If
    rstTestSections.Close
    Set rstTestSections = Nothing
    SetMouseNormal
Exit Sub

SetTabsMarks1:
    With rtfResult
        .SelItalic = True
        .SelTabCount = 4
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 8 * 567
        .SelTabs(2) = 10 * 567
        .SelTabs(3) = 16 * 567
        iRiderWidth = .SelTabs(2) - .SelTabs(0) - 350
        If iRiderWidth < 200 Then
            iRiderWidth = 200
        End If
    End With
Return

SetTabsMarks2:
    With rtfResult
        .SelItalic = True
        .SelTabCount = 3
        .SelTabs(0) = 8 * 567
        .SelTabs(1) = 10 * 567
        .SelTabs(2) = 16 * 567
        iRiderWidth = .SelTabs(0) - 350
        If iRiderWidth < 200 Then
            iRiderWidth = 200
        End If
    End With
Return

PrintTable:
    With rtfResult
        GoSub SetTabsMarks1
        .SelFontSize = 10
        .SelBold = True
        .SelText = EventName & " - " & dtaTest.Recordset.Fields("Code") & " " & Translate(dtaTest.Recordset.Fields("Test"), mcLanguage) & IIf(dtaTest.Recordset.Fields("Type_pre") <= 2, " - " & ClipAmp(tbsSelFin.SelectedItem.Caption), "")
        If dtaNotYet.Recordset.AbsolutePosition + 1 > 0 Then
            .SelText = " - #" & dtaNotYet.Recordset.AbsolutePosition + 1 & "/" & dtaNotYet.Recordset.RecordCount
        End If
        .SelText = vbCrLf
        .SelBold = True
        .SelUnderline = True
        .SelText = vbTab & vbTab & vbTab & vbTab & vbCrLf
        .SelBold = True
        .SelItalic = True
        .SelUnderline = False
        .SelText = "| " & Translate("NR", mcLanguage) & vbTab & "| " & Translate("Participant", mcLanguage) & vbTab & "| " & Translate("Mark", mcLanguage) & vbTab & "| " & Translate("Comment", mcLanguage) & vbTab & "|" & vbCrLf
        .SelFontSize = 6
        .SelBold = True
        .SelUnderline = True
        .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab & "|" & vbTab & "|" & vbCrLf
        .SelBold = False
        .SelItalic = False
        .SelUnderline = False
        .SelText = "| "
        .SelBold = True
        .SelText = cNr
        .SelBold = False
        .SelText = vbTab & "| "
        .SelBold = True
        .SelText = FitString(Me, cRider, iRiderWidth, 2)
        .SelBold = False
        .SelText = vbTab & "" & vbTab & "| " & Translate("Judge", mcLanguage) & ":" & vbTab & "|" & vbCrLf
        .SelUnderline = False
        .SelText = "| " & cColor & vbTab & "| "
        .SelBold = True
        .SelText = FitString(Me, cHorse, iRiderWidth, 2)
        .SelBold = False
        .SelText = vbTab & "" & vbTab & "|" & vbTab & "|" & vbCrLf
        .SelUnderline = True
        .SelText = "| " & cGroup & vbTab & "| " & FitString(Me, cTeam, iRiderWidth, 2) & vbTab & "" & vbTab & "|" & vbTab & "|" & vbCrLf
        GoSub SetTabsMarks2
        rstTestSections.MoveFirst
        iMaxfactor = 0
        For iColumn = 1 To iMaxSections
            .SelFontSize = 9
            .SelItalic = True
            .SelText = "|" & FitString(Me, Translate(rstTestSections.Fields("Name"), mcLanguage), iRiderWidth, 2) & IIf(rstTestSections.Fields("Factor") <> 1, " (x " & rstTestSections.Fields("Factor") & ")", "") & vbTab & "|" & vbTab & "|" & vbTab & "|" & vbCrLf
            iMaxfactor = iMaxfactor + rstTestSections.Fields("Factor")
            .SelFontSize = 9
            .SelItalic = False
            .SelUnderline = False
            .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab & "|" & vbCrLf
            .SelUnderline = True
            .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab & "|" & vbCrLf
            rstTestSections.MoveNext
        Next iColumn
        If IsNull(dtaTest.Recordset.Fields("Out_fin")) = False Then
            iMaxfactor = iMaxfactor - dtaTest.Recordset.Fields("Out_fin")
        End If
        .SelFontSize = 10
        .SelBold = True
        .SelText = "|" & Translate("Total", mcLanguage) & vbTab & "|" & vbTab & "| :" & iMaxfactor & " =" & vbTab & "|" & vbCrLf
        .SelFontSize = 9
        .SelBold = True
        .SelUnderline = False
        .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab & "|" & vbCrLf
        .SelBold = True
        .SelUnderline = True
        .SelText = "|" & vbTab & "|" & vbTab & "|" & vbTab & "|" & vbCrLf
        .SelUnderline = False
        .SelItalic = True
        .SelFontSize = 8
        .SelText = App.EXEName & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") & " - " & App.LegalCopyright & " - " & Translate("Composed", mcLanguage) & " " & Format$(Now, "d mmmm yyyy hh:mm:ss") & vbCrLf
        .SelItalic = False
        
        If Not (dtaNotYet.Recordset.AbsolutePosition + 1) Mod iMaxRows = 0 Then
            For iColumn = 1 To IIf(iMaxRows = 2, 4, 2)
                .SelText = vbCrLf
            Next iColumn
        End If
    End With
Return

End Sub
Sub PrintStartForm(Optional iEquipmentCheck As Integer = 0)
    Dim iKey As Integer
    Dim iMaxRows As Integer
    Dim iTemp As Integer
    Dim iFound As Integer
    Dim rstTime As DAO.Recordset
    Dim rstEntry As DAO.Recordset
    Dim rstWithDrawn As DAO.Recordset
    Dim rstPosition As DAO.Recordset
    Dim rstEquip As DAO.Recordset

    Dim dtaTemp As data

    Dim cNr As String
    Dim cQry As String
    Dim cGroup As String
    Dim cColor As String
    Dim cRider As String
    Dim cHorse As String
    Dim cClub As String
    Dim cClass As String
    Dim cTeam As String
    Dim cHand As String
    Dim cEquip As String
    Dim cPosition As String
    Dim cOldSta As String
    Dim iCounter As Integer
    Dim iGroup As Integer
    Dim iWithDrawn As Integer
    Dim cTemp As String
    Dim iEquipmentPercentage As Integer
    Dim iNoStart As Integer
    Dim iLog As Integer
    
    Randomize Timer

    If fraMarks.Visible = True Then
        If dtaNotYet.Recordset.RecordCount = 0 Then
            MsgBox Translate("No list of relevant participants available", mcLanguage), vbInformation
            Exit Sub
        End If
    ElseIf fraTime.Visible = True Then
        If dtaNotYet.Recordset.RecordCount = 0 And dtaAlready.Recordset.RecordCount = 0 Then
            MsgBox Translate("No list of relevant participants available", mcLanguage), vbInformation
            Exit Sub
        End If
    End If

    Set rstEntry = mdbMain.OpenRecordset("SELECT DISTINCT Group FROM Entries WHERE Code='" & TestCode & "' AND Status=0")
    If rstEntry.RecordCount > 0 Then
        rstEntry.MoveLast
        If rstEntry.RecordCount = 1 Then
            iGroup = rstEntry.Fields("Group")
        End If
    End If
    rstEntry.Close
    Set rstEntry = Nothing

    If TestStatus = 0 And iGroup = 1 Then
        iKey = MsgBox(Translate("Do you want to compose the starting order first?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton2)
        If iKey = vbYes Then
           cmdComposeGroups_Click
        End If
    End If

    iMaxRows = 25
    iCounter = 0

    If iEquipmentCheck = 2 Then 'only participants to be checked
        PrintRtfHeader Translate("Equipment Check", mcLanguage) & " - " & Format$(Now, "ddd dd-mm hh:mm")
    ElseIf iEquipmentCheck = 1 Then 'complete starting order
        PrintRtfHeader Translate("Starting Order", mcLanguage) & "/" & Translate("Equipment Check", mcLanguage) & " - " & Format$(Now, "ddd dd-mm hh:mm")
        Else
        PrintRtfHeader Translate("Starting Order", mcLanguage) & " - " & Format$(Now, "ddd dd-mm hh:mm")
    End If

    If iEquipmentCheck > 0 Then
        Set rstEquip = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=0 AND Check>0")
        If rstEquip.RecordCount > 0 Then
            iKey = MsgBox("Do you want to re-assign participants to the equipment check?", vbQuestion + vbYesNo + vbDefaultButton2)
        Else
            iKey = vbYes
        End If
        If iKey = vbYes Then
            cTemp = GetVariable("Equipmentcheck")
            If cTemp = "" Then cTemp = "25"
            Do
                cTemp = InputBox("How much % of the participants have to be checked?", "", cTemp)
                iTemp = InStr(cTemp, ".")
                If iTemp > 0 Then cTemp = Left$(cTemp, iTemp - 1)
                iTemp = InStr(cTemp, ",")
                If iTemp > 0 Then cTemp = Left$(cTemp, iTemp - 1)
                cTemp = Val(cTemp)
            Loop While Val(cTemp) < 0 Or Val(cTemp) > 100 Or cTemp = ""

            SetVariable "Equipmentcheck", CInt(cTemp)
            iEquipmentPercentage = Val(cTemp)

            mdbMain.Execute "UPDATE Entries SET Entries.[Check] = 0 WHERE Entries.Code='" & TestCode & "' AND Entries.Status=0"
            mdbMain.Execute "UPDATE Entries SET Entries.[TimeStamp] = NOW() WHERE Entries.Code='" & TestCode & "' AND Entries.Status=0 AND ISNULL(Entries.TimeStamp)"

            Randomize Timer

            Set rstEquip = mdbMain.OpenRecordset("SELECT TOP " & iEquipmentPercentage & " PERCENT Check,TimeStamp FROM Entries WHERE Code='" & TestCode & "' AND Status=0 ORDER BY Rnd(TimeStamp)")
            If rstEquip.RecordCount > 0 Then
                With rstEquip
                    Do While Not .EOF
                        .Edit
                        .Fields("Check") = 1
                        .Fields("TimeStamp") = Now
                        .Update
                        .MoveNext
                    Loop
                End With
            End If
        End If
        rstEquip.Close
        Set rstEquip = Nothing
    End If

    GoSub SetTabsMarks

    With rtfResult
        .SelItalic = False
        .SelBold = False
        .SelFontSize = 10
        .SelUnderline = True
        For iTemp = 1 To .SelTabCount
            .SelText = vbTab
        Next iTemp
        .SelText = vbCrLf
        .SelUnderline = False
        .SelText = vbCrLf
        .SelItalic = True
        If TestStatus = 0 Then
            .SelText = Left$(UCase$(Translate("Sequence", mcLanguage)), 3) & vbTab
        Else
            .SelText = Left$(UCase$(Translate("Position", mcLanguage)), 3) & vbTab
        End If
        If TestStatus = 0 Then
            .SelText = UCase$(Translate("Grp", mcLanguage))
            .SelText = vbTab
        End If
        If fraMarks.Visible = True Then
            .SelText = IIf(chkRein.Enabled = True And chkRein.Value = 1, UCase$(Translate("L/R", mcLanguage)), "") & vbTab
        End If
        .SelText = IIf(chkColor.Value = 1 And miUseColors = 1, UCase$(Translate("Clr", mcLanguage)), "") & vbTab
        .SelText = UCase$(Translate("Nr", mcLanguage)) & vbTab & UCase$(Translate("participant", mcLanguage)) & vbTab & vbCrLf
        .SelItalic = False
        .SelUnderline = True
        For iTemp = 1 To .SelTabCount
            .SelText = vbTab
        Next iTemp
        .SelText = vbCrLf
        .SelUnderline = False
        .SelItalic = False
        .SelBold = False
        .SelFontSize = 11
        .SelText = vbCrLf
    End With

    If fraTime.Visible = True Then
        iKey = MsgBox(Translate("Randomize starting order for participants with equal times or no times?", mcLanguage), vbYesNo + vbQuestion)
        If iKey = vbYes Then
            Randomize Timer
            cQry = "SELECT Entries.*,Alltimes,Disq,NoStart "
            cQry = cQry & " FROM Entries "
            cQry = cQry & " LEFT JOIN Results ON (Entries.STA = Results.STA) AND (Entries.Code = Results.Code) "
            cQry = cQry & " WHERE Entries.Code='" & TestCode & "' "
            cQry = cQry & " AND (IsNull(Disq) Or Disq = 0)"
            cQry = cQry & " AND (IsNull(NoStart) Or NoStart = 0)"
            cQry = cQry & " ORDER BY Alltimes&''='',Alltimes DESC, Val(Entries.STA) MOD ((RND(1)*11)+1)"
            Set rstTime = mdbMain.OpenRecordset(cQry)
            If rstTime.RecordCount > 0 Then
                With rstTime
                    Do While Not .EOF
                        Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Sta='" & .Fields("Sta") & "'")
                        If rstEntry.RecordCount > 0 Then
                            With rstEntry
                                .Edit
                                .Fields("Position") = rstTime.AbsolutePosition * 2 + 1
                                .Fields("Group") = 0
                                .Update
                            End With
                        End If
                        .MoveNext
                    Loop
                End With
                rstEntry.Close
                Set rstEntry = Nothing
            End If
            cmdComposeGroups_Click
        Else
            cQry = "SELECT Entries.*,Alltimes,Disq "
            cQry = cQry & " FROM Entries "
            cQry = cQry & " LEFT JOIN Results ON (Entries.STA = Results.STA) AND (Entries.Code = Results.Code) "
            cQry = cQry & " WHERE Entries.Code='" & TestCode & "' "
            cQry = cQry & " AND (IsNull(Disq) Or Disq = 0)"
            cQry = cQry & " AND (IsNull(NoStart) Or NoStart = 0)"
            cQry = cQry & " ORDER BY ISNULL(Results.AllTimes), Results.AllTimes DESC, Entries.Group, Entries.Position;"
            Set rstTime = mdbMain.OpenRecordset(cQry)
        End If
        If rstTime.RecordCount > 0 Then
            With rstTime
                .MoveFirst
                Do While Not .EOF
                    If IsNull(rstTime.Fields("Disq")) Or rstTime.Fields("Disq") = 0 Then
                        iFound = True
                        dtaNotYet.Recordset.FindFirst "Sta LIKE '" & .Fields("Sta") & "'"
                        If dtaNotYet.Recordset.NoMatch = True Then
                            dtaAlready.Recordset.FindFirst "Marks.Sta LIKE '" & .Fields("Sta") & "'"
                            If dtaAlready.Recordset.NoMatch = True Then
                                iFound = False
                            Else
                                Set dtaTemp = dtaAlready
                            End If
                        Else
                            Set dtaTemp = dtaNotYet
                        End If
                        If iFound = True Then
                            With dtaTemp
                                cNr = rstTime.Fields("Sta")
                                cGroup = IIf(rstTime.Fields("Group") > 0, Format$(rstTime.Fields("Group"), "00"), "  ")
                                If miUseColors = 1 Then
                                    cColor = UCase$(Left$(rstTime.Fields("Color") & "  ", 2))
                                End If
                                cRider = .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")

                                'Mark riders who have not declared their presence (IPZV only):
                                If .Recordset.Fields("pStatus") = 0 And mcVersionSwitch = "ipzv" Then
                                    cRider = "? " & .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")
                                Else
                                    cRider = .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")
                                End If

                                If .Recordset.Fields("Class") & "" <> "" Then
                                    cRider = cRider & " [" & .Recordset.Fields("Class") & "]"
                                End If
                                cHorse = .Recordset.Fields("Name_Horse") & ""
                                If miShowHorseId <> 0 Then
                                    cHorse = cHorse & " [" & GetHorseId(.Recordset) & "]"
                                End If
                                If miShowHorseAge <> 0 Then
                                    cHorse = cHorse & " " & getHorseAge(.Recordset)
                                End If
                                
                                If miShowRidersTeam <> 0 Then
                                    cTeam = UCase$(Left$(GetRidersTeam(.Recordset), 2))
                                Else
                                    cTeam = ""
                                End If
                                'IPZV LK output:
                                If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                                    cTeam = Trim$(cTeam & " " & GetRidersLk(.Recordset))
                                End If
                                cHand = IIf(rstTime.Fields("RR") = True, Left$(Translate("Right", mcLanguage), 1), Left$(Translate("Left", mcLanguage), 1))
                                If iEquipmentCheck > 0 Then
                                    cEquip = IIf(.Recordset.Fields("Check") > 0, "[" & Left$(Translate("Equipment check", mcLanguage), 1) & "]", "")
                                End If
                            End With
                            GoSub PrintParticipant
                            If iEquipmentCheck <> 2 Or (iEquipmentCheck = 2 And cEquip <> "") Then
                                With rtfResult
                                    .SelText = vbTab & vbTab & vbTab & vbTab & vbTab & Replace(Replace(rstTime.Fields("Alltimes") & "", ",", "."), "99.99", " -.- ") & vbCrLf
                                End With
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End With
            rstTime.Close
            Set rstTime = Nothing
        End If
        
        cQry = "SELECT Entries.*,Alltimes,Disq "
        cQry = cQry & " FROM Entries "
        cQry = cQry & " LEFT JOIN Results ON (Entries.STA = Results.STA) AND (Entries.Code = Results.Code) "
        cQry = cQry & " WHERE Entries.Code='" & TestCode & "' "
        cQry = cQry & " AND (IsNull(Disq) Or Disq = 0)"
        cQry = cQry & " AND (NoStart = -1)"
        cQry = cQry & " ORDER BY ISNULL(Results.AllTimes), Results.AllTimes DESC, Entries.Group, Entries.Position;"
        Set rstTime = mdbMain.OpenRecordset(cQry)

        If rstTime.RecordCount > 0 Then
            iNoStart = True
            With rtfResult
                .SelText = vbCrLf & vbCrLf
            End With
            With rstTime
                .MoveFirst
                Do While Not .EOF
                    If IsNull(rstTime.Fields("Disq")) Or rstTime.Fields("Disq") = 0 Then
                        iFound = True
                        dtaNotYet.Recordset.FindFirst "Sta LIKE '" & .Fields("Sta") & "'"
                        If dtaNotYet.Recordset.NoMatch = True Then
                            dtaAlready.Recordset.FindFirst "Marks.Sta LIKE '" & .Fields("Sta") & "'"
                            If dtaAlready.Recordset.NoMatch = True Then
                                iFound = False
                            Else
                                Set dtaTemp = dtaAlready
                            End If
                        Else
                            Set dtaTemp = dtaNotYet
                        End If
                        If iFound = True Then
                            With dtaTemp
                                cNr = rstTime.Fields("Sta")
                                cColor = ""
                                cGroup = ""
                                cRider = .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")

                                'Mark riders who have not declared their presence (IPZV only):
                                If .Recordset.Fields("pStatus") = 0 And mcVersionSwitch = "ipzv" Then
                                    cRider = "? " & .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")
                                Else
                                    cRider = .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")
                                End If

                                If .Recordset.Fields("Class") & "" <> "" Then
                                    cRider = cRider & " [" & .Recordset.Fields("Class") & "]"
                                End If
                                cHorse = .Recordset.Fields("Name_Horse") & ""
                                If miShowHorseId <> 0 Then
                                    cHorse = cHorse & " [" & GetHorseId(.Recordset) & "]"
                                End If
                                If miShowHorseAge <> 0 Then
                                    cHorse = cHorse & " " & getHorseAge(.Recordset)
                                End If
                                If miShowRidersTeam <> 0 Then
                                    cTeam = UCase$(Left$(GetRidersTeam(.Recordset), 2))
                                Else
                                    cTeam = ""
                                End If
                                'IPZV LK output:
                                If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                                    cTeam = Trim$(cTeam & " " & GetRidersLk(.Recordset))
                                End If
                                cHand = ""
                                cEquip = ""
                            End With
                            GoSub PrintParticipant
                            If iEquipmentCheck <> 2 Or (iEquipmentCheck = 2 And cEquip <> "") Then
                                With rtfResult
                                    .SelText = vbTab & vbTab & vbTab & vbTab & vbTab & Replace(Replace(rstTime.Fields("Alltimes") & "", ",", "."), "99.99", " -.- ") & vbCrLf
                                End With
                            End If
                        End If
                    End If
                    .MoveNext
                Loop
            End With
            rstTime.Close
            Set rstTime = Nothing
        End If
        
        ChangeCaption True
    Else
        If dtaNotYet.Recordset.RecordCount > 0 Then
            With dtaNotYet
                .Recordset.MoveLast
                If TestStatus = 0 Then
                    .Recordset.MoveFirst
                    Do While Not .Recordset.EOF
                        cNr = .Recordset.Fields("Sta")
                        If TestStatus = 0 Then
                            cGroup = IIf(.Recordset.Fields("Group") > 0, Format$(.Recordset.Fields("Group"), "00"), "  ")
                        End If
                        If miUseColors = 1 Then
                            cColor = UCase$(Left$(.Recordset.Fields("Color") & "  ", 2))
                        End If
                        cRider = .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")

                        'Mark riders who have not declared their presence (IPZV only):
                        If .Recordset.Fields("pStatus") = 0 And mcVersionSwitch = "ipzv" Then
                            cRider = "? " & .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")
                        Else
                            cRider = .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")
                        End If


                        If .Recordset.Fields("Class") & "" <> "" Then
                            cRider = cRider & " [" & .Recordset.Fields("Class") & "]"
                        End If
                        cHorse = .Recordset.Fields("Name_Horse") & ""
                        If miShowHorseId <> 0 Then
                            cHorse = cHorse & " [" & GetHorseId(.Recordset) & "]"
                        End If
                        If miShowHorseAge <> 0 Then
                            cHorse = cHorse & " " & getHorseAge(.Recordset)
                        End If
                        If miShowRidersTeam <> 0 Then
                            cTeam = UCase$(Left$(GetRidersTeam(.Recordset), 2))
                        Else
                            cTeam = ""
                        End If
                        'IPZV LK output:
                        If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                            cTeam = Trim$(cTeam & " " & GetRidersLk(.Recordset))
                        End If
                        cHand = IIf(.Recordset.Fields("RR") = True, Left$(Translate("Right", mcLanguage), 1), Left$(Translate("Left", mcLanguage), 1))
                        If iEquipmentCheck > 0 Then
                            cEquip = IIf(.Recordset.Fields("Check") > 0, "[" & Left$(Translate("Equipment check", mcLanguage), 1) & "]", "")
                        End If
                        GoSub PrintParticipant
                        .Recordset.MoveNext
                    Loop
                Else
                    If miFinalsSequence <> 0 Then
                        .Recordset.MoveFirst
                        Do While Not .Recordset.EOF
                            cNr = .Recordset.Fields("Sta")
                            If TestStatus = 0 Then
                                cGroup = IIf(.Recordset.Fields("Group") > 0, Format$(.Recordset.Fields("Group"), "00"), "  ")
                            End If
                            If miUseColors = 1 Then
                                cColor = UCase$(Left$(.Recordset.Fields("Color") & "  ", 2))
                            End If
                            cRider = .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")
                            If .Recordset.Fields("Class") & "" <> "" Then
                                cRider = cRider & " [" & .Recordset.Fields("Class") & "]"
                            End If
                            cHorse = .Recordset.Fields("Name_Horse")
                            If miShowHorseId <> 0 Then
                                cHorse = cHorse & " [" & GetHorseId(.Recordset) & "]"
                            End If
                            If miShowHorseAge <> 0 Then
                                cHorse = cHorse & " " & getHorseAge(.Recordset)
                            End If
                            If miShowRidersTeam <> 0 Then
                                cTeam = UCase$(Left$(GetRidersTeam(.Recordset), 2))
                            Else
                                cTeam = ""
                            End If
                            'IPZV LK output:
                            If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                                cTeam = Trim$(cTeam & " " & GetRidersLk(.Recordset))
                            End If
                            cHand = IIf(.Recordset.Fields("RR") = True, Left$(Translate("Right", mcLanguage), 1), Left$(Translate("Left", mcLanguage), 1))
                            GoSub PrintParticipant
                            .Recordset.MoveNext
                        Loop
                    Else
                        Do While Not .Recordset.BOF
                            cNr = .Recordset.Fields("Sta")
                            If TestStatus = 0 Then
                                cGroup = IIf(.Recordset.Fields("Group") > 0, Format$(.Recordset.Fields("Group"), "00"), "  ")
                            End If
                            If miUseColors = 1 Then
                                cColor = UCase$(Left$(.Recordset.Fields("Color") & "  ", 2))
                            End If
                            cRider = .Recordset.Fields("Name_First") & " " & .Recordset.Fields("Name_Last")
                            If .Recordset.Fields("Class") & "" <> "" Then
                                cRider = cRider & " [" & .Recordset.Fields("Class") & "]"
                            End If
                            cHorse = .Recordset.Fields("Name_Horse")
                            If miShowHorseId <> 0 Then
                                cHorse = cHorse & " [" & GetHorseId(.Recordset) & "]"
                            End If
                            If miShowHorseAge <> 0 Then
                                cHorse = cHorse & " " & getHorseAge(.Recordset)
                            End If
                            If miShowRidersTeam <> 0 Then
                                cTeam = UCase$(Left$(GetRidersTeam(.Recordset), 2))
                            Else
                                cTeam = ""
                            End If
                            'IPZV LK output:
                            If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                                cTeam = Trim$(cTeam & " " & GetRidersLk(.Recordset))
                            End If
                            cHand = IIf(.Recordset.Fields("RR") = True, Left$(Translate("Right", mcLanguage), 1), Left$(Translate("Left", mcLanguage), 1))
                            GoSub PrintParticipant
                            .Recordset.MovePrevious
                        Loop
                    End If
                    
                    'Print note about reverse order only when it applies:
                    If miFinalsSequence = 0 Then
                        With rtfResult
                            .SelText = vbCrLf
                            .SelItalic = True
                            .SelText = Translate("Participants will enter the track in reverse order.", mcLanguage) & vbCrLf
                            .SelItalic = False
                        End With
                    End If
                End If
                If TestStatus > 0 Then
                    Set rstWithDrawn = mdbMain.OpenRecordset("SELECT * FROM Results INNER JOIN Participants ON Results.STA=Participants.STA WHERE Results.Code='" & TestCode & "' AND Results.Status=" & TestStatus & " AND Results.Disq=-2 ORDER BY Results.STA")
                    If rstWithDrawn.RecordCount > 0 Then
                        iWithDrawn = True
                        With rtfResult
                            .SelText = vbCrLf
                            .SelBold = True
                            .SelText = Translate("Withdrawn", mcLanguage) & vbCrLf
                            .SelBold = False
                        End With
                        Do While Not rstWithDrawn.EOF
                            With rstWithDrawn
                                cNr = .Fields("Results.Sta")
                                If TestStatus = 0 Then
                                    cGroup = ""
                                End If
                                If miUseColors = 1 Then
                                    cColor = ""
                                End If
                                cRider = GetPersonsName(.Fields("PersonId"))
                                cHorse = GetHorsesName(.Fields("HorseId"))
                                If miShowHorseId <> 0 Then
                                    cHorse = cHorse & " [" & GetHorseId(rstWithDrawn) & "]"
                                End If
                                If miShowRidersTeam <> 0 Then
                                    cTeam = UCase$(Left$(GetRidersTeam(rstWithDrawn), 2))
                                Else
                                    cTeam = ""
                                End If
                                'IPZV LK output:
                                If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                                    cTeam = Trim$(cTeam & " " & GetRidersLk(rstWithDrawn))
                                End If
                                cHand = "-"
                                GoSub PrintParticipant
                                .MoveNext
                            End With
                        Loop
                    End If
                    rstWithDrawn.Close
                    Set rstWithDrawn = Nothing
                End If
            End With
        End If
    End If
    Set rstPosition = Nothing
    
    If iEquipmentCheck <> 0 Then
        With rtfResult
            .SelText = vbCrLf
            .SelItalic = True
            .SelFontSize = 9
            .SelText = "[" & Left$(Translate("Equipment check", mcLanguage), 1) & "]" & " = " & Translate("Equipment Check", mcLanguage) & vbCrLf
            .SelItalic = False
        End With
    End If

    PrintRtfFooter Translate("Starting Order", mcLanguage)

Exit Sub

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        .SelTabCount = 6
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 2.5 * 567
        .SelTabs(2) = 3.5 * 567
        .SelTabs(3) = 5 * 567
        .SelTabs(4) = 6 * 567
        .SelTabs(5) = 16 * 567
    End With
Return

ComposeParticipant:
Return

PrintParticipant:
    iCounter = iCounter + 1
    If iEquipmentCheck <> 2 Or (iEquipmentCheck = 2 And cEquip <> "") Then
        With rtfResult
            If Left$(cColor, 2) = Left$(TestColors, 2) Or (iEquipmentCheck = 2 And cEquip <> "") Then
                .SelFontSize = 6
                .SelText = vbCrLf
                .SelFontSize = 11
            End If
            .SelFontSize = 11
            .SelBold = True
            If iWithDrawn = False And iNoStart = False Then
                If TestStatus = 0 Then
                    .SelText = Format$(iCounter, "00")
                Else
                    Set rstPosition = mdbMain.OpenRecordset("SELECT Position FROM Results WHERE Code='" & TestCode & "' AND Status=0 AND STA='" & dtaNotYet.Recordset.Fields("Sta") & " '")
                    If rstPosition.RecordCount > 0 Then
                        .SelText = Format$(rstPosition.Fields("Position") & "", "00")
                    Else
                        .SelText = "--"
                    End If
                    rstPosition.Close
                End If
            End If
            If iNoStart = True Then
                .SelText = Translate("No Start", mcLanguage)
                .SelBold = False
            Else
                .SelText = vbTab
            End If
            If TestStatus = 0 Then
                .SelText = cGroup
                If cEquip <> "" Then
                    .SelText = " " & cEquip
                End If
                .SelText = vbTab
            End If
            If fraMarks.Visible = True Then
                .SelText = IIf(chkRein.Enabled = True And chkRein.Value = 1, cHand, "")
                .SelText = vbTab
            End If
            .SelText = IIf(chkColor.Value = 1 And miUseColors = 1, cColor, "")
            .SelText = vbTab
            .SelBold = False
            .SelText = cNr
            .SelText = vbTab
            If iNoStart = True Then
            Else
               .SelBold = True
            End If
            .SelText = cRider
            .SelText = " - "
            .SelText = cHorse
            If miShowRidersTeam <> 0 And cTeam <> "" Then
                .SelText = " - "
                .SelText = cTeam
            End If
            .SelBold = False
            .SelText = vbCrLf
            
             'Provide Information for Logging-DB:
             If miWriteLogDB Then
                 If iWithDrawn = 0 Then
                     iLog = WriteLogDBStart(EventName, TestCode, TestStatus, cNr, iCounter, 0, "")
                 End If
             End If
             
        End With
    End If
Return

End Sub
Public Function StartedParticipants(Status As Integer) As String
    Dim cQry As String
    Dim iTemp As Integer
    
    On Local Error Resume Next
    
    'select started riders
    cQry = "SELECT Marks.Sta "
    cQry = cQry & " & '  -  ' & Persons.Name_First "
    cQry = cQry & " & ' ' & Persons.Name_Last "
    cQry = cQry & " & IIF(Participants.Class<>'',' [' &  Participants.Class & ']','')"
    If frmMain.chkTeam.Value = 1 Then
        cQry = cQry & " & IIF(Participants.Club<>'',' [' &  Participants.Club & ']','')"
        cQry = cQry & " & IIF(Participants.Team<>'',' [' &  Participants.Team & ']','')"
    End If
    cQry = cQry & " & ' - ' & Horses.Name_Horse "
    If frmMain.chkFeifId.Value = 1 Then
        cQry = cQry & " & ' [' & Horses.FEIFId & ']' "
    End If
    cQry = cQry & " as cList,"
    cQry = cQry & " Format(Marks.Score,'" & TestTotalFormat & "') as cTotal,"
    
    If fraMarks.Visible = True Then
        cQry = cQry & " Format(Marks.Mark1,'" & Me.TestMarkFormat & "') "
        If dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus)) > 1 Then
            For iTemp = 2 To dtaTestInfo.Recordset.Fields("Num_J_" & Format$(TestStatus))
                cQry = cQry & " & ' - ' & Format(Marks.Mark" & Format$(iTemp) & ",'" & Me.TestMarkFormat & "')"
            Next iTemp
        End If
    ElseIf fraTime.Visible = True Then
        cQry = cQry & " Format(Marks.Mark1,'" & Me.TestTimeFormat & "') "
    End If
    
    cQry = cQry & " As cMarks"
    cQry = cQry & ",Results.Score"
    cQry = cQry & ",Results.Disq"
    cQry = cQry & ",Results.Alltimes"
    cQry = cQry & ",Results.Position"
    cQry = cQry & ",Marks.Section"
    cQry = cQry & ",Marks.Score"
    cQry = cQry & ",Marks.Out"
    cQry = cQry & ",Marks.Sta"
    cQry = cQry & ",Marks.Flag"
    cQry = cQry & ",Marks.Mark1"
    cQry = cQry & ",Marks.Mark2"
    cQry = cQry & ",Marks.Mark3"
    cQry = cQry & ",Marks.Mark4"
    cQry = cQry & ",Marks.Mark5"
    cQry = cQry & ",Participants.Sta"
    cQry = cQry & ",Participants.Club"
    cQry = cQry & ",Participants.Team"
    cQry = cQry & ",Participants.Class"
    cQry = cQry & ",Participants.Status AS pStatus"
    cQry = cQry & ",Testsections.Name"
    cQry = cQry & ",Testsections.Section"
    cQry = cQry & ",Testsections.Factor"
    cQry = cQry & ",Persons.Name_First"
    cQry = cQry & ",Persons.Name_Middle"
    cQry = cQry & ",Persons.Name_Last"
    cQry = cQry & ",Horses.Name_Horse"
    cQry = cQry & ",Horses.HorseID"
    cQry = cQry & ",Horses.FEIFID"
    cQry = cQry & ",Horses.Sex_Horse"
    cQry = cQry & " FROM Testsections "
    cQry = cQry & " INNER JOIN (Results "
    cQry = cQry & " INNER JOIN (((Marks "
    cQry = cQry & " INNER JOIN Participants "
    cQry = cQry & " ON Marks.STA = Participants.STA) "
    cQry = cQry & " INNER JOIN Persons "
    cQry = cQry & " ON Participants.PersonID = Persons.PersonID) "
    cQry = cQry & " INNER JOIN Horses "
    cQry = cQry & " ON Participants.HorseID = Horses.HorseID) "
    cQry = cQry & " ON Results.Status = Marks.Status "
    cQry = cQry & " AND Results.STA = Marks.STA "
    cQry = cQry & " AND Results.Code = Marks.Code) "
    If Status = 3 Then
         cQry = cQry & " ON Testsections.Status = Marks.Status-2"
    ElseIf Status = 2 Then
         cQry = cQry & " ON Testsections.Status = Marks.Status-1"
    Else
         cQry = cQry & " ON Testsections.Status = Marks.Status"
    End If
    cQry = cQry & " AND Testsections.Section = Marks.Section "
    cQry = cQry & " AND Testsections.Code = Marks.Code"
    cQry = cQry & " Where Marks.Status = " & Status
    cQry = cQry & " AND Marks.Code='" & Me.TestCode & "'"
    
    If fraMarks.Visible = True Then
        If Status = 0 And dtaTest.Recordset.Fields("Type_time") = 3 Then
            'MM: Corrected for FIPO 6.10
            cQry = cQry & " ORDER BY Results.Disq DESC,Results.Score DESC,Results.Alltimes DESC,Marks.STA,Testsections.Section;"
        ElseIf (Status = 0 And dtaTest.Recordset.Fields("Type_pre") = 2) Or (Status <> 0 And dtaTest.Recordset.Fields("Type_Final") = 2) Then
            cQry = cQry & " ORDER BY Results.Disq DESC,Results.Score,Results.Alltimes,Marks.STA,Testsections.Section;"
        Else
            cQry = cQry & " ORDER BY Results.Disq DESC,Results.Score DESC,Results.Alltimes,Marks.STA,Testsections.Section;"
        End If
    ElseIf fraTime.Visible = True Then
        cQry = cQry & " ORDER BY Results.Disq DESC,Results.Score=0 DESC,Results.Score,Results.Alltimes,Marks.STA,Testsections.Section;"
    End If
    
    StartedParticipants = cQry
    
End Function
Public Sub PrintLogFile()
    Dim iTemp As Integer
    
    Dim rstMarks As DAO.Recordset
    Dim cNr As String
    Dim cQry As String
    
    Dim iCounter As Integer
    
    cQry = "SELECT Participants.Sta,Marks.*,Persons.Name_First,Persons.Name_Last,Horses.Name_Horse,Testsections.Name,Results.Score"
    cQry = cQry & " FROM (((Participants "
    cQry = cQry & " INNER JOIN (Marks "
    cQry = cQry & " INNER JOIN Testsections ON Marks.Code = Testsections.Code "
    cQry = cQry & " AND Marks.Section = Testsections.Section "
    If TestStatus = 3 Then
         cQry = cQry & " AND Testsections.Status = Marks.Status-2"
    ElseIf TestStatus = 2 Then
         cQry = cQry & " AND Testsections.Status = Marks.Status-1"
    Else
         cQry = cQry & " AND Testsections.Status = Marks.Status"
    End If
    cQry = cQry & " AND Marks.Section = Testsections.Section) "
    cQry = cQry & " ON Participants.STA = Marks.STA) "
    cQry = cQry & " INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) "
    cQry = cQry & " INNER JOIN Horses ON Participants.HorseID = Horses.HorseID) "
    cQry = cQry & " INNER JOIN Results ON Marks.Status = Results.Status "
    cQry = cQry & " AND Marks.STA = Results.STA AND Marks.Code = Results.Code"
    cQry = cQry & " WHERE Marks.Code='" & TestCode & "' AND Marks.Status=" & TestStatus
    cQry = cQry & " ORDER BY Participants.Sta,Marks.Section;"

    Set rstMarks = mdbMain.OpenRecordset(cQry)
    If rstMarks.RecordCount > 0 Then
        
        PrintRtfHeader Translate("Log File", mcLanguage) & " - " & Format$(Now, "ddd dd-mm hh:mm")
        
        GoSub SetTabsMarks
        
        With rtfResult
            .SelItalic = False
            .SelBold = False
            .SelFontSize = 11
            .SelUnderline = True
            For iTemp = 1 To .SelTabCount
                .SelText = vbTab
            Next iTemp
            .SelText = vbCrLf
            .SelUnderline = False
            .SelText = vbCrLf
            .SelItalic = True
            .SelText = Translate("This is not a result list!", mcLanguage) & vbCrLf
            .SelItalic = False
            .SelUnderline = True
            For iTemp = 1 To .SelTabCount
                .SelText = vbTab
            Next iTemp
            .SelText = vbCrLf
            .SelUnderline = False
            .SelItalic = False
            .SelBold = False
            .SelFontSize = 12
            .SelText = vbCrLf
        End With
        
        Do While Not rstMarks.EOF
            With rtfResult
                .SelBold = False
                If rstMarks.Fields("Participants.Sta") <> cNr Then
                    .SelFontSize = 11
                    .SelText = rstMarks.Fields("Participants.Sta")
                    .SelText = vbTab
                    .SelText = rstMarks.Fields("Name_First") & " " & rstMarks.Fields("Name_Last")
                    .SelText = " - "
                    .SelText = rstMarks.Fields("Name_Horse")
                    .SelText = " - "
                    .SelText = Format$(rstMarks.Fields("Results.Score"), TestTotalFormat)
                    .SelText = vbCrLf
                End If
                cNr = rstMarks.Fields("Participants.Sta")
                .SelFontSize = 10
                .SelText = vbTab
                .SelText = UCase$(Left$(Translate(rstMarks.Fields("Name"), mcLanguage), 5))
                .SelText = vbTab
                .SelText = Format$(rstMarks.Fields("Mark1"), TestMarkFormat)
                .SelText = vbTab
                .SelText = Format$(rstMarks.Fields("Mark2"), TestMarkFormat)
                .SelText = vbTab
                .SelText = Format$(rstMarks.Fields("Mark3"), TestMarkFormat)
                .SelText = vbTab
                .SelText = Format$(rstMarks.Fields("Mark4"), TestMarkFormat)
                .SelText = vbTab
                .SelText = Format$(rstMarks.Fields("Mark5"), TestMarkFormat)
                .SelText = vbTab
                .SelText = Format$(rstMarks.Fields("Marks.Score"), TestTotalFormat)
                .SelText = vbTab
                .SelText = Format$(rstMarks.Fields("Timestamp"), "dd-mm-yy hh:mm:ss")
                .SelText = vbCrLf
            End With
            rstMarks.MoveNext
        Loop
        
        PrintRtfFooter Translate("Log File", mcLanguage)
    Else
        MsgBox Translate("No marks entered yet!", mcLanguage), vbInformation
    End If
    rstMarks.Close
    Set rstMarks = Nothing
    
Exit Sub

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        .SelTabCount = 9
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 3 * 567
        .SelTabs(2) = 4 * 567
        .SelTabs(3) = 5 * 567
        .SelTabs(4) = 6 * 567
        .SelTabs(5) = 7 * 567
        .SelTabs(6) = 9 * 567
        .SelTabs(7) = 12 * 567
        .SelTabs(8) = 16 * 567
    End With
Return

End Sub
Sub PrintAllMerge()
    Dim rstAll As DAO.Recordset
    Dim rstEntry As DAO.Recordset
    Dim cQry As String
    Dim cNr As String
    Dim cTemp As String
    Dim cFilename As String
    Dim iFileNum As Integer
        
   ReadIniFile gcIniFile, "Print", "Merge", cFilename
   If cFilename = "" Then
        cFilename = NameOfFile(mcDatabaseName) & ".Csv"
   End If
   
   On Local Error Resume Next
   With frmMain.CommonDialog1
        .CancelError = True
        .DefaultExt = ".Csv"
        .DialogTitle = "Select a folder"
        .Filter = "Merge files|*.Csv|"
        .FilterIndex = 1
        .Flags = cdlOFNHideReadOnly
        .FileName = cFilename
        .ShowOpen
        If Err = cdlCancel Then
            Exit Sub
        End If
        If .FileName <> "" Then
            cFilename = NameOfFile(.FileName) & ".Csv"
        End If
    End With
    On Local Error GoTo 0
    
    If cFilename = "" Or cFilename = Chr$(27) Then
    Else
        cQry = "SELECT Participants.*, Persons.*, Horses.*"
        cQry = cQry & " FROM (Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
        cQry = cQry & " ORDER BY Participants.Sta;"
    
        Set rstAll = mdbMain.OpenRecordset(cQry)
        If rstAll.RecordCount > 0 Then
            iFileNum = FreeFile
            Open cFilename For Output Access Write Shared As #iFileNum
            Print #iFileNum, Chr$(34); "Sta"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Rider"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Horse"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Club"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Team"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Class"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Tests"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Sex"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Year"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Color"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Breeder"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "Owner"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "F"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "M"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "FF"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "FM"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "MF"; Chr$(34);
            Print #iFileNum, mcExcelSeparator; Chr$(34); "MM"; Chr$(34)
            Do While Not rstAll.EOF
                With rtfResult
                    Print #iFileNum, Chr$(34) & rstAll.Fields("Sta") & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("Name_First") & " " & RTrim$(rstAll.Fields("Name_Middle") & " ") & rstAll.Fields("Name_Last") & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("Name_Horse") & " [" & GetHorseId(rstAll) & "]" & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & GetRidersClub(rstAll) & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & GetRidersTeam(rstAll) & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("Class") & Chr$(34);
                    
                    cTemp = ""
                    Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Sta='" & rstAll.Fields("Sta") & "' AND Status=0 ORDER BY Code")
                    If rstEntry.RecordCount > 0 Then
                        Do While Not rstEntry.EOF
                            If rstEntry.AbsolutePosition = 0 Then
                                cTemp = rstEntry.Fields("Code")
                            Else
                                cTemp = cTemp & ", " & rstEntry.Fields("Code")
                            End If
                            rstEntry.MoveNext
                        Loop
                    End If
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & cTemp & Chr$(34);
                    
                    cTemp = ""
                    Select Case rstAll.Fields("Sex_horse")
                    Case 1
                        cTemp = Translate("Stallion", mcLanguage)
                    Case 2
                        cTemp = Translate("Mare", mcLanguage)
                    Case 3
                        cTemp = Translate("Gelding", mcLanguage)
                    End Select
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & cTemp & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & Format$(rstAll.Fields("Birthday_Horse") & "", "YYYY") & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("Color") & "" & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("Breeder") & "" & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("Owner") & "" & Chr$(34);
                    
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("F") & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("M") & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("FF") & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("FM") & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("MF") & Chr$(34);
                    Print #iFileNum, mcExcelSeparator & Chr$(34) & rstAll.Fields("MM") & Chr$(34);
                    
                    Print #iFileNum, ""
                    rstEntry.Close
                End With
                rstAll.MoveNext
            Loop
            Close #iFileNum
        Else
            MsgBox Translate("No participants entered yet!", mcLanguage), vbInformation
        End If
        rstAll.Close
        Set rstEntry = Nothing
        Set rstAll = Nothing
        WriteIniFile gcIniFile, "Print", "Merge", cFilename
    End If
    
End Sub
Sub PrintAllPrinter()
    Dim iTemp As Integer
    
    Dim rstAll As DAO.Recordset
    Dim rstEntry As DAO.Recordset
    Dim cQry As String
    Dim cNr As String
        
    cQry = "SELECT Participants.*, Persons.*, Horses.*"
    cQry = cQry & " FROM (Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
    cQry = cQry & " ORDER BY Participants.Sta;"

    Set rstAll = mdbMain.OpenRecordset(cQry)
    If rstAll.RecordCount > 0 Then
        SetMouseHourGlass
        
        frmMain.Enabled = False
        
        With rtfResult
            .Text = ""
            .SelBold = True
            .SelFontSize = 18
            .SelText = EventName & vbCrLf
            .SelBold = True
            .SelFontSize = 18
            .SelText = Translate("Participants", mcLanguage) & vbCrLf
            .SelBold = False
            .SelFontSize = 11
        End With
    
        GoSub SetTabsParticipant
        
        With rtfResult
            .SelItalic = False
            .SelBold = False
            .SelFontSize = 11
            .SelUnderline = True
            For iTemp = 1 To .SelTabCount
                .SelText = vbTab
            Next iTemp
            .SelText = vbCrLf
            .SelUnderline = False
            .SelItalic = False
            .SelBold = False
            .SelFontSize = 12
            .SelText = vbCrLf
        End With
        
        Do While Not rstAll.EOF
            GoSub SetTabsParticipant
            With rtfResult
                If rstAll.AbsolutePosition Mod 10 = 1 Then
                    .SelFontSize = 6
                    .SelText = vbCrLf
                End If
                
                .SelBold = True
                .SelItalic = False
                .SelFontSize = 11
                .SelText = rstAll.Fields("Sta")
                .SelText = vbTab
                .SelText = rstAll.Fields("Name_First") & " " & RTrim$(rstAll.Fields("Name_Middle") & " ") & rstAll.Fields("Name_Last")
                If rstAll.Fields("Class") & "" <> "" Then
                    .SelText = " [" & rstAll.Fields("Class") & "]"
                End If
                .SelText = vbTab
                .SelText = rstAll.Fields("Name_Horse") & " [" & GetHorseId(rstAll) & "]"
                If miShowRidersClub <> 0 Then
                    .SelText = vbTab
                    .SelText = GetRidersClub(rstAll)
                End If
                If miShowRidersTeam <> 0 Then
                    .SelText = vbTab
                    .SelText = GetRidersTeam(rstAll)
                End If
                'IPZV LK output:
                If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                    .SelText = vbTab
                    .SelText = Left$(GetRidersLk(rstAll), 2)
                End If
                
                .SelBold = False
                .SelText = vbCrLf
                
                .SelFontSize = 9
                .SelText = vbTab
                Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Sta='" & rstAll.Fields("Sta") & "' AND Status=0 ORDER BY Code")
                If rstEntry.RecordCount > 0 Then
                    Do While Not rstEntry.EOF
                        If rstEntry.AbsolutePosition = 0 Then
                            .SelText = rstEntry.Fields("Code")
                        Else
                            .SelText = ", " & rstEntry.Fields("Code")
                        End If
                        rstEntry.MoveNext
                    Loop
                End If
                
                .SelText = vbTab
                Select Case rstAll.Fields("Sex_horse")
                Case 1
                    .SelText = Translate("Stallion", mcLanguage)
                Case 2
                    .SelText = Translate("Mare", mcLanguage)
                Case 3
                    .SelText = Translate("Gelding", mcLanguage)
                End Select
                
                .SelText = " - "
                .SelText = Format$(rstAll.Fields("Birthday_horse") & "", "YYYY")
                .SelText = " - "
                .SelText = rstAll.Fields("Color") & ""
                .SelText = vbCrLf
                
                If Len(rstAll.Fields("Owner") & rstAll.Fields("Breeder") & "") > 2 Then
                    .SelFontSize = 9
                    .SelText = vbTab
                    If Len(rstAll.Fields("Breeder") & "") > 1 Then
                        .SelItalic = True
                        .SelText = Translate("Breeder", mcLanguage) & ": "
                        .SelItalic = False
                        .SelText = rstAll.Fields("Breeder") & ""
                    End If
                    .SelText = vbTab
                    If Len(rstAll.Fields("Owner") & "") > 1 Then
                        .SelItalic = True
                        .SelText = Translate("Owner", mcLanguage) & ": "
                        .SelItalic = False
                        .SelText = rstAll.Fields("Owner") & ""
                    End If
                    .SelText = vbCrLf
                End If
                
                If Len(rstAll.Fields("F") & rstAll.Fields("M") & "") > 2 Then
                    GoSub SetTabsHorse
                    .SelFontSize = 9
                    
                    .SelText = vbTab
                    If rstAll.Fields("F") & "" <> "" Then
                        .SelText = "F: " & rstAll.Fields("F")
                    End If
                    .SelText = vbTab
                    If rstAll.Fields("M") & "" <> "" Then
                        .SelText = "M: " & rstAll.Fields("M")
                    End If
                    .SelText = vbCrLf
                    
                    If Len(rstAll.Fields("FF") & rstAll.Fields("FM") & "") > 2 Then
                        GoSub SetTabsParents
                        .SelText = vbTab
                        If rstAll.Fields("FF") & "" <> "" Then
                            .SelText = "-FF: " & rstAll.Fields("FF")
                        End If
                        .SelText = vbTab
                        If rstAll.Fields("FM") & "" <> "" Then
                            .SelText = "-FM: " & rstAll.Fields("FM")
                        End If
                        .SelText = vbCrLf
                    End If
                    
                    If Len(rstAll.Fields("MF") & rstAll.Fields("MM")) & "" > 2 Then
                        GoSub SetTabsParents
                        .SelText = vbTab
                        If rstAll.Fields("MF") & "" <> "" Then
                            .SelText = "-MF: " & rstAll.Fields("MF")
                        End If
                        .SelText = vbTab
                        If rstAll.Fields("MM") & "" <> "" Then
                            .SelText = "-MM: " & rstAll.Fields("MM")
                        End If
                        .SelText = vbCrLf
                    End If
                End If
                rstEntry.Close
            End With
            rstAll.MoveNext
        Loop
        
        PrintRtfFooter Translate("Participants", mcLanguage)
    Else
        MsgBox Translate("No participants entered yet!", mcLanguage), vbInformation
    End If
    SetMouseNormal
    frmMain.Enabled = True
    rstAll.Close
    Set rstEntry = Nothing
    Set rstAll = Nothing
    
Exit Sub

SetTabsParticipant:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        .SelTabCount = 4
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 7 * 567
        .SelTabs(2) = 13 * 567
        .SelTabs(3) = 16 * 567
    End With
Return

SetTabsHorse:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        .SelTabCount = 3
        .SelTabs(0) = 1.5 * 567
        .SelTabs(1) = 7.5 * 567
        .SelTabs(2) = 16 * 567
    End With
Return

SetTabsParents:
    With rtfResult
        .SelFontSize = 9
        .SelItalic = True
        .SelTabCount = 3
        .SelTabs(0) = 2 * 567
        .SelTabs(1) = 8 * 567
        .SelTabs(2) = 16 * 567
    End With
Return


End Sub
Sub CreateCombinationMenu()
    Dim rstComb As DAO.Recordset
    Dim iItem As Integer
    
    iItem = -1
    For iItem = mnuFilePrintCombComb.Count - 1 To 1 Step -1
        Unload Me.mnuFilePrintCombComb(iItem)
    Next iItem
    
    If TableExist(mdbMain, "Combinations") Then
        Set rstComb = mdbMain.OpenRecordset("SELECT Combination,Code FROM Combinations WHERE USERLEVEL>=0 ORDER BY Combination")
        If rstComb.RecordCount > 0 Then
           Do While Not rstComb.EOF
                If rstComb.Fields("Combination") & "" <> "" Then
                    If iItem > 0 Then
                       Load mnuFilePrintCombComb(mnuFilePrintCombComb.Count)
                    End If
                    iItem = iItem + 1
                    With mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1)
                        .Caption = Translate(rstComb.Fields("Combination"), mcLanguage)
                        .Visible = True
                    End With
                    ChangeTagItem mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1), "Comb", rstComb.Fields("Code")
                    ChangeTagItem mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1), "PopUp", "No"
                End If
                rstComb.MoveNext
           Loop
        Else
           MsgBox Translate("No list of combinations available.", mcLanguage), vbExclamation
        End If
        rstComb.Close
        Set rstComb = Nothing
    End If
    
    If iItem > 0 Then
       Load mnuFilePrintCombComb(mnuFilePrintCombComb.Count)
        iItem = iItem + 1
        With mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1)
            .Caption = "-"
            .Visible = True
        End With
        ChangeTagItem mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1), "Comb", "-"
        ChangeTagItem mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1), "PopUp", "No"
    End If
    
    If iItem > 0 Then
       Load mnuFilePrintCombComb(mnuFilePrintCombComb.Count)
    End If
    iItem = iItem + 1
    With mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1)
        .Caption = Translate("Team Combination", mcLanguage)
        .Visible = True
    End With
    ChangeTagItem mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1), "Comb", "Team"
    ChangeTagItem mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1), "PopUp", "No"
    
    If iItem > 0 Then
       Load mnuFilePrintCombComb(mnuFilePrintCombComb.Count)
    End If
    iItem = iItem + 1
    With mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1)
        .Caption = Translate("Club Combination", mcLanguage)
        .Visible = True
    End With
    ChangeTagItem mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1), "Comb", "Club"
    ChangeTagItem mnuFilePrintCombComb(mnuFilePrintCombComb.Count - 1), "PopUp", "No"
    
End Sub
Sub CreateTestMenu()
    Dim iQual As Integer
    Dim iItem As Integer
    Dim cPrevQual As String

    If mnuTestAll.Checked = True Or TableExist(mdbMain, "Testinfo") = False Then
        dtaTest.RecordSource = "SELECT * FROM Tests WHERE (Removed=False or ISNULL(Removed)) ORDER BY Qualification,Code"
    Else
        dtaTest.RecordSource = "SELECT Tests.* FROM Tests INNER JOIN TestInfo ON Tests.Code=TestInfo.Code WHERE TestInfo.Nr>0 AND (Removed=False or ISNULL(Removed)) ORDER BY TestInfo.Nr"
        dtaTest.Refresh
        If dtaTest.Recordset.RecordCount = 0 Then
            mnuTestAll.Checked = True
            dtaTest.RecordSource = "SELECT * FROM Tests WHERE (Removed=False or ISNULL(Removed)) ORDER BY Qualification,Code"
        End If
    End If
    dtaTest.Refresh
    If dtaTest.Recordset.RecordCount > 0 Then
       mnutestQual1.Visible = False
       mnuTestQual2.Visible = False
       mnuTestQual3.Visible = False
       mnuTestQual4.Visible = False
       mnuTestQual5.Visible = False
       mnuTestQual6.Visible = False
       mnuTestQual7.Visible = False
       mnuTestQual8.Visible = False
       mnuTestQual9.Visible = False
       mnuTestQual10.Visible = False
       mnutestQual11.Visible = False
       mnuTestQual12.Visible = False
       mnuTestQual13.Visible = False
       mnuTestQual14.Visible = False
       mnuTestQual15.Visible = False
       mnuTestQual16.Visible = False
       mnuTestQual17.Visible = False
       mnuTestQual18.Visible = False
       mnuTestQual19.Visible = False
       mnuTestQual20.Visible = False
       mnutestQual21.Visible = False
       mnuTestQual22.Visible = False
       mnuTestQual23.Visible = False
       mnuTestQual24.Visible = False
       mnuTestQual25.Visible = False
       mnuTestQual26.Visible = False
       mnuTestQual27.Visible = False
       mnuTestQual28.Visible = False
       mnuTestQual29.Visible = False
       
       For iItem = mnuTestQual1Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual1Test(iItem)
       Next iItem
       For iItem = mnuTestQual2Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual2Test(iItem)
       Next iItem
       For iItem = mnuTestQual3Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual3Test(iItem)
       Next iItem
       For iItem = mnuTestQual4Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual4Test(iItem)
       Next iItem
       For iItem = mnuTestQual5Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual5Test(iItem)
       Next iItem
       For iItem = mnuTestQual6Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual6Test(iItem)
       Next iItem
       For iItem = mnuTestQual7Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual7Test(iItem)
       Next iItem
       For iItem = mnuTestQual8Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual8Test(iItem)
       Next iItem
       For iItem = mnuTestQual9Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual9Test(iItem)
       Next iItem
       For iItem = mnuTestQual10Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual10Test(iItem)
       Next iItem
       For iItem = mnuTestQual11Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual11Test(iItem)
       Next iItem
       For iItem = mnuTestQual12Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual12Test(iItem)
       Next iItem
       For iItem = mnuTestQual13Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual13Test(iItem)
       Next iItem
       For iItem = mnuTestQual14Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual14Test(iItem)
       Next iItem
       For iItem = mnuTestQual15Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual15Test(iItem)
       Next iItem
       For iItem = mnuTestQual16Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual16Test(iItem)
       Next iItem
       For iItem = mnuTestQual17Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual17Test(iItem)
       Next iItem
       For iItem = mnuTestQual18Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual18Test(iItem)
       Next iItem
       For iItem = mnuTestQual19Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual19Test(iItem)
       Next iItem
       For iItem = mnuTestQual20Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual20Test(iItem)
       Next iItem
       For iItem = mnuTestQual21Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual21Test(iItem)
       Next iItem
       For iItem = mnuTestQual22Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual22Test(iItem)
       Next iItem
       For iItem = mnuTestQual23Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual23Test(iItem)
       Next iItem
       For iItem = mnuTestQual24Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual24Test(iItem)
       Next iItem
       For iItem = mnuTestQual25Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual25Test(iItem)
       Next iItem
       For iItem = mnuTestQual26Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual26Test(iItem)
       Next iItem
       For iItem = mnuTestQual27Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual27Test(iItem)
       Next iItem
       For iItem = mnuTestQual28Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual28Test(iItem)
       Next iItem
       For iItem = mnuTestQual29Test.Count - 1 To 1 Step -1
           Unload Me.mnuTestQual29Test(iItem)
       Next iItem
       
       If mnuTestAll.Checked = True Then
            iQual = 0
            iItem = -1
       Else
            iQual = 1
            iItem = -1
            With mnutestQual1
                .Caption = Translate("&This event", mcLanguage)
                .Visible = True
            End With
       End If
       Do While Not dtaTest.Recordset.EOF
          If dtaTest.Recordset.Fields("Code") & "" <> "" Then
            If mnuTestAll.Checked = True Then
              If dtaTest.Recordset.Fields("Qualification") <> cPrevQual Then
                 iQual = iQual + 1
                 iItem = -1
                 cPrevQual = dtaTest.Recordset.Fields("Qualification")
              End If
              iItem = iItem + 1
                Select Case iQual
                Case 1
                    With mnutestQual1
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                        .Visible = True
                    End With
                    If iItem > 0 Then
                       Load mnuTestQual1Test(mnuTestQual1Test.Count)
                    End If
                    With mnuTestQual1Test(mnuTestQual1Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 2
                   With mnuTestQual2
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual2Test(mnuTestQual2Test.Count)
                    End If
                    With mnuTestQual2Test(mnuTestQual2Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 3
                   With mnuTestQual3
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual3Test(mnuTestQual3Test.Count)
                    End If
                    With mnuTestQual3Test(mnuTestQual3Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 4
                   With mnuTestQual4
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                      .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual4Test(mnuTestQual4Test.Count)
                    End If
                    With mnuTestQual4Test(mnuTestQual4Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 5
                   With mnuTestQual5
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual5Test(mnuTestQual5Test.Count)
                    End If
                    With mnuTestQual5Test(mnuTestQual5Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 6
                   With mnuTestQual6
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual6Test(mnuTestQual6Test.Count)
                    End If
                    With mnuTestQual6Test(mnuTestQual6Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 7
                   With mnuTestQual7
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual7Test(mnuTestQual7Test.Count)
                    End If
                    With mnuTestQual7Test(mnuTestQual7Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 8
                   With mnuTestQual8
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual8Test(mnuTestQual8Test.Count)
                    End If
                    With mnuTestQual8Test(mnuTestQual8Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 9
                   With mnuTestQual9
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual9Test(mnuTestQual9Test.Count)
                    End If
                    With mnuTestQual9Test(mnuTestQual9Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 10
                   With mnuTestQual10
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual10Test(mnuTestQual10Test.Count)
                    End If
                    With mnuTestQual10Test(mnuTestQual10Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 11
                    With mnutestQual11
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                        .Visible = True
                    End With
                    If iItem > 0 Then
                       Load mnuTestQual11Test(mnuTestQual11Test.Count)
                    End If
                    With mnuTestQual11Test(mnuTestQual11Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 12
                   With mnuTestQual12
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual12Test(mnuTestQual12Test.Count)
                    End If
                    With mnuTestQual12Test(mnuTestQual12Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 13
                   With mnuTestQual13
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual13Test(mnuTestQual13Test.Count)
                    End If
                    With mnuTestQual13Test(mnuTestQual13Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 14
                   With mnuTestQual14
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                      .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual14Test(mnuTestQual14Test.Count)
                    End If
                    With mnuTestQual14Test(mnuTestQual14Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 15
                   With mnuTestQual15
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual15Test(mnuTestQual15Test.Count)
                    End If
                    With mnuTestQual15Test(mnuTestQual15Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 16
                   With mnuTestQual16
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual16Test(mnuTestQual16Test.Count)
                    End If
                    With mnuTestQual16Test(mnuTestQual16Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 17
                   With mnuTestQual17
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual17Test(mnuTestQual17Test.Count)
                    End If
                    With mnuTestQual17Test(mnuTestQual17Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 18
                   With mnuTestQual18
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual18Test(mnuTestQual18Test.Count)
                    End If
                    With mnuTestQual18Test(mnuTestQual18Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 19
                   With mnuTestQual19
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual19Test(mnuTestQual19Test.Count)
                    End If
                    With mnuTestQual19Test(mnuTestQual19Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 20
                   With mnuTestQual20
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual20Test(mnuTestQual20Test.Count)
                    End If
                    With mnuTestQual20Test(mnuTestQual20Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 21
                    With mnutestQual21
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                        .Visible = True
                    End With
                    If iItem > 0 Then
                       Load mnuTestQual21Test(mnuTestQual21Test.Count)
                    End If
                    With mnuTestQual21Test(mnuTestQual21Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 22
                   With mnuTestQual22
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual22Test(mnuTestQual22Test.Count)
                    End If
                    With mnuTestQual22Test(mnuTestQual22Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 23
                   With mnuTestQual23
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual23Test(mnuTestQual23Test.Count)
                    End If
                    With mnuTestQual23Test(mnuTestQual23Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 24
                   With mnuTestQual24
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                      .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual24Test(mnuTestQual24Test.Count)
                    End If
                    With mnuTestQual24Test(mnuTestQual24Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 25
                   With mnuTestQual25
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual25Test(mnuTestQual25Test.Count)
                    End If
                    With mnuTestQual25Test(mnuTestQual25Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 26
                   With mnuTestQual26
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual26Test(mnuTestQual26Test.Count)
                    End If
                    With mnuTestQual26Test(mnuTestQual26Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 27
                   With mnuTestQual27
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual27Test(mnuTestQual27Test.Count)
                    End If
                    With mnuTestQual27Test(mnuTestQual27Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case 28
                   With mnuTestQual28
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual28Test(mnuTestQual28Test.Count)
                    End If
                    With mnuTestQual28Test(mnuTestQual28Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                Case Else
                   With mnuTestQual29
                        .Caption = dtaTest.Recordset.Fields("Qualification")
                       .Visible = True
                   End With
                    If iItem > 0 Then
                       Load mnuTestQual29Test(mnuTestQual29Test.Count)
                    End If
                    With mnuTestQual29Test(mnuTestQual29Test.Count - 1)
                        .Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                        .Tag = dtaTest.Recordset.Fields("Code")
                        .Visible = True
                    End With
                End Select
                  frmMain.mnuTestAddNew.Enabled = True
                  frmMain.mnuTestEdit.Enabled = True
                  frmMain.mnuTestRemove.Enabled = True
              Else
                  iItem = iItem + 1
                  If iItem > 0 Then
                     Load mnuTestQual1Test(mnuTestQual1Test.Count)
                  End If
                  With mnuTestQual1Test(mnuTestQual1Test.Count - 1)
                      .Caption = "&" & Format$(iItem + 1) & ": " & dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
                      .Tag = dtaTest.Recordset.Fields("Code")
                      .Visible = True
                  End With
                  frmMain.mnuTestAddNew.Enabled = False
                  frmMain.mnuTestEdit.Enabled = False
                  frmMain.mnuTestRemove.Enabled = False
              End If
            End If
          dtaTest.Recordset.MoveNext
       Loop
    Else
       MsgBox Translate("No valid list of tests available. Download Sport Rules first", mcLanguage), vbCritical
       Unload Me
       End
   End If

End Sub
Sub CreateNewTest()
   Dim iKey As Integer
   Dim iTemp As Integer
   Dim iWr As Integer
   Dim cTest As String
   Dim cOld As String
   Dim cTestList As String
   Dim cTemp As String
   
   Dim rstOld As DAO.Recordset
   Dim rstNew As DAO.Recordset
   Dim fld As DAO.Field
   
   iKey = MsgBox(Translate("Add a new test?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
   If iKey = vbYes Then
        Set rstOld = mdbMain.OpenRecordset("SELECT Code FROM Tests WHERE (Removed=False or ISNULL(Removed)) ORDER BY Code")
        If rstOld.RecordCount > 0 Then
            Do While Not rstOld.EOF
                If cTestList = "" Then
                    cTestList = rstOld.Fields("Code")
                Else
                    cTestList = cTestList & ", " & rstOld.Fields("Code")
                End If
                rstOld.MoveNext
            Loop
        End If
        iKey = MsgBox(Translate("Do you want to base this test upon an existing test?", mcLanguage), vbYesNo + vbQuestion)
        If iKey = vbYes Then
            Do
                cOld = InputBox(Translate("The new test is based upon which existing test?", mcLanguage) & vbCrLf & Translate("Enter one of these existing tests", mcLanguage) & ":" & vbCrLf & cTestList, , cOld)
                cOld = UCase$(cOld)
            Loop While cOld <> "" And InStr(", " & cTestList & ",", ", " & cOld & ",") = 0
        Else
            cOld = ""
        End If
        Do
            cTest = InputBox(Translate("Enter the code for this test, max. 8 positions (only letters and numbers), like ", mcLanguage) & cOld & "X", , cOld & "X")
            cTest = UCase$(Left$(UnDotSpace(cTest), 8))
        Loop While cTest <> "" And InStr(", " & cTestList & ",", ", " & cTest & ",") > 0
        iKey = vbNo
        If cOld <> "" And cTest <> "" Then
            iKey = MsgBox(Translate("Add new test", mcLanguage) & ": " & cTest & " " & Translate("based upon", mcLanguage) & " " & cOld & "?", vbYesNo + vbExclamation + vbDefaultButton2)
        Else
            iKey = MsgBox(Translate("Add new test", mcLanguage) & ": " & cTest & "?", vbYesNo + vbExclamation + vbDefaultButton2)
        End If
        If iKey = vbYes Then
            Set rstNew = mdbMain.OpenRecordset("SELECT * FROM Tests")
            If cOld <> "" Then
                Set rstOld = mdbMain.OpenRecordset("SELECT * FROM Tests WHERE Code LIKE '" & cOld & "'")
                If rstOld.RecordCount > 0 Then
                    If rstOld.Fields("WR").Value > 0 Then
                        iKey = MsgBox(Translate("Do you want this test to be a WorldRanking test as well?", mcLanguage), vbYesNo + vbQuestion)
                        If iKey = vbYes Then
                            iWr = True
                        End If
                    End If
                    With rstNew
                        cTemp = InputBox(Translate("What is the name of the new test", mcLanguage) & " (" & Translate("Like", mcLanguage) & " '" & Translate(rstOld.Fields("Test"), mcLanguage) & " " & Translate("for senior riders", mcLanguage) & "')?", , Translate(rstOld.Fields("Test"), mcLanguage))
                        .AddNew
                        For Each fld In rstOld.Fields
                            If fld.Name = "Code" Then
                                .Fields(fld.Name) = cTest
                            ElseIf fld.Name = "Test" Then
                                .Fields(fld.Name) = cTemp
                            ElseIf fld.Name = "WR" Then
                                If iWr = True Then
                                    CopyField fld, rstNew.Fields(fld.Name)
                                End If
                            Else
                                CopyField fld, rstNew.Fields(fld.Name)
                            End If
                        Next
                        If iWr = True Then
                            .Fields("WRTest") = rstOld.Fields("WRTest")
                        Else
                            .Fields("WRTest") = ""
                        End If
                        .Update
                    End With
                End If
            Else
                cTemp = InputBox(Translate("What is the name of the new test", mcLanguage) & "?")
                With rstNew
                    .AddNew
                    .Fields("Code") = cTest
                    .Fields("Test") = cTemp
                    .Fields("RR") = True
                    .Fields("Userlevel") = 1
                    .Fields("Type_Pre") = 1
                    .Fields("Type_Final") = 1
                    .Fields("Qualification") = "Extra"
                    .Fields("Type_Time") = 0
                    .Fields("Out_fin") = 0
                    .Fields("Num_J") = 5
                    .Fields("Removed") = False
                    .Fields("Type_Special") = 0
                    .Update
                End With
            End If
            Set rstNew = mdbMain.OpenRecordset("SELECT * FROM TestSections")
            If cOld <> "" Then
                Set rstOld = mdbMain.OpenRecordset("SELECT * FROM TestSections WHERE Code LIKE '" & cOld & "'")
                If rstOld.RecordCount > 0 Then
                    Do While Not rstOld.EOF
                        With rstNew
                            .AddNew
                            For Each fld In rstOld.Fields
                                If fld.Name = "Code" Then
                                    rstNew.Fields(fld.Name) = cTest
                                Else
                                    rstNew.Fields(fld.Name) = rstOld.Fields(fld.Name)
                                End If
                            Next
                            .Update
                        End With
                        rstOld.MoveNext
                    Loop
                End If
                Set rstOld = mdbMain.OpenRecordset("SELECT * FROM TestTimeTables WHERE Code LIKE '" & cOld & "'")
                Set rstNew = mdbMain.OpenRecordset("SELECT * FROM TestTimeTables")
                If rstOld.RecordCount > 0 Then
                    Do While Not rstOld.EOF
                        With rstNew
                            .AddNew
                            For Each fld In rstOld.Fields
                                If fld.Name = "Code" Then
                                    rstNew.Fields(fld.Name) = cTest
                                Else
                                    rstNew.Fields(fld.Name) = rstOld.Fields(fld.Name)
                                End If
                            Next
                            .Update
                        End With
                        rstOld.MoveNext
                    Loop
                End If
            Else
                With rstNew
                    .AddNew
                    .Fields("Code") = cTest
                    .Fields("Status") = 0
                    .Fields("Section") = 1
                    .Fields("Name") = "Preliminary Round"
                    .Fields("Mark_low") = 0
                    .Fields("Mark_hi") = 10
                    .Fields("Factor") = 1
                    .Fields("Out") = False
                    .Fields("Recycle") = False
                    .Update
                    
                    For iTemp = 1 To 2
                        .AddNew
                        .Fields("Code") = cTest
                        .Fields("Status") = 1
                        .Fields("Section") = iTemp
                        .Fields("Name") = "Section " & iTemp
                        .Fields("Mark_low") = 0
                        .Fields("Mark_hi") = 10
                        .Fields("Factor") = 1
                        .Fields("Out") = False
                        .Fields("Recycle") = False
                        .Update
                    Next iTemp
                End With
            End If
            
            CreateTestInfo cTest
            
            Set rstNew = mdbMain.OpenRecordset("SELECT Nr FROM TestInfo WHERE Nr>0 ORDER BY Nr DESC")
            If rstNew.RecordCount > 0 Then
                iTemp = rstNew.Fields("Nr")
                Set rstNew = mdbMain.OpenRecordset("SELECT Nr FROM TestinFo WHERE Code='" & cTest & "'")
                If rstNew.RecordCount > 0 Then
                    With rstNew
                        .Edit
                        .Fields("Nr") = iTemp + 1
                        .Update
                    End With
                End If
            End If
            
            rstNew.Close
            Set rstNew = Nothing
            CreateTestMenu
            frmTests.fcInitCode = cTest
            frmTests.Caption = ClipAmp(mnuTestAddNew.Caption)
            frmTests.Show 1, Me
            
        End If
        rstOld.Close
        Set rstOld = Nothing
   End If
End Sub
Sub CreateNewCombination()
   Dim iKey As Integer
   Dim iTemp As Integer
   Dim cCombination As String
   Dim cCode As String
   Dim cOld As String
   Dim cTestList As String
   Dim cTemp As String
   
   Dim rstOld As DAO.Recordset
   Dim rstNew As DAO.Recordset
   Dim fld As DAO.Field
   
   iKey = MsgBox(Translate("Add a new combination of tests?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
   If iKey = vbYes Then
        Set rstOld = mdbMain.OpenRecordset("SELECT Combination FROM Combinations ORDER BY Combination")
        If rstOld.RecordCount > 0 Then
            Do While Not rstOld.EOF
                If cTestList = "" Then
                    cTestList = Translate(rstOld.Fields("Combination") & "", mcLanguage)
                Else
                    cTestList = cTestList & ", " & Translate(rstOld.Fields("Combination") & "", mcLanguage)
                End If
                rstOld.MoveNext
            Loop
        End If
        cOld = ""
        cCombination = InputBox(Translate("Enter the name for this combination of tests, max. 25 positions, preferrably in English.", mcLanguage), , "")
        If cCombination <> "" And cCombination <> Chr$(27) Then
            cCombination = StrConv(cCombination, vbProperCase)
            If InStr(", " & cTestList & ",", ", " & cCombination & ",") > 0 Then
                frmCombination.fcCombination = cCombination
                frmCombination.Show 1, Me
            Else
                cTestList = ""
                Set rstOld = mdbMain.OpenRecordset("SELECT Distinct Code FROM Combinations")
                If rstOld.RecordCount > 0 Then
                    Do While Not rstOld.EOF
                        If cTestList = "" Then
                            cTestList = Translate(rstOld.Fields("Code") & "", mcLanguage)
                        Else
                            cTestList = cTestList & "|" & Translate(rstOld.Fields("Code") & "", mcLanguage)
                        End If
                        rstOld.MoveNext
                    Loop
                End If
                Do
                    cCode = InputBox(Translate("Enter the code for", mcLanguage) & " '" & cCombination & "', " & Translate("max. 8 positions.", mcLanguage), , "")
                Loop While InStr("|" & cTestList & "|", "|" & cCode & "|") <> 0 And Len(cCode) > 8
                If cCode <> "" And cCode <> Chr$(27) Then
                    cCode = StrConv(cCode, vbUpperCase)
                    iKey = MsgBox(Translate("Add new combination of tests", mcLanguage) & ": '" & cCode & "-" & cCombination & "'?", vbYesNo + vbExclamation + vbDefaultButton2)
                    If iKey = vbYes Then
                        Set rstOld = mdbMain.OpenRecordset("SELECT * FROM Combinations ORDER BY Combination")
                        With rstOld
                            .AddNew
                            .Fields("Combination") = Left$(cCombination, .Fields("Combination").Size)
                            .Fields("Code") = Left$(cCode, .Fields("Code").Size)
                            .Fields("Userlevel") = 1
                            .Update
                        End With
                        frmCombination.fcCombination = cCombination
                        frmCombination.Show 1, Me
                    End If
                End If
            End If
        End If
        rstOld.Close
        Set rstOld = Nothing
   End If

End Sub
Sub RemoveTest()
   Dim iKey As Integer
   Dim cTest As String
   Dim cOld As String
   Dim cTestList As String
   Dim cTemp As String
   
   Dim rstOld As DAO.Recordset
   Dim fld As DAO.Field
   
   iKey = MsgBox(Translate("Remove a test?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
   If iKey = vbYes Then
        Set rstOld = mdbMain.OpenRecordset("SELECT Code FROM Tests WHERE Userlevel=1 ORDER BY Code")
        If rstOld.RecordCount > 0 Then
            Do While Not rstOld.EOF
                If cTestList = "" Then
                    cTestList = rstOld.Fields("Code")
                Else
                    cTestList = cTestList & ", " & rstOld.Fields("Code")
                End If
                rstOld.MoveNext
            Loop
        End If
        If cTestList <> "" Then
            Do
                cOld = InputBox(Translate("Which of these tests should be removed?", mcLanguage) & vbCrLf & Translate("Enter one of these existing tests", mcLanguage) & ":" & vbCrLf & cTestList, , cOld)
                cOld = UCase$(cOld)
            Loop While cOld <> "" And InStr(", " & cTestList & ",", ", " & cOld & ",") = 0
            If cOld <> "" Then
                iKey = MsgBox(Translate("Remove test", mcLanguage) & ": '" & cOld & "'?" & vbCrLf & "(" & Translate("This is only possible when no marks have been entered", mcLanguage) & "!)", vbYesNo + vbExclamation + vbDefaultButton2)
                If iKey = vbYes Then
                    Set rstOld = mdbMain.OpenRecordset("SELECT Code FROM MARKS WHERE Code LIKE '" & cOld & "'")
                    If rstOld.RecordCount = 0 Then
                        mdbMain.Execute ("DELETE * FROM TestTimeTables WHERE Code Like '" & cOld & "'")
                        mdbMain.Execute ("DELETE * FROM Testsections WHERE Code Like '" & cOld & "'")
                        mdbMain.Execute ("DELETE * FROM Tests WHERE Code Like '" & cOld & "'")
                        mdbMain.Execute ("DELETE * FROM Results WHERE Code Like '" & cOld & "'")
                    Else
                        MsgBox Translate("Test cannot be removed; marks have already been entered", mcLanguage) & "!", vbExclamation
                    End If
                    CreateTestMenu
                End If
            End If
        Else
            MsgBox Translate("No tests to remove.", mcLanguage), vbInformation
        End If
        rstOld.Close
        Set rstOld = Nothing
   End If

End Sub

Private Sub RemoveCombination()
   Dim iKey As Integer
   Dim cTest As String
   Dim cOld As String
   Dim cTemp As String
   Dim cTestList As String
   
   Dim rstOld As DAO.Recordset
   Dim fld As DAO.Field
   
   iKey = MsgBox(Translate("Remove a combination of tests?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
   If iKey = vbYes Then
        Set rstOld = mdbMain.OpenRecordset("SELECT Distinct Code FROM Combinations WHERE Userlevel=1 ORDER BY Code")
        If rstOld.RecordCount > 0 Then
            Do While Not rstOld.EOF
                If cTestList = "" Then
                    cTestList = rstOld.Fields("Code") & ""
                Else
                    cTestList = cTestList & ", " & rstOld.Fields("Code") & ""
                End If
                rstOld.MoveNext
            Loop
        End If
        Do
            cOld = InputBox(Translate("Which of these combinations should be removed?", mcLanguage) & vbCrLf & Translate("Enter one of these existing combinations", mcLanguage) & ":" & vbCrLf & cTestList, , cOld)
            cOld = UCase$(cOld)
        Loop While cOld <> "" And InStr(", " & cTestList & ",", ", " & cOld & ",") = 0
        If cOld <> "" Then
            iKey = MsgBox(Translate("Remove combination", mcLanguage) & ": '" & cOld & "'?", vbYesNo + vbExclamation + vbDefaultButton2)
            If iKey = vbYes Then
                mdbMain.Execute ("DELETE * FROM CombinationSections WHERE Code LIKE '" & cOld & "'")
                mdbMain.Execute ("DELETE * FROM Combinations WHERE Code LIKE '" & cOld & "'")
                If mdbMain.RecordsAffected > 0 Then
                    CreateCombinationMenu
                End If
            End If
        End If
        rstOld.Close
        Set rstOld = Nothing
   End If
End Sub
Private Sub txtScore_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub txtTime_Change()
   lblParticipant.Tag = "*"
End Sub

Private Sub txtTime_DblClick()
    If txtTime.BackColor <> mlAlertColor Then
       txtTime.BackColor = mlAlertColor
       miNoBackupNow = True
    End If

End Sub

Private Sub txtTime_GotFocus()
   txtTime.SelLength = 5
   If txtTime.BackColor <> mlAlertColor Then
        txtTime.BackColor = mlAlertColor
        miNoBackupNow = True
   End If
End Sub

Private Sub txtTime_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case vbKeyReturn
      If ValidateTime = True Then
        SetFocusTo cmdOK
      End If
      KeyCode = 0
   End Select
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
    'suppress annoying beep when using <Enter> in stead of <Tab>
    If KeyAscii = 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtTime_LostFocus()
    Set mctlActive = txtTime
    txtTime.BackColor = QBColor(15)
    miNoBackupNow = False
End Sub

Private Sub txtTime_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
   If Button = vbRightButton Then
      StartMenuPopUp
   End If
End Sub

Private Sub txtTime_Validate(Cancel As Boolean)
    ValidateTime
End Sub
Public Function CreateCaption() As String
    Dim cTemp As String
    
    cTemp = Me.TestCode & " - " & Translate(Me.TestName, mcLanguage)
    If dtaTest.Recordset.Fields("Type_pre") <= 2 Then
        cTemp = cTemp & ": " & ClipAmp(tbsSelFin.SelectedItem.Caption)
    End If
    If ClipAmp(tbsSelFin.SelectedItem.Caption) <> ClipAmp(Me.tbsSection(tbsSelFin.SelectedItem.Index - 1).SelectedItem.Caption) Then
        cTemp = cTemp & " (" & ClipAmp(Me.tbsSection(tbsSelFin.SelectedItem.Index - 1).SelectedItem.Caption) & ")"
    End If
    CreateCaption = cTemp & " [" & EventName & "]"
End Function
Sub WriteLogDBFinals()
    'Write starting lists of finals to the LogDB
    Dim iWrite As Integer
    Dim rstColor As DAO.Recordset
    
    Set rstColor = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Deleted=0 ORDER BY Position DESC")
    While Not rstColor.EOF
        iWrite = WriteLogDBStart(EventName, rstColor("Code"), rstColor("Status"), rstColor("STA"), rstColor("Position"), rstColor("Group"), "")
        rstColor.MoveNext
    Wend
    
    rstColor.Close
    Set rstColor = Nothing
End Sub
Sub AddColorsToFinals()
    Dim iColor As Integer
    Dim cColor() As String
    
    Dim rstColor As DAO.Recordset
    cColor = Split(TestColors, ",")
    Set rstColor = mdbMain.OpenRecordset("SELECT Color,Group,Position FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " ORDER BY Position DESC")
    If rstColor.RecordCount > 0 Then
        Do While Not rstColor.EOF
            rstColor.Edit
            If iColor <= UBound(cColor) Then
                rstColor.Fields("Color") = cColor(iColor)
            Else
                rstColor.Fields("Color") = mcNoColor
            End If
            rstColor.Fields("Group") = 1
            rstColor.Update
            rstColor.MoveNext
            iColor = iColor + 1
        Loop
    End If
    rstColor.Close
    Set rstColor = Nothing
End Sub
Public Function PrintCombination(CombinationCode As String, TableName As String) As Integer
    Dim rstComb As DAO.Recordset
    Dim rstTemp As DAO.Recordset
    Dim rstResults As DAO.Recordset
    Dim rstMarks As DAO.Recordset
    
    Dim iOldGroup As Integer
    Dim cPosition As String
    Dim iPosition As Integer
    Dim iGroup As Integer
    Dim iGroupcount As Integer
    Dim iTemp As Integer
    Dim curOldScore As Currency
    
    If TableExist(mdbMain, TableName) = True Then
        Set rstComb = mdbMain.OpenRecordset("SELECT Participants.STA,Participants.Team,Participants.Club,Participants.Class, [" & TableName & "].*, Horses.Name_Horse,Persons.Name_First & ' ' & Persons.Name_Last AS Name_rider FROM (([" & TableName & "] INNER JOIN Participants ON [" & TableName & "].STA = Participants.STA) INNER JOIN Horses ON Participants.HorseID = Horses.HorseID) INNER JOIN Persons ON Participants.PersonID = Persons.PersonID WHERE Score>0 ORDER BY Score DESC")
        If rstComb.RecordCount = 0 Then
            SetMouseNormal
            MsgBox Translate("No participants with sufficient marks in this combination", mcLanguage)
        Else
            Set rstTemp = mdbMain.OpenRecordset("SELECT * FROM Combinations WHERE Code Like '" & CombinationCode & "'")
            With rtfResult
                .Text = ""
                .SelBold = True
                .SelFontSize = 18
                .SelText = EventName & vbCrLf
                .SelBold = True
                .SelFontSize = 18
                .SelText = Translate("Combination", mcLanguage) & ": " & Translate(rstTemp.Fields("Combination"), mcLanguage) & vbCrLf
                .SelBold = False
                .SelFontSize = 11
            End With
            
            Set rstTemp = mdbMain.OpenRecordset("SELECT * FROM CombinationSections WHERE Code Like '" & CombinationCode & "' ORDER BY [Group],Test")
            If rstTemp.RecordCount > 0 Then
                With rtfResult
                    .SelFontSize = 11
                    .SelBold = True
                    .SelTabCount = 11
                    .SelTabs(0) = 2.5 * 567
                    For iTemp = 1 To 10
                        .SelTabs(iTemp) = (2.5 + (iTemp * 1.25)) * 567
                    Next iTemp
                    .SelText = vbCrLf & vbCrLf
                End With
                Do While Not rstTemp.EOF
                    With rtfResult
                        If iOldGroup <> rstTemp.Fields("Group") Then
                            .SelText = vbCrLf
                            .SelBold = True
                            .SelText = Translate("Group", mcLanguage) & " " & rstTemp.Fields("Group")
                            .SelBold = False
                        End If
                        iOldGroup = rstTemp.Fields("Group")
                        .SelText = vbTab & rstTemp.Fields("Test")
                        If rstTemp.Fields("Factor") <> 1 Then
                            .SelText = " (x " & Format$(rstTemp.Fields("Factor"), TestMarkFormat) & ")"
                        End If
                        .SelText = "; "
                    End With
                    rstTemp.MoveNext
                    iGroupcount = iOldGroup
                Loop
            End If
            rstTemp.Close
            
            With rtfResult
                .SelText = vbCrLf
                .SelTabCount = 1
                .SelTabs(0) = 16 * 567
                .SelUnderline = True
                .SelText = vbTab & vbCrLf
                .SelUnderline = False
                GoSub SetTabsRider
                .SelItalic = True
                .SelBold = False
                .SelFontSize = 8
                .SelText = "POS" & vbTab & "#" & vbTab & Translate("RIDER / HORSE", mcLanguage) & vbTab & "TOT" & vbCrLf
                GoSub SetTabsMarks
                .SelFontSize = 8
                .SelItalic = True
                .SelUnderline = True
                For iGroup = 1 To .SelTabCount
                    If iGroup <= iGroupcount Then
                        .SelText = vbTab & Format$(iGroup)
                    Else
                        .SelText = vbTab
                    End If
                Next iGroup
                .SelText = vbCrLf & vbCrLf
                .SelUnderline = False
                .SelItalic = False
            End With
            
            iPosition = 0
            Do While Not rstComb.EOF
                iPosition = iPosition + 1
                If curOldScore <> rstComb.Fields("Score") Then
                    cPosition = iPosition
                End If
                curOldScore = rstComb.Fields("Score")
                With rtfResult
                    GoSub SetTabsRider
                    .SelText = cPosition & vbTab & rstComb.Fields("Participants.STA") & vbTab & rstComb.Fields("Name_rider")
                    
                    If rstComb.Fields("Class") & "" <> "" Then
                        .SelText = " [" & rstComb.Fields("Class") & "]"
                    End If
                    
                    .SelText = vbCrLf
                    GoSub SetTabsRider
                    .SelText = vbTab & vbTab & rstComb.Fields("Name_Horse")
                    If miShowRidersClub <> 0 Then
                        .SelText = " / " & GetRidersClub(rstComb)
                    End If
                    If miShowRidersTeam <> 0 Then
                        .SelText = " / " & GetRidersTeam(rstComb)
                    End If
                    'IPZV LK output:
                    If mcVersionSwitch = "ipzv" And miShowRidersLK = 1 Then
                        .SelText = " / " & Left$(GetRidersLk(rstComb), 2)
                    End If
                    .SelText = vbTab & Format$(rstComb.Fields("Score"), "0.000") & vbCrLf
                    GoSub SetTabsMarks
                    
                    For iGroup = 1 To iGroupcount
                        .SelText = vbTab & rstComb.Fields("Code" & Format$(iGroup)) & ": " & Format$(rstComb.Fields("Score" & Format$(iGroup)), TestTotalFormat)
                    Next iGroup
                    .SelText = vbCrLf
                End With
                rstComb.MoveNext
            Loop
            
            PrintRtfFooter Translate("Combination", mcLanguage), "ZZ-" & CombinationCode
        End If
        rstComb.Close
        Set rstComb = Nothing
        Set rstTemp = Nothing
    Else
        SetMouseNormal
        MsgBox Translate("No participants with sufficient marks in this combination", mcLanguage)
    End If
    SetMouseNormal

Exit Function

SetTabsRider:
    With rtfResult
        .SelFontSize = 11
        .SelBold = True
        .SelTabCount = 3
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 2.5 * 567
        .SelTabs(2) = 14 * 567
    End With
Return

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelBold = False
        .SelTabCount = 6
        .SelTabs(0) = 3 * 567
        .SelTabs(1) = 5.5 * 567
        .SelTabs(2) = 8 * 567
        .SelTabs(3) = 10.5 * 567
        .SelTabs(4) = 13 * 567
        .SelTabs(5) = 16 * 567
    End With
Return

End Function
Public Function PrintTeamCombination(CombinationCode As String, TableName As String, iNumTests As Integer, cTeamClub As String) As Integer
    Dim rstComb As DAO.Recordset
    Dim rstTemp As DAO.Recordset
    Dim rstResults As DAO.Recordset
    Dim rstMarks As DAO.Recordset
    
    Dim iOldGroup As Integer
    Dim cPosition As String
    Dim iPosition As Integer
    Dim iGroup As Integer
    Dim iGroupcount As Integer
    Dim iTemp As Integer
    Dim cTemp As String
    Dim curOldScore As Currency
    Dim fld As DAO.Field
    
    If TableExist(mdbMain, TableName) = True Then
        Set rstComb = mdbMain.OpenRecordset("SELECT * FROM [" & TableName & "] WHERE Score>0 ORDER BY Score DESC")
        If rstComb.RecordCount = 0 Then
            SetMouseNormal
            MsgBox Translate("No teams with sufficient marks in this combination", mcLanguage)
        Else
            Set rstTemp = mdbMain.OpenRecordset("SELECT * FROM Combinations WHERE Code Like '" & CombinationCode & "'")
            With rtfResult
                .Text = ""
                .SelBold = True
                .SelFontSize = 18
                .SelText = EventName & vbCrLf
                .SelBold = True
                .SelFontSize = 18
                .SelText = Translate("Combination", mcLanguage) & ": " & Translate(rstTemp.Fields("Combination"), mcLanguage) & vbCrLf
                .SelBold = False
                .SelFontSize = 11
            End With
            rstTemp.Close
            
            With rtfResult
                .SelText = vbCrLf
                .SelTabCount = 1
                .SelTabs(0) = 16 * 567
                .SelUnderline = True
                .SelText = vbTab & vbCrLf
                .SelUnderline = False
                GoSub SetTabsRider
                .SelItalic = True
                .SelBold = False
                .SelFontSize = 8
                .SelText = "POS" & vbTab & " " & vbTab & Translate(cTeamClub, mcLanguage) & vbTab & "TOT" & vbCrLf
                GoSub SetTabsMarks
                .SelFontSize = 8
                .SelItalic = True
                .SelUnderline = True
                For iGroup = 1 To .SelTabCount
                    If iGroup <= iGroupcount Then
                        .SelText = vbTab & Format$(iGroup)
                    Else
                        .SelText = vbTab
                    End If
                Next iGroup
                .SelText = vbCrLf & vbCrLf
                .SelUnderline = False
                .SelItalic = False
            End With
            
            iPosition = 0
            Do While Not rstComb.EOF
                iPosition = iPosition + 1
                If curOldScore <> rstComb.Fields("Score") Then
                    cPosition = iPosition
                End If
                curOldScore = rstComb.Fields("Score")
                With rtfResult
                    GoSub SetTabsRider
                    .SelText = cPosition & vbTab & " " & vbTab & rstComb.Fields("Team")
                    .SelText = vbTab & Format$(rstComb.Fields("Score"), "0.000") & vbCrLf
                End With
                rstComb.MoveNext
            Loop
            
            rtfResult.SelText = "$#@!"
    
            Set rstComb = mdbMain.OpenRecordset("SELECT * FROM [" & TableName & "] ORDER BY Team,Score DESC")
            If rstComb.RecordCount > 0 Then
                
                With rtfResult
                    .SelText = vbCrLf
                    GoSub SetTabsMarks
                    .SelFontSize = 8
                    .SelItalic = True
                    .SelBold = True
                    .SelUnderline = True
                    For Each fld In rstComb.Fields
                        .SelText = fld.Name & vbTab
                    Next
                    .SelText = vbCrLf & vbCrLf
                    .SelUnderline = False
                    .SelItalic = False
                    .SelBold = False
                End With
                
                Do While Not rstComb.EOF
                    For Each fld In rstComb.Fields
                        If fld.OrdinalPosition = 0 Then
                            If cTemp <> fld.Value Then
                                rtfResult.SelText = vbCrLf
                                rtfResult.SelText = Left$(fld.Value, 15) & vbTab
                            Else
                                rtfResult.SelText = " " & vbTab
                            End If
                            cTemp = fld.Value & ""
                        ElseIf fld.OrdinalPosition = 1 Then
                            rtfResult.SelText = Format$(fld.Value, "0.000") & vbTab
                        ElseIf fld.Value > 0 Then
                            rtfResult.SelText = Format$(fld.Value, "0.00") & vbTab
                        Else
                            rtfResult.SelText = " " & vbTab
                        End If
                    Next
                    rtfResult.SelText = vbCrLf
                    
                    rstComb.MoveNext
                
                Loop
            End If
            PrintRtfFooter Translate("Combination", mcLanguage), "ZZ-" & CombinationCode
        End If
        rstComb.Close
        
        Set rstComb = Nothing
        Set rstTemp = Nothing
    Else
        SetMouseNormal
        MsgBox Translate("No participants with sufficient marks in this combination", mcLanguage)
    End If
    SetMouseNormal

Exit Function

SetTabsRider:
    With rtfResult
        .SelFontSize = 11
        .SelBold = True
        .SelTabCount = 3
        .SelTabs(0) = 1 * 567
        .SelTabs(1) = 2.5 * 567
        .SelTabs(2) = 14 * 567
    End With
Return

SetTabsMarks:
    With rtfResult
        .SelFontSize = 9
        .SelBold = False
        .SelTabCount = iNumTests + 2
        .SelTabs(0) = 0
        For iTemp = 1 To .SelTabCount - 1
            .SelTabs(iTemp) = (iTemp + 1) * 567 * 1.5 + 567
        Next iTemp
    End With
Return

End Function
Public Function getHorseAge(r As DAO.Recordset) As String
    On Local Error Resume Next
    Dim iTemp As Integer
    Dim iHorseage As Integer
    getHorseAge = ""
    iHorseage = 999
    If r.Fields("Horses.FEIFID") & "" <> "" Then
        iTemp = Val(Mid$(r.Fields("Horses.FEIFID") & "00000000", 3, 4))
        If iTemp > 1900 And iTemp <= Year(Now) Then
            iHorseage = DateDiff("y", iTemp, Year(Now()))
        End If
    End If
    If iHorseage = 999 Then
        If r.Fields("FEIFID") & "" <> "" Then
            iTemp = Val(Mid$(r.Fields("FEIFID") & "00000000", 3, 4))
            If iTemp > 1900 And iTemp <= Year(Now) Then
                iHorseage = DateDiff("y", iTemp, Year(Now()))
            End If
        End If
    End If
    If iHorseage < miHorseAgeLimit Then
        getHorseAge = "**" & Format$(iHorseage) & "**"
    End If
    On Local Error GoTo 0

End Function
Public Function GetHorseId(r As DAO.Recordset) As String
    On Local Error Resume Next
    GetHorseId = ""
    If r.Fields("Horses.FEIFID") & "" <> "" Then
        GetHorseId = r.Fields("Horses.FEIFID") & ""
    End If
    If GetHorseId = "" Then
        If r.Fields("FEIFID") & "" <> "" Then
            GetHorseId = r.Fields("FEIFID") & ""
        Else
            GetHorseId = "-"
        End If
    End If
    
    'ÖIV: print HorseID instead of FEIFID
    If mcVersionSwitch = "öiv" Then
        If r.Fields("HorseID") & "" <> "" Then
            GetHorseId = r.Fields("HorseID") & ""
        End If
    End If
    
   
    
    On Local Error GoTo 0
End Function

Public Function GetRidersClub(r As DAO.Recordset) As String
    On Local Error Resume Next
    If r.Fields("Club") & "" <> "" Then
        GetRidersClub = r.Fields("Club") & ""
    End If
    On Local Error GoTo 0
End Function
Public Function GetRidersTeam(r As DAO.Recordset) As String
    On Local Error Resume Next
    If r.Fields("Team") & "" <> "" Then
        GetRidersTeam = r.Fields("Team") & ""
    End If
        
    On Local Error GoTo 0
End Function
Public Function GetRidersLk(r As DAO.Recordset) As String
    On Local Error Resume Next

    GetRidersLk = GetLK(dtaTest.Recordset.Fields("Code"), Left(r.Fields(0), 3))
    
    On Local Error GoTo 0
End Function
Public Function GetLK(ipo As String, startnr As String) As String
    On Local Error Resume Next
    Dim rstLK As DAO.Recordset
    Dim temp As Currency
    
    Set rstLK = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE code='" & ipo & "' AND sta='" & startnr & "'")
    
    GetLK = ""
    If rstLK.RecordCount <> 0 Then
        temp = rstLK("qualification")
        If temp > 0 And CInt(temp) = temp Then
            GetLK = "[LK " & CStr(CInt(temp)) & "]"
        End If
    End If
    rstLK.Close
    Set rstLK = Nothing

    On Local Error GoTo 0
End Function
Public Function GetJudgeId(cTestCode As String, iTestStatus As Integer, iJudgePosition As Integer) As String
    Dim rstJudge As DAO.Recordset
    Dim cTemp As String
    
    On Local Error GoTo 0
    
    If cTestCode <> "" Then
        Set rstJudge = mdbMain.OpenRecordset("Select JudgeId FROM TestJudges WHERE Code LIKE '" & cTestCode & "' AND Status=" & iTestStatus & " AND [Position]=" & iJudgePosition)
        If rstJudge.RecordCount = 0 And iTestStatus > 0 Then
                cTemp = ""
                Set rstJudge = mdbMain.OpenRecordset("Select * FROM TestJudges WHERE Code LIKE '" & cTestCode & "' AND Status=0 AND Position=" & iJudgePosition)
                If rstJudge.RecordCount > 0 Then
                    cTemp = rstJudge.Fields("JudgeId") & ""
                    With rstJudge
                        .AddNew
                        .Fields("Code") = cTestCode
                        .Fields("Status") = iTestStatus
                        .Fields("Position") = iJudgePosition
                        .Fields("JudgeId") = cTemp
                        .Update
                    End With
                End If
                GetJudgeId = cTemp
        ElseIf iTestStatus > 0 Then
            cTemp = rstJudge.Fields("JudgeId") & ""
            If cTemp = "" Then
                Set rstJudge = mdbMain.OpenRecordset("Select * FROM TestJudges WHERE Code LIKE '" & cTestCode & "' AND Status=0 AND [Position]=" & iJudgePosition)
                If rstJudge.RecordCount > 0 Then
                    cTemp = rstJudge.Fields("JudgeId") & ""
                    If cTemp <> "" Then
                        With rstJudge
                            .AddNew
                            .Fields("Code") = cTestCode
                            .Fields("Status") = iTestStatus
                            .Fields("Position") = iJudgePosition
                            .Fields("JudgeId") = cTemp
                            .Update
                        End With
                    End If
                End If
            End If
            GetJudgeId = cTemp
        Else
            GetJudgeId = rstJudge.Fields("JudgeId") & ""
        End If
        rstJudge.Close
    End If
    Set rstJudge = Nothing
End Function
Public Function GetHorsesName(cHorseId As String) As String
    Dim rstPerson As DAO.Recordset
    If cHorseId <> "" Then
        Set rstPerson = mdbMain.OpenRecordset("Select Name_Horse FROM Horses WHERE HorseId LIKE '" & cHorseId & "'")
        If rstPerson.RecordCount > 0 Then
            GetHorsesName = rstPerson.Fields(0)
        End If
        rstPerson.Close
    End If
    Set rstPerson = Nothing
End Function
Public Function GetPersonsName(cPersonId As String) As String
    Dim rstPerson As DAO.Recordset
    Dim cTemp As String
    If cPersonId <> "" Then
        Set rstPerson = mdbMain.OpenRecordset("SELECT Persons.Name_First, Persons.Name_Middle, Persons.Name_Last, Participants.Class FROM Persons LEFT JOIN Participants ON Persons.PersonID = Participants.PersonID WHERE Persons.PersonID LIKE '" & cPersonId & "';")
        'Set rstPerson = mdbMain.OpenRecordset("Select Persons.Name_First,Persons.Name_Middle,Persons.Name_Last,Participants.Class FROM Persons INNER JOIN Participants ON Persons.PersonId=Participants.PersonId WHERE Persons.PersonId LIKE '" & cPersonId & "'")
        If rstPerson.RecordCount > 0 Then
            cTemp = rstPerson.Fields("Name_first") & RTrim$(" " & rstPerson.Fields("Name_middle")) & " " & rstPerson.Fields("Name_last")
            If rstPerson.Fields("Class") & "" <> "" Then
                cTemp = cTemp & " [" & rstPerson.Fields("Class") & "]"
            End If
        End If
        GetPersonsName = cTemp
        rstPerson.Close
    End If
    Set rstPerson = Nothing
End Function

Public Sub MakeRtfFooter(Optional iPagenum As Integer = 0)
    PrintRtfLine
    With rtfResult
        .SelTabCount = 1
        .SelTabs(0) = 2.5 * 567
        .SelText = vbCrLf
        .SelUnderline = False
        .SelFontSize = 9
        .SelItalic = True
        .SelText = Translate("Composed", mcLanguage) & " " & Format$(Now, "d mmmm yyyy hh:mm:ss") & vbCrLf
        .SelFontSize = 9
        .SelItalic = True
        .SelText = App.EXEName & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000") & " - " & App.CompanyName & vbCrLf
        .SelFontSize = 9
        .SelItalic = True
        .SelText = vbTab & GetVariable("Vision") & vbCrLf
        .SelItalic = False
        .SelUnderline = False
        .SelBold = False
        If iPagenum > 0 Then
            .SelAlignment = rtfRight
            .SelText = Translate("Page", mcLanguage) & " " & iPagenum & vbCrLf
        End If
        .SelFontSize = 12
        .SelText = vbCrLf
    End With
End Sub
Public Sub PrintRtfLine()
    Dim iBold As Integer
    With rtfResult
        iBold = .SelBold
        .SelBold = False
        .SelTabCount = 1
        .SelTabs(0) = 16 * 567
        .SelUnderline = True
        .SelText = vbTab & vbCrLf
        .SelBold = iBold
    End With
End Sub
Private Sub CheckLanguageDb()
    Dim cPath As String
    Dim cLanguages As String
    Dim cLang As String
    Dim mdbLang As DAO.Database
    Dim tdfLang As DAO.TableDef
    
    ReadIniFile gcIniHorseFile, "Database", "Language", cPath
    If cPath = "" Then
        cPath = App.Path
    End If
    If Right$(cPath, 1) <> "\" Then
        cPath = cPath & "\"
    End If
    
    cLanguages = GetVariable("Languages")
    
    If Dir$(cPath & "Languages.Mdb") <> "" Then
        Set mdbLang = DBEngine.OpenDatabase(cPath & "Languages.Mdb", False, False)
        Set tdfLang = mdbLang.TableDefs("Translation")
        AppendDeleteField tdfLang, "APPEND", "StringId", dbLong
        Do While cLanguages <> ""
            Parse cLang, cLanguages, " "
            If cLang <> "" Then
                AppendDeleteField tdfLang, "APPEND", cLang, dbText, 255
            End If
        Loop
        mdbLang.Close
        Set mdbLang = Nothing
    End If
End Sub
Public Sub CorrectColorsAndGroups()
    Dim rstEntry As DAO.Recordset
    Dim iOldGroup As Integer
    Dim iNewGroup As Integer
    Dim iColor As String
    Dim cColor() As String
    
    Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND NOT Sta IN (SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & ") ORDER BY Group,Position")
    If rstEntry.RecordCount > 0 Then
        iOldGroup = 0
        iNewGroup = 0
        iColor = 0
        cColor = Split(TestColors, ",")
        With rstEntry
            'correct colors and groups
            Do While Not .EOF
                If iOldGroup <> .Fields("Group") Then
                    'restart for every group
                    iColor = 0
                    iOldGroup = .Fields("Group")
                    iNewGroup = iNewGroup + 1
                End If
                .Edit
                If Val(cmbGroupSize.Text) <= 1 Then
                    'no groups, no colors
                    .Fields("Color") = ""
                ElseIf iColor <= UBound(cColor) Then
                    'take color from list
                    .Fields("Color") = Left$(cColor(iColor), .Fields("Color").Size)
                Else
                    'list too short
                    .Fields("Color") = mcNoColor
                End If
                If iNewGroup <> .Fields("Group") Then
                    .Fields("Group") = iNewGroup
                End If
                .Update
                .MoveNext
                iColor = iColor + 1
            Loop
            
        End With
        
    End If
    rstEntry.Close
    Set rstEntry = Nothing
    
End Sub
Public Sub CreateLogDBParticipant()
   Dim iExcelFileNum As Integer
    Dim iKey As Integer
    Dim iTemp As Integer
    Dim iErrCounter As Integer
    Dim cSex As String
    Dim cSta As String
    Dim cStaList As String
    Dim cText As String
    
    Dim rstPart As DAO.Recordset
    Dim rstMarks As DAO.Recordset
    Dim rstTest As DAO.Recordset
    
    On Local Error Resume Next
    
    If mcExcelDir = "" Then Exit Sub
    
    If Dir$(mcExcelDir, vbDirectory) = "" Then
        If Err > 0 Then
            MsgBox Translate("Cannot access", mcLanguage) & " " & mcExcelDir, vbCritical
            mcExcelDir = ""
            Exit Sub
        Else
            MkDir mcExcelDir
        End If
    End If
    
    'mch: cmbGroupSize is not set correctly when using IceSort. Check if the statement below causes trouble...
    If (chkColor.Value = 0 And chkColor.Enabled = True) Or chkColor.Enabled = False Or Val(cmbGroupSize.Text) <= 1 And TestStatus = 0 Then
        cStaList = Format$(Val(txtParticipant.Text), "000")
    Else
        Set rstPart = mdbMain.OpenRecordset("SELECT Group FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND STA='" & Format$(Val(txtParticipant.Text), "000") & "'")
        If rstPart.RecordCount > 0 Then
            Set rstPart = mdbMain.OpenRecordset("SELECT Sta FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND [Group]=" & rstPart.Fields("Group") & " ORDER BY Position")
            If rstPart.RecordCount > 0 Then
                Do While Not rstPart.EOF
                    cStaList = cStaList & rstPart.Fields(0) & " "
                    rstPart.MoveNext
                Loop
            End If
        Else
            cStaList = Format$(Val(txtParticipant.Text), "000")
        End If
    End If
    
    Do While cStaList <> ""
        Parse cSta, cStaList, " "
        cText = cText & cSta
        Set rstPart = mdbMain.OpenRecordset("SELECT Persons.*, Horses.*, Participants.Class, Participants.Club, Participants.Team FROM Horses INNER JOIN (Participants INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) ON Horses.HorseID = Participants.HorseID WHERE Participants.STA='" & cSta & "'")
        If rstPart.RecordCount > 0 Then
            With rstPart
                cText = cText & mcExcelSeparator & .Fields("Name_First") & " " & rstPart.Fields("Name_Last")
                cText = cText & mcExcelSeparator & .Fields("Class") & ""
                cText = cText & mcExcelSeparator & .Fields("Club") & ""
                cText = cText & mcExcelSeparator & .Fields("Team") & ""
                cText = cText & mcExcelSeparator & .Fields("Name_horse")
                cText = cText & mcExcelSeparator & .Fields("FEIFId")
                Select Case .Fields("Sex_Horse")
                Case 1
                   cSex = Translate("Stallion", mcLanguage)
                Case 2
                   cSex = Translate("Mare", mcLanguage)
                Case 3
                   cSex = Translate("Gelding", mcLanguage)
                Case Else
                   cSex = "--"
                End Select
                cText = cText & mcExcelSeparator & cSex
                
                Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Section=" & TestSection & " AND STA='" & cSta & "'")
                If rstMarks.RecordCount > 0 Then
                    Set rstTest = mdbMain.OpenRecordset("SELECT Name FROM Testsections  WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Section=" & TestSection)
                    If rstTest.RecordCount > 0 Then
                        cText = cText & mcExcelSeparator & rstTest.Fields("Name")
                    Else
                        cText = cText & mcExcelSeparator
                    End If
                    rstTest.Close
                    For iTemp = 1 To 5
                        If fraTime.Visible = True Then
                            cText = cText & mcExcelSeparator & ""
                        Else
                            cText = cText & mcExcelSeparator & Replace(Format$(rstMarks.Fields("Mark" & Format$(iTemp)), TestMarkFormat), ",", ".")
                        End If
                    Next iTemp
                    cText = cText & mcExcelSeparator & Replace(Format$(rstMarks.Fields("Score"), TestTotalFormat), ",", ".")
                    
                    Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND STA='" & cSta & "'")
                    If rstMarks.RecordCount > 0 Then
                        cText = cText & mcExcelSeparator
                        If rstMarks.Fields("Disq") = -1 Then
                            cText = cText & mcExcelSeparator & UCase$(Translate("ELIMINATED", mcLanguage))
                        ElseIf rstMarks.Fields("Disq") = -2 Then
                            cText = cText & mcExcelSeparator & Translate("Withdrawn", mcLanguage)
                        Else
                            cText = cText & Replace(Format$(rstMarks.Fields("Score"), TestTotalFormat), ",", ".")
                        End If
                    End If
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                Else
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                End If
                cText = cText & vbCrLf
                rstMarks.Close
            End With
        End If
        rstPart.Close
    Loop
    
    Set rstPart = Nothing
    Set rstMarks = Nothing
    Set rstTest = Nothing
    
    On Local Error GoTo CreateLogDBError
    
    iExcelFileNum = FreeFile
    Open StrConv(mcExcelDir & "participant.csv", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
    Print #iExcelFileNum, cText
    Close #iExcelFileNum
    
Exit Sub

CreateLogDBError:
    If iErrCounter < 5 Then
        iErrCounter = iErrCounter + 1
        Sleep 1
        Resume
    Else
        Exit Sub
    End If
Return
End Sub
Public Sub CreateExcelParticipant()
    Dim iExcelFileNum As Integer
    Dim iHtmlFileNum As Integer
    Dim iKey As Integer
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iErrCounter As Integer
    Dim iCreateHtml As Integer
    
    Dim cSex As String
    Dim cSta As String
    Dim cStaList As String
    Dim cText As String
    Dim cHtmlNew As String
    Dim cHtmlTop As String
    Dim cHtmlMiddle As String
    Dim cHtmlBottom As String
    Dim cHtml As String
    Dim cQry As String
    Dim cTemp As String
    
    Dim rstPart As DAO.Recordset
    Dim rstMarks As DAO.Recordset
    Dim rstTest As DAO.Recordset
    
    On Local Error Resume Next
    
    If mcExcelDir = "" Then Exit Sub
    
    If Dir$(mcExcelDir, vbDirectory) = "" Then
        If Err > 0 Then
            MsgBox Translate("Cannot access", mcLanguage) & " " & mcExcelDir, vbCritical
            mcExcelDir = ""
            Exit Sub
        Else
            MkDir mcExcelDir
        End If
    End If
    
    If Dir$(mcExcelDir & "template-participant.html") <> "" Then
        cHtml = GetHTMLTemplate(mcExcelDir & "template-participant.html")
        cHtml = Replace(cHtml, "{Title}", frmMain.EventName)
        cHtml = Replace(cHtml, "{Test}", frmMain.TestCode & " - " & frmMain.TestName)
        
        iTemp = InStr(cHtml, "<tr")
        If iTemp > 1 Then
            cHtmlTop = Left$(cHtml, iTemp - 1)
            cHtml = Mid$(cHtml, iTemp)
            iTemp = InStrRev(cHtml, "</table>")
            If iTemp > 1 Then
                cHtmlBottom = Mid$(cHtml, iTemp)
                cHtmlMiddle = Left(cHtml, iTemp - 1)
            End If
        End If
        cHtmlNew = cHtmlTop
        iCreateHtml = 1
    Else
        cHtml = ""
        cHtmlTop = ""
        cHtmlMiddle = ""
        cHtmlBottom = ""
        cHtmlNew = ""
        iCreateHtml = 0
    End If

    'mch: cmbGroupSize is not set correctly when using IceSort. Check if the statement below causes trouble...
    If (chkColor.Value = 0 And chkColor.Enabled = True) Or chkColor.Enabled = False Or Val(cmbGroupSize.Text) <= 1 And TestStatus = 0 Then
        cStaList = Format$(Val(txtParticipant.Text), "000")
    Else
        Set rstPart = mdbMain.OpenRecordset("SELECT Group FROM Entries WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND STA='" & Format$(Val(txtParticipant.Text), "000") & "'")
        If rstPart.RecordCount > 0 Then
            Set rstPart = mdbMain.OpenRecordset("SELECT Entries.Sta FROM Entries WHERE Entries.Code='" & TestCode & "' AND Entries.Status=" & TestStatus & " AND [Entries.Group]=" & rstPart.Fields("Group") & " ORDER BY  Entries.Position")
            If rstPart.RecordCount > 0 Then
                Do While Not rstPart.EOF
                    cStaList = cStaList & rstPart.Fields(0) & " "
                    rstPart.MoveNext
                Loop
            End If
        Else
            cStaList = Format$(Val(txtParticipant.Text), "000")
        End If
    End If
    
    
    cText = ""
    Do While cStaList <> ""
        If iCreateHtml = 1 Then cHtml = cHtmlMiddle
        Parse cSta, cStaList, " "
        cText = cText & mcExcelSeparator & cSta
        If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{startnum}", cSta)
        cQry = "SELECT Persons.*, Horses.*, Participants.Class, Participants.Club, Participants.Team "
        cQry = cQry & " FROM (Participants"
        cQry = cQry & " INNER JOIN Persons"
        cQry = cQry & " ON Participants.PersonId=Persons.PersonId)"
        cQry = cQry & " INNER JOIN Horses"
        cQry = cQry & " ON Participants.HorseId=Horses.HorseId"
        cQry = cQry & " WHERE Participants.STA='" & cSta & "'"
        Set rstPart = mdbMain.OpenRecordset(cQry)
        If rstPart.RecordCount > 0 Then
            With rstPart
                If miUseColors = 1 Then
                    cTemp = UCase$(Left$(GetParticipantsColor(cSta), 2))
                    cText = cText & mcExcelSeparator & cTemp & ""
                    If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Color}", cTemp)
                Else
                    cText = cText & mcExcelSeparator & ""
                    If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Color}", "")
                End If
                
                cText = cText & mcExcelSeparator & PrepareForExcel(.Fields("Name_First") & " " & rstPart.Fields("Name_Last"))
                cText = cText & mcExcelSeparator & .Fields("Class") & ""
                cText = cText & mcExcelSeparator & .Fields("Club") & ""
                cText = cText & mcExcelSeparator & .Fields("Team") & ""
                cText = cText & mcExcelSeparator & PrepareForExcel(.Fields("Name_horse"))
                cText = cText & mcExcelSeparator & .Fields("Horses.FEIFId")
                Select Case .Fields("Sex_Horse")
                Case 1
                   cSex = Translate("Stallion", mcLanguage)
                Case 2
                   cSex = Translate("Mare", mcLanguage)
                Case 3
                   cSex = Translate("Gelding", mcLanguage)
                Case Else
                   cSex = "--"
                End Select
                cText = cText & mcExcelSeparator & cSex
                If iCreateHtml = 1 Then
                    cHtml = Replace(cHtml, "{Rider}", .Fields("Name_First") & " " & rstPart.Fields("Name_Last"))
                    cHtml = Replace(cHtml, "{Class}", .Fields("Class") & "")
                    cHtml = Replace(cHtml, "{Club}", .Fields("Club") & "")
                    cHtml = Replace(cHtml, "{Team}", .Fields("Team") & "")
                    cHtml = Replace(cHtml, "{Horse}", .Fields("Name_horse"))
                    cHtml = Replace(cHtml, "{FEIFId}", .Fields("Horses.FEIFId") & "")
                    cHtml = Replace(cHtml, "{Gender}", cSex)
                End If

                Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Section=" & TestSection & " AND STA='" & cSta & "' ORDER BY Score DESC")
                If rstMarks.RecordCount > 0 Then
                    If TestStatus > 0 Then
                        Set rstTest = mdbMain.OpenRecordset("SELECT Name,Type_special,Type_Time FROM Testsections  WHERE Code='" & TestCode & "' AND Status=" & IIf(TestStatus > 1, 1, TestStatus) & " AND Section=" & TestSection)
                    Else
                        Set rstTest = mdbMain.OpenRecordset("SELECT Code + '-' + Test AS Name, Type_special,Type_Time from Tests WHERE Code='" & TestCode & "'")
                    End If
                    If rstTest.RecordCount > 0 Then
                        cText = cText & mcExcelSeparator & PrepareForExcel(rstTest.Fields("Name"))
                        If iCreateHtml = 1 Then
                            If IsNull(rstTest.Fields("Name")) Then
                                cHtml = Replace(cHtml, "{Section}", "")
                            Else
                                cHtml = Replace(cHtml, "{Section}", rstTest.Fields("Name"))
                            End If
                        End If
                    Else
                        cText = cText & mcExcelSeparator
                        If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Section}", "")
                    End If
                    
                    If rstTest.Fields("Type_Special") = 2 Then '---PP
                        Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND STA='" & cSta & "' ORDER BY Section")
                        If rstMarks.RecordCount > 0 Then
                            iTemp2 = 0
                            With rstMarks
                                Do While Not .EOF
                                    If iTemp2 = 1 Then
                                        If iCreateHtml = 1 Then
                                            cHtmlNew = cHtmlNew + vbCrLf + cHtml
                                            cHtml = cHtmlMiddle
                                        End If
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & vbCrLf
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                        cText = cText & mcExcelSeparator
                                    End If
                                    For iTemp = 1 To 5
                                        cText = cText & mcExcelSeparator & Replace(Format$(rstMarks.Fields("Mark" & Format$(iTemp)), TestMarkFormat), ",", ".")
                                        If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Mark" & Format$(iTemp) & "}", Replace(Format$(rstMarks.Fields("Mark" & Format$(iTemp)), TestMarkFormat), ",", "."))
                                    Next iTemp
                                    cText = cText & mcExcelSeparator & Replace(Format$(rstMarks.Fields("Score"), TestTotalFormat), ",", ".")
                                    If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Score}", Replace(Format$(rstMarks.Fields("Score"), TestTotalFormat), ",", "."))
                                    iTemp2 = 1
                                    .MoveNext
                                Loop
                            End With
                        End If
                    Else
                        For iTemp = 1 To 5
                            If rstTest.Fields("Type_Time") = 1 Then  '---races
                                Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND Section=" & Format$(iTemp) & " AND STA='" & cSta & "'")
                                If rstMarks.RecordCount > 0 Then
                                    cText = cText & mcExcelSeparator & Replace(Format$(rstMarks.Fields("Mark1"), TestTimeFormat), ",", ".")
                                    If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Mark" & Format$(iTemp) & "}", Replace(Format$(rstMarks.Fields("Mark1"), TestTimeFormat), ",", "."))
                                Else
                                    cText = cText & mcExcelSeparator & ""
                                    If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Mark" & Format$(iTemp) & "}", "")
                                End If
                            Else
                                cText = cText & mcExcelSeparator & Replace(Format$(rstMarks.Fields("Mark" & Format$(iTemp)), TestMarkFormat), ",", ".")
                                If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Mark" & Format$(iTemp) & "}", Replace(Format$(rstMarks.Fields("Mark" & Format$(iTemp)), TestMarkFormat), ",", "."))
                            End If
                        Next iTemp
                        cText = cText & mcExcelSeparator & Replace(Format$(rstMarks.Fields("Score"), TestTotalFormat), ",", ".")
                        If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Score}", Replace(Format$(rstMarks.Fields("Score"), TestTotalFormat), ",", "."))
                    End If
                    Set rstMarks = mdbMain.OpenRecordset("SELECT * FROM Results WHERE Code='" & TestCode & "' AND Status=" & TestStatus & " AND STA='" & cSta & "'")
                    If rstMarks.RecordCount > 0 Then
                        cText = cText & mcExcelSeparator
                        If rstMarks.Fields("Disq") = -1 Then
                            cText = cText & mcExcelSeparator & UCase$(Translate("ELIMINATED", mcLanguage))
                            If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Total}", UCase$(Translate("ELIMINATED", mcLanguage)))
                        ElseIf rstMarks.Fields("Disq") = -2 Then
                            cText = cText & mcExcelSeparator & Translate("Withdrawn", mcLanguage)
                            If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Total}", Translate("Withdrawn", mcLanguage))
                        Else
                            cText = cText & Replace(Format$(rstMarks.Fields("Score"), TestTotalFormat), ",", ".")
                            If iCreateHtml = 1 Then cHtml = Replace(cHtml, "{Total}", Replace(Format$(rstMarks.Fields("Score"), TestTotalFormat), ",", "."))
                        End If
                    End If
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    
                    rstTest.Close

                Else
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    If iCreateHtml = 1 Then
                        cHtml = Replace(cHtml, "{Test}", "")
                        For iTemp = 1 To 5
                            cHtml = Replace(cHtml, "{Mark" & Format$(iTemp) & "}", "")
                        Next iTemp
                        cHtml = Replace(cHtml, "{Score}", "")
                        cHtml = Replace(cHtml, "{Total}", "")
                    End If
                End If

                    
                cText = cText & vbCrLf
                rstMarks.Close
            End With
        End If
        rstPart.Close
        If iCreateHtml = 1 Then
            cHtmlNew = cHtmlNew + vbCrLf + cHtml
        End If
    Loop
    
    
    Set rstPart = Nothing
    Set rstMarks = Nothing
    Set rstTest = Nothing
    
    On Local Error GoTo CreateExcelError
    
    iExcelFileNum = FreeFile
    Open StrConv(mcExcelDir & "participant.csv", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
    Print #iExcelFileNum, cText
    Close #iExcelFileNum
    
    If iCreateHtml = 1 Then
        cHtmlNew = cHtmlNew + vbCrLf + cHtmlBottom
        cHtmlNew = RemoveFieldsFromTemplate(cHtmlNew)
        iExcelFileNum = FreeFile
        Open StrConv(mcExcelDir & "participant.html", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
        Print #iExcelFileNum, cHtmlNew
        Close #iExcelFileNum
    End If
    
Exit Sub

CreateExcelError:
    If iErrCounter < 5 Then
        iErrCounter = iErrCounter + 1
        Sleep 1
        Resume
    Else
        Exit Sub
    End If
Return

End Sub

Public Sub CreateExcelRanking(Optional iCount As Integer = 10)
    Dim iExcelFileNum As Integer
    Dim iKey As Integer
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim iTmpCount As Integer
    
    Dim iErrCounter As Integer
    Dim iEmptyline As Integer
    
    Dim cText As String
    Dim cTemp As String
    Dim cSex As String
    Dim cQry As String
    
    Dim curOldScore As Currency
    Dim cPosition As String
    Dim iPosition As Integer
    Dim iHighestPosition As Integer
    Dim cOldSta As String
    
    Dim cHtmlNew As String
    Dim cHtmlTop As String
    Dim cHtmlMiddle As String
    Dim cHtmlBottom As String
    Dim cHtml As String
    Dim cStart As String
    
    Dim rstStart As DAO.Recordset
    
    On Local Error Resume Next
    
    If mcExcelDir = "" Then Exit Sub
    
    If Dir$(mcExcelDir, vbDirectory) = "" Then
        If Err > 0 Then
            MsgBox Translate("Cannot access", mcLanguage) & " " & mcExcelDir, vbCritical
            mcExcelDir = ""
            Exit Sub
        Else
            MkDir mcExcelDir
        End If
    End If
    
    'mm: added complete starting order
    
    cStart = ""
    iTmpCount = 0
    
    cQry = "SELECT E.Sta, UCASE(LEFT(E.Color,2)), E.Group, PE.Name_First + ' '  + PE.Name_Last as Rider, PA.Class, PA.Club,PA.Team,HO.Name_horse,HO.FEIFID,IIF (HO.Sex_Horse=3,'Gelding' ,IIF (HO.Sex_Horse=2, 'Mare',IIF (HO.Sex_Horse=1, 'Stallion','')))"
    cQry = cQry & " FROM (((Entries AS E"
    cQry = cQry & " INNER JOIN Participants AS PA ON PA.STA=E.STA)"
    cQry = cQry & " INNER JOIN Persons PE ON PE.PersonId=PA.PersonId)"
    cQry = cQry & " INNER JOIN Horses AS HO ON HO.HorseId=PA.HorseId)"
    cQry = cQry & " WHERE E.Code='" & TestCode & "'"
    cQry = cQry & " AND E.Status=" & TestStatus
    cQry = cQry & " AND E.Sta Not In (Select Sta FROM Results WHERE disq>0)"
    cQry = cQry & " AND (IsNull(E.NoStart) Or E.NoStart = 0)"
    cQry = cQry & " ORDER BY E.Group, E.Position;"
    Set rstStart = mdbMain.OpenRecordset(cQry)
    If rstStart.RecordCount > 0 Then
        Do While Not rstStart.EOF
            iTmpCount = iTmpCount + 1
            If iTmpCount > 250 Then Exit Do
            For iTemp = 0 To 9
                cStart = cStart & mcExcelSeparator & PrepareForExcel(rstStart.Fields(iTemp))
            Next iTemp
            cStart = cStart & vbCrLf
            rstStart.MoveNext
        Loop
        iExcelFileNum = FreeFile
        Open StrConv(mcExcelDir & "start.csv", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
        Print #iExcelFileNum, cStart
        Close #iExcelFileNum
        If fraTime.Visible = True Then
            Open StrConv(mcExcelDir & "start_" & UnDotSpace(TestCode) & "_" & TestStatus & "_" & TestSection & ".csv", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
        Else
            Open StrConv(mcExcelDir & "start_" & UnDotSpace(TestCode) & "_" & TestStatus & ".csv", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
        End If
        Print #iExcelFileNum, cStart
        Close #iExcelFileNum
    End If
    rstStart.Close
    Set rstStart = Nothing
       
    If Dir$(mcExcelDir & "template-current.html") <> "" Then
        cHtml = GetHTMLTemplate(mcExcelDir & "template-current.html")
        cHtml = Replace(cHtml, "{Title}", frmMain.EventName)
        cHtml = Replace(cHtml, "{Test}", frmMain.TestCode & " - " & frmMain.TestName)
        iTemp = InStr(cHtml, "<tr")
        If iTemp > 1 Then
            cHtmlTop = Left$(cHtml, iTemp - 1)
            cHtml = Mid$(cHtml, iTemp)
            iTemp = InStrRev(cHtml, "</table>")
            If iTemp > 1 Then
                cHtmlBottom = Mid$(cHtml, iTemp)
                cHtmlMiddle = Left(cHtml, iTemp - 1)
            End If
        End If
    Else
        cHtml = ""
        cHtmlTop = ""
        cHtmlMiddle = ""
        cHtmlBottom = ""
        cHtmlNew = ""
    End If
    
    cHtmlNew = cHtmlTop
  
    
    iHighestPosition = GetHighestPosition(TestCode, TestStatus) - 1
    
    If dtaAlready.Recordset.RecordCount > 0 Then
        iPosition = 0
        dtaAlready.Recordset.MoveFirst
        For iTemp = 0 To dtaAlready.Recordset.RecordCount - 1
            cHtml = cHtmlMiddle
            With dtaAlready.Recordset
                If cOldSta <> .Fields("Participants.STA") Then
                    iPosition = iPosition + 1
                    If .Fields("Disq") < 0 Then
                        cPosition = "--"
                    ElseIf curOldScore <> .Fields("Results.Score") Then
                        cPosition = Format$(iPosition + iHighestPosition, "00")
                    End If
                    cText = cText & cPosition & mcExcelSeparator
                    cHtml = Replace(cHtml, "{Position}", cPosition)

                    cText = cText & .Fields("Participants.STA") & mcExcelSeparator
                    cHtml = Replace(cHtml, "{startnum}", .Fields("Participants.STA"))
                    
                    If miUseColors = 1 And TestStatus > 0 Then
                        cTemp = UCase$(Left$(GetParticipantsColor(.Fields("Participants.STA")), 2))
                        cText = cText & cTemp & mcExcelSeparator
                        cHtml = Replace(cHtml, "{Color}", cTemp)
                    Else
                        cText = cText & mcExcelSeparator
                        cHtml = Replace(cHtml, "{Color}", "")
                    End If

                    cText = cText & Trim$(.Fields("Name_First") & " " & .Fields("Name_Middle")) & " " & .Fields("Name_Last") & mcExcelSeparator
                    cHtml = Replace(cHtml, "{Rider}", Trim$(.Fields("Name_First") & " " & .Fields("Name_Middle")) & " " & .Fields("Name_Last"))
                    
                    cText = cText & .Fields("Class") & "" & mcExcelSeparator
                    cHtml = Replace(cHtml, "{Class}", .Fields("Class"))
                    
                    cText = cText & .Fields("Club") & "" & mcExcelSeparator
                    cHtml = Replace(cHtml, "{Club}", .Fields("Club"))
                    
                    cText = cText & .Fields("Team") & "" & mcExcelSeparator
                    cHtml = Replace(cHtml, "{Team}", .Fields("Team"))
                    
                    cText = cText & .Fields("Name_Horse") & mcExcelSeparator
                    cHtml = Replace(cHtml, "{Horse}", .Fields("Name_Horse"))
                    
                    cText = cText & .Fields("FEIFId") & mcExcelSeparator
                    cHtml = Replace(cHtml, "{FEIFId}", .Fields("FEIFId"))
                    
                    Select Case .Fields("Sex_Horse")
                    Case 1
                       cSex = Translate("Stallion", mcLanguage)
                    Case 2
                       cSex = Translate("Mare", mcLanguage)
                    Case 3
                       cSex = Translate("Gelding", mcLanguage)
                    Case Else
                       cSex = "--"
                    End Select
                    cText = cText & cSex & mcExcelSeparator
                    cHtml = Replace(cHtml, "{Gender}", cSex)
                 Else
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cText = cText & mcExcelSeparator
                    cHtml = Replace(cHtml, "{Position}", "")
                    cHtml = Replace(cHtml, "{Startnum}", "")
                    cHtml = Replace(cHtml, "{Color}", "")
                    cHtml = Replace(cHtml, "{Rider}", "")
                    cHtml = Replace(cHtml, "{Horse}", "")
                    cHtml = Replace(cHtml, "{FEIFId}", "")
                    cHtml = Replace(cHtml, "{Gender}", "")
                    cHtml = Replace(cHtml, "{Class}", "")
                    cHtml = Replace(cHtml, "{Club}", "")
                    cHtml = Replace(cHtml, "{Team}", "")
                    cHtml = Replace(cHtml, "{Total}", "")
                End If
                
                cText = cText & Replace(.Fields("Name"), ", ", " - ") & mcExcelSeparator
                cHtml = Replace(cHtml, "{Section}", .Fields("Name"))
                If fraMarks.Visible = True Then
                    For iTemp2 = 1 To 5
                        cText = cText & Replace(Format$(.Fields("Mark" & Format$(iTemp2)), TestMarkFormat), ",", ".") & mcExcelSeparator
                        cHtml = Replace(cHtml, "{Mark" & Format$(iTemp2) & "}", Replace(Format$(.Fields("Mark" & Format$(iTemp2)), TestMarkFormat), ",", "."))
                    Next iTemp2
                End If
                cText = cText & Replace(Format$(.Fields("Marks.Score"), TestTotalFormat), ",", ".") & mcExcelSeparator
                cHtml = Replace(cHtml, "{Score}", Replace(Format$(.Fields("Marks.Score"), TestTotalFormat), ",", "."))
                
                If cOldSta <> .Fields("Participants.STA") Then
                    If .Fields("Disq") = -2 Then
                        cText = cText & Translate("withdrawn", mcLanguage)
                        cHtml = Replace(cHtml, "{Total}", Translate("withdrawn", mcLanguage))
                    ElseIf .Fields("Disq") = -1 Then
                        cText = cText & UCase$(Translate("ELIMINATED", mcLanguage))
                        cHtml = Replace(cHtml, "{Total}", Translate("ELIMINATED", mcLanguage))
                    Else
                        cText = cText & Replace(Format$(.Fields("Results.Score"), TestTotalFormat), ",", ".")
                        cHtml = Replace(cHtml, "{Total}", Replace(Format$(.Fields("Results.Score"), TestTotalFormat), ",", "."))
                    End If
                End If
                cText = cText & mcExcelSeparator
                cText = cText & mcExcelSeparator
                cText = cText & vbCrLf
                curOldScore = .Fields("Results.Score")
                cOldSta = .Fields("Participants.STA")
                .MoveNext
            End With
            If cHtml <> "" Then
                cHtmlNew = cHtmlNew + vbCrLf + cHtml
            End If
        Next iTemp
    Else
        cText = cText & mcExcelSeparator
        cText = cText & Translate("No results available yet", mcLanguage) & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & mcExcelSeparator
        cText = cText & vbCrLf
    End If
    
    If cHtml <> "" Then
        cHtmlNew = cHtmlNew + vbCrLf + cHtmlBottom
    End If
    
    On Local Error GoTo CreateExcelError
        
    iExcelFileNum = FreeFile
    Open StrConv(mcExcelDir & "current.csv", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
    Print #iExcelFileNum, cText;
    Close #iExcelFileNum
    
    Open StrConv(mcExcelDir & UnDotSpace(dtaTest.Recordset.Fields("Code")) & "_" & TestStatus & ".csv", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
    Print #iExcelFileNum, cText;
    Close #iExcelFileNum
    
    If cHtmlNew <> "" Then
        iExcelFileNum = FreeFile
        Open StrConv(mcExcelDir & "current.html", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
        Print #iExcelFileNum, cHtmlNew
        Close #iExcelFileNum
        Open StrConv(mcExcelDir & UnDotSpace(dtaTest.Recordset.Fields("Code")) & "_" & TestStatus & ".html", vbLowerCase) For Output Access Write Shared As #iExcelFileNum
        Print #iExcelFileNum, cHtmlNew
        Close #iExcelFileNum
    End If
Exit Sub

CreateExcelError:
    If iErrCounter < 5 Then
        iErrCounter = iErrCounter + 1
        Sleep 1
        
        Resume
    Else
        Exit Sub
    End If
Return

End Sub
Function ComposeFinals(iStatus As Integer, Optional iNoReply As Integer = vbNo) As Integer
    Dim iKey As Integer
    Dim rstEntry As DAO.Recordset
    Dim rstStarted As DAO.Recordset
    Dim rstSplit As DAO.Recordset
    Dim iLowest As Integer
    Dim iHighest As Integer
    Dim iPosition As Integer
    Dim iFound As Integer
    Dim cOldSta As Integer
    Dim curOldMark As Currency
    Dim cQry As String
    Dim cStaList As String
    Dim iSplit As Integer
    Dim cSplitCode As String
    
    If iStatus = 3 Then
        If dtaTest.Recordset.Fields("Type_Special") = 3 Then
            iLowest = 23
            iHighest = 16
        Else
            iLowest = 15
            iHighest = 11
        End If
    ElseIf iStatus = 2 Then
        If dtaTest.Recordset.Fields("Type_Special") = 3 Then
            iLowest = 15
            iHighest = 8
        Else
            iLowest = 10
            iHighest = 6
        End If
    Else
        If dtaTest.Recordset.Fields("Type_Special") = 3 Then
            iLowest = 7
            iHighest = 1
        Else
            iLowest = 5
            iHighest = 1
        End If
    End If
    
    If iNoReply = vbNo Then
        If iStatus = 2 Then
            iKey = MsgBox(Translate("Compose starting order for this final?", mcLanguage) & vbCrLf & "(" & iHighest & "-" & iLowest & IIf(iStatus = 2, " " & Translate("plus the winner of the C-Final when applicable", mcLanguage) & ".", "") & ")", vbQuestion + vbYesNo + vbDefaultButton2)
        Else
            iKey = MsgBox(Translate("Compose starting order for this final?", mcLanguage) & vbCrLf & "(" & iHighest & "-" & iLowest & IIf(iStatus = 1, " " & Translate("plus the winner of the B-Final when applicable", mcLanguage) & ".", "") & ")", vbQuestion + vbYesNo + vbDefaultButton2)
        End If
    Else
        iKey = vbYes
    End If
    
    SetMouseHourGlass
    
    If iKey = vbYes Then
        If frmMain.chkSplitFinals = 1 Then
            '* prepare staring order for all finals involved
            '*
            mdbMain.Execute ("DELETE * FROM Entries WHERE Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "') And Status = " & iStatus)
            Set rstSplit = mdbMain.OpenRecordset("SELECT DISTINCT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "'")
            If rstSplit.RecordCount > 0 Then
                Do While Not rstSplit.EOF
                    cSplitCode = rstSplit.Fields(0)
                    GoSub SelectParticipants
                    rstSplit.MoveNext
                Loop
            Else
                cSplitCode = TestCode
                GoSub SelectParticipants
            End If
        Else
            mdbMain.Execute ("DELETE * FROM Entries WHERE Code='" & TestCode & "' AND Status=" & iStatus)
            cSplitCode = TestCode
            GoSub SelectParticipants
        End If
    Else
        ComposeFinals = False
    End If
    SetMouseNormal
    
Exit Function

SelectParticipants:
    'select participants
    cQry = "SELECT * FROM Results"
    cQry = cQry & " INNER JOIN Participants ON Results.STA=Participants.Sta"
    cQry = cQry & " WHERE Results.Code='" & TestCode & "'"
    cQry = cQry & " AND Results.Status=0"
    cQry = cQry & " AND NOT Results.STA IN (SELECT STA From Results WHERE Code='" & TestCode & "' AND Disq<0)"
    
    'check if preliminaries should be split into different finals
    If frmMain.chkSplitFinals = 1 Then
        cQry = cQry & " AND NOT Participants.Class & '' IN (SELECT Class & '' AS Class1 FROM TestSplits WHERE TestToSplit LIKE '" & TestCode & "' AND SplitToTest<>'" & cSplitCode & "')"
    End If
    If dtaTest.Recordset.Fields("Type_pre") = 2 Then
        cQry = cQry & " ORDER BY Results.Score ASC"
    Else
        cQry = cQry & " ORDER BY Results.Score DESC"
    End If
    
    cStaList = "|"
    Set rstStarted = mdbMain.OpenRecordset(cQry)
    If rstStarted.RecordCount > 0 Then
        Set rstEntry = mdbMain.OpenRecordset("SELECT * FROM Entries")
        iPosition = 0
        curOldMark = 0
        If rstStarted.RecordCount > 0 Then
            With rstStarted
                .MoveFirst
                Do While Not .EOF
                    If .Fields("Score") <> curOldMark Then
                        iPosition = .AbsolutePosition + 1
                        curOldMark = .Fields("Score")
                    End If
                    If iPosition > iLowest Then
                        Exit Do
                    ElseIf iPosition >= iHighest Then
                        With rstEntry
                            If iFound = False Then
                                If iStatus > 1 Then
                                    SetHighestPosition cSplitCode, iPosition, iStatus
                                End If
                            End If
                            iFound = True
                            .AddNew
                            .Fields("Code") = cSplitCode
                            .Fields("Sta") = rstStarted.Fields("Results.STA")
                            .Fields("Group") = 0
                            .Fields("Position") = iLowest - rstStarted.AbsolutePosition
                            .Fields("Status") = iStatus
                            .Fields("Deleted") = 0
                            .Fields("Timestamp") = Now
                            .Update
                            cStaList = cStaList & rstStarted.Fields("Results.STA") & "|"
                        End With
                    End If
                    .MoveNext
                Loop
            End With
        End If
        
        'MM
        'add winner C-Final to B-Final
        curOldMark = 0
        If iStatus = 2 And iFound = True Then
           cQry = "SELECT * FROM Results WHERE Code='" & cSplitCode & "' AND Status=3"
           If dtaTest.Recordset.Fields("Type_final") = 2 Then
                cQry = cQry & " ORDER BY Score ASC"
           Else
                cQry = cQry & " ORDER BY Score DESC"
           End If
           Set rstStarted = mdbMain.OpenRecordset(cQry)
           If rstStarted.RecordCount > 0 Then
                Do While Not rstStarted.EOF
                    If rstStarted.Fields("STA") <> cOldSta And InStr(cStaList, "|" & rstStarted.Fields("STA") & "|") = 0 Then
                        iPosition = iPosition + 1
                        With rstEntry
                            .AddNew
                            .Fields("Code") = cSplitCode
                            .Fields("Sta") = rstStarted.Fields("STA")
                            .Fields("Group") = 0
                            .Fields("Position") = 0
                            .Fields("Status") = iStatus
                            .Fields("Deleted") = 0
                            .Fields("Timestamp") = Now
                            curOldMark = rstStarted.Fields("Score")
                            .Update
                            cStaList = cStaList & rstStarted.Fields("STA") & "|"
                        End With
                        cOldSta = rstStarted.Fields("STA")
                    End If
                    rstStarted.MoveNext
                    If Not rstStarted.EOF Then
                        If rstStarted.Fields("Score") <> curOldMark Then Exit Do
                    End If
                Loop
            End If
        End If
        
        'add winner previous final to this final
        curOldMark = 0
        If iStatus = 1 And iFound = True Then
           cQry = "SELECT * FROM Results WHERE Code='" & cSplitCode & "' AND Status=2"
           If dtaTest.Recordset.Fields("Type_final") = 2 Then
                cQry = cQry & " ORDER BY Score ASC"
           Else
                cQry = cQry & " ORDER BY Score DESC"
           End If
           Set rstStarted = mdbMain.OpenRecordset(cQry)
           If rstStarted.RecordCount > 0 Then
                Do While Not rstStarted.EOF
                    If rstStarted.Fields("STA") <> cOldSta And InStr(cStaList, "|" & rstStarted.Fields("STA") & "|") = 0 Then
                        iPosition = iPosition + 1
                        With rstEntry
                            .AddNew
                            .Fields("Code") = cSplitCode
                            .Fields("Sta") = rstStarted.Fields("STA")
                            .Fields("Group") = 0
                            .Fields("Position") = 0
                            .Fields("Status") = iStatus
                            .Fields("Deleted") = 0
                            .Fields("Timestamp") = Now
                            curOldMark = rstStarted.Fields("Score")
                            .Update
                            cStaList = cStaList & rstStarted.Fields("STA") & "|"
                        End With
                        cOldSta = rstStarted.Fields("STA")
                    End If
                    rstStarted.MoveNext
                    If Not rstStarted.EOF Then
                        If rstStarted.Fields("Score") <> curOldMark Then Exit Do
                    End If
                Loop
            End If
        End If
        rstEntry.Close
        If iFound = False Then
            MsgBox Translate("No participants found that qualify for this final.", mcLanguage)
        Else
            AddColorsToFinals
            If miWriteLogDB Then
                WriteLogDBFinals
            End If
        End If
    Else
        MsgBox Translate("No participants found in preliminary rounds.", mcLanguage)
    End If
    rstStarted.Close
    Set rstStarted = Nothing
    Set rstEntry = Nothing
    
    ChangeCaption True
    ComposeFinals = True
Return

End Function
Sub CheckTieBreak(Optional iWarnAnyWay As Integer = False)
    Dim rstPos As DAO.Recordset
    Dim rstMarks As DAO.Recordset
    
    Dim cQry As String
    Dim iKey As Integer
    Dim iFirstPosCount As Integer
    Dim iFinished As Integer
    Dim curHighestScore As Currency
    
    'MM: the check should check if all sections are finished for all participants
    'problem: what to do with participants that give up half way?
    
    If TestStatus = 1 And fraMarks.Visible = True And (miDoNotCheckTieBreakAgain = False Or iWarnAnyWay = True) Then
        
        ' avoid repeted asking
        miDoNotCheckTieBreakAgain = True
        
        ' what is tye mark for position 1
        Set rstPos = mdbMain.OpenRecordset("SELECT Sta,Score FROM Results WHERE Code='" & TestCode & "' AND Status=1 AND Position=1")
        If rstPos.RecordCount > 0 Then
            curHighestScore = rstPos.Fields("Score")
            ' are there more participants with this marks
            Set rstPos = mdbMain.OpenRecordset("SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=1 AND Score=" & Replace(Format(curHighestScore), ",", "."))
            If rstPos.RecordCount > 0 Then
                rstPos.MoveLast
                iFirstPosCount = rstPos.RecordCount
                If iFirstPosCount > 1 Then
                    If iWarnAnyWay = False Then
                        ' did all participants finish all sections?
                        ' but what to do with participants that did give up?
                        Set rstPos = mdbMain.OpenRecordset("SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=1 AND Score>0")
                        If rstPos.RecordCount > 0 Then
                            iFinished = True
                            Do While Not rstPos.EOF And iFinished = True
                                Set rstMarks = mdbMain.OpenRecordset("SELECT COUNT(STA) FROM Marks WHERE Code='" & TestCode & "' AND Status=1 AND STA='" & rstPos.Fields("STA") & "' AND Score>0")
                                If rstMarks.Fields(0) <> frmMain.tbsSection(2).Tabs.Count Then
                                    iFinished = False
                                End If
                                rstMarks.Close
                                rstPos.MoveNext
                            Loop
                            Set rstMarks = Nothing
                       End If
                    Else
                        '  when printing a check is needed anyway
                        iFinished = True
                    End If
                    If iFinished = True Then
                        ' Looks like we are ready
                        Set rstPos = mdbMain.OpenRecordset("SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=1 AND Alltimes<>''")
                        If rstPos.RecordCount > 0 Then
                            iKey = MsgBox(iFirstPosCount & " " & Translate("Participants share the first position. The judges have already decided upon a tiebreak. Do you want to change the winner of that tiebreak ?", mcLanguage), vbExclamation + vbYesNo + vbDefaultButton2)
                        Else
                            iKey = MsgBox(iFirstPosCount & " " & Translate("Participants share the first position. The judges have to decide upon a tiebreak. Do you want to mark the winner of the tiebreak ?", mcLanguage), vbExclamation + vbYesNo + vbDefaultButton1)
                        End If
                        If iKey = vbYes Then
                            mdbMain.Execute "UPDATE Results SET Alltimes='' WHERE Results.Code='" & TestCode & "' AND Results.Status=1"
                            cQry = "SELECT Results.Sta "
                            cQry = cQry & " & '  -  ' & Persons.Name_First "
                            cQry = cQry & " & ' ' & Persons.Name_Last "
                            cQry = cQry & " & IIF(Participants.Class<>'',' [' & Participants.Class & ']','')"
                            cQry = cQry & " & ' - ' & Horses.Name_Horse "
                            cQry = cQry & " as cList"
                            cQry = cQry & " FROM ((Results INNER JOIN Participants ON Results.STA = Participants.STA) INNER JOIN Persons ON Participants.PersonID = Persons.PersonID) INNER JOIN Horses ON Participants.HorseID = Horses.HorseID"
                            cQry = cQry & " WHERE Results.Code='" & TestCode & "' AND Results.Status=1"
                            cQry = cQry & " AND Results.Score=" & Replace(Format(curHighestScore), ",", ".")
                            cQry = cQry & " ORDER BY Results.STA"
                            With frmToolBox
                                .intChecked = True
                                .intSingleChoice = True
                                .strQry = cQry
                                .Caption = ClipAmp("Tiebreak: mark the winner")
                                .Show 1, Me
                            End With
                            If Tempvar <> "" Then
                                mdbMain.Execute "UPDATE Results SET Alltimes='1' WHERE Results.Code='" & TestCode & "' AND Results.Status=1 AND Results.Position=1 AND Results.Sta='" & Left$(Tempvar, 3) & "'"
                                mdbMain.Execute "UPDATE Results SET Alltimes='2' WHERE Results.Code='" & TestCode & "' AND Results.Status=1 AND Results.Position=1 AND Results.AllTimes<>'1'"
                            End If
                            dtaAlready.Recordset.Requery
                            ChangeCaption True
                            TestInfoMessage = "Winner of the tiebreak has been set."
                            Tempvar = ""
                        Else
                            Set rstPos = mdbMain.OpenRecordset("SELECT Sta FROM Results WHERE Code='" & TestCode & "' AND Status=1 AND AllTimes<>''")
                            If rstPos.RecordCount > 0 Then
                                iKey = MsgBox(Translate("Do you want to remove previous winner of the tiebreak?", mcLanguage), vbQuestion + vbYesNo + vbDefaultButton2)
                                If iKey = vbYes Then
                                    mdbMain.Execute "UPDATE Results SET Alltimes='' WHERE Results.Code='" & TestCode & "' AND Results.Status=1 AND Results.AllTimes<>''"
                                    dtaAlready.Recordset.Requery
                                    ChangeCaption True
                                End If
                            End If
                            TestInfoMessage = "Winner of the tiebreak has been removed."
                            Tempvar = ""
                        End If
                    End If
                End If
            End If
        End If
        rstPos.Close
        Set rstPos = Nothing
    End If
        
    StatusMessage ""
    SetMouseNormal
End Sub
Public Sub cmdOkClick()
   Dim rstOk As Recordset
   Dim rstSectionMark As DAO.Recordset
   Dim cSta As String
   Dim iItem As Integer
   Dim cTemp As String
   Dim iKey As Integer
   Dim cLogLine As String
   Dim iCountZero As Integer
   Dim iDisq As Integer
   Dim iOldDisq As Integer
   Dim cLogDB As String
   Dim iLogDBAction As Integer
   Dim iDoLogDB As Boolean
   
   miDoNotCheckTieBreakAgain = False
   
   If fraJudges.Visible = True Then
        For iItem = 0 To TestJudges - 1
            cTemp = cTemp & txtMarks(iItem).Text
            cLogDB = cLogDB & txtMarks(iItem).Text
            If iItem < (TestJudges - 1) Then
                cLogDB = cLogDB & " - "
            Else
                cLogDB = cLogDB & ": Total "
            End If
        Next iItem
   ElseIf fraTime.Visible = True Then
        cTemp = txtTime
   End If
    
   If cTemp = "" Then
      If fraMarks.Visible = True Then
        iKey = MsgBox(Translate("No marks entered. Store results anyway?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
      ElseIf fraTime.Visible = True Then
        iKey = MsgBox(Translate("No time entered. Store results anyway?", mcLanguage), vbYesNo + vbQuestion + vbDefaultButton2)
      End If
      If iKey = vbNo Then
         SetFocusTo txtParticipant
         Exit Sub
      End If
   End If
   If lblParticipant.Caption <> "" Then
      cSta = Format$(Val(txtParticipant.Text), "000")
      'is this participant already an entry?
      Set rstOk = mdbMain.OpenRecordset("SELECT * FROM Entries WHERE Sta='" & cSta & "' AND Status=" & Me.TestStatus & " AND Code='" & Me.TestCode & "'")
      With rstOk
         If .RecordCount = 0 Then
            .AddNew
            .Fields("Sta") = cSta
            .Fields("Status") = Me.TestStatus
            .Fields("Code") = Me.TestCode
            .Fields("Position") = 0
            .Fields("RR") = False
            .Fields("Group") = 0
            .Fields("TimeStamp") = Now
            .Update
         End If
      End With
      
      'are there already marks?
      Set rstOk = mdbMain.OpenRecordset("SELECT * FROM Marks WHERE Sta='" & cSta & "' AND Status=" & Me.TestStatus & " AND Code='" & Me.TestCode & "' AND Section=" & Me.TestSection)
      With rstOk
         If .RecordCount > 0 Then
            .Edit
         Else
            .AddNew
         End If
         .Fields("Sta") = cSta
         cLogLine = cSta
         .Fields("Status") = Me.TestStatus
         .Fields("Code") = Me.TestCode
         cLogLine = cLogLine & " " & TestCode & "-" & TestStatus & "-" & TestSection
         .Fields("Section") = Me.TestSection
         
         If fraMarks.Visible = True Then
            For iItem = 0 To TestJudges - 1
               .Fields("Mark" & Format$(iItem + 1)) = MakeStringValue(txtMarks(iItem).Text)
               If .Fields("Mark" & Format$(iItem + 1)) = 0 Then
                    If iItem < 4 Then
                        iCountZero = iCountZero + 1
                    End If
               End If
               If dtaTest.Recordset.Fields("Type_Special") = 2 And chkFlag.Value <> 0 And iItem = 4 Then
                   cLogLine = cLogLine & " 0.00"
               Else
                   cLogLine = cLogLine & " " & .Fields("Mark" & Format$(iItem + 1))
               End If
            Next iItem
            .Fields("Score") = MakeStringValue(txtScore.Text)
            cLogDB = cLogDB & " " & txtScore.Text
         ElseIf fraTime.Visible = True Then
            .Fields("Mark1") = MakeStringValue(txtTime.Text)
            .Fields("Score") = MakeStringValue(txtScore.Text)
            cLogLine = cLogLine & " " & .Fields("Mark1")
            cLogDB = cLogDB & "> " & txtScore.Text
         End If
         .Fields("Flag") = IIf(chkFlag.Value = 0, False, True)
         .Fields("TimeStamp") = Now
         .Update
         LogLine cLogLine
         .Close
      End With
      
      'is there already a result record?
      Set rstOk = mdbMain.OpenRecordset("SELECT * FROM Results WHERE Sta='" & cSta & "' AND Status=" & Me.TestStatus & " AND Code='" & Me.TestCode & "'")
      With rstOk
         If .RecordCount > 0 Then
            .Edit
            iLogDBAction = 2 'changing existing result
         Else
            .AddNew
            iLogDBAction = 1 'new result
         End If
         
         iOldDisq = IIf(IsNull(.Fields("Disq")), 0, .Fields("Disq"))
         
         .Fields("Sta") = cSta
         .Fields("Status") = Me.TestStatus
         .Fields("Code") = Me.TestCode
         .Fields("Position") = 0
         If fraMarks.Visible = True Then
            .Fields("Score") = CalculateResult(cSta)
            If dtaTest.Recordset.Fields("Type_time") = 3 Then
                .Fields("AllTimes") = Left$(CalculatePP1Times(cSta), .Fields("AllTimes").Size)
            End If
         ElseIf fraTime.Visible = True Then
            .Fields("Score") = CalculateTime(cSta)
            .Fields("AllTimes") = Left$(CalculateAllTimes(cSta), .Fields("AllTimes").Size)
            .Fields("Time") = .Fields("Score")
         End If
         If IsNull(.Fields("DISQ")) Then
            .Fields("DISQ") = 0
         End If
         .Fields("TimeStamp") = Now
         .Fields("FR") = False
         .Update
         
         .Close
         
         If iCountZero >= 3 Then
            If dtaTest.Recordset.Fields("Type_special") = 2 And dtaTest.Recordset.Fields("Out_Fin") = 1 Then 'PP2
                'nothing happens as the rider has two chances anyway
            ElseIf dtaTest.Recordset.Fields("Type_special") = 2 Then 'PP1 or PP2
                'shit happens -> eliminated
                'changed in 2017 to allow second run allways
                '+++If Me.chkDisqualified.Value = 0 Then
                '+++    MsgBox Translate("This participant is eliminated.", mcLanguage)
                '+++    Me.chkDisqualified.Value = 1
                '+++    chkWithdrawn.Value = 0
                '+++    chkWithdrawn.Enabled = False
                '+++    ParticipantDisqWith cSta, TestCode, TestStatus, -1
                '+++End If
            End If
         End If
      End With
      
      Set rstOk = Nothing
            
      If miExcelFiles <> 0 Then
          CreateExcelParticipant
      End If
      
      If fraTime.Visible = False And TestStatus = 1 Then
          mdbMain.Execute "UPDATE Results SET Alltimes='' WHERE Results.Code='" & TestCode & "' AND Results.Status=1"
      End If
      
      ClearMarks
   End If
   
   LastUsedSta = cSta
   
   LookUpRelevantParticipants
   
   txtScore.BackColor = QBColor(15)
   miNoBackupNow = False
   
   CreateHTMLDetails cSta, True
      
   If miWriteLogDB Then
        'iDoLogDB = WriteLogDBMarks(EventName, TestCode, TestStatus, cSta, iLogDBAction, cLogDB, TestSection)
        iDoLogDB = WriteLogDBMarks2(EventName, TestCode, TestStatus, cSta)
   End If
   
   If dblstNotYet.Visible = True And dtaNotYet.Recordset.RecordCount > 0 Then
        SetFocusTo dblstNotYet
        dblstNotYet.BoundText = dtaNotYet.Recordset.Fields("cList")
    Else
        SetFocusTo txtParticipant
   End If
End Sub

Private Sub winsock_Close()
    blnConnected = False
    winsock.Close
End Sub

Private Sub winsock_Connect()
    blnConnected = True
End Sub

Private Sub winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strResponse As String

    winsock.GetData strResponse, vbString, bytesTotal
    strResponse = FormatLineEndings(strResponse)
    
    ' append to string since data arrives in multiple packets:
    winsockresponse = winsockresponse & strResponse
End Sub

Private Sub winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox Description, vbExclamation, "A problem with the internet connection occurred."
    winsock.Close
End Sub
' forms HTTP request to read the version text file on FEIF's server
Private Sub CheckWebForUpdate()
    Dim eUrl As URL
    
    Dim webupdates As String
    Dim strMethod As String
    Dim strData As String
    Dim strPostData As String
    Dim strHeaders As String
    
    Dim strHTTP As String
    Dim X As Integer
    
    strPostData = ""
    strHeaders = ""
    strMethod = "GET"
    
    ' abort if winsock is currently in use:
    If blnConnected Then Exit Sub
    
    'Is the update URL stored in the database?
    webupdates = GetVariable("webupdates")
    If webupdates = "" Then
        webupdates = "https://feif.glitnir.nl/icehorse/icehorsetools_version.txt"
    Else
        If Not Left$(LCase$(webupdates), 7) = "http://" Then
            webupdates = "https://" & webupdates
        ElseIf Not Left$(LCase$(webupdates), 7) = "https://" Then
            webupdates = "https://" & webupdates
        End If
        If Not Right$(webupdates, 4) = ".txt" Then
            If Not Right$(webupdates, 1) = "/" Then
                webupdates = webupdates & "/"
            End If
            webupdates = webupdates & "icehorsetools_version.txt"
        End If
    End If
    
    ' get the url
    eUrl = ExtractUrl(webupdates)
    
    If eUrl.Host = vbNullString Then
        MsgBox "Invalid Host", vbCritical, "ERROR"
        Exit Sub
    End If
    
    ' configure winsock
    winsock.Protocol = sckTCPProtocol
    winsock.RemoteHost = eUrl.Host
    
    If eUrl.Scheme = "http" Then
        If eUrl.Port > 0 Then
            winsock.RemotePort = eUrl.Port
        Else
            winsock.RemotePort = 80
        End If
    ElseIf eUrl.Scheme = vbNullString Then
        winsock.RemotePort = 80
    Else
        MsgBox "Invalid protocol schema"
    End If
    
    ' build encoded data the data is url encoded in the form
    ' var1=value&var2=value
    strData = ""
    'For X = 0 To txtVariableName.Count - 1
    '    If txtVariableName(X).Text <> vbNullString Then
    '        strData = strData & URLEncode(txtVariableName(X).Text) & "=" & _
    '                        URLEncode(txtVariableValue(X).Text) & "&"
    '    End If
    'Next X
    
    If eUrl.Query <> vbNullString Then
        eUrl.URI = eUrl.URI & "?" & eUrl.Query
    End If
    
    ' check if any variables were supplied
    If strData <> vbNullString Then
        strData = Left(strData, Len(strData) - 1)
        
        
        If strMethod = "GET" Then
            ' if this is a GET request then the URL encoded data
            ' is appended to the URI with a ?
            If eUrl.Query <> vbNullString Then
                eUrl.URI = eUrl.URI & "&" & strData
            Else
                eUrl.URI = eUrl.URI & "?" & strData
            End If
        Else
            ' if it is a post request, the data is appended to the
            ' body of the HTTP request and the headers Content-Type
            ' and Content-Length added
            strPostData = strData
            strHeaders = "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
                         "Content-Length: " & Len(strPostData) & vbCrLf
                         
        End If
    End If
            
    ' get any aditional headers and add them
    'For X = 0 To txtHeaderName.Count - 1
    '    If txtHeaderName(X).Text <> vbNullString Then
    '        strHeaders = strHeaders & txtHeaderName(X).Text & ": " & _
    '                        txtHeaderValue(X).Text & vbCrLf
    '    End If
    'Next X
    
    ' clear the old HTTP response
    winsockresponse = ""
    
    ' build the HTTP request in the form
    '
    ' {REQ METHOD} URI HTTP/1.0
    ' Host: {host}
    ' {headers}
    '
    ' {post data}
    strHTTP = strMethod & " " & eUrl.URI & " HTTP/1.0" & vbCrLf
    strHTTP = strHTTP & "Host: " & eUrl.Host & vbCrLf
    strHTTP = strHTTP & strHeaders
    strHTTP = strHTTP & vbCrLf
    strHTTP = strHTTP & strPostData

    
    winsock.Connect
    
    ' wait for a connection
    While Not blnConnected
        DoEvents
    Wend
    
    ' send the HTTP request
    winsock.SendData strHTTP
End Sub

Public Function ComposeFinalsAllowed() As Integer
    Dim cQry As String
    Dim rst As DAO.Recordset
    
    cQry = "SELECT * FROM Results "
    cQry = cQry & " WHERE Code Like '" & frmMain.TestCode & "' "
    cQry = cQry & " AND Status = " & frmMain.TestStatus
    cQry = cQry & " AND Disq > -2"
    Set rst = mdbMain.OpenRecordset(cQry)
    If rst.RecordCount = 0 Then
        ComposeFinalsAllowed = True
    Else
        ComposeFinalsAllowed = False
    End If
    rst.Close
    Set rst = Nothing
End Function
Public Function SplitFinalsAllowed() As Integer
    Dim cQry As String
    Dim rst As DAO.Recordset
    
    cQry = "SELECT DISTINCT Participants.Class FROM Results "
    cQry = cQry & " INNER JOIN Participants ON Results.STA = Participants.STA "
    cQry = cQry & " WHERE Results.Code Like '" & frmMain.TestCode & "' "
    cQry = cQry & " AND Results.Status = 0 "
    cQry = cQry & " AND Results.Disq > -1"
    
    Set rst = mdbMain.OpenRecordset(cQry)
    If rst.RecordCount > 0 Then
        rst.MoveLast
        If rst.RecordCount > 1 Then
            '* It is allowed to split finals
            '* but are there already marks entered
            cQry = "SELECT * FROM Results "
            cQry = cQry & " WHERE Code IN (SELECT SplitToTest FROM TestSplits WHERE TestToSplit LIKE '" & frmMain.TestCode & "')"
            cQry = cQry & " AND Status > 0"
            cQry = cQry & " AND Disq > -1"
            
            Set rst = mdbMain.OpenRecordset(cQry)
            If rst.RecordCount > 0 Then
                '* there are already marks entered, so not allowed to redefine split
                SplitFinalsAllowed = False
            Else
                SplitFinalsAllowed = True
            End If
        Else
            SplitFinalsAllowed = False
        End If
    End If
    rst.Close
    Set rst = Nothing
End Function

Public Function GetParticipantsPosition(cSta As String, cCode As String, iStatus As Integer) As Integer
    Dim rst As DAO.Recordset
    Dim cQry As String
    
    cQry = "SELECT Position "
    cQry = cQry & " FROM Results "
    cQry = cQry & " WHERE Sta='" & cSta & "' "
    cQry = cQry & " AND Code='" & cCode & "' "
    cQry = cQry & " AND Status=" & iStatus
    
    Set rst = mdbMain.OpenRecordset(cQry)
    If rst.RecordCount > 0 Then
        GetParticipantsPosition = rst.Fields(0)
    Else
        GetParticipantsPosition = 0
    End If
    rst.Close
    Set rst = Nothing
End Function

Public Function GetParticipantsColor(cSta As String) As String
    Dim rst As DAO.Recordset
        
    GetParticipantsColor = ""
    On Local Error Resume Next
    
    Set rst = mdbMain.OpenRecordset("SELECT Color FROM Entries WHERE Sta='" & cSta & "' AND Code='" & TestCode & "' AND Status=" & TestStatus)
    
    If rst.RecordCount > 0 Then
        GetParticipantsColor = rst.Fields(0)
    End If
    rst.Close
    Set rst = Nothing
End Function
Public Function PrepareForExcel(cExcel As String) As String
    Dim cTemp As String
    cTemp = cExcel & ""
    PrepareForExcel = Replace(Replace(Replace(Replace(cTemp, ",", " - "), ";", " - "), vbTab, "-"), "  ", " ")
    
End Function
Public Function RemoveFieldsFromTemplate(cHtml As String) As String
    Dim iTemp As Integer
    Dim iTemp2 As Integer
    Dim cTemp As String
    Dim cRemove As String
    
    cTemp = cHtml
    iTemp = InStr(cTemp, "{")
    Do While iTemp > 0
        iTemp2 = InStr(cTemp, "}")
        If iTemp2 > iTemp Then
            cRemove = Mid$(cTemp, iTemp, iTemp2 - iTemp + 1)
            cTemp = Replace(cTemp, cRemove, "&nbsp;")
            iTemp = InStr(cTemp, "{")
        Else
            iTemp = 0
        End If
    Loop
    RemoveFieldsFromTemplate = cTemp
    
End Function

