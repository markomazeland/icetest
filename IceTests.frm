VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frmTests 
   Caption         =   "Test Properties"
   ClientHeight    =   8265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   Begin VB.Data dtaTimeTables 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   9840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaType_Time 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaType_Special 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaHandling 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data dtaPrelim 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame fraTbsTest 
      Height          =   2775
      Index           =   2
      Left            =   5760
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtPrelimHi 
         Height          =   285
         Index           =   0
         Left            =   5160
         TabIndex        =   68
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtPrelimLo 
         Height          =   285
         Index           =   0
         Left            =   4560
         TabIndex        =   67
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkPrelimOut 
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   53
         ToolTipText     =   "May be taken out when calculating result (normally not)"
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cmbTimeDecimals 
         DataField       =   "Time_decimals"
         DataSource      =   "dtaTest"
         Height          =   315
         Index           =   0
         Left            =   3000
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox cmbMarkDecimals 
         DataField       =   "Mark_Decimals"
         DataSource      =   "dtaTest"
         Height          =   315
         Index           =   0
         ItemData        =   "IceTests.frx":0000
         Left            =   3000
         List            =   "IceTests.frx":000D
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox cmbPrelimName 
         CausesValidation=   0   'False
         Height          =   315
         Index           =   0
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   21
         ToolTipText     =   "Name of the section (preferrably in English)"
         Top             =   840
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtPrelimFactor 
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   22
         ToolTipText     =   "Multiplication factor (usually 1)"
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin MSDBCtls.DBCombo dbcPre 
         Bindings        =   "IceTests.frx":001A
         DataField       =   "Type_Pre"
         DataSource      =   "dtaTest"
         Height          =   315
         Left            =   720
         TabIndex        =   18
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Label"
         BoundColumn     =   "Code"
         Text            =   ""
      End
      Begin VB.Label lblPrelimHi 
         Caption         =   "&Highest"
         Height          =   375
         Left            =   5160
         TabIndex        =   66
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblPrelimLo 
         Caption         =   "&Lowest"
         Height          =   255
         Left            =   4560
         TabIndex        =   65
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblPrelimOut 
         Caption         =   "&Out"
         Height          =   255
         Left            =   6000
         TabIndex        =   52
         ToolTipText     =   "May be taken out when calculating result"
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblDec 
         Alignment       =   1  'Right Justify
         Caption         =   "Decimals"
         Height          =   375
         Index           =   0
         Left            =   2280
         TabIndex        =   19
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblPre 
         Alignment       =   1  'Right Justify
         Caption         =   "Judgement"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblPrelimSection 
         Caption         =   "&Section"
         Height          =   255
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Name of the section "
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblPrelimFactor 
         Caption         =   "&Factor"
         Height          =   255
         Left            =   3960
         TabIndex        =   38
         ToolTipText     =   "Multiplication factor "
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame fraTbsTest 
      Height          =   3255
      Index           =   4
      Left            =   5760
      TabIndex        =   32
      Top             =   4440
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtOther 
         DataField       =   "ScaleStep"
         DataSource      =   "dtaTimeTables"
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   80
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox txtOther 
         DataField       =   "ScaleRange"
         DataSource      =   "dtaTimeTables"
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   78
         Top             =   2400
         Width           =   1575
      End
      Begin VB.TextBox txtOther 
         DataField       =   "ScaleSlow"
         DataSource      =   "dtaTimeTables"
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   76
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtOther 
         DataField       =   "ScaleFast"
         DataSource      =   "dtaTimeTables"
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   74
         Top             =   1680
         Width           =   1575
      End
      Begin MSDBCtls.DBCombo dbcOther 
         Bindings        =   "IceTests.frx":0035
         DataField       =   "Out_Fin"
         DataSource      =   "dtaTest"
         Height          =   315
         Index           =   0
         Left            =   3240
         TabIndex        =   62
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         ListField       =   "Section"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo dbcOther 
         Bindings        =   "IceTests.frx":004D
         DataField       =   "Type_Special"
         DataSource      =   "dtaTest"
         Height          =   315
         Index           =   1
         Left            =   3240
         TabIndex        =   63
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Label"
         BoundColumn     =   "Code"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo dbcOther 
         Bindings        =   "IceTests.frx":006B
         DataField       =   "Type_time"
         DataSource      =   "dtaTest"
         Height          =   315
         Index           =   2
         Left            =   3240
         TabIndex        =   64
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Label"
         BoundColumn     =   "Code"
         Text            =   ""
      End
      Begin VB.Label lblOther 
         Caption         =   "Steps between marks (when applicable)"
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   79
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label lblOther 
         Caption         =   "Highest mark (default=10.0)"
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   77
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblOther 
         Caption         =   "Slowest time"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   75
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblOther 
         Caption         =   "Fastest time"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   73
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblGeneralText 
         AutoSize        =   -1  'True
         Caption         =   "Most of the time no special settings are required (leave them to their default settings)."
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
         Left            =   240
         TabIndex        =   59
         Top             =   240
         Width           =   7305
      End
      Begin VB.Label lblOther 
         Caption         =   "Special settings when times are used"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   58
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblOther 
         Caption         =   "Special settings for this test"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   57
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblOther 
         Caption         =   "Number of sections that will be taken out when calculating results (default=0)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   56
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.Frame fraTbsTest 
      Height          =   2895
      Index           =   3
      Left            =   5760
      TabIndex        =   23
      Top             =   3000
      Visible         =   0   'False
      Width           =   6255
      Begin VB.TextBox txtFinalsLo 
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   70
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtFinalsHi 
         Height          =   285
         Index           =   0
         Left            =   5400
         TabIndex        =   69
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkRecycle 
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   55
         ToolTipText     =   "Re-use marks from preliminary rounds (normally not)"
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CheckBox chkFinalsOut 
         Height          =   255
         Index           =   0
         Left            =   5880
         TabIndex        =   51
         ToolTipText     =   "May be taken out when calculating result (normally not)"
         Top             =   840
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox cmbTimeDecimals 
         DataField       =   "Time_Decimals"
         DataSource      =   "dtaTest"
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox cmbMarkDecimals 
         DataField       =   "Mark_decimals"
         DataSource      =   "dtaTest"
         Height          =   315
         Index           =   1
         ItemData        =   "IceTests.frx":0086
         Left            =   3240
         List            =   "IceTests.frx":0093
         TabIndex        =   41
         Top             =   120
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtFinalsFactor 
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   31
         ToolTipText     =   "Multiplication factor (usually 1)"
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.ComboBox cmbFinalsName 
         BackColor       =   &H80000009&
         CausesValidation=   0   'False
         Height          =   315
         Index           =   0
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   30
         ToolTipText     =   "Name of the section (preferrably in English)"
         Top             =   840
         Visible         =   0   'False
         Width           =   3735
      End
      Begin MSDBCtls.DBCombo dbcFin 
         Bindings        =   "IceTests.frx":00A0
         DataField       =   "Type_Final"
         DataSource      =   "dtaTest"
         Height          =   315
         Left            =   840
         TabIndex        =   25
         Top             =   120
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Label"
         BoundColumn     =   "Code"
         Text            =   ""
      End
      Begin VB.Label lblFinalsLo 
         Caption         =   "&Lowest"
         Height          =   255
         Left            =   4800
         TabIndex        =   72
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblFinalsHi 
         Caption         =   "&Highest"
         Height          =   375
         Left            =   5400
         TabIndex        =   71
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblRecycle 
         Caption         =   "&Re-use"
         Height          =   255
         Left            =   6120
         TabIndex        =   54
         ToolTipText     =   "Re-use marks from preliminary rounds"
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblFinalsOut 
         Caption         =   "&Out"
         Height          =   255
         Left            =   5760
         TabIndex        =   50
         ToolTipText     =   "May be taken out when calculating result"
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lblDec 
         Alignment       =   1  'Right Justify
         Caption         =   "Decimals"
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   26
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lblFin 
         Alignment       =   1  'Right Justify
         Caption         =   "Judgement"
         Height          =   375
         Left            =   9360
         TabIndex        =   24
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label lblFinalsFactor 
         Caption         =   "&Factor"
         Height          =   255
         Left            =   3960
         TabIndex        =   29
         ToolTipText     =   "Multiplication factor"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblFinalsSection 
         Caption         =   "&Section"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         ToolTipText     =   "Name of the section "
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame fraTbsTest 
      Height          =   4215
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   5295
      Begin VB.TextBox txtComments 
         DataField       =   "Comments"
         DataSource      =   "dtaTest"
         Height          =   495
         Left            =   1560
         TabIndex        =   61
         Top             =   3240
         Width           =   1215
      End
      Begin VB.ComboBox cmbGroup 
         DataField       =   "Groupsize"
         DataSource      =   "dtaTest"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   49
         Top             =   2880
         Width           =   1215
      End
      Begin VB.ComboBox cmbJudges 
         DataField       =   "Num_J"
         DataSource      =   "dtaTest"
         Height          =   315
         ItemData        =   "IceTests.frx":00BB
         Left            =   1560
         List            =   "IceTests.frx":00C2
         Sorted          =   -1  'True
         TabIndex        =   47
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtSponsor 
         DataField       =   "Sponsor"
         DataSource      =   "dtaTestInfo"
         Height          =   495
         Left            =   1560
         MaxLength       =   255
         TabIndex        =   44
         ToolTipText     =   "The sponsor of this test"
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton cmdSplitFinals 
         Height          =   375
         Left            =   4560
         Picture         =   "IceTests.frx":00C9
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkSplitFinals 
         Caption         =   "&Split finals"
         CausesValidation=   0   'False
         DataSource      =   "dtaTestInfo"
         Height          =   375
         Left            =   1560
         TabIndex        =   42
         Top             =   2040
         Width           =   2415
      End
      Begin VB.CommandButton cmdNewQual 
         Height          =   375
         Left            =   4560
         Picture         =   "IceTests.frx":0213
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Create a new qualification category"
         Top             =   1320
         Width           =   735
      End
      Begin VB.Data dtaName 
         Caption         =   "Data1"
         Connect         =   "Access 2000;"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4560
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4080
         Visible         =   0   'False
         Width           =   1140
      End
      Begin MSDBCtls.DBCombo dbcQual 
         Bindings        =   "IceTests.frx":0315
         DataField       =   "Qualification"
         DataSource      =   "dtaTest"
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         ToolTipText     =   "What is the qualification for this test?"
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Qualification"
         Text            =   ""
      End
      Begin MSDBCtls.DBCombo dbcHandling 
         Bindings        =   "IceTests.frx":032B
         DataField       =   "Handling"
         DataSource      =   "dtaTestInfo"
         Height          =   315
         Left            =   1560
         TabIndex        =   15
         ToolTipText     =   "How is the test handled at this event?"
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Label"
         BoundColumn     =   "Code"
         Text            =   ""
      End
      Begin VB.Label lblComments 
         Alignment       =   1  'Right Justify
         Caption         =   "&Comments"
         Height          =   375
         Left            =   360
         TabIndex        =   60
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblGroup 
         Alignment       =   1  'Right Justify
         Caption         =   "&Groups"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label lblJudges 
         Alignment       =   1  'Right Justify
         Caption         =   "&Judges"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label lblSponsor 
         Alignment       =   1  'Right Justify
         Caption         =   "&Sponsor"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblHandling 
         Alignment       =   1  'Right Justify
         Caption         =   "&Handling"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   "How is the test handled at this event?"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         Caption         =   "Status"
         DataField       =   "Label"
         DataSource      =   "dtaProgress"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblQual 
         Alignment       =   1  'Right Justify
         Caption         =   "&Qualification"
         Height          =   375
         Left            =   45
         TabIndex        =   11
         ToolTipText     =   "What is the qualification for this test?"
         Top             =   1320
         Width           =   1335
      End
   End
   Begin MSComctlLib.TabStrip tbsTest 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   7435
      TabWidthStyle   =   2
      MultiRow        =   -1  'True
      TabFixedHeight  =   529
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   1764
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "General"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Preliminary Round"
            Key             =   "prelim"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Finals"
            Key             =   "Finals"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Other information"
            Key             =   "Other"
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
   Begin VB.Data dtaProgress 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7440
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaTestInfo 
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
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaStatusFin 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaStatusPre 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      TabIndex        =   37
      ToolTipText     =   "Add section to test"
      Top             =   7320
      Width           =   5655
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "&OK"
         Height          =   375
         Left            =   4680
         TabIndex        =   35
         ToolTipText     =   "Close window"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   34
         ToolTipText     =   "Remove section from test"
         Top             =   120
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   33
         ToolTipText     =   "Add new section to this test"
         Top             =   120
         Width           =   975
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   36
      Top             =   7890
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Data dtaQual 
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
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data dtaFinals 
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame fraTest 
      Enabled         =   0   'False
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.TextBox txtWR 
         Alignment       =   1  'Right Justify
         DataField       =   "WRTest"
         DataSource      =   "dtaTest"
         Height          =   315
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   7
         Tag             =   "FEIF WorldRanking"
         Top             =   1080
         Width           =   735
      End
      Begin VB.CheckBox chkRein 
         Caption         =   "&Choice of rein"
         DataField       =   "rr"
         DataSource      =   "dtaTest"
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         ToolTipText     =   "Does the rider have a choice of rein?"
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtTest 
         DataField       =   "Test"
         DataSource      =   "dtaTest"
         Height          =   315
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   4
         ToolTipText     =   "What is the name of this test?"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCode 
         DataField       =   "Code"
         DataSource      =   "dtaTest"
         Height          =   315
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   2
         ToolTipText     =   "What is the code for this test?"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblWR 
         Alignment       =   1  'Right Justify
         Caption         =   "FEIF WorldRanking"
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label lblTest 
         Alignment       =   1  'Right Justify
         Caption         =   "&Test"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         Caption         =   "&Code"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Data dtaTest 
      Caption         =   "Tests"
      Connect         =   "Access 2000;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "frmTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple editor to create and edit new tests (needs more options)
'
' Copyright (C) Marko Mazeland 2003, 2006
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
Public fcCode As String
Public fcInitCode As String
Dim fiFormLoaded As Integer
Dim fiIndex As Integer

Private Sub chkFinalsOut_Click(Index As Integer)
    If fiFormLoaded = True Then
        MsgBox Translate("Don't forget to mark how many sections may be taken out.", mcLanguage)
        Check_out
    End If
End Sub
Private Sub chkPrelimOut_Click(Index As Integer)
    If fiFormLoaded = True Then
        MsgBox Translate("Don't forget to mark how many sections may be taken out.", mcLanguage)
        Check_out
    End If
End Sub

Private Sub chkRein_Click()
    If chkRein.Value = 0 Then
        Me.cmbGroup.Enabled = False
    Else
        Me.cmbGroup.Enabled = True
    End If
End Sub

Private Sub cmbFinalsName_Change(Index As Integer)
    SelectRow Index
End Sub

Private Sub cmbPrelimName_Change(Index As Integer)
    SelectRow Index
End Sub

Private Sub cmbprelimName_GotFocus(Index As Integer)
    SelectRow Index
End Sub
Private Sub chkSplitFinals_Click()
    If chkSplitFinals.Value = 0 Then
        cmdSplitFinals.Enabled = False
    Else
        cmdSplitFinals.Enabled = True
    End If
    With dtaTestInfo.Recordset
        .Edit
        .Fields("SplitFinals") = chkSplitFinals.Value
        .Update
    End With
End Sub
Private Sub cmbFinalsName_Click(Index As Integer)
    SelectRow Index
End Sub

Private Sub cmbMarkDecimals_Change(Index As Integer)
    If Val(cmbMarkDecimals(Index)) > 2 Or Val(cmbMarkDecimals(Index).Text) < 0 Then
        cmbMarkDecimals(Index).Text = 1
    End If
    If Index = 1 Then
        cmbMarkDecimals(0).Text = cmbMarkDecimals(1).Text
    Else
        cmbMarkDecimals(1).Text = cmbMarkDecimals(0).Text
    End If
End Sub


Private Sub cmbMarkDecimals_Click(Index As Integer)
    If Val(cmbMarkDecimals(Index)) > 2 Or Val(cmbMarkDecimals(Index).Text) < 0 Then
        cmbMarkDecimals(Index).Text = 1
    End If
    If Index = 1 Then
        cmbMarkDecimals(0).Text = cmbMarkDecimals(1).Text
    Else
        cmbMarkDecimals(1).Text = cmbMarkDecimals(0).Text
    End If
End Sub

Private Sub cmbTimeDecimals_Change(Index As Integer)
    If Val(cmbTimeDecimals(Index).Text) > 2 Or Val(cmbTimeDecimals(Index).Text) < 0 Then
        cmbTimeDecimals(Index).Text = 2
    End If
    If Index = 1 Then
        cmbTimeDecimals(0).Text = cmbTimeDecimals(1).Text
    Else
        cmbTimeDecimals(1).Text = cmbTimeDecimals(0).Text
    End If
End Sub

Private Sub cmbTimeDecimals_Click(Index As Integer)
    If Val(cmbTimeDecimals(Index).Text) > 2 Or Val(cmbTimeDecimals(Index).Text) < 0 Then
        cmbTimeDecimals(Index).Text = 2
    End If
    If Index = 1 Then
        cmbTimeDecimals(0).Text = cmbTimeDecimals(1).Text
    Else
        cmbTimeDecimals(1).Text = cmbTimeDecimals(0).Text
    End If
End Sub

Private Sub cmdNewQual_Click()
    Dim cTemp As String
    Dim iKey As Integer
    
    cTemp = Trim$(InputBox(Translate("Create a new qualification class?", mcLanguage), "", dbcQual.Text))
    If cTemp <> "" And cTemp <> Chr$(27) Then
        iKey = MsgBox(Translate("Add a new qualification class:", mcLanguage) & " " & cTemp, vbYesNo + vbQuestion)
        If iKey = vbYes Then
            dbcQual.Text = Left$(cTemp, Me.dtaTest.Recordset.Fields("Qualification").Size)
        End If
    End If
End Sub

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmbprelimName_Click(Index As Integer)
    SelectRow Index
End Sub

Private Sub cmdSplitFinals_Click()
    frmSplitFinals.fiTakeAllClasses = True
    frmSplitFinals.fcTestCode = fcInitCode
    frmSplitFinals.Show 1, Me
End Sub

Private Sub dbcFin_KeyUp(KeyCode As Integer, Shift As Integer)
    CheckDecimals
End Sub

Private Sub dbcFin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CheckDecimals
End Sub


Private Sub dbcHandling_Click(Area As Integer)
    If dbcHandling.BoundText = 0 Then
        chkSplitFinals.Enabled = False
        cmdSplitFinals.Enabled = False
    Else
        If Me.dtaTestInfo.Recordset.Fields("Handling") > 4 Then
            chkSplitFinals.Enabled = False
        Else
            chkSplitFinals.Enabled = True
            If chkSplitFinals.Value = 1 Then
                cmdSplitFinals.Enabled = True
            End If
        End If
    End If
End Sub

Private Sub dbcOther_Change(Index As Integer)
    Dim iTemp As Integer
    If Index = 1 And dbcOther(1).BoundText = "2" Then
        dbcOther(2).BoundText = "3"
        chkSplitFinals.Enabled = False
        cmdSplitFinals.Enabled = False
        dbcHandling.Enabled = False
    ElseIf Index = 2 And dbcOther(2).BoundText = "3" Then
        dbcOther(1).BoundText = "2"
        chkSplitFinals.Enabled = False
        cmdSplitFinals.Enabled = False
        dbcHandling.Enabled = False
    ElseIf dbcOther(2).BoundText <> "0" Then
        chkSplitFinals.Enabled = False
        cmdSplitFinals.Enabled = False
        dbcHandling.Enabled = False
    Else
        chkSplitFinals.Enabled = True
        cmdSplitFinals.Enabled = True
        dbcHandling.Enabled = True
    End If
    For iTemp = 0 To txtOther.Count - 1
        If dbcOther(2).BoundText > "0" Then
            lblOther(iTemp + dbcOther.Count).Visible = True
            txtOther(iTemp).Visible = True
            If dtaTimeTables.Recordset.RecordCount = 0 Then
                With dtaTimeTables.Recordset
                    .AddNew
                    .Fields("Code") = fcInitCode
                    .Fields("ScaleRange") = 10
                    .Update
                End With
                dtaTimeTables.Refresh
            End If
        Else
            lblOther(iTemp + dbcOther.Count).Visible = False
            txtOther(iTemp).Visible = False
        End If
    Next iTemp

End Sub

Private Sub dbcPre_Change()
    If dbcPre.BoundText = "3" Then
        cmdAdd.Enabled = True
    Else
        cmdAdd.Enabled = False
    End If
End Sub

Private Sub dbcPre_KeyUp(KeyCode As Integer, Shift As Integer)
    CheckDecimals
End Sub

Private Sub dbcPre_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CheckDecimals
End Sub

Private Sub Form_Load()
    Dim rst As DAO.Recordset
    Dim iTemp As Integer
        
    fiFormLoaded = False
        
    ReadFormPosition Me
    ChangeFontSize Me, msFontSize
    TranslateControls Me
       
    For iTemp = 1 To tbsTest.Tabs.Count
        tbsTest.Tabs(iTemp).Caption = TranslateCaption(tbsTest.Tabs(iTemp).Caption, 0, True)
    Next iTemp
    
    With Me.cmbGroup
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "4"
        .AddItem "5"
    End With
    
    With Me.cmbJudges
        .Clear
        .AddItem "1"
        .AddItem "2"
        .AddItem "3"
        .AddItem "5"
    End With
    
    For iTemp = 0 To 1
        With Me.cmbMarkDecimals(iTemp)
            .Clear
            .AddItem "0"
            .AddItem "1"
            .AddItem "2"
        End With
    
        With Me.cmbTimeDecimals(iTemp)
            .Clear
            .AddItem "0"
            .AddItem "1"
            .AddItem "2"
        End With
    Next iTemp
    DoEvents
    
    If TableExist(mdbMain, "TestInfo") = False Then
        MsgBox Translate("Please download a new set of Sport Rules first!", mcLanguage), vbCritical
        Unload Me
    End If
    
    mdbMain.Execute ("UPDATE TestInfo SET SplitFinals=0 WHERE IsNull(SplitFinals)<>False")
    mdbMain.Execute ("UPDATE TestInfo SET SplitFinals=0 WHERE Handling>4")
 
    If TableExist(mdbMain, "[Values]") = True Then
        Set rst = mdbMain.OpenRecordset("SELECT * FROM [Values] WHERE [Field]='Status'")
        If rst.RecordCount = 0 Then
            With rst
                .AddNew
                .Fields("Field") = "Status"
                .Fields("Code") = "1"
                .Fields("Label") = "Marks"
                .Update
                
                .AddNew
                .Fields("Field") = "Status"
                .Fields("Code") = "2"
                .Fields("Label") = "Place Marks"
                .Update
                
                .AddNew
                .Fields("Field") = "Status"
                .Fields("Code") = "3"
                .Fields("Label") = "Time"
                .Update
            End With
        End If
        rst.Close
    Else
        MsgBox Translate("Please download a new set of Sport Rules first!", mcLanguage), vbCritical
        Unload Me
    End If
    
    With dtaTimeTables
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT * FROM TestTimeTables WHERE Code LIKE '" & fcInitCode & "'"
        .Refresh
    End With
    
    With dtaTest
        .Caption = "Tests"
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT Code,Test,RR,Qualification,Userlevel,Type_Pre,Type_Final,Mark_decimals,Time_decimals,WRTest,Num_J,Groupsize,Out_Fin,Type_Special,Type_Time,Comments FROM Tests WHERE Tests.Code LIKE '" & fcInitCode & "'"
        .Refresh
    End With
    
    With dtaTestInfo
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT * FROM TestInfo WHERE Code LIKE '" & fcInitCode & "'"
        .Refresh
    End With
    chkSplitFinals.Value = dtaTestInfo.Recordset.Fields("Splitfinals")
    
    With dtaHandling
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT [Code],[Label] FROM [Values] WHERE [Field]='Handling' ORDER BY [Label]"
        .Refresh
    End With
    
    lblStatus.Caption = ""
    Set rst = mdbMain.OpenRecordset("SELECT [Label] FROM [Values] WHERE [Field]='Status' and Code='" & dtaTestInfo.Recordset.Fields("Status") & "'")
    If rst.RecordCount > 0 Then
        lblStatus.Caption = Translate("Status", mcLanguage) & ": " & Translate(rst.Fields(0), mcLanguage)
    End If
    Set rst = mdbMain.OpenRecordset("SELECT COUNT(STA) FROM Entries WHERE Code LIKE '" & fcInitCode & "' AND Status=0")
    If rst.RecordCount > 0 Then
        If lblStatus.Caption = "" Then
            lblStatus.Caption = Translate("Status", mcLanguage) & ": "
        End If
        lblStatus.Caption = lblStatus.Caption & "; " & rst.Fields(0) & " " & Translate("participants", mcLanguage)
    End If
    rst.Close
    
    If dtaTestInfo.Recordset.Fields("Status") > 1 Then
        dbcHandling.Enabled = False
    End If
    
    With dtaQual
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT DISTINCT Qualification FROM Tests ORDER BY Qualification"
        .Refresh
    End With
    
    With dtaStatusPre
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT [Code],[Label] FROM [Values] WHERE [Field]='Type_PF' ORDER BY [Code]"
        .Refresh
    End With
    
    With dtaStatusFin
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT [Code],[Label] FROM [Values] WHERE [Field]='Type_PF' ORDER BY [Code]"
        .Refresh
    End With
    
    With dtaType_Special
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT [Code],[Label] FROM [Values] WHERE [Field]='Type_Special' ORDER BY [Code]"
        .Refresh
    End With
    
    With dtaType_Time
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT [Code],[Label] FROM [Values] WHERE [Field]='Type_Time' ORDER BY [Code]"
        .Refresh
    End With
    
    LoadRecord
        
    Form_Resize
    
    CheckDecimals
        
    Check_out
    
    tbsTest.SelectedItem = tbsTest.Tabs.Item(2)
    tbsTest.SelectedItem = tbsTest.Tabs.Item(1)
    tbsTest.SelectedItem.HighLighted = miUseHighLights
    tbsTest_Click
    
    Set rst = Nothing
            
    fiFormLoaded = True
End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    
    On Local Error Resume Next
    
    With fraTest
        .Top = 50
        .Left = 50
        .Width = ScaleWidth - 100
        .Height = (txtCode.Height + 50) * 4
    End With
    
    With lblCode
        .Container = fraTest
        .Left = 50
        .Top = 250
        .Width = .Container.Width \ 6 - 100
    End With
    
    With txtCode
        .Container = fraTest
        .Top = lblCode.Top
        .Left = lblCode.Left + lblCode.Width + 50
        .Width = .Container.Width - .Left - 50
    End With
    
    With lblTest
        .Container = fraTest
        .Top = txtCode.Top + txtCode.Height + 50
        .Left = lblCode.Left
        .Width = lblCode.Width
    End With
    
    With txtTest
        .Container = fraTest
        .Top = lblTest.Top
        .Left = txtCode.Left
        .Width = .Container.Width - .Left - 50
    End With
    
    With chkRein
        .Container = fraTest
        .Top = txtTest.Top + txtTest.Height + 50
        .Left = txtTest.Left
        .Width = (.Container.Width - .Left - 50) \ 2
    End With
    
    With txtWR
        .Container = fraTest
        .Top = chkRein.Top
        .Left = .Container.Width - .Width - 50
    End With
    
    With lblWR
        .Container = fraTest
        .Top = chkRein.Top
        .Left = chkRein.Left + chkRein.Width + 50
        .Width = txtWR.Left - .Left - 50
    End With
    
    With tbsTest
        .Left = 50
        .Top = fraTest.Top + fraTest.Height + 50
        .Width = ScaleWidth - 100
        .Height = ScaleHeight - .Top - 100 - fraButtons.Height - Me.StatusBar1.Height
        .TabFixedWidth = (.Width - 100) \ (.Tabs.Count + 1)
        .TabMinWidth = frmMain.miTabMinWidth
        .TabFixedHeight = frmMain.miTabMinHeight
    End With
    
    For iTemp = 1 To fraTbsTest.Count
        With fraTbsTest(iTemp)
            .Top = Me.tbsTest.ClientTop
            .Left = Me.tbsTest.ClientLeft
            .Height = Me.tbsTest.ClientHeight
            .Width = Me.tbsTest.ClientWidth
        End With
    Next iTemp
            
    '*frame general
    With lblStatus
        .Container = fraTbsTest(1)
        .Top = 250
        .Left = 50 + lblCode.Width
        .Width = .Container.Width - 100
    End With
    
    With lblSponsor
        .Container = fraTbsTest(1)
        .Top = lblStatus.Top + lblStatus.Height + 50
        .Width = lblCode.Width
        .Left = 50
    End With
    
    With txtSponsor
        .Container = fraTbsTest(1)
        .Top = lblSponsor.Top
        .Left = lblSponsor.Left + lblSponsor.Width + 50
        .Width = .Container.Width - .Left - cmdNewQual.Width - 150
    End With
    
    With lblQual
        .Container = fraTbsTest(1)
        .Top = txtSponsor.Top + txtSponsor.Height + 50
        .Width = lblSponsor.Width
        .Left = 50
    End With
    
    With dbcQual
        .Container = fraTbsTest(1)
        .Top = lblQual.Top
        .Left = lblQual.Left + lblQual.Width + 50
        .Width = .Container.Width - .Left - cmdNewQual.Width - 150
    End With
    
    With cmdNewQual
        .Container = fraTbsTest(1)
        .Top = lblQual.Top
        .Left = .Container.Width - .Width - 50
    End With
    
    With lblHandling
        .Left = lblQual.Left
        .Top = dbcQual.Top + dbcQual.Height + 50
        .Width = lblQual.Width
    End With
    
    With dbcHandling
        Set .Container = fraTbsTest(1)
        .Left = dbcQual.Left
        .Top = lblHandling.Top
        .Width = dbcQual.Width
    End With
    
    With lblJudges
        .Container = fraTbsTest(1)
        .Left = lblHandling.Left
        .Top = dbcHandling.Top + dbcHandling.Height + 50
        .Width = lblHandling.Width
    End With
    
    With cmbJudges
        .Container = fraTbsTest(1)
        .Left = dbcHandling.Left
        .Top = lblJudges.Top
        .Width = .Container.Width \ 10
    End With
        
    With lblGroup
        .Container = fraTbsTest(1)
        .Left = lblJudges.Left
        .Top = cmbJudges.Top + cmbJudges.Height + 50
        .Width = lblHandling.Width
    End With
    
    With cmbGroup
        .Container = fraTbsTest(1)
        .Left = cmbJudges.Left
        .Top = lblGroup.Top
        .Width = .Container.Width \ 10
    End With
        
    With chkSplitFinals
        .Top = cmbGroup.Top + cmbGroup.Height + 50
        .Left = cmbGroup.Left
        .Width = dbcQual.Width
        .Height = cmbGroup.Height
    End With
    
    With cmdSplitFinals
        .Container = fraTbsTest(1)
        .Top = chkSplitFinals.Top
        .Left = .Container.Width - .Width - 50
    End With
    
    With lblComments
        .Container = fraTbsTest(1)
        .Left = lblJudges.Left
        .Top = chkSplitFinals.Top + chkSplitFinals.Height + 50
        .Width = lblJudges.Width
    End With
    
    With txtComments
        .Container = fraTbsTest(1)
        .Top = lblComments.Top
        .Left = lblComments.Left + lblComments.Width + 50
        .Width = .Container.Width - .Left - cmdNewQual.Width - 150
        .Height = Max(500, .Container.Height - .Top - 50)
    End With
    
    ' frame for prelim
    With lblPre
        .Container = fraTbsTest(2)
        .Top = 250
        .Left = 50
        .Width = .Container.Width \ 6 - 100
        .Visible = True
    End With
    
    With dbcPre
        .Container = fraTbsTest(2)
        .Top = lblPre.Top
        .Left = lblPre.Left + lblPre.Width + 50
        .Width = .Container.Width \ 2 - .Left - 50
        .Visible = True
    End With
    
    With cmbTimeDecimals(0)
        .Container = fraTbsTest(2)
        .Top = lblPre.Top
        .Left = .Container.Width * 0.7 - 150
        .Width = .Container.Width \ 10
    End With
    
    With cmbMarkDecimals(0)
        .Container = fraTbsTest(2)
        .Top = lblPre.Top
        .Left = .Container.Width * 0.7 - 150
        .Width = .Container.Width \ 10
    End With
    
    With lblDec(0)
        .Container = fraTbsTest(2)
        .Top = lblPre.Top
        .Width = .Container.Width \ 6 - 100
        .Left = cmbTimeDecimals(0).Left - .Width - 50
    End With
    
    With lblPrelimSection
        .Container = fraTbsTest(2)
        .Top = dbcPre.Top + dbcPre.Height + 50
        .Left = 50
        .Width = .Container.Width * 0.5 - 200
    End With
    
    With lblPrelimFactor
        .Container = fraTbsTest(2)
        .Top = lblPrelimSection.Top
        .Left = lblPrelimSection.Left + lblPrelimSection.Width + 50
        .Width = .Container.Width \ 10
    End With
    
    With lblPrelimLo
        .Container = fraTbsTest(2)
        .Top = lblPrelimFactor.Top
        .Left = lblPrelimFactor.Left + lblPrelimFactor.Width + 50
        .Width = .Container.Width \ 10
    End With
    
    With lblPrelimHi
        .Container = fraTbsTest(2)
        .Top = lblPrelimLo.Top
        .Left = lblPrelimLo.Left + lblPrelimHi.Width + 50
        .Width = .Container.Width \ 10
    End With
    
    With lblPrelimOut
        .Container = fraTbsTest(2)
        .Top = lblPrelimHi.Top
        .Left = lblPrelimHi.Left + lblPrelimHi.Width + 50
        .Width = .Container.Width \ 10
    End With
    
    For iTemp = 1 To cmbPrelimName.Count - 1
        With cmbPrelimName(iTemp)
            .Container = fraTbsTest(2)
            .Top = lblPrelimSection.Top + lblPrelimSection.Height + cmbPrelimName(0).Height * (iTemp - 1) + 50
            .Width = .Container.Width * 0.5 - 200
            .Left = 50
        End With
        With txtPrelimFactor(iTemp)
            .Container = fraTbsTest(2)
            .Top = cmbPrelimName(iTemp).Top
            .Height = cmbPrelimName(0).Height
            .Left = cmbPrelimName(iTemp).Left + cmbPrelimName(iTemp).Width + 50
            .Width = .Container.Width \ 20
        End With
        With txtPrelimLo(iTemp)
            .Container = fraTbsTest(2)
            .Top = txtPrelimFactor(iTemp).Top
            .Height = cmbPrelimName(0).Height
            .Width = .Container.Width \ 20
            .Left = lblPrelimLo.Left
        End With
        With txtPrelimHi(iTemp)
            .Container = fraTbsTest(2)
            .Top = txtPrelimLo(iTemp).Top
            .Height = cmbPrelimName(0).Height
            .Width = .Container.Width \ 20
            .Left = lblPrelimHi.Left
        End With
        With chkPrelimOut(iTemp)
            .Container = fraTbsTest(2)
            .Top = cmbPrelimName(iTemp).Top
            .Height = cmbPrelimName(0).Height
            .Width = .Container.Width \ 20
            .Left = lblPrelimOut.Left
        End With
    Next iTemp
    
    ' frame for finals
    With lblFin
        .Container = fraTbsTest(3)
        .Top = 250
        .Left = 50
        .Width = .Container.Width \ 6 - 100
        .Visible = True
    End With
    
    With dbcFin
        .Container = fraTbsTest(3)
        .Top = lblFin.Top
        .Left = lblFin.Left + lblFin.Width + 50
        .Width = .Container.Width \ 2 - .Left - 50
        .Visible = True
    End With
    
    With cmbTimeDecimals(1)
        .Container = fraTbsTest(3)
        .Top = lblFin.Top
        .Left = .Container.Width * 0.7 - 150
        .Width = .Container.Width \ 10
    End With
    
    With cmbMarkDecimals(1)
        .Container = fraTbsTest(3)
        .Top = lblFin.Top
        .Left = .Container.Width * 0.7 - 150
        .Width = .Container.Width \ 10
    End With
    
    With lblDec(1)
        .Container = fraTbsTest(3)
        .Top = lblFin.Top
        .Width = .Container.Width \ 6 - 100
        .Left = cmbTimeDecimals(1).Left - .Width - 50
    End With
    
    With lblFinalsSection
        .Container = fraTbsTest(3)
        .Top = dbcFin.Top + dbcFin.Height + 50
        .Left = 50
        .Width = .Container.Width * 0.5 - 200
    End With
    
    With lblFinalsFactor
        .Container = fraTbsTest(3)
        .Top = lblFinalsSection.Top
        .Left = lblFinalsSection.Left + lblFinalsSection.Width + 50
        .Width = .Container.Width \ 10
    End With
        
    With lblFinalsLo
        .Container = fraTbsTest(2)
        .Top = lblFinalsFactor.Top
        .Left = lblFinalsFactor.Left + lblFinalsFactor.Width + 50
        .Width = .Container.Width \ 10
    End With
    
    With lblFinalsHi
        .Container = fraTbsTest(2)
        .Top = lblFinalsLo.Top
        .Left = lblFinalsLo.Left + lblFinalsHi.Width + 50
        .Width = .Container.Width \ 10
    End With
    
    With lblFinalsOut
        .Container = fraTbsTest(2)
        .Top = lblFinalsHi.Top
        .Left = lblFinalsHi.Left + lblFinalsHi.Width + 50
        .Width = .Container.Width \ 10
    End With
    
    With lblRecycle
        .Container = fraTbsTest(3)
        .Top = lblFinalsOut.Top
        .Left = lblFinalsOut.Left + lblFinalsOut.Width + 50
        .Width = .Container.Width \ 10
    End With
    
    For iTemp = 1 To cmbFinalsName.Count - 1
        With cmbFinalsName(iTemp)
            .Container = fraTbsTest(3)
            .Top = lblFinalsSection.Top + lblFinalsSection.Height + cmbFinalsName(0).Height * (iTemp - 1) + 50
            .Width = .Container.Width * 0.5 - 200
            .Left = 50
        End With
        With txtFinalsFactor(iTemp)
            .Container = fraTbsTest(3)
            .Top = cmbFinalsName(iTemp).Top
            .Height = cmbFinalsName(0).Height
            .Width = .Container.Width \ 20
            .Left = cmbFinalsName(iTemp).Left + cmbFinalsName(iTemp).Width + 50
        End With
        With txtFinalsLo(iTemp)
            .Container = fraTbsTest(2)
            .Top = txtFinalsFactor(iTemp).Top
            .Height = cmbFinalsName(0).Height
            .Width = .Container.Width \ 20
            .Left = lblFinalsLo.Left
        End With
        With txtFinalsHi(iTemp)
            .Container = fraTbsTest(2)
            .Top = txtFinalsLo(iTemp).Top
            .Height = cmbFinalsName(0).Height
            .Width = .Container.Width \ 20
            .Left = lblFinalsHi.Left
        End With
        With chkFinalsOut(iTemp)
            .Container = fraTbsTest(2)
            .Top = cmbFinalsName(iTemp).Top
            .Height = cmbFinalsName(0).Height
            .Width = .Container.Width \ 20
            .Left = lblFinalsOut.Left
        End With
        
        With chkRecycle(iTemp)
            .Container = fraTbsTest(3)
            .Top = chkFinalsOut(iTemp).Top
            .Height = cmbFinalsName(0).Height
            .Width = .Container.Width \ 20
            .Left = lblRecycle.Left
        End With
    Next iTemp
    
    '* frame other information
    With lblGeneralText
        .Left = 50
        .Height = dbcOther(0).Height
        .Top = 250
    End With
    
    For iTemp = 0 To lblOther.Count - 1
        With lblOther(iTemp)
            .Container = fraTbsTest(3)
            If iTemp = 0 Then
                .Top = lblGeneralText.Top + lblGeneralText.Height + 50
            Else
                .Top = lblOther(iTemp - 1).Top + dbcOther(0).Height + 50
            End If
            .Left = 50
            .Width = .Container.Width * 0.5 - 200
        End With
    Next iTemp
    For iTemp = 0 To dbcOther.Count - 1
        With dbcOther(iTemp)
            .Container = fraTbsTest(3)
            .Top = lblOther(iTemp).Top
            .Left = lblOther(iTemp).Left + lblOther(iTemp).Width + 50
            .Width = .Container.Width - .Left - 50
        End With
    Next iTemp
    For iTemp = 0 To txtOther.Count - 1
        With txtOther(iTemp)
            .Container = fraTbsTest(3)
            .Top = lblOther(iTemp + dbcOther.Count).Top
            .Left = lblOther(iTemp).Left + lblOther(iTemp).Width + 50
            .Width = .Container.Width - .Left - 50
        End With
    Next iTemp
    
    '*frame for buttons
    With fraButtons
        .Left = 50
        .Width = ScaleWidth - 100
        .Height = cmdAdd.Height + 50
        .Top = Me.StatusBar1.Top - .Height - 50
    End With
      
    With cmdOK
        .Container = fraButtons
        .Top = 0
        .Left = .Container.Width - cmdOK.Width - 50
    End With
    
    With cmdDelete
        .Container = fraButtons
        .Top = 0
        .Left = cmdOK.Left - cmdDelete.Width - 50
    End With
    
    With cmdAdd
        .Container = fraButtons
        .Top = 0
        .Left = cmdDelete.Left - 50 - cmdAdd.Width
    End With
    On Local Error GoTo 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    SaveRecord
    
    WriteFormPosition Me

End Sub

Private Sub txtFactor_Click(Index As Integer)
    SelectRow Index
End Sub


Private Sub SaveRecord()
    Dim iTemp As Integer
    Dim rstTestSections As DAO.Recordset
    
    On Local Error Resume Next
    
    If fraTest.Enabled = True Then
        mdbMain.Execute "DELETE * FROM Testsections WHERE Code='" & fcCode & "'"
        Set rstTestSections = mdbMain.OpenRecordset("SELECT * FROM Testsections")
        With rstTestSections
            For iTemp = 1 To cmbPrelimName.Count - 1
                If cmbPrelimName(iTemp).Tag = "" Then
                    .AddNew
                    .Fields("Code") = fcCode
                    .Fields("Status") = 0
                    .Fields("Section") = iTemp
                    .Fields("Name") = Left$(cmbPrelimName(iTemp).Text, .Fields("Name").Size)
                    .Fields("Factor") = Val(txtPrelimFactor(iTemp).Text)
                    .Fields("Mark_low") = Val(txtPrelimLo(iTemp).Text)
                    .Fields("Mark_hi") = Val(txtPrelimHi(iTemp).Text)
                    .Fields("Out") = IIf(chkPrelimOut(iTemp) = 1, True, False)
                    .Fields("Recycle") = False
                    .Update
                End If
            Next iTemp
            For iTemp = 1 To cmbFinalsName.Count - 1
                If cmbFinalsName(iTemp).Tag = "" Then
                    .AddNew
                    .Fields("Code") = fcCode
                    .Fields("Status") = 1
                    .Fields("Section") = iTemp
                    .Fields("Name") = Left$(cmbFinalsName(iTemp).Text, .Fields("Name").Size)
                    .Fields("Mark_low") = Val(txtFinalsLo(iTemp).Text)
                    .Fields("Mark_hi") = Val(txtFinalsHi(iTemp).Text)
                    .Fields("Factor") = Val(txtFinalsFactor(iTemp).Text)
                    .Fields("Out") = IIf(chkFinalsOut(iTemp) = 1, True, False)
                    .Fields("Recycle") = IIf(chkRecycle(iTemp) = 1, True, False)
                    .Update
                End If
            Next iTemp
        End With
        rstTestSections.Close
        Set rstTestSections = Nothing
    End If
    fiIndex = -1
    
    On Local Error GoTo 0
End Sub

Sub LoadRecord()
    
    Dim iTemp As Integer
    
    On Local Error Resume Next
    
    Caption = dtaTest.Recordset.Fields("Code") & " - " & dtaTest.Recordset.Fields("Test")
    If dtaTest.Recordset.Fields("Userlevel") = 1 Then
        fraTest.Enabled = True
    Else
        fraTest.Enabled = False
        Caption = Caption & " [" & Translate("Read Only", mcLanguage) & "/FIPO]"
    End If
        
    For iTemp = 2 To 4
        Me.fraTbsTest(iTemp).Enabled = fraTest.Enabled
    Next iTemp
    
    With dtaName
        .DatabaseName = mcDatabaseName
        .RecordSource = "SELECT DISTINCT Name FROM TestSections WHERE Name<>'' ORDER BY Name"
        .Refresh
    End With
    
    If dtaTest.Recordset.Fields("Code") <> "" Then
        fcCode = dtaTest.Recordset.Fields("Code")
        
        dtaPrelim.DatabaseName = mcDatabaseName
        dtaPrelim.RecordSource = "SELECT * FROM TestSections WHERE Code LIKE '" & fcCode & "' AND Status=0 ORDER BY Section"
        dtaPrelim.Refresh
        If dtaPrelim.Recordset.RecordCount > 0 Then
            Do While Not dtaPrelim.Recordset.EOF
                Load cmbPrelimName(cmbPrelimName.Count)
                Load txtPrelimFactor(txtPrelimFactor.Count)
                Load txtPrelimLo(txtPrelimLo.Count)
                Load txtPrelimHi(txtPrelimHi.Count)
                Load chkPrelimOut(chkPrelimOut.Count)
                With cmbPrelimName(dtaPrelim.Recordset.AbsolutePosition + 1)
                    .Visible = True
                    .Clear
                    .Tag = ""
                    With dtaName.Recordset
                        .MoveFirst
                        Do While Not .EOF
                            cmbPrelimName(dtaPrelim.Recordset.AbsolutePosition + 1).AddItem .Fields(0)
                            .MoveNext
                        Loop
                    End With
                    DoEvents
                    .Text = dtaPrelim.Recordset.Fields("Name")
                End With
                With txtPrelimFactor(dtaPrelim.Recordset.AbsolutePosition + 1)
                    .Visible = True
                    .Text = dtaPrelim.Recordset.Fields("Factor")
                End With
                With txtPrelimLo(dtaPrelim.Recordset.AbsolutePosition + 1)
                    .Visible = True
                    .Text = dtaPrelim.Recordset.Fields("Mark_Low")
                End With
                With txtPrelimHi(dtaPrelim.Recordset.AbsolutePosition + 1)
                    .Visible = True
                    .Text = dtaPrelim.Recordset.Fields("Mark_Hi")
                End With
                With chkPrelimOut(dtaPrelim.Recordset.AbsolutePosition + 1)
                    If dtaPrelim.Recordset.RecordCount > 1 Then
                        .Visible = True
                    Else
                        lblPrelimOut.Visible = False
                    End If
                    .Value = IIf(dtaPrelim.Recordset.Fields("Out") = True, 1, 0)
                End With
                dtaPrelim.Recordset.MoveNext
            Loop
        End If
        
        dtaFinals.DatabaseName = mcDatabaseName
        dtaFinals.RecordSource = "SELECT * FROM TestSections WHERE Code LIKE '" & fcCode & "' AND Status=1 ORDER BY Section"
        dtaFinals.Refresh
        If dtaFinals.Recordset.RecordCount > 0 Then
            dtaFinals.Recordset.MoveLast
            dtaFinals.Recordset.MoveFirst
            Do While Not dtaFinals.Recordset.EOF
                Load cmbFinalsName(cmbFinalsName.Count)
                Load txtFinalsFactor(txtFinalsFactor.Count)
                Load txtFinalsLo(txtFinalsLo.Count)
                Load txtFinalsHi(txtFinalsHi.Count)
                Load chkFinalsOut(chkFinalsOut.Count)
                Load chkRecycle(chkRecycle.Count)
                With cmbFinalsName(dtaFinals.Recordset.AbsolutePosition + 1)
                    .Visible = True
                    .Clear
                    .Tag = ""
                    With dtaName.Recordset
                        .MoveFirst
                        Do While Not .EOF
                            cmbFinalsName(dtaFinals.Recordset.AbsolutePosition + 1).AddItem .Fields(0)
                            .MoveNext
                        Loop
                    End With
                    DoEvents
                    .Text = dtaFinals.Recordset.Fields("Name")
                End With
                With txtFinalsFactor(dtaFinals.Recordset.AbsolutePosition + 1)
                    .Visible = True
                    .Text = dtaFinals.Recordset.Fields("Factor")
                End With
                With txtFinalsLo(dtaFinals.Recordset.AbsolutePosition + 1)
                    .Visible = True
                    .Text = dtaFinals.Recordset.Fields("Mark_Low")
                End With
                With txtFinalsHi(dtaFinals.Recordset.AbsolutePosition + 1)
                    .Visible = True
                    .Text = dtaFinals.Recordset.Fields("Mark_Hi")
                End With
                With chkFinalsOut(dtaFinals.Recordset.AbsolutePosition + 1)
                    If dtaFinals.Recordset.RecordCount > 1 Then
                        .Visible = True
                    Else
                        lblFinalsOut.Visible = False
                    End If
                    .Value = IIf(dtaFinals.Recordset.Fields("Out") = True, 1, 0)
                End With
                dtaFinals.Recordset.MoveNext
            Loop
        End If
        Form_Resize
    End If
    fiIndex = -1
End Sub
Sub SelectRow(Index As Integer)
    Dim iTemp As Integer
    If tbsTest.SelectedItem.Index = 2 Then
        For iTemp = 1 To txtPrelimFactor.Count - 1
            If iTemp = Index Then
                fiIndex = Index
                Me.txtPrelimFactor(iTemp).FontBold = True
                Me.txtPrelimLo(iTemp).FontBold = True
                Me.txtPrelimHi(iTemp).FontBold = True
            Else
                Me.txtPrelimFactor(iTemp).FontBold = False
                Me.txtPrelimLo(iTemp).FontBold = False
                Me.txtPrelimHi(iTemp).FontBold = False
                With Me.cmbPrelimName(iTemp)
                    .SelStart = 0
                End With
            End If
        Next iTemp
    Else
        For iTemp = 1 To txtFinalsFactor.Count - 1
            If iTemp = Index Then
                fiIndex = Index
                Me.txtFinalsFactor(iTemp).FontBold = True
                Me.txtFinalsLo(iTemp).FontBold = True
                Me.txtFinalsHi(iTemp).FontBold = True
            Else
                Me.txtFinalsFactor(iTemp).FontBold = False
                Me.txtFinalsLo(iTemp).FontBold = False
                Me.txtFinalsHi(iTemp).FontBold = False
                With Me.cmbFinalsName(iTemp)
                    .SelStart = 0
                End With
            End If
        Next iTemp
    End If
End Sub


Private Sub txtSection_Click(Index As Integer)
    SelectRow Index
End Sub
Private Sub cmdDelete_Click()
    Dim iLast As Integer
    Dim iKey As Integer
    Dim iVisible As Integer
    Dim iTemp As Integer
    
    If fiIndex >= 0 Then
        If tbsTest.SelectedItem.Index = 2 Then
            iKey = MsgBox(cmbPrelimName(fiIndex).Text & ": " & Translate("Remove this section from this test?", mcLanguage), vbQuestion + vbYesNo)
            If iKey = vbYes Then
                If cmbPrelimName.Count > 1 Then
                For iTemp = 0 To cmbPrelimName.Count - 1
                        If cmbPrelimName(iTemp).Visible = True Then
                            iVisible = iVisible + 1
                        End If
                    Next iTemp
                    If iVisible > 1 Then
                        cmbPrelimName(fiIndex).Visible = False
                        txtPrelimFactor(fiIndex).Visible = False
                        txtPrelimLo(fiIndex).Visible = False
                        txtPrelimHi(fiIndex).Visible = False
                        Me.chkPrelimOut(fiIndex).Visible = False
                        cmbPrelimName(fiIndex).Tag = "Deleted"
                        SaveRecord
                    Else
                        MsgBox Translate("You cannot remove the last section of a test.", mcLanguage)
                    End If
                Else
                    MsgBox Translate("You cannot remove the last section of a test.", mcLanguage)
                End If
            End If
        Else
            iKey = MsgBox(cmbFinalsName(fiIndex).Text & ": " & Translate("Remove this section from this test?", mcLanguage), vbQuestion + vbYesNo)
            If iKey = vbYes Then
                If cmbFinalsName.Count > 1 Then
                    cmbFinalsName(fiIndex).Visible = False
                    Me.txtFinalsFactor(fiIndex).Visible = False
                    txtFinalsLo(fiIndex).Visible = False
                    txtFinalsHi(fiIndex).Visible = False
                    Me.chkFinalsOut(fiIndex).Visible = False
                    Me.chkRecycle(fiIndex).Visible = False
                    cmbFinalsName(fiIndex).Tag = "Deleted"
                    SaveRecord
                Else
                    MsgBox Translate("You cannot remove the last section of a test.", mcLanguage)
                End If
            End If
        End If
    Else
        MsgBox Translate("Select a section to remove first!", mcLanguage)
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim iTemp As Integer
    Dim cTemp As String
    Dim iKey As Integer
    
    On Local Error Resume Next
        
    If tbsTest.SelectedItem.Index = 2 Then
        If cmbPrelimName.Count = 1 Then
            Load cmbPrelimName(cmbPrelimName.Count)
            Load txtPrelimFactor(txtPrelimFactor.Count)
            Load txtPrelimLo(txtPrelimLo.Count)
            Load txtPrelimHi(txtPrelimHi.Count)
            Load chkPrelimOut(chkPrelimOut.Count)
        End If
        For iTemp = 1 To cmbPrelimName.Count - 1
            If chkPrelimOut(iTemp).Visible = False Then
                lblPrelimOut.Visible = True
                chkPrelimOut(iTemp).Visible = True
            End If
            If cmbPrelimName(iTemp).Visible = False Then
                cmbPrelimName(iTemp).Visible = True
                cmbPrelimName(iTemp).Tag = ""
                txtPrelimFactor(iTemp).Visible = True
                txtPrelimLo(iTemp).Visible = True
                txtPrelimHi(iTemp).Visible = True
                chkPrelimOut(iTemp).Visible = True
                GoSub AddOneMoreSection
                Exit For
            ElseIf iTemp = cmbPrelimName.Count - 1 Then
                Load cmbPrelimName(cmbPrelimName.Count)
                Load txtPrelimFactor(txtPrelimFactor.Count)
                Load txtPrelimLo(txtPrelimLo.Count)
                Load txtPrelimHi(txtPrelimHi.Count)
                Load chkPrelimOut(chkPrelimOut.Count)
                cmbPrelimName(cmbPrelimName.Count - 1).Visible = True
                txtPrelimFactor(txtPrelimFactor.Count - 1).Visible = True
                txtPrelimLo(txtPrelimLo.Count - 1).Visible = True
                txtPrelimHi(txtPrelimHi.Count - 1).Visible = True
                chkPrelimOut(chkPrelimOut.Count - 1).Visible = True
                Form_Resize
                iTemp = iTemp + 1
                GoSub AddOneMoreSection
                Exit For
            End If
        Next iTemp
    Else
        If cmbFinalsName.Count = 1 Then
            Load cmbFinalsName(cmbFinalsName.Count)
            Load txtFinalsFactor(txtFinalsFactor.Count)
            Load txtFinalsLo(txtFinalsLo.Count)
            Load txtFinalsHi(txtFinalsHi.Count)
            Load chkFinalsOut(chkFinalsOut.Count)
            Load chkRecycle(chkRecycle.Count)
        End If
        For iTemp = 1 To cmbFinalsName.Count - 1
            If chkFinalsOut(iTemp).Visible = False Then
                lblFinalsOut.Visible = True
                chkFinalsOut(iTemp).Visible = True
            End If
            If cmbFinalsName(iTemp).Visible = False Then
                cmbFinalsName(iTemp).Visible = True
                cmbFinalsName(iTemp).Tag = ""
                txtFinalsFactor(iTemp).Visible = True
                txtFinalsLo(iTemp).Visible = True
                txtFinalsHi(iTemp).Visible = True
                chkFinalsOut(iTemp).Visible = True
                GoSub AddOneMoreSection
                Exit For
            ElseIf iTemp = cmbFinalsName.Count - 1 Then
                Load cmbFinalsName(cmbFinalsName.Count)
                Load txtFinalsFactor(txtFinalsFactor.Count)
                Load txtFinalsLo(txtFinalsLo.Count)
                Load txtFinalsHi(txtFinalsHi.Count)
                Load chkFinalsOut(chkFinalsOut.Count)
                Load chkRecycle(chkRecycle.Count)
                cmbFinalsName(cmbFinalsName.Count - 1).Visible = True
                txtFinalsLo(txtFinalsLo.Count - 1).Visible = True
                txtFinalsHi(txtFinalsHi.Count - 1).Visible = True
                txtFinalsFactor(txtFinalsFactor.Count - 1).Visible = True
                chkFinalsOut(chkFinalsOut.Count - 1).Visible = True
                Form_Resize
                iTemp = iTemp + 1
                GoSub AddOneMoreSection
                Exit For
            End If
        Next iTemp
    End If
    Form_Resize
    On Local Error GoTo 0
Exit Sub

AddOneMoreSection:
    If tbsTest.SelectedItem.Index = 2 Then
        txtPrelimFactor(iTemp).Text = "1"
        txtPrelimLo(iTemp).Text = Format$(0, "0." & String(Val(cmbMarkDecimals(0).Text), "0"))
        txtPrelimHi(iTemp).Text = Format$(10, "0." & String(Val(cmbMarkDecimals(0).Text), "0"))
        cmbPrelimName(iTemp).Clear
        With dtaName.Recordset
            .MoveFirst
            Do While Not .EOF
                cmbPrelimName(iTemp).AddItem .Fields(0)
                .MoveNext
            Loop
        End With
        cmbPrelimName(iTemp).Text = Translate("Section", mcLanguage) & " " & cmbPrelimName.Count
        SelectRow iTemp
    Else
        txtFinalsFactor(iTemp).Text = "1"
        txtFinalsLo(iTemp).Text = Format$(0, "0." & String(Val(cmbMarkDecimals(1).Text), "0"))
        txtFinalsHi(iTemp).Text = Format$(10, "0." & String(Val(cmbMarkDecimals(1).Text), "0"))
        cmbFinalsName(iTemp).Clear
        With dtaName.Recordset
            .MoveFirst
            Do While Not .EOF
                cmbFinalsName(iTemp).AddItem .Fields(0)
                .MoveNext
            Loop
        End With
        cmbFinalsName(iTemp).Text = Translate("Section", mcLanguage) & " " & cmbFinalsName.Count
        SelectRow iTemp
    End If
Return

End Sub
Private Sub CheckDecimals()
    Dim iTemp As Integer
    Dim iMarkVisible(0 To 1) As Integer
    Dim iMarkEnabled(0 To 1) As Integer
    Dim iTimeVisible(0 To 1) As Integer
    Dim iTimeEnabled(0 To 1) As Integer
    
    If dbcPre.BoundText = "3" Then 'time
        iTimeVisible(0) = True
        iMarkVisible(0) = False
        If cmbTimeDecimals(0).Text = "" Then
            With dtaTest.Recordset
                .Edit
                .Fields("Time_decimals") = 2
                .Update
            End With
        End If
    Else
        iTimeVisible(0) = False
        iMarkVisible(0) = True
    End If
    If dbcPre.BoundText = "1" Then 'marks
        iMarkVisible(0) = True
        iTimeVisible(0) = False
        If cmbMarkDecimals(0).Text = "" Then
            With dtaTest.Recordset
                .Edit
                .Fields("Mark_decimals") = 1
                .Update
            End With
        End If
    Else
        iMarkVisible(0) = False
        iTimeVisible(0) = True
    End If
    
    If dbcFin.BoundText = "3" Then 'time
        iTimeVisible(1) = True
        iMarkVisible(1) = False
        If cmbTimeDecimals(1).Text = "" Then
            With dtaTest.Recordset
                .Edit
                .Fields("Time_decimals") = 2
                .Update
            End With
        End If
    Else
        iTimeVisible(1) = False
        iMarkVisible(1) = True
    End If
    If dbcFin.BoundText = "1" Then 'marks
        iMarkVisible(1) = True
        iTimeVisible(1) = False
        If cmbMarkDecimals(1).Text = "" Then
            With dtaTest.Recordset
                .Edit
                .Fields("Mark_decimals") = 1
                .Update
            End With
        End If
    Else
        iMarkVisible(1) = False
        iTimeVisible(1) = True
    End If
    
    If dbcPre.BoundText = "0" Or dbcPre.BoundText = "2" Then
        iMarkEnabled(0) = False
        iTimeVisible(0) = False
    End If
    
    If dbcFin.BoundText = "0" Or dbcFin.BoundText = "2" Then
        iMarkEnabled(1) = False
        iTimeVisible(1) = False
    End If
    
    cmbTimeDecimals(0).Visible = iTimeVisible(0)
    cmbMarkDecimals(0).Visible = iMarkVisible(0)
    lblPrelimLo.Visible = iMarkVisible(0)
    lblPrelimHi.Visible = iMarkVisible(0)
    For iTemp = 1 To txtPrelimLo.Count - 1
        txtPrelimLo(iTemp).Visible = iMarkVisible(0)
        txtPrelimHi(iTemp).Visible = iMarkVisible(0)
    Next iTemp
    
    cmbTimeDecimals(1).Visible = iTimeVisible(1)
    cmbMarkDecimals(1).Visible = iMarkVisible(1)
    lblFinalsLo.Visible = iMarkVisible(1)
    lblFinalsHi.Visible = iMarkVisible(1)
    For iTemp = 1 To txtFinalsLo.Count - 1
        txtFinalsLo(iTemp).Visible = iMarkVisible(1)
        txtFinalsHi(iTemp).Visible = iMarkVisible(1)
    Next iTemp
End Sub

Private Sub tbsTest_Click()
    Dim cTemp As String
    Dim iTemp As Integer
    
    For iTemp = 1 To tbsTest.Tabs.Count
        tbsTest.Tabs(iTemp).HighLighted = False
        fraTbsTest(iTemp).Visible = False
    Next iTemp
    tbsTest.SelectedItem.HighLighted = miUseHighLights
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    
    fraTbsTest(tbsTest.SelectedItem.Index).Visible = True
    If tbsTest.SelectedItem.Index = 3 Then
        If fraTbsTest(tbsTest.SelectedItem.Index).Enabled = True Then
            cmdAdd.Enabled = True
            cmdDelete.Enabled = True
        End If
    End If
        
End Sub

Private Sub txtFinalsFactor_Change(Index As Integer)
    SelectRow Index
End Sub

Private Sub txtFinalsFactor_Click(Index As Integer)
    SelectRow Index
End Sub

Private Sub txtFinalsHi_Change(Index As Integer)
    SelectRow Index
End Sub

Private Sub txtFinalsHi_Click(Index As Integer)
    SelectRow Index

End Sub

Private Sub txtFinalsLo_Change(Index As Integer)
   SelectRow Index
End Sub

Private Sub txtFinalsLo_Click(Index As Integer)
    SelectRow Index

End Sub

Private Sub txtOther_Validate(Index As Integer, Cancel As Boolean)
    If Index = 2 And txtOther(2) <> "" Then
        txtOther(2).Text = Format$(Val(txtOther(2).Text), "0.0")
    End If

End Sub

Private Sub txtPrelimFactor_Change(Index As Integer)
    SelectRow Index
End Sub

Private Sub txtPrelimFactor_Click(Index As Integer)
    SelectRow Index
End Sub
Public Function Check_out()
    Dim rst As DAO.Recordset
    Dim iTemp As Integer
    Dim iCount As Integer
    
    For iTemp = 0 To chkFinalsOut.Count - 1
        If chkFinalsOut(iTemp).Value <> 0 Then
            iCount = iCount + 1
        End If
    Next iTemp
    
    If iCount > 0 Then
        Me.dbcOther(0).Enabled = True
        If Me.dbcOther(0).Text = "" Or Me.dbcOther(0).Text = "0" Then
            With dtaTest.Recordset
                .Edit
                .Fields("Out_Fin") = 1
                .Update
            End With
        End If
    Else
        For iTemp = 0 To chkPrelimOut.Count - 1
            If chkPrelimOut(iTemp).Value <> 0 Then
                iCount = iCount + 1
            End If
        Next iTemp
        If iCount > 0 Then
            MsgBox Translate("Please make sure that settings for finals are the ame as for preliminary rounds.", mcLanguage)
            Me.dbcOther(0).Enabled = True
            If Me.dbcOther(0).Text = "" Or Me.dbcOther(0).Text = "0" Then
                With dtaTest.Recordset
                    .Edit
                    .Fields("Out_Fin") = 1
                    .Update
                End With
            End If
        Else
            With dtaTest.Recordset
                .Edit
                .Fields("Out_Fin") = 0
                .Update
            End With
            Me.dbcOther(0).Enabled = False
        End If
    End If
End Function
Private Sub txtPrelimHi_Change(Index As Integer)
    SelectRow Index

End Sub
Private Sub txtPrelimHi_Click(Index As Integer)
    SelectRow Index

End Sub
Private Sub txtPrelimLo_Change(Index As Integer)
    SelectRow Index

End Sub

Private Sub txtPrelimLo_Click(Index As Integer)
    SelectRow Index
End Sub
