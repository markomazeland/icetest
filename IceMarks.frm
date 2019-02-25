VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMarks 
   Caption         =   "Marks"
   ClientHeight    =   9840
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   12405
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Index           =   0
      Left            =   3720
      Top             =   8280
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "IceMarks.frx":0000
      Height          =   1095
      Index           =   4
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   1931
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "IceMarks.frx":0018
      Height          =   1095
      Index           =   3
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   1931
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "IceMarks.frx":0030
      Height          =   1095
      Index           =   2
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   1931
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "IceMarks.frx":0048
      Height          =   1095
      Index           =   1
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6000
      Visible         =   0   'False
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   1931
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "IceMarks.frx":0060
      Height          =   1095
      Index           =   0
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7080
      Visible         =   0   'False
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   1931
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         AllowRowSizing  =   0   'False
         RecordSelectors =   0   'False
         BeginProperty Column00 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   9345
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   873
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.TabStrip tbsJudges 
      Height          =   735
      Left            =   3000
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      TabWidthStyle   =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Index           =   2
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Index           =   3
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Index           =   4
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   2
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Menu mnuOrder 
      Caption         =   "Order"
      Begin VB.Menu mnuOrderOrder 
         Caption         =   "by Starting order"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuOrderOrder 
         Caption         =   "by Startnumber"
         Index           =   1
      End
      Begin VB.Menu mnuOrderOrder 
         Caption         =   "by Rider"
         Index           =   2
      End
      Begin VB.Menu mnuOrderOrder 
         Caption         =   "by Horse"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Entry form for individual marks per participant in preliminary rounds
'
' Copyright (C) Marko Mazeland 2006
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

Dim miRowColBusy As Integer
Dim miHorizontal As Integer

Private Sub DataGrid1_AfterColEdit(Index As Integer, ByVal ColIndex As Integer)
    With Adodc1(Index)
        If ColIndex > 3 Then
            DoEvents
            ValidateMark DataGrid1(Index)
            .Recordset.Bookmark = .Recordset.Bookmark
            Calculatetotal Index, Adodc1(Index).Recordset.AbsolutePosition
        End If
    End With
End Sub

Private Sub DataGrid1_Keydown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim iTemp As Integer
    If miRowColBusy = False Then
        miRowColBusy = True
        If KeyCode = vbKeyReturn Then
            With Adodc1(Index)
                iTemp = ValidateMark(DataGrid1(Index))
                .Recordset.Bookmark = .Recordset.Bookmark
            End With
            If iTemp = True Then
                Calculatetotal Index, Adodc1(Index).Recordset.AbsolutePosition
                DataGrid1(Index).Col = DataGrid1(Index).Col + 1
                If DataGrid1(Index).Col = DataGrid1(Index).Columns.Count - 2 Then
                    On Local Error Resume Next
                    DataGrid1(Index).Row = DataGrid1(Index).Row + 1
                    On Local Error GoTo 0
                    DataGrid1(Index).Col = 4
                End If
            End If
        End If
        miRowColBusy = False
    End If
End Sub

Private Sub DataGrid1_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    
    If miRowColBusy = False Then
        miRowColBusy = True
        With Adodc1(Index)
           .Recordset.Bookmark = .Recordset.Bookmark
        End With
        If DataGrid1(Index).Col = DataGrid1(Index).Columns.Count - 2 Then
            On Local Error Resume Next
            DataGrid1(Index).Row = DataGrid1(Index).Row + 1
            On Local Error GoTo 0
            DataGrid1(Index).Col = 4
        ElseIf DataGrid1(Index).Col < 4 Then
            If DataGrid1(Index).Row > 0 Then
                DataGrid1(Index).Row = DataGrid1(Index).Row - 1
                DataGrid1(Index).Col = DataGrid1(Index).Columns.Count - 3
            Else
                DataGrid1(Index).Col = 4
            End If
        End If
        DataGrid1(Index).SelStart = 0
        DataGrid1(Index).SelLength = 5
        miRowColBusy = False
    End If
End Sub

Private Sub Form_Load()
    Dim iTemp As Integer
    Dim cTemp As String
    Dim rstTest As DAO.Recordset
    
    SetMouseHourGlass
    
    ReadFormPosition Me
    
    ChangeFontSize Me, msFontSize
    
    TranslateControls Me
        
    ReadIniFile gcIniFile, "Sectionmarks", "Order", cTemp
    For iTemp = 0 To mnuOrderOrder.Count - 1
        If iTemp = Val(cTemp) Then
            mnuOrderOrder(iTemp).Checked = True
        Else
            mnuOrderOrder(iTemp).Checked = False
        End If
    Next iTemp
    
    Caption = frmMain.TestCode & "-" & frmMain.TestName & " (" & frmMain.tbsSelFin.SelectedItem.Caption & ")"
    
    With tbsJudges
        For iTemp = 1 To frmMain.TestJudges
            If .Tabs.Count < iTemp Then
                .Tabs.Add
            End If
            .Tabs(iTemp).Caption = Translate("Judge", mcLanguage) & " " & iTemp
        Next iTemp
    End With
        
    If PopulateTempTable = False Then
        Unload Me
    End If
    
    '+++PopulateGrid
    
    With tbsJudges.Tabs(1)
        .Selected = True
        .HighLighted = miUseHighLights
    End With
        
    Form_Resize
    
    SetMouseNormal
    
End Sub

Private Sub Form_Resize()
    Dim iTemp As Integer
    Dim iField As Integer
    
    On Local Error Resume Next
    
    With tbsJudges
        .Top = 0
        .Left = 0
        .Width = ScaleWidth
        .Height = ScaleHeight - Me.StatusBar1.Height
        .TabMinWidth = frmMain.miTabMinWidth
        .TabFixedHeight = frmMain.miTabMinHeight
        .TabFixedWidth = Me.ScaleWidth \ (frmMain.TestJudges + 1)
    End With
    
    For iTemp = 1 To frmMain.TestJudges
        With DataGrid1(iTemp - 1)
            .Top = tbsJudges.ClientTop
            .Width = tbsJudges.ClientWidth
            .Left = tbsJudges.ClientLeft
            .Height = tbsJudges.ClientHeight
            For iField = 0 To Adodc1(iTemp - 1).Recordset.Fields.Count - 1
                With .Columns(iField)
                    If iField = 1 Or iField = 2 Then
                        .Width = tbsJudges.ClientWidth \ 4
                    Else
                        .Width = tbsJudges.ClientWidth \ ((Adodc1(iTemp - 1).Recordset.Fields.Count - 3.5) * 2)
                    End If
                End With
            Next iField
        End With
    Next iTemp
    
    On Local Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim iTemp As Integer
    
    If Not Cancel = True Then
        For iTemp = 1 To tbsJudges.Tabs.Count
            DataGrid1_Keydown (iTemp - 1), vbKeyReturn, 0
            On Local Error Resume Next
            DataGrid1(iTemp - 1).Row = DataGrid1(iTemp - 1).Row + 1
            DataGrid1(iTemp - 1).Row = 0
            On Local Error GoTo 0
            
        Next iTemp
        
        ExtractFromTempTable
    End If
    
    WriteFormPosition Me

End Sub

Private Sub mnuOrderOrder_Click(Index As Integer)
    Dim iTemp As Integer
    For iTemp = 0 To mnuOrderOrder.Count - 1
        If iTemp = Index Then
            mnuOrderOrder(iTemp).Checked = True
        Else
            mnuOrderOrder(iTemp).Checked = False
        End If
    Next iTemp
    
    WriteIniFile gcIniFile, "Sectionmarks", "Order", Format$(Index)
    
    PopulateGrid

End Sub
Private Sub tbsJudges_Click()
    Dim iTemp As Integer
    
    For iTemp = 1 To tbsJudges.Tabs.Count
        If tbsJudges.Tabs(iTemp).Selected = True Then
            tbsJudges.Tabs(iTemp).HighLighted = miUseHighLights
            DataGrid1(iTemp - 1).Visible = True
        Else
            tbsJudges.Tabs(iTemp).HighLighted = False
            DataGrid1(iTemp - 1).Visible = False
        End If
    Next iTemp
    
End Sub
Sub PopulateGrid()
    Dim iJudge As Integer
    Dim iField As Integer
    Dim cQry As String
    
    DoEvents
    
    For iJudge = 1 To frmMain.TestJudges
        If Me.mnuOrderOrder(3).Checked = True Then
            cQry = "SELECT * FROM [" & mcTempTableName & "] WHERE Judge=" & iJudge & " ORDER BY Horse"
        ElseIf Me.mnuOrderOrder(2).Checked = True Then
            cQry = "SELECT * FROM [" & mcTempTableName & "] WHERE Judge=" & iJudge & " ORDER BY Rider"
        ElseIf Me.mnuOrderOrder(1).Checked = True Then
            cQry = "SELECT * FROM [" & mcTempTableName & "] WHERE Judge=" & iJudge & " ORDER BY STA"
        Else
            cQry = "SELECT * FROM [" & mcTempTableName & "] WHERE Judge=" & iJudge & " ORDER BY POS"
        End If
        
        With Adodc1(iJudge - 1)
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mcDatabaseName & ";Persist Security Info=False"
            .RecordSource = cQry
            .Refresh
        End With
           
        '***assign the connection to the database; horizontal
        With DataGrid1(iJudge - 1)
            For iField = 0 To Adodc1(iJudge - 1).Recordset.Fields.Count - 1
                If .Columns.Count < iField + 1 Then
                    .Columns.Add .Columns.Count
                End If
                With .Columns(iField)
                    If iField > 3 And iField < Adodc1(iJudge - 1).Recordset.Fields.Count - 2 Then
                        .Caption = mcSectionName(iField - 3)
                    Else
                        .Caption = Translate(Adodc1(iJudge - 1).Recordset.Fields(iField).Name, mcLanguage)
                    End If
                    .DataField = Adodc1(iJudge - 1).Recordset.Fields(iField).Name
                    If iField <= 3 Or iField = Adodc1(iJudge - 1).Recordset.Fields.Count - 2 Then
                        .Locked = True
                    Else
                        .Locked = False
                    End If
                    If iField = 3 Or iField = Adodc1(iJudge - 1).Recordset.Fields.Count - 1 Then
                        .Visible = False
                    Else
                        .Visible = True
                    End If
               End With

                DoEvents
            Next iField
            .Refresh
            .ReBind
        End With
    Next iJudge
    
End Sub
Private Sub Calculatetotal(iIndex As Integer, iRow As Integer)
    Dim iTemp As Integer
    Dim curTemp As Currency
    Dim cMarks As String
    
    curTemp = 0
    For iTemp = 4 To DataGrid1(iIndex).Columns.Count - 3
        If Val(DataGrid1(iIndex).Columns(iTemp)) >= 0 And Trim$(DataGrid1(iIndex).Columns(iTemp)) <> "" Then
            DataGrid1(iIndex).Columns(iTemp) = Format$(Val(Replace(DataGrid1(iIndex).Columns(iTemp), ",", ".")), "0.0")
            curTemp = curTemp + Val(DataGrid1(iIndex).Columns(iTemp)) * mcurSectionFactor(iTemp - 3)
            cMarks = cMarks & DataGrid1(iIndex).Columns(iTemp)
        End If
    Next iTemp
    If cMarks <> "" Then
        DataGrid1(iIndex).Columns(DataGrid1(iIndex).Columns.Count - 2) = Format$(curTemp / mcurTestFactor, frmMain.TestMarkFormat)
    Else
        DataGrid1(iIndex).Columns(DataGrid1(iIndex).Columns.Count - 2) = ""
    End If
    
    DoEvents
    
End Sub

Private Function ValidateMark(tObj As Object) As Integer

   Dim curMark As Currency
   
   ValidateMark = True
   
   If Trim$(tObj.Text) <> "" Then
      If (frmMain.TestStatus = 0 And frmMain.dtaTest.Recordset.Fields("Type_pre") = 2) Or (frmMain.TestStatus <> 0 And frmMain.dtaTest.Recordset.Fields("Type_Final") = 2) Then
            tObj.Text = Val(tObj.Text)
      ElseIf InStr(tObj.Text, ",") = 0 And InStr(tObj.Text, ".") = 0 And Val(tObj.Text) > 10 Then
          If Err > 0 Then
            tObj.Text = ""
          End If
          If tObj.Text <= 10 ^ (frmMain.TestMarkDecimals) Then
             tObj.Text = Val(tObj.Text) / 10 ^ (frmMain.TestMarkDecimals - 1)
          ElseIf tObj.Text <= 10 ^ (frmMain.TestMarkDecimals + 1) Then
             tObj.Text = Val(tObj.Text) / 10 ^ frmMain.TestMarkDecimals
          ElseIf tObj.Text <= 10 ^ (frmMain.TestMarkDecimals + 2) Then
             tObj.Text = Val(tObj.Text) / 10 ^ (frmMain.TestMarkDecimals + 1)
          End If
      End If
      curMark = MakeStringValue(tObj.Text)
      If ((frmMain.TestStatus = 0 And frmMain.dtaTest.Recordset.Fields("Type_pre") = 1) Or (frmMain.TestStatus <> 0 And frmMain.dtaTest.Recordset.Fields("Type_Final") = 1)) And curMark < frmMain.dtaTestSection.Recordset.Fields("Mark_low") Or curMark > frmMain.dtaTestSection.Recordset.Fields("Mark_hi") Then
         ValidateMark = False
         StatusMessage Translate("Only marks between", mcLanguage) & " " & frmMain.dtaTestSection.Recordset.Fields("Mark_low") & " - " & frmMain.dtaTestSection.Recordset.Fields("Mark_hi") & " " & Translate("are accepted", mcLanguage) & " !", 2
         Me.StatusBar1.SimpleText = Translate("Only marks between", mcLanguage) & " " & frmMain.dtaTestSection.Recordset.Fields("Mark_low") & " - " & frmMain.dtaTestSection.Recordset.Fields("Mark_hi") & " " & Translate("are accepted", mcLanguage) & " !"
         tObj.Text = ""
         miNoBackupNow = True
         PlaySound "SYSTEMEXCLAMATION", 0, 1
      Else
         miNoBackupNow = False
      End If
   Else
         miNoBackupNow = True
   End If
   
   DoEvents
   
End Function

