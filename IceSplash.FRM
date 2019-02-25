VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      Height          =   4590
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   7380
      Begin VB.Image Image1 
         Height          =   2415
         Left            =   120
         Picture         =   "IceSplash.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Rechts
         Caption         =   "LicenseTo"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Tag             =   "LicenseTo"
         Top             =   300
         Width           =   6855
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   29.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2520
         TabIndex        =   7
         Tag             =   "Product"
         Top             =   480
         Width           =   2220
      End
      Begin VB.Label lblCompanyProduct 
         Caption         =   "CompanyProduct"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2520
         TabIndex        =   6
         Tag             =   "CompanyProduct"
         Top             =   1200
         Width           =   4665
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   6285
         TabIndex        =   5
         Tag             =   "Version"
         Top             =   2400
         Width           =   930
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         Height          =   795
         Left            =   120
         TabIndex        =   2
         Tag             =   "Warning"
         Top             =   3720
         Width           =   7080
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Rechts
         Caption         =   "Company"
         Height          =   495
         Left            =   2760
         TabIndex        =   4
         Tag             =   "Company"
         Top             =   3120
         Width           =   4455
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Rechts
         Caption         =   "Copyright"
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Tag             =   "Copyright"
         Top             =   2760
         Width           =   4575
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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


Private Sub Form_Click()
   Hide
End Sub

Private Sub Form_Load()
    Me.lblCopyright.Caption = App.LegalCopyright
    Me.lblCompany.Caption = App.CompanyName
    Me.lblCompanyProduct.Caption = App.FileDescription
    Me.lblLicenseTo.Caption = MachineName & " - [" & UserName & "]"
    Me.lblProductName.Caption = App.Title
    Me.lblVersion.Caption = "Version" & " " & App.Major & "." & App.Minor & "." & Format$(App.Revision, "000")
    Me.lblWarning.Caption = "This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2 of the License, or (at your option) any later version."

End Sub

Private Sub fraMainFrame_DragDrop(Source As Control, X As Single, Y As Single)
   Hide
End Sub
