Attribute VB_Name = "modProgBar"
' Copyright (C) Marko Mazeland and/or Datawerken Holding BV 2003
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
Public Sub IncreaseProgressbarValue(p As ProgressBar, Optional l As Long = 1)
    If p.Value + l <= p.Max Then
        p.Value = p.Value + l
    Else
        p.Value = p.Max
    End If
    If p.Value Mod 100 = 0 Or p < 10 Then
        DoEvents
    End If
End Sub
Public Sub SetProgressBarMax(p As ProgressBar, lMax As Long)
    
    p.Max = lMax
    p.Min = 0
    p.Value = p.Min
End Sub
Public Sub ShowProgressbar(F As Form, iPanelNum As Integer, lMaximum As Long)
    If lMaximum > 0 Then
        Dim c As Control
        Dim iGevonden As Integer
        
        F.ProgressBar1.Visible = True
        SetProgressBarMax F.ProgressBar1, lMaximum
        
        iGevonden = False
        For Each c In F.Controls
            If c.Name = "StatusBar1" Or Left$(c.Name, 3) = "stb" Then
                iGevonden = True
                Exit For
            End If
        Next
        If iGevonden = True Then
            Do While c.Panels.Count < iPanelNum
                c.Panels.Add
            Loop
            c.Panels(iPanelNum).AutoSize = 1
            If iPanelNum > 1 Then
                c.Panels(iPanelNum - 1).AutoSize = 2
                c.Panels(iPanelNum - 1).MinWidth = 1
            End If
            F.ProgressBar1.Top = c.Top + (c.Height * 0.1)
            F.ProgressBar1.Left = c.Panels(iPanelNum).Left
            F.ProgressBar1.Width = c.Panels(iPanelNum).Width
            F.ProgressBar1.Height = c.Height * 0.9
        End If
        DoEvents
    Else
        F.ProgressBar1.Visible = False
    End If
End Sub



