Attribute VB_Name = "IceDiv"
Option Explicit
Sub CreateDivisionTable()
    Dim curFactor As Currency
    Dim curMax As Currency
    Dim curValue As Currency
    Dim curResult As Currency
    Dim cTemp As String
    
    Dim iFilenum As Integer
    
    Const curHalve = 0.5
    
    iFilenum = FreeFile
    Open "c:\klad\div.txt" For Output As #iFilenum
    For curFactor = 2 To 7
         curMax = curFactor * 10
         curValue = curMax
         Print #iFilenum, curFactor
         Do While curValue > 0
            cTemp = Format$(curValue, "00.0") & " / " & curFactor & " = " & Format$(curValue / curFactor, "00.0")
            If Left$(cTemp, 1) = "0" Then Mid$(cTemp, 1, 1) = " "
            If Right$(Format$(curValue, "00.0"), 3) = Format$(0, "0.0") Or Right$(Format$(curValue, "00.0"), 3) = Format$(0, "5.0") Then
                Print #iFilenum, ""
            End If
            Print #iFilenum, Replace(cTemp, "= 0", "=  ")
            curValue = curValue - curHalve
         Loop
         Print #iFilenum, Chr$(14)
    Next curFactor
    
End Sub
