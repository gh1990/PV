Option Explicit

' Workbook-ul gazdă care folosește add-in-ul.
Public HostWB As Workbook

Public Sub InitHost(ByVal wb As Workbook)
    Set HostWB = wb
End Sub

' Returnează workbook-ul gazdă, cu fallback pe ThisWorkbook.
Public Function GetHost() As Workbook
    If HostWB Is Nothing Then
        Set GetHost = ThisWorkbook
    Else
        Set GetHost = HostWB
    End If
End Function
