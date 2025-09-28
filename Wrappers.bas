Option Explicit

' Numele add-in-ului (fișierul .xlam). Schimbă dacă folosești alt nume.
Private Const ADDIN_NAME As String = "PV_Addin.xlam"

' Încearcă să încarce add-in-ul dacă nu este deja activat.
' Poți furniza opțional o cale completă (addinPath) dacă nu este în lista Add-ins.
Public Sub EnsureAddinLoaded(Optional ByVal addinPath As String = "")
    Dim ai As AddIn
    Dim found As Boolean
    For Each ai In Application.AddIns
        If StrComp(ai.Name, ADDIN_NAME, vbTextCompare) = 0 Then
            If Not ai.Installed Then ai.Installed = True
            found = True
            Exit For
        End If
    Next ai
    If Not found And Len(addinPath) > 0 Then
        On Error Resume Next
        Application.Workbooks.Open addinPath
        On Error GoTo 0
        For Each ai In Application.AddIns
            If StrComp(ai.Name, ADDIN_NAME, vbTextCompare) = 0 Then
                ai.Installed = True
                Exit For
            End If
        Next ai
    End If
End Sub

' ——— Callback-uri Ribbon rulate în workbook, delegă în add-in
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!RibbonOnLoad", ribbon
    On Error GoTo 0
End Sub

Public Sub AddProcesVerbalNou(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddProcesVerbalNou", control
    On Error GoTo 0
End Sub

Public Sub AddFormularNorma(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddFormularNorma", control
    On Error GoTo 0
End Sub

Public Sub AddFormularTransport(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddFormularTransport", control
    On Error GoTo 0
End Sub

Public Sub AddObiectPV(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddObiectPV", control
    On Error GoTo 0
End Sub

Public Sub AddNormaPV(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddNormaPV", control
    On Error GoTo 0
End Sub

Public Sub AddMaterialePV(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddMaterialePV", control
    On Error GoTo 0
End Sub

Public Sub AddTransportPV(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddTransportPV", control
    On Error GoTo 0
End Sub

Public Sub AddUtilajPV(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddUtilajPV", control
    On Error GoTo 0
End Sub

Public Sub AddNormaBD(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddNormaBD", control
    On Error GoTo 0
End Sub

Public Sub AddMaterialeBD(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddMaterialeBD", control
    On Error GoTo 0
End Sub

Public Sub AddUtilajBD(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddUtilajBD", control
    On Error GoTo 0
End Sub

Public Sub AddTransportBD(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddTransportBD", control
    On Error GoTo 0
End Sub

Public Sub AddListeBD(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!AddListeBD", control
    On Error GoTo 0
End Sub

Public Sub FormularPV(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!FormularPV", control
    On Error GoTo 0
End Sub

Public Sub FormularFise(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!FormularFise", control
    On Error GoTo 0
End Sub

Public Sub CalcMateriale(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!CalcMateriale", control
    On Error GoTo 0
End Sub

Public Sub FormeazaFisa(control As IRibbonControl)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!FormeazaFisa", control
    On Error GoTo 0
End Sub

Public Sub StergePVsiFise(Optional ByVal cereConfirmare As Boolean = True)
    On Error Resume Next
    Application.Run "'" & ADDIN_NAME & "'!StergePVsiFise", cereConfirmare
    On Error GoTo 0
End Sub
