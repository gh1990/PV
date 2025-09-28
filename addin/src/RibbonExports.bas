Option Explicit

' Proceduri exportate de add-in; apelează logica din Module1 (importat în add-in).

Public ribbonUI As IRibbonUI

' ——— Ribbon
Public Sub RibbonOnLoad(r As IRibbonUI)
    Set ribbonUI = r
    On Error Resume Next
    Module1.RibbonOnLoad r
    On Error GoTo 0
End Sub

' ——— Butoane / callback-uri
Public Sub AddProcesVerbalNou(control As IRibbonControl): Module1.AddProcesVerbalNou control: End Sub
Public Sub AddFormularNorma(control As IRibbonControl): Module1.AddFormularNorma control: End Sub
Public Sub AddFormularTransport(control As IRibbonControl): Module1.AddFormularTransport control: End Sub
Public Sub AddObiectPV(control As IRibbonControl): Module1.AddObiectPV control: End Sub
Public Sub AddNormaPV(control As IRibbonControl): Module1.AddNormaPV control: End Sub
Public Sub AddMaterialePV(control As IRibbonControl): Module1.AddMaterialePV control: End Sub
Public Sub AddTransportPV(control As IRibbonControl): Module1.AddTransportPV control: End Sub
Public Sub AddUtilajPV(control As IRibbonControl): Module1.AddUtilajPV control: End Sub
Public Sub AddNormaBD(control As IRibbonControl): Module1.AddNormaBD control: End Sub
Public Sub AddMaterialeBD(control As IRibbonControl): Module1.AddMaterialeBD control: End Sub
Public Sub AddUtilajBD(control As IRibbonControl): Module1.AddUtilajBD control: End Sub
Public Sub AddTransportBD(control As IRibbonControl): Module1.AddTransportBD control: End Sub
Public Sub AddListeBD(control As IRibbonControl): Module1.AddListeBD control: End Sub
Public Sub FormularPV(control As IRibbonControl): Module1.FormularPV control: End Sub
Public Sub FormularFise(control As IRibbonControl): Module1.FormularFise control: End Sub
Public Sub CalcMateriale(control As IRibbonControl): Module1.CalcMateriale control: End Sub
Public Sub FormeazaFisa(control As IRibbonControl): Module1.FormeazaFisa control: End Sub

Public Sub StergePVsiFise(Optional ByVal cereConfirmare As Boolean = True)
    On Error Resume Next
    Module1.StergePVsiFise cereConfirmare
    On Error GoTo 0
End Sub

' ——— Evenimente (opțional, dacă vrei logică de eveniment în add-in)
Public Sub Handle_SheetChange(ByVal Sh As Object, ByVal Target As Range)
End Sub

Public Sub Handle_SheetActivate(ByVal Sh As Object)
End Sub
