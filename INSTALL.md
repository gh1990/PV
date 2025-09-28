# INSTALL — PV_Addin.xlam

Pașii pentru instalarea add-in-ului în Excel și legarea workbook-ului.

1) Copiază PV_Addin.xlam
- Pune fișierul într-un folder “trusted”, de exemplu:
  - Windows: `%AppData%\Microsoft\AddIns\PV_Addin.xlam`
  - Sau într-un folder local (ex. Documente\PV\PV_Addin.xlam)

2) Activează add-in-ul în Excel
- Excel > File > Options > Add-ins
- Manage: Excel Add-ins > Go…
- Browse… > alege `PV_Addin.xlam` > OK
- Bifează `PV_Addin` în listă > OK

3) Deschide workbook-ul PV.xlsm
- La `Workbook_Open`, workbook-ul cheamă `InitHost` din add-in și va delega toate callback-urile de Ribbon și evenimentele.
- Dacă add-in-ul nu e găsit în lista de Add-ins, wrapper-ul din workbook încearcă să-l încarce dacă îi dai o cale.

4) Verificări
- Butonul “AddProcesVerbalNou” creează PV nou corect.
- “FormeazaFisa” funcționează.
- Foile model/BD rămân neschimbate ca nume/poziție.

5) Parolă/semnătură
- Codul add-in-ului este protejat pentru vizualizare (parolă setată la build).
- Opțional: semnează digital proiectul VBA (Tools > Digital Signature) pentru medii cu politici stricte de macro-uri.
