# BUILD — PV_Addin.xlam

Acest document descrie cum construiești add-in-ul `PV_Addin.xlam` din sursele din `addin/src`.

1) Creează un proiect Excel Add-In
- Deschide Excel > un workbook gol.
- Alt+F11 (VBE) > File > Save As > alege tip: "Excel Add-In (*.xlam)".
- Salvează ca `PV_Addin.xlam` într-un folder “trusted” (ex. %AppData%\Microsoft\AddIns\PV_Addin.xlam sau un folder local pe care îl vei încărca apoi în Excel).

2) Importă sursele din acest repo
- În VBE, selectează proiectul add-in > File > Import File:
  - `addin/src/Module_Init.bas`
  - `addin/src/RibbonExports.bas`
- Exportă din workbook-ul tău actual modulele cu logică:
  - `Module1.bas` (toată logica PV/Fise/Materiale)
  - `UserForm1.frm` + `UserForm1.frx`
- Importă-le în proiectul add-in (File > Import File).

3) Ajustează referințele (dacă e cazul)
- Dacă există referințe la `ThisWorkbook` în logică:
  - Poți înlocui cu `GetHost()` din `Module_Init.bas`, sau
  - Le poți lăsa cum sunt dacă workbook-ul țintă e “gazda” (vezi inițializarea din workbook prin `InitHost`).

4) Protejează codul (parolă)
- VBE: Tools > VBAProject Properties > Protection
  - Bifează “Lock project for viewing”
  - Setează parola și confirmă
- Salvează add-in-ul, închide complet Excel, redeschide (protecția intră în vigoare după redeschidere).

5) Teste
- Încarcă add-in-ul în Excel (vezi INSTALL.md).
- Deschide workbook-ul `PV.xlsm` și validează:
  - butoanele de Ribbon rulează (creare PV, recreare Fișe etc.),
  - nu apar erori de încărcare,
  - codul nu este vizibil în add-in fără parolă.

Note
- Poți ține add-in-ul în control de versiune ca binar prin Git LFS (configurat în repo).
- Pentru update-uri: suprascrii `PV_Addin.xlam` și reîncarci add-in-ul în Excel (dacă era activ, un restart de Excel ajută).