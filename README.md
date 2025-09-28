# PV

Repo pentru workbook-ul Excel (Procese Verbale + Fișe).

## Cum încarci workbook-ul

1. (Recomandat) Activează Git LFS local:
   - Instalează Git LFS: https://git-lfs.com
   - Rulează o singură dată: `git lfs install`
2. Creează o ramură, de ex. `upload/xlsm`.
3. Adaugă fișierul Excel (ex. `PV.xlsm`) și fă commit + push.
4. Deschide un Pull Request către `main`.

Notă: Fișierele `.xlsm/.xlam/.xlsb` sunt binare; Git LFS previne “umflarea” repo-ului.

## Pasul următor (recomandat)

Mutarea logicii VBA într-un add-in `.xlam` parolat:
- Codul principal se mută într-un add-in protejat (vizualizare blocată).
- Workbook-ul rămâne UI + date; apelează rutinele din add-in (prin `Application.Run`).
- Beneficii: mai sigur, actualizări ușoare, dif-uri clare în PR-uri.

Vom furniza:
- Structura proiect add-in (`addin/`), module sursă și instrucțiuni BUILD/INSTALL.
- Wrappers în workbook (callback-uri Ribbon, evenimente) care rutează spre add-in.
