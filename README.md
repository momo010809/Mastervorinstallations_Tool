# Setup Tool

Ein kleines Windows-Setup-Programm in C# mit WinForms. Es erstellt eine Installationsordnerstruktur, schreibt ein Beispielprotokoll und kann später beliebige Setup-Aktionen ersetzen.

## Eigenschaften

- Windows GUI mit einem sauberen Setup-Button
- Admin-Rechte beim Start angefordert über `app.manifest`
- Beispielaktionen in `MainForm.cs` können durch eigene Setup-Logik ersetzt werden

## Dateien

- `SetupTool.csproj` - .NET WinForms Projektdatei
- `Program.cs` - Programmstart
- `MainForm.cs` - Logik und Setup-Ablauf
- `MainForm.Designer.cs` - UI-Layout
- `app.manifest` - UAC / Admin-Anforderung

## Build

1. Installiere das .NET SDK 11.0 Preview oder neuer.
2. Öffne eine Kommandozeile im Ordner `c:\SetupTool`.
3. Führe aus:

```powershell
dotnet publish -c Release
```

4. Die Ausgabedateien liegen unter `bin\Release\net11.0-windows\win-x64\publish`.

## Start Auf Zielsystem

- Starte direkt `SetupTool.exe`.
- Es ist als Self-Contained Single-File veröffentlicht (eine ausführbare Datei, keine separate .NET-Installation nötig).

## Anpassung

- Ersetze die Methode `PerformSetupActions()` in `MainForm.cs` mit deiner eigenen Installationslogik.
- Verwende `ExecuteCommand("cmd.exe", "/c ...")`, um Batch-Äquivalente auszuführen.
- Entferne oder ändere `app.manifest`, wenn du keine Admin-Elevation benötigst.
