using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SetupTool
{
    public partial class MainForm : Form
    {
        private int currentPage = 0;
        private string newComputerName = string.Empty;
        private bool disableFastStartup;
        private bool enableNumLock;
        private bool disableWindowsUpdates;
        private bool isExecuting;
        private readonly Dictionary<string, bool> desktopIcons = new()
        {
            { "🖥️ Dieser PC", false },
            { "🗑️ Papierkorb", false },
            { "📁 Benutzerordner", false },
            { "⚙️ Systemsteuerung", false }
        };
        private readonly Dictionary<string, bool> localInstallerSelections = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, CheckBox> localInstallerCheckboxes = new(StringComparer.OrdinalIgnoreCase);
        private const string EmbeddedInstallerPrefix = "Install.";
        private const string DesktopCopyProgramName = "FixfaxQs15.exe";
        private FlowLayoutPanel panelLocalInstallers;
        private Label labelLocalInstallers;
        private readonly Dictionary<string, bool> initialAutostartStates = new(StringComparer.OrdinalIgnoreCase);

        private sealed class StartupEntry
        {
            public string Name { get; init; } = string.Empty;
            public string Command { get; init; } = string.Empty;
            public string Location { get; init; } = string.Empty;
            public RegistryHive? Hive { get; init; }
            public bool IsEnabled { get; set; }
            public bool InitialIsEnabled { get; set; }
        }

        public MainForm()
        {
            InitializeComponent();
            InitializeAutostartUi();
            InitializeLocalInstallersUi();
            LoadLocalInstallers();
            currentPage = 0; // Start mit Welcome Screen
            UpdatePage();
            LoadPhoto();
            InitializeFinishCredit();
        }

        private void RegisterMouseWheelNavigation(Control root)
        {
            root.MouseWheel += MainForm_MouseWheel;
            foreach (Control child in root.Controls)
            {
                RegisterMouseWheelNavigation(child);
            }
        }

        private void MainForm_MouseWheel(object sender, MouseEventArgs e)
        {
            if (isExecuting)
            {
                return;
            }

            if (sender is TextBoxBase)
            {
                return;
            }

            if (currentPage < 1 || currentPage > 8)
            {
                return;
            }

            int targetPage = currentPage + (e.Delta > 0 ? -1 : 1);
            targetPage = Math.Max(1, Math.Min(8, targetPage));

            if (targetPage == currentPage)
            {
                return;
            }

            SavePageSelections();
            currentPage = targetPage;
            UpdatePage();
        }

        private void InitializeFinishCredit()
        {
            try
            {
                if (this.Icon != null)
                {
                    pictureBoxMadeBy.Image = this.Icon.ToBitmap();
                }
            }
            catch
            {
            }
        }

        private void LoadPhoto()
        {
            try
            {
                string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Icon.ico");
                if (File.Exists(imagePath))
                {
                    this.Icon = new Icon(imagePath);
                }
                else
                {
                    this.Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                }
            }
            catch
            {
                // Fehler beim Laden des Icons ignorieren
            }
        }

        private void buttonStart_Click(object sender, EventArgs e)
        {
            currentPage = 1;
            UpdatePage();
        }

        private void tabControl_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (isExecuting)
            {
                return;
            }

            SavePageSelections();
            currentPage = tabControl.SelectedIndex + 1;
            UpdatePage();
        }

        private void buttonNext_Click(object sender, EventArgs e)
        {
            if (isExecuting)
            {
                return;
            }

            SavePageSelections();
            currentPage++;
            UpdatePage();
        }

        private void buttonBack_Click(object sender, EventArgs e)
        {
            if (isExecuting)
            {
                return;
            }

            if (currentPage > 0)
            {
                SavePageSelections();
                currentPage--;
                UpdatePage();
            }
        }

        private void buttonRestart_Click(object sender, EventArgs e)
        {
            var psi = new ProcessStartInfo("shutdown.exe", "/r /t 0")
            {
                CreateNoWindow = true,
                UseShellExecute = false
            };
            Process.Start(psi);
        }

        private void buttonFinish_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void buttonRestartExplorer_Click(object sender, EventArgs e)
        {
            try
            {
                RestartExplorer();
                MessageBox.Show("Explorer wurde neu gestartet.", "Erfolg", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Explorer-Neustart fehlgeschlagen: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void buttonRun_Click(object sender, EventArgs e)
        {
            if (isExecuting)
            {
                return;
            }

            isExecuting = true;
            SavePageSelections();
            currentPage = 9;
            UpdatePage();
            labelPageInfo.Text = "Installation läuft...";

            try
            {
                var tasks = BuildSetupTasks();
                progressBarMain.Maximum = Math.Max(tasks.Count, 1);
                progressBarMain.Value = 0;
                flowLayoutPanelProgress.Controls.Clear();
                textBoxLog.Clear();

                foreach (var task in tasks)
                {
                    var itemLabel = new Label
                    {
                        AutoSize = true,
                        Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point),
                        Text = task.Name + "...",
                        Margin = new Padding(3, 3, 3, 3)
                    };

                    Invoke(() => flowLayoutPanelProgress.Controls.Add(itemLabel));
                    try
                    {
                        await task.Action();
                        Invoke(() => itemLabel.Text = task.Name + " ✓");
                        AppendLog(task.Name + " abgeschlossen.");
                    }
                    catch (Exception ex)
                    {
                        Invoke(() => itemLabel.Text = task.Name + " ✗");
                        AppendLog($"Fehler bei {task.Name}: {ex.Message}");
                    }

                    Invoke(() =>
                    {
                        progressBarMain.Value = Math.Min(progressBarMain.Maximum, progressBarMain.Value + 1);
                        labelExecutionStatus.Text = $"{progressBarMain.Value} von {tasks.Count} Aufgaben erledigt";
                    });
                    await Task.Delay(300);
                }

                labelExecutionStatus.Text = "Alle Schritte abgeschlossen.";
                AppendLog("Der Vorgang wurde erfolgreich abgeschlossen.");
            }
            finally
            {
                isExecuting = false;
                currentPage = 10;
                UpdatePage();
                Application.DoEvents();
                this.Refresh();
            }
        }

        private void SavePageSelections()
        {
            switch (currentPage)
            {
                case 1:
                    newComputerName = textBoxComputerName.Text.Trim();
                    break;
                case 2:
                    disableFastStartup = checkBoxDisableFastStartup.Checked;
                    break;
                case 3:
                    disableWindowsUpdates = checkBoxDisableWindowsUpdates.Checked;
                    break;
                case 4:
                    enableNumLock = checkBoxEnableNumLock.Checked;
                    break;
                case 5:
                    desktopIcons["🖥️ Dieser PC"] = checkBoxIconThisPC.Checked;
                    desktopIcons["🗑️ Papierkorb"] = checkBoxIconRecycleBin.Checked;
                    desktopIcons["📁 Benutzerordner"] = checkBoxIconNetwork.Checked;
                    desktopIcons["⚙️ Systemsteuerung"] = checkBoxIconControlPanel.Checked;
                    break;
                case 6:
                    foreach (var item in localInstallerCheckboxes)
                    {
                        localInstallerSelections[item.Key] = item.Value.Checked;
                    }
                    break;
            }
        }

        private void LoadPageSelections()
        {
            switch (currentPage)
            {
                case 1:
                    textBoxComputerName.Text = newComputerName;
                    break;
                case 2:
                    checkBoxDisableFastStartup.Checked = disableFastStartup;
                    break;
                case 3:
                    checkBoxDisableWindowsUpdates.Checked = disableWindowsUpdates;
                    break;
                case 4:
                    checkBoxEnableNumLock.Checked = enableNumLock;
                    break;
                case 5:
                    checkBoxIconThisPC.Checked = desktopIcons["🖥️ Dieser PC"];
                    checkBoxIconRecycleBin.Checked = desktopIcons["🗑️ Papierkorb"];
                    checkBoxIconNetwork.Checked = desktopIcons["📁 Benutzerordner"];
                    checkBoxIconControlPanel.Checked = desktopIcons["⚙️ Systemsteuerung"];
                    break;
                case 6:
                    foreach (var item in localInstallerCheckboxes)
                    {
                        item.Value.Checked = localInstallerSelections.TryGetValue(item.Key, out var selected) && selected;
                    }
                    break;
            }
        }

        private void UpdatePage()
        {
            panelWelcome.Visible = currentPage == 0;
            panelComputerName.Visible = currentPage == 1;
            panelFastStartup.Visible = currentPage == 2;
            panelWindowsUpdates.Visible = currentPage == 3;
            panelNumLock.Visible = currentPage == 4;
            panelDesktopIcons.Visible = currentPage == 5;
            panelApps.Visible = currentPage == 6;
            panelAutostart.Visible = currentPage == 7;
            panelSummary.Visible = currentPage == 8;
            panelExecution.Visible = currentPage == 9;
            panelRestart.Visible = currentPage == 10;

            // TabControl synchronisieren
            tabControl.Visible = currentPage >= 1 && currentPage <= 8;
            if (currentPage >= 1 && currentPage <= 8)
            {
                tabControl.SelectedIndex = currentPage - 1;
            }

            buttonStart.Visible = currentPage == 0;
            buttonBack.Visible = currentPage > 0 && currentPage < 9;
            buttonNext.Visible = currentPage >= 1 && currentPage < 8;
            buttonRun.Visible = currentPage == 8;
            buttonRestart.Visible = currentPage == 10;
            buttonFinish.Visible = currentPage == 10;
            labelCredit.Visible = currentPage == 0;
            pictureBoxPhoto.Visible = currentPage == 0;

            if (currentPage == 0)
            {
                labelPageInfo.Text = "Hallo, ich bin das Master Vorinstallations Tool";
            }
            else if (currentPage == 1)
            {
                labelPageInfo.Text = "1) Computername ändern";
            }
            else if (currentPage == 2)
            {
                labelPageInfo.Text = "2) Schnellstart";
            }
            else if (currentPage == 3)
            {
                labelPageInfo.Text = "3) Windows Updates";
            }
            else if (currentPage == 4)
            {
                labelPageInfo.Text = "4) NumLock";
            }
            else if (currentPage == 5)
            {
                labelPageInfo.Text = "5) Desktop-Icons";
            }
            else if (currentPage == 6)
            {
                labelPageInfo.Text = "6) Programme";
            }
            else if (currentPage == 7)
            {
                labelPageInfo.Text = "7) Autostart";
            }
            else if (currentPage == 8)
            {
                labelPageInfo.Text = "8) Zusammenfassung";
            }
            else if (currentPage == 9)
            {
                labelPageInfo.Text = "Ausführung";
            }
            else if (currentPage == 10)
            {
                labelPageInfo.Text = "Fertig";
            }

            if (currentPage == 6)
            {
                LoadLocalInstallers();
            }

            if (currentPage == 7)
            {
                LoadAutostartEntries();
            }

            LoadPageSelections();
            UpdateSummaryText();
        }

        private void UpdateSummaryText()
        {
            if (currentPage != 8)
                return;

            var selectedIcons = desktopIcons.Where(p => p.Value).Select(p => p.Key).ToList();
            var selectedLocalInstallers = localInstallerSelections
                .Where(p => p.Value)
                .Select(p => Path.GetFileNameWithoutExtension(p.Key))
                .ToList();

            textBoxSummary.Clear();

            void Append(string text, Color color)
            {
                textBoxSummary.SelectionStart = textBoxSummary.TextLength;
                textBoxSummary.SelectionLength = 0;
                textBoxSummary.SelectionColor = color;
                textBoxSummary.AppendText(text);
            }

            var fg = SystemColors.WindowText;

            Append($"Computername: {(!string.IsNullOrEmpty(newComputerName) ? newComputerName : "(Nicht geändert)")}\n", fg);
            Append($"Schnellstart deaktivieren: {(disableFastStartup ? "Ja" : "Nein")}\n", fg);
            Append($"Windows Updates deaktivieren: {(disableWindowsUpdates ? "Ja" : "Nein")}\n", fg);
            Append($"NumLock aktivieren: {(enableNumLock ? "Ja" : "Nein")}\n\n", fg);

            Append("Desktop-Icons:\n", fg);
            if (selectedIcons.Any())
                foreach (var icon in selectedIcons)
                    Append("  " + icon + "\n", fg);
            else
                Append("  Keine Desktop-Icons ausgewählt\n", fg);

            Append("\nLokale Installer:\n", fg);
            if (selectedLocalInstallers.Any())
                foreach (var inst in selectedLocalInstallers)
                    Append("  " + inst + "\n", fg);
            else
                Append("  Keine lokalen Installer ausgewählt\n", fg);

            Append("\nAutostart:\n", fg);
            var changedAutostartEntries = dataGridViewAutostart.Rows
                .Cast<DataGridViewRow>()
                .Where(r => r.Tag is StartupEntry)
                .Select(r => (StartupEntry)r.Tag!)
                .Where(e => e.IsEnabled != e.InitialIsEnabled)
                .OrderBy(e => e.IsEnabled ? 0 : 1)
                .ThenBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (changedAutostartEntries.Any())
            {
                foreach (var entry in changedAutostartEntries)
                {
                    Color entryColor = entry.IsEnabled
                        ? Color.FromArgb(33, 150, 243)
                        : Color.FromArgb(130, 130, 130);
                    string status = entry.IsEnabled ? "● AN  " : "● AUS ";
                    Append("  " + status, entryColor);
                    Append(entry.Name + "\n", fg);
                }
            }
            else
            {
                Append("  Keine geänderten Autostart-Einträge\n", fg);
            }
        }

        private List<SetupTask> BuildSetupTasks()
        {
            var tasks = new List<SetupTask>();

            if (!string.IsNullOrEmpty(newComputerName))
            {
                tasks.Add(new SetupTask($"Computername ändern in: {newComputerName}", async () => await RenameComputerAsync(newComputerName)));
            }
            else
            {
                tasks.Add(new SetupTask("Computername beibehalten", () => Task.CompletedTask));
            }

            if (disableFastStartup)
            {
                tasks.Add(new SetupTask("Schnellstart deaktivieren", async () => await DisableFastStartupAsync()));
            }
            else
            {
                tasks.Add(new SetupTask("Schnellstart beibehalten", () => Task.CompletedTask));
            }

            if (disableWindowsUpdates)
            {
                tasks.Add(new SetupTask("Windows Updates deaktivieren", async () => await DisableWindowsUpdatesAsync()));
            }
            else
            {
                tasks.Add(new SetupTask("Windows Updates beibehalten", () => Task.CompletedTask));
            }

            if (enableNumLock)
            {
                tasks.Add(new SetupTask("NumLock aktivieren", async () => await SetNumLockAsync()));
            }
            else
            {
                tasks.Add(new SetupTask("NumLock beibehalten", () => Task.CompletedTask));
            }

            if (desktopIcons.Values.Any(v => v))
            {
                tasks.Add(new SetupTask("Desktop-Icons aktivieren", async () => await ApplyDesktopIconsAsync()));
            }
            else
            {
                tasks.Add(new SetupTask("Keine Desktop-Icons aktivieren", () => Task.CompletedTask));
            }

            var selectedLocalInstallers = localInstallerSelections.Where(p => p.Value).Select(p => p.Key).ToList();
            if (selectedLocalInstallers.Any())
            {
                foreach (var installerName in selectedLocalInstallers)
                {
                    tasks.Add(new SetupTask($"Lokalen Installer starten: {Path.GetFileNameWithoutExtension(installerName)}", async () => await RunLocalInstallerAsync(installerName)));
                }
            }
            else
            {
                tasks.Add(new SetupTask("Keine lokalen Installer starten", () => Task.CompletedTask));
            }

            return tasks;
        }

        private void InitializeLocalInstallersUi()
        {
            checkBoxAppOffice.Visible = false;
            checkBoxAppEset.Visible = false;
            checkBoxAppChrome.Visible = false;
            checkBoxAppFoxit.Visible = false;
            checkBoxAppAcrobat.Visible = false;
            checkBoxAppVeeam.Visible = false;
            labelApps.Text = "Programme zum Installieren";

            labelLocalInstallers = new Label
            {
                AutoSize = true,
                Font = new Font("Segoe UI", 10F, FontStyle.Regular, GraphicsUnit.Point),
                Location = new Point(32, 60),
                Name = "labelLocalInstallers",
                Text = "Wähle Programme aus (nebeneinander):"
            };

            panelLocalInstallers = new FlowLayoutPanel
            {
                AutoScroll = false,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                Location = new Point(32, 85),
                Name = "panelLocalInstallers",
                Size = new Size(530, 200)
            };

            panelApps.Controls.Add(labelLocalInstallers);
            panelApps.Controls.Add(panelLocalInstallers);
            labelLocalInstallers.BringToFront();
            panelLocalInstallers.BringToFront();
        }

        private void InitializeAutostartUi()
        {
            dataGridViewAutostart.AutoGenerateColumns = false;
            dataGridViewAutostart.Columns.Clear();

            var enabledColumn = new DataGridViewTextBoxColumn
            {
                Name = "Enabled",
                HeaderText = "Status",
                Width = 70,
                ReadOnly = true
            };

            var nameColumn = new DataGridViewTextBoxColumn
            {
                Name = "Name",
                HeaderText = "Name",
                Width = 220,
                ReadOnly = true,
                DataPropertyName = nameof(StartupEntry.Name)
            };

            var commandColumn = new DataGridViewTextBoxColumn
            {
                Name = "Command",
                HeaderText = "Pfad / Befehl",
                AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill,
                ReadOnly = true,
                DataPropertyName = nameof(StartupEntry.Command)
            };

            dataGridViewAutostart.Columns.Add(enabledColumn);
            dataGridViewAutostart.Columns.Add(nameColumn);
            dataGridViewAutostart.Columns.Add(commandColumn);
            dataGridViewAutostart.CellPainting += dataGridViewAutostart_CellPainting;
            dataGridViewAutostart.CellClick += dataGridViewAutostart_CellClick;
            dataGridViewAutostart.Cursor = Cursors.Default;
            buttonRefreshAutostart.Click += buttonRefreshAutostart_Click;
        }

        private void buttonRefreshAutostart_Click(object sender, EventArgs e)
        {
            LoadAutostartEntries();
        }

        private void dataGridViewAutostart_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex != 0)
            {
                return;
            }

            e.PaintBackground(e.ClipBounds, true);

            var row = dataGridViewAutostart.Rows[e.RowIndex];
            bool isEnabled = row.Tag is StartupEntry paintEntry && paintEntry.IsEnabled;
            bool isReadOnly = row.Tag is StartupEntry roEntry && !roEntry.Hive.HasValue;

            int switchWidth = 46;
            int switchHeight = 22;
            int x = e.CellBounds.X + (e.CellBounds.Width - switchWidth) / 2;
            int y = e.CellBounds.Y + (e.CellBounds.Height - switchHeight) / 2;

            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            Color bgColor = isReadOnly
                ? Color.FromArgb(200, 200, 200)
                : (isEnabled ? Color.FromArgb(33, 150, 243) : Color.FromArgb(160, 160, 160));

            int radius = switchHeight / 2;
            var rect = new Rectangle(x, y, switchWidth, switchHeight);
            using (var bgBrush = new SolidBrush(bgColor))
            {
                DrawRoundedRectangle(g, bgBrush, rect, radius);
            }

            int circleMargin = 2;
            int circleDiameter = switchHeight - circleMargin * 2;
            int circleX = isEnabled ? x + switchWidth - circleDiameter - circleMargin : x + circleMargin;
            using (var circleBrush = new SolidBrush(Color.White))
            {
                g.FillEllipse(circleBrush, circleX, y + circleMargin, circleDiameter, circleDiameter);
            }

            e.Handled = true;
        }

        private static void DrawRoundedRectangle(Graphics g, Brush brush, Rectangle rect, int radius)
        {
            using var path = new System.Drawing.Drawing2D.GraphicsPath();
            path.AddArc(rect.X, rect.Y, radius * 2, radius * 2, 180, 90);
            path.AddArc(rect.Right - radius * 2, rect.Y, radius * 2, radius * 2, 270, 90);
            path.AddArc(rect.Right - radius * 2, rect.Bottom - radius * 2, radius * 2, radius * 2, 0, 90);
            path.AddArc(rect.X, rect.Bottom - radius * 2, radius * 2, radius * 2, 90, 90);
            path.CloseFigure();
            g.FillPath(brush, path);
        }

        private void dataGridViewAutostart_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex != 0)
            {
                return;
            }

            var row = dataGridViewAutostart.Rows[e.RowIndex];
            if (row.Tag is not StartupEntry entry || !entry.Hive.HasValue)
            {
                return;
            }

            bool newValue = !entry.IsEnabled;
            try
            {
                SetAutostartEntryState(entry, newValue);
                entry.IsEnabled = newValue;
                dataGridViewAutostart.InvalidateCell(0, e.RowIndex);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Autostart-Eintrag konnte nicht geändert werden: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadAutostartEntries()
        {
            try
            {
                dataGridViewAutostart.Rows.Clear();

                foreach (var entry in GetAutostartEntries())
                {
                    string key = BuildStartupEntryKey(entry);
                    if (!initialAutostartStates.TryGetValue(key, out bool initialState))
                    {
                        initialState = entry.IsEnabled;
                        initialAutostartStates[key] = initialState;
                    }

                    entry.InitialIsEnabled = initialState;
                    int rowIndex = dataGridViewAutostart.Rows.Add("", entry.Name, entry.Command);
                    var row = dataGridViewAutostart.Rows[rowIndex];
                    row.Tag = entry;
                    row.Cells[1].ToolTipText = entry.Location;
                    row.Cells[2].ToolTipText = entry.Command;
                    if (!entry.Hive.HasValue)
                    {
                        row.Cells[0].ToolTipText = "Dieser Eintrag kann nur angezeigt werden.";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Autostart-Einträge konnten nicht geladen werden: {ex.Message}", "Fehler", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private List<StartupEntry> GetAutostartEntries()
        {
            var entries = new Dictionary<string, StartupEntry>(StringComparer.OrdinalIgnoreCase);

            foreach (var registryEntry in GetRegistryAutostartEntries(RegistryHive.CurrentUser))
            {
                entries[BuildStartupEntryKey(registryEntry)] = registryEntry;
            }

            foreach (var registryEntry in GetRegistryAutostartEntries(RegistryHive.LocalMachine))
            {
                entries[BuildStartupEntryKey(registryEntry)] = registryEntry;
            }

            foreach (var wmiEntry in GetWmiAutostartEntries())
            {
                string key = BuildStartupEntryKey(wmiEntry);
                if (!entries.ContainsKey(key))
                {
                    entries[key] = wmiEntry;
                }
            }

            return entries.Values
                .OrderBy(entry => entry.Name, StringComparer.OrdinalIgnoreCase)
                .ThenBy(entry => entry.Location, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static string BuildStartupEntryKey(StartupEntry entry)
        {
            string hiveName = entry.Hive?.ToString() ?? "Wmi";
            return $"{hiveName}|{entry.Name}|{entry.Command}";
        }

        private IEnumerable<StartupEntry> GetRegistryAutostartEntries(RegistryHive hive)
        {
            const string runPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
            using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default);
            using RegistryKey runKey = baseKey.OpenSubKey(runPath, false);
            if (runKey == null)
            {
                yield break;
            }

            foreach (string valueName in runKey.GetValueNames())
            {
                string command = runKey.GetValue(valueName)?.ToString() ?? string.Empty;
                yield return new StartupEntry
                {
                    Name = valueName,
                    Command = command,
                    Location = $"{hive}\\{runPath}",
                    Hive = hive,
                    IsEnabled = GetAutostartApprovedState(hive, valueName)
                };
            }
        }

        private IEnumerable<StartupEntry> GetWmiAutostartEntries()
        {
            using var searcher = new ManagementObjectSearcher("SELECT Name, Command, Location FROM Win32_StartupCommand");
            foreach (ManagementObject item in searcher.Get())
            {
                string name = item["Name"]?.ToString() ?? string.Empty;
                string command = item["Command"]?.ToString() ?? string.Empty;
                string location = item["Location"]?.ToString() ?? "WMI";
                RegistryHive? hive = ParseHiveFromLocation(location);
                yield return new StartupEntry
                {
                    Name = name,
                    Command = command,
                    Location = location,
                    Hive = hive,
                    IsEnabled = hive.HasValue ? GetAutostartApprovedState(hive.Value, name) : true
                };
            }
        }

        private static RegistryHive? ParseHiveFromLocation(string location)
        {
            if (location.Contains("HKCU", StringComparison.OrdinalIgnoreCase) ||
                location.Contains("HKEY_CURRENT_USER", StringComparison.OrdinalIgnoreCase))
            {
                return RegistryHive.CurrentUser;
            }

            if (location.Contains("HKLM", StringComparison.OrdinalIgnoreCase) ||
                location.Contains("HKEY_LOCAL_MACHINE", StringComparison.OrdinalIgnoreCase))
            {
                return RegistryHive.LocalMachine;
            }

            return null;
        }

        private static bool GetAutostartApprovedState(RegistryHive hive, string name)
        {
            const string approvedPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run";
            using RegistryKey baseKey = RegistryKey.OpenBaseKey(hive, RegistryView.Default);
            using RegistryKey approvedKey = baseKey.OpenSubKey(approvedPath, false);
            byte[] value = approvedKey?.GetValue(name) as byte[];
            if (value == null || value.Length == 0)
            {
                return true;
            }

            return value[0] != 2;
        }

        private static void SetAutostartEntryState(StartupEntry entry, bool enabled)
        {
            if (!entry.Hive.HasValue)
            {
                throw new InvalidOperationException("Dieser Autostart-Eintrag kann nicht direkt umgeschaltet werden.");
            }

            const string approvedPath = @"Software\Microsoft\Windows\CurrentVersion\Explorer\StartupApproved\Run";
            byte[] data = enabled
                ? new byte[] { 3, 0, 0, 0, 0, 0, 0, 0 }
                : new byte[] { 2, 0, 0, 0, 0, 0, 0, 0 };

            using RegistryKey baseKey = RegistryKey.OpenBaseKey(entry.Hive.Value, RegistryView.Default);
            using RegistryKey approvedKey = baseKey.CreateSubKey(approvedPath, true);
            if (approvedKey == null)
            {
                throw new InvalidOperationException("StartupApproved-Registry-Pfad konnte nicht geöffnet werden.");
            }

            approvedKey.SetValue(entry.Name, data, RegistryValueKind.Binary);
        }

        private static List<string> GetEmbeddedInstallerFileNames()
        {
            var names = Assembly.GetExecutingAssembly()
                .GetManifestResourceNames()
                .Where(name => name.StartsWith(EmbeddedInstallerPrefix, StringComparison.OrdinalIgnoreCase) &&
                               name.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                .Select(name => name.Substring(EmbeddedInstallerPrefix.Length))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (!names.Any(n => n.Equals(DesktopCopyProgramName, StringComparison.OrdinalIgnoreCase)))
            {
                names.Add(DesktopCopyProgramName);
            }

            return names.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private void LoadLocalInstallers()
        {
            if (panelLocalInstallers == null)
            {
                return;
            }

            panelLocalInstallers.Controls.Clear();
            localInstallerCheckboxes.Clear();

            var installerFiles = GetEmbeddedInstallerFileNames();
            if (!installerFiles.Any())
            {
                panelLocalInstallers.Controls.Add(new Label { AutoSize = true, Text = "Keine eingebetteten EXE-Dateien gefunden." });
                return;
            }

            foreach (var fileName in installerFiles)
            {
                string displayName = Path.GetFileNameWithoutExtension(fileName);
                if (!localInstallerSelections.ContainsKey(fileName))
                {
                    localInstallerSelections[fileName] = false;
                }

                var checkBox = new CheckBox
                {
                    AutoSize = false,
                    Width = 170,
                    Height = 24,
                    Text = displayName,
                    Checked = localInstallerSelections[fileName],
                    Margin = new Padding(3, 3, 3, 3)
                };

                checkBox.CheckedChanged += (_, _) =>
                {
                    localInstallerSelections[fileName] = checkBox.Checked;
                };

                localInstallerCheckboxes[fileName] = checkBox;
                panelLocalInstallers.Controls.Add(checkBox);
            }
        }

        private static string ExtractEmbeddedInstallerToTemp(string installerFileName)
        {
            string resourceName = EmbeddedInstallerPrefix + installerFileName;
            using Stream resourceStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName);
            if (resourceStream == null)
            {
                throw new Exception($"Eingebettete Datei nicht gefunden: {installerFileName}");
            }

            string tempDir = Path.Combine(Path.GetTempPath(), "SetupTool", "Installers");
            Directory.CreateDirectory(tempDir);

            string tempFilePath = Path.Combine(tempDir, installerFileName);
            using FileStream fileStream = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write, FileShare.None);
            resourceStream.CopyTo(fileStream);

            return tempFilePath;
        }

        private static string ResolveInstallerSourceForDesktopCopy(string installerFileName)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string directPath = Path.Combine(baseDir, installerFileName);
            if (File.Exists(directPath))
            {
                return directPath;
            }

            string installPath = Path.Combine(baseDir, "Install", installerFileName);
            if (File.Exists(installPath))
            {
                return installPath;
            }

            return ExtractEmbeddedInstallerToTemp(installerFileName);
        }

        private static void CopyProgramToDesktop(string installerFileName)
        {
            string sourcePath = ResolveInstallerSourceForDesktopCopy(installerFileName);
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string destinationPath = Path.Combine(desktopPath, installerFileName);
            File.Copy(sourcePath, destinationPath, true);
        }

        private static string GetDesktopIconGuid(string iconName)
        {
            return iconName switch
            {
                "🖥️ Dieser PC" => "{20D04FE0-3AEA-1069-A2D8-08002B30309D}",
                "📁 Benutzerordner" => "{59031A47-3F72-44A7-89C5-5595FE6B30EE}",
                "⚙️ Systemsteuerung" => "{5399E694-6CE5-4D6C-8FCE-1D8870FDCBA0}",
                "🗑️ Papierkorb" => "{645FF040-5081-101B-9F08-00AA002F954E}",
                _ => string.Empty,
            };
        }

        private async Task ApplyDesktopIconsAsync()
        {
            await Task.Run(() =>
            {
                try
                {
                    const string subKey = @"Software\Microsoft\Windows\CurrentVersion\Explorer\HideDesktopIcons\NewStartPanel";
                    using var key = Registry.CurrentUser.CreateSubKey(subKey, true);
                    if (key == null)
                    {
                        throw new Exception("Der Registry-Pfad für Desktop-Icons wurde nicht gefunden.");
                    }

                    foreach (var icon in desktopIcons.Where(i => i.Value))
                    {
                        string guid = GetDesktopIconGuid(icon.Key);
                        if (!string.IsNullOrWhiteSpace(guid))
                        {
                            key.SetValue(guid, 0, RegistryValueKind.DWord);
                        }
                    }

                    RestartExplorer();
                }
                catch (Exception ex)
                {
                    throw new Exception($"Fehler beim Aktivieren der Desktop-Icons: {ex.Message}");
                }
            });
        }

        private static void RestartExplorer()
        {
            Process.Start(new ProcessStartInfo("taskkill", "/f /im explorer.exe")
            {
                CreateNoWindow = true,
                UseShellExecute = false
            })?.WaitForExit();

            Process.Start(new ProcessStartInfo("explorer.exe")
            {
                CreateNoWindow = true,
                UseShellExecute = true
            });
        }

        private async Task DisableFastStartupAsync()
        {
            await Task.Run(() =>
            {
                try
                {
                    using var key = Registry.LocalMachine.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Power", true);
                    if (key == null)
                    {
                        throw new Exception("Der Registry-Pfad wurde nicht gefunden.");
                    }

                    var currentValue = key.GetValue("HiberbootEnabled");
                    if (currentValue is int intValue && intValue == 0)
                    {
                        return; // Schnellstart ist bereits deaktiviert.
                    }

                    key.SetValue("HiberbootEnabled", 0, RegistryValueKind.DWord);

                    var verifiedValue = key.GetValue("HiberbootEnabled");
                    if (!(verifiedValue is int verifiedInt) || verifiedInt != 0)
                    {
                        throw new Exception("Der Registry-Wert konnte nicht korrekt auf 0 gesetzt werden.");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Fehler beim Deaktivieren des Schnellstarts: {ex.Message}");
                }
            });
        }

        private async Task DisableWindowsUpdatesAsync()
        {
            await Task.Run(() =>
            {
                try
                {
                    const string basePath = @"SOFTWARE\Policies\Microsoft\Windows";

                    // Öffne oder erstelle den Basispfad
                    using var baseKey = Registry.LocalMachine.OpenSubKey(basePath, true) ?? Registry.LocalMachine.CreateSubKey(basePath, true);
                    if (baseKey == null)
                    {
                        throw new Exception($"Konnte Registry-Pfad '{basePath}' nicht öffnen oder erstellen.");
                    }

                    // Erstelle den WindowsUpdate-Schlüssel
                    using var updateKey = baseKey.OpenSubKey("WindowsUpdate") ?? baseKey.CreateSubKey("WindowsUpdate", true);
                    if (updateKey == null)
                    {
                        throw new Exception("Konnte Windows Update Registry-Schlüssel nicht erstellen.");
                    }

                    // Erstelle den AU-Unterschlüssel
                    using var auKey = updateKey.OpenSubKey("AU") ?? updateKey.CreateSubKey("AU", true);
                    if (auKey == null)
                    {
                        throw new Exception("Konnte AU Registry-Schlüssel nicht erstellen.");
                    }

                    // Setze NoAutoUpdate auf 1 (deaktiviert)
                    auKey.SetValue("NoAutoUpdate", 1, RegistryValueKind.DWord);

                    // Verifiziere, dass der Wert gesetzt wurde
                    var verifiedValue = auKey.GetValue("NoAutoUpdate");
                    if (!(verifiedValue is int verifiedInt) || verifiedInt != 1)
                    {
                        throw new Exception("Der NoAutoUpdate-Wert konnte nicht korrekt auf 1 gesetzt werden.");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Fehler beim Deaktivieren von Windows Updates: {ex.Message}");
                }
            });
        }

        private async Task RenameComputerAsync(string newName)
        {
            await Task.Run(() =>
            {
                try
                {
                    if (string.IsNullOrWhiteSpace(newName))
                    {
                        return; // Kein neuer Name angegeben
                    }

                    // Führe PowerShell-Befehl aus: Rename-Computer -NewName "NEUER-NAME"
                    var psi = new ProcessStartInfo("powershell.exe")
                    {
                        Arguments = $"-NoProfile -Command \"Rename-Computer -NewName '{newName}' -Force\"",
                        CreateNoWindow = true,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        Verb = "runas" // Mit Administrator-Rechten ausführen
                    };

                    using var process = Process.Start(psi);
                    if (process == null)
                    {
                        throw new Exception("PowerShell-Prozess konnte nicht gestartet werden.");
                    }

                    process.WaitForExit(30000); // Max 30 Sekunden warten

                    // Überprüfe Exit Code
                    if (process.ExitCode != 0)
                    {
                        string errorOutput = process.StandardError.ReadToEnd();
                        string output = process.StandardOutput.ReadToEnd();
                        string errorMsg = !string.IsNullOrEmpty(errorOutput) ? errorOutput : output;
                        throw new Exception($"PowerShell-Fehler: {errorMsg}");
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Fehler beim Ändern des Computernamens: {ex.Message}");
                }
            });
        }

        private async Task SetNumLockAsync()
        {
            await Task.Run(() =>
            {
                try
                {
                    const string subKey = ".DEFAULT\\Control Panel\\Keyboard";
                    const string valueName = "InitialKeyboardIndicators";

                    using var key = Registry.Users.OpenSubKey(subKey, true);
                    if (key == null)
                    {
                        throw new Exception("Der Registry-Pfad wurde nicht gefunden.");
                    }

                    var currentValueObj = key.GetValue(valueName);
                    if (currentValueObj == null)
                    {
                        throw new Exception($"Der Wert '{valueName}' wurde nicht gefunden.");
                    }

                    var rawValue = currentValueObj.ToString();
                    if (string.IsNullOrWhiteSpace(rawValue))
                    {
                        throw new Exception($"Der Wert '{valueName}' ist leer.");
                    }

                    string newValue = null;
                    if (rawValue == "2147483648")
                    {
                        newValue = "2147483650";
                    }
                    else if (rawValue == "2147483650" || rawValue == "2")
                    {
                        return; // NumLock ist bereits aktiviert.
                    }
                    else if (rawValue == "0")
                    {
                        newValue = "2";
                    }
                    else
                    {
                        newValue = "2"; // Unbekannter Wert, auf Standard 2 setzen.
                    }

                    if (!string.IsNullOrEmpty(newValue))
                    {
                        key.SetValue(valueName, newValue, RegistryValueKind.String);
                        var verifiedValueObj = key.GetValue(valueName);
                        var verifiedValue = verifiedValueObj?.ToString() ?? string.Empty;
                        if (verifiedValue != newValue)
                        {
                            throw new Exception($"Der Wert konnte nicht korrekt gesetzt werden. Aktueller Wert: {verifiedValue}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Fehler beim Aktivieren von NumLock: {ex.Message}");
                }
            });
        }

        private async Task RunLocalInstallerAsync(string installerFileName)
        {
            await Task.Run(() =>
            {
                try
                {
                    if (installerFileName.Equals(DesktopCopyProgramName, StringComparison.OrdinalIgnoreCase))
                    {
                        CopyProgramToDesktop(installerFileName);
                        return;
                    }

                    string installerPath = ExtractEmbeddedInstallerToTemp(installerFileName);
                    string processName = Path.GetFileNameWithoutExtension(installerPath);

                    var startInfo = new ProcessStartInfo(installerPath)
                    {
                        UseShellExecute = true,
                        Verb = "runas",
                        WorkingDirectory = Path.GetDirectoryName(installerPath) ?? AppDomain.CurrentDomain.BaseDirectory
                    };

                    using var process = Process.Start(startInfo);
                    process?.WaitForExit();

                    // Nach Beenden versuchen, das Fenster ggf. zu schließen (z.B. falls Setup offen bleibt)
                    try
                    {
                        foreach (var proc in Process.GetProcessesByName(processName))
                        {
                            if (!proc.HasExited)
                            {
                                proc.Kill();
                                proc.WaitForExit(5000);
                            }
                        }
                    }
                    catch { /* Ignorieren, falls kein Prozess mehr offen */ }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Fehler beim Starten von {installerFileName}: {ex.Message}");
                }
            });
        }

        private void AppendLog(string message)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>(AppendLog), message);
                return;
            }

            textBoxLog.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}");
        }

        private sealed class SetupTask
        {
            public string Name { get; }
            public Func<Task> Action { get; }

            public SetupTask(string name, Func<Task> action)
            {
                Name = name;
                Action = action;
            }
        }
    }
}
