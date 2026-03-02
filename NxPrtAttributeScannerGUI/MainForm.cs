using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

public class MainForm : Form
{
    TextBox tbNxBaseDir;
    Button btnPickNx;
    TextBox tbRoot;
    Button btnPickRoot;

    TextBox tbFilters;
    //CheckBox cbGroupSheets;
    ComboBox cmbGroupMode;

    ComboBox cmbMode;

    TextBox tbExcelOut;
    Button btnPickExcel;

    Button btnStart;
    Button btnStop;

    RichTextBox tbLog;
    Label lblStatus;

    Process _proc;

    public MainForm()
    {
        Text = "Сканер атрибутов NX PRT → Excel";
        Width = 980;
        Height = 720;
        StartPosition = FormStartPosition.CenterScreen;

        InitializeUi();
        LoadDefaults();
    }

    void InitializeUi()
    {
        Padding = new Padding(10);

        // ===== Split: сверху настройки+кнопки, снизу лог =====
        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            FixedPanel = FixedPanel.Panel1,
            IsSplitterFixed = false,
            Panel1MinSize = 260,
            Panel2MinSize = 150
        };

        Controls.Add(split);

        // Чтобы по умолчанию лог был больше
        split.FixedPanel = FixedPanel.Panel1;
        split.IsSplitterFixed = false;
        split.SplitterDistance = 300;

        // ===== Верхняя панель: 2 строки (настройки + кнопки) =====
        var topLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(0)
        };
        topLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // настройки
        topLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // кнопки

        split.Panel1.Controls.Add(topLayout);

        // ===== Таблица настроек =====
        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 0,
            AutoSize = false,
            Padding = new Padding(0)
        };

        // Колонки: Label | Input | Button
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        topLayout.Controls.Add(grid, 0, 0);

        // ===== Row: NX =====
        tbNxBaseDir = new TextBox { Dock = DockStyle.Fill };
        btnPickNx = new Button { Text = "Выбрать…", AutoSize = true };
        btnPickNx.Click += (s, e) => PickNxFolder();
        AddRow(grid, "Папка NX (UGII_BASE_DIR):", tbNxBaseDir, btnPickNx);

        // ===== Row: Root =====
        tbRoot = new TextBox { Dock = DockStyle.Fill };
        btnPickRoot = new Button { Text = "Выбрать…", AutoSize = true };
        btnPickRoot.Click += (s, e) => PickRootFolder();
        AddRow(grid, "Корневая папка (Root):", tbRoot, btnPickRoot);

        // ===== Row: Excel Out =====
        tbExcelOut = new TextBox { Dock = DockStyle.Fill };
        btnPickExcel = new Button { Text = "Выбрать…", AutoSize = true };
        btnPickExcel.Click += (s, e) => PickExcelOut();
        AddRow(grid, "Файл Excel (Out):", tbExcelOut, btnPickExcel);

        // ===== Параметры (2 строки, без чекбокса) =====
        var paramsGrid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0)
        };
        paramsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        paramsGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        paramsGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        paramsGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var lblGroup = new Label
        {
            Text = "Группировка:",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 6, 10, 0)
        };

        cmbGroupMode = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 3, 0, 3)
        };
        cmbGroupMode.Items.Add("Первый уровень (FirstLevel)");
        cmbGroupMode.Items.Add("Всё в одном (AllInOne)");
        cmbGroupMode.SelectedIndex = 0;

        paramsGrid.Controls.Add(lblGroup, 0, 0);
        paramsGrid.Controls.Add(cmbGroupMode, 1, 0);

        var lblMode = new Label
        {
            Text = "Режим:",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 6, 10, 0)
        };

        cmbMode = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 3, 0, 3)
        };
        cmbMode.Items.Add("Сканировать и выгрузить в Excel (ScanAndExport)");
        cmbMode.Items.Add("Только сканировать (ScanOnly)");
        cmbMode.SelectedIndex = 0;

        paramsGrid.Controls.Add(lblMode, 0, 1);
        paramsGrid.Controls.Add(cmbMode, 1, 1);

        AddRow(grid, "Параметры:", paramsGrid, null);

        // ===== Фильтры: делаем строку растягиваемой =====
        // 1) строка-лейбл (обычная)
        AddRow(grid, "Фильтр папок (каждая с новой строки или через ;):", new Panel { Height = 1 }, null);

        // 2) строка-текстбокс (растягивается)
        tbFilters = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            MinimumSize = new System.Drawing.Size(0, 70)
        };
        AddRow(grid, "", tbFilters, null, makeRowFill: true);

        // ===== Нижняя панель с кнопками (в отдельной строке topLayout) =====
        var bottomBar = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            AutoSize = true,
            Padding = new Padding(0, 6, 0, 0)
        };
        bottomBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottomBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bottomBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        btnStart = new Button { Text = "Запустить", AutoSize = true };
        btnStart.Click += (s, e) => StartRun();

        btnStop = new Button { Text = "Остановить", AutoSize = true, Enabled = false };
        btnStop.Click += (s, e) => StopRun();

        lblStatus = new Label
        {
            Text = "Готово.",
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Padding = new Padding(8, 6, 0, 0)
        };

        bottomBar.Controls.Add(btnStart, 0, 0);
        bottomBar.Controls.Add(btnStop, 1, 0);
        bottomBar.Controls.Add(lblStatus, 2, 0);

        topLayout.Controls.Add(bottomBar, 0, 1);

        // ===== Log (нижняя панель split) =====
        tbLog = new RichTextBox
        {
            Dock = DockStyle.Fill,
            ReadOnly = true
        };
        split.Panel2.Controls.Add(tbLog);

        // При закрытии — убить процесс, если он жив
        FormClosing += (s, e) =>
        {
            try { if (_proc != null && !_proc.HasExited) _proc.Kill(); } catch { }
        };
    }

    void LoadDefaults()
    {
        // NX from env
        tbNxBaseDir.Text = Environment.GetEnvironmentVariable("UGII_BASE_DIR") ?? "";

        // default paths relative to launcher folder
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        tbExcelOut.Text = Path.Combine(baseDir, "reports", "Parts.xlsx");
        tbRoot.Text = ""; // пользователь выберет
        tbFilters.Text = ""; // без фильтра по умолчанию
    }

    void PickNxFolder()
    {
        using (var dlg = new FolderBrowserDialog())
        {
            dlg.Description = "Выберите папку установки NX (например: ...\\NX1899)";
            dlg.SelectedPath = Directory.Exists(tbNxBaseDir.Text) ? tbNxBaseDir.Text : "";
            if (dlg.ShowDialog(this) == DialogResult.OK)
                tbNxBaseDir.Text = dlg.SelectedPath;
        }
    }

    void PickRootFolder()
    {
        using (var dlg = new FolderBrowserDialog())
        {
            dlg.Description = "Выберите корневую папку, где искать *.prt";
            dlg.SelectedPath = Directory.Exists(tbRoot.Text) ? tbRoot.Text : "";
            if (dlg.ShowDialog(this) == DialogResult.OK)
                tbRoot.Text = dlg.SelectedPath;
        }
    }

    void PickExcelOut()
    {
        using (var dlg = new SaveFileDialog())
        {
            dlg.Title = "Куда сохранить Excel";
            dlg.Filter = "Excel (*.xlsx)|*.xlsx";
            dlg.FileName = Path.GetFileName(tbExcelOut.Text);
            dlg.InitialDirectory = Directory.Exists(Path.GetDirectoryName(tbExcelOut.Text))
                ? Path.GetDirectoryName(tbExcelOut.Text)
                : AppDomain.CurrentDomain.BaseDirectory;

            if (dlg.ShowDialog(this) == DialogResult.OK)
                tbExcelOut.Text = dlg.FileName;
        }
    }

    void StartRun()
    {
        if (_proc != null && !_proc.HasExited)
        {
            AppendLog("Уже запущено.\n");
            return;
        }

        string nxBase = (tbNxBaseDir.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(nxBase) || !Directory.Exists(nxBase))
        {
            MessageBox.Show(this, "Не найдена папка NX. Укажите UGII_BASE_DIR вручную.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        string runJournal = Path.Combine(nxBase, "NXBIN", "run_journal.exe");
        if (!File.Exists(runJournal))
        {
            MessageBox.Show(this, "Не найден run_journal.exe:\n" + runJournal, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        string baseDir = AppDomain.CurrentDomain.BaseDirectory;

        string runScannerCs = Path.Combine(baseDir, "RunScanner.cs");
        if (!File.Exists(runScannerCs))
        {
            MessageBox.Show(this, "Не найден файл RunScanner.cs рядом с Launcher.exe.\nПоложите RunScanner.cs в папку программы.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        string scannerDll = Path.Combine(baseDir, "NxPrtAttributeScanner.dll");
        if (!File.Exists(scannerDll))
        {
            MessageBox.Show(this, "Не найден NxPrtAttributeScanner.dll рядом с Launcher.exe.\nПоложите DLL (и зависимости) в папку программы.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        string root = (tbRoot.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
        {
            MessageBox.Show(this, "Укажите существующую корневую папку Root.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        string excelOut = (tbExcelOut.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(excelOut))
        {
            MessageBox.Show(this, "Укажите путь для Excel (Out).", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Создаём папки, чтобы не падать на записи ini/out
        try
        {
            Directory.CreateDirectory(Path.Combine(baseDir, "cache"));
            Directory.CreateDirectory(Path.Combine(baseDir, "reports"));
            Directory.CreateDirectory(Path.GetDirectoryName(excelOut) ?? baseDir);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Не удалось создать папки:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Пишем scan.ini рядом с Launcher.exe
        string iniPath = Path.Combine(baseDir, "scan.ini");
        try
        {
            File.WriteAllText(iniPath, BuildIni(baseDir, root, excelOut), Encoding.UTF8);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Не удалось записать scan.ini:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        AppendLog("=== Запуск ===\n");
        AppendLog("NX: " + nxBase + "\n");
        AppendLog("Root: " + root + "\n");
        AppendLog("INI:  " + iniPath + "\n\n");

        // Команда
        // run_journal.exe "RunScanner.cs" -args config="...\scan.ini" scannerDll="...\NxPrtAttributeScanner.dll"
        string args =
            Quote(runScannerCs) +
            " -args " +
            "config=" + Quote(iniPath) + " " +
            "scannerDll=" + Quote(scannerDll);

        var psi = new ProcessStartInfo
        {
            FileName = runJournal,
            Arguments = args,
            WorkingDirectory = baseDir,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        // Критично для MiniSnap и прочих managed dll: добавим NXBIN и NXBIN\managed в PATH на время запуска
        string nxb = Path.Combine(nxBase, "NXBIN");
        string nxManaged = Path.Combine(nxb, "managed");

        try
        {
            string oldPath = psi.EnvironmentVariables["PATH"] ?? Environment.GetEnvironmentVariable("PATH") ?? "";
            psi.EnvironmentVariables["PATH"] = nxManaged + ";" + nxb + ";" + baseDir + ";" + oldPath;
            psi.EnvironmentVariables["UGII_BASE_DIR"] = nxBase; // на всякий случай
        }
        catch { /* не критично */ }

        _proc = new Process();
        _proc.StartInfo = psi;
        _proc.EnableRaisingEvents = true;

        _proc.OutputDataReceived += (s, e) => { if (e.Data != null) AppendLog(e.Data + "\n"); };
        _proc.ErrorDataReceived += (s, e) => { if (e.Data != null) AppendLog("[ERR] " + e.Data + "\n"); };
        _proc.Exited += (s, e) =>
        {
            BeginInvoke((Action)(() =>
            {
                btnStart.Enabled = true;
                btnStop.Enabled = false;
                lblStatus.Text = "Готово. Код завершения: " + _proc.ExitCode;
                AppendLog("\n=== Завершено. ExitCode=" + _proc.ExitCode + " ===\n");
            }));
        };

        try
        {
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            lblStatus.Text = "Работаю…";

            _proc.Start();
            _proc.BeginOutputReadLine();
            _proc.BeginErrorReadLine();
        }
        catch (Exception ex)
        {
            btnStart.Enabled = true;
            btnStop.Enabled = false;
            lblStatus.Text = "Ошибка запуска.";
            AppendLog("[FATAL] " + ex + "\n");
        }
    }

    void StopRun()
    {
        try
        {
            if (_proc != null && !_proc.HasExited)
            {
                _proc.Kill();
                AppendLog("\n=== Остановлено пользователем ===\n");
            }
        }
        catch (Exception ex)
        {
            AppendLog("[ERR] Не удалось остановить: " + ex.Message + "\n");
        }
    }

    string BuildIni(string baseDir, string root, string excelOut)
    {
        // Важно: относительные пути тут можно писать как .\cache\parts.db и .\reports\Parts.xlsx
        // ScanOptionsLoader должен привязывать их к папке scan.ini (мы это уже обсуждали)
        var sb = new StringBuilder();

        sb.AppendLine("[Scan]");
        sb.AppendLine("Root=" + root);
        sb.AppendLine("Mode=" + (cmbMode.SelectedIndex == 1 ? "ScanOnly" : "ScanAndExport"));
        sb.AppendLine("DbPath=.\\cache\\parts.db");

        // Фильтры: каждая строка или через ;
        string raw = tbFilters.Text ?? "";
        var items = raw.Split(new[] { '\r', '\n', ';' }, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i < items.Length; i++)
        {
            string t = items[i].Trim();
            if (t.Length == 0) continue;
            sb.AppendLine("IncludeFolderName=" + t);
        }

        sb.AppendLine();
        sb.AppendLine("[Excel]");
        sb.AppendLine("Out=" + excelOut);
        //sb.AppendLine("GroupSheets=" + (cbGroupSheets.Checked ? "1" : "0"));
        sb.AppendLine("GroupMode=" + (cmbGroupMode.SelectedIndex == 1 ? "AllInOne" : "FirstLevel"));

        return sb.ToString();
    }

    void AppendLog(string text)
    {
        if (InvokeRequired)
        {
            BeginInvoke((Action)(() => AppendLog(text)));
            return;
        }

        tbLog.AppendText(text);
        tbLog.ScrollToCaret();
    }

    static string Quote(string s)
    {
        if (s == null) return "\"\"";
        if (s.Contains("\"")) s = s.Replace("\"", "\\\"");
        return "\"" + s + "\"";
    }
    private static void AddRow(
    TableLayoutPanel grid,
    string labelText,
    Control input,
    Control button,
    bool makeRowFill = false)
    {
        int row = grid.RowCount;
        grid.RowCount++;

        // Стиль строки: обычные строки AutoSize,
        // а одна "растягиваемая" строка — Percent 100%
        grid.RowStyles.Add(makeRowFill
            ? new RowStyle(SizeType.Percent, 100f)
            : new RowStyle(SizeType.AutoSize));

        var lbl = new Label
        {
            Text = labelText,
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 6, 10, 0)
        };

        // input
        input.Margin = new Padding(0, 3, 10, 3);
        input.Anchor = AnchorStyles.Left | AnchorStyles.Right;

        grid.Controls.Add(lbl, 0, row);
        grid.Controls.Add(input, 1, row);

        if (button != null)
        {
            button.Margin = new Padding(0, 2, 0, 2);
            button.Anchor = AnchorStyles.Left;
            grid.Controls.Add(button, 2, row);
        }
        else
        {
            // пустая ячейка для выравнивания (не обязательно)
            grid.Controls.Add(new Panel { Width = 1, Height = 1 }, 2, row);
        }
    }

    static Label MkLabel(string text, int x, int y, int w)
    {
        return new Label { Left = x, Top = y + 4, Width = w, Text = text };
    }

    static TextBox MkTextBox(int x, int y, int w)
    {
        return new TextBox { Left = x, Top = y, Width = w };
    }

    static Button MkButton(string text, int x, int y, int w)
    {
        return new Button { Left = x, Top = y, Width = w, Text = text };
    }
}