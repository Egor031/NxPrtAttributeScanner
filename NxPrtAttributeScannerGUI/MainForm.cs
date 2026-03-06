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
    // ===== DB / листы =====
    TextBox tbDbPath;
    Button btnPickDb;
    Button btnNewDb;

    SplitContainer splitFolders;    // внутри верхней панели
    TreeView tvFolders;

    ListBox lbSheets;
    ListBox lbSheetFolders;

    Button btnSheetAdd;
    Button btnSheetRename;
    Button btnSheetDelete;
    Button btnAddFolderToSheet;
    Button btnRemoveFolderFromSheet;

    CacheRepository _repo;
    string _rootFromDb;
    int _selectedSheetId = -1;

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

        // ===== Split: сверху UI, снизу лог =====
        var split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            FixedPanel = FixedPanel.Panel1,
            IsSplitterFixed = false
            // ВАЖНО: НЕ ставим Panel1MinSize/Panel2MinSize здесь
        };
        Controls.Add(split);

        // ===== Верхняя панель: настройки + кнопки =====
        var topLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Padding = new Padding(0)
        };
        topLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        topLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
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
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        topLayout.Controls.Add(grid, 0, 0);

        // ===== NX =====
        tbNxBaseDir = new TextBox { Dock = DockStyle.Fill };
        btnPickNx = new Button { Text = "Выбрать…", AutoSize = true };
        btnPickNx.Click += (s, e) => PickNxFolder();
        AddRow(grid, "Папка NX (UGII_BASE_DIR):", tbNxBaseDir, btnPickNx);

        // ===== ROOT =====
        tbRoot = new TextBox { Dock = DockStyle.Fill };
        btnPickRoot = new Button { Text = "Выбрать…", AutoSize = true };
        btnPickRoot.Click += (s, e) => PickRootFolder();
        AddRow(grid, "Корневая папка (Root):", tbRoot, btnPickRoot);

        // ===== DB =====
        tbDbPath = new TextBox { Dock = DockStyle.Fill };
        btnPickDb = new Button { Text = "Открыть .db…", AutoSize = true };
        btnPickDb.Click += (s, e) => PickDb();
        AddRow(grid, "База данных (.db):", tbDbPath, btnPickDb);

        btnNewDb = new Button { Text = "Новая база по ROOT", AutoSize = true };
        btnNewDb.Click += (s, e) => CreateDbFromRoot();
        AddRow(grid, "", new Panel { Height = 1 }, btnNewDb);

        // ===== Excel Out =====
        tbExcelOut = new TextBox { Dock = DockStyle.Fill };
        btnPickExcel = new Button { Text = "Выбрать…", AutoSize = true };
        btnPickExcel.Click += (s, e) => PickExcelOut();
        AddRow(grid, "Файл Excel (Out):", tbExcelOut, btnPickExcel);

        // ===== Параметры =====
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

        var lblGroup = new Label { Text = "Группировка:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 10, 0) };
        cmbGroupMode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, Margin = new Padding(0, 3, 0, 3) };
        cmbGroupMode.Items.Add("Первый уровень (FirstLevel)");
        cmbGroupMode.Items.Add("Всё в одном (AllInOne)");
        cmbGroupMode.SelectedIndex = 0;

        var lblMode = new Label { Text = "Режим:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 10, 0) };
        cmbMode = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Dock = DockStyle.Fill, Margin = new Padding(0, 3, 0, 3) };
        cmbMode.Items.Add("Сканировать и выгрузить в Excel (ScanAndExport)");
        cmbMode.Items.Add("Только сканировать (ScanOnly)");
        cmbMode.SelectedIndex = 0;

        paramsGrid.Controls.Add(lblGroup, 0, 0);
        paramsGrid.Controls.Add(cmbGroupMode, 1, 0);
        paramsGrid.Controls.Add(lblMode, 0, 1);
        paramsGrid.Controls.Add(cmbMode, 1, 1);

        AddRow(grid, "Параметры:", paramsGrid, null);

        // ===== Дерево + листы =====
        splitFolders = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical
            // НЕ ставим Panel1MinSize/Panel2MinSize здесь
        };

        tvFolders = new TreeView { Dock = DockStyle.Fill };
        tvFolders.BeforeExpand += (s, e) => ExpandFolderNode(e.Node);
        splitFolders.Panel1.Controls.Add(tvFolders);

        var right = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2
        };
        right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        right.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        right.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        right.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        splitFolders.Panel2.Controls.Add(right);

        lbSheets = new ListBox { Dock = DockStyle.Fill };
        lbSheets.SelectedIndexChanged += (s, e) => OnSheetSelected();
        right.Controls.Add(lbSheets, 0, 0);

        lbSheetFolders = new ListBox { Dock = DockStyle.Fill };
        right.Controls.Add(lbSheetFolders, 1, 0);

        var sheetBtns = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            Padding = new Padding(0, 6, 0, 0)
        };

        btnSheetAdd = new Button { Text = "Создать лист", AutoSize = true };
        btnSheetAdd.Click += (s, e) => AddSheet();

        btnSheetRename = new Button { Text = "Переименовать", AutoSize = true };
        btnSheetRename.Click += (s, e) => RenameSheet();

        btnSheetDelete = new Button { Text = "Удалить", AutoSize = true };
        btnSheetDelete.Click += (s, e) => DeleteSheet();

        btnAddFolderToSheet = new Button { Text = "Добавить папку →", AutoSize = true };
        btnAddFolderToSheet.Click += (s, e) => AddSelectedFolderToSelectedSheet();

        btnRemoveFolderFromSheet = new Button { Text = "Убрать папку", AutoSize = true };
        btnRemoveFolderFromSheet.Click += (s, e) => RemoveSelectedFolderFromSheet();

        sheetBtns.Controls.Add(btnSheetAdd);
        sheetBtns.Controls.Add(btnSheetRename);
        sheetBtns.Controls.Add(btnSheetDelete);
        sheetBtns.Controls.Add(new Label { Text = "   ", AutoSize = true });
        sheetBtns.Controls.Add(btnAddFolderToSheet);
        sheetBtns.Controls.Add(btnRemoveFolderFromSheet);

        right.Controls.Add(sheetBtns, 0, 1);
        right.SetColumnSpan(sheetBtns, 2);

        AddRow(grid, "Листы и папки:", splitFolders, null, makeRowFill: true);

        // ===== Фильтры =====
        AddRow(grid, "Фильтр папок (каждая с новой строки или через ;):", new Panel { Height = 1 }, null);

        tbFilters = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            MinimumSize = new System.Drawing.Size(0, 70)
        };
        AddRow(grid, "", tbFilters, null, makeRowFill: true);

        // ===== Кнопки снизу =====
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

        lblStatus = new Label { Text = "Готово.", AutoSize = true, Anchor = AnchorStyles.Left, Padding = new Padding(8, 6, 0, 0) };

        bottomBar.Controls.Add(btnStart, 0, 0);
        bottomBar.Controls.Add(btnStop, 1, 0);
        bottomBar.Controls.Add(lblStatus, 2, 0);

        topLayout.Controls.Add(bottomBar, 0, 1);

        // ===== Лог =====
        tbLog = new RichTextBox { Dock = DockStyle.Fill, ReadOnly = true };
        split.Panel2.Controls.Add(tbLog);

        FormClosing += (s, e) =>
        {
            try { if (_proc != null && !_proc.HasExited) _proc.Kill(); } catch { }
        };

        // ===== ВСЕ настройки MinSize и SplitterDistance делаем ТОЛЬКО после появления размеров =====
        Load += (s, e) =>
        {
            // 1) главный split (верх/низ)
            split.Panel1MinSize = 260;
            split.Panel2MinSize = 150;
            SetSplitterDistanceSafe(split, Math.Max(split.Panel1MinSize, (int)(split.Height * 0.35)));

            // 2) внутренний split (лево/право)
            splitFolders.Panel1MinSize = 220;
            splitFolders.Panel2MinSize = 260;
            SetSplitterDistanceSafeVertical(splitFolders, splitFolders.Width / 2);
        };

        split.SizeChanged += (s, e) => SetSplitterDistanceSafe(split, split.SplitterDistance);
        splitFolders.SizeChanged += (s, e) => SetSplitterDistanceSafeVertical(splitFolders, splitFolders.SplitterDistance);
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

        string dbPath = (tbDbPath.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(dbPath))
        {
            MessageBox.Show(this, "Сначала выберите или создайте базу данных (.db).", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        dbPath = Path.GetFullPath(dbPath);

        if (!File.Exists(dbPath))
        {
            MessageBox.Show(this, "Файл базы данных не найден:\n" + dbPath, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        // Создаём папки, чтобы не падать на записи ini/out
        try
        {
            Directory.CreateDirectory(Path.Combine(baseDir, "reports"));
            Directory.CreateDirectory(Path.GetDirectoryName(excelOut) ?? baseDir);
            Directory.CreateDirectory(Path.GetDirectoryName(dbPath) ?? baseDir);
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
            File.WriteAllText(iniPath, BuildIni(baseDir, root, excelOut, dbPath), Encoding.UTF8);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "Не удалось записать scan.ini:\n" + ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        AppendLog("NX:   " + nxBase + "\n");
        AppendLog("Root: " + root + "\n");
        AppendLog("DB:   " + dbPath + "\n");
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

    string BuildIni(string baseDir, string root, string excelOut, string dbPath)
    {
        var sb = new StringBuilder();

        string fullRoot = Path.GetFullPath(root);
        string fullExcelOut = Path.GetFullPath(excelOut);
        string fullDbPath = Path.GetFullPath(dbPath);

        sb.AppendLine("[Scan]");
        sb.AppendLine("Root=" + fullRoot);
        sb.AppendLine("Mode=" + (cmbMode.SelectedIndex == 1 ? "ScanOnly" : "ScanAndExport"));
        sb.AppendLine("DbPath=" + fullDbPath);

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
        sb.AppendLine("Out=" + fullExcelOut);
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

        // если кнопки нет — растянем input на 2 колонки (input + пустая)
        if (button == null)
            grid.SetColumnSpan(input, 2);

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

    void PickDb()
    {
        using (var dlg = new OpenFileDialog())
        {
            dlg.Filter = "SQLite DB (*.db)|*.db|All files (*.*)|*.*";
            dlg.Title = "Открыть базу данных";
            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            OpenDb(dlg.FileName);
        }
    }

    void CreateDbFromRoot()
    {
        string root = (tbRoot.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
        {
            MessageBox.Show(this, "Сначала выберите существующую ROOT папку.", "ROOT", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        using (var dlg = new SaveFileDialog())
        {
            dlg.Filter = "SQLite DB (*.db)|*.db";
            dlg.Title = "Создать базу данных";
            dlg.FileName = Path.GetFileName(root.TrimEnd('\\', '/')) + ".db";
            dlg.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;

            if (dlg.ShowDialog(this) != DialogResult.OK) return;

            OpenDb(dlg.FileName);
            tbDbPath.Text = dlg.FileName;

            // сохраняем root в meta
            try { _repo.SetRootPath(root); } catch { }
            _rootFromDb = root;
            LoadRootTree(root);
        }
    }

    void OpenDb(string dbPath)
    {
        tbDbPath.Text = dbPath;

        _repo = new CacheRepository(dbPath);

        // root из базы (meta)
        try { _rootFromDb = _repo.GetRootPath(); } catch { _rootFromDb = null; }

        // если в базе root есть — покажем его в tbRoot (но не перетираем если пользователь уже ввёл?)
        if (!string.IsNullOrWhiteSpace(_rootFromDb))
            tbRoot.Text = _rootFromDb;

        RefreshSheets();

        if (!string.IsNullOrWhiteSpace(tbRoot.Text) && Directory.Exists(tbRoot.Text))
            LoadRootTree(tbRoot.Text);
        else
            tvFolders.Nodes.Clear();
    }

    void LoadRootTree(string root)
    {
        tvFolders.BeginUpdate();
        try
        {
            tvFolders.Nodes.Clear();
            var rootNode = new TreeNode(Path.GetFileName(root.TrimEnd(Path.DirectorySeparatorChar)))
            {
                Tag = root
            };
            rootNode.Nodes.Add(new TreeNode("..."));
            tvFolders.Nodes.Add(rootNode);
            rootNode.Expand();
        }
        finally { tvFolders.EndUpdate(); }
    }

    void ExpandFolderNode(TreeNode node)
    {
        if (node == null) return;
        if (node.Nodes.Count == 1 && node.Nodes[0].Text == "...")
        {
            node.Nodes.Clear();

            string path = node.Tag as string;
            if (string.IsNullOrWhiteSpace(path) || !Directory.Exists(path)) return;

            try
            {
                foreach (var dir in Directory.GetDirectories(path))
                {
                    var child = new TreeNode(Path.GetFileName(dir)) { Tag = dir };
                    try
                    {
                        if (Directory.GetDirectories(dir).Length > 0)
                            child.Nodes.Add(new TreeNode("..."));
                    }
                    catch { }
                    node.Nodes.Add(child);
                }
            }
            catch { /* на сетевых дисках/правах может падать — игнор */ }
        }
    }

    // ===== Sheets UI =====

    sealed class SheetListItem
    {
        public int Id;
        public string Name;
        public SheetListItem(int id, string name) { Id = id; Name = name; }
        public override string ToString() => Name;
    }

    void RefreshSheets()
    {
        lbSheets.Items.Clear();
        lbSheetFolders.Items.Clear();
        _selectedSheetId = -1;

        if (_repo == null) return;

        var sheets = _repo.ListSheets(); // ты реализуешь в CacheRepository
        foreach (var sh in sheets)
            lbSheets.Items.Add(new SheetListItem(sh.Id, sh.Name));
    }

    void OnSheetSelected()
    {
        lbSheetFolders.Items.Clear();
        var item = lbSheets.SelectedItem as SheetListItem;
        if (item == null || _repo == null) { _selectedSheetId = -1; return; }

        _selectedSheetId = item.Id;

        var sheets = _repo.GetSheetsWithFolders(); // ты реализуешь в CacheRepository
        var sInfo = sheets.Find(x => x.Id == _selectedSheetId);
        if (sInfo == null) return;

        foreach (var f in sInfo.Folders)
            lbSheetFolders.Items.Add(f.RelPath + (f.IncludeSubfolders ? "  (включая подпапки)" : ""));
    }

    void AddSheet()
    {
        if (_repo == null) { MessageBox.Show(this, "Откройте базу данных."); return; }
        string name = Prompt("Название листа:", "Новый лист");
        if (string.IsNullOrWhiteSpace(name)) return;

        _repo.CreateSheet(name.Trim());
        RefreshSheets();
    }

    void RenameSheet()
    {
        if (_repo == null) return;
        var item = lbSheets.SelectedItem as SheetListItem;
        if (item == null) return;

        string name = Prompt("Новое название листа:", item.Name);
        if (string.IsNullOrWhiteSpace(name)) return;

        _repo.RenameSheet(item.Id, name.Trim());
        RefreshSheets();
    }

    void DeleteSheet()
    {
        if (_repo == null) return;
        var item = lbSheets.SelectedItem as SheetListItem;
        if (item == null) return;

        var ok = MessageBox.Show(this, $"Удалить лист \"{item.Name}\"?", "Подтверждение",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

        if (ok != DialogResult.Yes) return;

        _repo.DeleteSheet(item.Id);
        RefreshSheets();
    }

    void AddSelectedFolderToSelectedSheet()
    {
        if (_repo == null) { MessageBox.Show(this, "Откройте базу данных."); return; }
        if (_selectedSheetId <= 0) { MessageBox.Show(this, "Выберите лист справа."); return; }
        if (tvFolders.SelectedNode == null) { MessageBox.Show(this, "Выберите папку слева."); return; }

        string root = (tbRoot.Text ?? "").Trim();
        if (string.IsNullOrWhiteSpace(root) || !Directory.Exists(root))
        {
            MessageBox.Show(this, "ROOT не задан или не существует."); return;
        }

        string abs = tvFolders.SelectedNode.Tag as string;
        if (string.IsNullOrWhiteSpace(abs)) return;

        string rel = MakeRel(root, abs);
        if (rel == null)
        {
            MessageBox.Show(this, "Папка не принадлежит ROOT."); return;
        }

        _repo.AddFolderToSheet(_selectedSheetId, rel, includeSubfolders: true);
        OnSheetSelected();
    }

    void RemoveSelectedFolderFromSheet()
    {
        if (_repo == null) return;
        if (_selectedSheetId <= 0) return;
        if (lbSheetFolders.SelectedItem == null) return;

        string s = lbSheetFolders.SelectedItem.ToString();
        string rel = s;
        int idx = s.IndexOf("  (", StringComparison.Ordinal);
        if (idx > 0) rel = s.Substring(0, idx);

        _repo.RemoveFolderFromSheet(_selectedSheetId, rel);
        OnSheetSelected();
    }

    static string MakeRel(string root, string abs)
    {
        try
        {
            if (!root.EndsWith(Path.DirectorySeparatorChar.ToString()))
                root += Path.DirectorySeparatorChar;

            var rootUri = new Uri(root);
            var absUri = new Uri(abs);

            if (rootUri.Scheme != absUri.Scheme) return null;

            var relUri = rootUri.MakeRelativeUri(absUri);
            return Uri.UnescapeDataString(relUri.ToString()).Replace('/', Path.DirectorySeparatorChar);
        }
        catch { return null; }
    }

    static string Prompt(string title, string defaultValue)
    {
        using (var f = new Form())
        {
            f.Text = title;
            f.Width = 420;
            f.Height = 140;
            f.StartPosition = FormStartPosition.CenterParent;
            f.FormBorderStyle = FormBorderStyle.FixedDialog;
            f.MinimizeBox = false;
            f.MaximizeBox = false;

            var tb = new TextBox { Left = 10, Top = 10, Width = 380, Text = defaultValue ?? "" };
            var ok = new Button { Text = "OK", Left = 230, Width = 75, Top = 45, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Отмена", Left = 315, Width = 75, Top = 45, DialogResult = DialogResult.Cancel };

            f.Controls.Add(tb);
            f.Controls.Add(ok);
            f.Controls.Add(cancel);
            f.AcceptButton = ok;
            f.CancelButton = cancel;

            return f.ShowDialog() == DialogResult.OK ? tb.Text : null;
        }
    }

    static void SetSplitterDistanceSafe(SplitContainer split, int desired)
    {
        if (split == null) return;

        // Важно: при Horizontal считаем по Height
        int min = split.Panel1MinSize;
        int max = split.Height - split.Panel2MinSize - split.SplitterWidth;

        if (max < min) max = min; // на случай совсем маленького окна

        int v = desired;
        if (v < min) v = min;
        if (v > max) v = max;

        // чтобы не ловить исключение в момент, когда Handle ещё не создан
        try { split.SplitterDistance = v; } catch { }
    }

    static void SetSplitterDistanceSafeVertical(SplitContainer split, int desired)
    {
        if (split == null) return;

        int min = split.Panel1MinSize;
        int max = split.Width - split.Panel2MinSize - split.SplitterWidth;
        if (max < min) max = min;

        int v = desired;
        if (v < min) v = min;
        if (v > max) v = max;

        try { split.SplitterDistance = v; } catch { }
    }
}