// CacheRepository.cs (updated)
// Зависимости: System.Data.SQLite, System.IO, System.Globalization, System.Collections.Generic
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Globalization;

public sealed class PartRow
{
    public string FullPath;
    public string PartNoFile;
    public string Designation;
    public string Match;
    public string FolderSheetKey; // для группировки по листам
    public DateTime LastWriteUtc;
    public DateTime ExtractedUtc;
    public Dictionary<string, string> Attrs;
}

public sealed class SheetInfo
{
    public int Id;
    public string Name;
    public int SortOrder;
    public List<SheetFolder> Folders = new List<SheetFolder>();
}

public sealed class SheetFolder
{
    public string RelPath; // относительный путь относительно root
    public bool IncludeSubfolders;
}

public sealed class FilterItem
{
    public int Id;
    public string Type; // "include" или "exclude"
    public string Pattern;
    public bool Enabled;
}

public sealed class CacheRepository
{
    private readonly string _dbPath;

    public CacheRepository(string dbPath)
    {
        _dbPath = dbPath;
        Directory.CreateDirectory(Path.GetDirectoryName(dbPath) ?? ".");
        Init();
    }

    private void Init()
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();

            using (var cmd = con.CreateCommand())
            {
                // существующие таблицы + новые для листов/фильтров/meta
                cmd.CommandText = @"
PRAGMA journal_mode = WAL;

CREATE TABLE IF NOT EXISTS meta (
  key TEXT PRIMARY KEY,
  value TEXT
);

CREATE TABLE IF NOT EXISTS files (
  full_path TEXT PRIMARY KEY,
  file_size INTEGER NOT NULL,
  last_write_time_utc TEXT NOT NULL,
  part_no_file TEXT NOT NULL,
  designation_attr TEXT,
  match TEXT,
  extracted_at_utc TEXT NOT NULL,
  status TEXT NOT NULL,
  error_message TEXT
);
CREATE INDEX IF NOT EXISTS idx_files_partno ON files(part_no_file);

CREATE TABLE IF NOT EXISTS attributes (
  full_path TEXT NOT NULL,
  attr_name TEXT NOT NULL,
  attr_value TEXT,
  PRIMARY KEY(full_path, attr_name)
);
CREATE INDEX IF NOT EXISTS idx_attr_name ON attributes(attr_name);

-- Листы (sheets) и связанные с ними папки (relative paths)
CREATE TABLE IF NOT EXISTS sheets (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  sort_order INTEGER NOT NULL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS sheet_folders (
  sheet_id INTEGER NOT NULL,
  rel_path TEXT NOT NULL,
  include_subfolders INTEGER NOT NULL DEFAULT 1,
  PRIMARY KEY(sheet_id, rel_path)
);

CREATE TABLE IF NOT EXISTS filters (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  type TEXT NOT NULL, -- 'include' or 'exclude'
  pattern TEXT NOT NULL,
  enabled INTEGER NOT NULL DEFAULT 1
);
";
                cmd.ExecuteNonQuery();
            }

            // ensure meta.schema_version exists
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "INSERT OR IGNORE INTO meta(key,value) VALUES('schema_version','1');";
                cmd.ExecuteNonQuery();
            }
        }
    }

    // -------------------------
    // meta helpers
    // -------------------------
    public string GetMeta(string key, string defaultValue = null)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT value FROM meta WHERE key = $k;";
                cmd.Parameters.AddWithValue("$k", key);
                var v = cmd.ExecuteScalar();
                return v == null ? defaultValue : (string)v;
            }
        }
    }

    public void SetMeta(string key, string value)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "INSERT INTO meta(key,value) VALUES($k,$v) ON CONFLICT(key) DO UPDATE SET value=excluded.value;";
                cmd.Parameters.AddWithValue("$k", key);
                cmd.Parameters.AddWithValue("$v", value ?? "");
                cmd.ExecuteNonQuery();
            }
        }
    }

    // Root path stored in meta as 'root_path'
    public string GetRootPath()
    {
        return GetMeta("root_path", "");
    }

    public void SetRootPath(string root)
    {
        SetMeta("root_path", root ?? "");
    }

    // -------------------------
    // sheets / sheet_folders
    // -------------------------
    public List<SheetInfo> ListSheets()
    {
        var res = new List<SheetInfo>();
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT id, name, sort_order FROM sheets ORDER BY sort_order, id;";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        res.Add(new SheetInfo
                        {
                            Id = r.GetInt32(0),
                            Name = r.IsDBNull(1) ? "" : r.GetString(1),
                            SortOrder = r.IsDBNull(2) ? 0 : r.GetInt32(2)
                        });
                    }
                }
            }

            // load folders for each sheet
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT sheet_id, rel_path, include_subfolders
FROM sheet_folders
ORDER BY sheet_id, rel_path;";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        int sid = r.GetInt32(0);
                        string rel = r.IsDBNull(1) ? "" : r.GetString(1);
                        bool inc = r.IsDBNull(2) ? true : r.GetInt32(2) != 0;
                        var s = res.Find(x => x.Id == sid);
                        if (s != null)
                            s.Folders.Add(new SheetFolder { RelPath = rel, IncludeSubfolders = inc });
                    }
                }
            }
        }
        return res;
    }

    public int CreateSheet(string name)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var tx = con.BeginTransaction())
            using (var cmd = con.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "INSERT INTO sheets(name, sort_order) VALUES($n, (SELECT IFNULL(MAX(sort_order),0)+1 FROM sheets));";
                cmd.Parameters.AddWithValue("$n", name ?? "");
                cmd.ExecuteNonQuery();

                cmd.CommandText = "SELECT last_insert_rowid();";
                long id = (long)cmd.ExecuteScalar();
                tx.Commit();
                return (int)id;
            }
        }
    }

    public void RenameSheet(int id, string newName)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "UPDATE sheets SET name=$n WHERE id=$id;";
                cmd.Parameters.AddWithValue("$n", newName ?? "");
                cmd.Parameters.AddWithValue("$id", id);
                cmd.ExecuteNonQuery();
            }
        }
    }

    public void DeleteSheet(int id)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var tx = con.BeginTransaction())
            {
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "DELETE FROM sheet_folders WHERE sheet_id = $id;";
                    cmd.Parameters.AddWithValue("$id", id);
                    cmd.ExecuteNonQuery();
                }
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = "DELETE FROM sheets WHERE id = $id;";
                    cmd.Parameters.AddWithValue("$id", id);
                    cmd.ExecuteNonQuery();
                }
                tx.Commit();
            }
        }
    }

    public void AddFolderToSheet(int sheetId, string relPath, bool includeSubfolders)
    {
        if (relPath == null) relPath = "";
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
INSERT OR REPLACE INTO sheet_folders(sheet_id, rel_path, include_subfolders)
VALUES($sid, $rel, $inc);";
                cmd.Parameters.AddWithValue("$sid", sheetId);
                cmd.Parameters.AddWithValue("$rel", relPath);
                cmd.Parameters.AddWithValue("$inc", includeSubfolders ? 1 : 0);
                cmd.ExecuteNonQuery();
            }
        }
    }

    public void RemoveFolderFromSheet(int sheetId, string relPath)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM sheet_folders WHERE sheet_id = $sid AND rel_path = $rel;";
                cmd.Parameters.AddWithValue("$sid", sheetId);
                cmd.Parameters.AddWithValue("$rel", relPath ?? "");
                cmd.ExecuteNonQuery();
            }
        }
    }

    // Возвращает все листы и папки (удобно для GUI)
    public List<SheetInfo> GetSheetsWithFolders()
    {
        return ListSheets();
    }

    // -------------------------
    // filters (simple)
    // -------------------------
    public List<FilterItem> ListFilters()
    {
        var res = new List<FilterItem>();
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "SELECT id, type, pattern, enabled FROM filters ORDER BY id;";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        res.Add(new FilterItem
                        {
                            Id = r.GetInt32(0),
                            Type = r.IsDBNull(1) ? "" : r.GetString(1),
                            Pattern = r.IsDBNull(2) ? "" : r.GetString(2),
                            Enabled = r.IsDBNull(3) ? true : r.GetInt32(3) != 0
                        });
                    }
                }
            }
        }
        return res;
    }

    public int AddFilter(string type, string pattern, bool enabled = true)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "INSERT INTO filters(type, pattern, enabled) VALUES($t,$p,$e);";
                cmd.Parameters.AddWithValue("$t", type ?? "include");
                cmd.Parameters.AddWithValue("$p", pattern ?? "");
                cmd.Parameters.AddWithValue("$e", enabled ? 1 : 0);
                cmd.ExecuteNonQuery();
                cmd.CommandText = "SELECT last_insert_rowid();";
                long id = (long)cmd.ExecuteScalar();
                return (int)id;
            }
        }
    }

    public void RemoveFilter(int id)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = "DELETE FROM filters WHERE id = $id;";
                cmd.Parameters.AddWithValue("$id", id);
                cmd.ExecuteNonQuery();
            }
        }
    }

    // -------------------------
    // прежние методы (без изменений по логике, минимально подправлены)
    // -------------------------
    public bool NeedsExtraction(string fullPath, long fileSize, DateTime lastWriteUtc)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();

            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT file_size, last_write_time_utc
FROM files
WHERE full_path = $p;";
                cmd.Parameters.AddWithValue("$p", fullPath);

                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) return true;

                    long oldSize = r.GetInt64(0);
                    string oldTime = r.GetString(1);

                    string newTime = lastWriteUtc.ToString("O");

                    return !(oldSize == fileSize && string.Equals(oldTime, newTime, StringComparison.Ordinal));
                }
            }
        }
    }

    public void UpsertOk(string fullPath, long fileSize, DateTime lastWriteUtc,
    string partNoFile, string designation, string match, Dictionary<string, string> attrs)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var tx = con.BeginTransaction())
            {
                // 1) upsert files
                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"
INSERT INTO files(full_path, file_size, last_write_time_utc, part_no_file, designation_attr, match, extracted_at_utc, status, error_message)
VALUES($path,$size,$lw,$pn,$des,$match,$ex,'OK',NULL)
ON CONFLICT(full_path) DO UPDATE SET
 file_size=excluded.file_size,
 last_write_time_utc=excluded.last_write_time_utc,
 part_no_file=excluded.part_no_file,
 designation_attr=excluded.designation_attr,
 match=excluded.match,
 extracted_at_utc=excluded.extracted_at_utc,
 status='OK',
 error_message=NULL;";

                    cmd.Parameters.AddWithValue("$path", fullPath);
                    cmd.Parameters.AddWithValue("$size", fileSize);
                    cmd.Parameters.AddWithValue("$lw", lastWriteUtc.ToString("O"));
                    cmd.Parameters.AddWithValue("$pn", partNoFile);
                    cmd.Parameters.AddWithValue("$des", designation ?? "");
                    cmd.Parameters.AddWithValue("$match", match ?? "");
                    cmd.Parameters.AddWithValue("$ex", DateTime.UtcNow.ToString("O"));
                    cmd.ExecuteNonQuery();
                }

                // 2) удалить старые атрибуты (на случай, если атрибут убрали/переименовали)
                using (var del = con.CreateCommand())
                {
                    del.Transaction = tx;
                    del.CommandText = "DELETE FROM attributes WHERE full_path = $path;";
                    del.Parameters.AddWithValue("$path", fullPath);
                    del.ExecuteNonQuery();
                }

                // 3) вставить новые атрибуты
                using (var ins = con.CreateCommand())
                {
                    ins.Transaction = tx;
                    ins.CommandText = @"
INSERT INTO attributes(full_path, attr_name, attr_value)
VALUES($path,$name,$value);";

                    var pPath = ins.Parameters.Add("$path", System.Data.DbType.String);
                    var pName = ins.Parameters.Add("$name", System.Data.DbType.String);
                    var pVal = ins.Parameters.Add("$value", System.Data.DbType.String);

                    foreach (var kv in attrs)
                    {
                        pPath.Value = fullPath;
                        pName.Value = kv.Key;
                        pVal.Value = kv.Value ?? "";
                        ins.ExecuteNonQuery();
                    }
                }

                tx.Commit();
            }
        }
    }

    public void UpsertError(string fullPath, long fileSize, DateTime lastWriteUtc,
        string partNoFile, string error)
    {
        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();

            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
INSERT INTO files(full_path, file_size, last_write_time_utc, part_no_file, designation_attr, match, extracted_at_utc, status, error_message)
VALUES($path,$size,$lw,$pn,'','',$ex,'ERROR',$err)
ON CONFLICT(full_path) DO UPDATE SET
 file_size=excluded.file_size,
 last_write_time_utc=excluded.last_write_time_utc,
 part_no_file=excluded.part_no_file,
 extracted_at_utc=excluded.extracted_at_utc,
 status='ERROR',
 error_message=excluded.error_message;";
                cmd.Parameters.AddWithValue("$path", fullPath);
                cmd.Parameters.AddWithValue("$size", fileSize);
                cmd.Parameters.AddWithValue("$lw", lastWriteUtc.ToString("O"));
                cmd.Parameters.AddWithValue("$pn", partNoFile);
                cmd.Parameters.AddWithValue("$ex", DateTime.UtcNow.ToString("O"));
                cmd.Parameters.AddWithValue("$err", error ?? "");

                cmd.ExecuteNonQuery();
            }
        }
    }

    public List<string> GetAllAttributeNames()
    {
        var names = new List<string>();

        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT DISTINCT attr_name
FROM attributes
ORDER BY attr_name;";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                        names.Add(r.GetString(0));
                }
            }
        }

        return names;
    }

    // Возвращает все части под указанным root (только записи статус='OK' и подпадающие под root)
    public List<PartRow> GetAllParts(string rootFolder, bool groupByFolderSheets)
    {
        var parts = new List<PartRow>();
        var byPath = new Dictionary<string, PartRow>(StringComparer.OrdinalIgnoreCase);

        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();

            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT full_path, part_no_file, designation_attr, match, last_write_time_utc, extracted_at_utc
FROM files
WHERE status='OK'
  AND full_path LIKE $root
ORDER BY full_path;";
                cmd.Parameters.AddWithValue("$root", NormalizeRootForLike(rootFolder) + "%");

                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        var fullPath = r.GetString(0);

                        var row = new PartRow
                        {
                            FullPath = fullPath,
                            PartNoFile = r.IsDBNull(1) ? "" : r.GetString(1),
                            Designation = r.IsDBNull(2) ? "" : r.GetString(2),
                            Match = r.IsDBNull(3) ? "" : r.GetString(3),
                            LastWriteUtc = ParseIsoUtc(r.IsDBNull(4) ? "" : r.GetString(4)),
                            ExtractedUtc = ParseIsoUtc(r.IsDBNull(5) ? "" : r.GetString(5)),
                            FolderSheetKey = groupByFolderSheets ? GetSheetKey(rootFolder, fullPath) : "ALL",
                            Attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                        };

                        parts.Add(row);
                        byPath[fullPath] = row;
                    }
                }
            }

            // 2) читаем attributes одним проходом
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT full_path, attr_name, attr_value
FROM attributes
ORDER BY full_path, attr_name;";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        string p = r.GetString(0);
                        PartRow row;
                        if (!byPath.TryGetValue(p, out row)) continue;

                        string name = r.GetString(1);
                        string value = r.IsDBNull(2) ? "" : r.GetString(2);
                        row.Attrs[name] = value;
                    }
                }
            }
        }

        return parts;
    }

    private static DateTime ParseIsoUtc(string iso)
    {
        if (string.IsNullOrWhiteSpace(iso))
            return DateTime.SpecifyKind(DateTime.MinValue, DateTimeKind.Utc);

        return DateTime.Parse(iso, null, DateTimeStyles.RoundtripKind).ToUniversalTime();
    }

    private static string NormalizeRootForLike(string rootFolder)
    {
        if (string.IsNullOrEmpty(rootFolder)) return "";
        string root = rootFolder;

        // чтобы D:\Root2 не матчился под D:\Root
        if (!root.EndsWith(Path.DirectorySeparatorChar.ToString()))
            root += Path.DirectorySeparatorChar;

        return root.Replace("'", "''"); // minimal escaping for LIKE pattern
    }

    // Для .NET Framework 4.7.2 нет Path.GetRelativePath → используем Uri
    private static string GetSheetKey(string rootFolder, string fullPath)
    {
        try
        {
            string rel = GetRelativePathSafe(rootFolder, fullPath); // A\B\C\file.prt
            if (string.IsNullOrEmpty(rel)) return "ROOT";

            // берём первую папку относительно root
            int idx = rel.IndexOf(Path.DirectorySeparatorChar);
            if (idx < 0) return "ROOT";
            string first = rel.Substring(0, idx);
            return string.IsNullOrWhiteSpace(first) ? "ROOT" : first;
        }
        catch
        {
            return "ROOT";
        }
    }

    private static string GetRelativePathSafe(string rootFolder, string fullPath)
    {
        if (string.IsNullOrEmpty(rootFolder) || string.IsNullOrEmpty(fullPath))
            return "";

        // гарантируем слэш в конце root
        if (!rootFolder.EndsWith(Path.DirectorySeparatorChar.ToString()))
            rootFolder += Path.DirectorySeparatorChar;

        var rootUri = new Uri(rootFolder);
        var fileUri = new Uri(fullPath);

        if (rootUri.Scheme != fileUri.Scheme) return "";

        var relUri = rootUri.MakeRelativeUri(fileUri);
        string rel = Uri.UnescapeDataString(relUri.ToString());

        // Uri даёт '/', приводим к '\'
        rel = rel.Replace('/', Path.DirectorySeparatorChar);
        return rel;
    }

    // При необходимости получить абсолютный путь по относительному, зная root
    public static string GetAbsolutePathFromRel(string rootFolder, string relPath)
    {
        if (string.IsNullOrEmpty(relPath)) return rootFolder;
        if (!rootFolder.EndsWith(Path.DirectorySeparatorChar.ToString()))
            rootFolder += Path.DirectorySeparatorChar;
        return Path.Combine(rootFolder, relPath);
    }

    public Dictionary<string, string> GetAttributesByPath(string fullPath)
    {
        var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        using (var con = new System.Data.SQLite.SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var cmd = con.CreateCommand())
            {
                cmd.CommandText = @"
SELECT attr_name, attr_value
FROM attributes
WHERE full_path = $p
ORDER BY attr_name;";
                cmd.Parameters.AddWithValue("$p", fullPath);

                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        var name = r.GetString(0);
                        var val = r.IsDBNull(1) ? "" : r.GetString(1);
                        dict[name] = val;
                    }
                }
            }
        }

        return dict;
    }

    // Удаляем записи, которые не были увидены в последнем проходе, НО ТОЛЬКО ПОД УКАЗАННЫМ ROOT
    public void RemoveNotSeenUnderRoot(string rootFolder, HashSet<string> seen)
    {
        string rootLike = NormalizeRootForLike(rootFolder) + "%";

        using (var con = new SQLiteConnection($"Data Source={_dbPath};Version=3;"))
        {
            con.Open();
            using (var tx = con.BeginTransaction())
            {
                var toDelete = new List<string>();

                using (var cmd = con.CreateCommand())
                {
                    cmd.Transaction = tx;
                    cmd.CommandText = @"
SELECT full_path
FROM files
WHERE full_path LIKE $root;";
                    cmd.Parameters.AddWithValue("$root", rootLike);

                    using (var r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            string p = r.GetString(0);
                            if (!seen.Contains(p))
                                toDelete.Add(p);
                        }
                    }
                }

                foreach (var p in toDelete)
                {
                    using (var del1 = con.CreateCommand())
                    {
                        del1.Transaction = tx;
                        del1.CommandText = "DELETE FROM attributes WHERE full_path=$p;";
                        del1.Parameters.AddWithValue("$p", p);
                        del1.ExecuteNonQuery();
                    }

                    using (var del2 = con.CreateCommand())
                    {
                        del2.Transaction = tx;
                        del2.CommandText = "DELETE FROM files WHERE full_path=$p;";
                        del2.Parameters.AddWithValue("$p", p);
                        del2.ExecuteNonQuery();
                    }
                }

                tx.Commit();
            }
        }
    }
}