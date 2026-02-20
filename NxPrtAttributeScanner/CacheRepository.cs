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
                cmd.CommandText = @"
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
";
                cmd.ExecuteNonQuery();
            }
        }
    }

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
INSERT INTO files(full_path, file_size, last_write_time_utc, part_no_file, designation_attr, match, extracted_at_utc, attrs_json, status, error_message)
VALUES($path,$size,$lw,$pn,'','',$ex,'','ERROR',$err)
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

    public List<PartRow> GetAllParts(string rootFolder, bool groupByFolderSheets)
    {
        // 1) читаем files
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
ORDER BY full_path;";
                using (var r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        var fullPath = r.GetString(0);

                        var row = new PartRow
                        {
                            FullPath = fullPath,
                            PartNoFile = r.GetString(1),
                            Designation = r.IsDBNull(2) ? "" : r.GetString(2),
                            Match = r.IsDBNull(3) ? "" : r.GetString(3),
                            LastWriteUtc = ParseIsoUtc(r.GetString(4)),
                            ExtractedUtc = ParseIsoUtc(r.GetString(5)),
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
        // Ты писал "O" при сохранении — это ISO 8601 round-trip
        // Важно: считаем как UTC
        return DateTime.Parse(iso, null, DateTimeStyles.RoundtripKind).ToUniversalTime();
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
}

