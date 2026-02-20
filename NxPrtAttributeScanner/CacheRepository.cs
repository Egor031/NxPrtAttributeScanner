using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;

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
}