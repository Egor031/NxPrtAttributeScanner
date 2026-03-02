# NxPrtAttributeScanner

Инструмент для пакетного сканирования атрибутов NX (*.prt) файлов  
с сохранением результатов в Excel.

Поддерживает:

- 🔍 Рекурсивный поиск *.prt
- 📦 Кэширование через SQLite
- 📊 Экспорт в Excel (EPPlus)
- 🗂 Группировку по папкам
- 🖥 GUI + headless режим (через run_journal.exe)

---

## Требования

- Siemens NX (поддерживался NX 1899)
- .NET Framework 4.7.2
- Доступ к `run_journal.exe`

---

# 🖥 GUI версия

1. Соберите проект `NxPrtAttributeScannerGUI`
2. Запустите `NxPrtAttributeScannerGUI.exe`
3. Укажите:
   - Папку NX
   - Корневую папку с *.prt
   - Путь сохранения Excel
4. Нажмите **Запустить**

---

# ⚙ Headless (без GUI)

Можно запускать напрямую через NX journal.

### 1️⃣ Создайте файл `scan.ini`

Пример:

```ini
[Scan]
Root=D:\Data\PRT
Mode=ScanAndExport
DbPath=.\cache\parts.db

IncludeFolderName=Кронштейны
IncludeFolderName=Корпуса

[Excel]
Out=.\reports\Parts.xlsx
GroupMode=FirstLevel
````

### 2️⃣ Запуск

```bash
NXBIN\run_journal.exe "RunScanner.cs" -args config="scan.ini" scannerDll="NxPrtAttributeScanner.dll"
```

---

# Структура проекта

```
NxPrtAttributeScanner/
│
├── NxPrtAttributeScanner/        # DLL проект (ядро)
├── NxPrtAttributeScannerGUI/     # GUI
├── RunScanner.cs                 # NX journal loader
├── sample.scan.ini               # пример конфигурации
```

---

# Кэширование

Программа использует SQLite базу для:

* предотвращения повторного чтения неизменённых файлов
* хранения атрибутов
* отслеживания ошибок

Удалённые файлы автоматически удаляются из базы при повторном сканировании.

---

# Режимы

### ScanAndExport

Сканирует и создаёт Excel

### ScanOnly

Только обновляет кэш

---

# Группировка

* `FirstLevel` — группировка по первой папке относительно Root
* `AllInOne` — один лист Excel
