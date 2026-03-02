using System;
using System.IO;
using System.Windows.Forms;

static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // Глобальный перехват, чтобы не "молча закрывалось"
        Application.ThreadException += (s, e) => ShowFatal(e.Exception);
        AppDomain.CurrentDomain.UnhandledException += (s, e) =>
            ShowFatal(e.ExceptionObject as Exception ?? new Exception("UnhandledException"));

        try
        {
            Application.Run(new MainForm());
        }
        catch (Exception ex)
        {
            ShowFatal(ex);
        }
    }

    static void ShowFatal(Exception ex)
    {
        try
        {
            File.WriteAllText(
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "last_gui_error.txt"),
                ex.ToString()
            );
        }
        catch { }

        MessageBox.Show(
            ex.ToString(),
            "FATAL ERROR (GUI)",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error
        );
    }
}