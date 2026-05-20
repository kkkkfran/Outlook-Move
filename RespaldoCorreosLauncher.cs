using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

internal static class RespaldoCorreosLauncher
{
    [STAThread]
    private static void Main()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string scriptPath = Path.Combine(baseDir, "RespaldoCorreosPorAno.ps1");

        if (!File.Exists(scriptPath))
        {
            MessageBox.Show(
                "No se encontro RespaldoCorreosPorAno.ps1 en la misma carpeta del ejecutable.",
                "Respaldo de correos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        string powerShellPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows),
            @"System32\WindowsPowerShell\v1.0\powershell.exe");

        if (!File.Exists(powerShellPath))
        {
            MessageBox.Show(
                "No se encontro Windows PowerShell.",
                "Respaldo de correos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        string arguments =
            "-NoProfile -ExecutionPolicy Bypass -STA -File " +
            Quote(scriptPath);

        try
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = powerShellPath;
            startInfo.Arguments = arguments;
            startInfo.WorkingDirectory = baseDir;
            startInfo.UseShellExecute = false;
            startInfo.CreateNoWindow = true;
            Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "No se pudo abrir la herramienta:\r\n\r\n" + ex.Message,
                "Respaldo de correos",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
    }

    private static string Quote(string value)
    {
        return "\"" + value.Replace("\"", "\\\"") + "\"";
    }
}
