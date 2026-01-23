using System;
using System.Diagnostics;
using System.Security.Principal;

namespace PCDiagnosticPro.Services
{
    /// <summary>
    /// Service pour gérer les privilèges administrateur
    /// </summary>
    public static class AdminService
    {
        /// <summary>
        /// Vérifie si l'application s'exécute en mode administrateur
        /// </summary>
        public static bool IsRunningAsAdmin()
        {
            try
            {
                using var identity = WindowsIdentity.GetCurrent();
                var principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Redémarre l'application en mode administrateur
        /// </summary>
        public static void RestartAsAdmin()
        {
            try
            {
                var exePath = Environment.ProcessPath;
                if (string.IsNullOrEmpty(exePath)) return;

                var startInfo = new ProcessStartInfo
                {
                    FileName = exePath,
                    UseShellExecute = true,
                    Verb = "runas"
                };

                Process.Start(startInfo);
                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                App.LogMessage($"Erreur lors du redémarrage en admin: {ex.Message}");
                throw;
            }
        }
    }
}
