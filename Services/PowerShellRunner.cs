using System;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PCDiagnosticPro.Services
{
    public class PowerShellRunner
    {
        public event Action<string>? OutputReceived;
        public event Action<string>? ErrorReceived;

        private Process? _currentProcess;

        public async Task<PowerShellRunResult> RunAsync(
            string scriptPath,
            string arguments,
            TimeSpan timeout,
            CancellationToken cancellationToken)
        {
            var outputBuilder = new StringBuilder();
            var errorBuilder = new StringBuilder();

            using var timeoutCts = new CancellationTokenSource(timeout);
            using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(timeoutCts.Token, cancellationToken);

            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\" {arguments}",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = new UTF8Encoding(false),
                    StandardErrorEncoding = new UTF8Encoding(false)
                };

                _currentProcess = new Process { StartInfo = startInfo, EnableRaisingEvents = true };

                _currentProcess.OutputDataReceived += (sender, e) =>
                {
                    if (string.IsNullOrEmpty(e.Data)) return;
                    outputBuilder.AppendLine(e.Data);
                    OutputReceived?.Invoke(e.Data);
                };

                _currentProcess.ErrorDataReceived += (sender, e) =>
                {
                    if (string.IsNullOrEmpty(e.Data)) return;
                    errorBuilder.AppendLine(e.Data);
                    ErrorReceived?.Invoke(e.Data);
                };

                _currentProcess.Start();
                _currentProcess.BeginOutputReadLine();
                _currentProcess.BeginErrorReadLine();

                await _currentProcess.WaitForExitAsync(linkedCts.Token);

                return new PowerShellRunResult(
                    exitCode: _currentProcess.ExitCode,
                    output: outputBuilder.ToString(),
                    error: errorBuilder.ToString(),
                    timedOut: false);
            }
            catch (OperationCanceledException)
            {
                var timedOut = timeoutCts.IsCancellationRequested && !cancellationToken.IsCancellationRequested;
                TryKillProcess();
                return new PowerShellRunResult(
                    exitCode: -1,
                    output: outputBuilder.ToString(),
                    error: errorBuilder.ToString(),
                    timedOut: timedOut);
            }
            finally
            {
                _currentProcess?.Dispose();
                _currentProcess = null;
            }
        }

        public void Cancel()
        {
            TryKillProcess();
        }

        private void TryKillProcess()
        {
            try
            {
                if (_currentProcess != null && !_currentProcess.HasExited)
                {
                    _currentProcess.Kill(true);
                }
            }
            catch
            {
                // Ignore kill failures
            }
        }
    }

    public record PowerShellRunResult(int ExitCode, string Output, string Error, bool TimedOut);
}
