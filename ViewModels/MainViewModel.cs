using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using PCDiagnosticPro.Models;
using PCDiagnosticPro.Services;

namespace PCDiagnosticPro.ViewModels
{
    /// <summary>
    /// ViewModel principal de l'application
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        #region Fields

        private readonly PowerShellService _powerShellService;
        private readonly ReportParserService _reportParserService;
        private readonly DispatcherTimer _liveFeedTimer;
        private readonly Stopwatch _scanStopwatch;

        // Process management pour Cancel
        private Process? _scanProcess;
        private CancellationTokenSource? _scanCts;
        private readonly object _scanLock = new object();

        // Chemins relatifs
        private readonly string _baseDir = AppContext.BaseDirectory;
        private string _scriptPath = string.Empty;
        private string _reportsDir = string.Empty;
        private string _resultJsonPath = string.Empty;
        private string _configPath = string.Empty;

        // Settings loading flag
        private bool _isLoadingSettings = false;

        // Progress tracking
        private int _totalSteps = 27;

        #endregion

        #region Properties

        // Navigation
        private string _currentView = "Home";
        public string CurrentView
        {
            get => _currentView;
            set
            {
                if (SetProperty(ref _currentView, value))
                {
                    OnPropertyChanged(nameof(IsScannerView));
                    OnPropertyChanged(nameof(IsResultsView));
                    OnPropertyChanged(nameof(IsSettingsView));
                    OnPropertyChanged(nameof(IsHealthcheckView));
                    OnPropertyChanged(nameof(IsChatView));
                }
            }
        }

        public bool IsScannerView => CurrentView == "Home";
        public bool IsResultsView => CurrentView == "Results";
        public bool IsSettingsView => CurrentView == "Settings";
        public bool IsHealthcheckView => CurrentView == "Healthcheck";
        public bool IsChatView => CurrentView == "Chat";

        private string _scanState = "Idle";
        public string ScanState
        {
            get => _scanState;
            set
            {
                if (SetProperty(ref _scanState, value))
                {
                    OnPropertyChanged(nameof(IsIdle));
                    OnPropertyChanged(nameof(IsScanning));
                    OnPropertyChanged(nameof(IsCompleted));
                    OnPropertyChanged(nameof(IsError));
                    OnPropertyChanged(nameof(CanStartScan));
                    OnPropertyChanged(nameof(ShowScanButtons));
                    OnPropertyChanged(nameof(HasAnyScan));
                }
            }
        }

        public bool IsIdle => ScanState == "Idle";
        public bool IsScanning => ScanState == "Scanning";
        public bool IsCompleted => ScanState == "Completed";
        public bool IsError => ScanState == "Error";
        public bool CanStartScan => !IsScanning;
        public bool ShowScanButtons => IsCompleted || IsError;
        public bool HasAnyScan => ScanHistory.Count > 0;

        private int _progress;
        public int Progress
        {
            get => _progress;
            set => SetProperty(ref _progress, value);
        }

        private int _progressCount;
        public int ProgressCount
        {
            get => _progressCount;
            set => SetProperty(ref _progressCount, value);
        }

        private string _currentSection = string.Empty;
        public string CurrentSection
        {
            get => _currentSection;
            set => SetProperty(ref _currentSection, value);
        }

        private string _currentStep = "Prêt à analyser";
        public string CurrentStep
        {
            get => _currentStep;
            set => SetProperty(ref _currentStep, value);
        }

        private string _statusMessage = "Cliquez sur ANALYSER pour démarrer le diagnostic";
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        private string _errorMessage = string.Empty;
        public string ErrorMessage
        {
            get => _errorMessage;
            set => SetProperty(ref _errorMessage, value);
        }

        private ScanResult? _scanResult;
        public ScanResult? ScanResult
        {
            get => _scanResult;
            set
            {
                if (SetProperty(ref _scanResult, value))
                {
                    OnPropertyChanged(nameof(HasScanResult));
                    OnPropertyChanged(nameof(ScoreDisplay));
                    OnPropertyChanged(nameof(GradeDisplay));
                    OnPropertyChanged(nameof(StatusWithScore));
                }
            }
        }

        public bool HasScanResult => ScanResult != null && ScanResult.IsValid;
        public string ScoreDisplay => ScanResult?.Summary?.Score.ToString() ?? "0";
        public string GradeDisplay => ScanResult?.Summary?.Grade ?? "N/A";
        public string StatusWithScore => HasScanResult 
            ? $"Score: {ScanResult!.Summary.Score}/100 | Grade: {ScanResult.Summary.Grade}" 
            : "Aucun scan effectué";

        private ScanHistoryItem? _selectedHistoryScan;
        public ScanHistoryItem? SelectedHistoryScan
        {
            get => _selectedHistoryScan;
            set
            {
                if (SetProperty(ref _selectedHistoryScan, value))
                {
                    OnPropertyChanged(nameof(IsViewingHistoryDetail));
                    if (value != null && value.Result != null)
                    {
                        ScanResult = value.Result;
                        UpdateScanItemsFromResult(value.Result);
                    }
                }
            }
        }

        public bool IsViewingHistoryDetail => SelectedHistoryScan != null && IsResultsView;

        private bool _isAdmin;
        public bool IsAdmin
        {
            get => _isAdmin;
            set => SetProperty(ref _isAdmin, value);
        }

        private string _elapsedTime = "00:00";
        public string ElapsedTime
        {
            get => _elapsedTime;
            set => SetProperty(ref _elapsedTime, value);
        }

        // Paramètres
        private string _reportDirectory = string.Empty;
        public string ReportDirectory
        {
            get => _reportDirectory;
            set
            {
                if (SetProperty(ref _reportDirectory, value) && !_isLoadingSettings)
                {
                    IsSettingsDirty = true;
                }
            }
        }

        private bool _isSettingsDirty = false;
        public bool IsSettingsDirty
        {
            get => _isSettingsDirty;
            set => SetProperty(ref _isSettingsDirty, value);
        }

        // Collections
        public ObservableCollection<string> LiveFeedItems { get; } = new ObservableCollection<string>();
        public ObservableCollection<ScanItem> ScanItems { get; } = new ObservableCollection<ScanItem>();
        public ObservableCollection<ScanHistoryItem> ScanHistory { get; } = new ObservableCollection<ScanHistoryItem>();

        #endregion

        #region Commands

        public ICommand StartScanCommand { get; }
        public ICommand CancelScanCommand { get; }
        public ICommand OpenReportCommand { get; }
        public ICommand RestartAsAdminCommand { get; }
        public ICommand ExportResultsCommand { get; }
        public ICommand NavigateToScannerCommand { get; }
        public ICommand NavigateToResultsCommand { get; }
        public ICommand NavigateToSettingsCommand { get; }
        public ICommand NavigateToHealthcheckCommand { get; }
        public ICommand NavigateToChatCommand { get; }
        public ICommand BrowseReportDirectoryCommand { get; }
        public ICommand SaveSettingsCommand { get; }
        public ICommand SelectHistoryScanCommand { get; }
        public ICommand BackToHistoryCommand { get; }

        #endregion

        #region Constructor

        public MainViewModel()
        {
            _powerShellService = new PowerShellService();
            _reportParserService = new ReportParserService();
            _scanStopwatch = new Stopwatch();

            _liveFeedTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _liveFeedTimer.Tick += (s, e) => UpdateElapsedTime();

            // Initialiser les chemins relatifs
            _scriptPath = Path.Combine(_baseDir, "Scripts", "Total_PS_PC_Scan.ps1");
            _reportsDir = Path.Combine(_baseDir, "Rapports");
            _resultJsonPath = Path.Combine(_reportsDir, "scan_result.json");
            _configPath = Path.Combine(_baseDir, "config.json");

            // Créer le dossier Rapports s'il n'existe pas
            if (!Directory.Exists(_reportsDir))
            {
                try
                {
                    Directory.CreateDirectory(_reportsDir);
                }
                catch { }
            }

            IsAdmin = AdminService.IsRunningAsAdmin();

            // Charger les paramètres
            LoadSettings();

            // Initialiser les commandes
            StartScanCommand = new AsyncRelayCommand(StartScanAsync, () => CanStartScan);
            CancelScanCommand = new RelayCommand(CancelScan, () => IsScanning);
            OpenReportCommand = new RelayCommand(OpenReport, () => HasScanResult);
            RestartAsAdminCommand = new RelayCommand(RestartAsAdmin);
            ExportResultsCommand = new RelayCommand(ExportResults, () => HasScanResult);
            NavigateToScannerCommand = new RelayCommand(() => { CurrentView = "Home"; SelectedHistoryScan = null; });
            NavigateToResultsCommand = new RelayCommand(() => { CurrentView = "Results"; SelectedHistoryScan = null; }, () => HasAnyScan);
            NavigateToSettingsCommand = new RelayCommand(() => { CurrentView = "Settings"; SelectedHistoryScan = null; });
            NavigateToHealthcheckCommand = new RelayCommand(() => { CurrentView = "Healthcheck"; SelectedHistoryScan = null; });
            NavigateToChatCommand = new RelayCommand(() => { CurrentView = "Chat"; SelectedHistoryScan = null; });
            BrowseReportDirectoryCommand = new RelayCommand(BrowseReportDirectory);
            SaveSettingsCommand = new RelayCommand(SaveSettings, () => IsSettingsDirty);
            SelectHistoryScanCommand = new RelayCommand<ScanHistoryItem>(SelectHistoryScan);
            BackToHistoryCommand = new RelayCommand(BackToHistory);

            // S'abonner aux événements
            _powerShellService.OutputReceived += OnOutputReceived;
            _powerShellService.ProgressChanged += OnProgressChanged;
            _powerShellService.StepChanged += OnStepChanged;

            if (!IsAdmin)
            {
                StatusMessage = "⚠️ Droits administrateur requis pour un scan complet";
            }

            App.LogMessage("MainViewModel initialisé");
        }

        #endregion

        #region Methods

        private async Task StartScanAsync()
        {
            lock (_scanLock)
            {
                if (_scanProcess != null && !_scanProcess.HasExited)
                {
                    App.LogMessage("Scan déjà en cours");
                    return;
                }
            }

            try
            {
                // Vérifier que le script existe
                if (!File.Exists(_scriptPath))
                {
                    ErrorMessage = $"Script introuvable: {_scriptPath}";
                    StatusMessage = "❌ Script PowerShell introuvable";
                    ScanState = "Error";
                    App.LogMessage($"Script non trouvé: {_scriptPath}");
                    return;
                }

                // Vérifier/Créer le dossier Rapports
                if (!Directory.Exists(_reportsDir))
                {
                    try
                    {
                        Directory.CreateDirectory(_reportsDir);
                    }
                    catch (Exception ex)
                    {
                        ErrorMessage = $"Impossible de créer le dossier Rapports: {ex.Message}";
                        StatusMessage = "❌ Erreur création dossier";
                        ScanState = "Error";
                        return;
                    }
                }

                if (!IsAdmin)
                {
                    var result = MessageBox.Show(
                        "L'application nécessite les droits administrateur pour un scan complet.\n\nVoulez-vous redémarrer en mode administrateur?",
                        "Droits insuffisants",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Warning);

                    if (result == MessageBoxResult.Yes)
                    {
                        RestartAsAdmin();
                        return;
                    }
                }

                // Réinitialiser
                ScanState = "Scanning";
                Progress = 0;
                ProgressCount = 0;
                CurrentStep = "Initialisation...";
                CurrentSection = string.Empty;
                StatusMessage = "🔄 Analyse en cours...";
                ErrorMessage = string.Empty;
                LiveFeedItems.Clear();
                ScanItems.Clear();
                ScanResult = null;

                _scanStopwatch.Restart();
                _liveFeedTimer.Start();

                AddLiveFeedItem("▶ Démarrage du scan...");

                App.LogMessage("Démarrage du scan");

                // Créer CancellationTokenSource
                _scanCts = new CancellationTokenSource();

                var outputBuilder = new StringBuilder();
                var errorBuilder = new StringBuilder();

                // Lancer le processus PowerShell
                var startInfo = new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{_scriptPath}\" -OutputDir \"{_reportsDir}\"",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };

                _scanProcess = new Process { StartInfo = startInfo };
                _scanProcess.EnableRaisingEvents = true;

                // CORRECTION: Utiliser les événements DataReceived au lieu de ReadLineAsync
                _scanProcess.OutputDataReceived += (sender, e) =>
                {
                    if (string.IsNullOrEmpty(e.Data)) return;
                    
                    Application.Current?.Dispatcher.Invoke(() =>
                    {
                        outputBuilder.AppendLine(e.Data);
                        ProcessOutputLine(e.Data);
                    });
                };

                _scanProcess.ErrorDataReceived += (sender, e) =>
                {
                    if (!string.IsNullOrEmpty(e.Data))
                    {
                        Application.Current?.Dispatcher.Invoke(() =>
                        {
                            errorBuilder.AppendLine(e.Data);
                            App.LogMessage($"ERREUR PS: {e.Data}");
                        });
                    }
                };

                _scanProcess.Start();
                _scanProcess.BeginOutputReadLine();
                _scanProcess.BeginErrorReadLine();

                // Attendre la fin du processus
                await _scanProcess.WaitForExitAsync(_scanCts.Token);

                _scanStopwatch.Stop();
                _liveFeedTimer.Stop();

                var exitCode = _scanProcess.ExitCode;

                if (exitCode != 0 && errorBuilder.Length > 0)
                {
                    App.LogMessage($"Script terminé avec erreur: {errorBuilder}");
                }

                AddLiveFeedItem("✅ Scan terminé");

                // Lire le JSON
                if (File.Exists(_resultJsonPath))
                {
                    await LoadJsonResultAsync();
                }
                else
                {
                    ErrorMessage = $"Rapport JSON non trouvé: {_resultJsonPath}";
                    StatusMessage = "⚠️ Scan terminé mais rapport JSON introuvable";
                    ScanState = "Completed";
                }
            }
            catch (OperationCanceledException)
            {
                _scanStopwatch.Stop();
                _liveFeedTimer.Stop();
                StatusMessage = "⏹️ Analyse annulée";
                ScanState = "Idle";
                AddLiveFeedItem("⏹️ Analyse annulée");
                App.LogMessage("Scan annulé");
            }
            catch (Exception ex)
            {
                _scanStopwatch.Stop();
                _liveFeedTimer.Stop();
                ErrorMessage = ex.Message;
                StatusMessage = "❌ Erreur lors de l'analyse";
                ScanState = "Error";
                App.LogMessage($"Erreur scan: {ex.Message}");
            }
            finally
            {
                lock (_scanLock)
                {
                    _scanProcess?.Dispose();
                    _scanProcess = null;
                    _scanCts?.Dispose();
                    _scanCts = null;
                }
            }
        }

        private void ProcessOutputLine(string line)
        {
            AddLiveFeedItem(line);

            // Parser PROGRESS|<count>|<section>
            if (line.StartsWith("PROGRESS|"))
            {
                var parts = line.Split('|');
                if (parts.Length >= 3)
                {
                    if (int.TryParse(parts[1], out int count))
                    {
                        ProgressCount = count;
                        CurrentSection = parts[2];
                        CurrentStep = CurrentSection;
                        
                        // Calculer le pourcentage
                        Progress = (int)Math.Round((count / (double)_totalSteps) * 100);
                    }
                }
            }
        }

        private async Task LoadJsonResultAsync()
        {
            try
            {
                var jsonContent = await File.ReadAllTextAsync(_resultJsonPath);
                var jsonDoc = JsonDocument.Parse(jsonContent);
                var root = jsonDoc.RootElement;

                var result = new ScanResult
                {
                    IsValid = true,
                    Summary = new ScanSummary
                    {
                        ScanDate = DateTime.Now,
                        ScanDuration = _scanStopwatch.Elapsed
                    },
                    Items = new List<ScanItem>()
                };

                // Parser summary
                if (root.TryGetProperty("summary", out var summaryEl))
                {
                    result.Summary.Score = summaryEl.TryGetProperty("score", out var scoreEl) ? scoreEl.GetInt32() : 0;
                    result.Summary.Grade = summaryEl.TryGetProperty("grade", out var gradeEl) ? gradeEl.GetString() ?? "N/A" : "N/A";
                    result.Summary.CriticalCount = summaryEl.TryGetProperty("criticalCount", out var critEl) ? critEl.GetInt32() : 0;
                    result.Summary.ErrorCount = summaryEl.TryGetProperty("errorCount", out var errEl) ? errEl.GetInt32() : 0;
                    result.Summary.WarningCount = summaryEl.TryGetProperty("warningCount", out var warnEl) ? warnEl.GetInt32() : 0;
                }

                // Parser items
                if (root.TryGetProperty("items", out var itemsEl) && itemsEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var itemEl in itemsEl.EnumerateArray())
                    {
                        var severityStr = itemEl.TryGetProperty("severity", out var sevEl) ? sevEl.GetString() ?? "Info" : "Info";
                        var severity = severityStr switch
                        {
                            "Critical" => ScanSeverity.Critical,
                            "Major" => ScanSeverity.Error,
                            "Minor" => ScanSeverity.Warning,
                            _ => ScanSeverity.Info
                        };

                        result.Items.Add(new ScanItem
                        {
                            Category = itemEl.TryGetProperty("category", out var catEl) ? catEl.GetString() ?? "" : "",
                            Name = itemEl.TryGetProperty("name", out var nameEl) ? nameEl.GetString() ?? "" : "",
                            Severity = severity,
                            Detail = itemEl.TryGetProperty("detail", out var detEl) ? detEl.GetString() ?? "" : "",
                            Recommendation = itemEl.TryGetProperty("recommendation", out var recEl) ? recEl.GetString() ?? "" : ""
                        });
                    }
                }

                ScanResult = result;
                UpdateScanItemsFromResult(result);

                if (result.IsValid)
                {
                    StatusMessage = $"Score: {result.Summary.Score}/100 | Grade: {result.Summary.Grade}";
                    ScanState = "Completed";
                    AddToHistory(result);
                }
                else
                {
                    ErrorMessage = "Erreur lors du parsing JSON";
                    StatusMessage = "⚠️ Analyse terminée avec des erreurs";
                    ScanState = "Completed";
                }

                App.LogMessage($"Scan terminé: Score={result.Summary.Score}");
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Erreur lecture JSON: {ex.Message}";
                StatusMessage = "⚠️ Erreur lors du chargement du rapport";
                ScanState = "Error";
                App.LogMessage($"Erreur parsing JSON: {ex.Message}");
            }
        }

        private void UpdateScanItemsFromResult(ScanResult result)
        {
            ScanItems.Clear();
            foreach (var item in result.Items)
            {
                ScanItems.Add(item);
            }
        }

        private void AddToHistory(ScanResult result)
        {
            var historyItem = new ScanHistoryItem
            {
                ScanDate = result.Summary.ScanDate,
                Score = result.Summary.Score,
                Grade = result.Summary.Grade,
                Result = result
            };

            ScanHistory.Insert(0, historyItem);

            // Limiter à 10 scans
            while (ScanHistory.Count > 10)
            {
                ScanHistory.RemoveAt(ScanHistory.Count - 1);
            }

            OnPropertyChanged(nameof(HasAnyScan));
        }

        private void CancelScan()
        {
            try
            {
                lock (_scanLock)
                {
                    // Annuler le CancellationToken
                    _scanCts?.Cancel();

                    // Tuer le processus si encore actif
                    if (_scanProcess != null && !_scanProcess.HasExited)
                    {
                        try
                        {
                            _scanProcess.Kill(true);
                        }
                        catch (Exception ex)
                        {
                            App.LogMessage($"Erreur kill process: {ex.Message}");
                        }
                    }
                }

                _scanStopwatch.Stop();
                _liveFeedTimer.Stop();
                
                // Reset UI
                Progress = 0;
                ProgressCount = 0;
                CurrentStep = "Prêt à analyser";
                CurrentSection = string.Empty;
                StatusMessage = "⏹️ Analyse annulée";
                ScanState = "Idle";
                AddLiveFeedItem("⏹️ Analyse annulée");
                App.LogMessage("Scan annulé");
            }
            catch (Exception ex)
            {
                App.LogMessage($"Erreur annulation: {ex.Message}");
            }
        }

        private void OpenReport()
        {
            if (HasAnyScan)
            {
                CurrentView = "Results";
                if (ScanHistory.Count > 0)
                {
                    SelectedHistoryScan = ScanHistory[0];
                }
            }
        }

        private void SelectHistoryScan(ScanHistoryItem? item)
        {
            if (item != null)
            {
                SelectedHistoryScan = item;
            }
        }

        private void BackToHistory()
        {
            SelectedHistoryScan = null;
        }

        private void RestartAsAdmin()
        {
            try
            {
                AdminService.RestartAsAdmin();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Impossible de redémarrer en administrateur: {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportResults()
        {
            try
            {
                if (ScanResult == null) return;

                var dialog = new Microsoft.Win32.SaveFileDialog
                {
                    FileName = $"Diagnostic_{DateTime.Now:yyyyMMdd_HHmmss}",
                    DefaultExt = ".txt",
                    Filter = "Fichiers texte (*.txt)|*.txt|Tous les fichiers (*.*)|*.*"
                };

                if (dialog.ShowDialog() == true)
                {
                    File.WriteAllText(dialog.FileName, ScanResult.RawReport);
                    MessageBox.Show("Rapport exporté avec succès!", "Exportation", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erreur d'exportation: {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BrowseReportDirectory()
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Sélectionner le dossier des rapports",
                SelectedPath = ReportDirectory,
                ShowNewFolderButton = true
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ReportDirectory = dialog.SelectedPath;
                IsSettingsDirty = true;
            }
        }

        private void SaveSettings()
        {
            try
            {
                var config = new
                {
                    ReportDirectory = ReportDirectory
                };

                var jsonContent = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_configPath, jsonContent);
                
                IsSettingsDirty = false;
                App.LogMessage("Paramètres sauvegardés");
                MessageBox.Show("Paramètres enregistrés avec succès!", "Paramètres", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                App.LogMessage($"Erreur sauvegarde paramètres: {ex.Message}");
                MessageBox.Show($"Erreur lors de la sauvegarde: {ex.Message}", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void LoadSettings()
        {
            try
            {
                _isLoadingSettings = true;

                if (File.Exists(_configPath))
                {
                    var jsonContent = File.ReadAllText(_configPath);
                    var jsonDoc = JsonDocument.Parse(jsonContent);
                    var root = jsonDoc.RootElement;

                    if (root.TryGetProperty("ReportDirectory", out var reportDirEl))
                    {
                        _reportDirectory = reportDirEl.GetString() ?? Path.Combine(_baseDir, "Rapports");
                    }
                    else
                    {
                        _reportDirectory = Path.Combine(_baseDir, "Rapports");
                    }
                }
                else
                {
                    // Valeur par défaut
                    _reportDirectory = Path.Combine(_baseDir, "Rapports");
                }

                OnPropertyChanged(nameof(ReportDirectory));
            }
            catch (Exception ex)
            {
                App.LogMessage($"Erreur chargement paramètres: {ex.Message}");
                _reportDirectory = Path.Combine(_baseDir, "Rapports");
            }
            finally
            {
                _isLoadingSettings = false;
            }
        }


        private void OnOutputReceived(string output)
        {
            Application.Current?.Dispatcher.Invoke(() => AddLiveFeedItem(output));
        }

        private void OnProgressChanged(int progress)
        {
            Application.Current?.Dispatcher.Invoke(() => Progress = progress);
        }

        private void OnStepChanged(string step)
        {
            Application.Current?.Dispatcher.Invoke(() =>
            {
                CurrentStep = step;
                AddLiveFeedItem($"📍 {step}");
            });
        }

        private void AddLiveFeedItem(string item)
        {
            Application.Current?.Dispatcher.Invoke(() =>
            {
                LiveFeedItems.Insert(0, $"[{DateTime.Now:HH:mm:ss}] {item}");
                while (LiveFeedItems.Count > 100)
                {
                    LiveFeedItems.RemoveAt(LiveFeedItems.Count - 1);
                }
            });
        }

        private void UpdateElapsedTime()
        {
            ElapsedTime = _scanStopwatch.Elapsed.ToString(@"mm\:ss");
        }

        #endregion
    }

    /// <summary>
    /// Élément d'historique de scan
    /// </summary>
    public class ScanHistoryItem
    {
        public DateTime ScanDate { get; set; }
        public int Score { get; set; }
        public string Grade { get; set; } = "N/A";
        public ScanResult? Result { get; set; }
        public string DateDisplay => ScanDate.ToString("dd/MM/yyyy HH:mm");
        public string ScoreDisplay => $"{Score}/100 ({Grade})";
    }
}
