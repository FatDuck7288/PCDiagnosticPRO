using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Windows.Media;
using System.Windows.Data;
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
        private bool _cancelHandled;

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

        private readonly Dictionary<string, Dictionary<string, string>> _localizedStrings = new()
        {
            ["fr"] = new Dictionary<string, string>
            {
                ["HomeTitle"] = "PC Diagnostic PRO",
                ["HomeSubtitle"] = "Outil de diagnostic système professionnel",
                ["HomeScanTitle"] = "Scan et Fix",
                ["HomeScanAction"] = "Action : Lancer un diagnostic",
                ["HomeScanDescription"] = "Analysez votre PC et corrigez les problèmes",
                ["HomeChatTitle"] = "Chat et Support",
                ["HomeChatAction"] = "Action : Ouvrir l'assistance",
                ["HomeChatDescription"] = "Discutez avec l'IA pour résoudre vos problèmes",
                ["NavHomeTooltip"] = "Tableau de bord",
                ["NavScanTooltip"] = "Scan Healthcheck",
                ["NavReportsTooltip"] = "Rapports",
                ["NavSettingsTooltip"] = "Paramètres",
                ["HealthProgressTitle"] = "Progression",
                ["ElapsedTimeLabel"] = "Temps écoulé",
                ["ConfigsScannedLabel"] = "Configurations scannées",
                ["CurrentSectionLabel"] = "Section courante",
                ["LiveFeedLabel"] = "Flux en direct",
                ["ReportButtonText"] = "Rapport",
                ["ExportButtonText"] = "Exporter",
                ["ScanButtonText"] = "ANALYSER",
                ["ScanButtonSubtext"] = "Cliquez pour démarrer",
                ["CancelButtonText"] = "Arrêt",
                ["ChatTitle"] = "Chat et Support",
                ["ChatSubtitle"] = "Cette fonctionnalité sera disponible prochainement",
                ["ResultsHistoryTitle"] = "Historique des scans",
                ["ResultsDetailTitle"] = "Résultats du diagnostic",
                ["ResultsScanDateFormat"] = "Scan du {0}",
                ["ResultsDetailsHeader"] = "Détail des éléments analysés",
                ["ResultsBackButton"] = "← Retour",
                ["ResultsCategoryHeader"] = "Catégorie",
                ["ResultsItemHeader"] = "Élément",
                ["ResultsLevelHeader"] = "Niveau",
                ["ResultsDetailHeader"] = "Détail",
                ["ResultsRecommendationHeader"] = "Recommandation",
                ["SettingsTitle"] = "Paramètres",
                ["ReportsDirectoryTitle"] = "Répertoire des rapports",
                ["ReportsDirectoryDescription"] = "Sélectionnez le dossier où les rapports seront recherchés.",
                ["BrowseButtonText"] = "Parcourir...",
                ["AdminRightsTitle"] = "Droits administrateur",
                ["AdminStatusLabel"] = "Statut actuel: ",
                ["AdminNoText"] = "NON ADMIN",
                ["AdminYesText"] = "ADMINISTRATEUR",
                ["RestartAdminButtonText"] = "🔐 Relancer en administrateur",
                ["SaveSettingsButtonText"] = "💾 Enregistrer",
                ["LanguageTitle"] = "Langue de l'application",
                ["LanguageDescription"] = "Choisissez la langue de l'interface.",
                ["LanguageLabel"] = "Langue",
                ["ReadyToScan"] = "Prêt à analyser",
                ["StatusReady"] = "Cliquez sur ANALYSER pour démarrer le diagnostic",
                ["AdminRequiredWarning"] = "⚠️ Droits administrateur requis pour un scan complet",
                ["InitStep"] = "Initialisation...",
                ["StatusScanning"] = "🔄 Analyse en cours...",
                ["StatusScriptMissing"] = "❌ Script PowerShell introuvable",
                ["StatusFolderError"] = "❌ Erreur création dossier",
                ["StatusCanceled"] = "⏹️ Analyse annulée",
                ["StatusScanError"] = "❌ Erreur lors de l'analyse",
                ["StatusJsonMissing"] = "⚠️ Scan terminé mais rapport JSON introuvable",
                ["StatusParsingError"] = "⚠️ Analyse terminée avec des erreurs",
                ["StatusLoadReportError"] = "⚠️ Erreur lors du chargement du rapport",
                ["ArchivesButtonText"] = "Archives",
                ["ArchivesTitle"] = "Archives",
                ["ArchiveMenuText"] = "Archiver",
                ["DeleteMenuText"] = "Supprimer",
                ["ScoreLegendTitle"] = "Légende / Calcul du score",
                ["ScoreRulesTitle"] = "Règles de score",
                ["ScoreGradesTitle"] = "Grades",
                ["ScoreRuleInitial"] = "• Score initial : 100",
                ["ScoreRuleCritical"] = "• Erreurs critiques : criticalCount × 25 (−25 par critique)",
                ["ScoreRuleError"] = "• Erreurs : errorCount × 10 (−10 par erreur)",
                ["ScoreRuleWarning"] = "• Avertissements : warningCount × 5 (−5 par avertissement)",
                ["ScoreRuleMin"] = "• Score min : 0",
                ["ScoreRuleMax"] = "• Score max : 100",
                ["ScoreGradeA"] = "• ❤️ ≥ 90 : A",
                ["ScoreGradeB"] = "• 👍 ≥ 75 et < 90 : B",
                ["ScoreGradeC"] = "• ⚠️ ≥ 60 et < 75 : C",
                ["ScoreGradeD"] = "• 💀 < 60 : D",
                ["DeleteScanConfirmTitle"] = "Confirmation",
                ["DeleteScanConfirmMessage"] = "Voulez-vous vraiment supprimer ce scan ?"
            },
            ["en"] = new Dictionary<string, string>
            {
                ["HomeTitle"] = "PC Diagnostic PRO",
                ["HomeSubtitle"] = "Professional system diagnostic tool",
                ["HomeScanTitle"] = "Scan & Fix",
                ["HomeScanAction"] = "Action: Run a diagnostic",
                ["HomeScanDescription"] = "Analyze your PC and fix issues",
                ["HomeChatTitle"] = "Chat & Support",
                ["HomeChatAction"] = "Action: Open support",
                ["HomeChatDescription"] = "Chat with AI to resolve your issues",
                ["NavHomeTooltip"] = "Dashboard",
                ["NavScanTooltip"] = "Healthcheck scan",
                ["NavReportsTooltip"] = "Reports",
                ["NavSettingsTooltip"] = "Settings",
                ["HealthProgressTitle"] = "Progress",
                ["ElapsedTimeLabel"] = "Elapsed time",
                ["ConfigsScannedLabel"] = "Scanned configurations",
                ["CurrentSectionLabel"] = "Current section",
                ["LiveFeedLabel"] = "Live Feed",
                ["ReportButtonText"] = "Report",
                ["ExportButtonText"] = "Export",
                ["ScanButtonText"] = "SCAN",
                ["ScanButtonSubtext"] = "Click to start",
                ["CancelButtonText"] = "Stop",
                ["ChatTitle"] = "Chat & Support",
                ["ChatSubtitle"] = "This feature will be available soon",
                ["ResultsHistoryTitle"] = "Scan history",
                ["ResultsDetailTitle"] = "Diagnostic results",
                ["ResultsScanDateFormat"] = "Scan from {0}",
                ["ResultsDetailsHeader"] = "Detailed analyzed items",
                ["ResultsBackButton"] = "← Back",
                ["ResultsCategoryHeader"] = "Category",
                ["ResultsItemHeader"] = "Item",
                ["ResultsLevelHeader"] = "Level",
                ["ResultsDetailHeader"] = "Detail",
                ["ResultsRecommendationHeader"] = "Recommendation",
                ["SettingsTitle"] = "Settings",
                ["ReportsDirectoryTitle"] = "Reports directory",
                ["ReportsDirectoryDescription"] = "Select the folder where reports will be searched.",
                ["BrowseButtonText"] = "Browse...",
                ["AdminRightsTitle"] = "Administrator rights",
                ["AdminStatusLabel"] = "Current status: ",
                ["AdminNoText"] = "NOT ADMIN",
                ["AdminYesText"] = "ADMINISTRATOR",
                ["RestartAdminButtonText"] = "🔐 Restart as administrator",
                ["SaveSettingsButtonText"] = "💾 Save",
                ["LanguageTitle"] = "Application language",
                ["LanguageDescription"] = "Choose the interface language.",
                ["LanguageLabel"] = "Language",
                ["ReadyToScan"] = "Ready to scan",
                ["StatusReady"] = "Click SCAN to start the diagnostic",
                ["AdminRequiredWarning"] = "⚠️ Administrator rights required for a full scan",
                ["InitStep"] = "Initializing...",
                ["StatusScanning"] = "🔄 Scan in progress...",
                ["StatusScriptMissing"] = "❌ PowerShell script not found",
                ["StatusFolderError"] = "❌ Error creating folder",
                ["StatusCanceled"] = "⏹️ Scan canceled",
                ["StatusScanError"] = "❌ Error during scan",
                ["StatusJsonMissing"] = "⚠️ Scan completed but JSON report not found",
                ["StatusParsingError"] = "⚠️ Scan completed with errors",
                ["StatusLoadReportError"] = "⚠️ Error while loading the report",
                ["ArchivesButtonText"] = "Archives",
                ["ArchivesTitle"] = "Archives",
                ["ArchiveMenuText"] = "Archive",
                ["DeleteMenuText"] = "Delete",
                ["ScoreLegendTitle"] = "Legend / Score calculation",
                ["ScoreRulesTitle"] = "Score rules",
                ["ScoreGradesTitle"] = "Grades",
                ["ScoreRuleInitial"] = "• Starting score: 100",
                ["ScoreRuleCritical"] = "• Critical errors: criticalCount × 25 (-25 per critical)",
                ["ScoreRuleError"] = "• Errors: errorCount × 10 (-10 per error)",
                ["ScoreRuleWarning"] = "• Warnings: warningCount × 5 (-5 per warning)",
                ["ScoreRuleMin"] = "• Minimum score: 0",
                ["ScoreRuleMax"] = "• Maximum score: 100",
                ["ScoreGradeA"] = "• ❤️ ≥ 90 : A",
                ["ScoreGradeB"] = "• 👍 ≥ 75 and < 90 : B",
                ["ScoreGradeC"] = "• ⚠️ ≥ 60 and < 75 : C",
                ["ScoreGradeD"] = "• 💀 < 60 : D",
                ["DeleteScanConfirmTitle"] = "Confirmation",
                ["DeleteScanConfirmMessage"] = "Do you really want to delete this scan?"
            },
            ["es"] = new Dictionary<string, string>
            {
                ["HomeTitle"] = "PC Diagnostic PRO",
                ["HomeSubtitle"] = "Herramienta profesional de diagnóstico del sistema",
                ["HomeScanTitle"] = "Escanear y reparar",
                ["HomeScanAction"] = "Acción: Ejecutar un diagnóstico",
                ["HomeScanDescription"] = "Analice su PC y corrija los problemas",
                ["HomeChatTitle"] = "Chat y soporte",
                ["HomeChatAction"] = "Acción: Abrir soporte",
                ["HomeChatDescription"] = "Chatee con la IA para resolver sus problemas",
                ["NavHomeTooltip"] = "Panel",
                ["NavScanTooltip"] = "Escaneo de salud",
                ["NavReportsTooltip"] = "Informes",
                ["NavSettingsTooltip"] = "Configuración",
                ["HealthProgressTitle"] = "Progreso",
                ["ElapsedTimeLabel"] = "Tiempo transcurrido",
                ["ConfigsScannedLabel"] = "Configuraciones escaneadas",
                ["CurrentSectionLabel"] = "Sección actual",
                ["LiveFeedLabel"] = "Feed en vivo",
                ["ReportButtonText"] = "Informe",
                ["ExportButtonText"] = "Exportar",
                ["ScanButtonText"] = "ESCANEAR",
                ["ScanButtonSubtext"] = "Haga clic para iniciar",
                ["CancelButtonText"] = "Detener",
                ["ChatTitle"] = "Chat y soporte",
                ["ChatSubtitle"] = "Esta función estará disponible pronto",
                ["ResultsHistoryTitle"] = "Historial de escaneos",
                ["ResultsDetailTitle"] = "Resultados del diagnóstico",
                ["ResultsScanDateFormat"] = "Escaneo del {0}",
                ["ResultsDetailsHeader"] = "Detalle de elementos analizados",
                ["ResultsBackButton"] = "← Volver",
                ["ResultsCategoryHeader"] = "Categoría",
                ["ResultsItemHeader"] = "Elemento",
                ["ResultsLevelHeader"] = "Nivel",
                ["ResultsDetailHeader"] = "Detalle",
                ["ResultsRecommendationHeader"] = "Recomendación",
                ["SettingsTitle"] = "Configuración",
                ["ReportsDirectoryTitle"] = "Directorio de informes",
                ["ReportsDirectoryDescription"] = "Seleccione la carpeta donde se buscarán los informes.",
                ["BrowseButtonText"] = "Examinar...",
                ["AdminRightsTitle"] = "Permisos de administrador",
                ["AdminStatusLabel"] = "Estado actual: ",
                ["AdminNoText"] = "SIN ADMIN",
                ["AdminYesText"] = "ADMINISTRADOR",
                ["RestartAdminButtonText"] = "🔐 Reiniciar como administrador",
                ["SaveSettingsButtonText"] = "💾 Guardar",
                ["LanguageTitle"] = "Idioma de la aplicación",
                ["LanguageDescription"] = "Elija el idioma de la interfaz.",
                ["LanguageLabel"] = "Idioma",
                ["ReadyToScan"] = "Listo para escanear",
                ["StatusReady"] = "Haga clic en ESCANEAR para iniciar el diagnóstico",
                ["AdminRequiredWarning"] = "⚠️ Se requieren permisos de administrador para un análisis completo",
                ["InitStep"] = "Inicializando...",
                ["StatusScanning"] = "🔄 Análisis en curso...",
                ["StatusScriptMissing"] = "❌ Script de PowerShell no encontrado",
                ["StatusFolderError"] = "❌ Error al crear la carpeta",
                ["StatusCanceled"] = "⏹️ Análisis cancelado",
                ["StatusScanError"] = "❌ Error durante el análisis",
                ["StatusJsonMissing"] = "⚠️ Escaneo completado pero no se encontró el informe JSON",
                ["StatusParsingError"] = "⚠️ Análisis completado con errores",
                ["StatusLoadReportError"] = "⚠️ Error al cargar el informe",
                ["ArchivesButtonText"] = "Archivos",
                ["ArchivesTitle"] = "Archivos",
                ["ArchiveMenuText"] = "Archivar",
                ["DeleteMenuText"] = "Eliminar",
                ["ScoreLegendTitle"] = "Leyenda / Cálculo del puntaje",
                ["ScoreRulesTitle"] = "Reglas de puntaje",
                ["ScoreGradesTitle"] = "Calificaciones",
                ["ScoreRuleInitial"] = "• Puntaje inicial: 100",
                ["ScoreRuleCritical"] = "• Errores críticos: criticalCount × 25 (-25 por crítico)",
                ["ScoreRuleError"] = "• Errores: errorCount × 10 (-10 por error)",
                ["ScoreRuleWarning"] = "• Advertencias: warningCount × 5 (-5 por advertencia)",
                ["ScoreRuleMin"] = "• Puntaje mínimo: 0",
                ["ScoreRuleMax"] = "• Puntaje máximo: 100",
                ["ScoreGradeA"] = "• ❤️ ≥ 90 : A",
                ["ScoreGradeB"] = "• 👍 ≥ 75 y < 90 : B",
                ["ScoreGradeC"] = "• ⚠️ ≥ 60 y < 75 : C",
                ["ScoreGradeD"] = "• 💀 < 60 : D",
                ["DeleteScanConfirmTitle"] = "Confirmación",
                ["DeleteScanConfirmMessage"] = "¿Desea eliminar este escaneo?"
            }
        };

        private bool _isUpdatingLanguage;

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
                    OnPropertyChanged(nameof(IsViewingHistoryDetail));
                    OnPropertyChanged(nameof(IsViewingHistoryList));
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
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        public bool IsIdle => ScanState == "Idle";
        public bool IsScanning => ScanState == "Scanning";
        public bool IsCompleted => ScanState == "Completed";
        public bool IsError => ScanState == "Error";
        public bool CanStartScan => !IsScanning;
        public bool ShowScanButtons => IsCompleted || IsError;
        public bool HasAnyScan => ScanHistory.Count > 0 || ArchivedScanHistory.Count > 0;

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
                    OnPropertyChanged(nameof(IsViewingHistoryList));
                    OnPropertyChanged(nameof(SelectedScanDateDisplay));
                    if (value != null && value.Result != null)
                    {
                        ScanResult = value.Result;
                        UpdateScanItemsFromResult(value.Result);
                    }
                }
            }
        }

        public bool IsViewingHistoryDetail => SelectedHistoryScan != null && IsResultsView;

        private bool _isViewingArchives;
        public bool IsViewingArchives
        {
            get => _isViewingArchives;
            set
            {
                if (SetProperty(ref _isViewingArchives, value))
                {
                    OnPropertyChanged(nameof(IsViewingHistoryList));
                }
            }
        }

        public bool IsViewingHistoryList => !IsViewingHistoryDetail && !IsViewingArchives && IsResultsView;

        private bool _isAdmin;
        public bool IsAdmin
        {
            get => _isAdmin;
            set
            {
                if (SetProperty(ref _isAdmin, value))
                {
                    OnPropertyChanged(nameof(AdminStatusText));
                    OnPropertyChanged(nameof(AdminStatusForeground));
                }
            }
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

        private string _currentLanguage = "fr";
        public string CurrentLanguage
        {
            get => _currentLanguage;
            set
            {
                if (SetProperty(ref _currentLanguage, value))
                {
                    UpdateLocalizedStrings();
                    if (!_isUpdatingLanguage)
                    {
                        _isUpdatingLanguage = true;
                        SelectedLanguage = AvailableLanguages.FirstOrDefault(l => l.Code == value)
                                           ?? AvailableLanguages.First();
                        _isUpdatingLanguage = false;
                    }

                    if (!_isLoadingSettings)
                    {
                        IsSettingsDirty = true;
                    }
                }
            }
        }

        public ObservableCollection<LanguageOption> AvailableLanguages { get; } =
            new ObservableCollection<LanguageOption>
            {
                new LanguageOption { Code = "fr", DisplayName = "Français" },
                new LanguageOption { Code = "en", DisplayName = "English" },
                new LanguageOption { Code = "es", DisplayName = "Español" }
            };

        private LanguageOption? _selectedLanguage;
        public LanguageOption? SelectedLanguage
        {
            get => _selectedLanguage;
            set
            {
                if (SetProperty(ref _selectedLanguage, value) && value != null)
                {
                    if (!_isUpdatingLanguage)
                    {
                        _isUpdatingLanguage = true;
                        CurrentLanguage = value.Code;
                        _isUpdatingLanguage = false;
                    }

                    if (!_isLoadingSettings)
                    {
                        IsSettingsDirty = true;
                    }
                }
            }
        }

        public string HomeTitle => GetString("HomeTitle");
        public string HomeSubtitle => GetString("HomeSubtitle");
        public string HomeScanTitle => GetString("HomeScanTitle");
        public string HomeScanAction => GetString("HomeScanAction");
        public string HomeScanDescription => GetString("HomeScanDescription");
        public string HomeChatTitle => GetString("HomeChatTitle");
        public string HomeChatAction => GetString("HomeChatAction");
        public string HomeChatDescription => GetString("HomeChatDescription");
        public string NavHomeTooltip => GetString("NavHomeTooltip");
        public string NavScanTooltip => GetString("NavScanTooltip");
        public string NavReportsTooltip => GetString("NavReportsTooltip");
        public string NavSettingsTooltip => GetString("NavSettingsTooltip");
        public string HealthProgressTitle => GetString("HealthProgressTitle");
        public string ElapsedTimeLabel => GetString("ElapsedTimeLabel");
        public string ConfigsScannedLabel => GetString("ConfigsScannedLabel");
        public string CurrentSectionLabel => GetString("CurrentSectionLabel");
        public string LiveFeedLabel => GetString("LiveFeedLabel");
        public string ReportButtonText => GetString("ReportButtonText");
        public string ExportButtonText => GetString("ExportButtonText");
        public string ScanButtonText => GetString("ScanButtonText");
        public string ScanButtonSubtext => GetString("ScanButtonSubtext");
        public string CancelButtonText => GetString("CancelButtonText");
        public string ChatTitle => GetString("ChatTitle");
        public string ChatSubtitle => GetString("ChatSubtitle");
        public string ResultsHistoryTitle => GetString("ResultsHistoryTitle");
        public string ResultsDetailTitle => GetString("ResultsDetailTitle");
        public string ResultsDetailsHeader => GetString("ResultsDetailsHeader");
        public string ResultsBackButton => GetString("ResultsBackButton");
        public string ResultsCategoryHeader => GetString("ResultsCategoryHeader");
        public string ResultsItemHeader => GetString("ResultsItemHeader");
        public string ResultsLevelHeader => GetString("ResultsLevelHeader");
        public string ResultsDetailHeader => GetString("ResultsDetailHeader");
        public string ResultsRecommendationHeader => GetString("ResultsRecommendationHeader");
        public string SettingsTitle => GetString("SettingsTitle");
        public string ReportsDirectoryTitle => GetString("ReportsDirectoryTitle");
        public string ReportsDirectoryDescription => GetString("ReportsDirectoryDescription");
        public string BrowseButtonText => GetString("BrowseButtonText");
        public string AdminRightsTitle => GetString("AdminRightsTitle");
        public string AdminStatusLabel => GetString("AdminStatusLabel");
        public string AdminStatusText => IsAdmin ? GetString("AdminYesText") : GetString("AdminNoText");
        public Brush AdminStatusForeground => IsAdmin
            ? new SolidColorBrush(Color.FromRgb(46, 213, 115))
            : new SolidColorBrush(Color.FromRgb(255, 71, 87));
        public string RestartAdminButtonText => GetString("RestartAdminButtonText");
        public string SaveSettingsButtonText => GetString("SaveSettingsButtonText");
        public string LanguageTitle => GetString("LanguageTitle");
        public string LanguageDescription => GetString("LanguageDescription");
        public string LanguageLabel => GetString("LanguageLabel");
        public string ArchivesButtonText => GetString("ArchivesButtonText");
        public string ArchivesTitle => GetString("ArchivesTitle");
        public string ArchiveMenuText => GetString("ArchiveMenuText");
        public string DeleteMenuText => GetString("DeleteMenuText");
        public string ScoreLegendTitle => GetString("ScoreLegendTitle");
        public string ScoreRulesTitle => GetString("ScoreRulesTitle");
        public string ScoreGradesTitle => GetString("ScoreGradesTitle");
        public string ScoreRuleInitial => GetString("ScoreRuleInitial");
        public string ScoreRuleCritical => GetString("ScoreRuleCritical");
        public string ScoreRuleError => GetString("ScoreRuleError");
        public string ScoreRuleWarning => GetString("ScoreRuleWarning");
        public string ScoreRuleMin => GetString("ScoreRuleMin");
        public string ScoreRuleMax => GetString("ScoreRuleMax");
        public string ScoreGradeA => GetString("ScoreGradeA");
        public string ScoreGradeB => GetString("ScoreGradeB");
        public string ScoreGradeC => GetString("ScoreGradeC");
        public string ScoreGradeD => GetString("ScoreGradeD");
        public string SelectedScanDateDisplay => SelectedHistoryScan != null
            ? string.Format(GetString("ResultsScanDateFormat"), SelectedHistoryScan.DateDisplay)
            : string.Empty;

        // Collections
        public ObservableCollection<string> LiveFeedItems { get; } = new ObservableCollection<string>();
        public ObservableCollection<ScanItem> ScanItems { get; } = new ObservableCollection<ScanItem>();
        public ObservableCollection<ScanHistoryItem> ScanHistory { get; } = new ObservableCollection<ScanHistoryItem>();
        public ObservableCollection<ScanHistoryItem> ArchivedScanHistory { get; } = new ObservableCollection<ScanHistoryItem>();
        public ICollectionView ArchivedScanHistoryView { get; }

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
        public ICommand NavigateToArchivesCommand { get; }
        public ICommand ArchiveScanCommand { get; }
        public ICommand DeleteScanCommand { get; }

        #endregion

        #region Constructor

        public MainViewModel()
        {
            _powerShellService = new PowerShellService();
            _reportParserService = new ReportParserService();
            _scanStopwatch = new Stopwatch();

            ArchivedScanHistoryView = CollectionViewSource.GetDefaultView(ArchivedScanHistory);
            ArchivedScanHistoryView.GroupDescriptions.Add(new PropertyGroupDescription(nameof(ScanHistoryItem.MonthYearDisplay)));
            ArchivedScanHistoryView.SortDescriptions.Add(new SortDescription(nameof(ScanHistoryItem.ScanDate), ListSortDirection.Descending));

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
            _isUpdatingLanguage = true;
            SelectedLanguage = AvailableLanguages.FirstOrDefault(l => l.Code == CurrentLanguage)
                               ?? AvailableLanguages.First();
            _isUpdatingLanguage = false;
            UpdateLocalizedStrings();

            // Initialiser les commandes
            StartScanCommand = new AsyncRelayCommand(StartScanAsync, () => CanStartScan);
            CancelScanCommand = new RelayCommand(CancelScan, () => IsScanning);
            OpenReportCommand = new RelayCommand(OpenReport, () => HasScanResult);
            RestartAsAdminCommand = new RelayCommand(RestartAsAdmin);
            ExportResultsCommand = new RelayCommand(ExportResults, () => HasScanResult);
            NavigateToScannerCommand = new RelayCommand(() => { CurrentView = "Home"; SelectedHistoryScan = null; IsViewingArchives = false; });
            NavigateToResultsCommand = new RelayCommand(() => { CurrentView = "Results"; SelectedHistoryScan = null; IsViewingArchives = false; }, () => HasAnyScan);
            NavigateToSettingsCommand = new RelayCommand(() => { CurrentView = "Settings"; SelectedHistoryScan = null; IsViewingArchives = false; });
            NavigateToHealthcheckCommand = new RelayCommand(() => { CurrentView = "Healthcheck"; SelectedHistoryScan = null; IsViewingArchives = false; });
            NavigateToChatCommand = new RelayCommand(() => { CurrentView = "Chat"; SelectedHistoryScan = null; IsViewingArchives = false; });
            BrowseReportDirectoryCommand = new RelayCommand(BrowseReportDirectory);
            SaveSettingsCommand = new RelayCommand(SaveSettings, () => IsSettingsDirty);
            SelectHistoryScanCommand = new RelayCommand<ScanHistoryItem>(SelectHistoryScan);
            BackToHistoryCommand = new RelayCommand(BackToHistory);
            NavigateToArchivesCommand = new RelayCommand(NavigateToArchives, () => ScanHistory.Count > 0 || ArchivedScanHistory.Count > 0);
            ArchiveScanCommand = new RelayCommand<ScanHistoryItem>(ArchiveScan, item => item != null);
            DeleteScanCommand = new RelayCommand<ScanHistoryItem>(DeleteScan, item => item != null);

            ScanHistory.CollectionChanged += OnHistoryCollectionChanged;
            ArchivedScanHistory.CollectionChanged += OnHistoryCollectionChanged;

            // S'abonner aux événements
            _powerShellService.OutputReceived += OnOutputReceived;
            _powerShellService.ProgressChanged += OnProgressChanged;
            _powerShellService.StepChanged += OnStepChanged;

            if (!IsAdmin)
            {
                StatusMessage = GetString("AdminRequiredWarning");
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
                    StatusMessage = GetString("StatusScriptMissing");
                    ScanState = "Error";
                    App.LogMessage($"Script non trouvé: {_scriptPath}");
                    return;
                }

                var outputDir = string.IsNullOrWhiteSpace(ReportDirectory) ? _reportsDir : ReportDirectory;
                _resultJsonPath = Path.Combine(outputDir, "scan_result.json");

                // Vérifier/Créer le dossier Rapports
                if (!Directory.Exists(outputDir))
                {
                    try
                    {
                        Directory.CreateDirectory(outputDir);
                    }
                    catch (Exception ex)
                    {
                        ErrorMessage = $"Impossible de créer le dossier Rapports: {ex.Message}";
                        StatusMessage = GetString("StatusFolderError");
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
                CurrentStep = GetString("InitStep");
                CurrentSection = string.Empty;
                StatusMessage = GetString("StatusScanning");
                ErrorMessage = string.Empty;
                LiveFeedItems.Clear();
                ScanItems.Clear();
                ScanResult = null;
                _cancelHandled = false;

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
                    Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{_scriptPath}\" -OutputDir \"{outputDir}\"",
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
                    StatusMessage = GetString("StatusJsonMissing");
                    ScanState = "Completed";
                }
            }
            catch (OperationCanceledException)
            {
                if (!_cancelHandled)
                {
                    ResetAfterCancel();
                    _cancelHandled = true;
                }
                App.LogMessage("Scan annulé");
            }
            catch (Exception ex)
            {
                _scanStopwatch.Stop();
                _liveFeedTimer.Stop();
                ErrorMessage = ex.Message;
                StatusMessage = GetString("StatusScanError");
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
                var jsonContent = await File.ReadAllTextAsync(_resultJsonPath, Encoding.UTF8);
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
                    Items = new List<ScanItem>(),
                    RawReport = jsonContent,
                    ReportFilePath = _resultJsonPath
                };

                // Parser summary
                if (root.TryGetProperty("summary", out var summaryEl))
                {
                    result.Summary.Score = summaryEl.TryGetProperty("score", out var scoreEl) ? scoreEl.GetInt32() : 0;
                    result.Summary.Grade = summaryEl.TryGetProperty("grade", out var gradeEl) ? gradeEl.GetString() ?? "N/A" : "N/A";
                    result.Summary.CriticalCount = summaryEl.TryGetProperty("criticalCount", out var critEl) ? critEl.GetInt32() : 0;
                    result.Summary.ErrorCount = summaryEl.TryGetProperty("errorCount", out var errEl) ? errEl.GetInt32() : 0;
                    result.Summary.WarningCount = summaryEl.TryGetProperty("warningCount", out var warnEl) ? warnEl.GetInt32() : 0;

                    if (summaryEl.TryGetProperty("scanDate", out var dateEl))
                    {
                        if (DateTimeOffset.TryParse(dateEl.GetString(), out var parsedDate))
                        {
                            result.Summary.ScanDate = parsedDate.LocalDateTime;
                        }
                    }
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

                        var status = severity switch
                        {
                            ScanSeverity.Critical => "CRITIQUE",
                            ScanSeverity.Error => "ERREUR",
                            ScanSeverity.Warning => "AVERTISSEMENT",
                            _ => "INFO"
                        };

                        result.Items.Add(new ScanItem
                        {
                            Category = itemEl.TryGetProperty("category", out var catEl) ? catEl.GetString() ?? "" : "",
                            Name = itemEl.TryGetProperty("name", out var nameEl) ? nameEl.GetString() ?? "" : "",
                            Severity = severity,
                            Status = status,
                            Detail = itemEl.TryGetProperty("detail", out var detEl) ? detEl.GetString() ?? "" : "",
                            Recommendation = itemEl.TryGetProperty("recommendation", out var recEl) ? recEl.GetString() ?? "" : ""
                        });
                    }
                }

                result.Summary.TotalItems = result.Items.Count;
                result.Summary.OkCount = result.Items.Count(item => item.Severity == ScanSeverity.Info);

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
                    StatusMessage = GetString("StatusParsingError");
                    ScanState = "Completed";
                }

                App.LogMessage($"Scan terminé: Score={result.Summary.Score}");
            }
            catch (Exception ex)
            {
                ErrorMessage = $"Erreur lecture JSON: {ex.Message}";
                StatusMessage = GetString("StatusLoadReportError");
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

                if (!_cancelHandled)
                {
                    ResetAfterCancel();
                    _cancelHandled = true;
                }
                App.LogMessage("Scan annulé");
            }
            catch (Exception ex)
            {
                App.LogMessage($"Erreur annulation: {ex.Message}");
            }
        }

        private void ResetAfterCancel()
        {
            _scanStopwatch.Stop();
            _liveFeedTimer.Stop();

            // Reset UI
            Progress = 0;
            ProgressCount = 0;
            CurrentStep = GetString("ReadyToScan");
            CurrentSection = string.Empty;
            StatusMessage = GetString("StatusCanceled");
            ScanState = "Idle";
            AddLiveFeedItem("⏹️ Analyse annulée");
        }

        private void OpenReport()
        {
            if (HasAnyScan)
            {
                CurrentView = "Results";
                if (ScanHistory.Count > 0)
                {
                    IsViewingArchives = false;
                    SelectedHistoryScan = ScanHistory[0];
                }
            }
        }

        private void SelectHistoryScan(ScanHistoryItem? item)
        {
            if (item != null)
            {
                IsViewingArchives = false;
                SelectedHistoryScan = item;
            }
        }

        private void BackToHistory()
        {
            SelectedHistoryScan = null;
            IsViewingArchives = false;
        }

        private void NavigateToArchives()
        {
            SelectedHistoryScan = null;
            IsViewingArchives = true;
        }

        private void ArchiveScan(ScanHistoryItem? item)
        {
            if (item == null) return;

            if (ScanHistory.Remove(item))
            {
                ArchivedScanHistory.Insert(0, item);
                SelectedHistoryScan = null;
                IsViewingArchives = true;
                OnPropertyChanged(nameof(HasAnyScan));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private void DeleteScan(ScanHistoryItem? item)
        {
            if (item == null) return;

            var confirmation = MessageBox.Show(
                GetString("DeleteScanConfirmMessage"),
                GetString("DeleteScanConfirmTitle"),
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirmation != MessageBoxResult.Yes)
            {
                return;
            }

            if (SelectedHistoryScan == item)
            {
                SelectedHistoryScan = null;
            }

            if (ScanHistory.Remove(item))
            {
                OnPropertyChanged(nameof(HasAnyScan));
            }
            else if (ArchivedScanHistory.Remove(item))
            {
                OnPropertyChanged(nameof(HasAnyScan));
            }

            CommandManager.InvalidateRequerySuggested();
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
                    File.WriteAllText(dialog.FileName, ScanResult.RawReport, Encoding.UTF8);
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
                    ReportDirectory = ReportDirectory,
                    Language = CurrentLanguage
                };

                var jsonContent = JsonSerializer.Serialize(config, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_configPath, jsonContent, Encoding.UTF8);
                
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
                    var jsonContent = File.ReadAllText(_configPath, Encoding.UTF8);
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

                    if (root.TryGetProperty("Language", out var languageEl))
                    {
                        CurrentLanguage = languageEl.GetString() ?? "fr";
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

        private string GetString(string key)
        {
            if (_localizedStrings.TryGetValue(CurrentLanguage, out var languageSet) &&
                languageSet.TryGetValue(key, out var value))
            {
                return value;
            }

            if (_localizedStrings.TryGetValue("fr", out var fallback) &&
                fallback.TryGetValue(key, out var fallbackValue))
            {
                return fallbackValue;
            }

            return key;
        }

        private void UpdateLocalizedStrings()
        {
            var properties = new[]
            {
                nameof(HomeTitle),
                nameof(HomeSubtitle),
                nameof(HomeScanTitle),
                nameof(HomeScanAction),
                nameof(HomeScanDescription),
                nameof(HomeChatTitle),
                nameof(HomeChatAction),
                nameof(HomeChatDescription),
                nameof(NavHomeTooltip),
                nameof(NavScanTooltip),
                nameof(NavReportsTooltip),
                nameof(NavSettingsTooltip),
                nameof(HealthProgressTitle),
                nameof(ElapsedTimeLabel),
                nameof(ConfigsScannedLabel),
                nameof(CurrentSectionLabel),
                nameof(LiveFeedLabel),
                nameof(ReportButtonText),
                nameof(ExportButtonText),
                nameof(ScanButtonText),
                nameof(ScanButtonSubtext),
                nameof(CancelButtonText),
                nameof(ChatTitle),
                nameof(ChatSubtitle),
                nameof(ResultsHistoryTitle),
                nameof(ResultsDetailTitle),
                nameof(ResultsDetailsHeader),
                nameof(ResultsBackButton),
                nameof(ResultsCategoryHeader),
                nameof(ResultsItemHeader),
                nameof(ResultsLevelHeader),
                nameof(ResultsDetailHeader),
                nameof(ResultsRecommendationHeader),
                nameof(SettingsTitle),
                nameof(ReportsDirectoryTitle),
                nameof(ReportsDirectoryDescription),
                nameof(BrowseButtonText),
                nameof(AdminRightsTitle),
                nameof(AdminStatusLabel),
                nameof(AdminStatusText),
                nameof(AdminStatusForeground),
                nameof(RestartAdminButtonText),
                nameof(SaveSettingsButtonText),
                nameof(LanguageTitle),
                nameof(LanguageDescription),
                nameof(LanguageLabel),
                nameof(ArchivesButtonText),
                nameof(ArchivesTitle),
                nameof(ArchiveMenuText),
                nameof(DeleteMenuText),
                nameof(ScoreLegendTitle),
                nameof(ScoreRulesTitle),
                nameof(ScoreGradesTitle),
                nameof(ScoreRuleInitial),
                nameof(ScoreRuleCritical),
                nameof(ScoreRuleError),
                nameof(ScoreRuleWarning),
                nameof(ScoreRuleMin),
                nameof(ScoreRuleMax),
                nameof(ScoreGradeA),
                nameof(ScoreGradeB),
                nameof(ScoreGradeC),
                nameof(ScoreGradeD),
                nameof(SelectedScanDateDisplay)
            };

            foreach (var prop in properties)
            {
                OnPropertyChanged(prop);
            }

            if (IsIdle)
            {
                CurrentStep = GetString("ReadyToScan");
                StatusMessage = IsAdmin ? GetString("StatusReady") : GetString("AdminRequiredWarning");
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

        private void OnHistoryCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(HasAnyScan));
            ArchivedScanHistoryView.Refresh();
            CommandManager.InvalidateRequerySuggested();
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
        public string DateDisplay => ScanDate.ToString("dd/MM/yyyy HH:mm", CultureInfo.CurrentCulture);
        public string DayDisplay => ScanDate.ToString("dd", CultureInfo.CurrentCulture);
        public string MonthYearDisplay => ScanDate.ToString("MMMM yyyy", CultureInfo.CurrentCulture);
        public string ScoreDisplay => $"{Score}/100 ({Grade})";
    }
}
