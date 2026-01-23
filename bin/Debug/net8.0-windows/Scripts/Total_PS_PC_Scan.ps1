#Requires -Version 5.1
<#
.SYNOPSIS
    Script de diagnostic système complet pour PC Diagnostic Pro
.DESCRIPTION
    Analyse le système et génère un rapport JSON structuré
.PARAMETER OutputDir
    Répertoire où sauvegarder le rapport JSON
#>

param(
    [Parameter(Mandatory=$false)]
    [string]$OutputDir = $PSScriptRoot
)

# Configuration
$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"

# Initialisation
$script:stepCount = 0
$script:totalSteps = 27
$script:items = @()
$script:criticalCount = 0
$script:errorCount = 0
$script:warningCount = 0

function Write-Progress-Step {
    param([string]$Section)
    $script:stepCount++
    Write-Output "PROGRESS|$($script:stepCount)|$Section"
}

function Add-ScanItem {
    param(
        [string]$Category,
        [string]$Name,
        [string]$Severity,  # Info, Minor, Major, Critical
        [string]$Detail,
        [string]$Recommendation = ""
    )
    
    $script:items += @{
        category = $Category
        name = $Name
        severity = $Severity
        detail = $Detail
        recommendation = $Recommendation
    }
    
    switch ($Severity) {
        "Critical" { $script:criticalCount++ }
        "Major" { $script:errorCount++ }
        "Minor" { $script:warningCount++ }
    }
}

# ==================== DÉBUT DU SCAN ====================

Write-Output "=== PC Diagnostic Pro - Scan démarré ==="
Write-Output ""

# 1. Informations système de base
Write-Progress-Step "Informations système"
try {
    $os = Get-CimInstance Win32_OperatingSystem
    $cs = Get-CimInstance Win32_ComputerSystem
    
    Add-ScanItem -Category "Système" -Name "Nom du PC" -Severity "Info" -Detail $cs.Name
    Add-ScanItem -Category "Système" -Name "Version Windows" -Severity "Info" -Detail "$($os.Caption) - Build $($os.BuildNumber)"
    
    # Uptime
    $uptime = (Get-Date) - $os.LastBootUpTime
    $uptimeText = "$($uptime.Days) jours, $($uptime.Hours) heures"
    
    if ($uptime.Days -gt 30) {
        Add-ScanItem -Category "Système" -Name "Uptime" -Severity "Minor" -Detail $uptimeText -Recommendation "Redémarrage recommandé (uptime > 30 jours)"
    } else {
        Add-ScanItem -Category "Système" -Name "Uptime" -Severity "Info" -Detail $uptimeText
    }
} catch {
    Add-ScanItem -Category "Système" -Name "Informations système" -Severity "Major" -Detail "Erreur de lecture" -Recommendation "Vérifier les permissions"
}

# 2. CPU
Write-Progress-Step "Processeur"
try {
    $cpu = Get-CimInstance Win32_Processor
    Add-ScanItem -Category "Système" -Name "Processeur" -Severity "Info" -Detail $cpu.Name
    
    # Charge CPU
    $cpuLoad = (Get-CimInstance Win32_Processor).LoadPercentage
    if ($cpuLoad -gt 90) {
        Add-ScanItem -Category "Performance" -Name "Charge CPU" -Severity "Major" -Detail "$cpuLoad%" -Recommendation "CPU surchargé, vérifier les processus"
    } elseif ($cpuLoad -gt 70) {
        Add-ScanItem -Category "Performance" -Name "Charge CPU" -Severity "Minor" -Detail "$cpuLoad%" -Recommendation "Charge CPU élevée"
    } else {
        Add-ScanItem -Category "Performance" -Name "Charge CPU" -Severity "Info" -Detail "$cpuLoad%"
    }
} catch {
    Add-ScanItem -Category "Système" -Name "CPU" -Severity "Info" -Detail "Information non disponible"
}

# 3. Mémoire RAM
Write-Progress-Step "Mémoire RAM"
try {
    $os = Get-CimInstance Win32_OperatingSystem
    $totalRAM = [math]::Round($os.TotalVisibleMemorySize / 1MB, 1)
    $freeRAM = [math]::Round($os.FreePhysicalMemory / 1MB, 1)
    $usedPercent = [math]::Round((1 - ($freeRAM / $totalRAM)) * 100, 1)
    
    if ($usedPercent -gt 90) {
        Add-ScanItem -Category "Mémoire" -Name "Utilisation RAM" -Severity "Critical" -Detail "$usedPercent% utilisé ($totalRAM Go total)" -Recommendation "Mémoire critique! Fermer des applications"
    } elseif ($usedPercent -gt 80) {
        Add-ScanItem -Category "Mémoire" -Name "Utilisation RAM" -Severity "Major" -Detail "$usedPercent% utilisé ($totalRAM Go total)" -Recommendation "Mémoire élevée"
    } elseif ($usedPercent -gt 70) {
        Add-ScanItem -Category "Mémoire" -Name "Utilisation RAM" -Severity "Minor" -Detail "$usedPercent% utilisé ($totalRAM Go total)"
    } else {
        Add-ScanItem -Category "Mémoire" -Name "Utilisation RAM" -Severity "Info" -Detail "$usedPercent% utilisé ($totalRAM Go total)"
    }
} catch {
    Add-ScanItem -Category "Mémoire" -Name "RAM" -Severity "Info" -Detail "Information non disponible"
}

# 4-6. Disques
Write-Progress-Step "Stockage"
try {
    $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
    foreach ($disk in $disks) {
        $freePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1)
        $freeGB = [math]::Round($disk.FreeSpace / 1GB, 1)
        $totalGB = [math]::Round($disk.Size / 1GB, 1)
        
        if ($freePercent -lt 10) {
            Add-ScanItem -Category "Stockage" -Name "Disque $($disk.DeviceID)" -Severity "Critical" -Detail "$freeGB Go libre sur $totalGB Go ($freePercent%)" -Recommendation "Espace disque critique!"
        } elseif ($freePercent -lt 20) {
            Add-ScanItem -Category "Stockage" -Name "Disque $($disk.DeviceID)" -Severity "Major" -Detail "$freeGB Go libre sur $totalGB Go ($freePercent%)" -Recommendation "Libérer de l'espace disque"
        } elseif ($freePercent -lt 30) {
            Add-ScanItem -Category "Stockage" -Name "Disque $($disk.DeviceID)" -Severity "Minor" -Detail "$freeGB Go libre sur $totalGB Go ($freePercent%)"
        } else {
            Add-ScanItem -Category "Stockage" -Name "Disque $($disk.DeviceID)" -Severity "Info" -Detail "$freeGB Go libre sur $totalGB Go ($freePercent%)"
        }
    }
} catch {
    Add-ScanItem -Category "Stockage" -Name "Disques" -Severity "Info" -Detail "Information non disponible"
}

# 7-10. Services critiques
Write-Progress-Step "Services Windows"
$criticalServices = @(
    @{Name="wuauserv"; DisplayName="Windows Update"},
    @{Name="WinDefend"; DisplayName="Windows Defender"},
    @{Name="MpsSvc"; DisplayName="Pare-feu Windows"},
    @{Name="BITS"; DisplayName="Service de transfert intelligent"}
)

foreach ($svc in $criticalServices) {
    Write-Progress-Step "Service $($svc.DisplayName)"
    try {
        $service = Get-Service -Name $svc.Name -ErrorAction SilentlyContinue
        if ($service) {
            if ($service.Status -eq "Running") {
                Add-ScanItem -Category "Services" -Name $svc.DisplayName -Severity "Info" -Detail "En cours d'exécution"
            } else {
                Add-ScanItem -Category "Services" -Name $svc.DisplayName -Severity "Major" -Detail "Arrêté" -Recommendation "Démarrer le service $($svc.DisplayName)"
            }
        } else {
            Add-ScanItem -Category "Services" -Name $svc.DisplayName -Severity "Minor" -Detail "Non trouvé"
        }
    } catch {
        Add-ScanItem -Category "Services" -Name $svc.DisplayName -Severity "Info" -Detail "Non vérifié"
    }
}

# 11-13. Réseau
Write-Progress-Step "Connectivité réseau"
try {
    $networkAdapters = Get-NetAdapter | Where-Object { $_.Status -eq "Up" }
    if ($networkAdapters) {
        Add-ScanItem -Category "Réseau" -Name "Adaptateurs actifs" -Severity "Info" -Detail "$($networkAdapters.Count) adaptateur(s) connecté(s)"
    } else {
        Add-ScanItem -Category "Réseau" -Name "Adaptateurs" -Severity "Major" -Detail "Aucun adaptateur actif" -Recommendation "Vérifier la connexion réseau"
    }
} catch {
    Add-ScanItem -Category "Réseau" -Name "Adaptateurs" -Severity "Info" -Detail "Information non disponible"
}

Write-Progress-Step "Test connexion Internet"
try {
    $pingTest = Test-Connection -ComputerName "8.8.8.8" -Count 1 -Quiet
    if ($pingTest) {
        Add-ScanItem -Category "Réseau" -Name "Connexion Internet" -Severity "Info" -Detail "Connecté"
    } else {
        Add-ScanItem -Category "Réseau" -Name "Connexion Internet" -Severity "Major" -Detail "Non connecté" -Recommendation "Vérifier la connexion Internet"
    }
} catch {
    Add-ScanItem -Category "Réseau" -Name "Connexion Internet" -Severity "Minor" -Detail "Test non effectué"
}

Write-Progress-Step "Configuration DNS"
try {
    $dns = Get-DnsClientServerAddress -AddressFamily IPv4 | Where-Object { $_.ServerAddresses } | Select-Object -First 1
    if ($dns) {
        Add-ScanItem -Category "Réseau" -Name "DNS" -Severity "Info" -Detail ($dns.ServerAddresses -join ", ")
    }
} catch {
    Add-ScanItem -Category "Réseau" -Name "DNS" -Severity "Info" -Detail "Information non disponible"
}

# 14-16. Sécurité
Write-Progress-Step "Windows Defender"
try {
    $defender = Get-MpComputerStatus -ErrorAction SilentlyContinue
    if ($defender) {
        if ($defender.RealTimeProtectionEnabled) {
            Add-ScanItem -Category "Sécurité" -Name "Protection temps réel" -Severity "Info" -Detail "Activée"
        } else {
            Add-ScanItem -Category "Sécurité" -Name "Protection temps réel" -Severity "Critical" -Detail "Désactivée" -Recommendation "Activer Windows Defender"
        }
        
        if ($defender.AntivirusSignatureAge -gt 7) {
            Add-ScanItem -Category "Sécurité" -Name "Signatures antivirus" -Severity "Major" -Detail "$($defender.AntivirusSignatureAge) jours" -Recommendation "Mettre à jour les définitions"
        } else {
            Add-ScanItem -Category "Sécurité" -Name "Signatures antivirus" -Severity "Info" -Detail "À jour ($($defender.AntivirusSignatureAge) jours)"
        }
    }
} catch {
    Add-ScanItem -Category "Sécurité" -Name "Windows Defender" -Severity "Info" -Detail "Non vérifié"
}

Write-Progress-Step "Pare-feu"
try {
    $firewall = Get-NetFirewallProfile | Where-Object { $_.Enabled -eq $true }
    if ($firewall) {
        Add-ScanItem -Category "Sécurité" -Name "Pare-feu" -Severity "Info" -Detail "Actif ($($firewall.Count) profil(s))"
    } else {
        Add-ScanItem -Category "Sécurité" -Name "Pare-feu" -Severity "Critical" -Detail "Désactivé" -Recommendation "Activer le pare-feu Windows"
    }
} catch {
    Add-ScanItem -Category "Sécurité" -Name "Pare-feu" -Severity "Info" -Detail "Non vérifié"
}

# 17-19. Windows Update
Write-Progress-Step "Windows Update"
try {
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $pendingUpdates = $updateSearcher.Search("IsInstalled=0").Updates.Count
    
    if ($pendingUpdates -gt 10) {
        Add-ScanItem -Category "Mises à jour" -Name "Windows Update" -Severity "Major" -Detail "$pendingUpdates mises à jour en attente" -Recommendation "Installer les mises à jour Windows"
    } elseif ($pendingUpdates -gt 0) {
        Add-ScanItem -Category "Mises à jour" -Name "Windows Update" -Severity "Minor" -Detail "$pendingUpdates mises à jour en attente" -Recommendation "Mises à jour disponibles"
    } else {
        Add-ScanItem -Category "Mises à jour" -Name "Windows Update" -Severity "Info" -Detail "Système à jour"
    }
} catch {
    Add-ScanItem -Category "Mises à jour" -Name "Windows Update" -Severity "Info" -Detail "Non vérifié"
}

# 20-22. Applications au démarrage
Write-Progress-Step "Applications au démarrage"
try {
    $startupApps = Get-CimInstance Win32_StartupCommand
    $startupCount = ($startupApps | Measure-Object).Count
    
    if ($startupCount -gt 20) {
        Add-ScanItem -Category "Performance" -Name "Programmes au démarrage" -Severity "Major" -Detail "$startupCount programmes" -Recommendation "Désactiver les programmes inutiles"
    } elseif ($startupCount -gt 10) {
        Add-ScanItem -Category "Performance" -Name "Programmes au démarrage" -Severity "Minor" -Detail "$startupCount programmes"
    } else {
        Add-ScanItem -Category "Performance" -Name "Programmes au démarrage" -Severity "Info" -Detail "$startupCount programmes"
    }
} catch {
    Add-ScanItem -Category "Performance" -Name "Démarrage" -Severity "Info" -Detail "Non vérifié"
}

# 23-25. Fichiers temporaires
Write-Progress-Step "Fichiers temporaires"
try {
    $tempPath = $env:TEMP
    $tempSize = (Get-ChildItem -Path $tempPath -Recurse -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    $tempSizeMB = [math]::Round($tempSize / 1MB, 1)
    
    if ($tempSizeMB -gt 1000) {
        Add-ScanItem -Category "Maintenance" -Name "Fichiers temporaires" -Severity "Minor" -Detail "$tempSizeMB Mo" -Recommendation "Nettoyage recommandé"
    } else {
        Add-ScanItem -Category "Maintenance" -Name "Fichiers temporaires" -Severity "Info" -Detail "$tempSizeMB Mo"
    }
} catch {
    Add-ScanItem -Category "Maintenance" -Name "Fichiers temporaires" -Severity "Info" -Detail "Non vérifié"
}

# 26. Événements système critiques
Write-Progress-Step "Événements système"
try {
    $criticalEvents = Get-WinEvent -FilterHashtable @{LogName='System'; Level=1,2; StartTime=(Get-Date).AddDays(-7)} -MaxEvents 10 -ErrorAction SilentlyContinue
    $criticalCount = ($criticalEvents | Measure-Object).Count
    
    if ($criticalCount -gt 5) {
        Add-ScanItem -Category "Système" -Name "Événements critiques" -Severity "Major" -Detail "$criticalCount erreurs critiques (7 derniers jours)" -Recommendation "Vérifier les journaux système"
    } elseif ($criticalCount -gt 0) {
        Add-ScanItem -Category "Système" -Name "Événements critiques" -Severity "Minor" -Detail "$criticalCount erreurs critiques (7 derniers jours)"
    } else {
        Add-ScanItem -Category "Système" -Name "Événements critiques" -Severity "Info" -Detail "Aucune erreur critique récente"
    }
} catch {
    Add-ScanItem -Category "Système" -Name "Événements" -Severity "Info" -Detail "Non vérifié"
}

# 27. Finalisation
Write-Progress-Step "Génération du rapport"

# Calcul du score
$score = 100
$score -= $script:criticalCount * 25
$score -= $script:errorCount * 10
$score -= $script:warningCount * 5
$score = [Math]::Max(0, [Math]::Min(100, $score))

# Déterminer le grade
$grade = switch ($score) {
    { $_ -ge 90 } { "A" }
    { $_ -ge 75 } { "B" }
    { $_ -ge 60 } { "C" }
    default { "D" }
}

# Créer le résultat JSON
$result = @{
    summary = @{
        score = $score
        grade = $grade
        criticalCount = $script:criticalCount
        errorCount = $script:errorCount
        warningCount = $script:warningCount
        scanDate = (Get-Date).ToString("o")
    }
    items = $script:items
}

# Sauvegarder le JSON
$jsonPath = Join-Path $OutputDir "scan_result.json"
$result | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8 -Force

Write-Output ""
Write-Output "=== Scan terminé ==="
Write-Output "Score: $score/100 (Grade: $grade)"
Write-Output "Critique: $($script:criticalCount) | Majeur: $($script:errorCount) | Mineur: $($script:warningCount)"
Write-Output "Rapport sauvegardé: $jsonPath"
