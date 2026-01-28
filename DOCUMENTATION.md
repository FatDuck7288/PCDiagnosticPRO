# Intégration PowerShell + HardwareProbe

## Exécuter un scan depuis l'application
- Lancez l'application en mode administrateur.
- Cliquez sur **ANALYSER**. Le script PowerShell officiel est exécuté localement via `powershell.exe -NoProfile -ExecutionPolicy Bypass`.
- Le script génère son rapport **TXT** et le **JSON** d’origine dans le dossier des rapports (par défaut `Rapports`).

## Emplacements des fichiers générés
- Rapport PowerShell JSON (source de vérité) : `Rapports/scan_result.json`.
- Snapshot final fusionné : `Rapports/Snapshot_Final_<timestamp>.json`.

## Fonctionnement de la fusion
1. Le script PowerShell produit son JSON d’origine sans aucune modification.
2. Le module `VirtualIT.HardwareProbe` collecte localement des métriques hardware (températures, GPU, VRAM, carte mère, ventilateurs).
3. `FinalSnapshotBuilder` ajoute un nœud racine `hardwareProbe` au JSON d’origine et écrit un nouveau fichier `Snapshot_Final_*.json`.
4. En cas d’échec du module hardware, le JSON final est tout de même généré (best effort) avec `status=ERROR` et des erreurs détaillées.

## Tests rapides (manuel)
- Exécution admin OK : lancer un scan en mode administrateur.
- Si `HardwareProbe` échoue (capteurs indisponibles), le scan PowerShell doit toujours se terminer.
- Vérifier que `Snapshot_Final_*.json` est bien généré dans le dossier des rapports.
