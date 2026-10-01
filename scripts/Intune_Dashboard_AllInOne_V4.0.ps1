# ============================================================
# Script : Intune Dashboard All-In-One - v4.0
#          (socle "Script_Global_V3.3_Pages" + module "Proactivité / Santé des postes")
# Description : Interroge Microsoft Graph (REST, sans module Microsoft.Graph) et
#               génère, au choix, un dashboard HTML interactif (PSWriteHTML, pages
#               à onglets collants) et/ou des exports CSV : conformité, chiffrement,
#               applications, update rings, hardware, santé et proactivité des postes
#
# REPÈRES DE FUSION (rechercher ces balises dans le fichier) :
#   [FUSION v4 - Proactivité] : logique reprise du script "Script_proactivité_V2"
#                               (Endpoint Analytics, disque, démarrage, écrans bleus,
#                               batteries, inactivité, BitLocker, Defender, mises à
#                               jour, fiabilité applicative, profils, conformité)
#   [v4.0]                    : nouveau code propre à la fusion (REST, throttling
#                               partagé, $batch, HTML basculable, exports CSV...)
#   [v3.x] / [MODIF v3] / [UI v3.x] : historique du socle V3.3, conservé
# Nouveautés v2 : options plateformes, paliers inactifs configurables,
#                 affichage Low Storage au choix
# Nouveautés v3 : section "Non Encrypted" améliorée :
#                 - filtre (case à cocher dans le dashboard, cochée par
#                   défaut) masquant les postes inactifs 30j+ avec
#                   compteur mis à jour dynamiquement
#                 - colonne "RootCause" : cause exacte de non-conformité
#                   (BitLocker, Firewall, version OS...) récupérée via
#                   l'API Graph (deviceCompliancePolicyStates)
# Nouveautés v3.1 : refonte visuelle (police Inter, palette harmonisée,
#                   icônes SVG, hero glassmorphism, cartes Update Rings,
#                   bandeau Non Encrypted, CSS global) - logique inchangée
# Nouveautés v3.2 : dashboard découpé en 4 pages virtuelles dans un seul
#                   fichier HTML (Vue d'ensemble, Sécurité & conformité,
#                   Mises à jour & applications, Optimisation du parc) :
#                   menu d'onglets collant, transitions, liens directs
#                   (#securite...), bouton Précédent, impression de
#                   toutes les pages - aucune donnée supprimée
# Nouveautés v3.3 : page "Sécurité & conformité" enrichie par l'analyse des
#                   postes non conformes (logique reprise du script de
#                   référence "Non_compliant.ps1") :
#                   - données interrogées en direct via l'API Graph (rapport
#                     Intune "Noncompliant devices and settings", repli
#                     automatique sur l'API par appareil) : plus d'export
#                     manuel ni de fichier intermédiaire
#                   - classement par catégorie et par raison, ancienneté de
#                     synchronisation (0-5 j / 6-29 j / 30 j et +),
#                     actionnabilité (hors réseau) et tri par urgence
#                   - KPI, graphiques et synthèses dans le dashboard, export
#                     CSV (indicateurs, catégories, raisons, postes, détail)
#                   - appels Graph fiabilisés : nouvelles tentatives sur
#                     429 / 5xx / erreur réseau (Retry-After, backoff),
#                     renouvellement du jeton, messages d'erreur explicites
# Nouveautés v4.0 (All-In-One) :
#                   - nouvelle page "Santé & proactivité" (14 vérifications du
#                     script Proactivité), qui absorbe "Optimisation du parc" :
#                     paliers d'inactivité conservés, espace disque jugé en %
#                   - collecte 100 % REST : plus de module Microsoft.Graph
#                     (démarrage plus rapide, $select sur les appareils) ;
#                     PSWriteHTML n'est chargé que si le HTML est demandé
#                   - génération HTML désactivable : exécution des requêtes
#                     Graph et exports CSV seuls (dossier horodaté)
#                   - export CSV de toutes les sections collectées
#                   - throttling Graph partagé entre flux parallèles (un 429 met
#                     tous les flux en pause), Retry-After respecté, $batch pour
#                     les appels par appareil, fenêtre réactive pendant les pauses
#                   - journal C:\temp\dashboard-log.txt, pseudonymes stables
# Auteur : ECONOCOM
# ============================================================

# ===== IMPORTS ET ASSEMBLIES =====
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===== [FUSION v4 - Proactivité] RÉGLAGES RÉSEAU (.NET) =====
# .NET Framework n'autorise que DEUX connexions HTTP simultanées vers un même hôte :
# sans cette ligne, quatre flux de collecte parallèles n'en font travailler que deux.
try { [System.Net.ServicePointManager]::DefaultConnectionLimit = 24 } catch { }
# TLS 1.2 ajouté sans retirer ce qui est déjà négociable (TLS 1.3 sur les OS récents).
try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 } catch { }

# ===== CONFIGURATION GLOBALE =====
$ConfigFolder = "C:\temp\clients-id"
$OutputFolder = "C:\temp"
# [FUSION v4 - Proactivité] Journal d'exécution (rotation à 5 Mo) : à joindre à toute
# demande de diagnostic plutôt que des captures d'écran de la console.
$LogFile      = "C:\temp\dashboard-log.txt"

# ===== [MODIF v3] SEUIL D'INACTIVITÉ - SECTION "NON ENCRYPTED" =====
# Nombre de jours sans synchronisation Intune au-delà duquel un poste est
# considéré comme inactif dans la section "Non Encrypted" du dashboard.
# Utilisé par le filtre (case à cocher) et le compteur dynamique.
$NonEncryptedInactiveDays = 30

# ===== [v3.3] ANALYSE DES NON-CONFORMITÉS (PAGE "SÉCURITÉ & CONFORMITÉ") =====
# Logique reprise du script de référence "Non_compliant.ps1" ; les données ne
# proviennent plus d'un export manuel mais de l'API Graph.
# Seuils d'ancienneté de la dernière synchronisation Intune :
$NcSeuilVert       = 5     # <= 5 jours   -> synchronisation récente, priorité haute
$NcSeuilOrange     = 29    # 6 à 29 jours -> vigilance
$NcSeuilRouge      = 30    # >= 30 jours  -> poste fantôme
$NcSeuilHorsReseau = 15    # > 15 jours + erreur "réseau" -> action impossible

# Mots-clés d'une erreur qui EXIGE que le poste soit joignable (réseau,
# pare-feu, BitLocker, inactivité). Comparaison sans casse, accents ni
# ponctuation, sur la raison et le paramètre Intune. Liste librement extensible.
$NcKeywordsHorsReseau = @(
    "reseau", "network", "wifi", "vpn", "connexion", "connectivity",
    "pare feu", "parefeu", "firewall",
    "bitlocker", "chiffrement", "encryption",
    "inactif", "contact requis", "is active", "inactivity",
    "remaincontact", "activefirewall", "firewallrequired",
    "bitlockerenabled", "isactive"
)

# Source des données : rapport Intune "Noncompliant devices and settings"
# (quelques appels paginés pour tout le tenant). $false = analyse appareil par
# appareil uniquement (un appel par poste non conforme, plus lent).
$NcUseBulkReport  = $true
$NcReportPageSize = 500

# Export CSV (toutes sections) : séparateur ";" (Excel en français), UTF-8 avec BOM
$NcCsvDelimiter = ";"

# ===== [v3.3] RÉSILIENCE DES APPELS À L'API GRAPH =====
$GraphMaxRetries        = 5     # nouvelles tentatives sur 429 / 5xx / erreur réseau
$GraphRequestTimeoutSec = 120   # délai maximal d'une requête (secondes)

# ===== [FUSION v4 - Proactivité] PERFORMANCE DE LA COLLECTE =====
# Les listes Endpoint Analytics (scores, performances, batteries, fiabilité,
# historique de démarrage) et BitLocker sont indépendantes : elles sont lues de
# front dans un pool de runspaces. Au-delà de 4 à 6 flux, on déclenche le
# throttling Graph et le gain se retourne en perte.
# (Case "Collecte parallèle" de l'onglet Proactivité : décochée = séquentiel.)
$MaxParallelCollections    = 4
# Garde-fou de durée de la vérification "profils de configuration en erreur"
$MaxConfigProfilesAnalyzed = 300

# ===== [FUSION v4 - Proactivité] SEUILS DE LA PAGE "SANTÉ & PROACTIVITÉ" =====
# Ajustez-les librement à votre contexte.
$RemediationThresholds = @{
    DiskFreePctCritical  = 10     # % d'espace libre en-dessous duquel c'est critique
    DiskFreePctWarning   = 20     # % d'espace libre à surveiller
    DiskFreeGbCritical   = 5      # critique aussi si moins de N Go libres, quel que soit le %
    StaleDaysWarning     = 30     # jours sans synchronisation -> à surveiller
    StaleDaysCritical    = 90     # jours sans synchronisation -> critique
    ScoreLow             = 50     # score Endpoint Analytics global en-dessous duquel on alerte
    BootSlowSeconds      = 90     # démarrage (core boot) au-delà duquel on alerte
    BatteryPoor          = 50     # score batterie en-dessous duquel on alerte
    BsodCritical         = 3      # écrans bleus (fenêtre 14 j) à partir desquels c'est critique
    RestartsHigh         = 15     # redémarrages (fenêtre 14 j) jugés anormalement fréquents
    SignatureStaleDays   = 7      # ancienneté des signatures antivirus (Defender) tolérée
    UptimeWarningDays    = 14     # jours depuis le dernier démarrage connu -> à surveiller (estimation)
    AppCrashWarning      = 5      # plantages d'une même application (fenêtre ~14 j) jugés anormaux
    BatteryCapacityPoor  = 70     # capacité max restante (%) sous laquelle la batterie est à remplacer
    BatteryCapacityCrit  = 50     # capacité max restante (%) sous laquelle le remplacement est urgent
    BatteryRuntimeLowMin = 120    # autonomie estimée (minutes) sous laquelle le poste devient sédentaire
    ConfigErrorsCritical = 3      # profils de configuration en échec rendant l'alerte critique
}

# ===== COORDONNÉES ENTREPRISE PAR DÉFAUT =====
$DefaultCompanyName   = "ECONOCOM"
$DefaultContactPerson = "Nom si nécessaire"
$DefaultContactEmail  = "support-Intune-2IP@econocom.com"
$DefaultContactPhone  = "+33 ...."

# ===== CLÉ DE CHIFFREMENT AES =====
$AESKey = @(
    0x4D, 0x79, 0x53, 0x65, 0x63, 0x72, 0x65, 0x74,
    0x4B, 0x65, 0x79, 0x31, 0x32, 0x33, 0x34, 0x35,
    0x36, 0x37, 0x38, 0x39, 0x30, 0x41, 0x42, 0x43,
    0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4A, 0x4B
)

# ===== PALETTE DE COULEURS =====
# [UI v3.1] Palette harmonisée "Nuit / Indigo" : tons profonds et désaturés pour
# les en-têtes de section, couleurs sémantiques réservées aux statuts.
$Colors = @{
    Primary         = "#1e1b4b"   # Nuit (indigo profond) - en-têtes principaux, titres de graphiques
    Secondary       = "#4f46e5"   # Indigo - accent interactif
    Accent          = "#0e7490"   # Pétrole
    Success         = "#10b981"
    Warning         = "#f59e0b"
    Danger          = "#ef4444"
    Compliance      = "#4c1d95"   # Violet profond
    Applications    = "#9d174d"   # Framboise profond
    UpdateRings     = "#0e7490"   # Pétrole
    Hardware        = "#92400e"   # Ambre brûlé
    DetailTables    = "#334155"   # Ardoise - tables de détail
    Health          = "#0f766e"   # [v4.0] Sarcelle profonde - page Santé & proactivité
    # Couleurs d'identification des plateformes (cartes Overview)
    PlatformWindows = "#2563eb"
    PlatformIOS     = "#0e7490"
    PlatformAndroid = "#15803d"
    PlatformMac     = "#64748b"
}

# ===== CODE COULEUR PALIERS INACTIFS =====
# [UI v3.1] Rampe de sévérité progressive (bleu -> rouge) au lieu d'un mélange arbitraire
$InactiveColorMap = @{
    "30"  = "#0284c7"   # bleu ciel
    "60"  = "#4f46e5"   # indigo
    "90"  = "#d97706"   # ambre
    "120" = "#ea580c"   # orange
    "150" = "#dc2626"   # rouge
    "180" = "#991b1b"   # rouge sombre
}

# ===== CRÉATION DES DOSSIERS =====
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
if (-not (Test-Path $ConfigFolder)) {
    New-Item -ItemType Directory -Path $ConfigFolder | Out-Null
    Write-Host "Dossier créé : $ConfigFolder - Veuillez générer des configurations client" -ForegroundColor Yellow
}

# ========================================
# FONCTIONS DE CHIFFREMENT/DÉCHIFFREMENT
# ========================================

function Decrypt-StringAES {
    param([string]$EncryptedString)
    if ([string]::IsNullOrWhiteSpace($EncryptedString)) { return "" }
    try {
        $SecureString = ConvertTo-SecureString -String $EncryptedString -Key $AESKey
        $BSTR      = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
        $PlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        return $PlainText
    } catch {
        Write-Host "Erreur de déchiffrement AES : $($_.Exception.Message)" -ForegroundColor Red
        return ""
    }
}

# ========================================
# FONCTIONS DE GESTION DES CONFIGURATIONS
# ========================================

function Get-ClientConfig {
    param([string]$ClientName)
    if ([string]::IsNullOrWhiteSpace($ClientName) -or $ClientName -eq "-- Sélectionnez un client --") { return $null }
    $SafeClientName = $ClientName -replace '[\\/:*?"<>|]', '_' -replace '\s+', '_'
    $ConfigFile = Join-Path $ConfigFolder "$SafeClientName.clientconfig"
    if (-not (Test-Path $ConfigFile)) {
        Write-Host "Fichier de configuration introuvable : $ConfigFile" -ForegroundColor Red
        return $null
    }
    try {
        $config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json
        return [PSCustomObject]@{
            ClientName   = $config.ClientName
            TenantId     = Decrypt-StringAES -EncryptedString $config.TenantId
            ClientId     = Decrypt-StringAES -EncryptedString $config.ClientId
            ClientSecret = Decrypt-StringAES -EncryptedString $config.ClientSecret
        }
    } catch {
        Write-Host "Erreur lors de la lecture de la configuration : $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Load-ClientConfigs {
    $cmbClients.Items.Clear()
    [void]$cmbClients.Items.Add("-- Sélectionnez un client --")
    if (-not (Test-Path $ConfigFolder)) { $cmbClients.SelectedIndex = 0; $cmbClients.Enabled = $false; return }
    $configFiles = Get-ChildItem -Path $ConfigFolder -Filter "*.clientconfig" -ErrorAction SilentlyContinue
    if ($configFiles.Count -eq 0) {
        $cmbClients.SelectedIndex = 0; $cmbClients.Enabled = $false
        $lblStatus.Text = "⚠ Aucune configuration client trouvée"; $lblStatus.ForeColor = [System.Drawing.Color]::Orange
        return
    }
    foreach ($file in $configFiles) {
        try {
            $config = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
            [void]$cmbClients.Items.Add($config.ClientName)
        } catch { Write-Host "Erreur lecture config : $($file.Name)" -ForegroundColor Red }
    }
    $cmbClients.SelectedIndex = 0; $cmbClients.Enabled = $true
    $lblStatus.Text = "✓ $($configFiles.Count) configuration(s) client chargée(s)"
    $lblStatus.ForeColor = [System.Drawing.Color]::Green
}

# ========================================
# FONCTIONS UTILITAIRES
# ========================================

function Show-ErrorMessage([string]$message) {
    [System.Windows.Forms.MessageBox]::Show($message, "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
function Show-InfoMessage([string]$message) {
    [System.Windows.Forms.MessageBox]::Show($message, "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

# [v3.3] Met à jour la ligne d'état de la fenêtre (sans effet hors interface)
function Set-UiStatus([string]$Text) {
    if ($lblStatus) { $lblStatus.Text = $Text }
    if ($form)      { $form.Refresh() }
}

# [FUSION v4 - Proactivité] Journal horodaté : console + $LogFile (rotation à 5 Mo).
# Niveaux : INFO, WARN, ERROR, OK. Un échec d'écriture n'interrompt jamais le script.
function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "OK")][string]$Level = "INFO"
    )
    $line  = "[{0}] [{1}] {2}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $Level, $Message
    $color = switch ($Level) { "WARN" { "Yellow" } "ERROR" { "Red" } "OK" { "Green" } default { "Gray" } }
    Write-Host $line -ForegroundColor $color
    if (-not $LogFile) { return }
    try {
        $folder = Split-Path -Path $LogFile -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null }
        if ((Test-Path $LogFile) -and ((Get-Item $LogFile).Length -gt 5MB)) {
            $archive = Join-Path $folder ("dashboard-log_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".txt")
            Move-Item -Path $LogFile -Destination $archive -Force -ErrorAction SilentlyContinue
        }
        Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch { }
}

# [v4.0] Étape de génération : ligne d'état de la fenêtre + journal
function Write-Step([string]$Text) {
    Set-UiStatus $Text
    Write-Log $Text
}

# [FUSION v4 - Proactivité] Attente SANS geler la fenêtre WinForms : un simple
# Start-Sleep bloque le thread d'interface, Windows affiche "Ne répond pas" pendant
# un backoff de plusieurs dizaines de secondes. Ici les messages Windows sont pompés
# toutes les 200 ms et un compte à rebours s'affiche. Hors interface (runspace de
# collecte parallèle), c'est une attente ordinaire.
function Start-ResponsiveSleep {
    param([Parameter(Mandatory = $true)][int]$Seconds, [string]$Message = "Patientez")
    if ($Seconds -le 0) { return }
    if (-not $form) { Start-Sleep -Seconds $Seconds; return }
    $endTime = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $endTime) {
        $remaining = [Math]::Max(0, [int][Math]::Ceiling(($endTime - (Get-Date)).TotalSeconds))
        try {
            if ($lblStatus) { $lblStatus.Text = "$Message ($remaining s restantes)..." }
            [System.Windows.Forms.Application]::DoEvents()
        } catch { }
        Start-Sleep -Milliseconds 200
    }
}

# ========================================
# [UI v3.1] COMPOSANTS VISUELS DU DASHBOARD
# ========================================
# Fonctions de rendu uniquement : elles transforment des valeurs déjà
# calculées en fragments HTML. Aucune logique métier ici.

# Icônes SVG inline (tracés Feather / Lucide, licence MIT), colorées via currentColor
$IxIcons = @{
    'monitor'    = '<rect x="2" y="3" width="20" height="14" rx="2" ry="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/>'
    'windows'    = '<rect x="3" y="3" width="8" height="8" rx="1.5"/><rect x="13" y="3" width="8" height="8" rx="1.5"/><rect x="3" y="13" width="8" height="8" rx="1.5"/><rect x="13" y="13" width="8" height="8" rx="1.5"/>'
    'smartphone' = '<rect x="5" y="2" width="14" height="20" rx="2" ry="2"/><line x1="12" y1="18" x2="12.01" y2="18"/>'
    'bot'        = '<path d="M12 8V4H8"/><rect width="16" height="12" x="4" y="8" rx="2"/><path d="M2 14h2"/><path d="M20 14h2"/><path d="M15 13v2"/><path d="M9 13v2"/>'
    'laptop'     = '<path d="M20 16V7a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v9m16 0H4m16 0 1.28 2.55a1 1 0 0 1-.9 1.45H3.62a1 1 0 0 1-.9-1.45L4 16"/>'
    'clock'      = '<circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/>'
    'hard-drive' = '<line x1="22" y1="12" x2="2" y2="12"/><path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/><line x1="6" y1="16" x2="6.01" y2="16"/><line x1="10" y1="16" x2="10.01" y2="16"/>'
    'shield'     = '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>'
    'shield-ok'  = '<path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><polyline points="9 12 11 14 15 10"/>'
    'unlock'     = '<rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 9.9-1"/>'
    'lock'       = '<rect x="3" y="11" width="18" height="11" rx="2" ry="2"/><path d="M7 11V7a5 5 0 0 1 10 0v4"/>'
    'user'       = '<path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="12" cy="7" r="4"/>'
    'mail'       = '<path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/>'
    'phone'      = '<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/>'
    'calendar'   = '<rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/>'
    'filter'     = '<polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3"/>'
    'eye-off'    = '<path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/>'
    'check'      = '<path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/>'
    'x-circle'   = '<circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/>'
    'alert'      = '<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>'
    'refresh'    = '<polyline points="23 4 23 10 17 10"/><polyline points="1 20 1 14 7 14"/><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"/>'
    'circle'     = '<circle cx="12" cy="12" r="10"/>'
    'layout'     = '<rect width="7" height="9" x="3" y="3" rx="1"/><rect width="7" height="5" x="14" y="3" rx="1"/><rect width="7" height="9" x="14" y="12" rx="1"/><rect width="7" height="5" x="3" y="16" rx="1"/>'
    'package'    = '<line x1="16.5" y1="9.4" x2="7.5" y2="4.21"/><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/>'
    'gauge'      = '<path d="m12 14 4-4"/><path d="M3.34 19a10 10 0 1 1 17.32 0"/>'
    'printer'    = '<polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/>'
    # [v3.3] Analyse des non-conformités
    'file-text'  = '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/>'
    'database'   = '<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>'
    'info'       = '<circle cx="12" cy="12" r="10"/><line x1="12" y1="16" x2="12" y2="12"/><line x1="12" y1="8" x2="12.01" y2="8"/>'
    'wifi-off'   = '<line x1="1" y1="1" x2="23" y2="23"/><path d="M16.72 11.06A10.94 10.94 0 0 1 19 12.55"/><path d="M5 12.55a10.94 10.94 0 0 1 5.17-2.39"/><path d="M10.71 5.05A16 16 0 0 1 22.58 9"/><path d="M1.42 9a15.91 15.91 0 0 1 4.7-2.88"/><path d="M8.53 16.11a6 6 0 0 1 6.95 0"/><line x1="12" y1="20" x2="12.01" y2="20"/>'
    'list'       = '<line x1="8" y1="6" x2="21" y2="6"/><line x1="8" y1="12" x2="21" y2="12"/><line x1="8" y1="18" x2="21" y2="18"/><line x1="3" y1="6" x2="3.01" y2="6"/><line x1="3" y1="12" x2="3.01" y2="12"/><line x1="3" y1="18" x2="3.01" y2="18"/>'
    'zap'        = '<polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"/>'
    'tool'       = '<path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/>'
}

# Retourne une icône SVG inline. $Stroke est une chaîne pour éviter la virgule décimale (culture fr-FR).
function Get-IconSvg {
    param([string]$Name, [int]$Size = 20, [string]$Stroke = '2')
    $path = $IxIcons[$Name]
    if (-not $path) { $path = $IxIcons['circle'] }
    return ('<svg class="ix-icon" xmlns="http://www.w3.org/2000/svg" width="{0}" height="{0}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="{1}" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{2}</svg>' -f $Size, $Stroke, $path)
}

# Échappement HTML des valeurs texte injectées dans les blocs personnalisés
function ConvertTo-HtmlSafe {
    param([AllowNull()][object]$Value)
    return [System.Net.WebUtility]::HtmlEncode("$Value")
}

# Pourcentage entier (affichage uniquement)
function Get-IxPercent {
    param([double]$Part, [double]$Total)
    if ($Total -le 0) { return 0 }
    return [int][math]::Round(100 * $Part / $Total)
}

# Carte KPI de la section "Devices Overview"
function New-IxStatCard {
    param([string]$Label, [string]$Value, [string]$Icon, [string]$Color, [string]$Caption = "")
    $cap = if ($Caption) { "<div class=`"ix-stat__caption`">$Caption</div>" } else { "" }
    return @"
<div class="ix-stat" style="--c:$Color; --c-soft:${Color}14;">
  <div class="ix-stat__top">
    <span class="ix-stat__label">$Label</span>
    <span class="ix-stat__icon">$(Get-IconSvg -Name $Icon -Size 20 -Stroke '1.75')</span>
  </div>
  <div class="ix-stat__value">$Value</div>
  $cap
</div>
"@
}

# État vide (aucune donnée) : message orienté action plutôt qu'un simple texte
function New-IxEmptyState {
    param([string]$Icon, [string]$Title, [string]$Text = "", [ValidateSet('success', 'neutral', 'danger')][string]$Tone = 'neutral')
    $txt = if ($Text) { "<div class=`"ix-empty__text`">$Text</div>" } else { "" }
    return @"
<div class="ix-root ix-empty ix-empty--$Tone">
  <div class="ix-empty__icon">$(Get-IconSvg -Name $Icon -Size 24)</div>
  <div class="ix-empty__title">$Title</div>
  $txt
</div>
"@
}

# Feuille de style globale, injectée une seule fois en tête du dashboard.
# NB : commentaires CSS en /* */ uniquement ; aucune syntaxe [texte](lien) pour
# éviter toute interprétation par PSWriteHTML.
function Get-IxGlobalCss {
    return @"
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap">
<style>
/* ---------- Tokens ---------- */
:root {
  --ix-font: 'Inter', system-ui, -apple-system, 'Segoe UI', Roboto, Arial, sans-serif;
  --ix-bg: #f4f6fb;
  --ix-surface: #ffffff;
  --ix-ink: #0f172a;
  --ix-ink-2: #334155;
  --ix-muted: #64748b;
  --ix-border: #e2e8f0;
  --ix-border-soft: #eef2f7;
  --ix-night: $($Colors.Primary);
  --ix-indigo: $($Colors.Secondary);
  --ix-success: $($Colors.Success);
  --ix-warning: $($Colors.Warning);
  --ix-danger: $($Colors.Danger);
  --ix-shadow-sm: 0 1px 2px rgba(15,23,42,.04), 0 2px 6px -2px rgba(15,23,42,.06);
  --ix-shadow: 0 1px 2px rgba(15,23,42,.04), 0 10px 28px -10px rgba(30,27,75,.14);
  --ix-shadow-lg: 0 2px 4px rgba(15,23,42,.04), 0 22px 44px -16px rgba(30,27,75,.28);
}

/* ---------- Base ---------- */
html body { background: var(--ix-bg) !important; color: var(--ix-ink); -webkit-font-smoothing: antialiased; }
html body, html body div, html body span, html body td, html body th,
html body input, html body select, html body button, html body label, html body a {
  font-family: var(--ix-font) !important;
}
.apexcharts-text, .apexcharts-title-text, .apexcharts-legend-text, .apexcharts-datalabel, .apexcharts-tooltip {
  font-family: var(--ix-font) !important;
}
.apexcharts-title-text { font-weight: 600 !important; }

/* ---------- Sections PSWriteHTML ---------- */
.defaultSection {
  background: var(--ix-surface);
  border: 1px solid var(--ix-border) !important;
  border-radius: 16px !important;
  box-shadow: var(--ix-shadow) !important;
}
.defaultSectionHead {
  border-radius: 15px 15px 0 0;
  padding: 12px 20px !important;
  font-weight: 600 !important;
  letter-spacing: .01em;
  background-image: linear-gradient(100deg, rgba(255,255,255,.12) 0%, rgba(255,255,255,0) 55%);
}
.defaultSection .defaultPanel { box-shadow: none !important; border-radius: 12px !important; }
.defaultSection:has(.ix-hero), .defaultPanel:has(.ix-hero) {
  background: transparent !important; border: 0 !important; box-shadow: none !important;
}

/* ---------- Composants : socle ---------- */
.ix-root { font-size: 14px; line-height: 1.5; color: var(--ix-ink); text-align: left; }
.ix-root *, .ix-root *::before, .ix-root *::after { box-sizing: border-box; }
.ix-icon { display: inline-block; vertical-align: middle; flex-shrink: 0; }

/* ---------- En-tête (glassmorphism) ---------- */
.ix-hero {
  position: relative; overflow: hidden;
  border-radius: 22px; padding: 40px 44px; margin: 4px 0 6px; color: #fff;
  background:
    radial-gradient(900px 380px at 0% 0%, rgba(79,70,229,.55) 0%, rgba(79,70,229,0) 60%),
    linear-gradient(135deg, #0b1026 0%, var(--ix-night) 48%, #3b0764 100%);
  box-shadow: var(--ix-shadow-lg);
}
.ix-hero::after {
  content: ""; position: absolute; inset: 0; pointer-events: none;
  background-image: radial-gradient(rgba(255,255,255,.07) 1px, transparent 1px);
  background-size: 22px 22px;
  -webkit-mask-image: linear-gradient(90deg, transparent 0%, #000 70%);
          mask-image: linear-gradient(90deg, transparent 0%, #000 70%);
}
.ix-hero__glow { position: absolute; border-radius: 50%; filter: blur(70px); pointer-events: none; }
.ix-hero__glow--a { width: 380px; height: 380px; right: -90px; top: -170px; background: #7c3aed; opacity: .55; }
.ix-hero__glow--b { width: 320px; height: 320px; left: 38%; bottom: -220px; background: #0e7490; opacity: .45; }
.ix-hero__inner {
  position: relative; z-index: 1;
  display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 28px;
  max-width: 1400px; margin: 0 auto;
}
.ix-hero__brand { flex: 1 1 480px; min-width: 0; }
.ix-hero__brandline { display: flex; align-items: center; gap: 10px; font-size: 13px; font-weight: 500; color: rgba(255,255,255,.72); }
.ix-hero__mark {
  display: inline-flex; align-items: center; justify-content: center; width: 30px; height: 30px; border-radius: 9px;
  background: rgba(255,255,255,.12); border: 1px solid rgba(255,255,255,.22); color: #c7d2fe;
}
.ix-hero__title { margin-top: 14px; font-size: 44px; font-weight: 800; letter-spacing: -.025em; line-height: 1.02; color: #fff; }
.ix-hero__subtitle { margin-top: 10px; font-size: 16px; color: rgba(255,255,255,.78); }
.ix-hero__subtitle strong { color: #fff; font-weight: 600; }
.ix-hero__meta { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 22px; }
.ix-chip {
  display: inline-flex; align-items: center; gap: 7px; padding: 6px 12px; border-radius: 999px;
  font-size: 12.5px; font-weight: 500; color: #fff;
  background: rgba(255,255,255,.10); border: 1px solid rgba(255,255,255,.18);
  -webkit-backdrop-filter: blur(8px); backdrop-filter: blur(8px);
}
.ix-chip .ix-icon { color: #c7d2fe; }
.ix-chip--warn { background: rgba(245,158,11,.16); border-color: rgba(251,191,36,.45); color: #fde68a; }
.ix-chip--warn .ix-icon { color: #fbbf24; }
.ix-hero__contact {
  flex: 0 0 auto; min-width: 290px; padding: 20px 22px; border-radius: 18px;
  background: rgba(255,255,255,.08); border: 1px solid rgba(255,255,255,.18);
  -webkit-backdrop-filter: blur(18px) saturate(140%); backdrop-filter: blur(18px) saturate(140%);
  box-shadow: inset 0 1px 0 rgba(255,255,255,.14), 0 10px 30px -12px rgba(0,0,0,.45);
}
.ix-hero__contact-title { font-size: 13px; font-weight: 600; color: rgba(255,255,255,.66); margin-bottom: 10px; }
.ix-hero__contact-row { display: flex; align-items: center; gap: 12px; padding: 6px 0; font-size: 14px; color: #fff; }
.ix-hero__contact-row + .ix-hero__contact-row { border-top: 1px solid rgba(255,255,255,.10); }
.ix-hero__contact-row .ix-icon { color: #a5b4fc; }
.ix-hero__contact-row a { color: #fff; text-decoration: none; border-bottom: 1px dashed rgba(255,255,255,.35); }
.ix-hero__contact-row a:hover { border-bottom-color: #fff; }

/* ---------- Devices Overview ---------- */
.ix-overview { display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 16px; padding: 10px 4px 6px; }
.ix-stat {
  position: relative; overflow: hidden;
  padding: 18px 18px 16px; border-radius: 14px;
  background: linear-gradient(180deg, #ffffff 0%, #fafbfe 100%);
  border: 1px solid var(--ix-border);
  box-shadow: var(--ix-shadow-sm);
}
.ix-stat::before { content: ""; position: absolute; left: 0; right: 0; top: 0; height: 3px; background: linear-gradient(90deg, var(--c), transparent 85%); }
.ix-stat__top { display: flex; align-items: center; justify-content: space-between; gap: 12px; margin-bottom: 16px; }
.ix-stat__label { font-size: 13px; font-weight: 600; color: var(--ix-ink-2); }
.ix-stat__icon {
  display: inline-flex; align-items: center; justify-content: center;
  width: 40px; height: 40px; border-radius: 12px; background: var(--c-soft); color: var(--c);
}
.ix-stat__value { font-size: 38px; font-weight: 700; line-height: 1; letter-spacing: -.03em; color: var(--ix-ink); font-variant-numeric: tabular-nums; }
.ix-stat__caption { margin-top: 8px; font-size: 12.5px; color: var(--ix-muted); }
.ix-stat__caption b { color: var(--c); font-weight: 600; }

/* ---------- Update Rings ---------- */
.ix-rings { display: grid; grid-template-columns: repeat(auto-fit, minmax(min(340px, 100%), 1fr)); gap: 18px; padding: 12px 6px; }
.ix-ring {
  --c: var(--ix-success);
  position: relative; display: flex; flex-direction: column; gap: 16px;
  padding: 20px 20px 18px 24px; border-radius: 18px; background: #fff;
  border: 1px solid var(--ix-border); box-shadow: var(--ix-shadow-sm);
  max-width: 560px;
  transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease;
}
.ix-ring--warn   { --c: var(--ix-warning); }
.ix-ring--danger { --c: var(--ix-danger); }
.ix-ring::before { content: ""; position: absolute; left: 0; top: 20px; bottom: 20px; width: 4px; border-radius: 0 4px 4px 0; background: var(--c); }
.ix-ring:hover { transform: translateY(-4px); box-shadow: var(--ix-shadow-lg); border-color: #c7d2fe; }
.ix-ring__head { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; }
.ix-ring__name { font-size: 15px; font-weight: 600; line-height: 1.3; color: var(--ix-ink); word-break: break-word; }
.ix-badge { flex-shrink: 0; display: inline-flex; align-items: center; gap: 6px; padding: 3px 10px; border-radius: 999px; font-size: 12px; font-weight: 600; }
.ix-badge::before { content: ""; width: 6px; height: 6px; border-radius: 50%; background: currentColor; }
.ix-badge--ok     { color: #047857; background: #ecfdf5; }
.ix-badge--warn   { color: #b45309; background: #fffbeb; }
.ix-badge--danger { color: #b91c1c; background: #fef2f2; }
.ix-ring__hero { display: flex; align-items: flex-end; justify-content: space-between; gap: 12px; }
.ix-ring__rate { font-size: 46px; font-weight: 800; line-height: .95; letter-spacing: -.035em; color: var(--ix-ink); font-variant-numeric: tabular-nums; }
.ix-ring__rate small { font-size: 22px; font-weight: 700; color: var(--ix-muted); margin-left: 2px; letter-spacing: 0; }
.ix-ring__rate-label { margin-top: 6px; font-size: 12.5px; font-weight: 500; color: var(--ix-muted); }
.ix-ring__devices {
  display: inline-flex; align-items: center; gap: 7px; padding: 6px 10px; border-radius: 10px;
  font-size: 13px; color: var(--ix-muted); background: #f8fafc; border: 1px solid var(--ix-border-soft);
}
.ix-ring__devices strong { color: var(--ix-ink); font-weight: 700; font-variant-numeric: tabular-nums; }
.ix-bar { display: flex; gap: 2px; height: 8px; border-radius: 999px; overflow: hidden; background: #f1f5f9; }
.ix-bar__seg { flex-basis: 0; min-width: 0; }
.ix-bar__seg--ok     { background: var(--ix-success); }
.ix-bar__seg--warn   { background: var(--ix-warning); }
.ix-bar__seg--danger { background: var(--ix-danger); }
.ix-bar__seg--other  { background: #cbd5e1; }
.ix-ring__metrics { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; }
.ix-metric { min-width: 0; display: flex; flex-direction: column; gap: 6px; padding: 10px; border-radius: 12px; background: var(--m-soft); }
.ix-metric--ok     { --m: #047857; --m-soft: #ecfdf5; }
.ix-metric--warn   { --m: #b45309; --m-soft: #fffbeb; }
.ix-metric--danger { --m: #b91c1c; --m-soft: #fef2f2; }
.ix-metric__head { display: flex; align-items: center; gap: 5px; font-size: 11.5px; white-space: nowrap; overflow: hidden; font-weight: 600; color: var(--m); }
.ix-metric__value { font-size: 24px; font-weight: 700; line-height: 1.05; color: var(--ix-ink); font-variant-numeric: tabular-nums; }
.ix-metric--zero { background: #f8fafc; }
.ix-metric--zero .ix-metric__head { color: #94a3b8; }
.ix-metric--zero .ix-metric__value { color: #cbd5e1; }
.ix-ring__foot { display: flex; flex-wrap: wrap; gap: 6px 16px; font-size: 12px; color: var(--ix-muted); }
.ix-ring__foot span::before { content: ""; display: inline-block; width: 8px; height: 8px; border-radius: 2px; margin-right: 6px; background: #cbd5e1; }

/* ---------- Bandeau "Non Encrypted" ---------- */
.ix-alert {
  position: relative; overflow: hidden;
  display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 20px;
  padding: 20px 24px 20px 28px; margin: 6px 0 14px; border-radius: 16px;
  background: linear-gradient(120deg, #fff1f2 0%, #ffffff 65%);
  border: 1px solid #fecdd3; box-shadow: var(--ix-shadow-sm);
}
.ix-alert::before { content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px; background: linear-gradient(180deg, #fb7185, #dc2626); }
.ix-alert__main { display: flex; align-items: center; gap: 18px; }
.ix-alert__icon {
  display: inline-flex; align-items: center; justify-content: center; flex-shrink: 0;
  width: 56px; height: 56px; border-radius: 16px; color: #dc2626; background: #ffe4e6;
  box-shadow: inset 0 0 0 1px #fecdd3;
}
.ix-alert__title { font-size: 14px; font-weight: 600; color: #9f1239; }
.ix-alert__count { display: flex; align-items: baseline; flex-wrap: wrap; gap: 10px; margin-top: 2px; }
#nonEncCounter { font-size: 42px; font-weight: 800; line-height: 1; letter-spacing: -.035em; color: #dc2626; font-variant-numeric: tabular-nums; }
.ix-alert__count-label { font-size: 14px; font-weight: 500; color: #881337; }
.ix-pill {
  display: inline-block; margin-top: 10px; padding: 4px 11px; border-radius: 999px;
  font-size: 12.5px; font-weight: 500; color: #475569; background: #f1f5f9; border: 1px solid var(--ix-border);
}
.ix-pill .ix-icon { margin-right: 5px; vertical-align: -2px; }
.ix-alert__side { display: flex; flex-direction: column; align-items: flex-end; gap: 14px; }
.ix-kpis { display: flex; gap: 8px; }
.ix-kpi {
  display: flex; flex-direction: column; align-items: center; min-width: 88px;
  padding: 8px 12px; border-radius: 12px; background: #fff; border: 1px solid var(--ix-border);
}
.ix-kpi__v { font-size: 20px; font-weight: 700; line-height: 1.15; color: var(--ix-ink); font-variant-numeric: tabular-nums; }
.ix-kpi__l { font-size: 12px; font-weight: 500; color: var(--ix-muted); white-space: nowrap; }
.ix-kpi--danger .ix-kpi__v { color: #dc2626; }
.ix-kpi--muted  .ix-kpi__v { color: #94a3b8; }
.ix-switch { position: relative; display: inline-flex; align-items: center; gap: 10px; cursor: pointer; user-select: none; font-size: 13px; font-weight: 500; color: var(--ix-ink-2); }
.ix-switch input { position: absolute; opacity: 0; width: 1px; height: 1px; margin: 0; }
.ix-switch__track { position: relative; flex-shrink: 0; width: 40px; height: 22px; border-radius: 999px; background: #cbd5e1; transition: background .2s ease; }
.ix-switch__thumb { position: absolute; top: 3px; left: 3px; width: 16px; height: 16px; border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(15,23,42,.3); transition: transform .2s ease; }
.ix-switch input:checked + .ix-switch__track { background: var(--ix-indigo); }
.ix-switch input:checked + .ix-switch__track .ix-switch__thumb { transform: translateX(18px); }
.ix-switch input:focus-visible + .ix-switch__track { outline: 2px solid var(--ix-indigo); outline-offset: 2px; }
.ix-alert--ok { background: linear-gradient(120deg, #ecfdf5 0%, #ffffff 65%); border-color: #a7f3d0; }
.ix-alert--ok::before { background: linear-gradient(180deg, #34d399, #059669); }
.ix-alert--ok .ix-alert__icon { color: #059669; background: #d1fae5; box-shadow: inset 0 0 0 1px #a7f3d0; }
.ix-alert--ok .ix-alert__title, .ix-alert--ok .ix-alert__count-label { color: #065f46; }
.ix-alert--ok #nonEncCounter { color: #059669; }

/* ---------- États vides ---------- */
.ix-empty { display: flex; flex-direction: column; align-items: center; gap: 10px; padding: 36px 16px; text-align: center; }
.ix-empty__icon { display: inline-flex; align-items: center; justify-content: center; width: 52px; height: 52px; border-radius: 16px; }
.ix-empty--success .ix-empty__icon { color: #059669; background: #ecfdf5; }
.ix-empty--neutral .ix-empty__icon { color: #475569; background: #f1f5f9; }
.ix-empty--danger .ix-empty__icon { color: #dc2626; background: #fef2f2; }
.ix-empty__title { font-size: 16px; font-weight: 600; color: var(--ix-ink); }
.ix-empty__text { max-width: 60ch; font-size: 13px; color: var(--ix-muted); }

/* ---------- [v3.3] Analyse des non-conformités : note de lecture, source, avertissements ---------- */
.ix-note {
  display: flex; align-items: flex-start; gap: 12px;
  margin: 14px 4px 4px; padding: 14px 18px; border-radius: 14px;
  font-size: 13px; color: #78350f; background: #fffbeb; border: 1px solid #fde68a;
}
.ix-note .ix-icon { flex-shrink: 0; margin-top: 1px; color: #d97706; }
.ix-note strong { font-weight: 600; }
.ix-note ul { margin: 6px 0 0; padding-left: 18px; }
.ix-note li + li { margin-top: 4px; }
.ix-note--warn { color: #7f1d1d; background: #fef2f2; border-color: #fecaca; }
.ix-note--warn .ix-icon { color: #dc2626; }
.ix-meta { display: flex; flex-wrap: wrap; gap: 8px; margin: 12px 4px 2px; }
.ix-meta__item {
  display: inline-flex; align-items: center; gap: 7px; max-width: 100%; padding: 4px 11px; border-radius: 999px;
  font-size: 12.5px; font-weight: 500; color: #475569; background: #f1f5f9; border: 1px solid var(--ix-border);
  overflow-wrap: anywhere;
}
.ix-meta__item .ix-icon { color: var(--ix-indigo); }

/* ---------- Tables DataTables ---------- */
.dataTables_wrapper { font-size: 13px; color: var(--ix-ink-2); }
table.dataTable { border-collapse: separate !important; border-spacing: 0; }
table.dataTable thead th, table.dataTable thead td {
  background-color: #f8fafc !important; color: #475569 !important;
  font-size: 12px; font-weight: 600 !important;
  border-bottom: 1px solid var(--ix-border) !important; padding: 10px 12px !important;
}
table.dataTable tbody td { padding: 9px 12px !important; border-top: 1px solid var(--ix-border-soft) !important; color: var(--ix-ink-2); }
table.dataTable tbody tr:hover > td { background-color: #f5f7ff !important; }
.dataTables_wrapper .dataTables_filter input, .dataTables_wrapper .dataTables_length select {
  border: 1px solid var(--ix-border) !important; border-radius: 10px !important; padding: 6px 12px !important;
  background-color: #fff; outline: none; transition: border-color .15s ease, box-shadow .15s ease;
}
.dataTables_wrapper .dataTables_filter input:focus { border-color: var(--ix-indigo) !important; box-shadow: 0 0 0 3px rgba(79,70,229,.15); }
.dataTables_wrapper .dataTables_paginate .paginate_button { border-radius: 8px !important; border: 1px solid transparent !important; padding: 4px 10px !important; margin: 0 2px; }
.dataTables_wrapper .dataTables_paginate .paginate_button:hover { background: #eef2ff !important; color: var(--ix-night) !important; border-color: #e0e7ff !important; }
.dataTables_wrapper .dataTables_paginate .paginate_button.current,
.dataTables_wrapper .dataTables_paginate .paginate_button.current:hover { background: var(--ix-night) !important; color: #fff !important; border-color: var(--ix-night) !important; }
button.dt-button, .dt-buttons .dt-button {
  border-radius: 8px !important; border: 1px solid var(--ix-border) !important; background: #fff !important;
  color: var(--ix-ink-2) !important; font-size: 12px !important; font-weight: 500 !important;
  padding: 6px 12px !important; box-shadow: var(--ix-shadow-sm);
}
button.dt-button:hover, .dt-buttons .dt-button:hover { background: #f8fafc !important; }

/* ---------- [v3.2] Navigation entre pages (onglets collants) ---------- */
.ix-nav-sentinel { height: 1px; margin-bottom: -1px; }
.ix-nav { display: none; }
.ix-spa .ix-nav {
  position: sticky; top: 8px; z-index: 60;
  display: flex; align-items: center; justify-content: space-between; gap: 12px;
  margin: 10px 5px 8px; padding: 6px;
  background: rgba(255,255,255,.84);
  -webkit-backdrop-filter: blur(14px) saturate(160%); backdrop-filter: blur(14px) saturate(160%);
  border: 1px solid var(--ix-border); border-radius: 16px;
  box-shadow: var(--ix-shadow-sm);
  transition: box-shadow .25s ease, background-color .25s ease;
}
.ix-spa .ix-nav::before {
  content: ""; position: absolute; left: -5px; right: -5px; top: -10px; height: 10px;
  background: var(--ix-bg); pointer-events: none;
}
.ix-spa .ix-nav.is-stuck { background: rgba(255,255,255,.92); box-shadow: 0 14px 34px -16px rgba(30,27,75,.38); }
.ix-tabs { position: relative; display: flex; gap: 4px; min-width: 0; overflow-x: auto; scrollbar-width: none; }
.ix-tabs::-webkit-scrollbar { display: none; }
.ix-tabs__indicator {
  position: absolute; left: 0; top: 0; bottom: 0; width: 0; z-index: 0;
  border-radius: 11px; opacity: 0; pointer-events: none;
  background: linear-gradient(135deg, var(--ix-night) 0%, var(--ix-indigo) 100%);
  box-shadow: 0 8px 18px -8px rgba(79,70,229,.7);
  transition: transform .42s cubic-bezier(.22,.8,.26,1), width .42s cubic-bezier(.22,.8,.26,1), opacity .2s ease;
}
.ix-tabs.has-indicator .ix-tabs__indicator { opacity: 1; }
.ix-tab {
  position: relative; z-index: 1; flex-shrink: 0;
  display: inline-flex; align-items: center; gap: 9px;
  padding: 10px 16px; margin: 0; border: 0; border-radius: 11px; background: transparent;
  font-size: 14px; font-weight: 600; line-height: 1.2; color: var(--ix-muted);
  cursor: pointer; white-space: nowrap;
  transition: color .25s ease, background-color .2s ease;
}
.ix-tab:hover { color: var(--ix-ink); background-color: rgba(15,23,42,.05); }
.ix-tab[aria-selected="true"], .ix-tab[aria-selected="true"]:hover { color: #fff; background-color: transparent; }
.ix-tabs:not(.has-indicator) .ix-tab[aria-selected="true"] { background: linear-gradient(135deg, var(--ix-night), var(--ix-indigo)); }
.ix-tab:focus-visible { outline: 2px solid var(--ix-indigo); outline-offset: 2px; }
.ix-nav__aside { display: flex; align-items: center; gap: 12px; flex-shrink: 0; padding-right: 2px; }
.ix-nav__context {
  max-width: 260px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  font-size: 13px; font-weight: 500; color: var(--ix-muted);
  opacity: 0; transform: translateY(4px); transition: opacity .25s ease, transform .25s ease;
}
.ix-nav.is-stuck .ix-nav__context { opacity: 1; transform: none; }
.ix-nav__btn {
  display: inline-flex; align-items: center; gap: 7px; padding: 8px 12px; margin: 0;
  border-radius: 10px; border: 1px solid var(--ix-border); background: #fff;
  font-size: 13px; font-weight: 500; color: var(--ix-ink-2); cursor: pointer;
  transition: background-color .15s ease, border-color .15s ease;
}
.ix-nav__btn:hover { background: #f8fafc; border-color: #cbd5e1; }
.ix-nav__btn:focus-visible { outline: 2px solid var(--ix-indigo); outline-offset: 2px; }

/* ---------- [v3.2] Pages virtuelles et transition ---------- */
@keyframes ixPageIn {
  from { opacity: 0; transform: translate3d(0, 14px, 0); }
  to   { opacity: 1; transform: none; }
}
/* Les pages inactives restent mises en page (largeur réelle) mais sont repliées à une
   hauteur nulle et rendues invisibles : pas de display:none, sinon ApexCharts redessine
   les graphiques à 0 px et DataTables calcule des colonnes fausses. La page active est
   définie par les règles générées dans New-IxNavigation (attribut data-ix-page). */
.ix-pages { position: relative; }
.ix-spa .ix-page {
  position: absolute; top: 0; left: 0; right: 0; height: 0;
  overflow: hidden; visibility: hidden; pointer-events: none;
}
.ix-page__print-head { display: none; }

/* ---------- [v3.2] Boutons Afficher / Masquer des sections repliables ---------- */
.defaultSectionText a[id^="show_"], .defaultSectionText a[id^="hide_"] {
  display: inline-flex; align-items: center; gap: 7px; margin-left: 8px; padding: 3px 11px 3px 10px;
  border-radius: 999px; font-size: 0; line-height: 1.5; text-decoration: none; color: rgba(255,255,255,.95);
  background: rgba(255,255,255,.14); border: 1px solid rgba(255,255,255,.26);
  transition: background-color .15s ease;
}
.defaultSectionText a[id^="show_"]::after { content: "Afficher"; font-size: 12px; font-weight: 500; }
.defaultSectionText a[id^="hide_"]::after { content: "Masquer"; font-size: 12px; font-weight: 500; }
.defaultSectionText a[id^="show_"]::before, .defaultSectionText a[id^="hide_"]::before {
  content: ""; width: 6px; height: 6px; border-right: 1.6px solid currentColor; border-bottom: 1.6px solid currentColor;
}
.defaultSectionText a[id^="show_"]::before { transform: translateY(-2px) rotate(45deg); }
.defaultSectionText a[id^="hide_"]::before { transform: translateY(2px) rotate(-135deg); }
.defaultSectionText a[id^="show_"]:hover, .defaultSectionText a[id^="hide_"]:hover { background: rgba(255,255,255,.28); }
.defaultSectionHead:has(a[id^="show_"]) { cursor: pointer; user-select: none; }

/* ---------- Responsive, accessibilité, impression ---------- */
@media (max-width: 760px) {
  .ix-hero { padding: 28px 22px; }
  .ix-hero__title { font-size: 32px; }
  .ix-hero__contact { min-width: 0; width: 100%; }
  .ix-alert__side { align-items: flex-start; }
  .ix-overview { grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
  .ix-stat { padding: 14px; }
  .ix-stat__value { font-size: 30px; }
  .ix-ring { max-width: none; }
  .ix-spa .ix-nav { top: 0; margin: 8px 0; border-radius: 14px; }
  .ix-spa .ix-nav::before { left: 0; right: 0; }
  .ix-tabs {
    -webkit-mask-image: linear-gradient(90deg, transparent 0, #000 14px, #000 calc(100% - 14px), transparent 100%);
            mask-image: linear-gradient(90deg, transparent 0, #000 14px, #000 calc(100% - 14px), transparent 100%);
  }
  .ix-tab { padding: 9px 12px; }
  .ix-nav__context, .ix-nav__btn span { display: none; }
}
@media (prefers-reduced-motion: reduce) {
  .ix-ring, .ix-switch__track, .ix-switch__thumb, .ix-tabs__indicator, .ix-tab, .ix-nav__context { transition: none; }
  .ix-ring:hover { transform: none; }
  .ix-spa .ix-page { animation: none !important; }
}
@media print {
  html body { background: #fff !important; }
  .ix-nav, .ix-nav-sentinel { display: none !important; }
  .ix-spa .ix-page { position: static !important; height: auto !important; overflow: visible !important; visibility: visible !important; animation: none !important; }
  .ix-page + .ix-page { break-before: page; }
  .ix-page .defaultSection { break-inside: avoid; page-break-inside: avoid; }
  .ix-page__print-head { display: flex !important; align-items: baseline; flex-wrap: wrap; gap: 4px 12px; margin: 14px 8px 10px; color: var(--ix-night); }
  .ix-page__print-head .ix-icon { align-self: center; color: var(--ix-indigo); }
  .ix-page__print-head span { font-size: 22px; font-weight: 700; }
  .ix-page__print-head small { font-size: 13px; color: var(--ix-muted); }
  .defaultSectionText a[id^="show_"], .defaultSectionText a[id^="hide_"] { display: none !important; }
  .defaultSection, .ix-stat, .ix-ring, .ix-alert { box-shadow: none !important; }
  .ix-hero, .ix-stat, .ix-ring, .ix-alert, .ix-bar { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
}
</style>
"@
}

# ========================================
# [UI v3.2] NAVIGATION ENTRE PAGES VIRTUELLES (SPA dans un seul fichier HTML)
# ========================================
# Ces fragments sont émis en HTML brut directement dans le bloc New-HTML
# (et non via New-HTMLText) : le menu reste ainsi "collant" sur toute la
# hauteur du rapport et les scripts ne passent pas par le traitement de texte
# de PSWriteHTML.

# Menu d'onglets + état initial (lu depuis l'URL, ex. rapport.html#securite).
# Le petit script d'amorçage s'exécute AVANT l'analyse des pages : la bonne
# page s'affiche directement, sans flash du rapport complet.
function New-IxNavigation {
    param([object[]]$Pages, [string]$ContextLabel = "")

    $tabsHtml = ($Pages | ForEach-Object {
        '<button type="button" class="ix-tab" role="tab" id="ix-tab-{0}" data-page="{0}" aria-controls="ix-page-{0}" aria-selected="false" tabindex="-1" title="{1}">{2}<span class="ix-tab__label">{3}</span></button>' -f $_.Id, (ConvertTo-HtmlSafe $_.Hint), (Get-IconSvg -Name $_.Icon -Size 17 -Stroke '1.9'), (ConvertTo-HtmlSafe $_.Label)
    }) -join "`n    "

    # Page active = celle dont l'id figure dans <html data-ix-page="...">. L'animation
    # d'entrée est attachée à cet état : elle rejoue à chaque changement d'onglet.
    $pageSelectors = ($Pages | ForEach-Object { '.ix-spa[data-ix-page="{0}"] #ix-page-{0}' -f $_.Id }) -join ",`n"
    $pageCss = $pageSelectors + " {`n  position: relative; height: auto; overflow: visible; visibility: visible; pointer-events: auto;`n  animation: ixPageIn .45s cubic-bezier(.22,.8,.26,1);`n}"
    $idsJs   = ($Pages | ForEach-Object { "'" + $_.Id + "'" }) -join ', '
    $context = ConvertTo-HtmlSafe $ContextLabel

    $html = @"
<div class="ix-nav-sentinel" id="ixNavSentinel" aria-hidden="true"></div>
<nav class="ix-root ix-nav" id="ixNav" aria-label="Navigation du rapport">
  <div class="ix-tabs" role="tablist" aria-label="Pages du rapport">
    <span class="ix-tabs__indicator" aria-hidden="true"></span>
    $tabsHtml
  </div>
  <div class="ix-nav__aside">
    <span class="ix-nav__context" title="$context">$context</span>
    <button type="button" class="ix-nav__btn" id="ixPrintBtn" title="Imprimer toutes les pages ou les exporter en PDF">$(Get-IconSvg -Name 'printer' -Size 16 -Stroke '1.9')<span>Imprimer</span></button>
  </div>
</nav>
<style>
$pageCss
</style>
"@

    # Chaîne littérale (aucune interpolation PowerShell dans le JavaScript)
    $boot = @'
<script>
/* [UI v3.2] Amorcage : choisit la page initiale avant l'affichage du contenu */
(function () {
    var ids  = [__IX_IDS__];
    var hash = (window.location.hash || '').replace('#', '');
    var id   = ids.indexOf(hash) >= 0 ? hash : ids[0];
    var root = document.documentElement;
    root.classList.add('ix-spa');
    root.setAttribute('data-ix-page', id);
    var tabs = document.querySelectorAll('#ixNav .ix-tab');
    for (var i = 0; i < tabs.length; i++) {
        var on = tabs[i].getAttribute('data-page') === id;
        tabs[i].setAttribute('aria-selected', on ? 'true' : 'false');
        tabs[i].tabIndex = on ? 0 : -1;
    }
})();
</script>
'@
    return $html + $boot.Replace('__IX_IDS__', $idsJs)
}

# Ouverture d'une page virtuelle (la balise fermante '</div>' est émise dans New-HTML).
# L'en-tête de page n'apparaît qu'à l'impression, où toutes les pages se suivent.
function Get-IxPageStart {
    param([object]$Page)
    return @"
<div class="ix-page" id="ix-page-$($Page.Id)" role="tabpanel" aria-labelledby="ix-tab-$($Page.Id)">
<div class="ix-root ix-page__print-head">$(Get-IconSvg -Name $Page.Icon -Size 20)<span>$(ConvertTo-HtmlSafe $Page.Label)</span><small>$(ConvertTo-HtmlSafe $Page.Hint)</small></div>
"@
}

# Comportements de navigation : clic et clavier sur les onglets, historique
# (bouton Précédent du navigateur), pastille glissante, redessin des graphiques
# et tables à l'affichage d'une page, impression de toutes les pages.
function Get-IxSpaScript {
    return @'
<script>
/* [UI v3.2] Navigation entre les pages virtuelles du dashboard */
(function () {
    var root = document.documentElement;
    var nav  = document.getElementById('ixNav');
    if (!nav) { return; }
    var tablist   = nav.querySelector('.ix-tabs');
    var indicator = nav.querySelector('.ix-tabs__indicator');
    var sentinel  = document.getElementById('ixNavSentinel');
    var printBtn  = document.getElementById('ixPrintBtn');
    var tabs      = Array.prototype.slice.call(nav.querySelectorAll('.ix-tab'));
    var ids       = tabs.map(function (t) { return t.getAttribute('data-page'); });
    var reduce    = !!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches);

    function current()  { return root.getAttribute('data-ix-page') || ids[0]; }
    function readHash() { return (window.location.hash || '').replace('#', ''); }
    function tabFor(id) {
        for (var i = 0; i < tabs.length; i++) { if (tabs[i].getAttribute('data-page') === id) { return tabs[i]; } }
        return null;
    }

    /* Pastille glissante positionnee sous l'onglet actif.
       On ne la repositionne que si la geometrie a change : un evenement "resize"
       declenche par le redessin des graphiques n'interrompt pas l'animation. */
    var lastX = -1, lastW = -1;
    function moveIndicator(instant) {
        var tab = tabFor(current());
        if (!tab || !indicator || !tab.offsetWidth) { return; }
        if (tab.offsetLeft === lastX && tab.offsetWidth === lastW) { return; }
        lastX = tab.offsetLeft; lastW = tab.offsetWidth;
        if (instant) { indicator.style.transition = 'none'; }
        indicator.style.width     = tab.offsetWidth + 'px';
        indicator.style.transform = 'translateX(' + tab.offsetLeft + 'px)';
        if (instant) { void indicator.offsetWidth; indicator.style.transition = ''; }
        tablist.classList.add('has-indicator');
    }

    /* Recalcule la taille des graphiques et des tables d'une page devenue visible.
       Utilise la fonction native de PSWriteHTML si elle existe, sinon un repli generique. */
    function refreshPage(id) {
        var pageId = 'ix-page-' + id;
        if (!document.getElementById(pageId)) { return; }
        if (typeof window.findObjectsToRedraw === 'function') {
            try { window.findObjectsToRedraw(pageId); return; } catch (e) { }
        }
        try { window.dispatchEvent(new Event('resize')); } catch (e) { }
        if (window.jQuery && jQuery.fn.dataTable) {
            try {
                var api = jQuery.fn.dataTable.tables({ visible: true, api: true });
                api.columns.adjust();
                if (api.responsive && api.responsive.recalc) { api.responsive.recalc(); }
            } catch (e) { }
        }
    }

    function show(id) {
        if (ids.indexOf(id) < 0) { id = ids[0]; }
        var changed = id !== current();
        root.setAttribute('data-ix-page', id);
        tabs.forEach(function (t) {
            var on = t.getAttribute('data-page') === id;
            t.setAttribute('aria-selected', on ? 'true' : 'false');
            t.tabIndex = on ? 0 : -1;
        });
        moveIndicator(false);

        /* Mobile : garde l'onglet actif visible dans la barre defilante */
        var tab = tabFor(id);
        if (tab && tablist.scrollWidth > tablist.clientWidth) {
            var left = tab.offsetLeft - (tablist.clientWidth - tab.offsetWidth) / 2;
            try { tablist.scrollTo({ left: left, behavior: reduce ? 'auto' : 'smooth' }); } catch (e) { tablist.scrollLeft = left; }
        }
        if (!changed) { return; }

        /* Si l'on a defile sous le menu, la nouvelle page s'ouvre en haut, juste sous les onglets */
        if (sentinel) {
            var top = sentinel.getBoundingClientRect().top + window.pageYOffset;
            if (window.pageYOffset > top) { window.scrollTo(0, top); }
        }
        requestAnimationFrame(function () { requestAnimationFrame(function () { refreshPage(id); }); });
    }

    function navigate(id) {
        if (id === current()) { return; }
        if (readHash() !== id) {
            try { window.history.pushState(null, '', '#' + id); }
            catch (e) { try { window.location.hash = id; } catch (e2) { } }
        }
        show(id);
    }

    /* Clic et clavier (fleches gauche/droite, Debut, Fin) sur les onglets */
    tabs.forEach(function (t) {
        t.addEventListener('click', function () { navigate(t.getAttribute('data-page')); });
        t.addEventListener('keydown', function (e) {
            var i = tabs.indexOf(t), n = -1;
            if (e.key === 'ArrowRight')     { n = (i + 1) % tabs.length; }
            else if (e.key === 'ArrowLeft') { n = (i - 1 + tabs.length) % tabs.length; }
            else if (e.key === 'Home')      { n = 0; }
            else if (e.key === 'End')       { n = tabs.length - 1; }
            if (n < 0) { return; }
            e.preventDefault();
            tabs[n].focus();
            navigate(tabs[n].getAttribute('data-page'));
        });
    });

    /* Boutons Precedent / Suivant du navigateur */
    window.addEventListener('popstate',   function () { show(readHash()); });
    window.addEventListener('hashchange', function () { show(readHash()); });

    /* Ombre du menu lorsqu'il est colle en haut de l'ecran */
    if (sentinel && 'IntersectionObserver' in window) {
        new IntersectionObserver(function (entries) {
            var e = entries[0];
            nav.classList.toggle('is-stuck', !e.isIntersecting && e.boundingClientRect.top < 10);
        }, { rootMargin: '-10px 0px 0px 0px', threshold: 0 }).observe(sentinel);
    }

    /* Toute la barre de titre d'une section repliable devient cliquable */
    document.addEventListener('click', function (e) {
        var t = e.target;
        if (!t || !t.closest) { return; }
        var head = t.closest('.defaultSectionHead');
        if (!head || t.closest('a')) { return; }
        var links = head.querySelectorAll('a[id^="show_"], a[id^="hide_"]');
        for (var i = 0; i < links.length; i++) {
            if (links[i].style.display !== 'none') { links[i].click(); break; }
        }
    });

    /* Impression (bouton ou Ctrl+P) : la feuille de style d'impression affiche toutes
       les pages a la suite, chacune precedee de son titre, comme avant la v3.2 */
    if (printBtn) {
        printBtn.addEventListener('click', function () { window.print(); });
    }

    /* Repositionne la pastille (redimensionnement, chargement de la police Inter) */
    var resizeTimer;
    window.addEventListener('resize', function () {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(function () { moveIndicator(true); }, 80);
    });
    if (document.fonts && document.fonts.ready) { document.fonts.ready.then(function () { moveIndicator(true); }); }
    window.addEventListener('load', function () { moveIndicator(true); });
    moveIndicator(true);
})();
</script>
'@
}

# [FUSION v4 - Proactivité] Pseudonymes STABLES le temps d'une génération : un même
# poste / utilisateur porte le même alias dans toutes les pages et tous les CSV
# (la v3 tirait un alias aléatoire par ligne : "Inactifs 30 j" et "Inactifs 60 j"
# montraient deux alias différents pour un même poste). Remis à zéro à chaque génération.
function Reset-AnonymizationMaps {
    $script:AnonNameMap = @{}   # nom de poste réel (minuscules) -> pseudonyme
    $script:AnonUpnMap  = @{}   # UPN réel (minuscules)          -> pseudonyme
}

function Get-AnonymizedIdentity {
    param([string]$RealName, [string]$RealUpn)
    if ($null -eq $script:AnonNameMap -or $null -eq $script:AnonUpnMap) { Reset-AnonymizationMaps }
    $alias = ""
    if (-not [string]::IsNullOrWhiteSpace($RealName)) {
        $k = $RealName.Trim().ToLowerInvariant()
        if (-not $script:AnonNameMap.ContainsKey($k)) { $script:AnonNameMap[$k] = "Poste-" + ([guid]::NewGuid().ToString().Substring(0, 8)) }
        $alias = $script:AnonNameMap[$k]
    }
    $upnAlias = ""
    if (-not [string]::IsNullOrWhiteSpace($RealUpn)) {
        $k = $RealUpn.Trim().ToLowerInvariant()
        if (-not $script:AnonUpnMap.ContainsKey($k)) { $script:AnonUpnMap[$k] = "User-" + ([guid]::NewGuid().ToString().Substring(0, 8)) }
        $upnAlias = $script:AnonUpnMap[$k]
    }
    return [PSCustomObject]@{ Name = $alias; Upn = $upnAlias }
}

function Anonymize-DeviceData {
    param([Parameter(Mandatory=$true)][AllowEmptyCollection()]$DeviceList)
    return $DeviceList | ForEach-Object {
        $anon = Get-AnonymizedIdentity -RealName $_.DeviceName -RealUpn $_.UserPrincipalName   # [v4.0] alias stables
        [PSCustomObject]@{
            DeviceName              = $anon.Name
            UserPrincipalName       = $anon.Upn
            OperatingSystem         = $_.OperatingSystem
            Manufacturer            = $_.Manufacturer
            Model                   = $_.Model
            OSVersion               = $_.OSVersion
            ComplianceState         = $_.ComplianceState
            IsEncrypted             = $_.IsEncrypted
            LastSyncDateTime        = $_.LastSyncDateTime
            EnrolledDateTime        = $_.EnrolledDateTime
            FreeStorageSpaceInBytes = $_.FreeStorageSpaceInBytes
            TotalStorageSpaceInBytes = $_.TotalStorageSpaceInBytes
            # [MODIF v3] Conservation des champs d'analyse "Non Encrypted".
            # Ils valent $null pour les listes qui ne les possèdent pas :
            # sans impact sur les autres tables (colonnes non sélectionnées).
            RootCause               = $_.RootCause
            DaysInactive            = $_.DaysInactive
        }
    }
}

# ========================================
# [v3.3] APPELS À L'API GRAPH : GESTION DES ERREURS
# ========================================
# [v4.0] TOUS les appels Graph du script passent par Invoke-GraphApiRequest (le
# module Microsoft.Graph n'est plus utilisé) :
#  - 429 (throttling), 500/502/503/504 et erreurs réseau : nouvelles tentatives
#    bornées ($GraphMaxRetries). Retry-After est un MINIMUM imposé par le service :
#    il n'est jamais raccourci (le script Proactivité repartait en moyenne au bout
#    de 23 s pour un Retry-After de 30 s). Sans Retry-After : backoff exponentiel
#    bruité, pour que des flux limités au même instant ne repartent pas ensemble
#  - 429 : pause COMMUNE à tous les flux parallèles (GraphContext.PauseUntilUtc)
#  - 401 : jeton renouvelé une fois puis requête rejouée
#  - autres erreurs (400, 403, 404...) : exception au message explicite (code
#    HTTP, code Graph, request-id, piste de résolution). Le code HTTP est exposé
#    dans Exception.Data['StatusCode'] pour les appelants.
#  - attentes via Start-ResponsiveSleep : la fenêtre ne passe plus en "Ne répond pas"
# Compatible Windows PowerShell 5.1 et PowerShell 7.

# [v4.0] Contexte partagé de la génération en cours (créé par New-GraphContext dans
# Generate-Dashboard). Table SYNCHRONISÉE, transmise telle quelle aux runspaces de
# la collecte parallèle : jeton (un renouvellement profite à tous les flux) et pause
# de throttling (un 429 reçu par un flux met tous les flux en pause).
$script:GraphContext = $null

# Fonctions exécutées dans les runspaces de la collecte parallèle (injectées via
# InitialSessionState : une seule implémentation des reprises pour tout le script)
$GraphWorkerFunctions = @(
    'Invoke-GraphApiRequest', 'Get-GraphErrorInfo', 'Format-GraphErrorMessage', 'Get-GraphAccessToken',
    'Wait-GraphThrottleGate', 'Get-GraphPagedResults', 'Start-ResponsiveSleep', 'Set-UiStatus', 'Write-Log'
)

function New-GraphContext {
    param([string]$TenantId, [string]$ClientId, [string]$ClientSecret)
    return [hashtable]::Synchronized(@{
        TenantId      = $TenantId
        ClientId      = $ClientId
        ClientSecret  = $ClientSecret
        AccessToken   = $null
        ExpiresAtUtc  = [datetime]::MinValue
        PauseUntilUtc = [datetime]::MinValue
    })
}

# [v4.0] Attend la fin d'une pause de throttling posée par n'importe quel flux
function Wait-GraphThrottleGate {
    $ctx = $script:GraphContext
    if (-not $ctx) { return }
    $now = [datetime]::UtcNow
    if ($ctx.PauseUntilUtc -le $now) { return }
    $wait = [int][math]::Ceiling(($ctx.PauseUntilUtc - $now).TotalSeconds)
    Start-ResponsiveSleep -Seconds $wait -Message "Limitation Microsoft Graph : pause commune à tous les flux"
}

# Extrait d'une erreur Invoke-RestMethod / Invoke-WebRequest : code HTTP,
# délai Retry-After, code et message Graph (ou OAuth), request-id, erreur réseau.
function Get-GraphErrorInfo {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)

    $ex   = $ErrorRecord.Exception
    $info = [pscustomobject]@{ StatusCode = 0; RetryAfter = 0; Code = ""; Message = "$($ex.Message)"; RequestId = ""; IsNetworkError = $false }

    $response = $null
    if ($ex.PSObject.Properties['Response']) { $response = $ex.Response }

    if ($response) {
        try { $info.StatusCode = [int]$response.StatusCode } catch { }
        try {
            $retryAfter = $null
            if ($response.Headers -is [System.Net.WebHeaderCollection]) {
                $retryAfter = $response.Headers['Retry-After']                        # Windows PowerShell 5.1
            } elseif ($response.Headers -and $response.Headers.RetryAfter) {
                $rc = $response.Headers.RetryAfter                                     # PowerShell 7
                if ($rc.Delta)    { $retryAfter = $rc.Delta.TotalSeconds }
                elseif ($rc.Date) { $retryAfter = ($rc.Date - [DateTimeOffset]::UtcNow).TotalSeconds }
            }
            $seconds = 0.0
            if ($null -ne $retryAfter -and
                [double]::TryParse("$retryAfter", [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$seconds) -and
                $seconds -gt 0) {
                $info.RetryAfter = [int][math]::Ceiling($seconds)
            }
        } catch { }
    } else {
        # Pas de réponse HTTP : délai dépassé, DNS, proxy, connexion interrompue...
        $inner = $ex
        while ($inner) {
            if ($inner.GetType().FullName -in @('System.Net.WebException', 'System.Net.Http.HttpRequestException',
                    'System.Threading.Tasks.TaskCanceledException', 'System.TimeoutException',
                    'System.IO.IOException', 'System.Net.Sockets.SocketException')) {
                $info.IsNetworkError = $true; break
            }
            $inner = $inner.InnerException
        }
    }

    # Corps de l'erreur : Graph { error: { code, message, innerError } } ou OAuth { error, error_description }
    $body = $null
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) { $body = $ErrorRecord.ErrorDetails.Message }
    if ($body) {
        try {
            $json = $body | ConvertFrom-Json -ErrorAction Stop
            if ($json.error -is [string]) {
                $info.Code = $json.error
                if ($json.error_description) { $info.Message = ("$($json.error_description)" -split "`r?`n")[0] }
            } elseif ($json.error) {
                if ($json.error.code)    { $info.Code    = "$($json.error.code)" }
                if ($json.error.message) { $info.Message = "$($json.error.message)" }
                $innerError = $json.error.innerError
                if ($innerError) {
                    $requestId = $innerError.'request-id'
                    if (-not $requestId) { $requestId = $innerError.requestId }
                    if ($requestId) { $info.RequestId = "$requestId" }
                }
            }
        } catch {
            if ($body.Length -le 300) { $info.Message = $body.Trim() }
        }
    }
    return $info
}

# Message d'erreur lisible, avec une piste de résolution selon le code HTTP
function Format-GraphErrorMessage {
    param([object]$Info, [string]$Method, [string]$Uri, [int]$Attempts)

    $path = $Uri
    try { $path = ([uri]$Uri).AbsolutePath } catch { }
    $status = if ($Info.StatusCode -gt 0) { "HTTP $($Info.StatusCode)" } else { "erreur réseau" }
    if ($Info.Code) { $status += ", $($Info.Code)" }

    $hint = switch ($Info.StatusCode) {
        400     { "Requête refusée par l'API (paramètres ou rapport non pris en charge par le tenant)." }
        401     { "Authentification refusée : vérifiez l'ID d'application et le secret client (expiré ?)." }
        403     { "Permission insuffisante : accordez à l'application les autorisations Microsoft Graph de type Application DeviceManagementManagedDevices.Read.All et DeviceManagementConfiguration.Read.All, avec consentement administrateur." }
        404     { "Ressource introuvable (appareil supprimé entre-temps ou point de terminaison indisponible)." }
        429     { "Limitation de débit (throttling) persistante après $Attempts tentative(s)." }
        0       { "graph.microsoft.com injoignable (proxy, pare-feu, DNS ou délai dépassé) après $Attempts tentative(s)." }
        default { if ($Info.StatusCode -ge 500) { "Service Microsoft Graph / Intune indisponible après $Attempts tentative(s)." } else { "" } }
    }

    $msg = "Échec de l'appel $Method $path ($status) : $($Info.Message)"
    if ($hint)           { $msg += " $hint" }
    if ($Info.RequestId) { $msg += " (request-id : $($Info.RequestId))" }
    return $msg
}

# Jeton applicatif (client credentials) de la génération en cours : obtenu au
# premier appel, renouvelé 5 minutes avant expiration ou sur demande (-ForceRefresh).
# Retourne $null hors génération (les appelants utilisent alors leur -AccessToken).
# Deux flux parallèles peuvent renouveler en même temps : sans conséquence, les deux
# jetons obtenus sont valides (pas de verrou, donc aucun risque d'interblocage).
function Get-GraphAccessToken {
    param([switch]$ForceRefresh)
    $auth = $script:GraphContext
    if (-not $auth) { return $null }

    if ($ForceRefresh -or -not $auth.AccessToken -or (Get-Date).ToUniversalTime().AddMinutes(5) -ge $auth.ExpiresAtUtc) {
        $body = @{ client_id = $auth.ClientId; scope = "https://graph.microsoft.com/.default"; client_secret = $auth.ClientSecret; grant_type = "client_credentials" }
        $resp = Invoke-GraphApiRequest -NoAuth -Method POST -Uri "https://login.microsoftonline.com/$($auth.TenantId)/oauth2/v2.0/token" `
                                       -Body $body -ContentType "application/x-www-form-urlencoded" -MaxRetries 2
        if (-not $resp.access_token) { throw "Réponse d'authentification inattendue : aucun jeton d'accès renvoyé par Microsoft Entra ID." }
        $lifetime = 3599
        try { if ($resp.expires_in) { $lifetime = [int]$resp.expires_in } } catch { }
        $auth.AccessToken  = $resp.access_token
        $auth.ExpiresAtUtc = (Get-Date).ToUniversalTime().AddSeconds($lifetime)
    }
    return $auth.AccessToken
}

# Appel REST unique avec reprises sur erreur transitoire (voir en-tête de section).
# -RawText  : renvoie le corps décodé en UTF-8 (rapports Intune servis en flux binaire).
# -RawBytes : [v4.0] renvoie le corps brut (archive ZIP d'un export de rapport).
function Invoke-GraphApiRequest {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateSet('GET', 'POST')][string]$Method = 'GET',
        [object]$Body = $null,
        [string]$ContentType = "application/json",
        [string]$AccessToken = "",
        [switch]$NoAuth,
        [switch]$RawText,
        [switch]$RawBytes,
        [int]$MaxRetries = $GraphMaxRetries
    )
    $attempt      = 0
    $tokenRenewed = $false
    while ($true) {
        $attempt++
        $params = @{ Method = $Method; Uri = $Uri; TimeoutSec = $GraphRequestTimeoutSec; ErrorAction = 'Stop' }
        if (-not $NoAuth) {
            Wait-GraphThrottleGate
            $token = Get-GraphAccessToken
            if (-not $token) { $token = $AccessToken }
            $params.Headers = @{ Authorization = "Bearer $token" }
        }
        if ($null -ne $Body) { $params.Body = $Body; $params.ContentType = $ContentType }

        try {
            if ($RawText -or $RawBytes) {
                $web   = Invoke-WebRequest @params -UseBasicParsing
                $bytes = $web.RawContentStream.ToArray()
                if ($RawBytes) { return , $bytes }
                return [System.Text.Encoding]::UTF8.GetString($bytes)
            }
            return (Invoke-RestMethod @params)
        } catch {
            $info = Get-GraphErrorInfo -ErrorRecord $_

            # 401 : jeton expiré ou révoqué -> un seul renouvellement, puis nouvelle tentative
            if ($info.StatusCode -eq 401 -and -not $NoAuth -and -not $tokenRenewed -and $script:GraphContext) {
                $tokenRenewed = $true
                $renewed = $false
                try { [void](Get-GraphAccessToken -ForceRefresh); $renewed = $true } catch { }
                if ($renewed) { continue }
            }

            $retryable = ($info.StatusCode -in @(429, 500, 502, 503, 504)) -or $info.IsNetworkError
            if ($retryable -and $attempt -le $MaxRetries) {
                if ($info.RetryAfter -gt 0) {
                    # Minimum imposé par Graph : jamais raccourci, seulement décalé de 0 à 2 s
                    $delay = [int][math]::Min($info.RetryAfter, 120) + (Get-Random -Minimum 0 -Maximum 3)
                } else {
                    # Backoff exponentiel plafonné, tiré dans [fenêtre/2 ; fenêtre]
                    $window = [math]::Min(60, [math]::Pow(2, $attempt))
                    $delay  = [int][math]::Max(1, [math]::Round(($window / 2) + (Get-Random -Minimum 0.0 -Maximum 1.0) * ($window / 2)))
                }
                # [v4.0] Un 429 signale une limite du tenant : tous les flux patientent
                if ($info.StatusCode -eq 429 -and $script:GraphContext) {
                    $until = [datetime]::UtcNow.AddSeconds($delay)
                    if ($until -gt $script:GraphContext.PauseUntilUtc) { $script:GraphContext.PauseUntilUtc = $until }
                }
                $why  = if ($info.StatusCode -gt 0) { "HTTP $($info.StatusCode)" } else { "erreur réseau" }
                $path = $Uri
                try { $path = ([uri]$Uri).AbsolutePath } catch { }
                Write-Log "[Graph] $why sur $Method $path - nouvelle tentative $attempt/$MaxRetries dans $delay s" -Level WARN
                Start-ResponsiveSleep -Seconds $delay -Message "API Graph : $why, nouvelle tentative $attempt/$MaxRetries"
                continue
            }

            $ex = [System.Exception]::new((Format-GraphErrorMessage -Info $info -Method $Method -Uri $Uri -Attempts $attempt), $_.Exception)
            $ex.Data['StatusCode'] = $info.StatusCode
            $ex.Data['GraphCode']  = $info.Code
            throw $ex
        }
    }
}

# Code HTTP porté par une exception levée par Invoke-GraphApiRequest (0 si inconnu)
function Get-GraphExceptionStatus {
    param([System.Exception]$Exception)
    $e = $Exception
    while ($e) {
        if ($e.Data -and $e.Data.Contains('StatusCode')) { return [int]$e.Data['StatusCode'] }
        $e = $e.InnerException
    }
    return 0
}

function Get-GraphPagedResults {
    param([string]$Url, [string]$AccessToken)
    # [v3.3] Reprises (429 / 5xx / réseau) et erreurs explicites déléguées à Invoke-GraphApiRequest
    $results = [System.Collections.Generic.List[object]]::new()
    $next    = $Url
    while ($next) {
        $resp = Invoke-GraphApiRequest -Method GET -Uri $next -AccessToken $AccessToken
        if ($resp.value) { foreach ($item in $resp.value) { $results.Add($item) } }
        $next = $resp.'@odata.nextLink'
    }
    return $results.ToArray()
}

# Convertit la réponse d'un rapport Intune (flux JSON binaire ou texte) en objet
function ConvertFrom-GraphReportPayload {
    param([object]$Payload)
    if ($null -eq $Payload) { return $null }
    if ($Payload -is [byte[]]) { $Payload = [System.Text.Encoding]::UTF8.GetString($Payload) }
    if ($Payload -is [string]) {
        $text = $Payload.Trim([char]0xFEFF, ' ', "`r", "`n", "`t")
        if (-not $text) { return $null }
        return ($text | ConvertFrom-Json)
    }
    return $Payload
}

# Interroge un rapport Intune (actions beta/deviceManagement/reports/get...Report).
# Réponse : { TotalRowCount, Schema: [{ Column }], Values: [[...]] }, paginée par skip/top.
# Retourne { Columns; Rows } où chaque ligne est un objet nommé selon le schéma.
function Invoke-GraphReportQuery {
    param([Parameter(Mandatory)][string]$ReportName, [string]$AccessToken, [int]$PageSize = 500, [int]$MaxPages = 1000)

    $uri     = "https://graph.microsoft.com/beta/deviceManagement/reports/$ReportName"
    $columns = $null
    $rows    = [System.Collections.Generic.List[object]]::new()
    $skip    = 0
    $total   = -1
    for ($page = 1; $page -le $MaxPages; $page++) {
        $body = @{ skip = $skip; top = $PageSize } | ConvertTo-Json -Compress
        $resp = ConvertFrom-GraphReportPayload (Invoke-GraphApiRequest -Method POST -Uri $uri -Body $body -AccessToken $AccessToken -RawText)
        if ($null -eq $resp -or $null -eq $resp.Schema) { throw "Réponse inattendue du rapport $ReportName (schéma de colonnes absent)." }

        if ($null -eq $columns) { $columns = @($resp.Schema | ForEach-Object { [string]$_.Column }) }
        if ($null -ne $resp.TotalRowCount) { $total = [int]$resp.TotalRowCount }

        # Affectation directe : "$values = if (...) { @($resp.Values) }" ferait passer le
        # tableau par le pipeline, qui déplie une page d'UNE seule ligne en ses cellules
        $values = @()
        if ($null -ne $resp.Values) { $values = @($resp.Values) }
        foreach ($v in $values) {
            $o = [ordered]@{}
            for ($i = 0; $i -lt $columns.Count; $i++) { $o[$columns[$i]] = $v[$i] }
            $rows.Add([pscustomobject]$o)
        }
        $skip += $values.Count

        # Fin : page vide, total atteint, ou (total inconnu) page incomplète
        if ($values.Count -eq 0 -or ($total -ge 0 -and $skip -ge $total) -or ($total -lt 0 -and $values.Count -lt $PageSize)) { break }
    }
    return [pscustomobject]@{ Columns = [string[]]$columns; Rows = $rows }
}

# [FUSION v4 - Proactivité] Requêtes Graph regroupées par lots de 20 via $batch
# (limite du service), avec reprises à deux niveaux :
#  - l'appel $batch lui-même passe par Invoke-GraphApiRequest (429/5xx/réseau/jeton) ;
#  - les sous-requêtes limitées (429/503/504 dans un lot revenu en 200) sont
#    rejouées jusqu'à $SubRequestRetries fois, en respectant leur Retry-After.
# Les sous-requêtes comptent chacune dans le quota Graph : le gain porte sur la
# latence (20 fois moins d'allers-retours), pas sur le throttling.
# $Requests : @{ id = "..."; method = "GET"; url = "/deviceManagement/..." }
# Retour    : sous-réponses { id; status; headers; body } (une sous-requête en erreur
#             définitive est rendue avec son statut : à l'appelant de décider).
function Invoke-GraphBatch {
    param(
        # AllowEmptyCollection : un lot vide est un cas normal (aucun poste à interroger)
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Requests,
        [string]$AccessToken = "",
        [ValidateSet('beta', 'v1.0')][string]$GraphVersion = "beta",
        [int]$BatchSize = 20,
        [int]$SubRequestRetries = 3,
        [string]$Label = "Requêtes groupées"
    )
    $all = [System.Collections.Generic.List[object]]::new()
    if (-not $Requests -or $Requests.Count -eq 0) { return $all.ToArray() }

    # L'endpoint $batch exige un id unique par lot (400 sinon) : premières occurrences seulement
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $Requests) { if ($seen.Add([string]$r.id)) { $list.Add($r) } }

    $endpoint = "https://graph.microsoft.com/$GraphVersion/`$batch"
    $total    = $list.Count
    for ($i = 0; $i -lt $total; $i += $BatchSize) {
        $chunk = $list.GetRange($i, [math]::Min($BatchSize, $total - $i))
        $byId  = @{}
        foreach ($c in $chunk) { $byId[[string]$c.id] = $c }

        $pending = @($chunk)
        for ($round = 0; $pending.Count -gt 0; $round++) {
            $body  = @{ requests = $pending } | ConvertTo-Json -Depth 6 -Compress
            $resp  = Invoke-GraphApiRequest -Method POST -Uri $endpoint -Body $body -AccessToken $AccessToken
            $retry = [System.Collections.Generic.List[object]]::new()
            $hint  = 0
            foreach ($sub in @($resp.responses)) {
                if ($null -eq $sub) { continue }
                $status = 0
                try { $status = [int]$sub.status } catch { }
                if ($status -in @(429, 503, 504) -and $round -lt $SubRequestRetries -and $byId.ContainsKey([string]$sub.id)) {
                    $ra = 0
                    try { [void][int]::TryParse([string]$sub.headers.'Retry-After', [ref]$ra) } catch { }
                    if ($ra -gt $hint) { $hint = $ra }
                    $retry.Add($byId[[string]$sub.id])
                } else {
                    $all.Add($sub)
                }
            }
            $pending = @($retry)
            if ($pending.Count -gt 0) {
                $delay = if ($hint -gt 0) { [int][math]::Min(60, $hint) } else { [int][math]::Min(30, 3 * [math]::Pow(2, $round)) }
                if ($script:GraphContext) {
                    $until = [datetime]::UtcNow.AddSeconds($delay)
                    if ($until -gt $script:GraphContext.PauseUntilUtc) { $script:GraphContext.PauseUntilUtc = $until }
                }
                Write-Log "$Label : $($pending.Count) sous-requête(s) limitée(s), nouvelle tentative dans $delay s." -Level WARN
                Start-ResponsiveSleep -Seconds $delay -Message "$Label : $($pending.Count) requête(s) limitée(s)"
            }
        }
        Set-UiStatus "$Label : $([math]::Min($i + $BatchSize, $total)) / $total..."
    }
    return $all.ToArray()
}

# [v4.0] Collecte séquentielle de listes Graph (repli de la collecte parallèle).
# Corrige le script Proactivité, dont le repli appelait ".ToArray()" sur un tableau
# PowerShell : l'erreur était attrapée et TOUTES les données de santé perdues.
function Invoke-SequentialGraphCollections {
    param([Parameter(Mandatory = $true)][hashtable]$Jobs, [string]$AccessToken = "")
    $result = @{}
    foreach ($name in @($Jobs.Keys)) {
        Write-Step "Collecte : $name..."
        try {
            $result[$name] = [pscustomobject]@{ Items = @(Get-GraphPagedResults -Url $Jobs[$name] -AccessToken $AccessToken); Error = $null }
        } catch {
            $result[$name] = [pscustomobject]@{ Items = @(); Error = $_.Exception.Message }
        }
    }
    return $result
}

# [FUSION v4 - Proactivité] Collecte PARALLÈLE de listes Graph indépendantes, dans un
# pool de runspaces (ForEach-Object -Parallel n'existe qu'en PowerShell 7).
# [v4.0] Différences avec le script Proactivité :
#  - les runspaces exécutent les VRAIES fonctions du script ($GraphWorkerFunctions,
#    injectées via InitialSessionState) au lieu d'une copie texte de la logique de
#    reprise : Retry-After lu sous 5.1 et 7, jeton renouvelé, messages explicites ;
#  - $script:GraphContext (table synchronisée) est partagé : un 429 met tous les
#    flux en pause, un jeton renouvelé sert à tous ;
#  - toute défaillance du pool bascule sur Invoke-SequentialGraphCollections.
# $Jobs : Nom -> URL de première page (pagination faite dans le runspace).
# Retour : Nom -> [pscustomobject]@{ Items = @(...); Error = <message> | $null }
function Invoke-ParallelGraphCollections {
    param([Parameter(Mandatory = $true)][hashtable]$Jobs, [string]$AccessToken = "", [int]$MaxConcurrency = 4)

    if ($Jobs.Count -eq 0) { return @{} }
    if ($Jobs.Count -eq 1 -or $MaxConcurrency -le 1) { return (Invoke-SequentialGraphCollections -Jobs $Jobs -AccessToken $AccessToken) }

    $pool    = $null
    $running = [System.Collections.Generic.List[object]]::new()
    try {
        $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
        foreach ($fn in $GraphWorkerFunctions) {
            $cmd = Get-Command -Name $fn -CommandType Function -ErrorAction Stop
            $iss.Commands.Add([System.Management.Automation.Runspaces.SessionStateFunctionEntry]::new($fn, $cmd.Definition))
        }
        $shared = @{
            GraphContext           = $script:GraphContext
            GraphMaxRetries        = $GraphMaxRetries
            GraphRequestTimeoutSec = $GraphRequestTimeoutSec
            LogFile                = $LogFile
        }
        foreach ($k in $shared.Keys) {
            $iss.Variables.Add([System.Management.Automation.Runspaces.SessionStateVariableEntry]::new($k, $shared[$k], ""))
        }
        $pool = [runspacefactory]::CreateRunspacePool(1, $MaxConcurrency, $iss, $Host)
        $pool.Open()

        $worker = {
            param($JobName, $JobUrl, $Token)
            $ProgressPreference = "SilentlyContinue"
            try {
                [pscustomobject]@{ Name = $JobName; Items = @(Get-GraphPagedResults -Url $JobUrl -AccessToken $Token); Error = $null }
            } catch {
                [pscustomobject]@{ Name = $JobName; Items = @(); Error = $_.Exception.Message }
            }
        }
        foreach ($name in @($Jobs.Keys)) {
            $shell = [powershell]::Create()
            $shell.RunspacePool = $pool
            [void]$shell.AddScript($worker.ToString()).AddArgument($name).AddArgument($Jobs[$name]).AddArgument($AccessToken)
            $running.Add([pscustomobject]@{ Name = $name; Shell = $shell; Handle = $shell.BeginInvoke() })
        }
    } catch {
        Write-Log "Collecte parallèle indisponible ($($_.Exception.Message)) : bascule en collecte séquentielle." -Level WARN
        foreach ($r in $running) { try { $r.Shell.Stop(); $r.Shell.Dispose() } catch { } }
        if ($pool) { try { $pool.Dispose() } catch { } }
        return (Invoke-SequentialGraphCollections -Jobs $Jobs -AccessToken $AccessToken)
    }

    Write-Log "Collecte parallèle : $($running.Count) jeu(x) de données, $MaxConcurrency flux simultanés." -Level INFO
    $total = $running.Count
    while ($true) {
        $done = @($running | Where-Object { $_.Handle.IsCompleted }).Count
        if ($done -ge $total) { break }
        Set-UiStatus "Collecte parallèle : $done / $total jeu(x) de données terminé(s)..."
        # Fenêtre réactive pendant l'attente (sans effet hors interface)
        try { if ($form) { [System.Windows.Forms.Application]::DoEvents() } } catch { }
        Start-Sleep -Milliseconds 200
    }

    $result = @{}
    foreach ($r in $running) {
        try {
            $payload = @($r.Shell.EndInvoke($r.Handle)) | Where-Object { $_ -and $_.PSObject.Properties['Items'] } | Select-Object -First 1
            if ($payload) { $result[$r.Name] = [pscustomobject]@{ Items = @($payload.Items); Error = $payload.Error } }
            else          { $result[$r.Name] = [pscustomobject]@{ Items = @(); Error = "Aucune donnée renvoyée par le flux de collecte." } }
        } catch {
            $result[$r.Name] = [pscustomobject]@{ Items = @(); Error = $_.Exception.Message }
        } finally {
            try { $r.Shell.Dispose() } catch { }
        }
    }
    try { $pool.Close(); $pool.Dispose() } catch { }
    return $result
}

# [FUSION v4 - Proactivité] Lecture défensive d'un résultat de collecte : une source
# absente, vide ou en erreur rend un tableau vide, jamais $null.
function Get-CollectedItems {
    param([hashtable]$Bag, [string]$Name)
    if ($Bag -and $Bag.ContainsKey($Name) -and $Bag[$Name]) { return @($Bag[$Name].Items) }
    return @()
}

# ========================================
# [MODIF v3] ANALYSE DES CAUSES DE NON-CONFORMITÉ (ROOT CAUSE)
# ========================================

# [v3.3] Table unique de traduction des paramètres de conformité Intune en
# libellés lisibles, classés par catégorie. Utilisée par la colonne "RootCause"
# (section "Non Encrypted") et par l'analyse des non-conformités.
# Reconnaît les noms techniques (ex. "Windows10CompliancePolicy.BitLockerEnabled")
# et les libellés anglais des rapports Intune (ex. "Is active", "Minimum OS version"),
# ce qui reprend la table $ErrorLabelMap du script de référence.
# Ordre significatif : la première règle qui correspond l'emporte.
$ComplianceReasonRules = @(
    @{ Pattern = 'BitLocker';                                                          Category = 'Chiffrement';              Reason = 'Chiffrement BitLocker désactivé' }
    @{ Pattern = 'StorageRequireEncryption|RequireEncryption|Encryption of data storage'; Category = 'Chiffrement';           Reason = 'Chiffrement du stockage non conforme' }
    @{ Pattern = 'Firewall';                                                           Category = 'Pare-feu';                 Reason = 'Pare-feu (Firewall) inactif' }
    @{ Pattern = 'OsMinimumVersion|Minimum OS version';                                Category = "Système d'exploitation";   Reason = 'Version OS obsolète (< minimum requis)' }
    @{ Pattern = 'OsMaximumVersion|Maximum OS version';                                Category = "Système d'exploitation";   Reason = 'Version OS non autorisée (> maximum)' }
    @{ Pattern = 'ValidOperatingSystemBuildRanges|Valid operating system build';       Category = "Système d'exploitation";   Reason = 'Build OS hors plage autorisée' }
    @{ Pattern = 'Defender|AntiVirus|AntiSpyware|Antimalware';                         Category = 'Antivirus & menaces';      Reason = 'Antivirus / Defender non conforme' }
    @{ Pattern = 'RtpEnabled|Real-time protection';                                    Category = 'Antivirus & menaces';      Reason = 'Protection temps réel désactivée' }
    @{ Pattern = 'SignatureOutOfDate';                                                 Category = 'Antivirus & menaces';      Reason = 'Signatures antivirus obsolètes' }
    @{ Pattern = 'SecureBoot|Secure Boot';                                             Category = "Intégrité de l'appareil";  Reason = 'Secure Boot désactivé' }
    @{ Pattern = 'Tpm|Trusted Platform Module';                                        Category = "Intégrité de l'appareil";  Reason = 'TPM requis absent ou désactivé' }
    @{ Pattern = 'CodeIntegrity|Code integrity';                                       Category = "Intégrité de l'appareil";  Reason = 'Intégrité du code non conforme' }
    @{ Pattern = 'Password|Passcode';                                                  Category = 'Mot de passe';             Reason = 'Stratégie de mot de passe non conforme' }
    @{ Pattern = 'DeviceThreatProtection|threat level';                                Category = 'Antivirus & menaces';      Reason = 'Niveau de menace appareil trop élevé' }
    @{ Pattern = 'RequireRemainContact|^\s*Is ?active\s*$';                            Category = 'Inscription & activité';   Reason = 'Perte de contact Intune (poste inactif)' }
    @{ Pattern = 'RequireDeviceCompliancePolicyAssigned|compliance policy assigned';   Category = 'Inscription & activité';   Reason = 'Aucune stratégie de conformité assignée' }
    @{ Pattern = 'RequireUserExistence|Enrolled user exists';                          Category = 'Inscription & activité';   Reason = 'Utilisateur inscrit introuvable' }
    @{ Pattern = 'Jailbroken|Jailbreak|Rooted';                                        Category = "Intégrité de l'appareil";  Reason = 'Appareil jailbreaké / rooté' }
)

# [v3.3] Catégorie + raison d'un paramètre en écart. Le nom technique ($Setting)
# est essayé en premier (stable, indépendant de la langue), puis le libellé
# ($SettingName). Paramètre inconnu : libellé brut, catégorie "Autre".
function Get-ComplianceReasonInfo {
    param([string]$Setting, [string]$SettingName)
    foreach ($candidate in @($Setting, $SettingName)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        foreach ($rule in $ComplianceReasonRules) {
            if ($candidate -match $rule.Pattern) { return [pscustomobject]@{ Category = $rule.Category; Reason = $rule.Reason } }
        }
    }
    $raw = if (-not [string]::IsNullOrWhiteSpace($SettingName)) { $SettingName } else { $Setting }
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return [pscustomobject]@{ Category = 'Autre'; Reason = $raw.Trim() }
}

# [MODIF v3] Traduit un paramètre de conformité Intune (settingName brut,
# ex: "Windows10CompliancePolicy.BitLockerEnabled") en libellé lisible
# pour la colonne "RootCause" du dashboard.
# [v3.3] S'appuie sur la table $ComplianceReasonRules (mêmes libellés qu'en v3).
function Get-FriendlyComplianceReason {
    param([string]$Setting, [string]$SettingName)
    $info = Get-ComplianceReasonInfo -Setting $Setting -SettingName $SettingName
    if ($info) { return $info.Reason }
    return $null
}

# [v3.3] Paramètres en écart d'un appareil, via l'API Graph :
#   1) /managedDevices/{id}/deviceCompliancePolicyStates          (état par stratégie)
#   2) .../deviceCompliancePolicyStates/{policyId}/settingStates  (détail par paramètre,
#      appelé seulement s'il n'est pas déjà fourni dans la réponse n°1)
# Endpoint beta requis pour obtenir le détail settingStates de manière fiable.
# Retourne une ligne par paramètre en écart (ou par stratégie si le détail est
# indisponible). Une erreur API sur la liste des stratégies est propagée à
# l'appelant. Résultat mis en cache pour la génération en cours : la section
# "Non Encrypted" réutilise ainsi les postes déjà analysés.
$script:NcSettingsCache = $null
$NcBadStates = @('nonCompliant', 'error', 'conflict')

function Get-DeviceNonComplianceSettings {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [string]$AccessToken = ""
    )
    if ($null -ne $script:NcSettingsCache -and $script:NcSettingsCache.ContainsKey($DeviceId)) {
        return $script:NcSettingsCache[$DeviceId]
    }

    # 1) États des stratégies de conformité de l'appareil
    $polUri       = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/deviceCompliancePolicyStates"
    $policyStates = Get-GraphPagedResults -Url $polUri -AccessToken $AccessToken

    $result = @(ConvertFrom-PolicyStates -DeviceId $DeviceId -PolicyStates $policyStates -AccessToken $AccessToken)
    if ($null -ne $script:NcSettingsCache) { $script:NcSettingsCache[$DeviceId] = $result }
    return $result
}

# [v4.0] Préchargement du cache par lots de 20 appareils ($batch) au lieu d'un appel
# par appareil : utilisé avant l'analyse "Non Encrypted" et le repli par appareil de
# l'analyse des non-conformités. Un appareil dont la sous-requête échoue (404, 403,
# réponse paginée...) n'est pas mis en cache : il sera interrogé individuellement,
# avec la gestion d'erreurs habituelle.
function Initialize-NcSettingsCache {
    param([string[]]$DeviceIds, [string]$AccessToken = "")
    if ($null -eq $script:NcSettingsCache) { return }
    $todo = @($DeviceIds | Where-Object { $_ -and -not $script:NcSettingsCache.ContainsKey($_) } | Select-Object -Unique)
    if ($todo.Count -eq 0) { return }
    $requests = foreach ($id in $todo) { @{ id = $id; method = "GET"; url = "/deviceManagement/managedDevices/$id/deviceCompliancePolicyStates" } }
    try {
        $responses = Invoke-GraphBatch -Requests @($requests) -AccessToken $AccessToken -Label "États de conformité par appareil"
    } catch {
        Write-Log "Préchargement groupé des états de conformité impossible ($($_.Exception.Message)) : interrogation appareil par appareil." -Level WARN
        return
    }
    foreach ($r in $responses) {
        if ([int]$r.status -ne 200 -or -not $r.body -or $r.body.'@odata.nextLink') { continue }
        $script:NcSettingsCache[[string]$r.id] = @(ConvertFrom-PolicyStates -DeviceId ([string]$r.id) -PolicyStates @($r.body.value) -AccessToken $AccessToken)
    }
}

# Lignes "paramètre en écart" à partir des états de stratégies d'un appareil
function ConvertFrom-PolicyStates {
    param([string]$DeviceId, [object[]]$PolicyStates, [string]$AccessToken = "")
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($pol in @($PolicyStates)) {
        if ($null -eq $pol) { continue }
        # On ne détaille que les stratégies en écart (nonCompliant / error / conflict)
        if ("$($pol.state)" -notin $NcBadStates) { continue }

        # 2) Détail des paramètres (settingStates) de la stratégie en écart
        $settingStates = @($pol.settingStates | Where-Object { $null -ne $_ })
        if ($settingStates.Count -eq 0 -and $pol.id) {
            try {
                $ssUri         = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/deviceCompliancePolicyStates/$($pol.id)/settingStates"
                $settingStates = @(Get-GraphPagedResults -Url $ssUri -AccessToken $AccessToken)
            } catch { $settingStates = @() }
        }

        $badSettings = @($settingStates | Where-Object { "$($_.state)" -in $NcBadStates })
        if ($badSettings.Count -gt 0) {
            foreach ($s in $badSettings) {
                $rows.Add([pscustomobject]@{ PolicyName = "$($pol.displayName)"; Setting = "$($s.setting)"; SettingLabel = "$($s.settingName)"; State = "$($s.state)" })
            }
        } else {
            # Pas de granularité disponible : on remonte au moins la stratégie fautive
            $rows.Add([pscustomobject]@{ PolicyName = "$($pol.displayName)"; Setting = ""; SettingLabel = ""; State = "$($pol.state)" })
        }
    }
    return $rows.ToArray()
}

# [MODIF v3] Liste des raisons précises de non-conformité d'un appareil
# (colonne "RootCause"). Retourne un tableau de libellés, causes liées au
# chiffrement en premier.
function Get-DeviceNonComplianceReasons {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [string]$AccessToken = ""
    )
    $reasons = [System.Collections.Generic.List[string]]::new()
    try {
        foreach ($s in @(Get-DeviceNonComplianceSettings -DeviceId $DeviceId -AccessToken $AccessToken)) {
            $label = $null
            if ($s.Setting -or $s.SettingLabel) { $label = Get-FriendlyComplianceReason -Setting $s.Setting -SettingName $s.SettingLabel }
            if (-not $label) { $label = "Non conforme : $($s.PolicyName)" }
            if (-not $reasons.Contains($label)) { [void]$reasons.Add($label) }
        }
    } catch {
        # Une erreur API sur un poste ne doit pas interrompre la génération du dashboard
    }

    # Priorité d'affichage : causes liées au chiffrement en premier
    return @($reasons | Sort-Object { if ($_ -match 'BitLocker|Chiffrement') { 0 } else { 1 } })
}

# ========================================
# [v3.3] ANALYSE DES NON-CONFORMITÉS (CATÉGORIES / RAISONS) VIA L'API GRAPH
# ========================================
# Reprise de la logique du script de référence "Non_compliant.ps1", qui lisait
# un export manuel (CSV / XLSX) : classement par catégorie et par raison,
# ancienneté de synchronisation, actionnabilité, tri par urgence, synthèses et
# CSV. Les données sont désormais interrogées en direct, sans fichier :
#   1) rapport Intune "Noncompliant devices and settings" (tout le tenant) :
#      POST beta/deviceManagement/reports/getNoncompliantDevicesAndSettingsReport
#   2) repli automatique appareil par appareil si le rapport est indisponible :
#      GET beta/deviceManagement/managedDevices/{id}/deviceCompliancePolicyStates
# Aucune erreur de cette analyse n'interrompt la génération du dashboard : elle
# est affichée dans la section et dans le message de fin.

$NcLblVert          = "0 à $NcSeuilVert jours"
$NcLblOrange        = "$($NcSeuilVert + 1) à $NcSeuilOrange jours"
$NcLblRouge         = "$NcSeuilRouge jours et +"
$NcSyncOrder        = @($NcLblVert, $NcLblOrange, $NcLblRouge)
$NcLabelActionnable = "Actionnable"
$NcLabelImpossible  = "Action impossible (hors réseau)"
$NcPrioHaute        = "1 - Priorité haute (synchro ≤ $NcSeuilVert j)"
$NcPrioATraiter     = "2 - À traiter (synchro $($NcSeuilVert + 1)-$NcSeuilOrange j)"
$NcPrioHorsReseau   = "3 - Hors réseau (action impossible)"
$NcPrioFantome      = "4 - Poste fantôme (synchro ≥ $NcSeuilRouge j)"
$NcCategoryOther    = "Autre"
$NcCategoryUnknown  = "Non déterminée"
$NcReasonNotReported = "Raison non remontée par Intune"
$NcReasonApiError    = "Analyse impossible (erreur API)"
$NcSourceReport      = "Rapport Intune « Noncompliant devices and settings » (API Graph)"
$NcSourcePerDevice   = "API Graph, détail par appareil (deviceCompliancePolicyStates)"
$NcIgnoredStatuses   = @('compliant', 'conforme', 'not applicable', 'notapplicable', 'non applicable')

# Colonnes du rapport Intune : les noms varient selon la version du service
# (ex. "SettingNm_loc" / "SettingName"). Correspondance insensible à la casse, aux
# accents et à la ponctuation, dans l'ordre de priorité des alias (repris du
# $ColumnAliases du script de référence).
$NcReportColumnAliases = [ordered]@{
    DeviceId      = @("intunedeviceid", "deviceid", "device id", "managed device id")
    DeviceName    = @("devicename", "device name", "nom du poste", "nom de l'appareil", "computername", "hostname")
    User          = @("upn", "userprincipalname", "user principal name", "primaryuserupn", "useremail", "user email", "utilisateur (upn)")
    SettingLabel  = @("settingnm loc", "settingname loc", "setting name loc", "settingdisplayname", "setting display name")
    Setting       = @("settingname", "setting name", "settingnm", "setting nm", "settingid", "setting id", "parametre intune")
    Policy        = @("policyname", "policy name", "politique", "strategie")
    SettingStatus = @("settingstatus loc", "settingstatus", "setting status", "statut parametre")
    Compliance    = @("compliancestate loc", "compliancestate", "compliance state", "etat conformite")
    OS            = @("os loc", "os", "operatingsystem", "operating system", "platform", "plateforme")
    LastSync      = @("lastcontact", "last contact", "lastsyncdatetime", "last sync", "derniere synchro")
}

# Minuscules, sans accents ni ponctuation (repris du script de référence)
function Get-NormalizedKey {
    param([object]$Text)
    if ($null -eq $Text) { return "" }
    $s  = ([string]$Text).Trim().ToLowerInvariant().Normalize([Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($c in $s.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($c) -ne [Globalization.UnicodeCategory]::NonSpacingMark) { [void]$sb.Append($c) }
    }
    return (($sb.ToString() -replace '[^a-z0-9]+', ' ').Trim())
}

# Ensemble de clés (postes) insensible à la casse
function New-NcKeySet {
    return , ([System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase))
}

# Fait correspondre les colonnes logiques aux colonnes renvoyées par le rapport.
# Deux passes (reprises du script de référence) : correspondances exactes, puis
# correspondances partielles sur les colonnes encore libres.
function Resolve-NcReportColumns {
    param([string[]]$Headers)
    $normalized = @{}
    foreach ($h in $Headers) {
        $k = Get-NormalizedKey $h
        if ($k -and -not $normalized.ContainsKey($k)) { $normalized[$k] = $h }
    }
    $map  = @{}
    $used = New-NcKeySet

    # Passe 1 : correspondances exactes
    foreach ($logical in $NcReportColumnAliases.Keys) {
        $map[$logical] = $null
        foreach ($alias in $NcReportColumnAliases[$logical]) {
            $ak = Get-NormalizedKey $alias
            if ($ak -and $normalized.ContainsKey($ak) -and -not $used.Contains($ak)) {
                $map[$logical] = $normalized[$ak]; [void]$used.Add($ak); break
            }
        }
    }
    # Passe 2 : correspondances partielles (alias de 6 caractères et plus, pour que
    # "os" ne capte pas "osversion" ; jamais pour l'identifiant, "aaddeviceid" != "deviceid")
    foreach ($logical in $NcReportColumnAliases.Keys) {
        if ($map[$logical] -or $logical -eq 'DeviceId') { continue }
        foreach ($alias in $NcReportColumnAliases[$logical]) {
            $ak = Get-NormalizedKey $alias
            if ($ak.Length -lt 6) { continue }
            $hit = $normalized.Keys | Where-Object { -not $used.Contains($_) -and $_ -like "*$ak*" } | Sort-Object Length | Select-Object -First 1
            if ($hit) { $map[$logical] = $normalized[$hit]; [void]$used.Add($hit); break }
        }
    }
    return $map
}

# Valeur d'une colonne logique d'une ligne de rapport ($null si colonne absente)
function Get-NcCell {
    param([object]$Row, [hashtable]$Map, [string]$Name)
    if ($Map[$Name]) { return $Row.($Map[$Name]) }
    return $null
}

# Date hétérogène (DateTime, DateTimeOffset, ISO 8601, JJ/MM/AAAA) -> DateTime UTC.
# $null si vide ou antérieure à 2000 (Intune renvoie 0001-01-01 pour "jamais").
function ConvertTo-UtcDate {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    $d = $null
    if ($Value -is [DateTimeOffset]) {
        $d = $Value.UtcDateTime
    } elseif ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Local) { $d = $Value.ToUniversalTime() }
        else { $d = [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }   # Graph : UTC
    } else {
        $s = ([string]$Value).Trim()
        if (-not $s -or $s -in @('-', 'N/A', 'null', 'None', 'Never', 'Jamais')) { return $null }
        $dto = [DateTimeOffset]::MinValue
        if ($s -match '^\d{4}-\d{2}-\d{2}') {
            if ([DateTimeOffset]::TryParse($s, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)) { $d = $dto.UtcDateTime }
        } elseif ([DateTimeOffset]::TryParse($s, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [Globalization.DateTimeStyles]::AssumeLocal, [ref]$dto)) {
            $d = $dto.UtcDateTime
        }
    }
    if ($null -eq $d -or $d.Year -lt 2000) { return $null }
    return $d
}

# DateTime UTC -> "JJ/MM/AAAA HH:MM" en heure locale (vide si $null)
function Format-LocalDate {
    param([object]$UtcDate)
    if ($null -eq $UtcDate) { return "" }
    return $UtcDate.ToLocalTime().ToString('dd/MM/yyyy HH:mm', [Globalization.CultureInfo]::InvariantCulture)
}

# Traduction des états Intune (nonCompliant, error, conflict...) en français
function ConvertTo-NcStateLabel {
    param([object]$Value)
    $raw = ([string]$Value).Trim()
    switch (Get-NormalizedKey $raw) {
        { $_ -in @('noncompliant', 'not compliant') }      { return 'Non conforme' }
        'error'                                            { return 'Erreur' }
        'conflict'                                         { return 'Conflit' }
        { $_ -in @('ingraceperiod', 'in grace period') }   { return 'Période de grâce' }
        'compliant'                                        { return 'Conforme' }
        { $_ -in @('notapplicable', 'not applicable') }    { return 'Non applicable' }
        'unknown'                                          { return 'Inconnu' }
        default                                            { return $raw }
    }
}

# Pourcentage formaté selon la culture courante (virgule décimale en français)
function Format-NcPercent {
    param([double]$Part, [double]$Total)
    if ($Total -le 0) { return ('{0:0.0}' -f 0) }
    return ('{0:0.0}' -f (100.0 * $Part / $Total))
}

# Ligne brute "poste x paramètre en écart". Les propriétés de l'appareil géré
# (nom, UPN, OS, dernière synchro) font foi ; celles du rapport servent de repli.
function New-NcRawRow {
    param(
        [object]$Device,
        [string]$PolicyName = "", [string]$Setting = "", [string]$SettingLabel = "", [string]$SettingStatus = "",
        [string]$Reason = "", [string]$Category = "",
        [hashtable]$Fallback = @{}
    )
    $pick = {
        param($Primary, $Secondary)
        if (-not [string]::IsNullOrWhiteSpace("$Primary")) { return "$Primary".Trim() }
        if (-not [string]::IsNullOrWhiteSpace("$Secondary")) { return "$Secondary".Trim() }
        return ""
    }
    $lastSync = ConvertTo-UtcDate $Device.LastSyncDateTime
    if ($null -eq $lastSync) { $lastSync = ConvertTo-UtcDate $Fallback.LastSync }

    return [pscustomobject]@{
        DeviceKey         = "$($Device.Id)"
        DeviceName        = & $pick $Device.DeviceName        $Fallback.Name
        UserPrincipalName = & $pick $Device.UserPrincipalName $Fallback.User
        OperatingSystem   = & $pick $Device.OperatingSystem   $Fallback.OS
        ComplianceState   = & $pick $Device.ComplianceState   $Fallback.Compliance
        PolicyName        = $PolicyName.Trim()
        Setting           = $Setting.Trim()
        SettingLabel      = $SettingLabel.Trim()
        SettingStatus     = $SettingStatus.Trim()
        Reason            = $Reason
        Category          = $Category
        LastSyncUtc       = $lastSync
    }
}

# Source 1 : rapport Intune "Noncompliant devices and settings" pour tout le
# tenant, restreint aux postes non conformes du périmètre (exclusion des VM...).
# Lève une exception si le rapport est indisponible ou de structure inconnue.
function Get-NcRowsFromReport {
    param([object[]]$Devices, [string]$AccessToken)

    $report = Invoke-GraphReportQuery -ReportName 'getNoncompliantDevicesAndSettingsReport' -AccessToken $AccessToken -PageSize $NcReportPageSize
    $rows   = [System.Collections.Generic.List[object]]::new()
    if ($report.Rows.Count -eq 0) { return [pscustomobject]@{ Rows = $rows; OutOfScope = 0 } }

    $map     = Resolve-NcReportColumns -Headers $report.Columns
    $colList = $report.Columns -join ', '
    if (-not ($map.DeviceId -or $map.DeviceName)) { throw "Structure du rapport non reconnue : aucune colonne identifiant le poste (colonnes reçues : $colList)." }
    if (-not ($map.Setting -or $map.SettingLabel)) { throw "Structure du rapport non reconnue : aucune colonne de paramètre de conformité (colonnes reçues : $colList)." }

    # Index des postes du périmètre : identifiant Intune, sinon nom
    $byId   = @{}
    $byName = @{}
    foreach ($d in $Devices) {
        if ($d.Id) { $byId["$($d.Id)".ToLowerInvariant()] = $d }
        $n = "$($d.DeviceName)".Trim().ToLowerInvariant()
        if ($n -and -not $byName.ContainsKey($n)) { $byName[$n] = $d }
    }

    $seen       = New-NcKeySet
    $outOfScope = 0
    foreach ($r in $report.Rows) {
        # Paramètres conformes / non applicables : hors sujet
        $status = Get-NcCell $r $map 'SettingStatus'
        if ($null -ne $status -and (Get-NormalizedKey $status) -in $NcIgnoredStatuses) { continue }

        $dev = $null
        $id  = "$(Get-NcCell $r $map 'DeviceId')".Trim().ToLowerInvariant()
        if ($id -and $byId.ContainsKey($id)) { $dev = $byId[$id] }
        if (-not $dev) {
            $n = "$(Get-NcCell $r $map 'DeviceName')".Trim().ToLowerInvariant()
            if ($n -and $byName.ContainsKey($n)) { $dev = $byName[$n] }
        }
        # Poste hors périmètre : VM exclue, période de grâce, conforme depuis...
        if (-not $dev) { $outOfScope++; continue }

        $policy  = "$(Get-NcCell $r $map 'Policy')".Trim()
        $setting = "$(Get-NcCell $r $map 'Setting')".Trim()
        $label   = "$(Get-NcCell $r $map 'SettingLabel')".Trim()
        if (-not $setting -and -not $label) { continue }
        if (-not $seen.Add(("{0}|{1}|{2}|{3}" -f $dev.Id, $policy, $setting, $label))) { continue }   # doublon

        $fallback = @{
            Name       = Get-NcCell $r $map 'DeviceName'
            User       = Get-NcCell $r $map 'User'
            OS         = Get-NcCell $r $map 'OS'
            Compliance = Get-NcCell $r $map 'Compliance'
            LastSync   = Get-NcCell $r $map 'LastSync'
        }
        $rows.Add((New-NcRawRow -Device $dev -PolicyName $policy -Setting $setting -SettingLabel $label -SettingStatus "$status" -Fallback $fallback))
    }
    return [pscustomobject]@{ Rows = $rows; OutOfScope = $outOfScope }
}

# Source 2 (repli) : paramètres en écart, appareil par appareil. Un poste en
# erreur API est conservé avec la raison "Analyse impossible (erreur API)" ;
# une erreur 401 / 403 (droits) interrompt l'analyse : elle toucherait tous les postes.
function Get-NcRowsPerDevice {
    param([object[]]$Devices, [string]$AccessToken, [System.Collections.Generic.List[string]]$Warnings)

    $rows      = [System.Collections.Generic.List[object]]::new()
    $failed    = 0
    $lastError = ""
    $idx       = 0
    # [v4.0] Préchargement groupé ($batch, 20 appareils par appel) ; seuls les postes
    # non résolus par lot sont ensuite interrogés un par un.
    Initialize-NcSettingsCache -DeviceIds @($Devices | ForEach-Object { "$($_.Id)" }) -AccessToken $AccessToken
    foreach ($dev in $Devices) {
        $idx++
        if ($idx -eq 1 -or $idx % 5 -eq 0 -or $idx -eq $Devices.Count) {
            Set-UiStatus "🧩 Non-conformités : analyse par appareil $idx / $($Devices.Count)..."
        }
        try {
            $settings = @(Get-DeviceNonComplianceSettings -DeviceId "$($dev.Id)" -AccessToken $AccessToken)
        } catch {
            if ((Get-GraphExceptionStatus $_.Exception) -in @(401, 403)) { throw }
            $failed++
            $lastError = $_.Exception.Message
            $rows.Add((New-NcRawRow -Device $dev -Reason $NcReasonApiError -Category $NcCategoryUnknown))
            continue
        }
        foreach ($s in $settings) {
            $rows.Add((New-NcRawRow -Device $dev -PolicyName $s.PolicyName -Setting $s.Setting -SettingLabel $s.SettingLabel -SettingStatus $s.State))
        }
    }
    if ($failed -gt 0) {
        if ($failed -eq $Devices.Count) { throw "Aucun des $failed poste(s) non conforme(s) n'a pu être analysé via l'API Graph. Dernière erreur : $lastError" }
        $Warnings.Add("$failed poste(s) sur $($Devices.Count) n'ont pas pu être analysés : ils apparaissent avec la raison « $NcReasonApiError ». Dernière erreur : $lastError")
    }
    return $rows.ToArray()
}

# Traitement (repris d'Invoke-DataProcessing du script de référence) : catégorie
# et raison, ancienneté de synchronisation, tranche, actionnabilité, puis tri
# du plus urgent (synchro la plus récente = corrigeable maintenant) au moins urgent.
function Invoke-NcDataProcessing {
    param([object[]]$Rows, [datetime]$NowUtc)

    $patterns = @($NcKeywordsHorsReseau | ForEach-Object { Get-NormalizedKey $_ } | Where-Object { $_ })
    $out      = [System.Collections.Generic.List[object]]::new()

    foreach ($r in $Rows) {
        # --- Catégorie et raison ---
        $category = $r.Category
        $reason   = $r.Reason
        if (-not $reason) {
            $info = Get-ComplianceReasonInfo -Setting $r.Setting -SettingName $r.SettingLabel
            if ($info)             { $category = $info.Category; $reason = $info.Reason }
            elseif ($r.PolicyName) { $category = $NcCategoryUnknown; $reason = "Non conforme : $($r.PolicyName)" }
            else                   { $category = $NcCategoryUnknown; $reason = $NcReasonNotReported }
        }
        if (-not $category) { $category = $NcCategoryOther }

        # --- Ancienneté de synchronisation (jamais synchronisé = poste fantôme) ---
        $days = $null
        if ($r.LastSyncUtc) {
            $days = [int][math]::Floor(($NowUtc - $r.LastSyncUtc).TotalDays)
            if ($days -lt 0) { $days = 0 }   # marge sur les fuseaux horaires
        }
        $daysRule   = if ($null -eq $days) { [int]::MaxValue } else { $days }
        $syncStatus = if ($daysRule -le $NcSeuilVert) { $NcLblVert } elseif ($daysRule -le $NcSeuilOrange) { $NcLblOrange } else { $NcLblRouge }

        # --- Actionnabilité : erreur "réseau" sur un poste hors ligne depuis > 15 jours ---
        $haystack = Get-NormalizedKey ("{0} {1} {2}" -f $reason, $r.Setting, $r.SettingLabel)
        $isReseau = $false
        foreach ($p in $patterns) { if ($haystack.Contains($p)) { $isReseau = $true; break } }
        $action = if ($daysRule -gt $NcSeuilHorsReseau -and $isReseau) { $NcLabelImpossible } else { $NcLabelActionnable }

        $settingShown = if ($r.Setting) { $r.Setting } else { $r.SettingLabel }
        $out.Add([pscustomobject]@{
            DeviceKey         = $r.DeviceKey
            DeviceName        = $r.DeviceName
            UserPrincipalName = $r.UserPrincipalName
            OperatingSystem   = $r.OperatingSystem
            Category          = $category
            Reason            = $reason
            Setting           = $settingShown
            PolicyName        = $r.PolicyName
            SettingStatus     = ConvertTo-NcStateLabel $r.SettingStatus
            ComplianceState   = ConvertTo-NcStateLabel $r.ComplianceState
            LastSyncUtc       = $r.LastSyncUtc
            DaysSinceSync     = $days
            SyncStatus        = $syncStatus
            Action            = $action
        })
    }

    return @($out | Sort-Object @{ Expression = { if ($null -eq $_.DaysSinceSync) { [int]::MaxValue } else { $_.DaysSinceSync } } },
                                @{ Expression = 'DeviceName' }, @{ Expression = 'Category' }, @{ Expression = 'Reason' })
}

# Indicateurs du tableau de bord (repris de Get-ComplianceKpi) : postes uniques
function Get-NcKpi {
    param([object[]]$Rows)

    $all = New-NcKeySet; $traitables = New-NcKeySet; $vert = New-NcKeySet; $orange = New-NcKeySet
    $fantomes = New-NcKeySet; $rouge30 = New-NcKeySet; $horsReseau = New-NcKeySet
    $minDays = @{}

    foreach ($r in $Rows) {
        $k  = [string]$r.DeviceKey
        $d  = if ($null -eq $r.DaysSinceSync) { [int]::MaxValue } else { [int]$r.DaysSinceSync }
        $ok = $r.Action -eq $NcLabelActionnable
        [void]$all.Add($k)
        if ($ok -and $d -lt $NcSeuilRouge) {
            [void]$traitables.Add($k)
            if ($d -le $NcSeuilVert) { [void]$vert.Add($k) } else { [void]$orange.Add($k) }
        }
        if ($d -ge $NcSeuilRouge -or -not $ok)   { [void]$fantomes.Add($k) }
        if ($d -ge $NcSeuilRouge)                { [void]$rouge30.Add($k) }
        if (-not $ok -and $d -lt $NcSeuilRouge)  { [void]$horsReseau.Add($k) }
        if (-not $minDays.ContainsKey($k) -or $d -lt $minDays[$k]) { $minDays[$k] = $d }
    }

    # Répartition par ancienneté : postes uniques, dans l'ordre des tranches
    $parSynchro = @(foreach ($lbl in $NcSyncOrder) { [pscustomobject]@{ Libelle = $lbl; Valeur = 0 } })
    foreach ($d in $minDays.Values) {
        if ($d -le $NcSeuilVert)       { $parSynchro[0].Valeur += 1 }
        elseif ($d -le $NcSeuilOrange) { $parSynchro[1].Valeur += 1 }
        else                           { $parSynchro[2].Valeur += 1 }
    }

    return [pscustomobject]@{
        Occurrences = @($Rows).Count
        PostesTotal = $all.Count
        Traitables  = $traitables.Count
        Vert        = $vert.Count
        Orange      = $orange.Count
        Fantomes    = $fantomes.Count
        Rouge30     = $rouge30.Count
        HorsReseau  = $horsReseau.Count
        ParSynchro  = $parSynchro
    }
}

# Synthèse par catégorie ET raison : nombre de postes impactés, occurrences,
# répartition par ancienneté et actionnabilité (repris de Get-ErrorSummary)
function Get-NcReasonSummary {
    param([object[]]$Rows, [int]$TotalDevices)

    $acc = [ordered]@{}
    foreach ($r in $Rows) {
        $key = "$($r.Category)|$($r.Reason)"
        if (-not $acc.Contains($key)) {
            $acc[$key] = [pscustomobject]@{
                Category = $r.Category; Reason = $r.Reason; Occurrences = 0
                Devices  = New-NcKeySet; Vert = New-NcKeySet; Orange = New-NcKeySet; Rouge = New-NcKeySet; Actionnables = New-NcKeySet
            }
        }
        $a = $acc[$key]
        $k = [string]$r.DeviceKey
        $d = if ($null -eq $r.DaysSinceSync) { [int]::MaxValue } else { [int]$r.DaysSinceSync }
        $a.Occurrences += 1
        [void]$a.Devices.Add($k)
        if ($d -le $NcSeuilVert)       { [void]$a.Vert.Add($k) }
        elseif ($d -le $NcSeuilOrange) { [void]$a.Orange.Add($k) }
        else                           { [void]$a.Rouge.Add($k) }
        if ($r.Action -eq $NcLabelActionnable -and $d -lt $NcSeuilRouge) { [void]$a.Actionnables.Add($k) }
    }

    $list = foreach ($a in $acc.Values) {
        [pscustomobject][ordered]@{
            'Catégorie'                     = $a.Category
            'Raison'                        = $a.Reason
            'Postes impactés'               = $a.Devices.Count
            'Part des postes NC (%)'        = Format-NcPercent $a.Devices.Count $TotalDevices
            'Occurrences'                   = $a.Occurrences
            "Synchro $NcLblVert"            = $a.Vert.Count
            "Synchro $NcLblOrange"          = $a.Orange.Count
            "Synchro $NcLblRouge"           = $a.Rouge.Count
            "Actionnables (< $NcSeuilRouge j)" = $a.Actionnables.Count
            'Intraitables'                  = $a.Devices.Count - $a.Actionnables.Count
        }
    }
    return @($list | Sort-Object @{ Expression = 'Postes impactés'; Descending = $true }, @{ Expression = 'Occurrences'; Descending = $true },
                                 @{ Expression = 'Catégorie' }, @{ Expression = 'Raison' })
}

# Synthèse par catégorie : postes impactés, raisons distinctes, principales raisons
function Get-NcCategorySummary {
    param([object[]]$Rows, [object[]]$ReasonSummary, [int]$TotalDevices)

    $acc = [ordered]@{}
    foreach ($r in $Rows) {
        $c = [string]$r.Category
        if (-not $acc.Contains($c)) {
            $acc[$c] = [pscustomobject]@{ Category = $c; Occurrences = 0; Devices = New-NcKeySet; Reasons = New-NcKeySet; Actionnables = New-NcKeySet }
        }
        $a = $acc[$c]
        $k = [string]$r.DeviceKey
        $d = if ($null -eq $r.DaysSinceSync) { [int]::MaxValue } else { [int]$r.DaysSinceSync }
        $a.Occurrences += 1
        [void]$a.Devices.Add($k)
        [void]$a.Reasons.Add([string]$r.Reason)
        if ($r.Action -eq $NcLabelActionnable -and $d -lt $NcSeuilRouge) { [void]$a.Actionnables.Add($k) }
    }

    $list = foreach ($a in $acc.Values) {
        $top = @($ReasonSummary | Where-Object { $_.'Catégorie' -eq $a.Category } | Select-Object -First 3 |
                 ForEach-Object { '{0} ({1})' -f $_.'Raison', $_.'Postes impactés' })
        [pscustomobject][ordered]@{
            'Catégorie'                        = $a.Category
            'Postes impactés'                  = $a.Devices.Count
            'Part des postes NC (%)'           = Format-NcPercent $a.Devices.Count $TotalDevices
            'Occurrences'                      = $a.Occurrences
            'Raisons distinctes'               = $a.Reasons.Count
            "Actionnables (< $NcSeuilRouge j)" = $a.Actionnables.Count
            'Intraitables'                     = $a.Devices.Count - $a.Actionnables.Count
            'Principales raisons'              = $top -join ' | '
        }
    }
    return @($list | Sort-Object @{ Expression = 'Postes impactés'; Descending = $true }, @{ Expression = 'Catégorie' })
}

# Synthèse par poste (reprise de Get-SyncSummary) avec priorité de traitement
function Get-NcDeviceSummary {
    param([object[]]$Rows)

    $acc = [ordered]@{}
    foreach ($r in $Rows) {
        $k = [string]$r.DeviceKey
        if (-not $acc.Contains($k)) { $acc[$k] = [System.Collections.Generic.List[object]]::new() }
        $acc[$k].Add($r)
    }

    $list = foreach ($k in $acc.Keys) {
        $g    = $acc[$k]
        $best = $g[0]
        foreach ($x in $g) {
            if ($null -ne $x.DaysSinceSync -and ($null -eq $best.DaysSinceSync -or $x.DaysSinceSync -lt $best.DaysSinceSync)) { $best = $x }
        }
        $daysRule   = if ($null -eq $best.DaysSinceSync) { [int]::MaxValue } else { [int]$best.DaysSinceSync }
        $actionable = @($g | Where-Object { $_.Action -eq $NcLabelActionnable }).Count -gt 0
        $priority   = if ($daysRule -ge $NcSeuilRouge) { $NcPrioFantome }
                      elseif (-not $actionable)        { $NcPrioHorsReseau }
                      elseif ($daysRule -le $NcSeuilVert) { $NcPrioHaute }
                      else                             { $NcPrioATraiter }

        [pscustomobject][ordered]@{
            'Priorité'             = $priority
            'Poste'                = $best.DeviceName
            'Utilisateur (UPN)'    = $best.UserPrincipalName
            'OS'                   = $best.OperatingSystem
            'Dernière synchro'     = Format-LocalDate $best.LastSyncUtc
            'Jours depuis synchro' = $best.DaysSinceSync
            'Statut synchro'       = $best.SyncStatus
            'Nb erreurs'           = $g.Count
            'Catégories'           = (@($g | ForEach-Object { $_.Category } | Sort-Object -Unique) -join ' | ')
            'Raisons'              = (@($g | ForEach-Object { $_.Reason } | Sort-Object -Unique) -join ' | ')
        }
    }
    return @($list | Sort-Object @{ Expression = 'Priorité' },
                                 @{ Expression = { if ($null -eq $_.'Jours depuis synchro') { [int]::MaxValue } else { $_.'Jours depuis synchro' } } },
                                 @{ Expression = 'Poste' })
}

# Détail "poste x paramètre en écart" (colonnes du tableau et du CSV)
function Select-NcDetailView {
    param([object[]]$Rows)
    foreach ($r in $Rows) {
        [pscustomobject][ordered]@{
            'Poste'                = $r.DeviceName
            'Utilisateur (UPN)'    = $r.UserPrincipalName
            'OS'                   = $r.OperatingSystem
            'Catégorie'            = $r.Category
            'Raison'               = $r.Reason
            'Paramètre Intune'     = $r.Setting
            'Stratégie'            = $r.PolicyName
            'Statut paramètre'     = $r.SettingStatus
            'État conformité'      = $r.ComplianceState
            'Dernière synchro'     = Format-LocalDate $r.LastSyncUtc
            'Jours depuis synchro' = $r.DaysSinceSync
            'Statut synchro'       = $r.SyncStatus
            'Action possible'      = $r.Action
        }
    }
}

# Indicateurs clés sous forme de tableau (CSV "0_Indicateurs")
function Get-NcKpiView {
    param([object]$Result, [string]$ClientName = "")
    $k   = $Result.Kpi
    $row = { param($Name, $Value, $Description) [pscustomobject][ordered]@{ 'Indicateur' = $Name; 'Valeur' = $Value; 'Description' = $Description } }
    & $row 'Client'             $ClientName ''
    & $row 'Date de référence'  ($Result.ReferenceDate.ToString('dd/MM/yyyy HH:mm', [Globalization.CultureInfo]::InvariantCulture)) "Interrogation de l'API Graph"
    & $row 'Source des données' $Result.Source ''
    & $row 'Postes non conformes'                            $k.PostesTotal 'Postes uniques non conformes dans le périmètre analysé'
    & $row "Postes actionnables (< $NcSeuilRouge j)"         $k.Traitables  'Joignables et corrigeables à distance'
    & $row "Dont synchro ≤ $NcSeuilVert j"                   $k.Vert        'Synchro très récente : priorité haute'
    & $row "Dont synchro $($NcSeuilVert + 1)-$NcSeuilOrange j" $k.Orange    'Encore actifs, à traiter rapidement'
    & $row 'Postes intraitables'                             $k.Fantomes    'Fantômes ou hors réseau : hors de portée'
    & $row "Dont synchro ≥ $NcSeuilRouge j"                  $k.Rouge30     "Hors ligne depuis au moins $NcSeuilRouge jours"
    & $row "Dont hors réseau > $NcSeuilHorsReseau j"         $k.HorsReseau  'Erreur réseau / pare-feu / BitLocker sur un poste hors ligne'
    & $row "Occurrences d'erreur"                            $k.Occurrences 'Un poste peut cumuler plusieurs erreurs'
}

# Anonymisation cohérente : un même poste / utilisateur garde le même alias
# sur toutes ses lignes (indispensable aux synthèses par poste)
# [v4.0] Mêmes alias que toutes les autres pages (table Get-AnonymizedIdentity)
function Protect-NcRows {
    param([object[]]$Rows)
    foreach ($r in $Rows) {
        $anon = Get-AnonymizedIdentity -RealName $r.DeviceName -RealUpn $r.UserPrincipalName
        $r.DeviceName        = $anon.Name
        $r.UserPrincipalName = $anon.Upn
    }
}

# Résultat vide de l'analyse
function New-NcResult {
    return [pscustomobject]@{
        Success         = $false
        Error           = ""
        Source          = ""
        ReferenceDate   = Get-Date
        Rows            = @()
        Kpi             = $null
        CategorySummary = @()
        ReasonSummary   = @()
        DeviceSummary   = @()
        Warnings        = [System.Collections.Generic.List[string]]::new()
    }
}

# Point d'entrée de l'analyse. Ne lève jamais d'exception : en cas d'échec,
# Success = $false et Error contient le message à afficher.
function Invoke-NonComplianceAnalysis {
    param([object[]]$ScopeDevices, [string]$AccessToken, [switch]$Anonymize)

    $result = New-NcResult
    try {
        $nonCompliant = @($ScopeDevices | Where-Object { "$($_.ComplianceState)" -eq 'noncompliant' })
        $raw          = [System.Collections.Generic.List[object]]::new()

        if ($nonCompliant.Count -gt 0) {
            # --- Source 1 : rapport Intune (tout le tenant en quelques appels) ---
            $haveData = $false
            if ($NcUseBulkReport) {
                try {
                    Set-UiStatus "🧩 Non-conformités : lecture du rapport Intune (API Graph)..."
                    $report = Get-NcRowsFromReport -Devices $nonCompliant -AccessToken $AccessToken
                    if ($report.Rows.Count -gt 0) {
                        foreach ($r in $report.Rows) { $raw.Add($r) }
                        $haveData      = $true
                        $result.Source = $NcSourceReport
                    } else {
                        $result.Warnings.Add("Le rapport Intune n'a renvoyé aucun paramètre en écart pour les $($nonCompliant.Count) poste(s) non conforme(s) du périmètre : analyse appareil par appareil.")
                    }
                } catch {
                    $result.Warnings.Add("Rapport Intune « Noncompliant devices and settings » indisponible, analyse appareil par appareil. Détail : $($_.Exception.Message)")
                }
            }

            # --- Source 2 : repli appareil par appareil ---
            if (-not $haveData) {
                foreach ($r in @(Get-NcRowsPerDevice -Devices $nonCompliant -AccessToken $AccessToken -Warnings $result.Warnings)) { $raw.Add($r) }
                $result.Source = $NcSourcePerDevice
            }

            # --- Postes non conformes sans raison remontée : conservés, pour que le
            #     total corresponde au graphique "Compliance Status" ---
            $withRows = New-NcKeySet
            foreach ($r in $raw) { [void]$withRows.Add($r.DeviceKey) }
            foreach ($d in $nonCompliant) {
                if (-not $withRows.Contains("$($d.Id)")) { $raw.Add((New-NcRawRow -Device $d -Reason $NcReasonNotReported -Category $NcCategoryUnknown)) }
            }

            if ($Anonymize) { Protect-NcRows -Rows $raw }
        }

        Set-UiStatus "🧮 Non-conformités : classement par catégorie et par raison..."
        $rows                   = @(Invoke-NcDataProcessing -Rows $raw -NowUtc $result.ReferenceDate.ToUniversalTime())
        $result.Rows            = $rows
        $result.Kpi             = Get-NcKpi -Rows $rows
        $result.ReasonSummary   = @(Get-NcReasonSummary -Rows $rows -TotalDevices $result.Kpi.PostesTotal)
        $result.CategorySummary = @(Get-NcCategorySummary -Rows $rows -ReasonSummary $result.ReasonSummary -TotalDevices $result.Kpi.PostesTotal)
        $result.DeviceSummary   = @(Get-NcDeviceSummary -Rows $rows)
        $result.Success         = $true
    } catch {
        $result.Success = $false
        $result.Error   = $_.Exception.Message
    }
    return $result
}

# Export CSV des synthèses (séparateur $NcCsvDelimiter, UTF-8 avec BOM pour Excel).
# Chaque fichier est écrit indépendamment : un échec (fichier ouvert dans Excel,
# droits...) est ajouté aux avertissements sans bloquer les autres.
# [v4.0] Jeux de données CSV de l'analyse des non-conformités, écrits par
# Export-AllDatasets avec toutes les autres sections (même dossier horodaté).
# Les clés deviennent les noms de fichiers (préfixe = ordre dans le dossier).
function Get-NcCsvDatasets {
    param([object]$Result, [string]$ClientName = "", [string]$Prefix = "")
    return [ordered]@{
        "${Prefix}0_Indicateurs"            = @(Get-NcKpiView -Result $Result -ClientName $ClientName)
        "${Prefix}1_Synthese_Categories"    = @($Result.CategorySummary)
        "${Prefix}2_Synthese_Raisons"       = @($Result.ReasonSummary)
        "${Prefix}3_Synthese_Postes"        = @($Result.DeviceSummary)
        "${Prefix}4_Detail_Non_Conformites" = @(Select-NcDetailView -Rows $Result.Rows)
    }
}

# Bloc d'en-tête de la section "Non-Compliance Analysis" : KPI, règle de lecture,
# source et exports, avertissements. Rendu uniquement (données déjà calculées).
function New-NcDashboardHeadHtml {
    param([object]$Result, [int]$TotalDevices, [string]$CsvFolder = "", [string[]]$CsvFiles = @())

    # Échappement HTML + crochets (évite l'interprétation [texte](lien) par PSWriteHTML)
    $esc = { param($Text) ((ConvertTo-HtmlSafe $Text) -replace '\[', '&#91;') -replace '\]', '&#93;' }

    $warnHtml = ""
    if ($Result.Warnings.Count -gt 0) {
        $items    = ($Result.Warnings | ForEach-Object { "<li>$(& $esc $_)</li>" }) -join ''
        $warnHtml = "<div class=`"ix-root ix-note ix-note--warn`">$(Get-IconSvg -Name 'alert' -Size 18)<div><strong>Avertissements de l'analyse</strong><ul>$items</ul></div></div>"
    }

    if (-not $Result.Success) {
        $detail = & $esc $Result.Error
        return (New-IxEmptyState -Icon 'alert' -Tone 'danger' -Title "Analyse des non-conformités indisponible" `
                    -Text "L'API Graph n'a pas pu fournir les données de non-conformité ; le reste du dashboard n'est pas affecté.<br><br><b>Détail :</b> $detail") + $warnHtml
    }

    $meta = [System.Collections.Generic.List[string]]::new()
    if ($Result.Source) { $meta.Add("<span class=`"ix-meta__item`">$(Get-IconSvg -Name 'database' -Size 14)Source : $(& $esc $Result.Source)</span>") }
    $meta.Add("<span class=`"ix-meta__item`">$(Get-IconSvg -Name 'calendar' -Size 14)Référence : $($Result.ReferenceDate.ToString("dd/MM/yyyy 'à' HH:mm"))</span>")
    if ($CsvFiles.Count -gt 0) {
        $meta.Add("<span class=`"ix-meta__item`">$(Get-IconSvg -Name 'file-text' -Size 14)$($CsvFiles.Count) synthèses CSV : $(& $esc $CsvFolder)</span>")
    }
    $metaHtml = "<div class=`"ix-root ix-meta`">$($meta -join '')</div>"

    $k = $Result.Kpi
    if ($k.PostesTotal -eq 0) {
        return (New-IxEmptyState -Icon 'shield-ok' -Tone 'success' -Title 'Aucun poste non conforme' `
                    -Text "Aucun appareil du périmètre analysé n'est à l'état « non conforme » dans Intune.") + $metaHtml + $warnHtml
    }

    $cards = [System.Text.StringBuilder]::new()
    [void]$cards.Append((New-IxStatCard -Label 'Postes non conformes' -Value $k.PostesTotal -Icon 'alert' -Color $Colors.Compliance -Caption "<b>$(Get-IxPercent $k.PostesTotal $TotalDevices) %</b> du parc analysé"))
    [void]$cards.Append((New-IxStatCard -Label "Actionnables (&lt; $NcSeuilRouge j)" -Value $k.Traitables -Icon 'tool' -Color $Colors.Accent -Caption 'Joignables et corrigeables à distance'))
    [void]$cards.Append((New-IxStatCard -Label "Dont synchro &le; $NcSeuilVert j" -Value $k.Vert -Icon 'zap' -Color $Colors.Success -Caption 'Synchro très récente : priorité haute'))
    [void]$cards.Append((New-IxStatCard -Label "Dont synchro $($NcSeuilVert + 1)-$NcSeuilOrange j" -Value $k.Orange -Icon 'clock' -Color $Colors.Secondary -Caption 'Encore actifs, à traiter rapidement'))
    [void]$cards.Append((New-IxStatCard -Label 'Intraitables' -Value $k.Fantomes -Icon 'x-circle' -Color $Colors.Danger -Caption 'Fantômes ou hors réseau : hors de portée'))
    [void]$cards.Append((New-IxStatCard -Label "Dont synchro &ge; $NcSeuilRouge j" -Value $k.Rouge30 -Icon 'eye-off' -Color '#ea580c' -Caption "Hors ligne depuis au moins $NcSeuilRouge jours"))
    [void]$cards.Append((New-IxStatCard -Label "Dont hors réseau &gt; $NcSeuilHorsReseau j" -Value $k.HorsReseau -Icon 'wifi-off' -Color $Colors.Warning -Caption 'Erreur réseau / pare-feu / BitLocker'))
    [void]$cards.Append((New-IxStatCard -Label "Occurrences d'erreur" -Value $k.Occurrences -Icon 'list' -Color $Colors.Primary -Caption 'Un poste peut cumuler plusieurs erreurs'))

    $note = "<div class=`"ix-root ix-note`">$(Get-IconSvg -Name 'info' -Size 18)<div><strong>Règle de lecture :</strong> " +
            "une erreur réseau, pare-feu, BitLocker ou d'inactivité ne peut être ni corrigée ni vérifiée tant que le poste ne s'est pas reconnecté. " +
            "Un poste hors ligne depuis plus de $NcSeuilHorsReseau jours porteur de ce type d'erreur est donc classé « $NcLabelImpossible ». " +
            "Les postes synchronisés depuis moins de $NcSeuilRouge jours sont traitables, en priorité ceux de $NcSeuilVert jours ou moins.</div></div>"

    return "<div class=`"ix-root ix-overview`">$($cards.ToString())</div>" + $note + $metaHtml + $warnHtml
}

# ========================================
# [v4.0] COLLECTES REST (REMPLACENT LE MODULE MICROSOFT.GRAPH)
# ========================================
# Les objets sont normalisés aux noms de propriétés du SDK utilisés en v3
# (PascalCase) : tables, filtres et graphiques restent inchangés. Les dates sont des
# DateTime UTC lues depuis la valeur brute, jamais via une conversion en texte (qui
# dépend de la culture régionale du poste : piège décrit dans le script Proactivité).

# Champs utiles des appareils : environ 13 au lieu des ~60 renvoyés sans $select
$ManagedDeviceSelect = "id,deviceName,userPrincipalName,operatingSystem,osVersion,complianceState,lastSyncDateTime,enrolledDateTime,manufacturer,model,isEncrypted,freeStorageSpaceInBytes,totalStorageSpaceInBytes"

# Entier long ou $null (champs de stockage absents sur certaines plateformes)
function ConvertTo-NullableLong {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq "") { return $null }
    $n = [long]0
    if ([long]::TryParse("$Value", [ref]$n)) { return $n }
    return $null
}

# Décimal ou $null. -NonNegative : Endpoint Analytics code "non mesuré" par -1
function ConvertTo-NullableDouble {
    param($Value, [switch]$NonNegative)
    if ($null -eq $Value -or "$Value" -eq "") { return $null }
    $d = 0.0
    if (-not [double]::TryParse(("$Value" -replace ',', '.'), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $null }
    if ($NonNegative -and $d -lt 0) { return $null }
    return $d
}

function Get-ManagedDevicesRest {
    param([string]$AccessToken = "")
    $url  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=$ManagedDeviceSelect&`$top=999"
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $list = [System.Collections.Generic.List[object]]::new()
    foreach ($d in @(Get-GraphPagedResults -Url $url -AccessToken $AccessToken)) {
        # Dédoublonnage par id dès la source (co-gestion, pagination)
        if ($null -eq $d -or -not $d.id -or -not $seen.Add([string]$d.id)) { continue }
        $list.Add([pscustomobject]@{
            Id                       = [string]$d.id
            DeviceName               = [string]$d.deviceName
            UserPrincipalName        = [string]$d.userPrincipalName
            OperatingSystem          = [string]$d.operatingSystem
            OSVersion                = [string]$d.osVersion
            ComplianceState          = [string]$d.complianceState
            LastSyncDateTime         = ConvertTo-UtcDate $d.lastSyncDateTime
            EnrolledDateTime         = ConvertTo-UtcDate $d.enrolledDateTime
            Manufacturer             = [string]$d.manufacturer
            Model                    = [string]$d.model
            IsEncrypted              = ($d.isEncrypted -eq $true)
            FreeStorageSpaceInBytes  = ConvertTo-NullableLong $d.freeStorageSpaceInBytes
            TotalStorageSpaceInBytes = ConvertTo-NullableLong $d.totalStorageSpaceInBytes
        })
    }
    return $list.ToArray()
}

function Get-MobileAppsRest {
    param([string]$AccessToken = "")
    foreach ($a in @(Get-GraphPagedResults -Url "https://graph.microsoft.com/v1.0/deviceAppManagement/mobileApps" -AccessToken $AccessToken)) {
        if ($null -eq $a) { continue }
        [pscustomobject]@{
            DisplayName     = [string]$a.displayName
            Publisher       = [string]$a.publisher
            Id              = [string]$a.id
            CreatedDateTime = ConvertTo-UtcDate $a.createdDateTime
        }
    }
}

# Profils de configuration (modèles) : servent aux Update Rings ET à la vérification
# "profils en erreur" de la page Santé ; lus une seule fois par génération.
function Get-DeviceConfigurationsRest {
    param([string]$AccessToken = "")
    $base  = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations"
    $items = @(Get-GraphPagedResults -Url "${base}?`$select=id,displayName" -AccessToken $AccessToken)
    # Le type (@odata.type) identifie les Update Rings : s'il manque avec $select, relecture complète
    if ($items.Count -gt 0 -and @($items | Where-Object { $_.'@odata.type' }).Count -eq 0) {
        $items = @(Get-GraphPagedResults -Url $base -AccessToken $AccessToken)
    }
    return $items
}

# [FUSION v4 - Proactivité] Rapports Intune : tableau d'objets, ou { Schema, Values }
# à lignes positionnelles. Toujours restitués en objets à propriétés nommées.
function ConvertFrom-IntuneReportPayload {
    param($Payload)
    if ($null -eq $Payload) { return @() }
    if ($Payload.PSObject.Properties['Schema'] -and $Payload.PSObject.Properties['Values']) {
        if ($null -eq $Payload.Values) { return @() }
        $columns = @($Payload.Schema | ForEach-Object { [string]$_.Column })
        $rows    = [System.Collections.Generic.List[object]]::new()
        foreach ($line in @($Payload.Values)) {
            $cells  = @($line)
            $record = [ordered]@{}
            for ($i = 0; $i -lt $columns.Count; $i++) {
                $cell = $null
                if ($i -lt $cells.Count) { $cell = $cells[$i] }
                $record[$columns[$i]] = $cell
            }
            $rows.Add([pscustomobject]$record)
        }
        return $rows.ToArray()
    }
    return @($Payload)
}

# Rapport des échecs d'installation d'applications (v3, désormais via Invoke-GraphApiRequest)
function Get-FailedAppsReportRest {
    param([string]$AccessToken = "")
    $body = @{ top = 5000; orderBy = @("FailedDeviceCount desc") } | ConvertTo-Json -Compress
    $resp = Invoke-GraphApiRequest -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/getFailedMobileAppsReport" `
                                   -Body $body -AccessToken $AccessToken -RawText
    return @(ConvertFrom-IntuneReportPayload (ConvertFrom-GraphReportPayload $resp))
}

# Statuts des Windows Update Rings (logique v3 : comptes utilisateurs réels seulement).
# [v4.0] Erreurs remontées en avertissements (la v3 les masquait par un "catch {}" vide).
function Get-UpdateRingData {
    param([object[]]$Configurations, [string]$AccessToken = "", [System.Collections.Generic.List[string]]$Warnings)
    $summary = [System.Collections.Generic.List[object]]::new()
    $devices = [System.Collections.Generic.List[object]]::new()
    $rings   = @($Configurations | Where-Object { "$($_.'@odata.type')" -like "*windowsUpdateForBusinessConfiguration*" })
    foreach ($pol in $rings) {
        try {
            $statuses = @(Get-GraphPagedResults -Url "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($pol.id)/deviceStatuses?`$top=1000" -AccessToken $AccessToken)
        } catch {
            if ($Warnings) { $Warnings.Add("Update Ring « $($pol.displayName) » : $($_.Exception.Message)") }
            continue
        }
        $userStatuses = @($statuses | Where-Object { $_ -and $_.userName -and "$($_.userName)".Trim() -ne "" -and $_.userName -ne "System account" -and $_.userName -match "@" })
        if ($userStatuses.Count -eq 0) { continue }
        $count = @{ compliant = 0; error = 0; conflict = 0; notApplicable = 0; inProgress = 0 }
        foreach ($ds in $userStatuses) {
            $st = "$($ds.status)"
            if ($count.ContainsKey($st)) { $count[$st]++ }
            $devices.Add([pscustomobject]@{
                RingName     = [string]$pol.displayName
                DeviceName   = [string]$ds.deviceDisplayName
                UserName     = [string]$ds.userName
                Status       = $st
                LastReported = ConvertTo-UtcDate $ds.lastReportedDateTime
            })
        }
        $summary.Add([pscustomobject]@{
            RingName      = [string]$pol.displayName
            DeviceCount   = $userStatuses.Count
            Succeeded     = $count.compliant
            Error         = $count.error
            Conflict      = $count.conflict
            NotApplicable = $count.notApplicable
            InProgress    = $count.inProgress
        })
    }
    return [pscustomobject]@{
        Summary   = @($summary | Sort-Object RingName)
        Devices   = @($devices | Sort-Object RingName, DeviceName)
        RingCount = $rings.Count
    }
}

# ========================================
# [FUSION v4 - Proactivité] SANTÉ & PROACTIVITÉ DES POSTES
# ========================================
# OBJECTIF : donner de quoi AGIR AVANT l'incident. Reprise de l'onglet "Remédiation &
# santé des postes" du script Proactivité (14 vérifications, seuils
# $RemediationThresholds), affichée dans la page "Santé & proactivité" qui absorbe
# l'ancienne page "Optimisation du parc".
# SOURCES :
#   * disque, inactivité, conformité : appareils déjà collectés (aucun appel en plus)
#   * Endpoint Analytics (beta) : scores, performances de démarrage, batteries,
#     fiabilité des applications, historique de démarrage ; + états BitLocker
#     -> listes indépendantes, lues en parallèle (Invoke-ParallelGraphCollections)
#   * Defender : windowsProtectionState, un appel par poste WINDOWS, regroupés en $batch
#   * Mises à jour Windows : export du rapport Intune QualityUpdateDeviceStatusByPolicy
#   * Profils de configuration : deviceStatuses de chaque profil, regroupés en $batch
# PRÉREQUIS : l'Analyse des points de terminaison doit être ACTIVÉE dans Intune ; à
# défaut, les vérifications concernées restent vides et un avertissement l'indique.
# Permissions : DeviceManagementManagedDevices.Read.All et
# DeviceManagementConfiguration.Read.All (déjà requises par le reste du script).
# Chaque vérification est en "meilleur effort" : une source indisponible n'arrête
# jamais la génération.
# [v4.0] Écarts volontaires avec le script Proactivité :
#   - accumulations en List : les "+=" coûtaient 8 s contre 0,2 s sur 20 000 postes ;
#   - Defender limité aux postes Windows (iOS / Android renvoyaient tous une erreur) ;
#   - historique de démarrage rapproché AUSSI par deviceId / startTime, les noms de
#     champs du schéma Graph documenté (le script Proactivité ne cherchait que
#     deviceName / startupDateTime) ;
#   - fiabilité applicative : seule la vue par poste, réellement exploitée, est lue ;
#   - conformité : motif exact repris de l'analyse des non-conformités (page Sécurité) ;
#   - postes présents dans Endpoint Analytics mais hors du périmètre analysé (VM
#     exclues, appareils supprimés) ignorés, pour des totaux cohérents entre pages ;
#   - pas de boutons Sync / Reboot / Remédier : l'outil reste en lecture seule.

# Vérifications activables une à une (onglet "Proactivité" de la fenêtre). Une
# vérification décochée n'est ni collectée ni affichée : aucun appel superflu.
$HealthCheckDefs = @(
    @{ Key = "Disk";           Text = "Espace disque";               Col = 0; Row = 0
       Hint = "Alerte sous $($RemediationThresholds.DiskFreePctWarning) % d'espace libre, critique sous $($RemediationThresholds.DiskFreePctCritical) % ou $($RemediationThresholds.DiskFreeGbCritical) Go." }
    @{ Key = "Inactivity";     Text = "Inactivité (synchro)";        Col = 0; Row = 1
       Hint = "Postes sans synchronisation depuis plus de $($RemediationThresholds.StaleDaysWarning) jours ; critique au-delà de $($RemediationThresholds.StaleDaysCritical) jours." }
    @{ Key = "Boot";           Text = "Démarrage lent / HDD";        Col = 0; Row = 2
       Hint = "Démarrage au-delà de $($RemediationThresholds.BootSlowSeconds) s, et postes encore équipés d'un disque mécanique (Endpoint Analytics)." }
    @{ Key = "Bsod";           Text = "Écrans bleus / redémarrages"; Col = 0; Row = 3
       Hint = "Écrans bleus sur 14 jours (critique à partir de $($RemediationThresholds.BsodCritical)) et redémarrages anormalement fréquents (Endpoint Analytics)." }
    @{ Key = "Battery";        Text = "Batterie (score)";            Col = 1; Row = 0
       Hint = "Score de santé composite Endpoint Analytics, alerte sous $($RemediationThresholds.BatteryPoor)." }
    @{ Key = "BatteryDetail";  Text = "Batterie (capacité, âge)";    Col = 1; Row = 1
       Hint = "Capacité maximale restante, âge et autonomie estimée. Alerte sous $($RemediationThresholds.BatteryCapacityPoor) % de capacité, critique sous $($RemediationThresholds.BatteryCapacityCrit) %." }
    @{ Key = "EaScore";        Text = "Score Endpoint Analytics";    Col = 1; Row = 2
       Hint = "Score global du poste, alerte sous $($RemediationThresholds.ScoreLow)." }
    @{ Key = "Uptime";         Text = "Uptime (estimation)";         Col = 1; Row = 3
       Hint = "Temps écoulé depuis le dernier démarrage connu, au-delà de $($RemediationThresholds.UptimeWarningDays) jours. Donnée agrégée quotidiennement par Microsoft : c'est une estimation." }
    @{ Key = "BitLocker";      Text = "BitLocker (états avancés)";   Col = 2; Row = 0
       Hint = "Chiffrement du volume système et véritables échecs de protection (managedDeviceEncryptionStates). La page Sécurité s'appuie, elle, sur l'indicateur isEncrypted de l'appareil." }
    @{ Key = "Defender";       Text = "Antivirus (Defender)";        Col = 2; Row = 1
       Hint = "Protection en temps réel désactivée et signatures en retard de plus de $($RemediationThresholds.SignatureStaleDays) jours. Un appel par poste Windows, regroupés par 20." }
    @{ Key = "WindowsUpdate";  Text = "Mises à jour Windows";        Col = 2; Row = 2
       Hint = "Échecs de mises à jour qualité (rapport Intune). Nécessite des postes rattachés à une stratégie de mises à jour Windows ; peut ajouter jusqu'à 90 s de génération du rapport." }
    @{ Key = "AppReliability"; Text = "Fiabilité des applications";  Col = 2; Row = 3
       Hint = "Application qui plante le plus souvent sur chaque poste, à partir de $($RemediationThresholds.AppCrashWarning) plantages (Endpoint Analytics)." }
    @{ Key = "Compliance";     Text = "Conformité Intune";           Col = 3; Row = 0
       Hint = "État de conformité du poste, avec le motif exact si l'analyse des non-conformités est active. Aucun appel réseau supplémentaire." }
    @{ Key = "ConfigProfile";  Text = "Profils de configuration";    Col = 3; Row = 1
       Hint = "Postes en erreur ou en conflit d'application d'un profil (modèles et catalogue de paramètres) : le poste reste conforme mais n'applique pas le paramétrage attendu." }
)

# Libellés de sévérité (préfixe numérique : le tri des tableaux suit l'urgence)
$HealthSeverityLabels = @{ crit = "1 - Critique"; warn = "2 - À surveiller"; ok = "3 - OK" }
$HealthSeverityRank   = @{ ok = 0; warn = 1; crit = 2 }

# Clé de rapprochement par nom de poste : minuscules, sans accents ni ponctuation
function Get-DeviceNameKey {
    param([string]$Name)
    return ((Get-NormalizedKey $Name) -replace ' ', '')
}

# [FUSION v4 - Proactivité] Valeur d'une propriété quelle que soit sa casse ou sa graphie
# (DeviceName / deviceName / deviceDisplayName). [v4.0] Valeur BRUTE, pas convertie en
# texte : une date convertie en texte dépendrait de la culture régionale du poste.
function Get-PropValue {
    param($Object, [string[]]$Names)
    if ($null -eq $Object) { return $null }
    foreach ($n in $Names) {
        $p = $Object.PSObject.Properties[$n]
        if ($p -and $null -ne $p.Value -and "$($p.Value)" -ne "") { return $p.Value }
    }
    return $null
}

# Index nom de poste normalisé -> premier élément rencontré
function New-NameIndex {
    param([object[]]$Items, [string[]]$NameProps)
    $index = @{}
    foreach ($it in @($Items)) {
        if ($null -eq $it) { continue }
        $k = Get-DeviceNameKey ([string](Get-PropValue $it $NameProps))
        if ($k -and -not $index.ContainsKey($k)) { $index[$k] = $it }
    }
    return $index
}

# [FUSION v4 - Proactivité] Échecs de mises à jour qualité Windows, via l'export de
# rapport Intune (job asynchrone : création, attente, téléchargement).
# [v4.0] Archive ZIP lue en mémoire (plus de fichiers temporaires), attente réactive,
# appels via Invoke-GraphApiRequest (reprises, jeton). Lève une exception si le
# rapport n'aboutit pas : l'appelant la transforme en avertissement.
function Get-WindowsUpdateFailures {
    param([string]$AccessToken = "", [int]$MaxWaitSeconds = 90)
    $jobsUri = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs"
    $body    = @{ reportName = "QualityUpdateDeviceStatusByPolicy"; format = "json" } | ConvertTo-Json -Compress
    $job     = Invoke-GraphApiRequest -Method POST -Uri $jobsUri -Body $body -AccessToken $AccessToken
    $status  = $job
    $waited  = 0
    while ("$($status.status)" -notin @('completed', 'failed') -and $waited -lt $MaxWaitSeconds) {
        Start-ResponsiveSleep -Seconds 5 -Message "Rapport des mises à jour Windows en cours de génération"
        $waited += 5
        $status = Invoke-GraphApiRequest -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$($job.id)')" -AccessToken $AccessToken
    }
    if ("$($status.status)" -ne 'completed' -or -not $status.url) {
        throw "rapport non abouti après $waited s (statut « $($status.status) »)"
    }

    # URL de téléchargement signée (SAS) : surtout pas d'en-tête d'autorisation
    $bytes = Invoke-GraphApiRequest -NoAuth -Uri $status.url -RawBytes
    Add-Type -AssemblyName System.IO.Compression
    $zip = [System.IO.Compression.ZipArchive]::new([System.IO.MemoryStream]::new($bytes))
    try {
        $entry = @($zip.Entries | Where-Object { $_.Name -like '*.json' }) | Select-Object -First 1
        if (-not $entry) { return @() }
        $reader = [System.IO.StreamReader]::new($entry.Open(), [System.Text.Encoding]::UTF8)
        try { $json = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } finally { $zip.Dispose() }
    return @(ConvertFrom-IntuneReportPayload (ConvertFrom-GraphReportPayload $json))
}

# [FUSION v4 - Proactivité] Postes en ERREUR ou en CONFLIT d'application d'un profil de
# configuration : modèles (deviceConfigurations, déjà lus pour les Update Rings) ET
# catalogue de paramètres (configurationPolicies). Statuts interrogés PAR PROFIL (quelques
# dizaines d'appels groupés en $batch) et non par appareil (des milliers d'appels).
function Get-ConfigurationProfileErrors {
    param(
        [object[]]$DeviceConfigurations, [string]$AccessToken = "", [int]$MaxProfiles = 300,
        [System.Collections.Generic.List[string]]$Warnings
    )
    $profiles = [System.Collections.Generic.List[object]]::new()
    foreach ($p in @($DeviceConfigurations)) {
        if ($p -and $p.id) { $profiles.Add([pscustomobject]@{ Id = [string]$p.id; Name = [string]$p.displayName; Segment = "deviceConfigurations" }) }
    }
    try {
        foreach ($p in @(Get-GraphPagedResults -Url "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$select=id,name&`$top=200" -AccessToken $AccessToken)) {
            if ($p -and $p.id) { $profiles.Add([pscustomobject]@{ Id = [string]$p.id; Name = [string]$p.name; Segment = "configurationPolicies" }) }
        }
    } catch {
        if ($Warnings) { $Warnings.Add("Catalogue de paramètres indisponible : $($_.Exception.Message)") }
    }

    $list = @($profiles)
    if ($list.Count -eq 0) { return @() }
    if ($list.Count -gt $MaxProfiles) {
        if ($Warnings) { $Warnings.Add("Profils de configuration : $($list.Count) profils, analyse limitée aux $MaxProfiles premiers (`$MaxConfigProfilesAnalyzed).") }
        $list = $list[0..($MaxProfiles - 1)]
    }

    # Pas de $select / $filter sur deviceStatuses : leur prise en charge varie selon le
    # type de profil, et un rejet en 400 ferait perdre tout le lot. Tri local.
    $requests = [System.Collections.Generic.List[object]]::new()
    $byId     = @{}
    $n        = 0
    foreach ($p in $list) {
        $n++
        $byId["cfg$n"] = $p
        $requests.Add(@{ id = "cfg$n"; method = "GET"; url = "/deviceManagement/$($p.Segment)/$($p.Id)/deviceStatuses?`$top=999" })
    }
    Write-Step "🧩 Santé des postes : application de $($list.Count) profil(s) de configuration..."
    $responses = Invoke-GraphBatch -Requests $requests.ToArray() -AccessToken $AccessToken -Label "Profils de configuration"

    $failures  = [System.Collections.Generic.List[object]]::new()
    $truncated = 0
    foreach ($r in @($responses)) {
        $status = 0
        try { $status = [int]$r.status } catch { }
        if ($status -ne 200 -or -not $r.body -or -not $byId.ContainsKey([string]$r.id)) { continue }
        $p = $byId[[string]$r.id]
        if ($r.body.'@odata.nextLink') { $truncated++ }
        foreach ($st in @($r.body.value)) {
            if ($null -eq $st) { continue }
            $state = "$($st.status)".Trim().ToLowerInvariant()
            if ($state -ne 'error' -and $state -ne 'conflict') { continue }
            $failures.Add([pscustomobject]@{
                DeviceName   = [string]$st.deviceDisplayName
                UserName     = [string]$st.userName
                ProfileName  = $p.Name
                Status       = $state
                LastReported = ConvertTo-UtcDate $st.lastReportedDateTime
            })
        }
    }
    if ($truncated -gt 0 -and $Warnings) { $Warnings.Add("Profils de configuration : $truncated profil(s) de plus de 999 postes, décompte partiel pour ceux-là.") }
    Write-Log "Profils de configuration : $($failures.Count) application(s) en erreur ou en conflit sur $($list.Count) profil(s)."
    return $failures.ToArray()
}

# [FUSION v4 - Proactivité] Collecte des données de santé selon les vérifications cochées
# (une vérification décochée ne déclenche aucun appel). Ne lève pas d'exception : toute
# source indisponible devient un avertissement affiché dans la page et le message final.
function Invoke-HealthCollection {
    param(
        [object[]]$Devices, [hashtable]$Checks, [object[]]$DeviceConfigurations,
        [string]$AccessToken = "", [bool]$Parallel = $true,
        [System.Collections.Generic.List[string]]$Warnings
    )
    $base = "https://graph.microsoft.com/beta/deviceManagement"
    $jobs = @{}
    if ($Checks.EaScore -or $Checks.Battery) { $jobs["Scores Endpoint Analytics"]  = "$base/userExperienceAnalyticsDeviceScores?`$top=999" }
    if ($Checks.Boot -or $Checks.Bsod)       { $jobs["Performances de démarrage"]  = "$base/userExperienceAnalyticsDevicePerformance?`$top=999" }
    if ($Checks.BatteryDetail)               { $jobs["Santé des batteries"]        = "$base/userExperienceAnalyticsBatteryHealthDevicePerformance?`$top=999" }
    if ($Checks.AppReliability)              { $jobs["Fiabilité des applications"] = "$base/userExperienceAnalyticsAppHealthDevicePerformanceDetails?`$top=999" }
    if ($Checks.Uptime)                      { $jobs["Historique de démarrage"]    = "$base/userExperienceAnalyticsDeviceStartupHistory?`$top=999" }
    if ($Checks.BitLocker)                   { $jobs["État BitLocker"]             = "$base/managedDeviceEncryptionStates?`$top=999" }

    $collected = @{}
    if ($jobs.Count -gt 0) {
        if ($Parallel) {
            Write-Step "🩺 Santé des postes : $($jobs.Count) jeu(x) de données lus en parallèle..."
            $collected = Invoke-ParallelGraphCollections -Jobs $jobs -AccessToken $AccessToken -MaxConcurrency $MaxParallelCollections
        } else {
            $collected = Invoke-SequentialGraphCollections -Jobs $jobs -AccessToken $AccessToken
        }
        foreach ($name in @($collected.Keys)) {
            if ($collected[$name].Error) {
                $Warnings.Add("$name indisponible : $($collected[$name].Error)")
                Write-Log "$name indisponible : $($collected[$name].Error)" -Level WARN
            } else {
                Write-Log "$name : $(@($collected[$name].Items).Count) élément(s)."
            }
        }
    }

    # ----- Defender : pas de liste globale côté Graph, un appel par poste Windows -----
    $defender = @{}
    if ($Checks.Defender) {
        $windows = @($Devices | Where-Object { $_.Id -and "$($_.OperatingSystem)" -like "Windows*" })
        if ($windows.Count -gt 0) {
            Write-Step "🛡 Santé des postes : état Defender de $($windows.Count) poste(s) Windows (lots de 20)..."
            $requests = foreach ($d in $windows) { @{ id = [string]$d.Id; method = "GET"; url = "/deviceManagement/managedDevices/$($d.Id)/windowsProtectionState" } }
            try {
                $denied = 0
                foreach ($r in @(Invoke-GraphBatch -Requests @($requests) -AccessToken $AccessToken -Label "État Defender")) {
                    $st = 0
                    try { $st = [int]$r.status } catch { }
                    if ($st -eq 200 -and $r.body) { $defender[[string]$r.id] = $r.body }
                    elseif ($st -in @(401, 403)) { $denied++ }
                }
                if ($denied -gt 0) { $Warnings.Add("État Defender : accès refusé pour $denied poste(s) (permission DeviceManagementManagedDevices.Read.All).") }
                Write-Log "État Defender : $($defender.Count) poste(s) sur $($windows.Count)."
            } catch {
                $Warnings.Add("État Defender indisponible : $($_.Exception.Message)")
            }
        }
    }

    # ----- Mises à jour qualité Windows (rapport asynchrone) -----
    $updateFailures = @()
    if ($Checks.WindowsUpdate) {
        Write-Step "🔄 Santé des postes : rapport des mises à jour qualité Windows..."
        try { $updateFailures = @(Get-WindowsUpdateFailures -AccessToken $AccessToken) }
        catch { $Warnings.Add("Mises à jour Windows : $($_.Exception.Message). Prérequis : postes rattachés à une stratégie de mises à jour qualité Windows.") }
    }

    # ----- Profils de configuration en erreur / conflit -----
    $configErrors = @()
    if ($Checks.ConfigProfile) {
        try { $configErrors = @(Get-ConfigurationProfileErrors -DeviceConfigurations $DeviceConfigurations -AccessToken $AccessToken -MaxProfiles $MaxConfigProfilesAnalyzed -Warnings $Warnings) }
        catch { $Warnings.Add("Profils de configuration indisponibles : $($_.Exception.Message)") }
    }

    return [pscustomobject]@{
        Scores              = @(Get-CollectedItems -Bag $collected -Name "Scores Endpoint Analytics")
        Performance         = @(Get-CollectedItems -Bag $collected -Name "Performances de démarrage")
        Battery             = @(Get-CollectedItems -Bag $collected -Name "Santé des batteries")
        AppReliability      = @(Get-CollectedItems -Bag $collected -Name "Fiabilité des applications")
        StartupHistory      = @(Get-CollectedItems -Bag $collected -Name "Historique de démarrage")
        BitLocker           = @(Get-CollectedItems -Bag $collected -Name "État BitLocker")
        Defender            = $defender
        UpdateFailures      = $updateFailures
        ConfigProfileErrors = $configErrors
    }
}

function New-HealthIssue([string]$Label, [string]$Sev, [string]$Detail) {
    return [pscustomobject]@{ Label = $Label; Sev = $Sev; Detail = $Detail }
}

# Nombre au format français pour les libellés ("12,5")
function Format-FrNumber($Value) { return ("$Value" -replace '\.', ',') }

# [FUSION v4 - Proactivité] Croise appareils et données de santé, en déduit par poste
# les indicateurs, une sévérité (crit / warn / ok) et des ACTIONS DE REMÉDIATION
# regroupées (une action -> la liste des postes concernés).
function Build-RemediationData {
    param(
        [object[]]$Devices, [object]$Health, [hashtable]$Checks,
        [hashtable]$ComplianceReasonsById = @{}, [bool]$AnonymizeData = $false
    )
    $T    = $RemediationThresholds
    $rank = $HealthSeverityRank

    $scoreByName = New-NameIndex -Items $Health.Scores      -NameProps @('deviceName')
    $perfByName  = New-NameIndex -Items $Health.Performance -NameProps @('deviceName')
    $blByName    = New-NameIndex -Items $Health.BitLocker   -NameProps @('deviceName')
    $battByName  = New-NameIndex -Items $Health.Battery     -NameProps @('deviceName', 'deviceDisplayName')

    # Fiabilité applicative : par poste, l'application qui plante le plus souvent
    $crashCounts = @{}
    foreach ($e in @($Health.AppReliability)) {
        if ($null -eq $e -or "$($e.eventType)" -notmatch '(?i)crash|hang') { continue }
        $k = Get-DeviceNameKey ([string](Get-PropValue $e @('deviceDisplayName', 'deviceName')))
        if (-not $k) { continue }
        if (-not $crashCounts.ContainsKey($k)) { $crashCounts[$k] = @{} }
        $app = [string]$e.appDisplayName
        if (-not $crashCounts[$k].ContainsKey($app)) { $crashCounts[$k][$app] = 0 }
        $crashCounts[$k][$app]++
    }
    $crashByName = @{}
    foreach ($k in @($crashCounts.Keys)) {
        $top = $crashCounts[$k].GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
        if ($top) { $crashByName[$k] = [pscustomobject]@{ AppName = [string]$top.Key; Count = [int]$top.Value } }
    }

    # Dernier démarrage connu : par deviceId / startTime (schéma Graph documenté), puis par nom
    $bootById   = @{}
    $bootByName = @{}
    foreach ($h in @($Health.StartupHistory)) {
        if ($null -eq $h) { continue }
        $dt = ConvertTo-UtcDate (Get-PropValue $h @('startTime', 'startupDateTime', 'eventDateTime', 'lastBootUpTime'))
        if (-not $dt) { continue }
        $id = "$(Get-PropValue $h @('deviceId', 'managedDeviceId'))".ToLowerInvariant()
        if ($id -and (-not $bootById.ContainsKey($id) -or $dt -gt $bootById[$id])) { $bootById[$id] = $dt }
        $k = Get-DeviceNameKey ([string](Get-PropValue $h @('deviceName', 'deviceDisplayName')))
        if ($k -and (-not $bootByName.ContainsKey($k) -or $dt -gt $bootByName[$k])) { $bootByName[$k] = $dt }
    }

    # Profils de configuration en échec, regroupés par poste (au plus 6 noms conservés)
    $cfgByName = @{}
    foreach ($e in @($Health.ConfigProfileErrors)) {
        if ($null -eq $e) { continue }
        $k = Get-DeviceNameKey $e.DeviceName
        if (-not $k) { continue }
        if (-not $cfgByName.ContainsKey($k)) {
            $cfgByName[$k] = [pscustomobject]@{ Count = 0; Profiles = [System.Collections.Generic.List[string]]::new(); HasConflict = $false }
        }
        $c = $cfgByName[$k]
        $c.Count += 1
        if ($c.Profiles.Count -lt 6 -and $e.ProfileName) { $c.Profiles.Add([string]$e.ProfileName) }
        if ($e.Status -eq 'conflict') { $c.HasConflict = $true }
    }

    # Mises à jour qualité en échec : colonnes du rapport cherchées sous plusieurs graphies
    $updByName = @{}
    foreach ($u in @($Health.UpdateFailures)) {
        if ($null -eq $u) { continue }
        $state = [string](Get-PropValue $u @('AggregateState', 'CurrentDeviceUpdateStatus', 'UpdateStatus', 'status'))
        if ($state -notmatch '(?i)error|fail|cancel') { continue }
        $k = Get-DeviceNameKey ([string](Get-PropValue $u @('DeviceName', 'Device', 'deviceDisplayName')))
        if (-not $k -or $updByName.ContainsKey($k)) { continue }
        $updByName[$k] = [pscustomobject]@{ State = $state; Message = [string](Get-PropValue $u @('LatestAlertMessage', 'ErrorCode')) }
    }

    # Un même nom de poste peut correspondre à plusieurs enregistrements Intune (réimagé,
    # ré-enrôlé...). On garde l'enregistrement le plus récemment synchronisé : garder "le
    # premier rencontré" produisait de faux positifs "inactif depuis N jours".
    $byName   = [ordered]@{}
    $dupCount = 0
    foreach ($d in @($Devices)) {
        if ($null -eq $d) { continue }
        $k = Get-DeviceNameKey $d.DeviceName
        if (-not $k) { continue }
        if (-not $byName.Contains($k)) { $byName[$k] = $d; continue }
        $dupCount++
        $kept = $byName[$k]
        if ($d.LastSyncDateTime -and (-not $kept.LastSyncDateTime -or $d.LastSyncDateTime -gt $kept.LastSyncDateTime)) { $byName[$k] = $d }
    }
    if ($dupCount -gt 0) {
        Write-Log "Santé des postes : $dupCount enregistrement(s) en doublon de nom de poste ; le plus récemment synchronisé a été conservé (doublons à nettoyer dans Intune)." -Level WARN
    }

    $nowUtc  = [datetime]::UtcNow
    $rows    = [System.Collections.Generic.List[object]]::new()
    $actions = [ordered]@{}

    foreach ($k in @($byName.Keys)) {
        $d      = $byName[$k]
        $sc     = $scoreByName[$k]
        $pf     = $perfByName[$k]
        $bl     = $blByName[$k]
        $bh     = $battByName[$k]
        $dv     = if ($Health.Defender) { $Health.Defender[[string]$d.Id] } else { $null }
        $crash  = $crashByName[$k]
        $cfg    = $cfgByName[$k]
        $upd    = $updByName[$k]
        $issues = [System.Collections.Generic.List[object]]::new()

        # ----- Espace disque -----
        $freeGB = $null; $totGB = $null; $freePct = $null; $diskSev = $null
        if ($null -ne $d.FreeStorageSpaceInBytes) { $freeGB = [math]::Round($d.FreeStorageSpaceInBytes / 1GB, 1) }
        if ($d.TotalStorageSpaceInBytes -gt 0) {
            $totGB = [math]::Round($d.TotalStorageSpaceInBytes / 1GB, 1)
            if ($null -ne $d.FreeStorageSpaceInBytes) { $freePct = [int][math]::Round(100.0 * $d.FreeStorageSpaceInBytes / $d.TotalStorageSpaceInBytes) }
        }
        if ($Checks.Disk -and $null -ne $freePct) {
            $txt = "$(Format-FrNumber $freeGB) Go libres sur $(Format-FrNumber $totGB) Go ($freePct %)"
            if ($freePct -lt $T.DiskFreePctCritical -or ($null -ne $freeGB -and $freeGB -lt $T.DiskFreeGbCritical)) {
                $diskSev = 'crit'
                $issues.Add((New-HealthIssue "Libérer de l'espace disque (moins de $($T.DiskFreePctCritical) % ou $($T.DiskFreeGbCritical) Go libres)" 'crit' $txt))
            } elseif ($freePct -lt $T.DiskFreePctWarning) {
                $diskSev = 'warn'
                $issues.Add((New-HealthIssue "Espace disque à surveiller (moins de $($T.DiskFreePctWarning) % libres)" 'warn' $txt))
            }
        }

        # ----- Inactivité -----
        $daysSince = $null
        if ($d.LastSyncDateTime) { $daysSince = [int][math]::Floor(($nowUtc - $d.LastSyncDateTime).TotalDays) }
        if ($Checks.Inactivity -and $null -ne $daysSince) {
            if ($daysSince -ge $T.StaleDaysCritical) {
                $issues.Add((New-HealthIssue "Appareil inactif depuis plus de $($T.StaleDaysCritical) jours - reprendre contact ou retirer d'Intune" 'crit' "$daysSince jours sans synchronisation"))
            } elseif ($daysSince -ge $T.StaleDaysWarning) {
                $issues.Add((New-HealthIssue "Appareil sans synchronisation depuis plus de $($T.StaleDaysWarning) jours" 'warn' "$daysSince jours sans synchronisation"))
            }
        }

        # ----- Scores Endpoint Analytics & démarrage -----
        $analytics = ConvertTo-NullableDouble (Get-PropValue $sc @('endpointAnalyticsScore')) -NonNegative
        $battScore = ConvertTo-NullableDouble (Get-PropValue $sc @('batteryHealthScore')) -NonNegative
        $bootSec = $null; $bootScore = $null; $bsod = $null; $restarts = $null; $diskType = ""
        if ($pf) {
            $ms = ConvertTo-NullableDouble $pf.coreBootTimeInMs -NonNegative
            if ($ms -gt 0) { $bootSec = [int][math]::Round($ms / 1000) }
            $bootScore = ConvertTo-NullableDouble $pf.bootScore -NonNegative
            $bsod      = ConvertTo-NullableLong $pf.blueScreenCount
            $restarts  = ConvertTo-NullableLong $pf.restartCount
            $diskType  = "$($pf.diskType)".Trim()
        }
        if ($Checks.Bsod) {
            if ($bsod -gt 0) {
                $sev = if ($bsod -ge $T.BsodCritical) { 'crit' } else { 'warn' }
                $det = "$bsod écran(s) bleu(s) sur 14 jours"
                if ($null -ne $restarts) { $det += ", $restarts redémarrage(s)" }
                $issues.Add((New-HealthIssue "Écrans bleus (BSOD) récents - analyser pilotes et matériel" $sev $det))
            } elseif ($restarts -ge $T.RestartsHigh) {
                $issues.Add((New-HealthIssue "Redémarrages anormalement fréquents" 'warn' "$restarts redémarrage(s) sur 14 jours"))
            }
        }
        if ($Checks.Boot -and $bootSec -ge $T.BootSlowSeconds) {
            $det = "Démarrage en $bootSec s"
            if ($null -ne $bootScore) { $det += " (score $([int]$bootScore))" }
            $issues.Add((New-HealthIssue "Démarrage lent (plus de $($T.BootSlowSeconds) s) - alléger le démarrage" 'warn' $det))
        }
        if ($Checks.Boot -and $diskType -match '(?i)hdd') {
            $det = "Disque mécanique (HDD)"
            if ($null -ne $bootSec) { $det += " - démarrage $bootSec s" }
            $issues.Add((New-HealthIssue "Disque mécanique (HDD) - envisager un remplacement SSD" 'warn' $det))
        }
        if ($Checks.Battery -and $null -ne $battScore -and $battScore -lt $T.BatteryPoor) {
            $issues.Add((New-HealthIssue "Batterie dégradée (score sous $($T.BatteryPoor)) - prévoir un remplacement" 'warn' "Score batterie $([int]$battScore)"))
        }
        if ($Checks.EaScore -and $null -ne $analytics -and $analytics -lt $T.ScoreLow) {
            $issues.Add((New-HealthIssue "Score Endpoint Analytics faible (sous $($T.ScoreLow))" 'warn' "Score global $([int]$analytics)"))
        }

        # ----- BitLocker -----
        # advancedBitLockerStates est un enum "flags" sérialisé en texte ("tpmNotReady,
        # loggedOnUserNonAdmin"). Tous les indicateurs ne sont pas des défauts :
        # loggedOnUserNonAdmin (utilisateur non administrateur) est une BONNE pratique ;
        # le compter comme une erreur produisait une majorité de faux positifs.
        if ($Checks.BitLocker -and $bl) {
            $encState      = "$($bl.encryptionState)"
            $flags         = @("$($bl.advancedBitLockerStates)" -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne 'success' })
            $criticalFlags = @('osVolumeUnprotected', 'recoveryKeyBackupFailed', 'fixedDriveNotEncrypted', 'tpmNotAvailable', 'tpmNotReady')
            $ignoredFlags  = @('noUserConsent', 'loggedOnUserNonAdmin')
            $relevant      = @($flags | Where-Object { $_ -notin $ignoredFlags })
            if ($encState -eq 'notEncrypted') {
                $issues.Add((New-HealthIssue "BitLocker désactivé - chiffrer le poste" 'crit' "Volume système non chiffré"))
            } elseif ($relevant.Count -gt 0) {
                if (@($relevant | Where-Object { $_ -in $criticalFlags }).Count -gt 0) {
                    $issues.Add((New-HealthIssue "BitLocker en échec réel - protection du volume compromise" 'crit' ($relevant -join ', ')))
                } else {
                    $issues.Add((New-HealthIssue "BitLocker : écart de stratégie à vérifier (non bloquant)" 'warn' ($relevant -join ', ')))
                }
            }
        }

        # ----- Defender -----
        if ($Checks.Defender -and $dv) {
            if ($dv.realTimeProtectionEnabled -eq $false) {
                $issues.Add((New-HealthIssue "Protection en temps réel Defender désactivée" 'crit' "realTimeProtectionEnabled = false"))
            }
            if ($dv.signatureUpdateOverdue -eq $true) {
                $sigDet = "Signatures en retard"
                $lastReport = ConvertTo-UtcDate $dv.lastReportedDateTime
                if ($lastReport) { $sigDet += " (dernier rapport : $(Format-LocalDate $lastReport))" }
                $issues.Add((New-HealthIssue "Signatures antivirus en retard (plus de $($T.SignatureStaleDays) jours)" 'warn' $sigDet))
            }
        }

        # ----- Fiabilité applicative -----
        if ($Checks.AppReliability -and $crash -and $crash.Count -ge $T.AppCrashWarning) {
            $issues.Add((New-HealthIssue "Application qui plante fréquemment - réinstaller ou mettre à jour" 'warn' "$($crash.AppName) : $($crash.Count) plantage(s)"))
        }

        # ----- Uptime (estimation via le dernier démarrage connu) -----
        $lastBoot = $bootById["$($d.Id)".ToLowerInvariant()]
        if (-not $lastBoot) { $lastBoot = $bootByName[$k] }
        $uptimeDays = if ($lastBoot) { [int][math]::Floor(($nowUtc - $lastBoot).TotalDays) } else { $null }
        if ($Checks.Uptime -and $null -ne $uptimeDays -and $uptimeDays -ge $T.UptimeWarningDays) {
            $issues.Add((New-HealthIssue "Poste non redémarré depuis longtemps (estimation) - planifier un redémarrage" 'warn' "~$uptimeDays jour(s) depuis le dernier démarrage détecté"))
        }

        # ----- Conformité Intune (motif repris de l'analyse des non-conformités) -----
        $compState = "$($d.ComplianceState)".Trim()
        if ($Checks.Compliance -and $compState) {
            $reason = $ComplianceReasonsById[[string]$d.Id]
            switch -Regex ($compState) {
                '(?i)^noncompliant$' {
                    $det = if ($reason) { $reason } else { "État Intune : non conforme (motif non analysé dans cette génération)" }
                    $issues.Add((New-HealthIssue "Non conforme aux stratégies de conformité Intune" 'crit' $det))
                }
                '(?i)^inGracePeriod$' {
                    $issues.Add((New-HealthIssue "En période de grâce - à traiter avant bascule en non conforme" 'warn' "État Intune : période de grâce avant bascule en non conforme"))
                }
                '(?i)^(error|conflict)$' {
                    $issues.Add((New-HealthIssue "Évaluation de conformité en erreur ou en conflit" 'warn' "État Intune : $compState"))
                }
            }
        }

        # ----- Profils de configuration en erreur / conflit -----
        if ($Checks.ConfigProfile -and $cfg -and $cfg.Count -gt 0) {
            $names = @($cfg.Profiles | Select-Object -First 3)
            $det   = "$($cfg.Count) profil(s) en échec"
            if ($names.Count -gt 0) { $det += " : " + ($names -join ' ; ') }
            if ($cfg.Count -gt $names.Count) { $det += " ..." }
            if ($cfg.Count -ge $T.ConfigErrorsCritical) {
                $issues.Add((New-HealthIssue "Profils de configuration en échec ($($T.ConfigErrorsCritical) ou plus) - paramétrage non appliqué" 'crit' $det))
            } elseif ($cfg.HasConflict) {
                $issues.Add((New-HealthIssue "Conflit entre profils de configuration - deux stratégies se contredisent" 'warn' $det))
            } else {
                $issues.Add((New-HealthIssue "Profil de configuration en erreur - paramétrage non appliqué" 'warn' $det))
            }
        }

        # ----- Batterie : capacité, âge, autonomie (valeurs négatives = non mesuré) -----
        $battCapacity = $null; $battAgeDays = $null; $battRuntime = $null
        if ($bh) {
            $battCapacity = ConvertTo-NullableDouble (Get-PropValue $bh @('maxCapacityPercentage', 'maxCapacityPercent')) -NonNegative
            $battAgeDays  = ConvertTo-NullableDouble (Get-PropValue $bh @('batteryAgeInDays')) -NonNegative
            $battRuntime  = ConvertTo-NullableDouble (Get-PropValue $bh @('estimatedRuntimeInMinutes')) -NonNegative
            if ($battCapacity -eq 0) { $battCapacity = $null }
            if ($battRuntime -eq 0)  { $battRuntime = $null }
        }
        if ($Checks.BatteryDetail -and $null -ne $battCapacity -and $battCapacity -lt $T.BatteryCapacityPoor) {
            $det = "Capacité maximale restante : $([int]$battCapacity) %"
            if ($null -ne $battAgeDays) { $det += " - batterie âgée de ~$([int][math]::Round($battAgeDays / 365.0, 0)) an(s)" }
            if ($null -ne $battRuntime) { $det += " - autonomie estimée $([int]$battRuntime) min" }
            if ($battCapacity -lt $T.BatteryCapacityCrit) {
                $issues.Add((New-HealthIssue "Batterie hors d'usage (capacité sous $($T.BatteryCapacityCrit) %) - remplacement à planifier" 'crit' $det))
            } else {
                $issues.Add((New-HealthIssue "Batterie usée (capacité sous $($T.BatteryCapacityPoor) %) - remplacement à prévoir" 'warn' $det))
            }
        } elseif ($Checks.BatteryDetail -and $null -ne $battRuntime -and $battRuntime -lt $T.BatteryRuntimeLowMin -and $null -ne $battCapacity) {
            $issues.Add((New-HealthIssue "Autonomie insuffisante (moins de $($T.BatteryRuntimeLowMin) min) - poste devenu sédentaire" 'warn' "Autonomie estimée $([int]$battRuntime) min pour $([int]$battCapacity) % de capacité"))
        }

        # ----- Mises à jour qualité Windows en échec -----
        if ($Checks.WindowsUpdate -and $upd) {
            $det = "État : $($upd.State)"
            if ($upd.Message) { $det += " - $($upd.Message)" }
            $issues.Add((New-HealthIssue "Mise à jour qualité Windows en échec - poste privé de correctifs de sécurité" 'crit' $det))
        }

        # ----- Sévérité globale du poste -----
        $sevMax = 0
        foreach ($i in $issues) { if ($rank[$i.Sev] -gt $sevMax) { $sevMax = $rank[$i.Sev] } }
        $severity = @('ok', 'warn', 'crit')[$sevMax]

        # ----- Anonymisation : mêmes alias que dans toutes les autres pages -----
        $name = $d.DeviceName
        $upn  = $d.UserPrincipalName
        if ($AnonymizeData) {
            $anon = Get-AnonymizedIdentity -RealName $name -RealUpn $upn
            $name = $anon.Name
            $upn  = $anon.Upn
        }

        foreach ($i in $issues) {
            if (-not $actions.Contains($i.Label)) {
                $actions[$i.Label] = [pscustomobject]@{ Severity = $i.Sev; Devices = [System.Collections.Generic.List[object]]::new() }
            }
            $a = $actions[$i.Label]
            if ($rank[$i.Sev] -gt $rank[$a.Severity]) { $a.Severity = $i.Sev }
            $a.Devices.Add([pscustomobject]@{
                DeviceName        = $name
                UserPrincipalName = $upn
                OperatingSystem   = $d.OperatingSystem
                Detail            = $i.Detail
                LastSyncDateTime  = $d.LastSyncDateTime
            })
        }

        $rows.Add([pscustomobject]@{
            Severity          = $severity
            DeviceName        = $name
            UserPrincipalName = $upn
            OperatingSystem   = $d.OperatingSystem
            OSVersion         = $d.OSVersion
            FreeGB            = $freeGB
            TotalGB           = $totGB
            FreePct           = $freePct
            DiskSeverity      = $diskSev
            AnalyticsScore    = $analytics
            BootSeconds       = $bootSec
            DiskType          = $diskType
            BlueScreens       = $bsod
            Restarts          = $restarts
            BatteryScore      = $battScore
            BatteryCapacity   = $battCapacity
            BatteryAgeDays    = $battAgeDays
            BatteryRuntimeMin = $battRuntime
            BitLockerState    = if ($bl) { "$($bl.encryptionState)" } else { $null }
            DefenderRealTime  = if ($dv) { $dv.realTimeProtectionEnabled } else { $null }
            DefenderSigStale  = if ($dv) { $dv.signatureUpdateOverdue } else { $null }
            CrashApp          = if ($crash) { $crash.AppName } else { $null }
            CrashCount        = if ($crash) { $crash.Count } else { $null }
            UptimeDays        = $uptimeDays
            ComplianceState   = $compState
            ConfigErrorCount  = if ($cfg) { $cfg.Count } else { $null }
            ConfigErrorNames  = if ($cfg) { (@($cfg.Profiles) -join ' ; ') } else { $null }
            UpdateFailure     = if ($upd) { $upd.State } else { $null }
            LastSyncDateTime  = $d.LastSyncDateTime
            DaysSinceSync     = $daysSince
            Issues            = (@($issues | ForEach-Object { $_.Label }) -join ' | ')
            DeviceId          = [string]$d.Id
        })
    }

    $sorted   = @($rows | Sort-Object @{ Expression = { $rank[$_.Severity] }; Descending = $true }, @{ Expression = { $_.DeviceName } })
    $crit     = @($sorted | Where-Object { $_.Severity -eq 'crit' }).Count
    $warn     = @($sorted | Where-Object { $_.Severity -eq 'warn' }).Count
    $avgScore = $null
    $scored   = @($sorted | Where-Object { $null -ne $_.AnalyticsScore })
    if ($scored.Count -gt 0) { $avgScore = [int][math]::Round(($scored | Measure-Object -Property AnalyticsScore -Average).Average) }

    return [pscustomobject]@{
        Rows           = $sorted
        Actions        = $actions
        CritCount      = $crit
        WarnCount      = $warn
        OkCount        = $sorted.Count - $crit - $warn
        AvgScore       = $avgScore
        DuplicateNames = $dupCount
    }
}

# Actions triées par sévérité puis nombre de postes
function Get-SortedHealthActions {
    param($Actions)
    if (-not $Actions -or $Actions.Count -eq 0) { return @() }
    return @($Actions.GetEnumerator() | Sort-Object @{ Expression = { $HealthSeverityRank[$_.Value.Severity] }; Descending = $true },
                                                    @{ Expression = { $_.Value.Devices.Count }; Descending = $true },
                                                    @{ Expression = { $_.Key } })
}

# Synthèse des actions (tableau, graphique et CSV)
function Get-HealthActionSummary {
    param($Actions)
    foreach ($entry in (Get-SortedHealthActions -Actions $Actions)) {
        [pscustomobject][ordered]@{
            'Sévérité'           = $HealthSeverityLabels[$entry.Value.Severity]
            'Action recommandée' = $entry.Key
            'Postes concernés'   = $entry.Value.Devices.Count
        }
    }
}

# Détail action x poste (tableau et CSV)
function Select-HealthActionDetailView {
    param($Actions)
    foreach ($entry in (Get-SortedHealthActions -Actions $Actions)) {
        foreach ($dev in $entry.Value.Devices) {
            [pscustomobject][ordered]@{
                'Sévérité'           = $HealthSeverityLabels[$entry.Value.Severity]
                'Action recommandée' = $entry.Key
                'Poste'              = $dev.DeviceName
                'Utilisateur (UPN)'  = $dev.UserPrincipalName
                'OS'                 = $dev.OperatingSystem
                'Détail'             = $dev.Detail
                'Dernière synchro'   = Format-LocalDate $dev.LastSyncDateTime
            }
        }
    }
}

# Vue "un poste, tous ses indicateurs" (tableau et CSV)
function Select-HealthDeviceView {
    param([object[]]$Rows)
    $yesNo = { param($v) if ($null -eq $v) { "" } elseif ($v -eq $true) { "Oui" } else { "Non" } }
    foreach ($r in $Rows) {
        [pscustomobject][ordered]@{
            'Sévérité'               = $HealthSeverityLabels[$r.Severity]
            'Poste'                  = $r.DeviceName
            'Utilisateur (UPN)'      = $r.UserPrincipalName
            'OS'                     = $r.OperatingSystem
            'Version OS'             = $r.OSVersion
            'Disque libre (Go)'      = $r.FreeGB
            'Disque total (Go)'      = $r.TotalGB
            'Disque libre (%)'       = $r.FreePct
            'Score Endpoint Analytics' = $r.AnalyticsScore
            'Démarrage (s)'          = $r.BootSeconds
            'Type de disque'         = $r.DiskType
            'Écrans bleus (14 j)'    = $r.BlueScreens
            'Redémarrages (14 j)'    = $r.Restarts
            'Score batterie'         = $r.BatteryScore
            'Capacité batterie (%)'  = $r.BatteryCapacity
            'Autonomie (min)'        = $r.BatteryRuntimeMin
            'BitLocker'              = $r.BitLockerState
            'Defender temps réel'    = & $yesNo $r.DefenderRealTime
            'Signatures en retard'   = & $yesNo $r.DefenderSigStale
            'Application instable'   = $r.CrashApp
            'Plantages'              = $r.CrashCount
            'Uptime estimé (j)'      = $r.UptimeDays
            'Conformité'             = ConvertTo-NcStateLabel $r.ComplianceState
            'Profils en échec'       = $r.ConfigErrorCount
            'Mise à jour en échec'   = $r.UpdateFailure
            'Dernière synchro'       = Format-LocalDate $r.LastSyncDateTime
            'Jours sans synchro'     = $r.DaysSinceSync
            'Actions recommandées'   = $r.Issues
        }
    }
}

# Postes à l'espace disque insuffisant (remplace la liste "Low Storage < 100 GB" de la v3)
function Select-HealthDiskView {
    param([object[]]$Rows)
    foreach ($r in @($Rows | Where-Object { $_.DiskSeverity } | Sort-Object FreePct)) {
        [pscustomobject][ordered]@{
            'Sévérité'          = $HealthSeverityLabels[$r.DiskSeverity]
            'Poste'             = $r.DeviceName
            'Utilisateur (UPN)' = $r.UserPrincipalName
            'OS'                = $r.OperatingSystem
            'Disque libre (Go)' = $r.FreeGB
            'Disque total (Go)' = $r.TotalGB
            'Disque libre (%)'  = $r.FreePct
            'Type de disque'    = $r.DiskType
            'Dernière synchro'  = Format-LocalDate $r.LastSyncDateTime
        }
    }
}

# Bloc d'en-tête de la page "Santé & proactivité" : KPI, lecture, avertissements
function New-HealthDashboardHeadHtml {
    param([object]$Data, [System.Collections.Generic.List[string]]$Warnings, [int]$ActiveChecks = 0)

    $esc = { param($Text) ((ConvertTo-HtmlSafe $Text) -replace '\[', '&#91;') -replace '\]', '&#93;' }
    $warnHtml = ""
    if ($Warnings -and $Warnings.Count -gt 0) {
        $items    = ($Warnings | ForEach-Object { "<li>$(& $esc $_)</li>" }) -join ''
        $warnHtml = "<div class=`"ix-root ix-note ix-note--warn`">$(Get-IconSvg -Name 'alert' -Size 18)<div><strong>Données partiellement disponibles</strong><ul>$items</ul></div></div>"
    }
    if (-not $Data -or @($Data.Rows).Count -eq 0) {
        return (New-IxEmptyState -Icon 'gauge' -Tone 'neutral' -Title "Aucun poste à analyser" -Text "Aucun appareil du périmètre n'a pu être évalué.") + $warnHtml
    }

    $total = @($Data.Rows).Count
    $avg   = if ($null -ne $Data.AvgScore) { "$($Data.AvgScore)" } else { "&mdash;" }
    $cards = [System.Text.StringBuilder]::new()
    [void]$cards.Append((New-IxStatCard -Label 'Postes analysés' -Value $total -Icon 'monitor' -Color $Colors.Secondary -Caption "$ActiveChecks vérification(s) active(s)"))
    [void]$cards.Append((New-IxStatCard -Label 'Critiques' -Value $Data.CritCount -Icon 'x-circle' -Color $Colors.Danger -Caption "<b>$(Get-IxPercent $Data.CritCount $total) %</b> du parc - à traiter en priorité"))
    [void]$cards.Append((New-IxStatCard -Label 'À surveiller' -Value $Data.WarnCount -Icon 'alert' -Color $Colors.Warning -Caption "<b>$(Get-IxPercent $Data.WarnCount $total) %</b> du parc"))
    [void]$cards.Append((New-IxStatCard -Label 'Sans alerte' -Value $Data.OkCount -Icon 'shield-ok' -Color $Colors.Success -Caption "Tous les indicateurs dans les seuils"))
    [void]$cards.Append((New-IxStatCard -Label 'Score Endpoint Analytics moyen' -Value $avg -Icon 'gauge' -Color $Colors.Accent -Caption "Sur 100, postes mesurés uniquement"))

    $note = "<div class=`"ix-root ix-note`">$(Get-IconSvg -Name 'info' -Size 18)<div><strong>Lecture :</strong> " +
            "chaque poste reçoit la sévérité de sa pire alerte ; les actions recommandées regroupent les postes concernés. " +
            "Les données Endpoint Analytics (scores, démarrages, batteries, plantages) sont agrégées quotidiennement par Microsoft " +
            "et nécessitent que l'Analyse des points de terminaison soit activée ; l'uptime est une estimation. " +
            "Seuils modifiables en tête de script (`$RemediationThresholds).</div></div>"

    return "<div class=`"ix-root ix-overview`">$($cards.ToString())</div>" + $note + $warnHtml
}

# ========================================
# [v4.0] EXPORTS CSV (TOUTES LES SECTIONS)
# ========================================
# Un dossier horodaté par génération, un fichier par jeu de données collecté.
# Séparateur ";" et UTF-8 avec BOM (ouverture directe dans Excel en français) ;
# dates en heure locale "JJ/MM/AAAA HH:MM", décimaux selon la culture du poste,
# booléens Oui / Non, listes jointes par " | ". Données anonymisées si l'option
# est cochée (comme le rapport HTML).

function Export-CsvFile {
    param([object[]]$Rows, [string]$Path)
    $culture = [Globalization.CultureInfo]::CurrentCulture
    $out = foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $o = [ordered]@{}
        foreach ($p in $r.PSObject.Properties) {
            $v = $p.Value
            if ($v -is [datetime]) { $v = Format-LocalDate $v }
            elseif ($v -is [double] -or $v -is [single] -or $v -is [decimal]) { $v = $v.ToString($culture) }
            elseif ($v -is [bool]) { $v = if ($v) { "Oui" } else { "Non" } }
            elseif ($null -ne $v -and $v -isnot [string] -and $v -is [System.Collections.IEnumerable]) { $v = (@($v) | ForEach-Object { "$_" }) -join ' | ' }
            $o[$p.Name] = $v
        }
        [pscustomobject]$o
    }
    $lines = [string[]]@($out | ConvertTo-Csv -NoTypeInformation -Delimiter $NcCsvDelimiter)
    [System.IO.File]::WriteAllLines($Path, $lines, [System.Text.UTF8Encoding]::new($true))
}

# Écrit chaque jeu de données ; un échec (fichier ouvert dans Excel, droits...) devient
# un avertissement sans bloquer les autres. Jeu vide : pas de fichier.
function Export-AllDatasets {
    param([System.Collections.Specialized.OrderedDictionary]$Datasets, [string]$Folder, [System.Collections.Generic.List[string]]$Warnings)
    $written = [System.Collections.Generic.List[string]]::new()
    try {
        if (-not (Test-Path -LiteralPath $Folder)) { New-Item -ItemType Directory -Path $Folder -Force -ErrorAction Stop | Out-Null }
    } catch {
        $Warnings.Add("Export CSV impossible : le dossier « $Folder » n'a pas pu être créé ($($_.Exception.Message)).")
        return @()
    }
    foreach ($name in @($Datasets.Keys)) {
        $rows = @($Datasets[$name] | Where-Object { $null -ne $_ })
        if ($rows.Count -eq 0) { Write-Log "CSV « $name » : aucune donnée, fichier non créé."; continue }
        $path = Join-Path $Folder ("{0}.csv" -f ($name -replace '[\\/:*?"<>|]', '_'))
        try {
            Export-CsvFile -Rows $rows -Path $path
            $written.Add($path)
        } catch {
            $Warnings.Add("Export CSV « $name » en échec : $($_.Exception.Message)")
        }
    }
    Write-Log "Export CSV : $($written.Count) fichier(s) dans $Folder" -Level OK
    return $written.ToArray()
}

# ========================================
# INSTALLATION DES MODULES
# ========================================
# [v4.0] Seul PSWriteHTML reste nécessaire, et uniquement si le rapport HTML est
# demandé : toute la collecte passe par l'API REST (le module Microsoft.Graph, dont
# l'import complet est lent et peut dépasser la limite de 4 096 fonctions de
# Windows PowerShell 5.1, n'est plus utilisé).

function Install-RequiredModules {
    param([string[]]$Modules = @("PSWriteHTML"))
    Write-Step "🔍 Vérification des modules PowerShell..."
    $missingModules = @($Modules | Where-Object { -not (Get-Module -Name $_ -ListAvailable) })
    if ($missingModules.Count -eq 0) { return $true }

    $lblStatus.Text = "⚠ Modules manquants détectés"; $lblStatus.ForeColor = [System.Drawing.Color]::Red; $form.Refresh()
    $moduleList = $missingModules -join ", "
    $message = "Modules PowerShell manquants : $moduleList`n`nCes modules sont nécessaires pour générer le rapport HTML.`n`nVoulez-vous les installer maintenant ?`n`nNote : L'installation peut prendre quelques minutes.`n(Sans HTML, décochez « Générer le rapport HTML » : aucun module n'est requis.)"
    $result = [System.Windows.Forms.MessageBox]::Show($message, "Installation des modules requis", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $lblStatus.Text = "📦 Installation des modules en cours..."; $lblStatus.ForeColor = [System.Drawing.Color]::Orange; $form.Refresh()
        try {
            foreach ($module in $missingModules) {
                $lblStatus.Text = "📦 Installation de $module... (patientez)"; $form.Refresh()
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                $lblStatus.Text = "✓ $module installé avec succès"; $lblStatus.ForeColor = [System.Drawing.Color]::Green; $form.Refresh()
                Write-Log "Module $module installé." -Level OK
            }
            $lblStatus.Text = "✓ Modules installés"; $lblStatus.ForeColor = [System.Drawing.Color]::Green
            return $true
        } catch {
            $lblStatus.Text = "❌ Erreur lors de l'installation"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
            Write-Log "Installation de module impossible : $($_.Exception.Message)" -Level ERROR
            Show-ErrorMessage "Erreur lors de l'installation :`n`n$($_.Exception.Message)`n`nInstallez manuellement :`nInstall-Module -Name $moduleList -Force -AllowClobber -Scope CurrentUser"
            return $false
        }
    } else {
        $lblStatus.Text = "❌ Installation annulée"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
        Show-InfoMessage "Installation annulée.`n`nPour installer manuellement :`nInstall-Module -Name $moduleList -Force -AllowClobber -Scope CurrentUser"
        return $false
    }
}

# [v4.0] Texte des avertissements de la génération (message de fin et journal)
function Get-RunWarningsText {
    param([object]$NcResult, [System.Collections.Generic.List[string]]$RunWarnings, [System.Collections.Generic.List[string]]$HealthWarnings)
    $text = ""
    if ($NcResult -and -not $NcResult.Success) {
        $text += "`n`n⚠ Analyse des non-conformités indisponible :`n$($NcResult.Error)"
    } elseif ($NcResult -and $NcResult.Warnings.Count -gt 0) {
        $text += "`n`n⚠ Analyse des non-conformités ($($NcResult.Warnings.Count) avertissement(s)) :`n- " + ($NcResult.Warnings -join "`n- ")
    }
    if ($HealthWarnings -and $HealthWarnings.Count -gt 0) {
        $text += "`n`n⚠ Santé des postes ($($HealthWarnings.Count) source(s) partiellement disponible(s)) :`n- " + ($HealthWarnings -join "`n- ")
    }
    if ($RunWarnings -and $RunWarnings.Count -gt 0) {
        $text += "`n`n⚠ Autres avertissements :`n- " + ($RunWarnings -join "`n- ")
    }
    return $text
}

# ========================================
# GÉNÉRATION DU DASHBOARD
# ========================================
# [v4.0] Déroulé en phases : options -> prérequis -> authentification -> collectes et
# analyses (une section décochée n'est pas collectée) -> anonymisation -> exports CSV
# -> rendu HTML (facultatif) -> bilan. Détail de chaque étape dans $LogFile.

function Generate-Dashboard {
    param([bool]$OpenAfterGeneration = $true)

    $selectedClient = $cmbClients.Text.Trim()
    if ($selectedClient -eq "-- Sélectionnez un client --" -or [string]::IsNullOrWhiteSpace($selectedClient)) {
        Show-ErrorMessage "Veuillez sélectionner un client !"; return
    }

    $config = Get-ClientConfig -ClientName $selectedClient
    if ($null -eq $config) {
        Show-ErrorMessage "Impossible de charger la configuration du client !`n`nAssurez-vous que le fichier a été généré avec le même script (même clé AES)."
        return
    }

    $ClientName   = $config.ClientName
    $TenantId     = $config.TenantId
    $ClientId     = $config.ClientId
    $ClientSecret = $config.ClientSecret

    # ===== [v4.0] OPTIONS DE SORTIE ET DE COLLECTE =====
    # HTML décoché : requêtes Graph + exports CSV uniquement (PSWriteHTML non requis).
    $GenerateHtml = $chkGenerateHtml.Checked
    $ExportCsv    = $chkExportCsv.Checked -or -not $GenerateHtml
    $HealthChecks = @{}
    foreach ($def in $HealthCheckDefs) { $HealthChecks[$def.Key] = ($chkHealth.Checked -and $script:ChkHealthChecks[$def.Key].Checked) }
    $ActiveHealthChecks = @($HealthChecks.Values | Where-Object { $_ }).Count
    $RunHealth          = $ActiveHealthChecks -gt 0
    $RunWarnings        = [System.Collections.Generic.List[string]]::new()
    $HealthWarnings     = [System.Collections.Generic.List[string]]::new()
    Reset-AnonymizationMaps

    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    Write-Step "⏳ Génération pour $ClientName..."
    Write-Log "================================================================"
    Write-Log "NOUVELLE GÉNÉRATION - Client : $ClientName - HTML=$GenerateHtml / CSV=$ExportCsv / Santé=$ActiveHealthChecks vérification(s) / Parallèle=$($chkParallel.Checked)"

    try {
        if ($GenerateHtml) {
            if (-not (Install-RequiredModules -Modules @("PSWriteHTML"))) { return }
            Write-Step "📚 Chargement du module PSWriteHTML..."
            try {
                Import-Module PSWriteHTML -ErrorAction Stop
            } catch {
                $lblStatus.Text = "❌ Erreur de chargement des modules"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
                Show-ErrorMessage "Erreur lors du chargement du module PSWriteHTML :`n`n$($_.Exception.Message)"; return
            }
        }

        # ===== CONNEXION GRAPH =====
        Write-Step "🔐 Connexion à Microsoft Graph..."
        # [v3.3] Jeton géré par Get-GraphAccessToken : erreur d'authentification explicite
        # (secret expiré, tenant inconnu, réseau) et renouvellement automatique avant
        # expiration pendant les analyses longues. Cache par appareil remis à zéro.
        # [v4.0] Contexte synchronisé partagé avec les flux de collecte parallèle.
        $script:GraphContext    = New-GraphContext -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
        $script:NcSettingsCache = @{}
        try {
            $AccessToken = Get-GraphAccessToken -ForceRefresh
        } catch {
            $lblStatus.Text = "❌ Échec de l'authentification Microsoft Graph"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
            Write-Log "Authentification impossible : $($_.Exception.Message)" -Level ERROR
            Show-ErrorMessage "Impossible d'obtenir un jeton d'accès Microsoft Graph pour $ClientName :`n`n$($_.Exception.Message)"
            return
        }
        Write-Log "Jeton d'accès Graph obtenu." -Level OK

        # ===== RÉCUPÉRATION DES APPAREILS =====
        Write-Step "💻 Récupération des appareils gérés..."
        # [v4.0] REST + $select (remplace Get-MgDeviceManagementManagedDevice -All)
        $ManagedDevices = @(Get-ManagedDevicesRest -AccessToken $AccessToken)
        $ExcludeVM      = $chkExcludeVM.Checked
        $AnonymizeData  = $chkAnonymize.Checked

        $DevicesScope = if ($ExcludeVM) {
            @($ManagedDevices | Where-Object {
                -not (
                    ($_.Manufacturer -match 'VMware|innotek|VirtualBox|Parallels|QEMU') -or
                    ($_.Model        -match 'Virtual|VMware|VirtualBox|Hyper-V|Parallels|QEMU|KVM')
                )
            })
        } else { $ManagedDevices }
        Write-Log "Périmètre d'analyse : $($DevicesScope.Count) appareil(s) sur $($ManagedDevices.Count) (exclusion des VM : $ExcludeVM)."

        # ===== LECTURE DES OPTIONS D'AFFICHAGE =====
        # Plateformes à afficher
        $ShowWindows = $chkShowWindows.Checked
        $ShowIOS     = $chkShowIOS.Checked
        $ShowAndroid = $chkShowAndroid.Checked
        $ShowMac     = $chkShowMac.Checked

        # Paliers inactifs sélectionnés
        $InactiveThresholds = [System.Collections.ArrayList]@()
        if ($chkInactive30.Checked)  { [void]$InactiveThresholds.Add(30)  }
        if ($chkInactive60.Checked)  { [void]$InactiveThresholds.Add(60)  }
        if ($chkInactive90.Checked)  { [void]$InactiveThresholds.Add(90)  }
        if ($chkInactive120.Checked) { [void]$InactiveThresholds.Add(120) }
        if ($chkInactive150.Checked) { [void]$InactiveThresholds.Add(150) }
        if ($chkInactive180.Checked) { [void]$InactiveThresholds.Add(180) }

        # ===== APPLICATIONS =====
        if ($chkApplications.Checked) {
            Write-Step "📱 Récupération des applications..."
            try {
                $ManagedApps = @(Get-MobileAppsRest -AccessToken $AccessToken)
            } catch {
                $ManagedApps = @()
                $RunWarnings.Add("Applications du tenant : $($_.Exception.Message)")
            }
        }

        # ===== CLASSIFICATION PAR PLATEFORME =====
        Write-Step "🔍 Analyse des types d'appareils..."
        $WindowsDevices = @($DevicesScope | Where-Object { $_.OperatingSystem -like "Windows*" })
        $iOSDevices     = @($DevicesScope | Where-Object { $_.OperatingSystem -like "iOS*" })
        $AndroidDevices = @($DevicesScope | Where-Object { $_.OperatingSystem -like "Android*" })
        $MacDevices     = @($DevicesScope | Where-Object { $_.OperatingSystem -like "macOS*" -or $_.OperatingSystem -like "Mac*" })

        # ===== CONFORMITÉ =====
        if ($chkCompliance.Checked) {
            Write-Step "✓ Analyse de la conformité..."
            $CompliantDevices    = @($DevicesScope | Where-Object ComplianceState -EQ "compliant")
            $NoncompliantDevices = @($DevicesScope | Where-Object ComplianceState -EQ "noncompliant")
        }

        # ===== [v3.3] ANALYSE DES NON-CONFORMITÉS PAR CATÉGORIE / RAISON (API GRAPH) =====
        # Placée avant le chiffrement : en mode "par appareil", les postes déjà analysés
        # sont réutilisés (cache) pour la colonne RootCause de "Non Encrypted".
        # L'analyse gère ses propres erreurs : un échec n'interrompt pas le dashboard.
        $NcResult = $null
        if ($chkNcAnalysis.Checked) {
            Write-Step "🧩 Analyse des raisons de non-conformité (API Graph)..."
            $NcResult = Invoke-NonComplianceAnalysis -ScopeDevices $DevicesScope -AccessToken $AccessToken -Anonymize:$AnonymizeData
        }

        # ===== APPAREILS INACTIFS (paliers dynamiques) =====
        $InactiveDevicesByThreshold = @{}
        if ($InactiveThresholds.Count -gt 0) {
            Write-Step "⏰ Détection des appareils inactifs..."
            foreach ($days in $InactiveThresholds) {
                # [v4.0] Seuil en UTC, comme les dates de synchronisation
                $cutoff = [datetime]::UtcNow.AddDays(-$days)
                $InactiveDevicesByThreshold["$days"] = @($DevicesScope | Where-Object { $_.LastSyncDateTime -lt $cutoff })
            }
        }

        # ===== CHIFFREMENT =====
        if ($chkEncryption.Checked) {
            Write-Step "🔒 Vérification du chiffrement..."
            $EncryptedDevices  = @($DevicesScope | Where-Object IsEncrypted -EQ $true)
            $UnecryptedDevices = @($DevicesScope | Where-Object IsEncrypted -EQ $false)

            # ===== [MODIF v3] ENRICHISSEMENT "NON ENCRYPTED" : ROOT CAUSE + INACTIVITÉ =====
            # Pour chaque poste non chiffré :
            #  - calcul du nombre de jours d'inactivité (lastSyncDateTime de l'appareil)
            #  - récupération des raisons exactes de non-conformité via l'API Graph
            #    (deviceCompliancePolicyStates + settingStates)
            # [v4.0] États préchargés par lots de 20 ($batch) au lieu d'un appel par poste
            Initialize-NcSettingsCache -DeviceIds @($UnecryptedDevices | ForEach-Object { "$($_.Id)" }) -AccessToken $AccessToken
            $NonEncryptedList = [System.Collections.Generic.List[object]]::new()
            $idx    = 0
            $nowUtc = [datetime]::UtcNow   # lastSyncDateTime est en UTC

            foreach ($dev in $UnecryptedDevices) {
                $idx++
                if ($idx -eq 1 -or $idx % 5 -eq 0 -or $idx -eq $UnecryptedDevices.Count) {
                    Set-UiStatus "🔎 Analyse des causes racines (Non Encrypted) : $idx / $($UnecryptedDevices.Count)..."
                }

                # --- Jours d'inactivité (null = jamais synchronisé => traité comme inactif) ---
                $daysInactive = if ($dev.LastSyncDateTime) {
                    [int][math]::Floor(($nowUtc - $dev.LastSyncDateTime).TotalDays)
                } else { $null }

                # --- Raisons précises de non-conformité (cache, sinon appel Graph) ---
                $reasons = @(Get-DeviceNonComplianceReasons -DeviceId $dev.Id -AccessToken $AccessToken)

                # --- Construction du libellé RootCause affiché dans le dashboard ---
                if ($reasons.Count -gt 0) {
                    $rootCause = $reasons -join "  +  "
                } elseif ("$($dev.ComplianceState)" -eq 'compliant') {
                    # Le poste est conforme mais remonte IsEncrypted = false :
                    # ce n'est pas un écart de conformité, juste l'état matériel remonté
                    $rootCause = "BitLocker non actif (état remonté par l'appareil, aucun écart de conformité)"
                } else {
                    $rootCause = "Cause non déterminée (détail de conformité indisponible)"
                }
                if ($null -eq $daysInactive -or $daysInactive -ge $NonEncryptedInactiveDays) {
                    $rootCause += " (poste inactif - données possiblement obsolètes)"
                }

                $NonEncryptedList.Add([PSCustomObject]@{
                    DeviceName        = $dev.DeviceName
                    UserPrincipalName = $dev.UserPrincipalName
                    OperatingSystem   = $dev.OperatingSystem
                    Manufacturer      = $dev.Manufacturer
                    Model             = $dev.Model
                    OSVersion         = $dev.OSVersion
                    ComplianceState   = $dev.ComplianceState
                    RootCause         = $rootCause
                    DaysInactive      = $daysInactive
                    LastSyncDateTime  = $dev.LastSyncDateTime
                    EnrolledDateTime  = $dev.EnrolledDateTime
                })
            }

            # Tri : postes actifs en premier (les null = jamais vus, à la fin)
            $NonEncryptedEnriched = @($NonEncryptedList | Sort-Object @{ Expression = { if ($null -eq $_.DaysInactive) { 999999 } else { $_.DaysInactive } } })
        }

        # ===== INVENTAIRE =====
        $ManagedDevicesTable = $DevicesScope |
            Select-Object DeviceName, UserPrincipalName, OperatingSystem, Manufacturer, Model, OSVersion, ComplianceState, IsEncrypted, LastSyncDateTime, EnrolledDateTime |
            Sort-Object -Descending ComplianceState

        # ===== HARDWARE =====
        if ($chkHardware.Checked -and $WindowsDevices.Count -gt 0) {
            $OSVersions = $WindowsDevices | Group-Object OSVersion | Sort-Object Count -Descending | Select-Object Name, Count -First 10
            $Model      = $WindowsDevices | Group-Object { "$($_.Manufacturer), $($_.Model)" } | Sort-Object Count -Descending | Select-Object Name, Count -First 10
        }

        # ===== ÉCHECS APPS =====
        if ($chkApplications.Checked) {
            Write-Step "📱 Analyse des échecs d'installation..."
            $FailedAppsAll = @()
            try {
                $FailedAppsAll = @(Get-FailedAppsReportRest -AccessToken $AccessToken)
            } catch {
                # [v4.0] Erreur remontée (la v3 la masquait par un "catch {}" vide)
                $RunWarnings.Add("Rapport des échecs d'installation : $($_.Exception.Message)")
            }

            $FailedAppsTop10       = $FailedAppsAll | Sort-Object -Property FailedDeviceCount -Descending | Select-Object -First 10
            $LatestFailedAppsTop10 = [System.Collections.Generic.List[object]]::new()
            $FailedByName          = @{}
            foreach ($f in $FailedAppsAll) { if ($f.DisplayName -and -not $FailedByName.ContainsKey("$($f.DisplayName)")) { $FailedByName["$($f.DisplayName)"] = $f } }
            $ManagedAppsSorted     = $ManagedApps | Sort-Object -Property CreatedDateTime -Descending
            foreach ($app in $ManagedAppsSorted) {
                $failed = $FailedByName["$($app.DisplayName)"]
                if ($failed) {
                    $LatestFailedAppsTop10.Add([PSCustomObject]@{
                        DisplayName       = $failed.DisplayName
                        FailedDeviceCount = $failed.FailedDeviceCount
                        CreatedDateTime   = $app.CreatedDateTime
                    })
                }
                if ($LatestFailedAppsTop10.Count -ge 10) { break }
            }
            $AppsForFailedDonut = @(if ($LatestFailedAppsTop10.Count -gt 0) { $LatestFailedAppsTop10 } else { $FailedAppsTop10 })
        }

        # ===== [v4.0] PROFILS DE CONFIGURATION (Update Rings + vérification "profils") =====
        $DeviceConfigurations = @()
        if ($chkUpdateRings.Checked -or $HealthChecks.ConfigProfile) {
            Write-Step "⚙ Lecture des profils de configuration..."
            try {
                $DeviceConfigurations = @(Get-DeviceConfigurationsRest -AccessToken $AccessToken)
            } catch {
                $RunWarnings.Add("Profils de configuration : $($_.Exception.Message)")
            }
        }

        # ===== UPDATE RINGS =====
        if ($chkUpdateRings.Checked) {
            Write-Step "🔄 Analyse des Windows Update Rings..."
            $RingData           = Get-UpdateRingData -Configurations $DeviceConfigurations -AccessToken $AccessToken -Warnings $RunWarnings
            $UpdateRingsSummary = $RingData.Summary
            $UpdateRingsDevices = $RingData.Devices
            if ($AnonymizeData -and $UpdateRingsDevices.Count -gt 0) {
                $UpdateRingsDevices = @($UpdateRingsDevices | ForEach-Object {
                    $anon = Get-AnonymizedIdentity -RealName $_.DeviceName -RealUpn $_.UserName
                    [PSCustomObject]@{
                        RingName     = $_.RingName
                        DeviceName   = $anon.Name
                        UserName     = $anon.Upn
                        Status       = $_.Status
                        LastReported = $_.LastReported
                    }
                })
            }
        }

        # ===== [FUSION v4 - Proactivité] SANTÉ & PROACTIVITÉ DES POSTES =====
        # Réutilise les appareils déjà collectés (disque, inactivité, conformité : aucun
        # appel en plus) ; seules les vérifications cochées déclenchent des appels.
        $HealthData = $null
        if ($RunHealth) {
            $HealthRaw = Invoke-HealthCollection -Devices $DevicesScope -Checks $HealthChecks -DeviceConfigurations $DeviceConfigurations `
                             -AccessToken $AccessToken -Parallel $chkParallel.Checked -Warnings $HealthWarnings
            # Motifs de non-conformité ramenés au niveau du poste (clé : id Intune réel)
            $ComplianceReasonsById = @{}
            if ($NcResult -and $NcResult.Success) {
                foreach ($r in $NcResult.Rows) {
                    $key = [string]$r.DeviceKey
                    if (-not $ComplianceReasonsById.ContainsKey($key)) { $ComplianceReasonsById[$key] = [string]$r.Reason }
                    elseif ($ComplianceReasonsById[$key] -notlike "*$($r.Reason)*") { $ComplianceReasonsById[$key] += " ; $($r.Reason)" }
                }
            }
            Write-Step "🩺 Analyse de la santé des postes..."
            $HealthData = Build-RemediationData -Devices $DevicesScope -Health $HealthRaw -Checks $HealthChecks `
                              -ComplianceReasonsById $ComplianceReasonsById -AnonymizeData $AnonymizeData
            $level = if ($HealthData.CritCount -gt 0) { "WARN" } else { "OK" }
            Write-Log "Santé des postes : $(@($HealthData.Rows).Count) poste(s) - critiques $($HealthData.CritCount) / à surveiller $($HealthData.WarnCount) / OK $($HealthData.OkCount)" -Level $level
        }

        # ===== ANONYMISATION =====
        # [v4.0] Alias stables (Get-AnonymizedIdentity) : identiques dans toutes les pages et tous les CSV
        if ($AnonymizeData) {
            Write-Step "🔐 Anonymisation des données sensibles..."
            $ManagedDevicesTable = Anonymize-DeviceData -DeviceList $ManagedDevicesTable
            if ($chkCompliance.Checked) {
                $CompliantDevices    = Anonymize-DeviceData -DeviceList $CompliantDevices
                $NoncompliantDevices = Anonymize-DeviceData -DeviceList $NoncompliantDevices
            }
            if ($InactiveThresholds.Count -gt 0) {
                $newDict = @{}
                foreach ($days in $InactiveThresholds) {
                    $newDict["$days"] = Anonymize-DeviceData -DeviceList $InactiveDevicesByThreshold["$days"]
                }
                $InactiveDevicesByThreshold = $newDict
            }
            # [MODIF v3] Anonymisation de la liste enrichie "Non Encrypted"
            # (RootCause et DaysInactive sont préservés par Anonymize-DeviceData)
            if ($chkEncryption.Checked -and $NonEncryptedEnriched.Count -gt 0) {
                $NonEncryptedEnriched = Anonymize-DeviceData -DeviceList $NonEncryptedEnriched
            }
        }

        # ===== APPLICATIONS - TABLE =====
        if ($chkApplications.Checked) {
            $ManagedAppTable = $ManagedApps |
                Select-Object DisplayName, Publisher, Id, CreatedDateTime |
                Sort-Object -Descending CreatedDateTime
        }

        $Timestamp        = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
        $AnonymizedSuffix = if ($AnonymizeData) { "_ANONYMIZED" } else { "" }
        $SafeFileClient   = $ClientName -replace '[\\/:*?"<>|]', '_'
        $ReportFileName   = "Intune-Dashboard_${SafeFileClient}${AnonymizedSuffix}_${Timestamp}.html"

        # ===== [v4.0] EXPORTS CSV DE TOUTES LES SECTIONS =====
        # Un dossier horodaté par génération ; une section décochée n'a pas de fichier.
        $ExportFolder = ""
        $ExportFiles  = @()
        $NcCsvFiles   = @()
        if ($ExportCsv) {
            Write-Step "📄 Export CSV des données collectées..."
            $ExportFolder = Join-Path $OutputFolder "Intune-Export_${SafeFileClient}${AnonymizedSuffix}_${Timestamp}"
            $Datasets = [ordered]@{ "01_Appareils" = @($ManagedDevicesTable) }
            if ($chkCompliance.Checked) {
                $Datasets["02_Conformite_Non_Conformes"] = @($NoncompliantDevices | Select-Object DeviceName, UserPrincipalName, OperatingSystem, Manufacturer, Model, OSVersion, ComplianceState, IsEncrypted, LastSyncDateTime, EnrolledDateTime)
            }
            if ($NcResult -and $NcResult.Success -and $NcResult.Kpi.PostesTotal -gt 0) {
                foreach ($entry in (Get-NcCsvDatasets -Result $NcResult -ClientName $ClientName -Prefix "03_NonConformites_").GetEnumerator()) { $Datasets[$entry.Key] = $entry.Value }
            }
            if ($chkEncryption.Checked) {
                $Datasets["04_Chiffrement_Non_Chiffres"] = @($NonEncryptedEnriched | Select-Object DeviceName, UserPrincipalName, OperatingSystem, Manufacturer, Model, OSVersion, ComplianceState, RootCause, DaysInactive, LastSyncDateTime, EnrolledDateTime)
            }
            if ($chkApplications.Checked) {
                $Datasets["05_Applications"]        = @($ManagedAppTable)
                $Datasets["06_Applications_Echecs"] = @($FailedAppsAll)
            }
            if ($chkUpdateRings.Checked) {
                $Datasets["07_UpdateRings_Synthese"] = @($UpdateRingsSummary)
                $Datasets["08_UpdateRings_Postes"]   = @($UpdateRingsDevices)
            }
            if ($chkHardware.Checked) {
                $Datasets["09_Materiel_Modeles"]     = @($Model)
                $Datasets["10_Materiel_Versions_OS"] = @($OSVersions)
            }
            if ($HealthData) {
                $Datasets["11_Sante_Postes"]           = @(Select-HealthDeviceView -Rows $HealthData.Rows)
                $Datasets["12_Sante_Actions_Synthese"] = @(Get-HealthActionSummary -Actions $HealthData.Actions)
                $Datasets["13_Sante_Actions_Postes"]   = @(Select-HealthActionDetailView -Actions $HealthData.Actions)
                $Datasets["14_Sante_Espace_Disque"]    = @(Select-HealthDiskView -Rows $HealthData.Rows)
            }
            foreach ($days in $InactiveThresholds) {
                $Datasets[("15_Inactifs_{0:000}j" -f $days)] = @($InactiveDevicesByThreshold["$days"] | Select-Object DeviceName, UserPrincipalName, OperatingSystem, Manufacturer, Model, OSVersion, ComplianceState, IsEncrypted, LastSyncDateTime, EnrolledDateTime)
            }
            $ExportFiles = @(Export-AllDatasets -Datasets $Datasets -Folder $ExportFolder -Warnings $RunWarnings)
            $NcCsvFiles  = @($ExportFiles | Where-Object { [IO.Path]::GetFileName($_) -like "03_NonConformites_*" })
        }

        # ===== [v4.0] SANS RAPPORT HTML : BILAN ET FIN =====
        if (-not $GenerateHtml) {
            $lblStatus.Text = "✓ Collecte et exports CSV terminés"; $lblStatus.ForeColor = [System.Drawing.Color]::Green
            $message = "Collecte terminée (rapport HTML désactivé)."
            if ($ExportFiles.Count -gt 0) { $message += "`n`n$($ExportFiles.Count) fichier(s) CSV :`n$ExportFolder" }
            else                          { $message += "`n`nAucun fichier CSV n'a pu être écrit (voir les avertissements)." }
            if ($OpenAfterGeneration -and $ExportFiles.Count -gt 0) { $message += "`n`nLe dossier des exports s'ouvre automatiquement." }
            $warnText = Get-RunWarningsText -NcResult $NcResult -RunWarnings $RunWarnings -HealthWarnings $HealthWarnings
            if ($warnText) { Write-Log ($warnText.Trim()) -Level WARN }
            Write-Log "SUCCÈS (sans HTML) - $($ExportFiles.Count) fichier(s) CSV dans $ExportFolder" -Level OK
            Show-InfoMessage ($message + $warnText)
            if ($OpenAfterGeneration -and $ExportFiles.Count -gt 0) { Invoke-Item -LiteralPath $ExportFolder }
            return
        }

        # ===== GÉNÉRATION HTML =====
        Write-Step "📄 Génération du fichier HTML..."

        # ===== [MODIF v3] COMPTEURS "NON ENCRYPTED" POUR LE FILTRE DU DASHBOARD =====
        # Calculés côté PowerShell puis injectés dans le JavaScript du dashboard :
        # le compteur bascule entre (Total - Inactifs) et Total selon la case à cocher.
        if ($chkEncryption.Checked) {
            $NonEncTotal    = $NonEncryptedEnriched.Count
            $NonEncInactive = @($NonEncryptedEnriched | Where-Object {
                $null -eq $_.DaysInactive -or $_.DaysInactive -ge $NonEncryptedInactiveDays
            }).Count
            $NonEncActive   = $NonEncTotal - $NonEncInactive
        }

        $ShowContactInfo = $chkShowContact.Checked
        $CompanyName     = if ([string]::IsNullOrWhiteSpace($txtCompanyName.Text))   { $DefaultCompanyName }   else { $txtCompanyName.Text }
        $ContactPerson   = if ([string]::IsNullOrWhiteSpace($txtContactPerson.Text)) { $DefaultContactPerson } else { $txtContactPerson.Text }
        $ContactEmail    = if ([string]::IsNullOrWhiteSpace($txtContactEmail.Text))  { $DefaultContactEmail }  else { $txtContactEmail.Text }
        $ContactPhone    = if ([string]::IsNullOrWhiteSpace($txtContactPhone.Text))  { $DefaultContactPhone }  else { $txtContactPhone.Text }

        # ===== CONSTRUCTION HTML =====

        # ============================================================
        # [UI v3.1] PRÉPARATION DES COMPOSANTS VISUELS
        # Mise en forme uniquement : on lit les variables calculées plus
        # haut, aucune donnée n'est re-récupérée ni retraitée.
        # ============================================================
        $GeneratedAt  = (Get-Date).ToString("dd/MM/yyyy 'à' HH:mm")
        $TotalDevices = @($DevicesScope).Count

        # --- En-tête (hero glassmorphism) ---
        $SafeCompany = ConvertTo-HtmlSafe $CompanyName
        $SafeClient  = ConvertTo-HtmlSafe $ClientName

        $HeroChips = @(
            "<span class=`"ix-chip`">$(Get-IconSvg -Name 'calendar' -Size 14) Mis à jour le $GeneratedAt</span>"
            "<span class=`"ix-chip`">$(Get-IconSvg -Name 'monitor' -Size 14) $TotalDevices appareils analysés</span>"
        )
        if ($ExcludeVM)     { $HeroChips += "<span class=`"ix-chip`">$(Get-IconSvg -Name 'filter' -Size 14) Machines virtuelles exclues</span>" }
        if ($AnonymizeData) { $HeroChips += "<span class=`"ix-chip ix-chip--warn`">$(Get-IconSvg -Name 'lock' -Size 14) Données anonymisées - rapport confidentiel</span>" }

        $ContactHtml = ""
        if ($ShowContactInfo) {
            $SafePerson  = ConvertTo-HtmlSafe $ContactPerson
            $SafeEmail   = ConvertTo-HtmlSafe $ContactEmail
            $SafePhone   = ConvertTo-HtmlSafe $ContactPhone
            $ContactHtml = @"
    <div class="ix-hero__contact">
      <div class="ix-hero__contact-title">Votre contact</div>
      <div class="ix-hero__contact-row">$(Get-IconSvg -Name 'user' -Size 16)<span>$SafePerson</span></div>
      <div class="ix-hero__contact-row">$(Get-IconSvg -Name 'mail' -Size 16)<a href="mailto:$SafeEmail">$SafeEmail</a></div>
      <div class="ix-hero__contact-row">$(Get-IconSvg -Name 'phone' -Size 16)<span>$SafePhone</span></div>
    </div>
"@
        }

        $HeroHtml = @"
<div class="ix-root ix-hero">
  <div class="ix-hero__glow ix-hero__glow--a"></div>
  <div class="ix-hero__glow ix-hero__glow--b"></div>
  <div class="ix-hero__inner">
    <div class="ix-hero__brand">
      <div class="ix-hero__brandline"><span class="ix-hero__mark">$(Get-IconSvg -Name 'shield' -Size 16)</span>Rapport de parc Microsoft Intune</div>
      <div class="ix-hero__title">$SafeCompany</div>
      <div class="ix-hero__subtitle">État de santé des appareils pour <strong>$SafeClient</strong></div>
      <div class="ix-hero__meta">$($HeroChips -join '')</div>
    </div>
$ContactHtml
  </div>
</div>
"@

        # --- Cartes "Devices Overview" : total + plateformes (page Vue d'ensemble) ---
        $ov = [System.Text.StringBuilder]::new()
        [void]$ov.Append((New-IxStatCard -Label 'Total managed devices' -Value $TotalDevices -Icon 'monitor' -Color $Colors.Secondary -Caption 'Périmètre Intune analysé'))
        if ($ShowWindows) { [void]$ov.Append((New-IxStatCard -Label 'Windows' -Value $WindowsDevices.Count -Icon 'windows'    -Color $Colors.PlatformWindows -Caption "<b>$(Get-IxPercent $WindowsDevices.Count $TotalDevices) %</b> du parc")) }
        if ($ShowIOS)     { [void]$ov.Append((New-IxStatCard -Label 'iOS'     -Value $iOSDevices.Count     -Icon 'smartphone' -Color $Colors.PlatformIOS     -Caption "<b>$(Get-IxPercent $iOSDevices.Count $TotalDevices) %</b> du parc")) }
        if ($ShowAndroid) { [void]$ov.Append((New-IxStatCard -Label 'Android' -Value $AndroidDevices.Count -Icon 'bot'        -Color $Colors.PlatformAndroid -Caption "<b>$(Get-IxPercent $AndroidDevices.Count $TotalDevices) %</b> du parc")) }
        if ($ShowMac)     { [void]$ov.Append((New-IxStatCard -Label 'macOS'   -Value $MacDevices.Count     -Icon 'laptop'     -Color $Colors.PlatformMac     -Caption "<b>$(Get-IxPercent $MacDevices.Count $TotalDevices) %</b> du parc")) }
        $OverviewHtml = "<div class=`"ix-root ix-overview`">$($ov.ToString())</div>"

        # --- [v3.2] Cartes inactivité + stockage (page Santé & proactivité depuis la v4.0) ---
        $op = [System.Text.StringBuilder]::new()
        foreach ($days in $InactiveThresholds) {
            $count = $InactiveDevicesByThreshold["$days"].Count
            [void]$op.Append((New-IxStatCard -Label "Inactifs $days+ jours" -Value $count -Icon 'clock' -Color $InactiveColorMap["$days"] -Caption "<b>$(Get-IxPercent $count $TotalDevices) %</b> sans synchronisation"))
        }
        # [v4.0] Espace disque jugé en % (seuils du script Proactivité) au lieu de
        # "moins de 100 Go libres", qui signalait presque tous les portables 256 Go
        $DiskCritCount = 0; $DiskWarnCount = 0
        if ($HealthData -and $HealthChecks.Disk) {
            $DiskCritCount = @($HealthData.Rows | Where-Object { $_.DiskSeverity -eq 'crit' }).Count
            $DiskWarnCount = @($HealthData.Rows | Where-Object { $_.DiskSeverity -eq 'warn' }).Count
            $T = $RemediationThresholds
            [void]$op.Append((New-IxStatCard -Label 'Disque critique' -Value $DiskCritCount -Icon 'hard-drive' -Color $Colors.Danger -Caption "Moins de $($T.DiskFreePctCritical) % ou $($T.DiskFreeGbCritical) Go libres"))
            [void]$op.Append((New-IxStatCard -Label 'Disque à surveiller' -Value $DiskWarnCount -Icon 'hard-drive' -Color $Colors.Warning -Caption "Moins de $($T.DiskFreePctWarning) % d'espace libre"))
        }
        $OptimHtml = if ($op.Length -gt 0) { "<div class=`"ix-root ix-overview`">$($op.ToString())</div>" } else { "" }

        # --- [FUSION v4 - Proactivité] Page Santé : KPI, lecture, avertissements, synthèses ---
        $HealthHeadHtml = ""
        $HealthActions  = @()
        if ($HealthData) {
            $HealthHeadHtml = New-HealthDashboardHeadHtml -Data $HealthData -Warnings $HealthWarnings -ActiveChecks $ActiveHealthChecks
            $HealthActions  = @(Get-HealthActionSummary -Actions $HealthData.Actions)
        }

        # --- Update Rings : cartes avec taux de succès, barre de répartition et métriques ---
        $RingsHtml = ""
        if ($chkUpdateRings.Checked -and $UpdateRingsSummary.Count -gt 0) {
            $rb = [System.Text.StringBuilder]::new()
            foreach ($ring in $UpdateRingsSummary) {
                $rTotal   = [int]$ring.DeviceCount
                $rOk      = [int]$ring.Succeeded
                $rCf      = [int]$ring.Conflict
                $rEr      = [int]$ring.Error
                $rPending = [int]$ring.InProgress
                $rNa      = [int]$ring.NotApplicable
                $rOther   = [math]::Max(0, $rTotal - $rOk - $rCf - $rEr)
                $rRate    = Get-IxPercent $rOk $rTotal
                $rName    = ConvertTo-HtmlSafe $ring.RingName

                if     ($rEr -gt 0) { $rState = 'danger'; $rStateLabel = 'Erreurs' }
                elseif ($rCf -gt 0) { $rState = 'warn';   $rStateLabel = 'Conflits' }
                else                { $rState = 'ok';     $rStateLabel = 'Sain' }

                $zOk = if ($rOk -eq 0) { ' ix-metric--zero' } else { '' }
                $zCf = if ($rCf -eq 0) { ' ix-metric--zero' } else { '' }
                $zEr = if ($rEr -eq 0) { ' ix-metric--zero' } else { '' }

                $card = @"
<div class="ix-ring ix-ring--$rState">
  <div class="ix-ring__head">
    <div class="ix-ring__name" title="$rName">$rName</div>
    <span class="ix-badge ix-badge--$rState">$rStateLabel</span>
  </div>
  <div class="ix-ring__hero">
    <div>
      <div class="ix-ring__rate">$rRate<small>%</small></div>
      <div class="ix-ring__rate-label">Taux de succès</div>
    </div>
    <div class="ix-ring__devices">$(Get-IconSvg -Name 'monitor' -Size 15)<strong>$rTotal</strong> appareils</div>
  </div>
  <div class="ix-bar" role="img" aria-label="Succès $rOk, conflits $rCf, erreurs $rEr, autres $rOther">
    <span class="ix-bar__seg ix-bar__seg--ok"     style="flex-grow:$rOk"    title="Succès : $rOk"></span>
    <span class="ix-bar__seg ix-bar__seg--warn"   style="flex-grow:$rCf"    title="Conflits : $rCf"></span>
    <span class="ix-bar__seg ix-bar__seg--danger" style="flex-grow:$rEr"    title="Erreurs : $rEr"></span>
    <span class="ix-bar__seg ix-bar__seg--other"  style="flex-grow:$rOther" title="Autres : $rOther"></span>
  </div>
  <div class="ix-ring__metrics">
    <div class="ix-metric ix-metric--ok$zOk">
      <div class="ix-metric__head">$(Get-IconSvg -Name 'check' -Size 14)Succès</div>
      <div class="ix-metric__value">$rOk</div>
    </div>
    <div class="ix-metric ix-metric--warn$zCf">
      <div class="ix-metric__head">$(Get-IconSvg -Name 'alert' -Size 14)Conflits</div>
      <div class="ix-metric__value">$rCf</div>
    </div>
    <div class="ix-metric ix-metric--danger$zEr">
      <div class="ix-metric__head">$(Get-IconSvg -Name 'x-circle' -Size 14)Erreurs</div>
      <div class="ix-metric__value">$rEr</div>
    </div>
  </div>
  <div class="ix-ring__foot"><span>En cours : $rPending</span><span>Non applicable : $rNa</span></div>
</div>
"@
                [void]$rb.Append($card)
            }
            $RingsHtml = "<div class=`"ix-root ix-rings`">$($rb.ToString())</div>"
        }

        # --- [v3.3] Analyse des non-conformités : KPI, règle de lecture, source, avertissements ---
        $NcHeadHtml = ""
        if ($chkNcAnalysis.Checked -and $NcResult) {
            $NcHeadHtml = New-NcDashboardHeadHtml -Result $NcResult -TotalDevices $TotalDevices -CsvFolder $ExportFolder -CsvFiles $NcCsvFiles
        }
        $NcHasData = $chkNcAnalysis.Checked -and $NcResult -and $NcResult.Success -and $NcResult.Kpi.PostesTotal -gt 0

        # ============================================================
        # [UI v3.2] PAGES VIRTUELLES
        # Une page n'est créée (et son onglet affiché) que si au moins une
        # des sections qu'elle regroupe est cochée dans l'interface.
        # ============================================================
        $HasSecurity = $chkCompliance.Checked -or $chkEncryption.Checked -or $chkNcAnalysis.Checked
        $HasDeploy   = $chkUpdateRings.Checked -or $chkApplications.Checked
        # [v4.0] "Santé & proactivité" absorbe "Optimisation du parc" (inactifs + disque)
        $HasSante    = ($InactiveThresholds.Count -gt 0) -or [bool]$HealthData

        $IxPages = [System.Collections.Generic.List[object]]::new()
        $IxPages.Add([PSCustomObject]@{ Id = 'vue-ensemble'; Icon = 'layout'; Label = "Vue d'ensemble"; Hint = 'Taille et composition du parc : plateformes, modèles et versions de Windows' })
        if ($HasSecurity) { $IxPages.Add([PSCustomObject]@{ Id = 'securite';     Icon = 'shield';  Label = 'Sécurité & conformité';       Hint = 'Conformité Intune, raisons de non-conformité et chiffrement BitLocker' }) }
        if ($HasDeploy)   { $IxPages.Add([PSCustomObject]@{ Id = 'deploiement';  Icon = 'package'; Label = 'Mises à jour & applications'; Hint = 'Windows Update Rings et déploiement des applications' }) }
        if ($HasSante)    { $IxPages.Add([PSCustomObject]@{ Id = 'sante';        Icon = 'gauge';   Label = 'Santé & proactivité';         Hint = 'Endpoint Analytics, disque, démarrage, écrans bleus, batteries, inactivité et actions de remédiation' }) }

        $IxPage = @{}
        foreach ($p in $IxPages) { $IxPage[$p.Id] = $p }

        $UseSpa    = $IxPages.Count -gt 1
        $NavHtml   = if ($UseSpa) { New-IxNavigation -Pages $IxPages -ContextLabel $ClientName } else { "" }
        $SpaScript = if ($UseSpa) { Get-IxSpaScript } else { "" }

        New-HTML -TitleText "Intune Dashboard - $ClientName" -Online -FilePath "$OutputFolder\$ReportFileName" {

            # EN-TÊTE COMMUN À TOUTES LES PAGES : feuille de style globale + hero
            # (la date de mise à jour et le bandeau "anonymisé" sont intégrés au hero)
            New-HTMLSection -Invisible {
                New-HTMLPanel -Invisible {
                    New-HTMLText -Text (Get-IxGlobalCss) -FontSize 1
                    New-HTMLText -Text $HeroHtml -FontSize 1
                }
            }

            # [v3.2] MENU DE NAVIGATION (HTML brut, hors section pour rester collant en haut d'écran)
            $NavHtml
            '<div class="ix-pages">'

            # ============================================================
            # PAGE 1 : VUE D'ENSEMBLE — taille et composition du parc
            # ============================================================
            Get-IxPageStart -Page $IxPage['vue-ensemble']

            New-HTMLSection -HeaderText "Devices Overview" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Primary -CanCollapse {
                New-HTMLPanel {
                    New-HTMLText -Text $OverviewHtml -FontSize 1
                }
            }

            if ($chkHardware.Checked) {
                New-HTMLSection -Height 350 -HeaderText "System & Hardware Insights" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Hardware -CanCollapse {
                    New-HTMLPanel {
                        New-HTMLChart -Gradient {
                            foreach ($PC in $Model) { New-ChartBar -Name $PC.Name -Value $PC.Count }
                            New-ChartLegend -Name "Model", "Number of Devices"
                        } -Title "Top 10 Device Models" -TitleAlignment center -TitleColor $Colors.Primary
                    }
                    New-HTMLPanel {
                        New-HTMLChart -Gradient {
                            foreach ($OS in $OSVersions) { New-ChartBar -Name $OS.Name -Value $OS.Count }
                            New-ChartLegend -Name "OS Version", "Number of Devices"
                        } -Title "Top 10 Windows Versions" -TitleAlignment center -TitleColor $Colors.Primary
                    }
                }
            }

            '</div>'

            # ============================================================
            # PAGE 2 : SÉCURITÉ & CONFORMITÉ — conformité et BitLocker
            # ============================================================
            if ($HasSecurity) {
                Get-IxPageStart -Page $IxPage['securite']

                # [v3.3] Section affichée seulement si l'un de ses deux graphiques est demandé
                # (la page peut désormais n'exister que pour l'analyse des non-conformités)
                if ($chkCompliance.Checked -or $chkEncryption.Checked) {
                    New-HTMLSection -Height 350 -HeaderText "Security & Compliance" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Compliance -CanCollapse {
                        if ($chkCompliance.Checked) {
                            New-HTMLPanel {
                                New-HTMLChart -Gradient {
                                    New-ChartDonut -Name "Compliant"     -Value $CompliantDevices.Count    -Color $Colors.Success
                                    New-ChartDonut -Name "Non-Compliant" -Value $NoncompliantDevices.Count -Color $Colors.Danger
                                } -Title "Compliance Status" -TitleAlignment center -TitleColor $Colors.Primary
                            }
                        }
                        if ($chkEncryption.Checked) {
                            New-HTMLPanel {
                                New-HTMLChart -Gradient {
                                    New-ChartDonut -Name "Encrypted (BitLocker)" -Value $EncryptedDevices.Count  -Color $Colors.Success
                                    New-ChartDonut -Name "Not Encrypted"         -Value $UnecryptedDevices.Count -Color $Colors.Danger
                                } -Title "BitLocker Status" -TitleAlignment center -TitleColor $Colors.Primary
                            }
                        }
                    }
                }

                # [v3.3] ANALYSE DES NON-CONFORMITÉS : KPI, graphiques et synthèses par catégorie / raison
                if ($chkNcAnalysis.Checked -and $NcResult) {
                    New-HTMLSection -HeaderText "Non-Compliance Analysis" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Compliance -CanCollapse {
                        New-HTMLPanel {
                            New-HTMLText -Text $NcHeadHtml -FontSize 1
                        }
                    }
                }

                if ($NcHasData) {
                    New-HTMLSection -Height 380 -HeaderText "Non-Compliance by Category & Reason" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Compliance -CanCollapse {
                        New-HTMLPanel {
                            New-HTMLChart -Gradient {
                                foreach ($ncCat in $NcResult.CategorySummary) { New-ChartBar -Name $ncCat.'Catégorie' -Value $ncCat.'Postes impactés' }
                                New-ChartLegend -Name "Postes impactés"
                            } -Title "Postes impactés par catégorie" -TitleAlignment center -TitleColor $Colors.Primary
                        }
                        New-HTMLPanel {
                            New-HTMLChart -Gradient {
                                foreach ($ncReason in ($NcResult.ReasonSummary | Select-Object -First 10)) { New-ChartBar -Name $ncReason.'Raison' -Value $ncReason.'Postes impactés' }
                                New-ChartLegend -Name "Postes impactés"
                            } -Title "Top 10 des raisons de non-conformité" -TitleAlignment center -TitleColor $Colors.Primary
                        }
                        New-HTMLPanel {
                            New-HTMLChart -Gradient {
                                New-ChartDonut -Name $NcResult.Kpi.ParSynchro[0].Libelle -Value $NcResult.Kpi.ParSynchro[0].Valeur -Color $Colors.Success
                                New-ChartDonut -Name $NcResult.Kpi.ParSynchro[1].Libelle -Value $NcResult.Kpi.ParSynchro[1].Valeur -Color $Colors.Warning
                                New-ChartDonut -Name $NcResult.Kpi.ParSynchro[2].Libelle -Value $NcResult.Kpi.ParSynchro[2].Valeur -Color $Colors.Danger
                            } -Title "Ancienneté de la dernière synchronisation" -TitleAlignment center -TitleColor $Colors.Primary
                        }
                    }

                    New-HTMLSection -HeaderText "Non-Compliance Summary by Category" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse {
                        New-HTMLTable -DataTable $NcResult.CategorySummary -Filtering -PagingLength 25
                    }
                    New-HTMLSection -HeaderText "Non-Compliance Summary by Reason" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse {
                        New-HTMLTable -DataTable $NcResult.ReasonSummary -Filtering -PagingLength 25
                    }
                    New-HTMLSection -HeaderText "Non-Compliant Devices - Triage by Priority" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                        New-HTMLTable -DataTable $NcResult.DeviceSummary -Filtering -PagingLength 50
                    }
                    New-HTMLSection -HeaderText "Non-Compliance Detail (Device x Setting)" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                        New-HTMLTable -DataTable @(Select-NcDetailView -Rows $NcResult.Rows) -Filtering -PagingLength 50
                    }
                }

                if ($chkCompliance.Checked) {
                    New-HTMLSection -HeaderText "Non-Compliant Devices" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                        New-HTMLTable -DataTable ($NoncompliantDevices | Select-Object DeviceName, UserPrincipalName, OperatingSystem, Manufacturer, Model, OSVersion, ComplianceState, IsEncrypted, LastSyncDateTime, EnrolledDateTime) -Filtering -PagingLength 50
                    }
                }

                # SECTION "NON ENCRYPTED" : bandeau (compteur dynamique, mini-KPI, interrupteur) + table.
                # Les IDs utilisés par le JavaScript sont inchangés :
                #   nonEncCounter / nonEncHiddenInfo / chkHideInactiveNonEnc / TableNonEncrypted
                if ($chkEncryption.Checked) {
                    $alertVariant = if ($NonEncTotal -eq 0) { ' ix-alert--ok' } else { '' }
                    $alertIcon    = if ($NonEncTotal -eq 0) { 'shield-ok' } else { 'unlock' }
                    New-HTMLSection -HeaderText "Non Encrypted Devices (BitLocker)" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                        New-HTMLPanel {

                            New-HTMLText -Text @"
<div class="ix-root ix-alert$alertVariant">
  <div class="ix-alert__main">
    <div class="ix-alert__icon">$(Get-IconSvg -Name $alertIcon -Size 26 -Stroke '1.75')</div>
    <div>
      <div class="ix-alert__title">Postes sans chiffrement BitLocker</div>
      <div class="ix-alert__count">
        <span id="nonEncCounter">$NonEncActive</span>
        <span class="ix-alert__count-label">poste(s) non chiffré(s) affiché(s)</span>
      </div>
      <span id="nonEncHiddenInfo" class="ix-pill">$(Get-IconSvg -Name 'eye-off' -Size 13)$NonEncInactive poste(s) inactif(s) ${NonEncryptedInactiveDays}j+ masqué(s)</span>
    </div>
  </div>
  <div class="ix-alert__side">
    <div class="ix-kpis">
      <div class="ix-kpi ix-kpi--danger"><span class="ix-kpi__v">$NonEncTotal</span><span class="ix-kpi__l">Total</span></div>
      <div class="ix-kpi"><span class="ix-kpi__v">$NonEncActive</span><span class="ix-kpi__l">Actifs</span></div>
      <div class="ix-kpi ix-kpi--muted"><span class="ix-kpi__v">$NonEncInactive</span><span class="ix-kpi__l">Inactifs ${NonEncryptedInactiveDays}j+</span></div>
    </div>
    <label class="ix-switch">
      <input type="checkbox" id="chkHideInactiveNonEnc" checked
             onchange="window.updateNonEncryptedFilter && window.updateNonEncryptedFilter()">
      <span class="ix-switch__track"><span class="ix-switch__thumb"></span></span>
      <span>Masquer les postes inactifs depuis ${NonEncryptedInactiveDays} jours ou plus</span>
    </label>
  </div>
</div>

<script>
/* [MODIF v3] Filtre client "Non Encrypted" :
   - masque les lignes dont DaysInactive >= seuil quand la case est cochee
   - met a jour dynamiquement le compteur de la section (Total - Inactifs) */
(function () {
    var TABLE_ID   = 'TableNonEncrypted';        /* ID fixe via -DataTableID cote PowerShell */
    var THRESHOLD  = $NonEncryptedInactiveDays;  /* seuil (jours) injecte par PowerShell */
    var TOTAL      = $NonEncTotal;               /* total postes non chiffres */
    var INACTIVE   = $NonEncInactive;            /* dont inactifs >= seuil */
    var daysColIdx = -1;                         /* index colonne "DaysInactive" (detecte a l'init) */

    /* Met a jour compteur + badge, puis redessine la table selon l'etat de la case */
    window.updateNonEncryptedFilter = function () {
        var chk     = document.getElementById('chkHideInactiveNonEnc');
        var counter = document.getElementById('nonEncCounter');
        var info    = document.getElementById('nonEncHiddenInfo');
        var checked = !!(chk && chk.checked);
        if (counter) { counter.textContent = checked ? (TOTAL - INACTIVE) : TOTAL; }
        if (info)    { info.style.display  = checked ? '' : 'none'; }   /* [UI v3.1] '' = style CSS de la pastille */
        if (window.jQuery && jQuery.fn.dataTable && jQuery.fn.dataTable.isDataTable('#' + TABLE_ID)) {
            jQuery('#' + TABLE_ID).DataTable().draw();
        }
    };

    function init() {
        /* On attend que jQuery + DataTables (PSWriteHTML) aient initialise la table */
        if (!(window.jQuery && jQuery.fn.dataTable && jQuery.fn.dataTable.isDataTable('#' + TABLE_ID))) {
            setTimeout(init, 300);
            return;
        }
        var dt = jQuery('#' + TABLE_ID).DataTable();

        /* Detection de l'index de la colonne "DaysInactive" (robuste au reordonnancement) */
        var headers = dt.columns().header().toArray();
        for (var i = 0; i < headers.length; i++) {
            if (headers[i] && headers[i].textContent && headers[i].textContent.trim() === 'DaysInactive') {
                daysColIdx = i;
                break;
            }
        }

        /* Filtre personnalise DataTables, applique uniquement a cette table */
        jQuery.fn.dataTable.ext.search.push(function (settings, data) {
            if (!settings.nTable || settings.nTable.id !== TABLE_ID) { return true; }
            var chk = document.getElementById('chkHideInactiveNonEnc');
            if (!chk || !chk.checked || daysColIdx < 0) { return true; }
            var d = parseFloat(data[daysColIdx]);
            if (isNaN(d)) { return false; }   /* pas de date de sync = inactif => masque */
            return d < THRESHOLD;             /* visible uniquement si < seuil */
        });

        window.updateNonEncryptedFilter();    /* application initiale (case cochee par defaut) */
    }

    if (document.readyState === 'complete') { init(); } else { window.addEventListener('load', init); }
})();
</script>
"@ -FontSize 1

                            # Tableau détaillé avec RootCause + DaysInactive (inchangé)
                            New-HTMLTable -DataTable ($NonEncryptedEnriched | Select-Object DeviceName, UserPrincipalName, OperatingSystem, Manufacturer, Model, OSVersion, ComplianceState, RootCause, DaysInactive, LastSyncDateTime, EnrolledDateTime) -Filtering -PagingLength 50 -DataTableID 'TableNonEncrypted'
                        }
                    }
                }

                '</div>'
            }

            # ============================================================
            # PAGE 3 : MISES À JOUR & APPLICATIONS — Update Rings et déploiements
            # ============================================================
            if ($HasDeploy) {
                Get-IxPageStart -Page $IxPage['deploiement']

                # UPDATE RINGS — cartes responsives (grille auto-fit)
                if ($chkUpdateRings.Checked) {
                    New-HTMLSection -HeaderText "Update Rings" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.UpdateRings -CanCollapse {
                        New-HTMLPanel {
                            if ($RingsHtml) {
                                New-HTMLText -Text $RingsHtml -FontSize 1
                            } else {
                                New-HTMLText -Text (New-IxEmptyState -Icon 'refresh' -Tone 'neutral' -Title "Aucun Update Ring à afficher" -Text "Aucune stratégie Windows Update for Business n'a remonté de statut pour un appareil utilisateur. Vérifiez les affectations dans Intune.") -FontSize 1
                            }
                        }
                    }
                }

                if ($chkUpdateRings.Checked -and $UpdateRingsDevices.Count -gt 0) {
                    New-HTMLSection -HeaderText "Windows Update Rings Status (Per Device)" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                        New-HTMLTable -DataTable ($UpdateRingsDevices | Select-Object RingName, DeviceName, UserName, Status, LastReported) -Filtering -PagingLength 50
                    }
                }

                # APPLICATIONS
                if ($chkApplications.Checked) {
                    New-HTMLSection -Height 350 -HeaderText "Applications" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Applications -CanCollapse {
                        if ($AppsForFailedDonut.Count -gt 0) {
                            New-HTMLPanel {
                                New-HTMLChart -Gradient {
                                    $chartColors = @($Colors.Danger, "#f87171", "#fb923c", $Colors.Warning, "#fbbf24", "#a3e635", "#34d399", "#0e7490", $Colors.Secondary, $Colors.DetailTables)
                                    $colorIndex  = 0
                                    foreach ($app in $AppsForFailedDonut) {
                                        New-ChartDonut -Name $app.DisplayName -Value $app.FailedDeviceCount -Color $chartColors[$colorIndex % $chartColors.Count]
                                        $colorIndex++
                                    }
                                } -Title "Top 10 Latest Apps with Installation Failures" -TitleAlignment center -TitleColor $Colors.Primary
                            }
                            New-HTMLPanel {
                                New-HTMLTable -DataTable ($FailedAppsAll | Select-Object DisplayName, FailedDeviceCount) -Filtering -PagingLength 10
                            }
                        } else {
                            New-HTMLPanel {
                                New-HTMLText -Text (New-IxEmptyState -Icon 'check' -Tone 'success' -Title "Aucun échec d'installation" -Text "Toutes les applications déployées se sont installées sans erreur remontée par Intune.") -FontSize 1
                            }
                        }
                    }

                    New-HTMLSection -HeaderText "All Intune Applications" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                        New-HTMLTable -DataTable $ManagedAppTable -Filtering -PagingLength 50
                    }
                }

                '</div>'
            }

            # ============================================================
            # PAGE 4 : SANTÉ & PROACTIVITÉ — [FUSION v4 - Proactivité]
            # Santé des postes (14 vérifications) + inactivité et stockage (ex-page
            # "Optimisation du parc", conservée à l'identique pour les paliers d'inactivité)
            # ============================================================
            if ($HasSante) {
                Get-IxPageStart -Page $IxPage['sante']

                if ($HealthData) {
                    New-HTMLSection -HeaderText "Fleet Health & Proactivity" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Health -CanCollapse {
                        New-HTMLPanel {
                            New-HTMLText -Text $HealthHeadHtml -FontSize 1
                        }
                    }
                }

                # [v3.2] Cartes inactivité / stockage (auparavant dans "Devices Overview")
                if ($OptimHtml) {
                    New-HTMLSection -HeaderText "Inactive Devices & Storage" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Primary -CanCollapse {
                        New-HTMLPanel {
                            New-HTMLText -Text $OptimHtml -FontSize 1
                        }
                    }
                }

                if ($HealthData -and @($HealthData.Rows).Count -gt 0) {
                    New-HTMLSection -Height 380 -HeaderText "Health Overview & Top Actions" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Health -CanCollapse {
                        New-HTMLPanel {
                            New-HTMLChart -Gradient {
                                New-ChartDonut -Name "Critique"     -Value $HealthData.CritCount -Color $Colors.Danger
                                New-ChartDonut -Name "À surveiller" -Value $HealthData.WarnCount -Color $Colors.Warning
                                New-ChartDonut -Name "OK"           -Value $HealthData.OkCount   -Color $Colors.Success
                            } -Title "Sévérité des postes" -TitleAlignment center -TitleColor $Colors.Primary
                        }
                        if ($HealthActions.Count -gt 0) {
                            New-HTMLPanel {
                                New-HTMLChart -Gradient {
                                    foreach ($hAction in ($HealthActions | Select-Object -First 10)) { New-ChartBar -Name $hAction.'Action recommandée' -Value $hAction.'Postes concernés' }
                                    New-ChartLegend -Name "Postes concernés"
                                } -Title "Top 10 des actions de remédiation" -TitleAlignment center -TitleColor $Colors.Primary
                            }
                        }
                    }

                    New-HTMLSection -HeaderText "Remediation Actions (Summary)" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse {
                        New-HTMLPanel {
                            if ($HealthActions.Count -gt 0) {
                                New-HTMLTable -DataTable $HealthActions -Filtering -PagingLength 25
                            } else {
                                New-HTMLText -Text (New-IxEmptyState -Icon 'check' -Tone 'success' -Title "Aucune action de remédiation nécessaire" -Text "Tous les indicateurs collectés sont dans les seuils.") -FontSize 1
                            }
                        }
                    }

                    if ($HealthActions.Count -gt 0) {
                        New-HTMLSection -HeaderText "Remediation Actions (Device Detail)" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                            New-HTMLTable -DataTable @(Select-HealthActionDetailView -Actions $HealthData.Actions) -Filtering -PagingLength 50
                        }
                    }

                    New-HTMLSection -HeaderText "Device Health (All Indicators)" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                        New-HTMLTable -DataTable @(Select-HealthDeviceView -Rows $HealthData.Rows) -Filtering -PagingLength 50
                    }

                    if ($HealthChecks.Disk -and ($DiskCritCount + $DiskWarnCount) -gt 0) {
                        New-HTMLSection -HeaderText "Devices with Low Disk Space (< $($RemediationThresholds.DiskFreePctWarning) % free)" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                            New-HTMLTable -DataTable @(Select-HealthDiskView -Rows $HealthData.Rows) -Filtering -PagingLength 50
                        }
                    }
                }

                # Tables inactifs - une par palier sélectionné
                foreach ($days in $InactiveThresholds) {
                    $color   = $InactiveColorMap["$days"]
                    $devices = $InactiveDevicesByThreshold["$days"]
                    New-HTMLSection -HeaderText "Inactive Devices ($days+ days)" -HeaderTextSize 14 -HeaderBackGroundColor $color -CanCollapse -Collapsed {
                        New-HTMLTable -DataTable ($devices | Select-Object DeviceName, UserPrincipalName, OperatingSystem, Manufacturer, Model, OSVersion, ComplianceState, IsEncrypted, LastSyncDateTime, EnrolledDateTime | Sort-Object LastSyncDateTime) -Filtering -PagingLength 50
                    }
                }

                '</div>'
            }

            '</div>'

            # [v3.2] Script de navigation (après toutes les pages)
            $SpaScript

        } -ShowHTML:$OpenAfterGeneration

        $lblStatus.Text = "✓ Dashboard généré avec succès !"; $lblStatus.ForeColor = [System.Drawing.Color]::Green
        $message = "Dashboard généré avec succès !`n`nFichier : $OutputFolder\$ReportFileName"
        if ($OpenAfterGeneration) { $message += "`n`nLe rapport s'ouvre automatiquement dans votre navigateur." }
        else                      { $message += "`n`nLe rapport est disponible dans le dossier de sortie." }

        # [v4.0] Exports CSV et avertissements de toutes les sections (API, droits, Endpoint Analytics)
        if ($ExportFiles.Count -gt 0) {
            $message += "`n`n$($ExportFiles.Count) fichier(s) CSV :`n$ExportFolder"
        }
        if ($NcResult -and -not $NcResult.Success) {
            $lblStatus.Text = "⚠ Dashboard généré - analyse des non-conformités indisponible"; $lblStatus.ForeColor = [System.Drawing.Color]::DarkOrange
        }
        $warnText = Get-RunWarningsText -NcResult $NcResult -RunWarnings $RunWarnings -HealthWarnings $HealthWarnings
        if ($warnText) { Write-Log ($warnText.Trim()) -Level WARN }
        Write-Log "SUCCÈS - Rapport : $OutputFolder\$ReportFileName" -Level OK
        Show-InfoMessage ($message + $warnText)

    } catch {
        $lblStatus.Text = "❌ Erreur lors de la génération"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "ÉCHEC - $($_.Exception.Message)" -Level ERROR
        Write-Log "Ligne : $($_.InvocationInfo.ScriptLineNumber) / Trace : $($_.ScriptStackTrace)" -Level ERROR
        Show-ErrorMessage "Erreur lors de la génération du dashboard :`n`n$($_.Exception.Message)`n`n(Détail complet dans $LogFile)"
    } finally {
        # [v3.3] Le secret client et le cache d'analyse ne survivent pas à la génération
        $script:GraphContext    = $null
        $script:NcSettingsCache = $null
    }
}

# ========================================
# INTERFACE GRAPHIQUE
# ========================================

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Intune Dashboard v4.0 - All-In-One"
$form.Size            = New-Object System.Drawing.Size(980, 700)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.BackColor       = [System.Drawing.Color]::FromArgb(240, 242, 245)

# Icône
$iconBitmap   = New-Object System.Drawing.Bitmap(32, 32)
$iconGraphics = [System.Drawing.Graphics]::FromImage($iconBitmap)
$iconGraphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$iconGraphics.Clear([System.Drawing.Color]::FromArgb(0, 120, 212))
$whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$iconGraphics.FillRectangle($whiteBrush, 6, 16, 4, 12)
$iconGraphics.FillRectangle($whiteBrush, 14, 10, 4, 18)
$iconGraphics.FillRectangle($whiteBrush, 22, 6, 4, 22)
$whiteBrush.Dispose(); $iconGraphics.Dispose()
$form.Icon = [System.Drawing.Icon]::FromHandle($iconBitmap.GetHicon())

# En-tête
$headerPanel           = New-Object System.Windows.Forms.Panel
$headerPanel.Location  = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size      = New-Object System.Drawing.Size(980, 90)
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$form.Controls.Add($headerPanel)

$lblTitle           = New-Object System.Windows.Forms.Label
$lblTitle.Location  = New-Object System.Drawing.Point(30, 20)
$lblTitle.Size      = New-Object System.Drawing.Size(900, 50)
$lblTitle.Text      = "Intune Dashboard  —  v4.0 All-In-One"
$lblTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$headerPanel.Controls.Add($lblTitle)

# TabControl
$tabControl          = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(20, 105)
$tabControl.Size     = New-Object System.Drawing.Size(940, 490)
$tabControl.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$form.Controls.Add($tabControl)

# ========================================
# ONGLET 1 : CONFIGURATION
# ========================================

$tabConfig           = New-Object System.Windows.Forms.TabPage
$tabConfig.Text      = "Configuration"
$tabConfig.BackColor = [System.Drawing.Color]::White
$tabControl.Controls.Add($tabConfig)

$lblClient          = New-Object System.Windows.Forms.Label
$lblClient.Location = New-Object System.Drawing.Point(30, 30)
$lblClient.Size     = New-Object System.Drawing.Size(150, 25)
$lblClient.Text     = "Client Intune :"
$lblClient.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabConfig.Controls.Add($lblClient)

$cmbClients               = New-Object System.Windows.Forms.ComboBox
$cmbClients.Location      = New-Object System.Drawing.Point(190, 28)
$cmbClients.Size          = New-Object System.Drawing.Size(700, 25)
$cmbClients.DropDownStyle = "DropDownList"
$cmbClients.Font          = New-Object System.Drawing.Font("Segoe UI", 10)
$tabConfig.Controls.Add($cmbClients)

$grpOptions          = New-Object System.Windows.Forms.GroupBox
$grpOptions.Location = New-Object System.Drawing.Point(30, 75)
$grpOptions.Size     = New-Object System.Drawing.Size(860, 90)
$grpOptions.Text     = " Options générales "
$grpOptions.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabConfig.Controls.Add($grpOptions)

$chkExcludeVM          = New-Object System.Windows.Forms.CheckBox
$chkExcludeVM.Location = New-Object System.Drawing.Point(20, 28)
$chkExcludeVM.Size     = New-Object System.Drawing.Size(820, 25)
$chkExcludeVM.Text     = "🖥️  Exclure les machines virtuelles des statistiques"
$chkExcludeVM.Checked  = $true
$chkExcludeVM.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpOptions.Controls.Add($chkExcludeVM)

$chkAnonymize          = New-Object System.Windows.Forms.CheckBox
$chkAnonymize.Location = New-Object System.Drawing.Point(20, 58)
$chkAnonymize.Size     = New-Object System.Drawing.Size(820, 25)
$chkAnonymize.Text     = "🔒  Anonymiser les données sensibles (format : User-XXXX / Poste-XXXX)"
$chkAnonymize.Checked  = $false
$chkAnonymize.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpOptions.Controls.Add($chkAnonymize)

# [v4.0] Sortie : rapport HTML facultatif, exports CSV de toutes les sections
$grpOutput          = New-Object System.Windows.Forms.GroupBox
$grpOutput.Location = New-Object System.Drawing.Point(30, 180)
$grpOutput.Size     = New-Object System.Drawing.Size(860, 125)
$grpOutput.Text     = " Sortie "
$grpOutput.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabConfig.Controls.Add($grpOutput)

$chkGenerateHtml          = New-Object System.Windows.Forms.CheckBox
$chkGenerateHtml.Location = New-Object System.Drawing.Point(20, 28)
$chkGenerateHtml.Size     = New-Object System.Drawing.Size(820, 25)
$chkGenerateHtml.Text     = "🌐  Générer le rapport HTML (dashboard à onglets, module PSWriteHTML)"
$chkGenerateHtml.Checked  = $true
$chkGenerateHtml.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpOutput.Controls.Add($chkGenerateHtml)

$chkExportCsv          = New-Object System.Windows.Forms.CheckBox
$chkExportCsv.Location = New-Object System.Drawing.Point(20, 58)
$chkExportCsv.Size     = New-Object System.Drawing.Size(820, 25)
$chkExportCsv.Text     = "📄  Exporter toutes les données collectées en CSV (dossier horodaté dans $OutputFolder)"
$chkExportCsv.Checked  = $true
$chkExportCsv.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpOutput.Controls.Add($chkExportCsv)

$lblOutputInfo           = New-Object System.Windows.Forms.Label
$lblOutputInfo.Location  = New-Object System.Drawing.Point(20, 88)
$lblOutputInfo.Size      = New-Object System.Drawing.Size(820, 30)
$lblOutputInfo.Text      = "Sans HTML, seules les requêtes Graph et les exports CSV sont exécutés (aucun module requis) ; « Générer et Ouvrir » ouvre alors le dossier des exports."
$lblOutputInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$lblOutputInfo.ForeColor = [System.Drawing.Color]::Gray
$grpOutput.Controls.Add($lblOutputInfo)

# Sans HTML, l'export CSV devient la seule sortie : il est forcé
$chkGenerateHtml.Add_CheckedChanged({
    if (-not $chkGenerateHtml.Checked) { $chkExportCsv.Checked = $true }
    $chkExportCsv.Enabled = $chkGenerateHtml.Checked
    $btnGenerateOnly.Text = if ($chkGenerateHtml.Checked) { "Générer le Dashboard" } else { "Collecter et exporter" }
    $btnGenerateOpen.Text = if ($chkGenerateHtml.Checked) { "Générer et Ouvrir" } else { "Exporter et ouvrir le dossier" }
})

# ========================================
# ONGLET 2 : CONTACT
# ========================================

$tabContact           = New-Object System.Windows.Forms.TabPage
$tabContact.Text      = "Contact"
$tabContact.BackColor = [System.Drawing.Color]::White
$tabControl.Controls.Add($tabContact)

$lblContactInfo           = New-Object System.Windows.Forms.Label
$lblContactInfo.Location  = New-Object System.Drawing.Point(30, 20)
$lblContactInfo.Size      = New-Object System.Drawing.Size(860, 25)
$lblContactInfo.Text      = "Informations de contact affichées en haut du dashboard"
$lblContactInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$lblContactInfo.ForeColor = [System.Drawing.Color]::Gray
$tabContact.Controls.Add($lblContactInfo)

$chkShowContact           = New-Object System.Windows.Forms.CheckBox
$chkShowContact.Location  = New-Object System.Drawing.Point(30, 52)
$chkShowContact.Size      = New-Object System.Drawing.Size(500, 25)
$chkShowContact.Text      = "Afficher les coordonnées de contact dans le dashboard"
$chkShowContact.Checked   = $true
$chkShowContact.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$chkShowContact.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$tabContact.Controls.Add($chkShowContact)

function Add-ContactField {
    param($Parent, $YPos, $Label, $DefaultValue, [ref]$TextBox)
    $lbl          = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(30, $YPos + 3)
    $lbl.Size     = New-Object System.Drawing.Size(190, 25)
    $lbl.Text     = $Label
    $lbl.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $Parent.Controls.Add($lbl)
    $txt              = New-Object System.Windows.Forms.TextBox
    $txt.Location     = New-Object System.Drawing.Point(230, $YPos)
    $txt.Size         = New-Object System.Drawing.Size(670, 25)
    $txt.Text         = $DefaultValue
    $txt.Font         = New-Object System.Drawing.Font("Segoe UI", 9)
    $txt.BorderStyle  = "FixedSingle"
    $Parent.Controls.Add($txt)
    $TextBox.Value = $txt
}

$txtCompanyName   = $null; Add-ContactField -Parent $tabContact -YPos  95 -Label "Nom de l'entreprise :"  -DefaultValue $DefaultCompanyName   -TextBox ([ref]$txtCompanyName)
$txtContactPerson = $null; Add-ContactField -Parent $tabContact -YPos 140 -Label "Nom du contact :"       -DefaultValue $DefaultContactPerson -TextBox ([ref]$txtContactPerson)
$txtContactEmail  = $null; Add-ContactField -Parent $tabContact -YPos 185 -Label "Email du contact :"     -DefaultValue $DefaultContactEmail  -TextBox ([ref]$txtContactEmail)
$txtContactPhone  = $null; Add-ContactField -Parent $tabContact -YPos 230 -Label "Téléphone du contact :" -DefaultValue $DefaultContactPhone  -TextBox ([ref]$txtContactPhone)

$btnResetContact           = New-Object System.Windows.Forms.Button
$btnResetContact.Location  = New-Object System.Drawing.Point(230, 275)
$btnResetContact.Size      = New-Object System.Drawing.Size(220, 35)
$btnResetContact.Text      = "Réinitialiser les valeurs"
$btnResetContact.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnResetContact.FlatStyle = "Flat"
$btnResetContact.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125)
$btnResetContact.ForeColor = [System.Drawing.Color]::White
$btnResetContact.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnResetContact.Add_Click({
    $txtCompanyName.Text   = $DefaultCompanyName
    $txtContactPerson.Text = $DefaultContactPerson
    $txtContactEmail.Text  = $DefaultContactEmail
    $txtContactPhone.Text  = $DefaultContactPhone
})
$tabContact.Controls.Add($btnResetContact)

# ========================================
# ONGLET 3 : CONTENU
# ========================================

$tabContent           = New-Object System.Windows.Forms.TabPage
$tabContent.Text      = "Contenu"
$tabContent.BackColor = [System.Drawing.Color]::White
$tabControl.Controls.Add($tabContent)

$lblContentInfo           = New-Object System.Windows.Forms.Label
$lblContentInfo.Location  = New-Object System.Drawing.Point(30, 15)
$lblContentInfo.Size      = New-Object System.Drawing.Size(860, 25)
$lblContentInfo.Text      = "Sélectionnez les sections à inclure dans le dashboard"
$lblContentInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$lblContentInfo.ForeColor = [System.Drawing.Color]::Gray
$tabContent.Controls.Add($lblContentInfo)

# Sections principales
$grpMainSections          = New-Object System.Windows.Forms.GroupBox
$grpMainSections.Location = New-Object System.Drawing.Point(30, 48)
$grpMainSections.Size     = New-Object System.Drawing.Size(430, 220)
$grpMainSections.Text     = " Sections principales "
$grpMainSections.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabContent.Controls.Add($grpMainSections)

$chkCompliance          = New-Object System.Windows.Forms.CheckBox
$chkCompliance.Location = New-Object System.Drawing.Point(20, 30); $chkCompliance.Size = New-Object System.Drawing.Size(390, 25)
$chkCompliance.Text     = "Compliance Status (Conformité)"; $chkCompliance.Checked = $true
$chkCompliance.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpMainSections.Controls.Add($chkCompliance)

$chkEncryption          = New-Object System.Windows.Forms.CheckBox
$chkEncryption.Location = New-Object System.Drawing.Point(20, 62); $chkEncryption.Size = New-Object System.Drawing.Size(390, 25)
$chkEncryption.Text     = "BitLocker Encryption Status"; $chkEncryption.Checked = $true
$chkEncryption.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpMainSections.Controls.Add($chkEncryption)

$chkApplications          = New-Object System.Windows.Forms.CheckBox
$chkApplications.Location = New-Object System.Drawing.Point(20, 94); $chkApplications.Size = New-Object System.Drawing.Size(390, 25)
$chkApplications.Text     = "Applications (Échecs d'installation)"; $chkApplications.Checked = $true
$chkApplications.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpMainSections.Controls.Add($chkApplications)

$chkUpdateRings          = New-Object System.Windows.Forms.CheckBox
$chkUpdateRings.Location = New-Object System.Drawing.Point(20, 126); $chkUpdateRings.Size = New-Object System.Drawing.Size(390, 25)
$chkUpdateRings.Text     = "Windows Update Rings"; $chkUpdateRings.Checked = $true
$chkUpdateRings.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpMainSections.Controls.Add($chkUpdateRings)

$chkHardware          = New-Object System.Windows.Forms.CheckBox
$chkHardware.Location = New-Object System.Drawing.Point(20, 158); $chkHardware.Size = New-Object System.Drawing.Size(390, 25)
$chkHardware.Text     = "Hardware & OS Versions"; $chkHardware.Checked = $true
$chkHardware.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpMainSections.Controls.Add($chkHardware)

# [FUSION v4 - Proactivité] Case maître de l'analyse de santé des postes (le détail des
# 14 vérifications se règle dans l'onglet « Proactivité »)
$chkHealth          = New-Object System.Windows.Forms.CheckBox
$chkHealth.Location = New-Object System.Drawing.Point(20, 190); $chkHealth.Size = New-Object System.Drawing.Size(390, 25)
$chkHealth.Text     = "Santé & proactivité des postes (Endpoint Analytics)"; $chkHealth.Checked = $true
$chkHealth.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpMainSections.Controls.Add($chkHealth)

# Sections supplémentaires (vides - déplacées dans Affichage)
$grpExtraSections          = New-Object System.Windows.Forms.GroupBox
$grpExtraSections.Location = New-Object System.Drawing.Point(480, 48)
$grpExtraSections.Size     = New-Object System.Drawing.Size(430, 220)
$grpExtraSections.Text     = " Info "
$grpExtraSections.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabContent.Controls.Add($grpExtraSections)

$lblExtraInfo           = New-Object System.Windows.Forms.Label
$lblExtraInfo.Location  = New-Object System.Drawing.Point(20, 35)
$lblExtraInfo.Size      = New-Object System.Drawing.Size(395, 140)
$lblExtraInfo.Text      = "Les plateformes et les paliers d'appareils inactifs se configurent dans l'onglet « Affichage ».`n`nLes 14 vérifications de santé (disque, Endpoint Analytics, batteries, BitLocker, Defender...) et la collecte parallèle se configurent dans l'onglet « Proactivité ».`n`nLa génération HTML et les exports CSV se règlent dans l'onglet « Configuration »."
$lblExtraInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblExtraInfo.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$grpExtraSections.Controls.Add($lblExtraInfo)

# Boutons tout sélectionner / tout désélectionner
$btnSelectAll           = New-Object System.Windows.Forms.Button
$btnSelectAll.Location  = New-Object System.Drawing.Point(30, 280)
$btnSelectAll.Size      = New-Object System.Drawing.Size(200, 35)
$btnSelectAll.Text      = "Tout sélectionner"
$btnSelectAll.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSelectAll.FlatStyle = "Flat"
$btnSelectAll.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnSelectAll.ForeColor = [System.Drawing.Color]::White
$btnSelectAll.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnSelectAll.Add_Click({
    $chkCompliance.Checked   = $true; $chkEncryption.Checked  = $true
    $chkApplications.Checked = $true; $chkUpdateRings.Checked = $true
    $chkHardware.Checked     = $true; $chkNcAnalysis.Checked  = $true
    $chkHealth.Checked       = $true
})
$tabContent.Controls.Add($btnSelectAll)

$btnDeselectAll           = New-Object System.Windows.Forms.Button
$btnDeselectAll.Location  = New-Object System.Drawing.Point(245, 280)
$btnDeselectAll.Size      = New-Object System.Drawing.Size(200, 35)
$btnDeselectAll.Text      = "Tout désélectionner"
$btnDeselectAll.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnDeselectAll.FlatStyle = "Flat"
$btnDeselectAll.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
$btnDeselectAll.ForeColor = [System.Drawing.Color]::White
$btnDeselectAll.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnDeselectAll.Add_Click({
    $chkCompliance.Checked   = $false; $chkEncryption.Checked  = $false
    $chkApplications.Checked = $false; $chkUpdateRings.Checked = $false
    $chkHardware.Checked     = $false; $chkNcAnalysis.Checked  = $false
    $chkHealth.Checked       = $false
})
$tabContent.Controls.Add($btnDeselectAll)

# [v3.3] Analyse des non-conformités (page "Sécurité & conformité")
$grpNonCompliance          = New-Object System.Windows.Forms.GroupBox
$grpNonCompliance.Location = New-Object System.Drawing.Point(30, 330)
$grpNonCompliance.Size     = New-Object System.Drawing.Size(880, 95)
$grpNonCompliance.Text     = " Sécurité & conformité — analyse des non-conformités (API Graph) "
$grpNonCompliance.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabContent.Controls.Add($grpNonCompliance)

$chkNcAnalysis          = New-Object System.Windows.Forms.CheckBox
$chkNcAnalysis.Location = New-Object System.Drawing.Point(20, 28); $chkNcAnalysis.Size = New-Object System.Drawing.Size(840, 25)
$chkNcAnalysis.Text     = "🔎  Analyser les raisons de non-conformité (catégories, raisons, ancienneté de synchro, actionnabilité)"; $chkNcAnalysis.Checked = $true
$chkNcAnalysis.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpNonCompliance.Controls.Add($chkNcAnalysis)

# [v4.0] L'export CSV des synthèses est désormais global (onglet Configuration > Sortie)
$lblNcInfo           = New-Object System.Windows.Forms.Label
$lblNcInfo.Location  = New-Object System.Drawing.Point(20, 58)
$lblNcInfo.Size      = New-Object System.Drawing.Size(840, 30)
$lblNcInfo.Text      = "Données lues en direct (rapport Intune « Noncompliant devices and settings », repli par appareil). Permissions Graph (Application) : DeviceManagementManagedDevices.Read.All, DeviceManagementConfiguration.Read.All."
$lblNcInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$lblNcInfo.ForeColor = [System.Drawing.Color]::Gray
$grpNonCompliance.Controls.Add($lblNcInfo)

# ========================================
# ONGLET 4 : AFFICHAGE (NOUVEAU)
# ========================================

$tabAffichage           = New-Object System.Windows.Forms.TabPage
$tabAffichage.Text      = "Affichage"
$tabAffichage.BackColor = [System.Drawing.Color]::White
$tabControl.Controls.Add($tabAffichage)

$lblAffichageInfo           = New-Object System.Windows.Forms.Label
$lblAffichageInfo.Location  = New-Object System.Drawing.Point(30, 15)
$lblAffichageInfo.Size      = New-Object System.Drawing.Size(860, 25)
$lblAffichageInfo.Text      = "Plateformes de la page « Vue d'ensemble » et paliers d'inactivité de la page « Santé & proactivité »"
$lblAffichageInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
$lblAffichageInfo.ForeColor = [System.Drawing.Color]::Gray
$tabAffichage.Controls.Add($lblAffichageInfo)

# --- Groupe Plateformes ---
$grpPlateformes          = New-Object System.Windows.Forms.GroupBox
$grpPlateformes.Location = New-Object System.Drawing.Point(30, 48)
$grpPlateformes.Size     = New-Object System.Drawing.Size(430, 130)
$grpPlateformes.Text     = " Plateformes à afficher "
$grpPlateformes.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabAffichage.Controls.Add($grpPlateformes)

$chkShowWindows           = New-Object System.Windows.Forms.CheckBox
$chkShowWindows.Location  = New-Object System.Drawing.Point(20, 30); $chkShowWindows.Size = New-Object System.Drawing.Size(185, 25)
$chkShowWindows.Text      = "🪟  Windows"; $chkShowWindows.Checked = $true
$chkShowWindows.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$grpPlateformes.Controls.Add($chkShowWindows)

$chkShowIOS           = New-Object System.Windows.Forms.CheckBox
$chkShowIOS.Location  = New-Object System.Drawing.Point(225, 30); $chkShowIOS.Size = New-Object System.Drawing.Size(185, 25)
$chkShowIOS.Text      = "📱  iOS"; $chkShowIOS.Checked = $true
$chkShowIOS.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$grpPlateformes.Controls.Add($chkShowIOS)

$chkShowAndroid           = New-Object System.Windows.Forms.CheckBox
$chkShowAndroid.Location  = New-Object System.Drawing.Point(20, 65); $chkShowAndroid.Size = New-Object System.Drawing.Size(185, 25)
$chkShowAndroid.Text      = "🤖  Android"; $chkShowAndroid.Checked = $true
$chkShowAndroid.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$grpPlateformes.Controls.Add($chkShowAndroid)

$chkShowMac           = New-Object System.Windows.Forms.CheckBox
$chkShowMac.Location  = New-Object System.Drawing.Point(225, 65); $chkShowMac.Size = New-Object System.Drawing.Size(185, 25)
$chkShowMac.Text      = "🍎  macOS"; $chkShowMac.Checked = $false
$chkShowMac.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
$grpPlateformes.Controls.Add($chkShowMac)

$btnSelectAllPlatforms           = New-Object System.Windows.Forms.Button
$btnSelectAllPlatforms.Location  = New-Object System.Drawing.Point(20, 98); $btnSelectAllPlatforms.Size = New-Object System.Drawing.Size(150, 24)
$btnSelectAllPlatforms.Text      = "Toutes"; $btnSelectAllPlatforms.FlatStyle = "Flat"
$btnSelectAllPlatforms.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69); $btnSelectAllPlatforms.ForeColor = [System.Drawing.Color]::White
$btnSelectAllPlatforms.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnSelectAllPlatforms.Add_Click({ $chkShowWindows.Checked = $true; $chkShowIOS.Checked = $true; $chkShowAndroid.Checked = $true; $chkShowMac.Checked = $true })
$grpPlateformes.Controls.Add($btnSelectAllPlatforms)

$btnDeselectAllPlatforms           = New-Object System.Windows.Forms.Button
$btnDeselectAllPlatforms.Location  = New-Object System.Drawing.Point(180, 98); $btnDeselectAllPlatforms.Size = New-Object System.Drawing.Size(150, 24)
$btnDeselectAllPlatforms.Text      = "Aucune"; $btnDeselectAllPlatforms.FlatStyle = "Flat"
$btnDeselectAllPlatforms.BackColor = [System.Drawing.Color]::FromArgb(108, 117, 125); $btnDeselectAllPlatforms.ForeColor = [System.Drawing.Color]::White
$btnDeselectAllPlatforms.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnDeselectAllPlatforms.Add_Click({ $chkShowWindows.Checked = $false; $chkShowIOS.Checked = $false; $chkShowAndroid.Checked = $false; $chkShowMac.Checked = $false })
$grpPlateformes.Controls.Add($btnDeselectAllPlatforms)

# [v4.0] Groupe "Low Storage (< 100 GB)" retiré : l'espace disque est évalué en %
# par la vérification "Espace disque" de l'onglet « Proactivité ».

# --- Groupe Inactive Devices ---
$grpInactive          = New-Object System.Windows.Forms.GroupBox
$grpInactive.Location = New-Object System.Drawing.Point(480, 48)
$grpInactive.Size     = New-Object System.Drawing.Size(430, 290)
$grpInactive.Text     = " Inactive Devices — Paliers à afficher "
$grpInactive.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabAffichage.Controls.Add($grpInactive)

$inactiveDays   = @(30, 60, 90, 120, 150, 180)
$inactiveChecks = @{}

for ($i = 0; $i -lt $inactiveDays.Count; $i++) {
    $d   = $inactiveDays[$i]
    $col = $InactiveColorMap["$d"]
    $chk = New-Object System.Windows.Forms.CheckBox
    $chk.Location = New-Object System.Drawing.Point(20, (28 + $i * 36))
    $chk.Size     = New-Object System.Drawing.Size(380, 28)
    $chk.Text     = "⏰  $d jours"
    $chk.Checked  = ($d -eq 30 -or $d -eq 90 -or $d -eq 180)  # défaut : 30 / 90 / 180
    $chk.Font     = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $chk.ForeColor = [System.Drawing.Color]::FromArgb(
        [Convert]::ToInt32($col.Substring(1,2), 16),
        [Convert]::ToInt32($col.Substring(3,2), 16),
        [Convert]::ToInt32($col.Substring(5,2), 16)
    )
    $grpInactive.Controls.Add($chk)
    $inactiveChecks["$d"] = $chk
}

# Référencer chaque checkbox pour Generate-Dashboard
$chkInactive30  = $inactiveChecks["30"]
$chkInactive60  = $inactiveChecks["60"]
$chkInactive90  = $inactiveChecks["90"]
$chkInactive120 = $inactiveChecks["120"]
$chkInactive150 = $inactiveChecks["150"]
$chkInactive180 = $inactiveChecks["180"]

$btnSelectAllInactive           = New-Object System.Windows.Forms.Button
$btnSelectAllInactive.Location  = New-Object System.Drawing.Point(20, 248); $btnSelectAllInactive.Size = New-Object System.Drawing.Size(175, 30)
$btnSelectAllInactive.Text      = "Tous les paliers"; $btnSelectAllInactive.FlatStyle = "Flat"
$btnSelectAllInactive.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69); $btnSelectAllInactive.ForeColor = [System.Drawing.Color]::White
$btnSelectAllInactive.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnSelectAllInactive.Add_Click({ foreach ($d in $inactiveDays) { $inactiveChecks["$d"].Checked = $true } })
$grpInactive.Controls.Add($btnSelectAllInactive)

$btnDeselectAllInactive           = New-Object System.Windows.Forms.Button
$btnDeselectAllInactive.Location  = New-Object System.Drawing.Point(205, 248); $btnDeselectAllInactive.Size = New-Object System.Drawing.Size(175, 30)
$btnDeselectAllInactive.Text      = "Aucun palier"; $btnDeselectAllInactive.FlatStyle = "Flat"
$btnDeselectAllInactive.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69); $btnDeselectAllInactive.ForeColor = [System.Drawing.Color]::White
$btnDeselectAllInactive.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
$btnDeselectAllInactive.Add_Click({ foreach ($d in $inactiveDays) { $inactiveChecks["$d"].Checked = $false } })
$grpInactive.Controls.Add($btnDeselectAllInactive)

# ========================================
# [FUSION v4 - Proactivité] ONGLET 5 : PROACTIVITÉ (santé des postes)
# ========================================
# Une case par vérification (seuil en info-bulle) : une vérification décochée n'est
# ni collectée ni affichée. Repris de l'onglet "Remédiation avancée" du script
# Proactivité (les imports de fichiers et les actions Sync / Reboot n'en font pas partie).

$tabProactivite           = New-Object System.Windows.Forms.TabPage
$tabProactivite.Text      = "Proactivité"
$tabProactivite.BackColor = [System.Drawing.Color]::White
$tabControl.Controls.Add($tabProactivite)

$tipHealth              = New-Object System.Windows.Forms.ToolTip
$tipHealth.AutoPopDelay = 25000
$tipHealth.InitialDelay = 350

$grpHealthChecks          = New-Object System.Windows.Forms.GroupBox
$grpHealthChecks.Location = New-Object System.Drawing.Point(30, 15)
$grpHealthChecks.Size     = New-Object System.Drawing.Size(880, 190)
$grpHealthChecks.Text     = " Vérifications de la page « Santé & proactivité » "
$grpHealthChecks.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabProactivite.Controls.Add($grpHealthChecks)

$lblHealthChecksInfo           = New-Object System.Windows.Forms.Label
$lblHealthChecksInfo.Location  = New-Object System.Drawing.Point(20, 25)
$lblHealthChecksInfo.Size      = New-Object System.Drawing.Size(840, 20)
$lblHealthChecksInfo.Text      = "Une vérification décochée n'est ni collectée ni affichée (aucun appel Graph). Survolez une case pour voir son seuil."
$lblHealthChecksInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$lblHealthChecksInfo.ForeColor = [System.Drawing.Color]::Gray
$grpHealthChecks.Controls.Add($lblHealthChecksInfo)

$script:ChkHealthChecks = @{}
foreach ($def in $HealthCheckDefs) {
    $c          = New-Object System.Windows.Forms.CheckBox
    $c.Location = New-Object System.Drawing.Point((20 + $def.Col * 212), (52 + $def.Row * 32))
    $c.Size     = New-Object System.Drawing.Size(205, 25)
    $c.Text     = $def.Text
    $c.Checked  = $true
    $c.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
    $tipHealth.SetToolTip($c, $def.Hint)
    $grpHealthChecks.Controls.Add($c)
    $script:ChkHealthChecks[$def.Key] = $c
}

$btnSelectAllHealth           = New-Object System.Windows.Forms.Button
$btnSelectAllHealth.Location  = New-Object System.Drawing.Point(30, 215)
$btnSelectAllHealth.Size      = New-Object System.Drawing.Size(200, 32)
$btnSelectAllHealth.Text      = "Toutes les vérifications"
$btnSelectAllHealth.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnSelectAllHealth.FlatStyle = "Flat"
$btnSelectAllHealth.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnSelectAllHealth.ForeColor = [System.Drawing.Color]::White
$btnSelectAllHealth.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnSelectAllHealth.Add_Click({ foreach ($c in $script:ChkHealthChecks.Values) { $c.Checked = $true } })
$tabProactivite.Controls.Add($btnSelectAllHealth)

$btnDeselectAllHealth           = New-Object System.Windows.Forms.Button
$btnDeselectAllHealth.Location  = New-Object System.Drawing.Point(245, 215)
$btnDeselectAllHealth.Size      = New-Object System.Drawing.Size(200, 32)
$btnDeselectAllHealth.Text      = "Aucune vérification"
$btnDeselectAllHealth.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$btnDeselectAllHealth.FlatStyle = "Flat"
$btnDeselectAllHealth.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
$btnDeselectAllHealth.ForeColor = [System.Drawing.Color]::White
$btnDeselectAllHealth.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnDeselectAllHealth.Add_Click({ foreach ($c in $script:ChkHealthChecks.Values) { $c.Checked = $false } })
$tabProactivite.Controls.Add($btnDeselectAllHealth)

$grpHealthCollect          = New-Object System.Windows.Forms.GroupBox
$grpHealthCollect.Location = New-Object System.Drawing.Point(30, 260)
$grpHealthCollect.Size     = New-Object System.Drawing.Size(880, 150)
$grpHealthCollect.Text     = " Collecte "
$grpHealthCollect.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabProactivite.Controls.Add($grpHealthCollect)

$chkParallel          = New-Object System.Windows.Forms.CheckBox
$chkParallel.Location = New-Object System.Drawing.Point(20, 28)
$chkParallel.Size     = New-Object System.Drawing.Size(840, 25)
$chkParallel.Text     = "⚡  Collecte parallèle des listes Endpoint Analytics et BitLocker ($MaxParallelCollections flux, pause commune en cas de limitation Graph)"
$chkParallel.Checked  = $true
$chkParallel.Font     = New-Object System.Drawing.Font("Segoe UI", 9)
$grpHealthCollect.Controls.Add($chkParallel)

$lblHealthCollectInfo           = New-Object System.Windows.Forms.Label
$lblHealthCollectInfo.Location  = New-Object System.Drawing.Point(20, 60)
$lblHealthCollectInfo.Size      = New-Object System.Drawing.Size(840, 82)
$lblHealthCollectInfo.Text      = "Prérequis : l'Analyse des points de terminaison doit être activée dans Intune (sinon, seuls le disque, l'inactivité, la conformité, Defender, les mises à jour et les profils sont évalués).`nPermissions Graph (Application) : DeviceManagementManagedDevices.Read.All, DeviceManagementConfiguration.Read.All.`nSeuils modifiables en tête de script (`$RemediationThresholds). Journal détaillé : $LogFile"
$lblHealthCollectInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
$lblHealthCollectInfo.ForeColor = [System.Drawing.Color]::Gray
$grpHealthCollect.Controls.Add($lblHealthCollectInfo)

# La case maître (onglet Contenu) active ou grise l'ensemble des vérifications
$chkHealth.Add_CheckedChanged({
    $grpHealthChecks.Enabled      = $chkHealth.Checked
    $btnSelectAllHealth.Enabled   = $chkHealth.Checked
    $btnDeselectAllHealth.Enabled = $chkHealth.Checked
})

# ========================================
# BOUTONS DE GÉNÉRATION
# ========================================

$btnGenerateOnly           = New-Object System.Windows.Forms.Button
$btnGenerateOnly.Location  = New-Object System.Drawing.Point(250, 610)
$btnGenerateOnly.Size      = New-Object System.Drawing.Size(210, 45)
$btnGenerateOnly.Text      = "Générer le Dashboard"
$btnGenerateOnly.BackColor = [System.Drawing.Color]::FromArgb(40, 167, 69)
$btnGenerateOnly.ForeColor = [System.Drawing.Color]::White
$btnGenerateOnly.FlatStyle = "Flat"
$btnGenerateOnly.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnGenerateOnly.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnGenerateOnly.Add_Click({ Generate-Dashboard -OpenAfterGeneration $false })
$form.Controls.Add($btnGenerateOnly)

$btnGenerateOpen           = New-Object System.Windows.Forms.Button
$btnGenerateOpen.Location  = New-Object System.Drawing.Point(475, 610)
$btnGenerateOpen.Size      = New-Object System.Drawing.Size(230, 45)
$btnGenerateOpen.Text      = "Générer et Ouvrir"
$btnGenerateOpen.BackColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$btnGenerateOpen.ForeColor = [System.Drawing.Color]::White
$btnGenerateOpen.FlatStyle = "Flat"
$btnGenerateOpen.Font      = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$btnGenerateOpen.Cursor    = [System.Windows.Forms.Cursors]::Hand
$btnGenerateOpen.Add_Click({ Generate-Dashboard -OpenAfterGeneration $true })
$form.Controls.Add($btnGenerateOpen)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(20, 660)
$lblStatus.Size      = New-Object System.Drawing.Size(940, 25)
$lblStatus.Text      = "Sélectionnez un client et configurez votre dashboard"
$lblStatus.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblStatus.ForeColor = [System.Drawing.Color]::Gray
$lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$form.Controls.Add($lblStatus)

# ========================================
# LANCEMENT
# ========================================

Load-ClientConfigs
[void]$form.ShowDialog()