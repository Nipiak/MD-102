# ============================================================
# Script : Intune Dashboard - Conformité, Apps Découvertes & Inventaire
# Description : Génère un dashboard HTML interactif à 4 onglets :
#                 1) Analyse de conformité (Non-compliant / Grace Period) groupée par motif
#                 2) Rapport Discovered Apps (accordéon déplier/replier par application)
#                 3) Inventaire des applications du tenant + audit de version (best-effort via winget)
#                 4) Remédiation & santé des postes (espace disque, scores Endpoint Analytics,
#                    temps de démarrage, écrans bleus, batteries, appareils inactifs)
#               COLLECTE : trois modes au choix dans l'onglet "Source des données" —
#                 * API Microsoft Graph (comportement historique) ;
#                 * Import de fichiers (CSV/TSV/JSON) : génération 100 % hors ligne,
#                   sans aucune connexion au tenant ni configuration client ;
#                 * Mixte : API pour ce qui n'a pas de fichier, fichier pour le reste.
#               Une collecte API peut en outre exporter ses données au format réimportable,
#               afin de rejouer/archiver un rapport sans redemander d'accès au tenant.
#               Réutilise le socle d'authentification / de configuration client du script de base.
#               ANONYMISATION (onglet "Client & Options") : postes/utilisateurs, noms
#               d'applications (apps_N) et d'exécutables (proc_N), nom du client et termes
#               sensibles (X), avec table de correspondance confidentielle à côté du rapport.
#               AUCUN module externe requis (ni PSWriteHTML, ni Microsoft.Graph) : rendu HTML
#               sur mesure, fichier 100% autonome consultable hors ligne.
# Nouveautés V2.2 (onglet 2 "Applications découvertes") :
#               - liste COMPLÈTE des postes de chaque application détaillée : toutes les pages
#                 Graph sont lues (fin de "Affichage limité aux 999 premiers postes"), les
#                 pages suivantes étant elles aussi groupées par 20 dans des appels $batch
#               - "Nombre d'applications à détailler" = 0 : toutes les applications
#               - postes affichés par tranches de 20 (pagination), recherche par poste /
#                 utilisateur / OS dans chaque application, export CSV de la liste complète
#                 ou filtrée ; la recherche globale trouve aussi les postes non affichés
#               - données des postes stockées une seule fois (bloc JSON compact) au lieu d'un
#                 tableau HTML par application : rapport plus léger et plus rapide à ouvrir
#               - masquage des termes sensibles étendu à ces données
#               Les ajouts sont repérés par la balise [V2.2].
# Nouveautés V2.3 : TOUTES les applications découvertes ont leur liste de postes par défaut
#               (case « Toutes les applications », cochée) ; la limite aux N plus répandues
#               (50 auparavant, appliquée sans le dire) n'est plus qu'une option. Balise [V2.3].
# Auteur : ECONOCOM
# ============================================================

# ===== IMPORTS ET ASSEMBLIES =====
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===== RÉGLAGES RÉSEAU (.NET) =====
# .NET n'autorise que DEUX connexions HTTP simultanées vers un même hôte (héritage de la
# RFC HTTP/1.1). C'est le tout premier goulot d'étranglement d'une collecte parallèle vers
# graph.microsoft.com : sans cette ligne, ouvrir six runspaces n'en fait travailler que deux.
try { [System.Net.ServicePointManager]::DefaultConnectionLimit = 24 } catch { }
# TLS 1.2 ajouté sans écraser ce qui est déjà négociable (TLS 1.3 sur les OS récents).
try { [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12 } catch { }

# ===== CONFIGURATION GLOBALE =====
# NOTE : ConfigFolder et AESKey sont identiques au script de base afin de pouvoir
# réutiliser directement les fichiers .clientconfig déjà générés.
$ConfigFolder = "C:\temp\clients-id"
$OutputFolder = "C:\temp"
$LogFile      = "C:\temp\dashboard-log.txt"
# Référentiel des versions saisies manuellement : chargé automatiquement à chaque
# génération et alimenté par la fenêtre "Versions manuelles" de l'onglet Rapports.
$ManualVersionsFile = "C:\temp\versions-manuelles.csv"
# Jeton GitHub facultatif (lecture seule, dépôt public) utilisé pour interroger le
# catalogue winget-pkgs sans être limité à 60 requêtes/heure.
$GitHubTokenFile = "C:\temp\github-token.txt"
# Applications volontairement masquées de l'onglet "Inventaire & versions" du rapport :
# liste cochée/décochée dans la fenêtre "Versions manuelles", persistée d'une session à l'autre.
$ExcludedAppsFile = "C:\temp\apps-masquees.csv"
# Anonymisation du rapport : termes propres au client (nom commercial, sigle, nom d'outil
# interne...) à masquer, mémorisés PAR CLIENT d'une session à l'autre.
$AnonTermsFile = "C:\temp\anonymisation-termes.json"
# Libellé qui remplace le nom du client et les termes masqués dans le rapport.
$AnonClientLabel = "X"
# Dossier des fichiers d'import / d'export de donnees brutes (mode hors ligne).
# Utilise par la collecte "Import de fichiers" et par l'option d'export des donnees collectees.
$ImportFolder = "C:\temp\imports"

# ===== PERFORMANCE DE LA COLLECTE =====
# Les collectes de listes indépendantes (scores, performances, BitLocker, batteries,
# fiabilité applicative, historique de démarrage) sont lancées de front dans un pool de
# runspaces plutôt qu'à la file. Sur un parc réel, c'est le principal gain de temps de
# génération. Passez à $false pour revenir à un enchaînement strictement séquentiel
# (diagnostic, hôte très contraint, ou tenant déjà fortement limité).
$script:UseParallelCollection = $true
# Nombre de collectes menées de front. Au-delà de 4-6, on déclenche le throttling Graph
# et le gain se retourne en perte : ne montez qu'en connaissance de cause.
$script:MaxParallelCollections = 4
# Nombre maximum de profils de configuration analysés pour la vérification "erreurs de
# profil" (garde-fou de durée sur les tenants qui en comptent plusieurs centaines).
$script:MaxConfigProfilesAnalyzed = 300

# ===== COORDONNÉES ENTREPRISE PAR DÉFAUT =====
$DefaultCompanyName   = "ECONOCOM"
$DefaultContactPerson = "Nom si nécessaire"
$DefaultContactEmail  = "support-Intune-2IP@econocom.com"
$DefaultContactPhone  = "+33 ...."

# ============================================================
# THÈME DE L'INTERFACE — PERSONNALISABLE
# Modifiez librement ces valeurs (titres, couleurs hexadécimales,
# police, rayon d'arrondi) pour adapter la fenêtre à votre charte.
# ============================================================
$Theme = @{
    WindowTitle    = "Intune Dashboard — Conformité, Apps & Inventaire"   # Titre de la barre de fenêtre
    HeaderTitle    = "Intune Dashboard"                                    # Grand titre de l'en-tête
    HeaderSubtitle = "Conformité  ·  Applications découvertes  ·  Inventaire & versions  ·  Remédiation"
    HeaderColor1   = "#452069"    # Dégradé de l'en-tête : couleur de départ
    HeaderColor2   = "#7B3FA8"    # Dégradé de l'en-tête : couleur d'arrivée
    Accent         = "#5B2C8F"    # Couleur principale (onglet actif, titres de cartes, bouton "Générer et Ouvrir")
    AccentHover    = "#6D37A8"    # Survol du bouton principal
    Success        = "#1E9E62"    # Bouton "Générer le Dashboard"
    SuccessHover   = "#23B571"
    FormBack       = "#F2F4F8"    # Fond de la fenêtre et des onglets
    CardBack       = "#FFFFFF"    # Fond des cartes (sections)
    CardBorder     = "#E1E6EF"    # Liseré des cartes et boutons secondaires
    GhostHover     = "#EBEFF6"    # Survol des boutons secondaires
    TextMain       = "#1B2430"    # Texte principal
    TextMuted      = "#5B6472"    # Texte secondaire (statut, notes)
    FontFamily     = "Segoe UI"
    CornerRadius   = 14           # Rayon d'arrondi des cartes (les boutons utilisent 10)
}

function ConvertTo-UIColor {
    param([string]$Hex)
    return [System.Drawing.ColorTranslator]::FromHtml($Hex)
}

function Get-RoundedPath {
    param([int]$Width, [int]$Height, [int]$Radius)
    $r = [int]([Math]::Max(2, [Math]::Min($Radius, [Math]::Min($Width, $Height) / 2))) * 2
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $path.AddArc(0, 0, $r, $r, 180, 90)
    $path.AddArc($Width - $r - 1, 0, $r, $r, 270, 90)
    $path.AddArc($Width - $r - 1, $Height - $r - 1, $r, $r, 0, 90)
    $path.AddArc(0, $Height - $r - 1, $r, $r, 90, 90)
    $path.CloseFigure()
    return $path
}

function Set-RoundedRegion {
    param($Control, [int]$Radius = 12)
    if ($Control.Width -le 4 -or $Control.Height -le 4) { return }
    $path = Get-RoundedPath -Width $Control.Width -Height $Control.Height -Radius $Radius
    $Control.Region = New-Object System.Drawing.Region($path)
}

function New-CardPanel {
    <# Carte blanche arrondie avec liseré et titre, façon dashboard. #>
    param([string]$Title, [int]$X, [int]$Y, [int]$W, [int]$H)
    $p = New-Object System.Windows.Forms.Panel
    $p.Location  = New-Object System.Drawing.Point($X, $Y)
    $p.Size      = New-Object System.Drawing.Size($W, $H)
    $p.BackColor = ConvertTo-UIColor $Theme.CardBack
    Set-RoundedRegion -Control $p -Radius $Theme.CornerRadius
    $p.Add_Paint({
        param($sender, $e)
        $e.Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
        $pen  = New-Object System.Drawing.Pen((ConvertTo-UIColor $Theme.CardBorder), 1)
        $path = Get-RoundedPath -Width $sender.Width -Height $sender.Height -Radius $Theme.CornerRadius
        $e.Graphics.DrawPath($pen, $path)
        $pen.Dispose(); $path.Dispose()
    })
    if ($Title) {
        $t = New-Object System.Windows.Forms.Label
        $t.Location  = New-Object System.Drawing.Point(20, 14)
        $t.Size      = New-Object System.Drawing.Size(($W - 40), 24)
        $t.Text      = $Title
        $t.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 10.5, [System.Drawing.FontStyle]::Bold)
        $t.ForeColor = ConvertTo-UIColor $Theme.Accent
        $t.BackColor = [System.Drawing.Color]::Transparent
        $p.Controls.Add($t)
    }
    return $p
}

function New-ThemedButton {
    <# Bouton plat arrondi avec effet de survol. -Ghost : variante claire avec liseré. #>
    param(
        [string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H,
        [string]$BackHex, [string]$HoverHex, [string]$ForeHex = "#FFFFFF",
        [single]$FontSize = 10, [switch]$Ghost
    )
    $b = New-Object System.Windows.Forms.Button
    $b.Location  = New-Object System.Drawing.Point($X, $Y)
    $b.Size      = New-Object System.Drawing.Size($W, $H)
    $b.Text      = $Text
    $b.FlatStyle = "Flat"
    $b.BackColor = ConvertTo-UIColor $BackHex
    $b.ForeColor = ConvertTo-UIColor $ForeHex
    $b.Font      = New-Object System.Drawing.Font($Theme.FontFamily, $FontSize, [System.Drawing.FontStyle]::Bold)
    $b.Cursor    = [System.Windows.Forms.Cursors]::Hand
    $b.FlatAppearance.BorderSize = 0
    $b.FlatAppearance.MouseOverBackColor = ConvertTo-UIColor $HoverHex
    $b.FlatAppearance.MouseDownBackColor = ConvertTo-UIColor $HoverHex
    if ($Ghost) {
        $b.FlatAppearance.BorderSize  = 1
        $b.FlatAppearance.BorderColor = ConvertTo-UIColor $Theme.CardBorder
    }
    Set-RoundedRegion -Control $b -Radius 10
    return $b
}

function Enable-ModernWindowCorners {
    <# Coins de fenêtre arrondis natifs (Windows 11). Ignoré silencieusement sur Windows 10. #>
    param($TargetForm)
    try {
        if (-not ("Win32.DwmApi" -as [type])) {
            Add-Type -Namespace Win32 -Name DwmApi -MemberDefinition '[DllImport("dwmapi.dll")] public static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);' -ErrorAction Stop
        }
        $pref = 2
        [void][Win32.DwmApi]::DwmSetWindowAttribute($TargetForm.Handle, 33, [ref]$pref, 4)
    } catch { }
}

# ===== CLÉ DE CHIFFREMENT AES (identique au script de base : compatibilité des .clientconfig) =====
$AESKey = @(
    0x4D, 0x79, 0x53, 0x65, 0x63, 0x72, 0x65, 0x74,
    0x4B, 0x65, 0x79, 0x31, 0x32, 0x33, 0x34, 0x35,
    0x36, 0x37, 0x38, 0x39, 0x30, 0x41, 0x42, 0x43,
    0x44, 0x45, 0x46, 0x47, 0x48, 0x49, 0x4A, 0x4B
)

# ===== PERMISSIONS GRAPH REQUISES (App Registration - Client Credentials) =====
# - DeviceManagementManagedDevices.Read.All   (managedDevices, deviceCompliancePolicyStates, detectedApps)
# - DeviceManagementConfiguration.Read.All    (deviceCompliancePolicies)
# - DeviceManagementApps.Read.All             (mobileApps / inventaire applicatif)
# - (Onglet Remédiation) les endpoints userExperienceAnalytics* sont couverts par
#   DeviceManagementManagedDevices.Read.All, mais l'Analyse des points de terminaison
#   doit être ACTIVÉE dans Intune pour renvoyer des données (sinon : disque + inactivité seulement).


# ========================================
# FONCTIONS DE CHIFFREMENT/DÉCHIFFREMENT (reprises telles quelles du script de base)
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
        Write-Log "Erreur de déchiffrement AES : $($_.Exception.Message)" -Level ERROR
        return ""
    }
}

# ========================================
# FONCTIONS DE GESTION DES CONFIGURATIONS (reprises du script de base)
# ========================================

function Get-ClientConfig {
    param([string]$ClientName)
    if ([string]::IsNullOrWhiteSpace($ClientName) -or $ClientName -eq "-- Sélectionnez un client --") { return $null }
    $SafeClientName = $ClientName -replace '[\\/:*?"<>|]', '_' -replace '\s+', '_'
    $ConfigFile = Join-Path $ConfigFolder "$SafeClientName.clientconfig"
    if (-not (Test-Path $ConfigFile)) {
        Write-Log "Fichier de configuration introuvable : $ConfigFile" -Level ERROR
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
        Write-Log "Erreur lors de la lecture de la configuration : $($_.Exception.Message)" -Level ERROR
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
        $lblStatus.Text = "Aucune configuration client trouvée"; $lblStatus.ForeColor = [System.Drawing.Color]::Orange
        return
    }
    foreach ($file in $configFiles) {
        try {
            $config = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
            [void]$cmbClients.Items.Add($config.ClientName)
        } catch { Write-Log "Erreur lecture config : $($file.Name)" -Level ERROR }
    }
    $cmbClients.SelectedIndex = 0; $cmbClients.Enabled = $true
    $lblStatus.Text = "$($configFiles.Count) configuration(s) client chargée(s)"
    $lblStatus.ForeColor = [System.Drawing.Color]::Green
}

# ========================================
# FONCTIONS UTILITAIRES
# ========================================

function Write-Log {
    <#
        Écrit un message à la fois dans la console ET dans $LogFile, avec horodatage.
        Objectif : permettre de diagnostiquer un blocage/une erreur en collant simplement
        le contenu du fichier, sans dépendre de captures d'écran de la console.
        Niveaux : INFO (défaut), WARN, ERROR, OK.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "OK")][string]$Level = "INFO"
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line      = "[$timestamp] [$Level] $Message"

    $color = switch ($Level) {
        "WARN"  { "Yellow" }
        "ERROR" { "Red" }
        "OK"    { "Green" }
        default { "Gray" }
    }
    Write-Host $line -ForegroundColor $color

    try {
        $folder = Split-Path -Path $LogFile -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }

        # Rotation simple : si le journal dépasse ~5 Mo, on l'archive avant de continuer
        if ((Test-Path $LogFile) -and ((Get-Item $LogFile).Length -gt 5MB)) {
            $archiveName = "dashboard-log_" + (Get-Date).ToString("yyyyMMdd_HHmmss") + ".txt"
            Move-Item -Path $LogFile -Destination (Join-Path $folder $archiveName) -Force -ErrorAction SilentlyContinue
        }

        Add-Content -Path $LogFile -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch {
        # Le logging ne doit jamais faire planter le script principal
    }
}

# ===== CRÉATION DES DOSSIERS =====
# NOTE : placé APRÈS la définition de Write-Log, sinon le tout premier lancement sur un
# poste neuf (dossiers absents) échouait sur un appel à une fonction pas encore définie.
if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null }
if (-not (Test-Path $ImportFolder)) { New-Item -ItemType Directory -Path $ImportFolder -Force | Out-Null }
if (-not (Test-Path $ConfigFolder)) {
    New-Item -ItemType Directory -Path $ConfigFolder -Force | Out-Null
    Write-Log "Dossier créé : $ConfigFolder - Veuillez générer des configurations client" -Level WARN
}

function Show-ErrorMessage([string]$message) {
    [System.Windows.Forms.MessageBox]::Show($message, "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
function Show-InfoMessage([string]$message) {
    [System.Windows.Forms.MessageBox]::Show($message, "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function Start-ResponsiveSleep {
    <#
        Attend $Seconds secondes SANS geler la fenêtre WinForms. Un simple Start-Sleep bloque
        le thread d'interface : Windows arrête de pomper les messages et affiche la fenêtre
        comme "Ne répond pas", ce qui donne l'impression que le script est planté/bouclé alors
        qu'il patiente simplement (cas typique lors d'un backoff de plusieurs dizaines de
        secondes après un 429). Cette fonction pompe régulièrement les messages Windows via
        Application::DoEvents() et affiche un compte à rebours dans $lblStatus.
    #>
    param(
        [Parameter(Mandatory = $true)][int]$Seconds,
        [string]$Message = "Patientez"
    )
    $endTime = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $endTime) {
        $remaining = [Math]::Max(0, [int][Math]::Ceiling(($endTime - (Get-Date)).TotalSeconds))
        try {
            if ($lblStatus) { $lblStatus.Text = "$Message ($remaining s restantes)..." }
            if ($form)      { [System.Windows.Forms.Application]::DoEvents() }
        } catch {}
        Start-Sleep -Milliseconds 200
        try { if ($form) { [System.Windows.Forms.Application]::DoEvents() } } catch {}
    }
}

function Get-GraphErrorDetail {
    <#
        Extrait le message d'erreur détaillé renvoyé par Microsoft Graph (corps JSON,
        souvent du type { "error": { "code": "...", "message": "..." } }) à partir d'un
        ErrorRecord PowerShell, quelle que soit la version (Windows PowerShell 5.1 / PS7).
        Sans cela, on n'obtient que "Response status code does not indicate success: 400 (Bad Request)"
        qui ne dit pas QUEL appel a échoué ni POURQUOI.
    #>
    param($ErrorRecord)
    $rawBody = $null
    try {
        if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            $rawBody = $ErrorRecord.ErrorDetails.Message
        } elseif ($ErrorRecord.Exception.Response) {
            $stream = $ErrorRecord.Exception.Response.GetResponseStream()
            if ($stream) {
                $reader  = New-Object System.IO.StreamReader($stream)
                $rawBody = $reader.ReadToEnd()
                $reader.Dispose()
            }
        }
    } catch {}

    if ($rawBody) {
        try {
            $json = $rawBody | ConvertFrom-Json
            if ($json.error.message) { return "$($json.error.code) : $($json.error.message)" }
        } catch {}
        return $rawBody
    }
    return $ErrorRecord.Exception.Message
}

function ConvertTo-HtmlSafe {
    param([string]$Text)
    if ($null -eq $Text) { return "" }
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
}

function Anonymize-DeviceData {
    param([Parameter(Mandatory = $true)]$DeviceList)
    return $DeviceList | ForEach-Object {
        [PSCustomObject]@{
            DeviceName        = "Poste-" + ([guid]::NewGuid().ToString().Substring(0, 8))
            UserPrincipalName = "User-"  + ([guid]::NewGuid().ToString().Substring(0, 8))
            OperatingSystem   = $_.OperatingSystem
            OSVersion         = $_.OSVersion
            ComplianceState   = $_.ComplianceState
            LastSyncDateTime  = $_.LastSyncDateTime
        }
    }
}

# ========================================
# HELPERS GRAPH API : PAGINATION + BATCH + THROTTLING + COLLECTE PARALLÈLE
# ========================================
#
# Trois principes tiennent toute cette section :
#   1. Le service dicte le rythme. Un HTTP 429 s'accompagne d'un en-tête Retry-After :
#      c'est la seule source fiable du délai à respecter. Le corps JSON de la réponse
#      contient parfois "RetryAfter": null, ce qui a longtemps laissé croire le contraire.
#   2. On ne repart jamais tous en même temps. Sans bruit aléatoire ("jitter"), les
#      requêtes limitées au même instant repartent au même instant et se re-limitent.
#   3. On n'accumule pas dans un tableau. "$tableau += $element" recopie l'intégralité du
#      tableau à chaque ajout : coût quadratique, très sensible au-delà de quelques
#      dizaines de milliers d'éléments. Partout ici : List[psobject].

function Get-HttpStatusCode {
    <# Code HTTP porté par une ErrorRecord, quel que soit le moteur (PS 5.1 : WebException ;
       PS 7 : HttpResponseMessage). Renvoie 0 si le code est indéterminable — panne réseau,
       DNS, coupure de proxy : autant de cas où il n'y a pas de réponse HTTP du tout. #>
    param($ErrorRecord)
    $code = 0
    try { $code = [int]$ErrorRecord.Exception.Response.StatusCode } catch { $code = 0 }
    if ($code -eq 0) { try { $code = [int]$ErrorRecord.Exception.StatusCode } catch { $code = 0 } }
    return $code
}

function Get-RetryAfterSeconds {
    <# Valeur de l'en-tête Retry-After, ou $null s'il est absent ou illisible. #>
    param($ErrorRecord)
    $raw = $null
    try {
        $responseHeaders = $ErrorRecord.Exception.Response.Headers
        if ($responseHeaders) {
            # PowerShell 5.1 : WebHeaderCollection, indexable par nom.
            try { $raw = $responseHeaders["Retry-After"] } catch { $raw = $null }
            if ([string]::IsNullOrWhiteSpace([string]$raw)) {
                # PowerShell 7 : HttpResponseHeaders, non indexable — il faut TryGetValues.
                $values = $null
                if ($responseHeaders.TryGetValues("Retry-After", [ref]$values)) { $raw = @($values)[0] }
            }
        }
    } catch { $raw = $null }
    if (-not [string]::IsNullOrWhiteSpace([string]$raw)) {
        $parsed = 0
        if ([int]::TryParse((([string]$raw).Trim()), [ref]$parsed) -and $parsed -gt 0) { return $parsed }
    }
    return $null
}

function Get-BackoffDelaySeconds {
    <# Délai avant nouvelle tentative : Retry-After s'il est fourni, sinon backoff
       exponentiel plafonné et bruité (tirage dans [fenêtre/2 ; fenêtre]). #>
    param($ErrorRecord, [int]$Attempt = 0, [int]$BaseSeconds = 5, [int]$CapSeconds = 90)
    $retryAfter = Get-RetryAfterSeconds -ErrorRecord $ErrorRecord
    if ($null -ne $retryAfter) { return [int][Math]::Min($CapSeconds, [Math]::Max(1, $retryAfter)) }
    $window   = [Math]::Min([double]$CapSeconds, $BaseSeconds * [Math]::Pow(2, $Attempt))
    $jittered = ($window / 2.0) + ((Get-Random -Minimum 0 -Maximum 1001) / 1000.0) * ($window / 2.0)
    return [int][Math]::Max(1, [Math]::Round($jittered))
}

function Test-RetryableStatus {
    <# 429 = throttling, 408 = délai dépassé, 5xx = incident côté service, 0 = pas de
       réponse HTTP (coupure réseau passagère). Tout le reste est une vraie erreur
       fonctionnelle (401, 403, 400...) qu'il serait absurde de rejouer. #>
    param([int]$StatusCode)
    return ($StatusCode -eq 429 -or $StatusCode -eq 408 -or $StatusCode -eq 0 -or $StatusCode -ge 500)
}

function Get-GraphPagedResults {
    <#
        Pagination Graph complète (suit @odata.nextLink jusqu'au bout).

        Par rapport à la version précédente :
          * la pause entre pages est ADAPTATIVE — nulle tant que le service répond
            normalement, elle n'apparaît qu'après un premier 429 puis se résorbe de
            moitié à chaque page réussie. L'ancienne pause fixe de 300 ms coûtait à elle
            seule une trentaine de secondes sur un parc de 100 000 appareils, sans rien
            apporter quand le tenant n'était pas limité ;
          * Retry-After est réellement lu et honoré ;
          * l'accumulation passe par une List[psobject].
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Url,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [int]$MaxRetries = 8,
        [scriptblock]$ProgressCallback,
        [string]$Label = ""
    )
    $results     = New-Object System.Collections.Generic.List[psobject]
    $authHeaders = @{ Authorization = "Bearer $AccessToken" }
    $labelPrefix = if ([string]::IsNullOrWhiteSpace($Label)) { "" } else { "$Label : " }
    $next        = $Url
    $pageNumber  = 0
    $coolDownMs  = 0

    while ($next) {
        $pageNumber++
        $attempt  = 0
        $response = $null
        while ($true) {
            try {
                $response = Invoke-RestMethod -Method GET -Uri $next -Headers $authHeaders -ErrorAction Stop
                break
            } catch {
                $status = Get-HttpStatusCode -ErrorRecord $_
                if ((Test-RetryableStatus -StatusCode $status) -and $attempt -lt $MaxRetries) {
                    $wait = Get-BackoffDelaySeconds -ErrorRecord $_ -Attempt $attempt -BaseSeconds 5 -CapSeconds 90
                    # Le service vient de nous freiner : on ralentit aussi les pages suivantes.
                    $coolDownMs = [int][Math]::Min(2000, [Math]::Max(400, $coolDownMs * 2))
                    if ($ProgressCallback) { & $ProgressCallback "${labelPrefix}throttling Graph (HTTP $status) - nouvelle tentative dans $wait s (essai $($attempt + 1)/$MaxRetries)..." }
                    Start-ResponsiveSleep -Seconds $wait -Message "Throttling Graph (HTTP $status) - essai $($attempt + 1)/$MaxRetries"
                    $attempt++
                } else {
                    $detail = Get-GraphErrorDetail -ErrorRecord $_
                    throw "Erreur Graph API (HTTP $status) sur `"$next`" : $detail"
                }
            }
        }

        if ($response.value) { foreach ($item in $response.value) { [void]$results.Add($item) } }
        $next = $response.'@odata.nextLink'

        if ($next) {
            if ($ProgressCallback -and ($pageNumber % 5 -eq 0)) {
                & $ProgressCallback "${labelPrefix}$($results.Count) élément(s) récupéré(s) (page $pageNumber)..."
            }
            if ($coolDownMs -gt 0) {
                Start-Sleep -Milliseconds $coolDownMs
                $coolDownMs = [int]($coolDownMs / 2)   # on relâche dès que le service redevient coopératif
            }
        }
    }
    return $results.ToArray()
}

function Invoke-GraphBatch {
    <#
        Envoie une série de requêtes Graph via l'endpoint $batch (v1.0 ou beta), par lots de
        20 (limite du service), avec gestion des 429/5xx à DEUX niveaux : sur l'appel batch
        lui-même, et sur chaque sous-réponse (un lot peut revenir en 200 avec certaines
        sous-requêtes en 429).

        $Requests : tableau de hashtables @{ id=...; method="GET"; url="/deviceManagement/..." }

        Corrections de performance par rapport à la version précédente :
          * les sous-requêtes limitées sont rejouées en boucle (jusqu'à $SubRequestRetries
            passes) au lieu d'une seule tentative, en honorant leur propre Retry-After ;
          * la fusion des résultats rejoués passe par une table id -> requête au lieu d'un
            "Where-Object" reconstruisant tout le tableau de réponses à chaque sous-requête
            (quadratique, et ruineux sur les gros lots) ;
          * la pause entre lots est adaptative, comme pour la pagination : nulle par défaut,
            elle n'apparaît qu'après un throttling avéré.
    #>
    param(
        # AllowEmptyCollection : sans lui, Mandatory refuse @() AVANT d'entrer dans la
        # fonction, et le garde-fou ci-dessous ne sert jamais. Or un lot vide est un cas
        # normal (aucune stratégie en échec, aucun profil à interroger) : il doit rendre
        # un tableau vide, pas faire échouer la génération.
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][array]$Requests,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [string]$GraphVersion = "beta",
        [int]$BatchSize = 20,
        [int]$MaxRetries = 8,
        [int]$PauseMs = 0,
        [int]$SubRequestRetries = 3,
        [scriptblock]$ProgressCallback,
        [string]$Label = ""
    )

    $allResponses = New-Object System.Collections.Generic.List[psobject]
    if ($null -eq $Requests -or $Requests.Count -eq 0) { return $allResponses.ToArray() }

    # Sécurité : l'endpoint $batch exige un "id" unique PAR LOT (400 BadRequest sinon).
    # On déduplique par précaution, quelle que soit la cause de doublons en amont
    # (appareil compté deux fois, application dupliquée...) : seule la première
    # occurrence de chaque id est conservée.
    $seenIds  = New-Object System.Collections.Generic.HashSet[string]
    $deduped  = New-Object System.Collections.Generic.List[psobject]
    $dupCount = 0
    foreach ($r in $Requests) {
        if ($seenIds.Add([string]$r.id)) { [void]$deduped.Add($r) } else { $dupCount++ }
    }
    if ($dupCount -gt 0) {
        Write-Log "Invoke-GraphBatch : $dupCount requête(s) en doublon ignorée(s) (id déjà présent)." -Level WARN
    }
    if ($deduped.Count -eq 0) { return $allResponses.ToArray() }

    $requestArray = $deduped.ToArray()
    $total        = $requestArray.Count
    $authHeaders  = @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" }
    $endpoint     = "https://graph.microsoft.com/$GraphVersion/`$batch"
    $labelPrefix  = if ([string]::IsNullOrWhiteSpace($Label)) { "" } else { "$Label : " }
    $coolDownMs   = $PauseMs

    for ($i = 0; $i -lt $total; $i += $BatchSize) {
        $endIndex = [Math]::Min($i + $BatchSize - 1, $total - 1)
        $chunk    = @($requestArray[$i..$endIndex])

        # Table id -> requête d'origine : permet de rejouer une sous-requête limitée sans
        # reparcourir le lot à chaque fois.
        $requestById = @{}
        foreach ($c in $chunk) { $requestById[[string]$c.id] = $c }

        $pending = $chunk
        $round   = 0
        while ($pending.Count -gt 0 -and $round -le $SubRequestRetries) {
            $body    = (@{ requests = @($pending) } | ConvertTo-Json -Depth 8)
            $attempt = 0
            $resp    = $null
            while ($true) {
                try {
                    $resp = Invoke-RestMethod -Method POST -Uri $endpoint -Headers $authHeaders -Body $body -ErrorAction Stop
                    break
                } catch {
                    $status = Get-HttpStatusCode -ErrorRecord $_
                    if ((Test-RetryableStatus -StatusCode $status) -and $attempt -lt $MaxRetries) {
                        $wait = Get-BackoffDelaySeconds -ErrorRecord $_ -Attempt $attempt -BaseSeconds 5 -CapSeconds 90
                        $coolDownMs = [int][Math]::Min(1500, [Math]::Max(300, $coolDownMs * 2))
                        if ($ProgressCallback) { & $ProgressCallback "${labelPrefix}throttling sur l'appel batch (HTTP $status) - nouvelle tentative dans $wait s (essai $($attempt + 1)/$MaxRetries)..." }
                        Start-ResponsiveSleep -Seconds $wait -Message "Throttling batch Graph (HTTP $status) - essai $($attempt + 1)/$MaxRetries"
                        $attempt++
                    } else {
                        $detail = Get-GraphErrorDetail -ErrorRecord $_
                        throw "Erreur Graph API batch (HTTP $status) : $detail"
                    }
                }
            }

            $retryNext = New-Object System.Collections.Generic.List[psobject]
            $waitHint  = 0
            foreach ($sub in @($resp.responses)) {
                $subStatus = 0
                try { $subStatus = [int]$sub.status } catch { $subStatus = 0 }
                if ($subStatus -eq 429 -and $round -lt $SubRequestRetries) {
                    # Une sous-réponse 429 porte son propre Retry-After dans ses en-têtes.
                    try {
                        if ($sub.headers -and $sub.headers.'Retry-After') {
                            $parsed = 0
                            if ([int]::TryParse([string]$sub.headers.'Retry-After', [ref]$parsed) -and $parsed -gt $waitHint) { $waitHint = $parsed }
                        }
                    } catch { }
                    $original = $requestById[[string]$sub.id]
                    if ($original) { [void]$retryNext.Add($original) } else { [void]$allResponses.Add($sub) }
                } else {
                    [void]$allResponses.Add($sub)
                }
            }

            $pending = @($retryNext)
            if ($pending.Count -gt 0) {
                $delay = if ($waitHint -gt 0) { [Math]::Min(60, $waitHint) } else { [Math]::Min(30, 3 * [Math]::Pow(2, $round)) }
                if ($ProgressCallback) { & $ProgressCallback "${labelPrefix}$($pending.Count) sous-requête(s) limitée(s) - nouvelle tentative dans $([int]$delay) s..." }
                Start-ResponsiveSleep -Seconds ([int]$delay) -Message "Nouvelle tentative pour $($pending.Count) requête(s) limitée(s)"
            }
            $round++
        }

        if ($ProgressCallback) { & $ProgressCallback "${labelPrefix}$([Math]::Min($endIndex + 1, $total)) / $total requête(s) traitée(s)..." }
        if ($coolDownMs -gt 0) {
            Start-Sleep -Milliseconds $coolDownMs
            $coolDownMs = [int]($coolDownMs / 2)
        }
    }

    return $allResponses.ToArray()
}

function Invoke-ParallelGraphCollections {
    <#
        Exécute PLUSIEURS collectes de listes Graph indépendantes EN PARALLÈLE, via un pool
        de runspaces.

        POURQUOI : les collectes de l'onglet Remédiation (scores Endpoint Analytics,
        performances de démarrage, BitLocker, fiabilité applicative, historique de
        démarrage, santé des batteries) ne dépendent pas les unes des autres. Les enchaîner
        en série revient à attendre la SOMME de leurs latences, alors que le service les
        sert très bien de front. C'est le principal poste de temps de la génération.

        Concurrence volontairement modérée (4 par défaut) : au-delà, c'est le throttling
        Graph qu'on déclenche, et le gain se retourne en perte.

        $Jobs   : table Nom -> URL de première page (la pagination est faite dans le runspace).
        Retour  : table Nom -> [PSCustomObject] @{ Items = @(...); Error = <message> | $null }

        Toute défaillance de l'infrastructure de parallélisme (hôte restreint, stratégie
        d'exécution, .NET partiel) bascule automatiquement en collecte séquentielle : cette
        optimisation ne peut jamais empêcher la génération du rapport.
    #>
    param(
        [Parameter(Mandatory = $true)][hashtable]$Jobs,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [int]$MaxConcurrency = 4,
        [scriptblock]$ProgressCallback
    )
    $result = @{}
    if ($Jobs.Count -eq 0) { return $result }

    # Corps exécuté dans chaque runspace. Volontairement AUTONOME : un runspace ne voit
    # aucune des fonctions du script hôte, tout doit tenir dans ce texte.
    $workerText = @'
param($JobName, $JobUrl, $Token, $MaxRetries)
$ProgressPreference = "SilentlyContinue"
$items = New-Object System.Collections.Generic.List[psobject]
$err   = $null
$next  = $JobUrl
try {
    while ($next) {
        $attempt  = 0
        $response = $null
        while ($true) {
            try {
                $response = Invoke-RestMethod -Method GET -Uri $next -Headers @{ Authorization = "Bearer $Token" } -ErrorAction Stop
                break
            } catch {
                $status = 0
                try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = 0 }
                if (($status -eq 429 -or $status -eq 408 -or $status -ge 500) -and $attempt -lt $MaxRetries) {
                    $wait = [Math]::Min(90, 5 * [Math]::Pow(2, $attempt))
                    try {
                        $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                        $parsed = 0
                        if ($retryAfter -and [int]::TryParse([string]$retryAfter, [ref]$parsed) -and $parsed -gt 0) { $wait = [Math]::Min(120, $parsed) }
                    } catch { }
                    # Bruit aléatoire : deux runspaces limités en même temps ne repartent pas ensemble.
                    $wait = ($wait / 2.0) + ((Get-Random -Minimum 0 -Maximum 1001) / 1000.0) * ($wait / 2.0)
                    Start-Sleep -Seconds ([int][Math]::Max(1, [Math]::Round($wait)))
                    $attempt++
                } else {
                    throw
                }
            }
        }
        if ($response.value) { foreach ($item in $response.value) { [void]$items.Add($item) } }
        $next = $response.'@odata.nextLink'
    }
} catch {
    $err = $_.Exception.Message
}
[PSCustomObject]@{ Name = $JobName; Items = $items.ToArray(); Error = $err }
'@

    $pool = $null
    try {
        $pool = [runspacefactory]::CreateRunspacePool(1, [Math]::Max(1, $MaxConcurrency))
        $pool.ApartmentState = [System.Threading.ApartmentState]::MTA
        $pool.Open()
    } catch {
        Write-Log "Collecte parallèle indisponible ($($_.Exception.Message)) : bascule en collecte séquentielle." -Level WARN
        if ($pool) { try { $pool.Dispose() } catch { } }
        foreach ($name in @($Jobs.Keys)) {
            try {
                if ($ProgressCallback) { & $ProgressCallback "Récupération : $name..." }
                $items = @(Get-GraphPagedResults -Url $Jobs[$name] -AccessToken $AccessToken -ProgressCallback $ProgressCallback -Label $name)
                $result[$name] = [PSCustomObject]@{ Items = $items.ToArray(); Error = $null }
            } catch {
                $result[$name] = [PSCustomObject]@{ Items = @(); Error = $_.Exception.Message }
            }
        }
        return $result
    }

    $running = New-Object System.Collections.Generic.List[psobject]
    foreach ($name in @($Jobs.Keys)) {
        $shell = [powershell]::Create()
        $shell.RunspacePool = $pool
        [void]$shell.AddScript($workerText).AddArgument($name).AddArgument($Jobs[$name]).AddArgument($AccessToken).AddArgument(8)
        [void]$running.Add([PSCustomObject]@{ Name = $name; Shell = $shell; Handle = $shell.BeginInvoke() })
    }

    $total = $running.Count
    Write-Log "Collecte parallèle démarrée : $total jeu(x) de données, concurrence $MaxConcurrency." -Level INFO
    $lastReported = -1
    while ($true) {
        $done = @($running | Where-Object { $_.Handle.IsCompleted }).Count
        if ($done -ge $total) { break }
        if ($ProgressCallback -and $done -ne $lastReported) {
            & $ProgressCallback "Collecte parallèle : $done / $total jeu(x) de données terminé(s)..."
            $lastReported = $done
        }
        # Garde la fenêtre de l'outil réactive pendant l'attente (elle n'est pas figée).
        # Sous try/catch : cette boucle tourne toutes les 200 ms, et l'assembly WinForms
        # peut ne pas être chargée (exécution sans IHM, hôte restreint). Une collecte de
        # plusieurs minutes ne doit pas échouer pour un simple rafraîchissement d'écran.
        try { [System.Windows.Forms.Application]::DoEvents() } catch { }
        Start-Sleep -Milliseconds 200
    }

    foreach ($r in $running) {
        try {
            $output  = $r.Shell.EndInvoke($r.Handle)
            $payload = @($output) | Where-Object { $_ -and $_.PSObject.Properties['Items'] } | Select-Object -First 1
            if ($payload) {
                $result[$r.Name] = [PSCustomObject]@{ Items = @($payload.Items); Error = $payload.Error }
            } else {
                $result[$r.Name] = [PSCustomObject]@{ Items = @(); Error = "Aucune donnée renvoyée par le runspace." }
            }
        } catch {
            $result[$r.Name] = [PSCustomObject]@{ Items = @(); Error = $_.Exception.Message }
        } finally {
            try { $r.Shell.Dispose() } catch { }
        }
    }
    try { $pool.Close(); $pool.Dispose() } catch { }

    foreach ($name in @($result.Keys)) {
        if ($result[$name].Error) {
            Write-Log "Collecte '$name' incomplète : $($result[$name].Error)" -Level WARN
        } else {
            Write-Log "Collecte '$name' : $(@($result[$name].Items).Count) élément(s)." -Level INFO
        }
    }
    return $result
}

function ConvertFrom-IntuneReportPayload {
    <#
        Les exports de rapports Intune (deviceManagement/reports/exportJobs) reviennent sous
        DEUX formes selon le rapport : soit un tableau d'objets directement exploitable, soit
        un objet { Schema:[{Column,PropertyType}], Values:[[...],[...]] } où chaque ligne est
        un tableau POSITIONNEL à recoller aux noms de colonnes. On normalise les deux ici,
        pour que le reste du script ne voie jamais que des objets à propriétés nommées.
    #>
    param($Payload)
    if ($null -eq $Payload) { return @() }
    if ($Payload.PSObject.Properties['Schema'] -and $Payload.PSObject.Properties['Values']) {
        $columns = @($Payload.Schema | ForEach-Object { [string]$_.Column })
        $rows    = New-Object System.Collections.Generic.List[psobject]
        foreach ($line in @($Payload.Values)) {
            $cells  = @($line)
            $record = [ordered]@{}
            for ($i = 0; $i -lt $columns.Count; $i++) {
                $record[$columns[$i]] = if ($i -lt $cells.Count) { $cells[$i] } else { $null }
            }
            [void]$rows.Add([PSCustomObject]$record)
        }
        return $rows.ToArray()
    }
    return @($Payload)
}

# ========================================
# PAGE 1 - CONFORMITÉ & GRACE PERIOD
# ========================================

# Table de correspondance "nom technique du setting Graph" -> "motif lisible"
# Basée sur les propriétés documentées de windows10CompliancePolicy (Microsoft Graph).
# Le matching se fait par sous-chaîne (regex, insensible à la casse) car le champ
# "setting"/"settingName" renvoyé par Graph inclut généralement ces noms de propriété.
$ComplianceReasonMap = [ordered]@{
    "bitLockerEnabled"                          = "BitLocker désactivé / non conforme"
    "storageRequireEncryption"                  = "Chiffrement du stockage désactivé"
    "secureBootEnabled"                         = "Secure Boot désactivé"
    "codeIntegrityEnabled"                      = "Code Integrity désactivé"
    "earlyLaunchAntiMalwareDriverEnabled"        = "Early Launch Anti-Malware désactivé"
    "requireHealthyDeviceReport"                = "Attestation d'intégrité (Device Health) échouée"
    "osMinimumVersion"                          = "Version d'OS trop ancienne"
    "osMaximumVersion"                          = "Version d'OS non autorisée (trop récente)"
    "mobileOsMinimumVersion"                    = "Version d'OS mobile trop ancienne"
    "mobileOsMaximumVersion"                    = "Version d'OS mobile non autorisée"
    "passwordRequiredType"                      = "Type de mot de passe non conforme"
    "passwordMinimumLength"                     = "Mot de passe trop court"
    "passwordRequired"                          = "Mot de passe requis manquant"
    "activeFirewallRequired"                    = "Pare-feu non actif"
    "defenderEnabled"                           = "Microsoft Defender désactivé"
    "defenderVersion"                           = "Version Microsoft Defender non conforme"
    "signatureOutOfDate"                        = "Signatures antivirus obsolètes"
    "rtpEnabled"                                = "Protection en temps réel (RTP) désactivée"
    "antivirusRequired"                         = "Antivirus non détecté / non actif"
    "antiSpywareRequired"                       = "Anti-spyware non détecté / non actif"
    "deviceThreatProtectionRequiredSecurityLevel" = "Niveau de menace au-dessus du seuil autorisé"
    "deviceThreatProtectionEnabled"              = "Protection contre les menaces désactivée"
    "tpmRequired"                                = "TPM manquant ou non conforme"
    "memoryIntegrityEnabled"                    = "Memory Integrity (HVCI) désactivé"
    "kernelDmaProtectionEnabled"                 = "Protection DMA du noyau désactivée"
    "virtualizationBasedSecurityEnabled"         = "Virtualization Based Security désactivée"
}

function Get-ComplianceCategoryLabel {
    param([string]$RawSettingName)
    if ([string]::IsNullOrWhiteSpace($RawSettingName)) { return $null }
    foreach ($key in $ComplianceReasonMap.Keys) {
        if ($RawSettingName -match [regex]::Escape($key)) { return $ComplianceReasonMap[$key] }
    }
    return "Autre motif : $RawSettingName"
}

function Get-ComplianceBreakdown {
    <#
        Pour une liste de managedDevices (déjà filtrés sur Non-compliant / In Grace Period),
        récupère - via des appels Graph $batch - les policyStates puis les settingStates
        en échec, et regroupe les appareils par motif de non-conformité.
        Retourne un [ordered] hashtable : "Motif" -> @{ Count = n; Devices = @(...) }
    #>
    param(
        [Parameter(Mandatory = $true)][array]$Devices,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [scriptblock]$ProgressCallback
    )

    $result = [ordered]@{}
    if ($Devices.Count -eq 0) { return $result }

    # Un même appareil ne doit être traité qu'une fois (co-gestion, doublons de pagination, etc.)
    # sans quoi il serait compté deux fois dans les statistiques par motif.
    $Devices = @($Devices | Sort-Object -Property Id -Unique)

    if ($ProgressCallback) { & $ProgressCallback "Récupération des états de conformité par appareil ($($Devices.Count) appareils)..." }

    # Étape 1 : policyStates par appareil
    $policyStateRequests = @()
    foreach ($d in $Devices) {
        $policyStateRequests += @{
            id     = $d.Id
            method = "GET"
            url    = "/deviceManagement/managedDevices/$($d.Id)/deviceCompliancePolicyStates"
        }
    }
    $policyStateResponses = Invoke-GraphBatch -Requests $policyStateRequests -AccessToken $AccessToken -ProgressCallback $ProgressCallback

    if ($ProgressCallback) { & $ProgressCallback "Analyse des règles de conformité en échec..." }

    # Étape 2 : pour chaque policyState en échec, on va chercher le détail des settingStates
    $settingRequests = @()
    $badPolicyByDevice = @{}
    foreach ($resp in $policyStateResponses) {
        $devId = $resp.id
        if (-not $resp.body -or -not $resp.body.value) { continue }
        $bad = $resp.body.value | Where-Object { $_.state -in @("nonCompliant", "error", "conflict") }
        if (-not $bad -or $bad.Count -eq 0) { continue }
        $badPolicyByDevice[$devId] = $bad
        foreach ($ps in $bad) {
            $settingRequests += @{
                id     = "$devId--$($ps.id)"
                method = "GET"
                url    = "/deviceManagement/managedDevices/$devId/deviceCompliancePolicyStates/$($ps.id)/settingStates"
            }
        }
    }

    $settingResponses = @()
    if ($settingRequests.Count -gt 0) {
        if ($ProgressCallback) { & $ProgressCallback "Récupération du détail des règles en échec ($($settingRequests.Count) requêtes)..." }
        $settingResponses = Invoke-GraphBatch -Requests $settingRequests -AccessToken $AccessToken -ProgressCallback $ProgressCallback
    }

    $settingsByDevice = @{}
    foreach ($resp in $settingResponses) {
        if (-not $resp.body -or -not $resp.body.value) { continue }
        $devId = ($resp.id -split "--")[0]
        if (-not $settingsByDevice.ContainsKey($devId)) { $settingsByDevice[$devId] = @() }
        $failing = $resp.body.value | Where-Object { $_.state -in @("nonCompliant", "error", "conflict") }
        if ($failing) { $settingsByDevice[$devId] += $failing }
    }

    # Étape 3 : regroupement par motif
    foreach ($d in $Devices) {
        $devId   = $d.Id
        $reasons = @()

        if ($settingsByDevice.ContainsKey($devId) -and $settingsByDevice[$devId].Count -gt 0) {
            foreach ($s in $settingsByDevice[$devId]) {
                $rawName = $s.setting
                if ([string]::IsNullOrWhiteSpace($rawName)) { $rawName = $s.settingName }
                $reasons += (Get-ComplianceCategoryLabel -RawSettingName $rawName)
            }
        } elseif ($badPolicyByDevice.ContainsKey($devId)) {
            foreach ($ps in $badPolicyByDevice[$devId]) {
                $reasons += "Politique non conforme : $($ps.displayName)"
            }
        } else {
            $reasons += "Motif non détaillé (aucune règle en échec retournée par l'API)"
        }

        $reasons = $reasons | Select-Object -Unique
        foreach ($r in $reasons) {
            if (-not $result.Contains($r)) { $result[$r] = @{ Count = 0; Devices = @() } }
            $result[$r].Count++
            $result[$r].Devices += [PSCustomObject]@{
                DeviceName        = $d.DeviceName
                UserPrincipalName = $d.UserPrincipalName
                OperatingSystem   = $d.OperatingSystem
                OSVersion         = $d.OSVersion
                ComplianceState   = $d.ComplianceState
                LastSyncDateTime  = $d.LastSyncDateTime
            }
        }
    }

    return $result
}

# ========================================
# COLLECTE PAR IMPORT DE FICHIERS (MODE HORS LIGNE)
# ========================================
#
# OBJECTIF : permettre de produire exactement le même dashboard SANS appeler
# Microsoft Graph, à partir de fichiers exportés (portail Intune, requête Graph
# Explorer, export réalisé par ce script lui-même, ou fichier maintenu à la main).
#
# TROIS MODES DE COLLECTE (onglet "Source des données") :
#   API      -> comportement historique : tout est récupéré via Microsoft Graph.
#   IMPORT   -> aucune connexion au tenant : tout provient des fichiers désignés.
#   HYBRIDE  -> connexion à Graph, mais chaque jeu de données pour lequel un
#               fichier est renseigné est lu depuis ce fichier au lieu de l'API.
#
# FORMATS ACCEPTÉS : .csv / .txt (séparateur ';', ',' ou tabulation, détecté
# automatiquement), et .json (tableau d'objets OU réponse Graph brute { "value": [...] }).
#
# NOMS DE COLONNES : la reconnaissance est tolérante (casse, accents, espaces et
# ponctuation ignorés) et accepte les libellés anglais du portail comme les
# libellés français, ainsi que les noms de propriétés Graph. Exemples acceptés
# pour un même champ : "Device name", "deviceName", "Nom de l'appareil", "Hostname".

function Get-NormalizedHeader {
    <# Normalise un libellé de colonne / une valeur : minuscules, sans accents, sans ponctuation. #>
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    $decomposed = $Name.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $decomposed.ToCharArray()) {
        if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return (($sb.ToString().ToLower()) -replace '[^a-z0-9]', '')
}

function Get-ImportFileEncodingName {
    <# Détecte l'encodage via le BOM (les exports Intune sont UTF-8 ; certains outils produisent de l'UTF-16). #>
    param([string]$Path)
    try {
        $reader = New-Object System.IO.StreamReader($Path, [System.Text.Encoding]::UTF8, $true)
        [void]$reader.Peek()
        $cp = $reader.CurrentEncoding.CodePage
        $reader.Close()
        switch ($cp) {
            1200  { return "Unicode" }
            1201  { return "BigEndianUnicode" }
            65001 { return "UTF8" }
            default { return "Default" }
        }
    } catch { return "UTF8" }
}

function Get-ImportDelimiter {
    <# Détermine le séparateur d'un CSV en comparant les occurrences dans la ligne d'en-tête. #>
    param([string]$Path, [string]$EncodingName = "UTF8")
    $header = ""
    try {
        $reader = New-Object System.IO.StreamReader($Path, [System.Text.Encoding]::UTF8, $true)
        $header = $reader.ReadLine()
        $reader.Close()
    } catch { }
    if ([string]::IsNullOrWhiteSpace($header)) { return "," }

    $counts = @{
        ';'    = ([regex]::Matches($header, ';')).Count
        ','    = ([regex]::Matches($header, ',')).Count
        "`t"   = ([regex]::Matches($header, "`t")).Count
    }
    $best  = ","
    $bestN = 0
    foreach ($k in $counts.Keys) {
        if ($counts[$k] -gt $bestN) { $bestN = $counts[$k]; $best = $k }
    }
    if ($bestN -eq 0) { return "," }
    return $best
}

function Get-ImportRows {
    <#
        Lit un fichier d'import (CSV/TSV/JSON) et renvoie un tableau d'objets.
        Le format est déduit de l'extension ; le séparateur et l'encodage sont détectés.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [string]$Label = "fichier"
    )

    if (-not (Test-Path $Path)) { throw "Fichier d'import introuvable ($Label) : $Path" }

    $ext  = [System.IO.Path]::GetExtension($Path).ToLower()
    $rows = @()

    if ($ext -eq ".json") {
        $raw = Get-Content -Path $Path -Raw -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace($raw)) { return @() }
        $obj = $raw | ConvertFrom-Json
        if ($obj -is [System.Array]) {
            $rows = @($obj)
        } elseif ($obj -and ($obj.PSObject.Properties.Name -contains 'value')) {
            # Réponse Microsoft Graph brute : { "@odata.context": ..., "value": [ ... ] }
            $rows = @($obj.value)
        } else {
            $rows = @($obj)
        }
        # Même piège qu'avec Invoke-RestMethod : ConvertFrom-Json convertit automatiquement
        # les dates ISO 8601 du JSON en [datetime] .NET. On les fige tout de suite en ISO
        # texte non ambigu, avant que Get-MappedValue ne les caste en [string] plus loin.
        Repair-GraphDateTimeFields -Items $rows -FieldNames @(
            'lastSyncDateTime', 'lastReportedDateTime', 'signatureUpdateDateTime',
            'startupDateTime', 'eventDateTime', 'createdDateTime'
        )
    } else {
        $encName = Get-ImportFileEncodingName -Path $Path
        $delim   = Get-ImportDelimiter -Path $Path -EncodingName $encName
        $shown   = if ($delim -eq "`t") { "TAB" } else { $delim }
        $rows    = @(Import-Csv -Path $Path -Delimiter $delim -Encoding $encName -ErrorAction Stop)
        Write-Log "Import ($Label) : séparateur '$shown', encodage $encName." -Level INFO
    }

    Write-Log "Import ($Label) : $($rows.Count) ligne(s) lue(s) depuis $Path" -Level OK
    return $rows
}

function New-ColumnMap {
    <# Construit la table "nom de colonne normalisé -> nom réel", à partir de la 1re ligne. #>
    param($Rows)
    $map = @{}
    if (-not $Rows -or @($Rows).Count -eq 0) { return $map }
    foreach ($p in (@($Rows)[0]).PSObject.Properties) {
        $k = Get-NormalizedHeader $p.Name
        if ($k -and -not $map.ContainsKey($k)) { $map[$k] = $p.Name }
    }
    return $map
}

function Get-MappedValue {
    <#
        Renvoie la valeur de la première colonne correspondant à l'un des alias fournis.
        1re passe : correspondance exacte (après normalisation). 2e passe : correspondance
        partielle, réservée aux alias suffisamment longs pour rester non ambigus
        (ex. "compliance" retrouve "Compliance state", mais "os" ne capture pas "OS version").
    #>
    param($Row, $Map, [string[]]$Names, $Default = $null)

    foreach ($n in $Names) {
        $k = Get-NormalizedHeader $n
        if ($k -and $Map.ContainsKey($k)) {
            $v = $Row.($Map[$k])
            if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) { return $v }
        }
    }
    foreach ($n in $Names) {
        $k = Get-NormalizedHeader $n
        if (-not $k -or $k.Length -lt 5) { continue }
        foreach ($key in @($Map.Keys)) {
            if ($key -like "*$k*") {
                $v = $Row.($Map[$key])
                if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) { return $v }
            }
        }
    }
    return $Default
}

function ConvertTo-ComplianceStateCode {
    <#
        Ramène un état de conformité exporté (FR ou EN, portail ou Graph) vers le code
        interne utilisé par le script : compliant / noncompliant / inGracePeriod / ...
    #>
    param($Value)
    $n = Get-NormalizedHeader ([string]$Value)
    if ([string]::IsNullOrWhiteSpace($n)) { return "unknown" }
    if ($n -match 'grace')                                  { return "inGracePeriod" }
    if ($n -match 'nonconform|noncompliant|notcompliant')   { return "noncompliant" }
    if ($n -match 'conflit|conflict')                       { return "conflict" }
    if ($n -match 'erreur|error')                           { return "error" }
    if ($n -match 'nonapplicable|notapplicable')            { return "notApplicable" }
    if ($n -match 'conforme|compliant')                     { return "compliant" }
    if ($n -match 'inconnu|unknown')                        { return "unknown" }
    return [string]$Value
}

function ConvertTo-ImportInt {
    <# Convertit "1 234", "1,234" ou "12 postes" en entier ; 0 si rien d'exploitable. #>
    param($Value)
    $s = [string]$Value
    if ([string]::IsNullOrWhiteSpace($s)) { return 0 }
    $digits = ($s -replace '[^0-9]', '')
    if ([string]::IsNullOrWhiteSpace($digits)) { return 0 }
    try { return [int]$digits } catch { return 0 }
}

function Import-ManagedDevicesFile {
    <#
        Import des appareils gérés. Alimente la Page 1 (conformité) exactement comme
        l'appel Graph /deviceManagement/managedDevices : les objets renvoyés portent
        les mêmes noms de propriétés, le reste du script est donc inchangé.

        Colonnes reconnues (alias les plus courants) :
          Nom du poste  : Device name / deviceName / Nom de l'appareil / Hostname
          Utilisateur   : Primary UPN / userPrincipalName / Utilisateur principal
          OS            : OS / operatingSystem / Système d'exploitation
          Version OS    : OS version / osVersion / Version du système
          Conformité    : Compliance / complianceState / État de conformité
          Synchro       : Last check-in / lastSyncDateTime / Dernière synchronisation
          Fabricant     : Manufacturer / Fabricant      Modèle : Model / Modèle
          Identifiant   : Device ID / id / Intune device ID
          Motif (option): Motif / Reason / Non-compliance reason  -> utilisé si aucun
                          fichier de détail de conformité n'est fourni.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $rows = Get-ImportRows -Path $Path -Label "appareils"
    if (@($rows).Count -eq 0) { return @() }
    $map = New-ColumnMap -Rows $rows

    $list = New-Object System.Collections.Generic.List[psobject]
    $i = 0
    foreach ($r in $rows) {
        $i++
        $name = [string](Get-MappedValue $r $map @('deviceName','device name','nom de l''appareil','nom du peripherique','hostname','computername','machine','name'))
        $id   = [string](Get-MappedValue $r $map @('id','deviceId','intuneDeviceId','device id','intune device id','managedDeviceId','azureADDeviceId'))
        if ([string]::IsNullOrWhiteSpace($name) -and [string]::IsNullOrWhiteSpace($id)) { continue }
        if ([string]::IsNullOrWhiteSpace($id)) { $id = "IMPORT-{0:D6}" -f $i }

        $displayName = $name
        if ([string]::IsNullOrWhiteSpace($displayName)) { $displayName = $id }

        $list.Add([PSCustomObject]@{
            id                = $id
            deviceName        = $displayName
            userPrincipalName = [string](Get-MappedValue $r $map @('userPrincipalName','user principal name','primaryUserUPN','primary user upn','primaryUPN','primary upn','userUPN','upn','utilisateur principal','utilisateur','email','adresse de messagerie'))
            operatingSystem   = [string](Get-MappedValue $r $map @('operatingSystem','systeme d''exploitation','plateforme','platform','os'))
            osVersion         = [string](Get-MappedValue $r $map @('osVersion','os version','version du systeme d''exploitation','version de l''os','version du systeme'))
            complianceState   = ConvertTo-ComplianceStateCode (Get-MappedValue $r $map @('complianceState','compliance state','etat de conformite','conformite','compliance','statut de conformite'))
            lastSyncDateTime  = [string](Get-MappedValue $r $map @('lastSyncDateTime','last check-in','last check in','derniere synchronisation','derniere connexion','last contact','last sync'))
            manufacturer      = [string](Get-MappedValue $r $map @('manufacturer','fabricant','constructeur','oem'))
            model             = [string](Get-MappedValue $r $map @('model','modele'))
            # Stockage : octets (Graph) ou Mo (export du portail) — l'unité est détectée.
            freeStorageSpaceInBytes  = ConvertTo-StorageBytes (Get-MappedValue $r $map @('freeStorageSpaceInBytes','free storage','espace libre','free disk space','stockage disponible','storage free'))
            totalStorageSpaceInBytes = ConvertTo-StorageBytes (Get-MappedValue $r $map @('totalStorageSpaceInBytes','total storage','espace total','total disk space','capacite de stockage','storage total'))
            complianceReason  = [string](Get-MappedValue $r $map @('motif','reason','non-compliance reason','noncompliancereason','motif de non-conformite','setting','parametre en echec'))
        })
    }

    Write-Log "Import appareils : $($list.Count) appareil(s) normalisé(s)." -Level OK
    return $list.ToArray()
}

function Import-ComplianceDetailFile {
    <#
        Import du détail de non-conformité (une ligne par couple appareil / règle en échec).
        Renvoie une table de hachage : clé d'appareil normalisée -> tableau de motifs lisibles.

        Colonnes reconnues :
          Appareil : Device name / deviceName / Device ID  (les deux clés sont indexées)
          Motif    : Motif / Reason            -> repris tel quel dans le rapport
          Règle    : Setting / Setting name / Paramètre -> traduit via la table des motifs
          Stratégie: Policy name / Stratégie   -> utilisé si aucune règle n'est précisée
          État     : State / Statut (facultatif) -> les lignes conformes sont ignorées
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $byDevice = @{}
    $rows = Get-ImportRows -Path $Path -Label "détail de conformité"
    if (@($rows).Count -eq 0) { return $byDevice }
    $map = New-ColumnMap -Rows $rows

    $kept = 0
    foreach ($r in $rows) {
        $state = [string](Get-MappedValue $r $map @('state','settingState','etat','statut','status'))
        if (-not [string]::IsNullOrWhiteSpace($state)) {
            $code = ConvertTo-ComplianceStateCode $state
            if ($code -eq "compliant" -or $code -eq "notApplicable" -or $code -eq "unknown") { continue }
        }

        $reason  = [string](Get-MappedValue $r $map @('motif','reason','motif de non-conformite'))
        $setting = [string](Get-MappedValue $r $map @('setting','settingName','setting name','parametre','nom du parametre','regle','rule'))
        $policy  = [string](Get-MappedValue $r $map @('policyName','policy name','displayName','strategie','politique','nom de la strategie','compliance policy'))

        $label = $null
        if (-not [string]::IsNullOrWhiteSpace($reason)) {
            # Motif déjà lisible (fichier produit par ce script ou saisi à la main) : repris tel quel.
            $label = $reason.Trim()
        } elseif (-not [string]::IsNullOrWhiteSpace($setting)) {
            $label = Get-ComplianceCategoryLabel -RawSettingName $setting
        } elseif (-not [string]::IsNullOrWhiteSpace($policy)) {
            $label = "Politique non conforme : $policy"
        }
        if ([string]::IsNullOrWhiteSpace($label)) { continue }

        $keys = @()
        $dn = [string](Get-MappedValue $r $map @('deviceName','device name','nom de l''appareil','hostname','computername','machine'))
        $di = [string](Get-MappedValue $r $map @('deviceId','device id','id','intuneDeviceId','managedDeviceId'))
        if (-not [string]::IsNullOrWhiteSpace($dn)) { $keys += (Get-NormalizedHeader $dn) }
        if (-not [string]::IsNullOrWhiteSpace($di)) { $keys += (Get-NormalizedHeader $di) }
        if ($keys.Count -eq 0) { continue }

        foreach ($k in $keys) {
            if (-not $byDevice.ContainsKey($k)) { $byDevice[$k] = New-Object System.Collections.Generic.List[string] }
            if (-not $byDevice[$k].Contains($label)) { $byDevice[$k].Add($label) }
        }
        $kept++
    }

    Write-Log "Import détail de conformité : $kept ligne(s) retenue(s) pour $($byDevice.Count) clé(s) d'appareil." -Level OK
    return $byDevice
}

function Add-ImportedComplianceReasons {
    <#
        Rattache à chaque appareil les motifs issus du fichier de détail.
        IMPORTANT : appelé AVANT une éventuelle anonymisation, puisque la correspondance
        se fait sur le nom réel du poste.
    #>
    param([array]$Devices, $ReasonsByDevice)
    if (-not $Devices -or $Devices.Count -eq 0) { return }
    foreach ($d in $Devices) {
        $found = New-Object System.Collections.Generic.List[string]
        if ($ReasonsByDevice) {
            foreach ($raw in @($d.Id, $d.DeviceName)) {
                if ([string]::IsNullOrWhiteSpace([string]$raw)) { continue }
                $k = Get-NormalizedHeader ([string]$raw)
                if ($k -and $ReasonsByDevice.ContainsKey($k)) {
                    foreach ($lbl in $ReasonsByDevice[$k]) { if (-not $found.Contains($lbl)) { $found.Add($lbl) } }
                }
            }
        }
        if ($found.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$d.ComplianceReason)) {
            $found.Add((Get-ComplianceCategoryLabel -RawSettingName ([string]$d.ComplianceReason)))
        }
        $d | Add-Member -NotePropertyName ImportReasons -NotePropertyValue ($found.ToArray()) -Force
    }
}

function Get-ComplianceBreakdownFromImport {
    <#
        Équivalent hors ligne de Get-ComplianceBreakdown : regroupe les appareils par motif
        à partir des motifs déjà rattachés par Add-ImportedComplianceReasons.
        Renvoie strictement la même structure ("Motif" -> @{ Count; Devices }).
    #>
    param([array]$Devices)

    $result = [ordered]@{}
    if (-not $Devices -or $Devices.Count -eq 0) { return $result }

    foreach ($d in $Devices) {
        $reasons = @($d.ImportReasons) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        if (-not $reasons -or $reasons.Count -eq 0) {
            $reasons = @("Motif non détaillé (aucun fichier de détail de conformité fourni)")
        }
        foreach ($r in ($reasons | Select-Object -Unique)) {
            if (-not $result.Contains($r)) { $result[$r] = @{ Count = 0; Devices = @() } }
            $result[$r].Count++
            $result[$r].Devices += [PSCustomObject]@{
                DeviceName        = $d.DeviceName
                UserPrincipalName = $d.UserPrincipalName
                OperatingSystem   = $d.OperatingSystem
                OSVersion         = $d.OSVersion
                ComplianceState   = $d.ComplianceState
                LastSyncDateTime  = $d.LastSyncDateTime
            }
        }
    }
    return $result
}

function Test-DevicesHaveGraphIds {
    <#
        Vrai si les identifiants des appareils importés sont de vrais GUID Intune : dans ce
        cas (mode HYBRIDE), le détail des règles en échec reste récupérable via l'API même
        si la liste des appareils provient d'un fichier.
    #>
    param([array]$Devices)
    if (-not $Devices -or $Devices.Count -eq 0) { return $false }
    $sample = @($Devices | Select-Object -First 20)
    foreach ($d in $sample) {
        if ([string]$d.Id -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { return $false }
    }
    return $true
}

function Import-DiscoveredAppsFile {
    <#
        Import des applications découvertes. Deux formats sont acceptés et détectés seuls :
          a) une ligne par application (colonne "Devices"/"Nombre d'appareils") ;
          b) une ligne par couple application/poste (export "Discovered apps" d'un appareil,
             ou export produit par ce script) : le nombre de postes est alors calculé et
             le détail par poste est reconstitué.
        Renvoie @{ Apps = @(...); DevicesByApp = @{ id -> @(postes) } }.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $out  = [PSCustomObject]@{ Apps = @(); DevicesByApp = @{} }
    $rows = Get-ImportRows -Path $Path -Label "applications découvertes"
    if (@($rows).Count -eq 0) { return $out }
    $map = New-ColumnMap -Rows $rows

    $appsByKey = [ordered]@{}
    $devsByKey = @{}
    $usedIds   = New-Object System.Collections.Generic.HashSet[string]
    $seq       = 0

    foreach ($r in $rows) {
        $name = [string](Get-MappedValue $r $map @('applicationName','application name','nom de l''application','displayName','appName','application','logiciel','name'))
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $version = [string](Get-MappedValue $r $map @('version','appVersion','app version','version de l''application'))
        $key     = ($name.Trim().ToLower() + '|' + $version.Trim().ToLower())

        if (-not $appsByKey.Contains($key)) {
            $seq++
            $id = [string](Get-MappedValue $r $map @('appId','app id','detectedAppId','id'))
            if ([string]::IsNullOrWhiteSpace($id) -or $usedIds.Contains($id)) { $id = "imp-app-{0:D5}" -f $seq }
            [void]$usedIds.Add($id)

            $appsByKey[$key] = [PSCustomObject]@{
                id            = $id
                displayName   = $name.Trim()
                version       = $version.Trim()
                publisher     = [string](Get-MappedValue $r $map @('publisher','editeur','fabricant','vendor'))
                platform      = [string](Get-MappedValue $r $map @('platform','plateforme','operatingSystem','os'))
                deviceCount   = 0
                MatchKey      = $key
                CountExplicit = $false
            }
            $devsByKey[$key] = New-Object System.Collections.Generic.List[psobject]
        }

        $app = $appsByKey[$key]

        # Nombre de postes fourni explicitement par le fichier (format a)
        $rawCount = Get-MappedValue $r $map @('deviceCount','device count','nombre d''appareils','nombre de postes','devices','appareils','postes','install count','installations')
        if ($null -ne $rawCount) {
            $n = ConvertTo-ImportInt $rawCount
            if ($n -gt $app.deviceCount) { $app.deviceCount = $n }
            $app.CountExplicit = $true
        }

        # Détail par poste (format b)
        $devName = [string](Get-MappedValue $r $map @('deviceName','device name','nom de l''appareil','hostname','computername','machine'))
        if (-not [string]::IsNullOrWhiteSpace($devName)) {
            $devsByKey[$key].Add([PSCustomObject]@{
                deviceName        = $devName
                userPrincipalName = [string](Get-MappedValue $r $map @('userPrincipalName','primary user upn','upn','utilisateur principal','utilisateur','email'))
                operatingSystem   = [string](Get-MappedValue $r $map @('operatingSystem','systeme d''exploitation','plateforme','platform','os'))
                osVersion         = [string](Get-MappedValue $r $map @('osVersion','os version','version du systeme d''exploitation'))
            })
        }
    }

    $apps = @()
    foreach ($key in $appsByKey.Keys) {
        $app  = $appsByKey[$key]
        $devs = @($devsByKey[$key] | Sort-Object -Property deviceName -Unique)
        if (-not $app.CountExplicit) { $app.deviceCount = $devs.Count }
        if ($devs.Count -gt 0) { $out.DevicesByApp[$app.id] = $devs }
        $apps += $app
    }

    $out.Apps = @($apps | Sort-Object -Property deviceCount -Descending)
    Write-Log "Import applications découvertes : $($out.Apps.Count) application(s), détail par poste pour $($out.DevicesByApp.Count) d'entre elles." -Level OK
    return $out
}

function Get-DiscoveredAppsReportFromImport {
    <#
        Équivalent hors ligne de Get-DiscoveredAppsReport : renvoie exactement la même
        structure (AllApps / DevicesByApp / DetailedIds / TruncatedApps), en fusionnant
        au besoin un fichier "liste des applications" et un fichier "postes par application".
    #>
    param(
        [string]$AppsPath,
        [string]$DevicesPath,
        [int]$TopNDetailed = 0,
        [bool]$AnonymizeData = $false,
        [scriptblock]$ProgressCallback
    )

    $apps         = @()
    $devicesByApp = @{}

    if (-not [string]::IsNullOrWhiteSpace($AppsPath)) {
        if ($ProgressCallback) { & $ProgressCallback "Import des applications découvertes..." }
        $primary      = Import-DiscoveredAppsFile -Path $AppsPath
        $apps         = @($primary.Apps)
        $devicesByApp = $primary.DevicesByApp
    }

    if (-not [string]::IsNullOrWhiteSpace($DevicesPath)) {
        if ($ProgressCallback) { & $ProgressCallback "Import du détail des postes par application..." }
        $detail = Import-DiscoveredAppsFile -Path $DevicesPath

        if ($apps.Count -eq 0) {
            $apps         = @($detail.Apps)
            $devicesByApp = $detail.DevicesByApp
        } else {
            # Rattachement au catalogue déjà chargé : d'abord nom+version, sinon nom seul.
            $byKey  = @{}
            $byName = @{}
            foreach ($a in $apps) {
                $byKey[$a.MatchKey] = $a
                $nk = $a.displayName.Trim().ToLower()
                if (-not $byName.ContainsKey($nk)) { $byName[$nk] = $a }
            }
            foreach ($a in $detail.Apps) {
                $devs = $detail.DevicesByApp[$a.id]
                if (-not $devs -or $devs.Count -eq 0) { continue }
                $target = $null
                if ($byKey.ContainsKey($a.MatchKey))                       { $target = $byKey[$a.MatchKey] }
                elseif ($byName.ContainsKey($a.displayName.Trim().ToLower())) { $target = $byName[$a.displayName.Trim().ToLower()] }

                if ($target) {
                    $devicesByApp[$target.id] = $devs
                    if ($target.deviceCount -lt $devs.Count) { $target.deviceCount = $devs.Count }
                } else {
                    $apps += $a
                    $devicesByApp[$a.id] = $devs
                }
            }
        }
    }

    if ($AnonymizeData -and $devicesByApp.Count -gt 0) {
        # [V2.2] Un objet anonymisé par poste, partagé par toutes ses applications
        $anonCache = @{}
        foreach ($k in @($devicesByApp.Keys)) {
            $devicesByApp[$k] = ConvertTo-AnonymizedDeviceList -Devices $devicesByApp[$k] -Cache $anonCache
        }
    }

    $apps = @($apps | Sort-Object -Property deviceCount -Descending)

    # Les données étant locales, aucune limite d'appels n'est nécessaire : le détail est
    # affiché pour toutes les applications qui en disposent, dans la limite du Top N demandé
    # lorsque celui-ci est renseigné (0 = pas de limite).
    $candidates = if ($TopNDetailed -gt 0) { @($apps | Select-Object -First $TopNDetailed) } else { @($apps) }
    $detailedIds = @($candidates | Where-Object { $devicesByApp.ContainsKey($_.id) -and @($devicesByApp[$_.id]).Count -gt 0 } | ForEach-Object { $_.id })

    return [PSCustomObject]@{
        AllApps       = $apps
        DevicesByApp  = $devicesByApp
        DetailedIds   = $detailedIds
        TruncatedApps = @{}
    }
}

function Import-MobileAppsFile {
    <#
        Import de l'inventaire applicatif du tenant (équivalent de /deviceAppManagement/mobileApps).
        Les objets renvoyés portent les noms de propriétés Graph (displayName, publisher,
        displayVersion, createdDateTime) : Get-TenantAppInventory les traite sans modification,
        y compris la recherche des dernières versions publiques.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)

    $rows = Get-ImportRows -Path $Path -Label "inventaire applicatif"
    if (@($rows).Count -eq 0) { return @() }
    $map = New-ColumnMap -Rows $rows

    $list = New-Object System.Collections.Generic.List[psobject]
    foreach ($r in $rows) {
        $name = [string](Get-MappedValue $r $map @('displayName','name','nom de l''application','application name','appName','application','logiciel','nom'))
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $list.Add([PSCustomObject]@{
            displayName     = $name.Trim()
            publisher       = [string](Get-MappedValue $r $map @('publisher','editeur','fabricant','vendor','developer','developpeur','auteur','owner'))
            displayVersion  = [string](Get-MappedValue $r $map @('displayVersion','currentVersion','version actuelle','version deployee','version','app version'))
            createdDateTime = [string](Get-MappedValue $r $map @('createdDateTime','created','date de creation','creation'))
        })
    }

    Write-Log "Import inventaire applicatif : $($list.Count) application(s)." -Level OK
    return $list.ToArray()
}

function Export-RawCollectedData {
    <#
        Écrit les données collectées via l'API dans des fichiers CSV directement
        RÉIMPORTABLES par ce même script (mêmes noms de colonnes que ceux reconnus
        à l'import). Permet de rejouer une génération hors ligne, d'archiver l'état
        d'un tenant à une date donnée, ou de préparer un rapport sans redemander
        d'accès au tenant.
    #>
    param(
        [string]$ClientName,
        [string]$Timestamp,
        [array]$Devices,
        $NonCompliantBreakdown,
        $GracePeriodBreakdown,
        $DiscoveredApps,
        [array]$Inventory,
        [array]$Scores,
        [array]$Performance
    )

    $folder = Join-Path $ImportFolder ("{0}_{1}" -f ($ClientName -replace '[\\/:*?"<>|]', '_'), $Timestamp)
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $written = @()

    # 1) Appareils
    if ($Devices -and $Devices.Count -gt 0) {
        $p = Join-Path $folder "appareils.csv"
        $Devices | Select-Object `
            @{ n = 'id';                e = { $_.id } },
            @{ n = 'deviceName';        e = { $_.deviceName } },
            @{ n = 'userPrincipalName'; e = { $_.userPrincipalName } },
            @{ n = 'operatingSystem';   e = { $_.operatingSystem } },
            @{ n = 'osVersion';         e = { $_.osVersion } },
            @{ n = 'complianceState';   e = { $_.complianceState } },
            @{ n = 'lastSyncDateTime';  e = { $_.lastSyncDateTime } },
            @{ n = 'manufacturer';      e = { $_.manufacturer } },
            @{ n = 'model';             e = { $_.model } },
            @{ n = 'freeStorageSpaceInBytes';  e = { $_.freeStorageSpaceInBytes } },
            @{ n = 'totalStorageSpaceInBytes'; e = { $_.totalStorageSpaceInBytes } } |
            Export-Csv -Path $p -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $written += $p
    }

    # 2) Détail de conformité (motifs déjà lisibles : réimportés tels quels)
    $detailRows = @()
    foreach ($bd in @($NonCompliantBreakdown, $GracePeriodBreakdown)) {
        if (-not $bd) { continue }
        foreach ($entry in $bd.GetEnumerator()) {
            foreach ($d in $entry.Value.Devices) {
                $detailRows += [PSCustomObject]@{
                    DeviceName = $d.DeviceName
                    Motif      = $entry.Key
                    State      = "nonCompliant"
                }
            }
        }
    }
    if ($detailRows.Count -gt 0) {
        $p = Join-Path $folder "conformite-details.csv"
        $detailRows | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $written += $p
    }

    # 3) Applications découvertes + 4) postes par application
    if ($DiscoveredApps -and $DiscoveredApps.AllApps -and @($DiscoveredApps.AllApps).Count -gt 0) {
        $p = Join-Path $folder "apps-decouvertes.csv"
        $DiscoveredApps.AllApps | Select-Object `
            @{ n = 'appId';       e = { $_.id } },
            @{ n = 'displayName'; e = { $_.displayName } },
            @{ n = 'publisher';   e = { $_.publisher } },
            @{ n = 'version';     e = { $_.version } },
            @{ n = 'platform';    e = { $_.platform } },
            @{ n = 'deviceCount'; e = { $_.deviceCount } } |
            Export-Csv -Path $p -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $written += $p

        # [V2.2] List au lieu de "+=" (recopie complète du tableau à chaque ligne) : les listes
        # de postes étant désormais complètes, elles peuvent compter des centaines de milliers de lignes
        $appDeviceRows = New-Object System.Collections.Generic.List[psobject]
        foreach ($app in $DiscoveredApps.AllApps) {
            if (-not $DiscoveredApps.DevicesByApp.ContainsKey($app.id)) { continue }
            foreach ($d in @($DiscoveredApps.DevicesByApp[$app.id])) {
                $appDeviceRows.Add([PSCustomObject]@{
                    appId             = $app.id
                    applicationName   = $app.displayName
                    version           = $app.version
                    deviceName        = $d.deviceName
                    userPrincipalName = $d.userPrincipalName
                    operatingSystem   = $d.operatingSystem
                    osVersion         = $d.osVersion
                })
            }
        }
        if ($appDeviceRows.Count -gt 0) {
            $p = Join-Path $folder "apps-decouvertes-postes.csv"
            $appDeviceRows | Export-Csv -Path $p -NoTypeInformation -Encoding UTF8 -Delimiter ';'
            $written += $p
        }
    }

    # 5) Inventaire applicatif
    if ($Inventory -and $Inventory.Count -gt 0) {
        $p = Join-Path $folder "inventaire-apps.csv"
        $Inventory | Select-Object `
            @{ n = 'displayName';     e = { $_.DisplayName } },
            @{ n = 'publisher';       e = { $_.Publisher } },
            @{ n = 'displayVersion';  e = { $_.CurrentVersion } },
            @{ n = 'createdDateTime'; e = { $_.CreatedDateTime } } |
            Export-Csv -Path $p -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $written += $p
    }

    # 6) Scores de santé (Endpoint Analytics)
    if ($Scores -and $Scores.Count -gt 0) {
        $p = Join-Path $folder "scores-sante.csv"
        $Scores | Select-Object `
            @{ n = 'deviceName';               e = { $_.deviceName } },
            @{ n = 'endpointAnalyticsScore';   e = { $_.endpointAnalyticsScore } },
            @{ n = 'startupPerformanceScore';  e = { $_.startupPerformanceScore } },
            @{ n = 'appReliabilityScore';      e = { $_.appReliabilityScore } },
            @{ n = 'batteryHealthScore';       e = { $_.batteryHealthScore } },
            @{ n = 'workFromAnywhereScore';    e = { $_.workFromAnywhereScore } } |
            Export-Csv -Path $p -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $written += $p
    }

    # 7) Performances de démarrage
    if ($Performance -and $Performance.Count -gt 0) {
        $p = Join-Path $folder "performances-demarrage.csv"
        $Performance | Select-Object `
            @{ n = 'deviceName';        e = { $_.deviceName } },
            @{ n = 'coreBootTimeInMs';  e = { $_.coreBootTimeInMs } },
            @{ n = 'coreLoginTimeInMs'; e = { $_.coreLoginTimeInMs } },
            @{ n = 'bootScore';         e = { $_.bootScore } },
            @{ n = 'loginScore';        e = { $_.loginScore } },
            @{ n = 'blueScreenCount';   e = { $_.blueScreenCount } },
            @{ n = 'restartCount';      e = { $_.restartCount } },
            @{ n = 'diskType';          e = { $_.diskType } } |
            Export-Csv -Path $p -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        $written += $p
    }

    Write-Log "Export des données collectées : $($written.Count) fichier(s) dans $folder" -Level OK
    return $folder
}

function New-ImportTemplates {
    <#
        Écrit dans $ImportFolder\modeles des fichiers CSV vides portant les en-têtes
        attendus : le technicien n'a plus qu'à les remplir (Excel) ou à y coller un
        export existant après avoir renommé les colonnes.
    #>
    $folder = Join-Path $ImportFolder "modeles"
    if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }

    $templates = [ordered]@{
        "modele-appareils.csv"               = "id;deviceName;userPrincipalName;operatingSystem;osVersion;complianceState;lastSyncDateTime;manufacturer;model;freeStorageSpaceInBytes;totalStorageSpaceInBytes;motif"
        "modele-conformite-details.csv"      = "deviceName;motif;setting;policyName;state"
        "modele-apps-decouvertes.csv"        = "appId;displayName;publisher;version;platform;deviceCount"
        "modele-apps-decouvertes-postes.csv" = "appId;applicationName;version;deviceName;userPrincipalName;operatingSystem;osVersion"
        "modele-inventaire-apps.csv"         = "displayName;publisher;displayVersion;createdDateTime"
        "modele-scores-sante.csv"            = "deviceName;endpointAnalyticsScore;startupPerformanceScore;appReliabilityScore;batteryHealthScore;workFromAnywhereScore"
        "modele-performances-demarrage.csv"  = "deviceName;coreBootTimeInMs;coreLoginTimeInMs;bootScore;loginScore;blueScreenCount;restartCount;diskType"
    }
    foreach ($name in $templates.Keys) {
        $p = Join-Path $folder $name
        [System.IO.File]::WriteAllText($p, ($templates[$name] + "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
    }
    Write-Log "Modèles de fichiers d'import générés dans $folder" -Level OK
    return $folder
}

function Update-SourceModeUi {
    <# Active/désactive les champs de l'onglet "Source des données" selon le mode choisi. #>
    $importActive = ($rbSourceImport.Checked -or $rbSourceHybrid.Checked)
    foreach ($c in @($txtImpDevices, $txtImpCompliance, $txtImpDiscovered, $txtImpAppDevices, $txtImpInventory, $txtImpScores, $txtImpPerf,
                     $btnImpDevices, $btnImpCompliance, $btnImpDiscovered, $btnImpAppDevices, $btnImpInventory, $btnImpScores, $btnImpPerf,
                     $txtImpBitLocker, $txtImpDefender, $txtImpAppReliab, $btnImpBitLocker, $btnImpDefender, $btnImpAppReliab)) {
        if ($c) { $c.Enabled = $importActive }
    }
    if ($txtImportClientName) { $txtImportClientName.Enabled = $rbSourceImport.Checked }
    if ($chkExportRaw)        { $chkExportRaw.Enabled        = (-not $rbSourceImport.Checked) }
}

# ========================================
# PAGE 2 - DISCOVERED APPS (APPLICATIONS DÉCOUVERTES)
# ========================================

function ConvertTo-GraphRelativeUrl {
    <# [V2.2] Lien Graph absolu (@odata.nextLink) -> chemin relatif, seule forme acceptée
       dans une sous-requête $batch. #>
    param([string]$Url)
    return ([string]$Url -replace '^https://graph\.microsoft\.com/(beta|v1\.0)', '')
}

function ConvertTo-AnonymizedDeviceList {
    <#
        [V2.2] Pseudonymise une liste de postes (deviceName / userPrincipalName). Le même poste
        figure dans des centaines de listes (une par application installée) : le cache
        $Cache garde UN objet anonymisé par poste, partagé par toutes ces listes, au lieu
        d'un appel et d'un objet par ligne (des centaines de milliers sur un grand parc).
    #>
    param([array]$Devices, [hashtable]$Cache)
    $out = New-Object System.Collections.Generic.List[psobject]
    foreach ($d in @($Devices)) {
        if ($null -eq $d) { continue }
        $realName = [string](Get-PropCI $d @('deviceName', 'DeviceName'))
        $realUpn  = [string](Get-PropCI $d @('userPrincipalName', 'UserPrincipalName'))
        $os       = [string](Get-PropCI $d @('operatingSystem', 'OperatingSystem'))
        $osVer    = [string](Get-PropCI $d @('osVersion', 'OSVersion'))
        $key      = "$realName`t$realUpn`t$os`t$osVer"
        if (-not $Cache.ContainsKey($key)) {
            $anon = Get-AnonymizedIdentity -RealName $realName -RealUpn $realUpn
            $Cache[$key] = [PSCustomObject]@{
                deviceName        = $anon.Name
                userPrincipalName = $anon.Upn
                operatingSystem   = $os
                osVersion         = $osVer
            }
        }
        $out.Add($Cache[$key])
    }
    return $out.ToArray()
}

function Get-DiscoveredAppsReport {
    <#
        Récupère l'ensemble des detectedApps du tenant (paginé), triées par nombre
        de postes décroissant (deviceCount est déjà fourni par Graph, sans appel
        supplémentaire), puis la liste des postes des $TopNDetailed premières
        applications (0 = toutes).

        [V2.2] La liste des postes de chaque application détaillée est désormais COMPLÈTE :
          * toutes les pages (@odata.nextLink) sont lues, et plus seulement la première
            (999 postes au plus : "Affichage limité aux 999 premiers postes collectés") ;
          * les pages suivantes passent, elles aussi, par $batch, regroupées par 20 TOUTES
            APPLICATIONS CONFONDUES : une application à 7 423 postes coûte 8 sous-requêtes,
            servies avec celles des autres applications, au lieu de 8 appels séquentiels.
            C'était la pagination séquentielle, application par application, qui provoquait
            le throttling quasi continu constaté avant (d'où l'abandon du nextLink), pas le
            nombre de postes lui-même ;
          * $TopNDetailed = 0 : toutes les applications sont détaillées (comme en mode import ;
            auparavant, 0 n'en détaillait aucune en mode API) ;
          * une page en échec définitif (après les nouvelles tentatives de Invoke-GraphBatch)
            marque la liste comme INCOMPLÈTE, avec la raison : le rapport l'indique au lieu
            de présenter une liste partielle comme complète.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [int]$TopNDetailed = 0,
        [bool]$AnonymizeData = $false,
        [scriptblock]$ProgressCallback,
        # Garde-fou : 200 pages de 999 postes = près de 200 000 postes pour UNE application
        [int]$MaxPagesPerApp = 200
    )

    if ($ProgressCallback) { & $ProgressCallback "Récupération de la liste des applications découvertes..." }

    $Url     = "https://graph.microsoft.com/beta/deviceManagement/detectedApps?`$top=999"
    $AllApps = @(Get-GraphPagedResults -Url $Url -AccessToken $AccessToken -ProgressCallback $ProgressCallback)
    $AllApps = @($AllApps | Sort-Object -Property deviceCount -Descending)

    $AppsToDetail = if ($TopNDetailed -gt 0) { @($AllApps | Select-Object -First $TopNDetailed) } else { @($AllApps) }
    $AppsToDetail = @($AppsToDetail | Where-Object { $_ -and $_.id })
    $DetailedIds  = @($AppsToDetail | ForEach-Object { [string]$_.id })

    $lists         = @{}   # id application -> List des postes
    $truncatedApps = @{}   # id application -> raison d'une liste incomplète
    $pending       = New-Object System.Collections.Generic.List[psobject]
    foreach ($app in $AppsToDetail) {
        $appId = [string]$app.id
        $lists[$appId] = New-Object System.Collections.Generic.List[psobject]
        # Une application annoncée sur aucun poste n'a rien à lister : aucun appel
        $declared = 0
        try { $declared = [int]$app.deviceCount } catch { }
        if ($declared -le 0) { continue }
        $pending.Add([PSCustomObject]@{
            AppId = $appId
            Page  = 1
            Url   = "/deviceManagement/detectedApps/$appId/managedDevices?`$select=deviceName,userPrincipalName,operatingSystem,osVersion&`$top=999"
        })
    }
    $expectedLinks = 0
    foreach ($app in $AppsToDetail) { try { $expectedLinks += [int]$app.deviceCount } catch { } }
    $scopeLabel = if ($TopNDetailed -gt 0) { "les $($AppsToDetail.Count) applications les plus répandues" } else { "les $($AppsToDetail.Count) applications (toutes)" }
    Write-Log "Postes par application : $scopeLabel, environ $expectedLinks couple(s) application/poste annoncé(s) par Intune." -Level INFO

    # Pagination "en largeur" : à chaque passe, la page suivante de TOUTES les listes
    # inachevées est demandée en même temps, par lots de 20.
    $pass         = 0
    $subRequests  = 0
    while ($pending.Count -gt 0) {
        $pass++
        $requests = New-Object System.Collections.Generic.List[psobject]
        $byReqId  = @{}
        $n        = 0
        foreach ($p in $pending) {
            $n++
            # Id unique DANS la passe : l'id d'application seul ne suffit plus (plusieurs pages)
            $rid = "p$n"
            $byReqId[$rid] = $p
            $requests.Add(@{ id = $rid; method = "GET"; url = $p.Url })
        }
        $subRequests += $requests.Count
        $label = if ($pass -eq 1) { "Postes par application" } else { "Postes par application (pages suivantes, passe $pass)" }
        if ($ProgressCallback) { & $ProgressCallback "$label : $($requests.Count) liste(s) à lire..." }

        $responses = @()
        try {
            $responses = @(Invoke-GraphBatch -Requests $requests.ToArray() -AccessToken $AccessToken -ProgressCallback $ProgressCallback -Label $label)
        } catch {
            # Appel $batch en échec définitif : les listes déjà lues sont gardées, les autres
            # sont signalées incomplètes plutôt que de faire échouer tout le rapport.
            Write-Log "Postes par application : $($_.Exception.Message)" -Level WARN
            foreach ($p in $pending) { $truncatedApps[$p.AppId] = "erreur Graph sur la page $($p.Page) ($($_.Exception.Message))" }
            break
        }

        $next     = New-Object System.Collections.Generic.List[psobject]
        $answered = New-Object System.Collections.Generic.HashSet[string]
        foreach ($resp in $responses) {
            if ($null -eq $resp) { continue }
            $rid = [string]$resp.id
            if (-not $byReqId.ContainsKey($rid)) { continue }
            [void]$answered.Add($rid)
            $p = $byReqId[$rid]
            $status = 0
            try { $status = [int]$resp.status } catch { }
            if ($status -ne 200 -or $null -eq $resp.body) {
                $truncatedApps[$p.AppId] = "HTTP $status sur la page $($p.Page)"
                continue
            }
            foreach ($d in @($resp.body.value)) { if ($null -ne $d) { $lists[$p.AppId].Add($d) } }
            $link = [string]$resp.body.'@odata.nextLink'
            if ($link) {
                if ($p.Page -ge $MaxPagesPerApp) {
                    $truncatedApps[$p.AppId] = "plus de $MaxPagesPerApp pages"
                } else {
                    $next.Add([PSCustomObject]@{ AppId = $p.AppId; Page = $p.Page + 1; Url = (ConvertTo-GraphRelativeUrl $link) })
                }
            }
        }
        foreach ($rid in @($byReqId.Keys)) {
            if (-not $answered.Contains($rid)) { $truncatedApps[$byReqId[$rid].AppId] = "page $($byReqId[$rid].Page) restée sans réponse" }
        }
        $pending = $next
    }

    # Assemblage (+ pseudonymisation, un objet partagé par poste)
    $devicesByApp = @{}
    $anonCache    = @{}
    $listed       = 0
    foreach ($appId in $DetailedIds) {
        $list = $lists[$appId].ToArray()
        if ($AnonymizeData -and $list.Count -gt 0) { $list = ConvertTo-AnonymizedDeviceList -Devices $list -Cache $anonCache }
        $devicesByApp[$appId] = $list
        $listed += $list.Count
    }
    $level = if ($truncatedApps.Count -gt 0) { "WARN" } else { "OK" }
    Write-Log "Postes par application : $listed couple(s) application/poste lu(s) en $subRequests sous-requête(s) ($pass passe(s)) ; $($truncatedApps.Count) liste(s) incomplète(s)." -Level $level

    return [PSCustomObject]@{
        AllApps        = $AllApps
        DevicesByApp   = $devicesByApp
        DetailedIds    = $DetailedIds
        TruncatedApps  = $truncatedApps
    }
}

# ========================================
# PAGE 3 - INVENTAIRE DES APPLICATIONS DU TENANT & AUDIT DES VERSIONS
# ========================================
#
# IMPORTANT - Limite connue : il n'existe pas d'API publique universelle donnant
# "la dernière version publique" de n'importe quel logiciel. Ce script utilise une
# approche "best effort" :
#   1) Si un fichier CSV de correspondance est fourni (séparateur ';', colonnes :
#      AppName;LatestVersion;ReleaseDate — colonnes Info* supplémentaires ignorées),
#      il est utilisé en priorité (source la plus fiable, maintenue par vos équipes).
#      -> Un MODELE pré-rempli avec les noms EXACTS du tenant est généré à chaque
#         exécution à côté du rapport pour toutes les applications non résolues :
#         il suffit de le compléter (Excel) puis de le sélectionner dans l'outil.
#   1bis) Pour les navigateurs d'entreprise (Chrome, Edge, Firefox), une API officielle
#      de l'éditeur est interrogée directement : cette voie fonctionne même sans winget.
#   2) Sinon, si la commande "winget" est disponible sur le poste qui exécute CE script,
#      on interroge le dépôt communautaire winget (winget show / winget search).
#      -> Cela nécessite que le technicien exécute le script sur un poste Windows avec
#         winget installé et un accès Internet vers le dépôt winget.
#      -> ATTENTION : winget est une application PAR UTILISATEUR ; elle est souvent absente
#         d'une console élevée ouverte sous un autre compte (admin) ou d'un Windows Server.
#         Tester "winget --version" dans la MEME console que celle qui lance ce script.
#      -> La date de sortie ("Release Date") n'est pas toujours renseignée par winget ;
#         elle apparaît "Non disponible" le cas échéant.
#   3) A défaut, l'application apparaît avec le statut "Non trouvé" et doit être
#      vérifiée manuellement.

function Get-NormalizedAppName {
    <#
        Nettoie un displayName Intune pour maximiser les chances de correspondance :
        retire les parenthèses, les jetons d'architecture, les numéros de version et
        les suffixes de packaging courants ("MSI", "Enterprise", "Setup", "Installer"...).
        Ex : "Google Chrome Enterprise 126.0.6478.127 (x64) MSI" -> "Google Chrome"
    #>
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }
    $n = " $Name "
    $n = $n -replace '\([^)]*\)', ' '
    $n = $n -replace '\[[^\]]*\]', ' '
    $n = $n -replace '(?i)\b(x64|x86|win64|win32|amd64|arm64|64[\s-]?bits?|32[\s-]?bits?|64bit|32bit)\b', ' '
    $n = $n -replace '(?i)\bv?\d+(\.\d+){1,3}[a-z]?\b', ' '
    $n = $n -replace '(?i)\b(msi|msix|exe|setup|installer|install|package|deploy(ment)?|intune|win32|update|silent|fr|français|french)\b', ' '
    $n = $n -replace '[_]+', ' '
    $n = $n -replace '\s{2,}', ' '
    return $n.Trim(' ', '-', '_', ',', '.')
}

function Get-CompactName {
    # Réduit un nom à ses seuls caractères alphanumériques minuscules, pour comparaison.
    param([string]$Name)
    if (-not $Name) { return "" }
    return ($Name.ToLower() -replace '[^a-z0-9]', '')
}

function Test-VersionEquality {
    <#
        Compare deux numéros de version segment par segment, en tolérant les zéros
        de fin : "24.08" == "24.08.00.0", "2.63.00.00" == "2.63".
    #>
    param([string]$A, [string]$B)
    if ($A -eq $B) { return $true }
    $ca = (($A -replace '[^\d.]', '.') -replace '\.{2,}', '.').Trim('.')
    $cb = (($B -replace '[^\d.]', '.') -replace '\.{2,}', '.').Trim('.')
    if (-not $ca -or -not $cb) { return $false }
    $pa = $ca.Split('.'); $pb = $cb.Split('.')
    $len = [Math]::Max($pa.Count, $pb.Count)
    for ($k = 0; $k -lt $len; $k++) {
        $va = 0; $vb = 0
        if ($k -lt $pa.Count -and $pa[$k] -match '^\d+$') { $va = [int64]$pa[$k] }
        if ($k -lt $pb.Count -and $pb[$k] -match '^\d+$') { $vb = [int64]$pb[$k] }
        if ($va -ne $vb) { return $false }
    }
    return $true
}

function Get-WingetExecutable {
    <#
        Localise winget.exe de façon robuste.
        winget est distribué comme application PAR UTILISATEUR (App Installer) : il est
        très fréquent qu'il soit absent du PATH d'une console élevée ouverte sous un autre
        compte. On tente donc, dans l'ordre :
          1) le PATH de la session courante ;
          2) l'alias d'exécution du profil utilisateur (WindowsApps) ;
          3) le dossier d'installation réel du paquet Microsoft.DesktopAppInstaller.
        Retourne le chemin complet, ou $null.
    #>
    if ($script:WingetExePath) { return $script:WingetExePath }

    $found = $null
    $cmd = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -and (Test-Path $cmd.Source)) { $found = $cmd.Source }

    if (-not $found) {
        $alias = Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\winget.exe"
        if (Test-Path $alias) { $found = $alias }
    }

    if (-not $found) {
        try {
            $pkg = Get-ChildItem -Path (Join-Path $env:ProgramW6432 "WindowsApps") -Filter "Microsoft.DesktopAppInstaller_*" -Directory -ErrorAction SilentlyContinue |
                   Sort-Object Name -Descending
            foreach ($p in $pkg) {
                $candidate = Join-Path $p.FullName "winget.exe"
                if (Test-Path $candidate) { $found = $candidate; break }
            }
        } catch { }
    }

    $script:WingetExePath = $found
    return $found
}

function Invoke-WingetRaw {
    <#
        Exécute winget et retourne TOUTE sa sortie, y compris les messages d'erreur
        (2>&1). Indispensable pour le diagnostic : sans cela, un "source non
        disponible" ou "échec de mise à jour de la source" reste invisible.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$WingetPath,
        [Parameter(Mandatory = $true)][string[]]$Arguments
    )
    try {
        $out = & $WingetPath @Arguments 2>&1
        return @($out | ForEach-Object { "$_" })
    } catch {
        return @("[EXCEPTION] $($_.Exception.Message)")
    }
}

function ConvertFrom-WingetTable {
    <#
        Convertit la sortie tabulaire de winget en objets (Name / Id / Version).

        Deux niveaux, pour résister aux évolutions de format entre versions de winget :
          1) mode normal  : on repère la ligne de séparation (tirets ASCII ou caractères
             semi-graphiques) et on lit les lignes suivantes ;
          2) mode secours : si aucune ligne de séparation n'est trouvée, on balaie
             l'intégralité de la sortie à la recherche de lignes contenant un identifiant
             de la forme "Editeur.Paquet" suivi d'un numéro de version.
        L'identification par la forme de l'ID rend le parsing indépendant de la langue
        d'affichage de winget.
    #>
    param([string[]]$Lines)

    $results = @()
    if (-not $Lines -or $Lines.Count -eq 0) { return $results }

    # Suppression des retours chariot et de l'indicateur de progression (1 à 3 caractères
    # seulement : il ne faut surtout pas neutraliser la ligne de tirets séparatrice).
    $clean = @($Lines | ForEach-Object { ($_ -replace "[`r`b]", '') -replace '^[\\|/\s-]{1,3}$', '' })

    $sep = -1
    for ($i = 0; $i -lt $clean.Count; $i++) {
        if ($clean[$i] -match '^\s*[-\u2500\u2014_]{5,}\s*$') { $sep = $i; break }
    }
    $startIdx = if ($sep -ge 0) { $sep + 1 } else { 0 }

    for ($i = $startIdx; $i -lt $clean.Count; $i++) {
        $line = $clean[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $cols = @(($line -split '\s{2,}') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' })
        if ($cols.Count -lt 2) { continue }

        $idIdx = -1
        for ($k = 1; $k -lt $cols.Count; $k++) {
            if ($cols[$k] -match '^[A-Za-z0-9][\w+-]*(\.[\w+-]+)+$') { $idIdx = $k; break }
        }
        if ($idIdx -lt 1) { continue }

        $version = if (($idIdx + 1) -lt $cols.Count) { $cols[$idIdx + 1] } else { $null }
        if ($version -and $version -notmatch '\d') { $version = $null }   # "Inconnu" / "Unknown"

        $results += [PSCustomObject]@{
            Name    = ($cols[0..($idIdx - 1)] -join ' ').Trim()
            Id      = $cols[$idIdx]
            Version = $version
        }
    }

    return $results
}

function Invoke-WingetSearch {
    <#
        Recherche une application dans le catalogue winget en essayant plusieurs
        formulations, car selon la version de winget et la configuration des sources
        d'entreprise, certaines échouent là où d'autres aboutissent :
          1) recherche par nom, restreinte à la source communautaire "winget" ;
          2) recherche libre (terme positionnel) sur la même source ;
          3) recherche par nom sur TOUTES les sources déclarées (utile si la source
             "winget" a été renommée ou remplacée par un dépôt REST interne).
        La dernière sortie brute est conservée pour le diagnostic.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$WingetPath,
        [Parameter(Mandatory = $true)][string]$Query
    )

    if ([string]::IsNullOrWhiteSpace($Query)) { return @() }
    $common = @('--accept-source-agreements', '--disable-interactivity')

    $strategies = @(
        (@('search', '--name', $Query, '--source', 'winget') + $common),
        (@('search', $Query, '--source', 'winget')           + $common),
        (@('search', '--name', $Query)                       + $common)
    )

    foreach ($st in $strategies) {
        $raw = Invoke-WingetRaw -WingetPath $WingetPath -Arguments $st
        $script:WingetLastRaw     = $raw
        $script:WingetLastCommand = "winget " + ($st -join ' ')
        $parsed = ConvertFrom-WingetTable -Lines $raw
        if ($parsed.Count -gt 0) { return $parsed }
    }

    return @()
}

function Get-NameMatchScore {
    <#
        Score de ressemblance (0-100) entre le nom d'application Intune et un candidat winget.
        Volontairement conservateur : on préfère ne rien proposer plutôt qu'une correspondance
        fausse, qui afficherait un faux "à jour" / "à vérifier".
    #>
    param([string]$SourceName, [string]$CandidateName, [string]$CandidateId, [string]$Publisher)

    $a = Get-CompactName $SourceName
    $b = Get-CompactName $CandidateName
    if (-not $a -or -not $b) { return 0 }

    $score = 0
    $ratio = [double]$b.Length / [double]$a.Length

    # Les correspondances partielles ne sont retenues que si les deux noms sont de
    # longueur comparable : sans ce garde-fou, "Java" correspondrait à "JavaScript".
    if     ($a -eq $b)                                                                     { $score = 100 }
    elseif ($b.StartsWith($a) -and $a.Length -ge 4 -and $ratio -le 1.7)                     { $score = 88 }
    elseif ($a.StartsWith($b) -and $b.Length -ge 4 -and $a.Length -le ($b.Length + 8))      { $score = 84 }
    elseif ($b.Contains($a)   -and $a.Length -ge 6 -and $ratio -le 1.8)                     { $score = 76 }
    elseif ($a.Contains($b)   -and $b.Length -ge 6)                                         { $score = 72 }

    # Rejet ferme si le candidat est un canal de pré-version alors que l'application du
    # tenant n'en est pas un : comparer un déploiement stable à une build Beta ou Canary
    # n'a aucun sens et produirait un faux "à vérifier" à chaque exécution.
    $channels = '(beta|dev|canary|nightly|insider|preview|alpha|rc)'
    if ($CandidateName -match "(?i)\b$channels\b" -and $SourceName -notmatch "(?i)\b$channels\b") {
        return 0
    }

    # Bonus : l'éditeur Intune correspond à la partie "Editeur" de l'identifiant winget
    if ($score -gt 0 -and $Publisher -and $CandidateId -and $CandidateId.Contains('.')) {
        $idPub = Get-CompactName ($CandidateId.Split('.')[0])
        $appPub = Get-CompactName $Publisher
        if ($idPub -and $appPub -and ($appPub.StartsWith($idPub) -or $idPub.StartsWith($appPub))) {
            $score = [Math]::Min(100, $score + 8)
        }
    }
    return $score
}

function Get-WingetLatestVersion {
    <#
        Résout la dernière version publique d'une application via winget.
        Un seul appel winget par requête (la version est lue directement dans le tableau
        de résultats), puis sélection du meilleur candidat par score de ressemblance.
        Les résultats sont mis en cache pour ne jamais interroger deux fois le même nom.
    #>
    param([Parameter(Mandatory = $true)][string]$WingetPath, [Parameter(Mandatory = $true)][string]$AppName, [string]$Publisher)

    if (-not $script:WingetQueryCache) { $script:WingetQueryCache = @{} }

    $normalized = Get-NormalizedAppName -Name $AppName
    if ([string]::IsNullOrWhiteSpace($normalized) -or $normalized.Length -lt 3) { return $null }

    $cacheKey = $normalized.ToLower()
    if ($script:WingetQueryCache.ContainsKey($cacheKey)) { return $script:WingetQueryCache[$cacheKey] }

    $candidates = Invoke-WingetSearch -WingetPath $WingetPath -Query $normalized

    $best = $null; $bestScore = 0
    foreach ($cand in $candidates) {
        if (-not $cand.Version) { continue }
        $s = Get-NameMatchScore -SourceName $normalized -CandidateName $cand.Name -CandidateId $cand.Id -Publisher $Publisher
        if ($s -gt $bestScore) { $bestScore = $s; $best = $cand }
    }

    $result = $null
    if ($best -and $bestScore -ge 70) {
        $result = [PSCustomObject]@{ Version = $best.Version; ReleaseDate = $null; MatchedId = $best.Id; Score = $bestScore }
    }
    $script:WingetQueryCache[$cacheKey] = $result
    return $result
}

function Initialize-ConsoleEncoding {
    <#
        winget écrit sa sortie en UTF-8. Sans cet ajustement, Windows PowerShell 5.1 la
        décode en page de codes OEM et les accents deviennent illisibles ("r├®sultats").
        Appelé aussi bien par le diagnostic que par la génération.
    #>
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
    # Windows PowerShell 5.1 négocie encore TLS 1.0 sur certaines images : sans cela,
    # les points d'accès HTTPS modernes refusent la connexion.
    try {
        if ([Net.ServicePointManager]::SecurityProtocol -notmatch 'Tls12') {
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        }
    } catch { }
}

function Get-GitHubLatestRelease {
    <#
        Dernière version publiée d'un projet hébergé sur GitHub.
        Un seul appel par dépôt et par exécution (mise en cache), pour rester très en
        dessous de la limite de 60 requêtes/heure de l'API publique GitHub.
        Le nom de tag est nettoyé : "v8.9.7", "Audacity-3.5.1" ou "v2.45.2.windows.1"
        donnent respectivement 8.9.7, 3.5.1 et 2.45.2.
    #>
    param([Parameter(Mandatory = $true)][string]$Repo)

    if (-not $script:GitHubReleaseCache) { $script:GitHubReleaseCache = @{} }
    if ($script:GitHubReleaseCache.ContainsKey($Repo)) { return $script:GitHubReleaseCache[$Repo] }

    $value = $null
    try {
        $r = Invoke-RestMethod -Uri "https://api.github.com/repos/$Repo/releases/latest" -TimeoutSec 20 `
                               -Headers @{ 'User-Agent' = 'IntuneDashboard' } -ErrorAction Stop
        if ($r.tag_name) {
            $v = "$($r.tag_name)" -replace '^[^0-9]*', ''
            if ($v -match '^(\d+(\.\d+)*)') { $v = $Matches[1] }
            $d = $null
            try { $d = ([datetime]$r.published_at).ToString('dd/MM/yyyy') } catch { }
            if ($v) { $value = [PSCustomObject]@{ Version = $v; ReleaseDate = $d } }
        }
    } catch {
        Write-Log "Dépôt GitHub '$Repo' injoignable : $($_.Exception.Message)" -Level WARN
    }

    $script:GitHubReleaseCache[$Repo] = $value
    return $value
}

function Get-VendorResolverTable {
    <#
        Table des sources officielles, PILOTÉE PAR LES DONNÉES.

        Chaque entrée décrit un résolveur par de simples champs texte plutôt que par des
        fermetures ("closures") : l'appariement se réduit à un test d'expression régulière
        effectué par l'appelant, ce qui supprime toute dépendance au comportement de
        GetNewClosure() et se comporte donc à l'identique sur Windows PowerShell 5.1 et
        sur PowerShell 7.

        POUR AJOUTER UNE APPLICATION HÉBERGÉE SUR GITHUB : ajoutez une ligne dans $ghMap
        (motif + dépôt "propriétaire/projet"). Rien d'autre à faire.
        POUR AJOUTER UN AUTRE ÉDITEUR : ajoutez une entrée de Type 'Script' avec son
        propre bloc Resolve renvoyant @{ Version = "..."; ReleaseDate = "..." }.

        Champs :
          Pattern  : expression régulière testée sur le nom normalisé EN MINUSCULES
          Exclude  : expression régulière d'exclusion (facultative)
          Type     : 'Script' (bloc Resolve) ou 'GitHub' (champ Repo)
    #>
    if ($script:VendorResolverTable) { return $script:VendorResolverTable }

    # --- Applications dont la version officielle est publiée sur GitHub ---
    # Motifs volontairement ancrés pour éviter les faux acheminements
    # ("GitHub Desktop" ou "Git Extensions" ne doivent pas viser le dépôt de Git).
    $ghMap = @(
        @{ Pattern = '^notepad\+\+$|^notepadplusplus$';               Repo = 'notepad-plus-plus/notepad-plus-plus' },
        @{ Pattern = '^7[ _-]?zip$';                                  Repo = 'ip7z/7zip' },
        @{ Pattern = '^(microsoft )?visual studio code$|^vs ?code$';  Repo = 'microsoft/vscode' },
        @{ Pattern = '^git( for windows)?$';                          Repo = 'git-for-windows/git' },
        @{ Pattern = '^(microsoft )?powershell$|^pwsh$';              Repo = 'PowerShell/PowerShell' },
        @{ Pattern = '^windows terminal$';                            Repo = 'microsoft/terminal' },
        @{ Pattern = '^winmerge$';                                    Repo = 'WinMerge/winmerge' },
        @{ Pattern = '^keepassxc$';                                   Repo = 'keepassxreboot/keepassxc' },
        @{ Pattern = '^obs studio$|^obs$';                            Repo = 'obsproject/obs-studio' },
        @{ Pattern = '^greenshot$';                                   Repo = 'greenshot/greenshot' },
        @{ Pattern = '^audacity$';                                    Repo = 'audacity/audacity' }
    )

    $table = @(
        [PSCustomObject]@{
            Key = 'chrome'; Label = 'Google Chrome'; Type = 'Script'
            Pattern = '\bchrome\b'; Exclude = 'chromium'
            Resolve = {
                $r = Invoke-RestMethod -Uri 'https://versionhistory.googleapis.com/v1/chrome/platforms/win64/channels/stable/versions' -TimeoutSec 20 -ErrorAction Stop
                if ($r.versions -and $r.versions.Count -gt 0) {
                    return [PSCustomObject]@{ Version = $r.versions[0].version; ReleaseDate = $null }
                }
                return $null
            }
        },
        [PSCustomObject]@{
            Key = 'firefox'; Label = 'Mozilla Firefox'; Type = 'Script'
            Pattern = '\bfirefox\b'; Exclude = $null
            Resolve = {
                param($OriginalName)
                $r = Invoke-RestMethod -Uri 'https://product-details.mozilla.org/1.0/firefox_versions.json' -TimeoutSec 20 -ErrorAction Stop
                $v = if ($OriginalName -match '(?i)\besr\b') { $r.FIREFOX_ESR } else { $r.LATEST_FIREFOX_VERSION }
                if ($v) { return [PSCustomObject]@{ Version = "$v"; ReleaseDate = $null } }
                return $null
            }
        },
        [PSCustomObject]@{
            Key = 'thunderbird'; Label = 'Mozilla Thunderbird'; Type = 'Script'
            Pattern = '\bthunderbird\b'; Exclude = $null
            Resolve = {
                # Le nom exact de la clé varie selon la version du service Mozilla : on
                # balaie les propriétés retournées et on retient la première version
                # stable. En cas d'échec, on remonte la liste des clés disponibles dans
                # le message d'erreur, ce qui permet de corriger sans deviner.
                $r = Invoke-RestMethod -Uri 'https://product-details.mozilla.org/1.0/thunderbird_versions.json' -TimeoutSec 20 -ErrorAction Stop
                $props = @($r.PSObject.Properties)
                foreach ($rank in @('^LATEST_THUNDERBIRD_VERSION$', '^THUNDERBIRD_VERSION$', 'THUNDERBIRD')) {
                    foreach ($prop in $props) {
                        if ($prop.Name -notmatch $rank) { continue }
                        if ($prop.Name -match '(?i)(alpha|beta|nightly|devel|aurora|esr)') { continue }
                        $v = "$($prop.Value)"
                        if ($v -match '^\d+(\.\d+)*') { return [PSCustomObject]@{ Version = $v; ReleaseDate = $null } }
                    }
                }
                $keys = (($props | Select-Object -First 10).Name) -join ', '
                throw "aucune clé de version exploitable. Clés reçues : $keys"
            }
        },
        [PSCustomObject]@{
            Key = 'edge'; Label = 'Microsoft Edge'; Type = 'Script'
            Pattern = '\bedge\b'; Exclude = 'edgewater'
            Resolve = {
                $r = Invoke-RestMethod -Uri 'https://edgeupdates.microsoft.com/api/products?view=enterprise' -TimeoutSec 20 -ErrorAction Stop
                $stable = @($r | Where-Object { $_.Product -eq 'Stable' })
                if ($stable.Count -eq 0) { return $null }
                $rel = @($stable[0].Releases |
                         Where-Object { $_.Platform -eq 'Windows' -and $_.Architecture -eq 'x64' } |
                         Sort-Object { [datetime]$_.PublishedTime } -Descending)
                if ($rel.Count -eq 0) { return $null }
                $d = $null
                try { $d = ([datetime]$rel[0].PublishedTime).ToString('dd/MM/yyyy') } catch { }
                return [PSCustomObject]@{ Version = $rel[0].ProductVersion; ReleaseDate = $d }
            }
        },
        [PSCustomObject]@{
            Key = 'vlc'; Label = 'VLC media player'; Type = 'Script'
            Pattern = '\bvlc\b'; Exclude = $null
            Resolve = {
                $r = Invoke-RestMethod -Uri 'https://update.videolan.org/vlc/status-win-x64' -TimeoutSec 20 -ErrorAction Stop
                foreach ($line in ("$r" -split "`n")) {
                    $t = $line.Trim()
                    if ($t -match '^\d+(\.\d+){1,3}$') { return [PSCustomObject]@{ Version = $t; ReleaseDate = $null } }
                }
                return $null
            }
        }
    )

    foreach ($entry in $ghMap) {
        $table += [PSCustomObject]@{
            Key     = "gh:$($entry.Repo)"
            Label   = $entry.Repo
            Type    = 'GitHub'
            Pattern = $entry.Pattern
            Exclude = $null
            Repo    = $entry.Repo
            Resolve = $null
        }
    }

    $script:VendorResolverTable = $table
    return $script:VendorResolverTable
}

function Get-XmlNodeText {
    <#
        Extrait le texte d'un noeud XML. Nécessaire car PowerShell renvoie une chaîne
        pour un élément simple, mais un objet XmlElement dès que l'élément porte un
        attribut (cas de <d:Published m:type="Edm.DateTime">) : une interpolation
        directe donnerait alors le nom de la classe au lieu de la valeur.
    #>
    param($Node)
    if ($null -eq $Node) { return "" }
    if ($Node -is [System.Xml.XmlElement]) { return $Node.InnerText }
    return "$Node"
}

function Invoke-ChocolateySearch {
    <#
        Interroge le flux public du dépôt communautaire Chocolatey (OData v2, Atom/XML).

        POURQUOI CETTE SOURCE : c'est un catalogue de logiciels Windows de couverture
        comparable à celui de winget, mais interrogé en HTTPS standard depuis PowerShell.
        Il passe donc par le proxy système, là où winget utilise sa propre pile réseau et
        échoue lorsque cdn.winget.microsoft.com est filtré.
        Aucune installation de Chocolatey n'est requise : on ne fait que lire le flux.
    #>
    param([Parameter(Mandatory = $true)][string]$Query)

    $results = @()
    if ([string]::IsNullOrWhiteSpace($Query)) { return $results }

    $term   = [System.Uri]::EscapeDataString($Query)
    $lower  = [System.Uri]::EscapeDataString($Query.ToLower())
    $base   = 'https://community.chocolatey.org/api/v2'

    # La fonction OData "Search()" de NuGet v2 exige targetFramework ET includePrerelease :
    # sans ces deux paramètres, le service répond 400 (Bad Request). Une requête
    # "Packages()" filtrée sert de repli si la première forme est refusée.
    $urls = @(
        "$base/Search()?searchTerm=%27$term%27&targetFramework=%27%27&includePrerelease=false&`$filter=IsLatestVersion&`$top=10",
        "$base/Packages()?`$filter=IsLatestVersion%20and%20substringof(%27$lower%27,tolower(Id))&`$top=10"
    )

    $resp = $null
    $lastErr = $null
    foreach ($url in $urls) {
        try {
            $resp = Invoke-WebRequest -Uri $url -TimeoutSec 25 -UseBasicParsing -ErrorAction Stop
            $lastErr = $null
            break
        } catch {
            $code = ""
            try { $code = [int]$_.Exception.Response.StatusCode } catch { }
            $lastErr = "$($_.Exception.Message)$(if ($code) { " (HTTP $code)" })"
            $resp = $null
        }
    }

    if (-not $resp) {
        $script:ChocoLastError = $lastErr
        return $results
    }

    try {
        $xml = [xml]$resp.Content
        $script:ChocoLastError = $null

        foreach ($entry in @($xml.feed.entry)) {
            if (-not $entry) { continue }
            $props = $entry.properties
            if (-not $props) { continue }

            # Le titre lisible est plus proche du nom Intune que l'identifiant technique.
            $id    = Get-XmlNodeText $entry.title
            $title = Get-XmlNodeText $props.Title
            if ([string]::IsNullOrWhiteSpace($title)) { $title = $id }
            $ver   = Get-XmlNodeText $props.Version
            if ([string]::IsNullOrWhiteSpace($ver)) { continue }

            $date = $null
            try {
                $published = Get-XmlNodeText $props.Published
                if ($published) { $date = ([datetime]$published).ToString('dd/MM/yyyy') }
            } catch { }

            $results += [PSCustomObject]@{ Name = $title; Id = $id; Version = $ver; ReleaseDate = $date }
        }
    } catch {
        $script:ChocoLastError = $_.Exception.Message
    }

    return $results
}

function Get-ChocolateyLatestVersion {
    <#
        Résout la dernière version publique d'une application via Chocolatey, en
        réutilisant exactement le même moteur de rapprochement de noms que winget
        (score de ressemblance, rejet des canaux de pré-version, bonus éditeur).
        Résultats mis en cache : un même nom n'est jamais interrogé deux fois.
    #>
    param([Parameter(Mandatory = $true)][string]$AppName, [string]$Publisher)

    if (-not $script:ChocoQueryCache) { $script:ChocoQueryCache = @{} }

    $normalized = Get-NormalizedAppName -Name $AppName
    if ([string]::IsNullOrWhiteSpace($normalized) -or $normalized.Length -lt 3) { return $null }

    $cacheKey = $normalized.ToLower()
    if ($script:ChocoQueryCache.ContainsKey($cacheKey)) { return $script:ChocoQueryCache[$cacheKey] }

    $candidates = Invoke-ChocolateySearch -Query $normalized

    $best = $null; $bestScore = 0
    foreach ($cand in $candidates) {
        if (-not $cand.Version) { continue }
        # On tente le rapprochement sur le titre ET sur l'identifiant du paquet
        # ("7-Zip" comme "7zip"), en gardant le meilleur des deux.
        $s1 = Get-NameMatchScore -SourceName $normalized -CandidateName $cand.Name -CandidateId $cand.Id -Publisher $Publisher
        $s2 = Get-NameMatchScore -SourceName $normalized -CandidateName $cand.Id   -CandidateId $cand.Id -Publisher $Publisher
        $s  = [Math]::Max($s1, $s2)
        if ($s -gt $bestScore) { $bestScore = $s; $best = $cand }
    }

    $result = $null
    # Seuil volontairement plus strict que winget : la recherche Chocolatey est floue
    # (elle peut renvoyer des paquets populaires sans réel rapport avec la requête), on
    # n'accepte donc que les correspondances de niveau "égalité ou préfixe" (>= 80).
    if ($best -and $bestScore -ge 80) {
        $result = [PSCustomObject]@{
            Version     = $best.Version
            ReleaseDate = $best.ReleaseDate
            MatchedId   = $best.Id
            Score       = $bestScore
        }
    }
    $script:ChocoQueryCache[$cacheKey] = $result
    return $result
}

# ------------------------------------------------------------
# CATALOGUE winget-pkgs LU DIRECTEMENT SUR GITHUB
# ------------------------------------------------------------
#
# POURQUOI : la commande "winget" est absente des sessions élevées, des comptes de
# service et de Windows Server, et son CDN (cdn.winget.microsoft.com) est fréquemment
# bloqué en entreprise. Le MÊME catalogue est pourtant public sur GitHub :
#     https://github.com/microsoft/winget-pkgs
# Cette source l'interroge en HTTPS standard via api.github.com, qui passe le proxy
# d'entreprise dans la quasi-totalité des cas.
#
# COMMENT : l'arborescence du dépôt est
#     manifests/<lettre>/<Éditeur>/<Paquet>[/<sous-espace>...]/<Version>/*.yaml
# On descend cet arbre avec l'API "git/trees" (et non "contents", plafonnée à
# 1 000 entrées, insuffisant pour des lettres comme "m"). Le dernier appel est
# récursif sur le seul sous-arbre du paquet : les versions sont alors les dossiers
# parents des fichiers .yaml, ce qui reste exact même pour les identifiants à
# rallonge (Microsoft.VisualStudio.2022.Community -> .../2022/Community/17.9.0/).
#
# QUOTA : l'API GitHub anonyme est limitée à 60 requêtes/heure et par adresse IP,
# soit une trentaine d'applications seulement. Un jeton d'accès personnel — même
# SANS AUCUNE PORTÉE, le dépôt étant public — porte cette limite à 5 000/heure.
# Le champ prévu dans l'interface est facultatif ; sans jeton, la source s'arrête
# proprement dès le quota atteint et le journal l'indique explicitement.

function Get-GitHubToken {
    <# Jeton lu depuis le fichier local (facultatif). #>
    param([string]$Path = $GitHubTokenFile)
    if (-not (Test-Path $Path)) { return "" }
    try {
        $t = (Get-Content -Path $Path -Raw -ErrorAction Stop).Trim()
        return $t
    } catch { return "" }
}

function Save-GitHubToken {
    <#
        Enregistre le jeton pour les exécutions suivantes. Stocké en clair : il s'agit
        d'un jeton de LECTURE SEULE sur un dépôt public, sans portée ni accès à vos
        données. Ne réutilisez pas ici un jeton disposant de droits d'écriture.
    #>
    param([string]$Token, [string]$Path = $GitHubTokenFile)
    try {
        if ([string]::IsNullOrWhiteSpace($Token)) {
            if (Test-Path $Path) { Remove-Item $Path -Force }
            return
        }
        $folder = Split-Path -Path $Path -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        [System.IO.File]::WriteAllText($Path, $Token.Trim(), (New-Object System.Text.UTF8Encoding($false)))
        Write-Log "Jeton GitHub enregistré ($Path)." -Level INFO
    } catch {
        Write-Log "Enregistrement du jeton GitHub impossible : $($_.Exception.Message)" -Level WARN
    }
}

function Invoke-GitHubTreeApi {
    <#
        Appel unitaire à l'API "git/trees" de winget-pkgs, avec cache par SHA.
        Toute erreur de quota (403/429) désactive la source pour le reste de
        l'exécution : sans cela, chaque application restante générerait un appel
        voué à l'échec, pour plusieurs minutes perdues.
    #>
    param([Parameter(Mandatory = $true)][string]$Sha, [switch]$Recursive)

    if ($script:GhPkgsDisabled) { return $null }
    if (-not $script:GhTreeCache) { $script:GhTreeCache = @{} }

    $cacheKey = if ($Recursive) { "R:$Sha" } else { "N:$Sha" }
    if ($script:GhTreeCache.ContainsKey($cacheKey)) { return $script:GhTreeCache[$cacheKey] }

    $url = "https://api.github.com/repos/microsoft/winget-pkgs/git/trees/$Sha"
    if ($Recursive) { $url += "?recursive=1" }

    $headers = @{ 'User-Agent' = 'IntuneDashboard'; 'Accept' = 'application/vnd.github+json' }
    if (-not [string]::IsNullOrWhiteSpace($script:GhPkgsToken)) {
        $headers['Authorization'] = "Bearer $($script:GhPkgsToken)"
    }

    $entries = $null
    try {
        $script:GhPkgsCallCount++
        $r = Invoke-RestMethod -Uri $url -Headers $headers -TimeoutSec 25 -ErrorAction Stop
        if ($r.tree) { $entries = @($r.tree) } else { $entries = @() }
        if ($r.truncated -eq $true) {
            Write-Log "Arborescence GitHub tronquée sur '$Sha' : certaines versions peuvent manquer." -Level WARN
        }
    } catch {
        $status = $null
        try { $status = [int]$_.Exception.Response.StatusCode } catch { }
        if ($status -eq 403 -or $status -eq 429) {
            $script:GhPkgsDisabled = $true
            $hint = if ([string]::IsNullOrWhiteSpace($script:GhPkgsToken)) {
                "Quota anonyme de 60 requêtes/heure atteint après $($script:GhPkgsCallCount) appel(s) : renseignez un jeton GitHub (5 000/heure) dans l'onglet Rapports & Versions."
            } else {
                "Quota GitHub atteint ou jeton refusé après $($script:GhPkgsCallCount) appel(s)."
            }
            Write-Log "Catalogue winget-pkgs (GitHub) désactivé pour cette exécution. $hint" -Level WARN
        } elseif ($status -eq 401) {
            $script:GhPkgsDisabled = $true
            Write-Log "Jeton GitHub refusé (HTTP 401) : catalogue winget-pkgs désactivé. Vérifiez le jeton saisi." -Level WARN
        } else {
            Write-Log "Appel GitHub en échec ($url) : $($_.Exception.Message)" -Level WARN
        }
        $entries = $null
    }

    $script:GhTreeCache[$cacheKey] = $entries
    return $entries
}

function Get-WingetPkgsLetterTree {
    <#
        SHA du dossier "manifests/<lettre>" (les éditeurs). Le chemin racine ->
        manifests -> lettre n'est parcouru qu'une fois par lettre et par exécution.
    #>
    param([Parameter(Mandatory = $true)][string]$Letter)

    if (-not $script:GhLetterShaCache) { $script:GhLetterShaCache = @{} }
    $key = $Letter.ToLower()
    if ($script:GhLetterShaCache.ContainsKey($key)) { return $script:GhLetterShaCache[$key] }

    $value = $null
    if (-not $script:GhManifestsSha) {
        $root = Invoke-GitHubTreeApi -Sha "master"
        if ($root) {
            $m = $root | Where-Object { $_.path -eq "manifests" -and $_.type -eq "tree" } | Select-Object -First 1
            if ($m) { $script:GhManifestsSha = $m.sha }
        }
    }
    if ($script:GhManifestsSha) {
        $letters = Invoke-GitHubTreeApi -Sha $script:GhManifestsSha
        if ($letters) {
            $l = $letters | Where-Object { $_.path -eq $key -and $_.type -eq "tree" } | Select-Object -First 1
            if ($l) { $value = $l.sha }
        }
    }

    $script:GhLetterShaCache[$key] = $value
    return $value
}

function Get-WingetPkgsPublisherSha {
    <# SHA du dossier éditeur, par rapprochement du nom (comparaison compacte). #>
    param([Parameter(Mandatory = $true)][string]$PublisherName)

    $compact = Get-CompactName $PublisherName
    if (-not $compact) { return $null }

    $letter = $compact.Substring(0, 1)
    $sha    = Get-WingetPkgsLetterTree -Letter $letter
    if (-not $sha) { return $null }

    $publishers = Invoke-GitHubTreeApi -Sha $sha
    if (-not $publishers) { return $null }

    $best = $null; $bestLen = 0
    foreach ($p in $publishers) {
        if ($p.type -ne "tree") { continue }
        $c = Get-CompactName $p.path
        if (-not $c) { continue }
        if ($c -eq $compact) { return [PSCustomObject]@{ Name = $p.path; Sha = $p.sha } }
        # Correspondance par préfixe : "Google" pour un éditeur Intune "Google LLC",
        # "Igor Pavlov" pour "IgorPavlov". On garde le préfixe le plus long.
        if (($compact.StartsWith($c) -or $c.StartsWith($compact)) -and $c.Length -ge 3 -and $c.Length -gt $bestLen) {
            $best = [PSCustomObject]@{ Name = $p.path; Sha = $p.sha }
            $bestLen = $c.Length
        }
    }
    return $best
}

function Get-WingetPkgsVersionsFromTree {
    <#
        Extrait les versions d'un sous-arbre de paquet : ce sont les dossiers qui
        contiennent directement des fichiers .yaml. Cette règle vaut à toute
        profondeur, donc aussi pour les identifiants à sous-espaces.
    #>
    param($TreeEntries)
    $versions = New-Object System.Collections.Generic.HashSet[string]
    foreach ($e in $TreeEntries) {
        if ($e.type -ne "blob") { continue }
        if ($e.path -notmatch '\.ya?ml$') { continue }
        $parts = $e.path -split '/'
        if ($parts.Count -lt 2) { continue }
        [void]$versions.Add($parts[$parts.Count - 2])
    }
    return @($versions)
}

function Select-HighestVersion {
    <#
        Version la plus élevée d'une liste, en comparaison numérique quand c'est
        possible ("1.10" > "1.9", qu'un tri texte inverserait). Les canaux de
        pré-version sont écartés : les comparer à un déploiement stable produirait
        un faux "à vérifier" à chaque exécution.
    #>
    param([string[]]$Versions)
    $stable = @($Versions | Where-Object { $_ -and ($_ -notmatch '(?i)(beta|alpha|rc\d|preview|nightly|canary|insider|dev)') })
    if ($stable.Count -eq 0) { $stable = @($Versions | Where-Object { $_ }) }
    if ($stable.Count -eq 0) { return $null }

    $best = $null; $bestObj = $null
    foreach ($v in $stable) {
        $numeric = ($v -replace '[^0-9.]', '')
        $numeric = ($numeric -replace '\.{2,}', '.').Trim('.')
        $parsed  = $null
        if ($numeric -and ($numeric -match '^\d+(\.\d+){0,3}$')) {
            try { $parsed = [version]$numeric } catch { $parsed = $null }
        }
        if ($parsed) {
            if (-not $bestObj -or $parsed -gt $bestObj) { $bestObj = $parsed; $best = $v }
        } elseif (-not $bestObj -and (-not $best -or ($v -gt $best))) {
            $best = $v
        }
    }
    return $best
}

function Get-WingetPkgsLatestVersion {
    <#
        Dernière version publique d'une application d'après le dépôt GitHub winget-pkgs.
        Coût : 1 appel pour la liste des paquets de l'éditeur + 1 appel récursif sur le
        paquet retenu (les niveaux racine/manifests/lettre sont mutualisés). Résultats
        et échecs mis en cache : un même nom n'est jamais interrogé deux fois.
    #>
    param([Parameter(Mandatory = $true)][string]$AppName, [string]$Publisher)

    if ($script:GhPkgsDisabled) { return $null }
    if (-not $script:GhPkgsQueryCache) { $script:GhPkgsQueryCache = @{} }

    $normalized = Get-NormalizedAppName -Name $AppName
    if ([string]::IsNullOrWhiteSpace($normalized) -or $normalized.Length -lt 3) { return $null }

    $cacheKey = ($normalized + '|' + $Publisher).ToLower()
    if ($script:GhPkgsQueryCache.ContainsKey($cacheKey)) { return $script:GhPkgsQueryCache[$cacheKey] }

    # Candidats "éditeur" : l'éditeur Intune, puis le premier mot du nom de
    # l'application (cas "Google Chrome"), puis le nom entier (cas "7-Zip", dont le
    # dossier éditeur est "7zip").
    $pubCandidates = @()
    if (-not [string]::IsNullOrWhiteSpace($Publisher)) { $pubCandidates += $Publisher }
    $firstWord = ($normalized -split '\s+')[0]
    if ($firstWord -and $firstWord.Length -ge 3) { $pubCandidates += $firstWord }
    $pubCandidates += $normalized
    $pubCandidates = @($pubCandidates | Where-Object { $_ } | Select-Object -Unique)

    $result = $null
    foreach ($cand in $pubCandidates) {
        if ($script:GhPkgsDisabled) { break }
        $pub = Get-WingetPkgsPublisherSha -PublisherName $cand
        if (-not $pub) { continue }

        $packages = Invoke-GitHubTreeApi -Sha $pub.Sha
        if (-not $packages) { continue }

        # Nom "court" : le nom de l'application débarrassé du nom de l'éditeur en
        # préfixe, pour que "Google Chrome" retrouve le paquet "Chrome".
        $shortName  = $normalized
        $pubCompact = Get-CompactName $pub.Name
        $normCompact = Get-CompactName $normalized
        if ($pubCompact -and $normCompact.StartsWith($pubCompact) -and $normCompact.Length -gt $pubCompact.Length) {
            $words = @($normalized -split '\s+')
            if ($words.Count -gt 1) { $shortName = ($words[1..($words.Count - 1)] -join ' ') }
        }

        $best = $null; $bestScore = 0
        foreach ($p in $packages) {
            if ($p.type -ne "tree") { continue }
            $candidateId = "$($pub.Name).$($p.path)"
            $s1 = Get-NameMatchScore -SourceName $normalized -CandidateName $p.path -CandidateId $candidateId -Publisher $Publisher
            $s2 = Get-NameMatchScore -SourceName $shortName  -CandidateName $p.path -CandidateId $candidateId -Publisher $Publisher
            $s  = [Math]::Max($s1, $s2)
            if ($s -gt $bestScore) { $bestScore = $s; $best = $p }
        }

        # Seuil aligné sur celui de winget : le rapprochement s'appuie sur des noms de
        # dossiers structurés, plus fiables qu'un résultat de recherche plein texte.
        if (-not $best -or $bestScore -lt 72) { continue }

        $subtree = Invoke-GitHubTreeApi -Sha $best.sha -Recursive
        if (-not $subtree) { continue }

        $versions = Get-WingetPkgsVersionsFromTree -TreeEntries $subtree
        $latest   = Select-HighestVersion -Versions $versions
        if ($latest) {
            $result = [PSCustomObject]@{
                Version     = $latest
                ReleaseDate = $null   # non publiée dans les manifestes winget-pkgs
                MatchedId   = "$($pub.Name).$($best.path)"
                Score       = $bestScore
            }
            break
        }
    }

    $script:GhPkgsQueryCache[$cacheKey] = $result
    return $result
}

function Test-WingetPkgsConnectivity {
    <#
        Sonde unique du catalogue GitHub sur une application très répandue (7-Zip).
        S'il ne répond pas, la source est écartée immédiatement plutôt que de faire
        échouer des centaines de requêtes une par une.
    #>
    param([string]$Token)

    $script:GhPkgsDisabled  = $false
    $script:GhPkgsCallCount = 0
    $script:GhTreeCache     = @{}
    $script:GhLetterShaCache = @{}
    $script:GhPkgsQueryCache = @{}
    $script:GhManifestsSha   = $null
    $script:GhPkgsToken      = $Token

    $probe = Get-WingetPkgsLatestVersion -AppName "7-Zip" -Publisher "Igor Pavlov"
    if ($probe -and $probe.Version) {
        $mode = if ([string]::IsNullOrWhiteSpace($Token)) { "sans jeton (60 requêtes/heure)" } else { "avec jeton (5 000 requêtes/heure)" }
        Write-Log "Catalogue winget-pkgs (GitHub) opérationnel $mode : test '7-Zip' -> $($probe.Version) via $($probe.MatchedId)." -Level OK
        return $true
    }

    if (-not $script:GhPkgsDisabled) {
        Write-Log "Catalogue winget-pkgs (GitHub) sans résultat au test '7-Zip' : source non utilisée pour cette exécution." -Level WARN
    }
    $script:GhPkgsDisabled = $true
    return $false
}

function Get-VendorLatestVersion {
    <#
        Interroge le résolveur officiel correspondant à l'application, s'il en existe un.

        Trois protections issues d'un incident réel (toutes les applications d'un tenant
        s'étaient vu attribuer la version de Chrome) :
          1. appariement par [regex]::IsMatch avec motif validé — un motif nul ou vide ne
             peut plus correspondre à tout ;
          2. traçage du résolveur retenu ($script:LastVendorKey), repris dans la colonne
             "Source" du rapport : toute anomalie d'appariement devient visible ;
          3. garde-fou : un même résolveur ne peut servir plus de 10 applications
             différentes par exécution — au-delà, ses correspondances sont refusées et
             signalées (un résolveur d'éditeur est par nature spécifique à UN produit).
    #>
    param([Parameter(Mandatory = $true)][string]$AppName, [string]$Publisher)

    if (-not $script:VendorCache)      { $script:VendorCache      = @{} }
    if (-not $script:VendorErrorCache) { $script:VendorErrorCache = @{} }
    if (-not $script:VendorKeyHits)    { $script:VendorKeyHits    = @{} }
    if (-not $script:VendorCapWarned)  { $script:VendorCapWarned  = @{} }

    $n = (Get-NormalizedAppName -Name $AppName).ToLower()
    $script:LastNormalizedName = $n
    $script:LastVendorKey      = $null
    if (-not $n) { return $null }

    foreach ($resolver in (Get-VendorResolverTable)) {
        $pattern = "$($resolver.Pattern)"
        if ([string]::IsNullOrWhiteSpace($pattern)) { continue }

        $isMatch = $false
        try { $isMatch = [regex]::IsMatch($n, $pattern) } catch { continue }
        if (-not $isMatch) { continue }

        if (-not [string]::IsNullOrWhiteSpace("$($resolver.Exclude)")) {
            $isExcluded = $false
            try { $isExcluded = [regex]::IsMatch($n, "$($resolver.Exclude)") } catch { }
            if ($isExcluded) { continue }
        }

        # Le cache distingue les variantes ESR / standard d'un même éditeur.
        $cacheKey = $resolver.Key
        if ($resolver.Key -eq 'firefox' -and $AppName -match '(?i)\besr\b') { $cacheKey = 'firefox-esr' }

        # --- Chemin cache : on restaure aussi le motif d'échec, et le garde-fou
        #     s'applique de la même façon qu'à une résolution fraîche. ---
        if ($script:VendorCache.ContainsKey($cacheKey)) {
            $script:LastVendorError = $script:VendorErrorCache[$cacheKey]
            $cached = $script:VendorCache[$cacheKey]
            if ($cached) {
                $script:VendorKeyHits[$cacheKey] = [int]$script:VendorKeyHits[$cacheKey] + 1
                if ($script:VendorKeyHits[$cacheKey] -gt 10) {
                    if (-not $script:VendorCapWarned[$cacheKey]) {
                        $script:VendorCapWarned[$cacheKey] = $true
                        Write-Log "GARDE-FOU : le résolveur '$cacheKey' correspond à plus de 10 applications différentes — ses correspondances suivantes sont refusées (anomalie d'appariement probable ; les traces [résolution] du journal montrent les applications concernées)." -Level WARN
                    }
                    $script:LastVendorError = "refusé par le garde-fou anti-sur-appariement ('$cacheKey')"
                    return $null
                }
                $script:LastVendorKey = $cacheKey
            }
            return $cached
        }

        # --- Résolution fraîche ---
        $value = $null
        $script:LastVendorError = $null
        try {
            if ($resolver.Type -eq 'GitHub') {
                $value = Get-GitHubLatestRelease -Repo $resolver.Repo
                if (-not $value -and -not $script:LastVendorError) {
                    $script:LastVendorError = "dépôt GitHub sans version exploitable"
                }
            } else {
                $value = & $resolver.Resolve $AppName
                if (-not $value) { $script:LastVendorError = "réponse reçue mais aucune version exploitable" }
            }
        } catch {
            $script:LastVendorError = $_.Exception.Message
            Write-Log "Résolution '$($resolver.Label)' indisponible : $($_.Exception.Message)" -Level WARN
        }

        $script:VendorCache[$cacheKey]      = $value
        $script:VendorErrorCache[$cacheKey] = $script:LastVendorError

        if ($value) {
            $script:VendorKeyHits[$cacheKey] = [int]$script:VendorKeyHits[$cacheKey] + 1
            $script:LastVendorKey = $cacheKey
        }
        return $value
    }
    return $null
}

function Test-VersionResolution {
    <#
        Diagnostic complet de la résolution des versions, destiné au bouton de l'interface.
        Affiche notamment la SORTIE BRUTE de winget : c'est le seul moyen de distinguer
        un catalogue injoignable d'un simple problème de format de sortie.
    #>
    Initialize-ConsoleEncoding

    # Un diagnostic doit refléter l'état RÉEL de l'instant : on vide tous les caches,
    # sinon un échec survenu lors d'un test précédent serait rejoué à l'identique.
    $script:VendorCache        = @{}
    $script:VendorErrorCache   = @{}
    $script:VendorKeyHits      = @{}
    $script:VendorCapWarned    = @{}
    $script:GitHubReleaseCache = @{}
    $script:ChocoQueryCache    = @{}
    $script:WingetQueryCache   = @{}
    $script:GhPkgsQueryCache   = @{}
    $script:GhTreeCache        = @{}
    $script:GhLetterShaCache   = @{}
    $script:GhManifestsSha     = $null
    $script:GhPkgsDisabled     = $false
    $script:GhPkgsCallCount    = 0

    $sb = New-Object System.Text.StringBuilder

    $wg = Get-WingetExecutable
    if (-not $wg) {
        [void]$sb.AppendLine("winget : INTROUVABLE dans cette session.")
        [void]$sb.AppendLine("  winget est une application PAR UTILISATEUR (App Installer).")
        [void]$sb.AppendLine("  Causes fréquentes : console élevée sous un autre compte,")
        [void]$sb.AppendLine("  Windows Server sans App Installer, alias d'exécution désactivé.")
    } else {
        $ver = ""
        try { $ver = ((& $wg --version) 2>$null | Select-Object -First 1) } catch { }
        [void]$sb.AppendLine("winget : TROUVÉ ($ver)")
        [void]$sb.AppendLine("  $wg")

        # --- Sources déclarées : révèle une source communautaire absente ou remplacée ---
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("--- Sources déclarées (winget source list) ---")
        $srcRaw = Invoke-WingetRaw -WingetPath $wg -Arguments @('source', 'list')
        if ($srcRaw.Count -eq 0) {
            [void]$sb.AppendLine("  (aucune sortie)")
        } else {
            foreach ($l in ($srcRaw | Select-Object -First 12)) {
                if (-not [string]::IsNullOrWhiteSpace($l)) { [void]$sb.AppendLine("  $l") }
            }
        }

        # --- Recherche de test + sortie brute ---
        [void]$sb.AppendLine("")
        [void]$sb.AppendLine("--- Test de recherche sur '7zip' ---")
        $test = Invoke-WingetSearch -WingetPath $wg -Query '7zip'
        [void]$sb.AppendLine("  Commande retenue : $($script:WingetLastCommand)")
        [void]$sb.AppendLine("  Résultats exploitables : $($test.Count)")
        if ($test.Count -gt 0) {
            foreach ($t in ($test | Select-Object -First 3)) {
                [void]$sb.AppendLine("    $($t.Name) [$($t.Id)] = $($t.Version)")
            }
        } else {
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("  SORTIE BRUTE DE WINGET (à me transmettre) :")
            if (-not $script:WingetLastRaw -or $script:WingetLastRaw.Count -eq 0) {
                [void]$sb.AppendLine("    (winget n'a rien renvoyé du tout)")
            } else {
                foreach ($l in ($script:WingetLastRaw | Select-Object -First 15)) {
                    [void]$sb.AppendLine("    |$l")
                }
            }
            [void]$sb.AppendLine("")
            [void]$sb.AppendLine("  DIAGNOSTIC : le catalogue communautaire est injoignable.")
            [void]$sb.AppendLine("  PowerShell utilise le proxy systeme, winget non : c'est pourquoi")
            [void]$sb.AppendLine("  les API editeurs repondent alors que winget echoue.")
            [void]$sb.AppendLine("  Deux pistes cote infrastructure :")
            [void]$sb.AppendLine("   1. Autoriser cdn.winget.microsoft.com sur le proxy/pare-feu.")
            [void]$sb.AppendLine("   2. Declarer le proxy a winget :")
            [void]$sb.AppendLine("        winget settings --enable ProxyCommandLineOptions")
            [void]$sb.AppendLine("        winget source update --proxy http://<proxy>:<port>")
            [void]$sb.AppendLine("  En attendant, les API editeurs et le CSV alimentent le rapport.")
        }
    }

    # --- API éditeurs ---
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- API éditeurs (fonctionnent sans winget) ---")
    foreach ($probe in @('Google Chrome', 'Mozilla Firefox', 'Mozilla Thunderbird', 'Microsoft Edge', 'Notepad++', 'VLC media player', '7-Zip')) {
        $script:LastVendorError = $null
        $r = Get-VendorLatestVersion -AppName $probe
        if ($r -and $r.Version) {
            [void]$sb.AppendLine("  $probe : $($r.Version)")
        } else {
            $why = if ($script:LastVendorError) { $script:LastVendorError } else { "aucun resolveur ne correspond" }
            if ($why.Length -gt 80) { $why = $why.Substring(0, 80) + "..." }
            [void]$sb.AppendLine("  $probe : ECHEC ($why)")
            [void]$sb.AppendLine("      nom normalise teste : [$($script:LastNormalizedName)]")
        }
    }

    # --- Catalogue winget-pkgs lu sur GitHub : meme catalogue que winget, en HTTPS standard ---
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Catalogue winget-pkgs sur GitHub (api.github.com, suit le proxy systeme) ---")
    $ghToken = ""
    if ($txtGhToken) { $ghToken = $txtGhToken.Text.Trim() }
    if ([string]::IsNullOrWhiteSpace($ghToken)) { $ghToken = Get-GitHubToken }
    $script:GhPkgsToken = $ghToken
    $ghTest = Get-WingetPkgsLatestVersion -AppName '7-Zip' -Publisher 'Igor Pavlov'
    if ($ghTest -and $ghTest.Version) {
        [void]$sb.AppendLine("  OPERATIONNEL : 7-Zip = $($ghTest.Version) via $($ghTest.MatchedId)")
        $ghTest2 = Get-WingetPkgsLatestVersion -AppName 'Notepad++' -Publisher 'Don Ho'
        if ($ghTest2 -and $ghTest2.Version) { [void]$sb.AppendLine("                 Notepad++ = $($ghTest2.Version) via $($ghTest2.MatchedId)") }
        [void]$sb.AppendLine("  Appels API consommes par ce test : $($script:GhPkgsCallCount)")
        if ([string]::IsNullOrWhiteSpace($ghToken)) {
            [void]$sb.AppendLine("  SANS JETON : 60 requetes/heure, soit environ 25 applications.")
            [void]$sb.AppendLine("  Renseignez un jeton GitHub (lecture seule, sans portee) pour passer a 5000/heure.")
        } else {
            [void]$sb.AppendLine("  AVEC JETON : 5000 requetes/heure.")
        }
    } else {
        [void]$sb.AppendLine("  INJOIGNABLE ou quota epuise (voir le journal).")
        [void]$sb.AppendLine("  Domaine a autoriser : api.github.com")
    }

    # --- Catalogue Chocolatey : remplacant du catalogue winget lorsque le CDN est filtre ---
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("--- Catalogue Chocolatey (HTTPS standard, suit le proxy systeme) ---")
    $chocoTest = Invoke-ChocolateySearch -Query '7zip'
    if ($chocoTest.Count -gt 0) {
        [void]$sb.AppendLine("  OPERATIONNEL : $($chocoTest.Count) resultat(s) sur '7zip'")
        foreach ($t in ($chocoTest | Select-Object -First 3)) {
            [void]$sb.AppendLine("    $($t.Name) [$($t.Id)] = $($t.Version)")
        }
    } else {
        $why = if ($script:ChocoLastError) { $script:ChocoLastError } else { "aucun resultat" }
        if ($why.Length -gt 120) { $why = $why.Substring(0, 120) + "..." }
        [void]$sb.AppendLine("  INJOIGNABLE : $why")
        [void]$sb.AppendLine("  Domaine a autoriser : community.chocolatey.org")
    }

    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("Rappel : pilotes constructeurs (Epson, Konica, Kyocera, Lexmark...) et")
    [void]$sb.AppendLine("applications métier internes ne figurent dans AUCUN catalogue public.")
    [void]$sb.AppendLine("Pour celles-ci, seul le fichier CSV de correspondance fait foi.")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("Ce diagnostic est aussi enregistré dans le journal.")

    return $sb.ToString()
}

function Get-TenantAppInventory {
    param(
        [string]$AccessToken,
        [string]$OverrideCsvPath,
        [bool]$UseOnlineSources = $true,
        [bool]$UseGitHubPkgs = $true,
        [string]$GitHubToken = "",
        [scriptblock]$ProgressCallback,
        # Liste d'applications déjà collectée (import de fichier) : si elle est fournie,
        # aucun appel Graph n'est effectué. Les objets doivent porter les mêmes noms de
        # propriétés que Graph (displayName, publisher, displayVersion...), ce que
        # garantit Import-MobileAppsFile.
        [array]$PreloadedApps
    )

    if ($PreloadedApps -and $PreloadedApps.Count -gt 0) {
        $AllApps = $PreloadedApps
        Write-Log "Inventaire applicatif : $($AllApps.Count) application(s) issues d'un import de fichier (aucun appel Graph)." -Level OK
    } else {
        if ([string]::IsNullOrWhiteSpace($AccessToken)) {
            throw "Inventaire applicatif : aucun jeton Graph ni fichier d'import fourni."
        }
        if ($ProgressCallback) { & $ProgressCallback "Récupération de l'inventaire des applications du tenant..." }
        $Url     = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$top=999"
        $AllApps = Get-GraphPagedResults -Url $Url -AccessToken $AccessToken -ProgressCallback $ProgressCallback
        Repair-GraphDateTimeFields -Items $AllApps -FieldNames @('createdDateTime')
    }

    # ----- Fichier de correspondance manuel (source prioritaire et faisant foi) -----
    # Séparateur ';' (Excel FR) essayé en premier, puis ',' en secours.
    # Les lignes dont LatestVersion est vide (modèle pas encore complété) sont ignorées.
    # Si aucun fichier n'est sélectionné dans l'interface, on reprend automatiquement
    # le référentiel manuel : les versions saisies une fois n'ont jamais à être
    # re-sélectionnées ni re-saisies.
    if ([string]::IsNullOrWhiteSpace($OverrideCsvPath) -and (Test-Path $ManualVersionsFile)) {
        $OverrideCsvPath = $ManualVersionsFile
        Write-Log "Référentiel des versions manuelles repris automatiquement : $ManualVersionsFile"
    }

    $Overrides = @{}
    if ($OverrideCsvPath -and (Test-Path $OverrideCsvPath)) {
        foreach ($delim in @(';', ',')) {
            try {
                Import-Csv -Path $OverrideCsvPath -Delimiter $delim | ForEach-Object {
                    if ($_.AppName -and -not [string]::IsNullOrWhiteSpace($_.LatestVersion)) {
                        $Overrides[$_.AppName.Trim().ToLower()] = [PSCustomObject]@{
                            LatestVersion = $_.LatestVersion.Trim()
                            ReleaseDate   = $_.ReleaseDate
                        }
                    }
                }
            } catch {
                Write-Log "Erreur lecture du fichier de correspondance (séparateur '$delim') : $($_.Exception.Message)" -Level ERROR
            }
            if ($Overrides.Count -gt 0) { break }
        }
        Write-Log "Fichier de correspondance : $($Overrides.Count) version(s) chargée(s) depuis $OverrideCsvPath."
    }

    Initialize-ConsoleEncoding

    # Chaque génération repart de zéro : aucun résultat (ni échec) d'une exécution ou
    # d'un diagnostic précédent ne peut se rejouer par le cache.
    $script:VendorCache        = @{}
    $script:VendorErrorCache   = @{}
    $script:VendorKeyHits      = @{}
    $script:VendorCapWarned    = @{}
    $script:GitHubReleaseCache = @{}
    $script:ChocoQueryCache    = @{}
    $script:WingetQueryCache   = @{}

    # ============================================================
    # SONDES DE SANTÉ DES CATALOGUES EN LIGNE
    # Chaque catalogue est testé UNE fois sur une application très répandue (7-Zip).
    # S'il ne répond pas, il est écarté immédiatement plutôt que de faire échouer
    # des centaines de requêtes une par une (plusieurs minutes perdues).
    # ============================================================
    $wingetPath   = $null
    $chocoEnabled = $false
    $ghPkgsEnabled = $false

    if ($UseOnlineSources) {
        $wingetPath = Get-WingetExecutable
        if ($wingetPath) {
            $wgVer = ""
            try { $wgVer = ((& $wingetPath --version) 2>$null | Select-Object -First 1) } catch { }
            Write-Log "winget détecté ($wgVer) : $wingetPath" -Level OK

            if ($ProgressCallback) { & $ProgressCallback "Vérification du catalogue winget..." }
            $probe = Invoke-WingetSearch -WingetPath $wingetPath -Query '7zip'
            if ($probe.Count -eq 0) {
                $rawHead = ""
                if ($script:WingetLastRaw) { $rawHead = (($script:WingetLastRaw | Select-Object -First 6) -join ' | ') }
                Write-Log "Catalogue winget injoignable (test '7zip' sans résultat) : recherche winget désactivée pour cette exécution. Commande : $($script:WingetLastCommand). Sortie : $rawHead" -Level WARN
                $wingetPath = $null
            } else {
                Write-Log "Catalogue winget opérationnel ($($probe.Count) résultat(s) au test '7zip')." -Level OK
            }
        } else {
            Write-Log "winget introuvable dans cette session : catalogue winget non utilisé." -Level WARN
        }

        if ($UseGitHubPkgs) {
            if ($ProgressCallback) { & $ProgressCallback "Vérification du catalogue winget-pkgs sur GitHub..." }
            $ghPkgsEnabled = Test-WingetPkgsConnectivity -Token $GitHubToken
        } else {
            $script:GhPkgsDisabled = $true
            Write-Log "Catalogue winget-pkgs (GitHub) désactivé par l'option." -Level INFO
        }

        if ($ProgressCallback) { & $ProgressCallback "Vérification du catalogue Chocolatey..." }
        $chocoProbe = Invoke-ChocolateySearch -Query '7zip'
        if ($chocoProbe.Count -gt 0) {
            $chocoEnabled = $true
            Write-Log "Catalogue Chocolatey opérationnel ($($chocoProbe.Count) résultat(s) au test '7zip')." -Level OK
        } else {
            $reason = if ($script:ChocoLastError) { $script:ChocoLastError } else { "aucun résultat" }
            Write-Log "Catalogue Chocolatey injoignable ($reason) : non utilisé pour cette exécution." -Level WARN
        }
    } else {
        Write-Log "Recherche en ligne désactivée (option décochée) : seul le fichier CSV sera utilisé."
    }

    $Inventory = @()
    $csvHit = 0; $vendorHit = 0; $wgHit = 0; $ghHit = 0; $chocoHit = 0; $noMatch = 0; $notSearched = 0
    $i = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    foreach ($app in $AllApps) {
        $i++
        if ($ProgressCallback -and ($i % 5 -eq 0 -or $i -eq 1)) {
            & $ProgressCallback "Analyse des versions... ($i / $($AllApps.Count))"
        }

        # Version actuellement déployée sur Intune : le champ varie selon le type d'application
        $currentVersion = $null
        if ($app.displayVersion)                    { $currentVersion = $app.displayVersion }
        elseif ($app.msiInformation.productVersion) { $currentVersion = $app.msiInformation.productVersion }
        elseif ($app.versionNumber)                 { $currentVersion = $app.versionNumber }
        elseif ($app.version)                       { $currentVersion = $app.version }
        if ([string]::IsNullOrWhiteSpace($currentVersion)) { $currentVersion = "N/A" }

        $latestVersion = $null
        $releaseDate   = $null
        $source        = "Non recherché"

        # ============================================================
        # CASCADE DE RÉSOLUTION, de la source la plus fiable à la plus large :
        #   1. Fichier CSV      -> maîtrisé par vos équipes, fait toujours foi
        #   2. API éditeur      -> donnée officielle (Chrome, Firefox, Edge, GitHub...)
        #   3. Catalogue winget -> large, mais nécessite l'accès au CDN Microsoft
        #   4. Chocolatey       -> large, joignable en HTTPS standard via le proxy
        # ============================================================
        $key = $app.displayName.Trim().ToLower()
        if ($Overrides.ContainsKey($key)) {
            $latestVersion = $Overrides[$key].LatestVersion
            $releaseDate   = $Overrides[$key].ReleaseDate
            $source        = "Fichier manuel"
            $csvHit++
        }
        elseif (-not $UseOnlineSources) {
            $source = "Recherche désactivée"
            $notSearched++
        }
        else {
            $resolved = $false

            $vendor = Get-VendorLatestVersion -AppName $app.displayName -Publisher $app.publisher
            if ($vendor -and $vendor.Version) {
                $latestVersion = $vendor.Version
                $releaseDate   = $vendor.ReleaseDate
                $source        = "API éditeur ($($script:LastVendorKey))"
                $vendorHit++
                $resolved = $true
            }

            if (-not $resolved -and $wingetPath) {
                $wg = Get-WingetLatestVersion -WingetPath $wingetPath -AppName $app.displayName -Publisher $app.publisher
                if ($wg) {
                    $latestVersion = $wg.Version
                    $releaseDate   = $wg.ReleaseDate
                    $source        = "winget ($($wg.MatchedId))"
                    $wgHit++
                    $resolved = $true
                }
            }

            if (-not $resolved -and $ghPkgsEnabled -and -not $script:GhPkgsDisabled) {
                $gh = Get-WingetPkgsLatestVersion -AppName $app.displayName -Publisher $app.publisher
                if ($gh) {
                    $latestVersion = $gh.Version
                    $releaseDate   = $gh.ReleaseDate
                    $source        = "winget-pkgs GitHub ($($gh.MatchedId))"
                    $ghHit++
                    $resolved = $true
                }
            }

            if (-not $resolved -and $chocoEnabled) {
                $ch = Get-ChocolateyLatestVersion -AppName $app.displayName -Publisher $app.publisher
                if ($ch) {
                    $latestVersion = $ch.Version
                    $releaseDate   = $ch.ReleaseDate
                    $source        = "Chocolatey ($($ch.MatchedId))"
                    $chocoHit++
                    $resolved = $true
                }
            }

            if (-not $resolved) {
                if ($wingetPath -or $chocoEnabled -or $ghPkgsEnabled) { $source = "Aucune correspondance"; $noMatch++ }
                else                               { $source = "Catalogues indisponibles"; $notSearched++ }
            }
        }

        if ([string]::IsNullOrWhiteSpace($latestVersion)) { $latestVersion = "Non trouvé" }
        if ([string]::IsNullOrWhiteSpace($releaseDate))   { $releaseDate   = "Non disponible" }

        # Traces de résolution : les 8 premières applications systématiquement, puis
        # chaque application effectivement résolue en ligne. En cas d'anomalie (comme une
        # même version attribuée partout), le journal désigne la source ligne par ligne.
        if ($i -le 8 -or ($latestVersion -ne "Non trouvé" -and $source -ne "Fichier manuel")) {
            Write-Log "[résolution] '$($app.displayName)' -> $source : $latestVersion"
        }

        $status = "Non vérifiable"
        if ($currentVersion -ne "N/A" -and $latestVersion -ne "Non trouvé") {
            # Comparaison tolérante : "24.08" et "24.08.00.0" sont considérées identiques.
            if (Test-VersionEquality -A $currentVersion -B $latestVersion) { $status = "A jour" }
            else { $status = "A vérifier (versions différentes)" }
        }

        $Inventory += [PSCustomObject]@{
            DisplayName         = $app.displayName
            Publisher           = $app.publisher
            CurrentVersion      = $currentVersion
            LatestPublicVersion = $latestVersion
            ReleaseDate         = $releaseDate
            Source              = $source
            Status              = $status
            CreatedDateTime     = $app.createdDateTime
        }
    }

    $sw.Stop()
    $resolvedTotal = $csvHit + $vendorHit + $wgHit + $ghHit + $chocoHit
    $level = if ($resolvedTotal -gt 0) { "OK" } else { "WARN" }
    Write-Log ("Versions publiques : {0} résolue(s) sur {1} (CSV={2}, API éditeur={3}, winget={4}, winget-pkgs GitHub={5}, Chocolatey={6}) / sans correspondance={7} / non recherchées={8} - durée {9} s." -f `
        $resolvedTotal, $AllApps.Count, $csvHit, $vendorHit, $wgHit, $ghHit, $chocoHit, $noMatch, $notSearched, [Math]::Round($sw.Elapsed.TotalSeconds)) -Level $level
    if ($ghHit -gt 0 -or $script:GhPkgsCallCount -gt 0) {
        Write-Log "Catalogue GitHub : $($script:GhPkgsCallCount) appel(s) API pour $ghHit résolution(s)." -Level INFO
    }

    return $Inventory
}

# ========================================
# PAGE 4 - REMÉDIATION & SANTÉ DES POSTES
# ========================================
#
# OBJECTIF : donner de quoi AGIR AVANT l'incident — espace disque, scores de
# performance Endpoint Analytics, temps de démarrage, écrans bleus, batteries,
# appareils qui ne se synchronisent plus.
#
# SOURCES (mêmes trois modes que le reste du script : API / Import / Mixte) :
#   * Espace disque & inactivité : champs freeStorageSpaceInBytes /
#     totalStorageSpaceInBytes / lastSyncDateTime des managedDevices (déjà collectés).
#   * Scores par poste : /beta/deviceManagement/userExperienceAnalyticsDeviceScores
#   * Démarrage / BSOD :  /beta/deviceManagement/userExperienceAnalyticsDevicePerformance
#
# PERMISSIONS : les endpoints userExperienceAnalytics* sont couverts par
# DeviceManagementManagedDevices.Read.All (déjà requis). En revanche, l'Analyse des
# points de terminaison doit être ACTIVÉE dans Intune (Rapports > Analyse des points
# de terminaison) : à défaut, Graph renvoie une erreur ou des listes vides, et le
# script continue avec les seules données des appareils (disque + inactivité).

# Seuils de déclenchement des alertes — ajustez-les librement à votre contexte.
$RemediationThresholds = @{
    DiskFreePctCritical = 10     # % d'espace libre en-dessous duquel c'est critique
    DiskFreePctWarning  = 20     # % d'espace libre à surveiller
    DiskFreeGbCritical  = 5      # critique aussi si moins de N Go libres, quel que soit le %
    StaleDaysWarning    = 30     # jours sans synchronisation -> à surveiller
    StaleDaysCritical   = 90     # jours sans synchronisation -> critique
    ScoreLow            = 50     # score Endpoint Analytics global en-dessous duquel on alerte
    BootSlowSeconds     = 90     # démarrage (core boot) au-delà duquel on alerte
    BatteryPoor         = 50     # score batterie en-dessous duquel on alerte
    BsodCritical        = 3      # nombre d'écrans bleus (fenêtre 14 j) à partir duquel c'est critique
    RestartsHigh        = 15     # redémarrages (fenêtre 14 j) jugés anormalement fréquents
    SignatureStaleDays  = 7      # ancienneté des signatures antivirus (Defender) tolérée
    UptimeWarningDays   = 14     # jours depuis le dernier démarrage connu -> à surveiller (estimation)
    AppCrashWarning     = 5      # plantages d'une même application (fenêtre ~14 j) jugés anormaux
    BatteryCapacityPoor = 70     # capacité max restante (%) sous laquelle la batterie est à remplacer
    BatteryCapacityCrit = 50     # capacité max restante (%) sous laquelle le remplacement est urgent
    BatteryAgeWarnDays  = 1095   # âge de la batterie (jours) au-delà duquel on la surveille (~3 ans)
    BatteryRuntimeLowMin = 120   # autonomie estimée (minutes) sous laquelle le poste devient sédentaire
    ConfigErrorsCritical = 3     # nombre de profils de configuration en échec rendant l'alerte critique
}

# Date de fin de support de Windows 10 (toutes éditions grand public / Entreprise hors
# LTSC et hors ESU payant). Un poste resté sous Windows 10 après cette date ne reçoit plus
# de correctifs de sécurité : c'est un indicateur de parc à part entière, pas un détail.
$script:Windows10EndOfSupport = [datetime]'2025-10-14'

# Vérifications activables/désactivables indépendamment (case à cocher par vérification dans
# l'onglet "Rapports & Versions"). Une vérification désactivée n'est ni collectée, ni affichée :
# ce n'est pas un simple filtre visuel, aucun appel réseau n'est fait pour elle.
$script:RemediationChecks = [ordered]@{
    Disk           = $true   # espace disque
    Inactivity     = $true   # dernière synchronisation
    Boot           = $true   # démarrage lent / disque mécanique
    Bsod           = $true   # écrans bleus / redémarrages fréquents
    Battery        = $true   # santé de la batterie
    EaScore        = $true   # score Endpoint Analytics global
    BitLocker      = $true   # chiffrement BitLocker
    Defender       = $true   # protection antivirus Defender
    WindowsUpdate  = $true   # échecs de mises à jour qualité
    AppReliability = $true   # applications qui plantent le plus
    Uptime         = $true   # temps depuis le dernier démarrage connu (estimation)
    Compliance     = $true   # non-conformité aux stratégies de conformité Intune
    ConfigProfile  = $true   # profils de configuration en erreur ou en conflit
    BatteryDetail  = $true   # capacité, âge et autonomie de la batterie (Endpoint Analytics)
}

# ========================================
# ANONYMISATION DU RAPPORT
# ========================================
#
# Quatre familles de données peuvent être masquées, chacune par sa propre case à cocher :
#   * postes et utilisateurs      -> Poste-xxxxxxxx / User-xxxxxxxx
#   * applications                -> apps_1, apps_2... (onglets 2 et 3)
#   * exécutables qui plantent    -> proc_1, proc_2... (onglet 4)
#   * nom du client + termes liés -> X (partout dans le HTML final)
# Les pseudonymes sont STABLES le temps d'une génération : un même poste, un même
# utilisateur ou une même application porte le même libellé dans tous les onglets.
# Les tables sont remises à zéro à chaque génération, et une table de correspondance
# (confidentielle) est écrite à côté du rapport pour pouvoir retrouver les vraies valeurs.

function Reset-AnonymizationMaps {
    $script:AnonNameMap     = @{}   # nom de poste réel (minuscules) -> pseudonyme
    $script:AnonUpnMap      = @{}   # UPN réel (minuscules)          -> pseudonyme
    $script:AnonRealByAlias = @{}   # pseudonyme de poste            -> nom réel
    $script:AppAnonMap      = @{}
    $script:AppAnonList     = New-Object System.Collections.Generic.List[psobject]
    $script:ProcAnonMap     = @{}
    $script:ProcAnonList    = New-Object System.Collections.Generic.List[psobject]
}

function Get-AnonymizedIdentity {
    <#
        Pseudonyme STABLE pour un poste et pour un utilisateur le temps d'une génération :
        le même appareil porte ainsi le même "Poste-xxxxxxxx" dans les onglets Conformité,
        Applications découvertes et Remédiation. Poste et utilisateur sont pseudonymisés
        séparément : un UPN absent reste absent au lieu de recevoir un faux utilisateur.
    #>
    param([string]$RealName, [string]$RealUpn)
    if ($null -eq $script:AnonNameMap) { Reset-AnonymizationMaps }

    $alias = ""
    if (-not [string]::IsNullOrWhiteSpace($RealName)) {
        $k = $RealName.Trim().ToLower()
        if (-not $script:AnonNameMap.ContainsKey($k)) {
            do { $candidate = "Poste-" + ([guid]::NewGuid().ToString().Substring(0, 8)) } while ($script:AnonRealByAlias.ContainsKey($candidate))
            $script:AnonNameMap[$k] = [PSCustomObject]@{ Alias = $candidate; Real = $RealName.Trim() }
            $script:AnonRealByAlias[$candidate] = $RealName.Trim()
        }
        $alias = $script:AnonNameMap[$k].Alias
    }

    $upnAlias = ""
    if (-not [string]::IsNullOrWhiteSpace($RealUpn)) {
        $k = $RealUpn.Trim().ToLower()
        if (-not $script:AnonUpnMap.ContainsKey($k)) {
            $script:AnonUpnMap[$k] = [PSCustomObject]@{ Alias = "User-" + ([guid]::NewGuid().ToString().Substring(0, 8)); Real = $RealUpn.Trim() }
        }
        $upnAlias = $script:AnonUpnMap[$k].Alias
    }

    return [PSCustomObject]@{ Name = $alias; Upn = $upnAlias }
}

function Get-RealDeviceName {
    <# Pseudonyme -> nom réel (ou la valeur reçue si ce n'est pas un pseudonyme connu).
       Sert aux jointures internes effectuées APRÈS l'anonymisation de l'onglet 1. #>
    param([string]$Name)
    if ($script:AnonRealByAlias -and $Name -and $script:AnonRealByAlias.ContainsKey($Name)) { return $script:AnonRealByAlias[$Name] }
    return $Name
}

function Get-AnonymizedLabel {
    <# Libellé neutre et stable pour une application (apps_N) ou un exécutable (proc_N). #>
    param([string]$RealValue, [ValidateSet('App', 'Proc')][string]$Kind = 'App')
    if ([string]::IsNullOrWhiteSpace($RealValue)) { return $RealValue }
    if ($null -eq $script:AppAnonMap -or $null -eq $script:ProcAnonMap) { Reset-AnonymizationMaps }
    if ($Kind -eq 'App') { $map = $script:AppAnonMap;  $list = $script:AppAnonList;  $prefix = 'apps_' }
    else                 { $map = $script:ProcAnonMap; $list = $script:ProcAnonList; $prefix = 'proc_' }
    $k = $RealValue.Trim().ToLower()
    if (-not $map.ContainsKey($k)) {
        $label = $prefix + ($list.Count + 1)
        $map[$k] = $label
        $list.Add([PSCustomObject]@{ Label = $label; Real = $RealValue.Trim() })
    }
    return $map[$k]
}

function ConvertTo-AnonymizedInventory {
    <#
        Copie de l'inventaire (onglet 3) avec des noms d'applications neutres. Les objets
        d'origine ne sont PAS modifiés : ils servent encore au modèle CSV de versions et à
        la fenêtre "Versions manuelles", qui ont besoin des noms réels.
    #>
    param([array]$Inventory)
    $out = New-Object System.Collections.Generic.List[psobject]
    foreach ($item in @($Inventory)) {
        if ($null -eq $item) { continue }
        $copy  = $item.PSObject.Copy()
        $label = Get-AnonymizedLabel -RealValue ([string]$item.DisplayName) -Kind App
        $copy.DisplayName = $label
        # La source cite l'identifiant du paquet trouvé (winget, GitHub, Chocolatey, API
        # éditeur) : il trahirait le nom réel de l'application.
        if (-not [string]::IsNullOrWhiteSpace([string]$copy.Source)) {
            $copy.Source = ([string]$copy.Source) -replace '\([^)]*\)', "($label)"
        }
        $out.Add($copy)
    }
    return $out.ToArray()
}

function ConvertTo-AnonymizedDiscoveredApps {
    <# Copie du rapport des applications découvertes (onglet 2) avec des noms neutres. #>
    param($Report)
    if ($null -eq $Report -or $null -eq $Report.AllApps) { return $Report }
    $apps = New-Object System.Collections.Generic.List[psobject]
    foreach ($a in @($Report.AllApps)) {
        if ($null -eq $a) { continue }
        $copy = $a.PSObject.Copy()
        $prop = $copy.PSObject.Properties['displayName']
        if ($prop) { $prop.Value = Get-AnonymizedLabel -RealValue ([string]$prop.Value) -Kind App }
        $apps.Add($copy)
    }
    $clone = $Report.PSObject.Copy()
    $clone.AllApps = $apps.ToArray()
    return $clone
}

function Split-AnonymizationTerms {
    <# "Allianz; AZ ;MOPAZ" -> @('Allianz','AZ','MOPAZ') (séparateurs : ';' ou retour à la ligne). #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $seen = @{}
    $out  = New-Object System.Collections.Generic.List[string]
    foreach ($t in ($Text -split '[;\r\n]+')) {
        $v = $t.Trim()
        if ($v.Length -lt 2) { continue }   # un seul caractère détruirait le rapport
        if ($seen.ContainsKey($v.ToLower())) { continue }
        $seen[$v.ToLower()] = $true
        $out.Add($v)
    }
    return $out.ToArray()
}

function Protect-HtmlSensitiveTerms {
    <#
        Remplace le nom du client et les termes sensibles dans le HTML FINAL : titre,
        en-tête, noms de stratégies et de profils, éditeurs, noms de paquets, attributs de
        recherche et d'export... Les blocs <style> et <script> ne sont jamais touchés, pour
        ne pas casser la mise en page ni le JavaScript.
        [V2.2] Exception : les blocs de DONNÉES <script type='application/json'> (postes par
        application) sont traités, mais uniquement dans leurs chaînes JSON : un nom de poste
        ou un UPN contenant le nom du client y est masqué comme ailleurs, sans jamais toucher
        à la structure (clés, numéros, ponctuation).
          * terme de 4 caractères ou plus : remplacé partout, y compris au sein d'un mot
            ("Allianz_CitrixWeb" -> "X_CitrixWeb") ;
          * terme plus court (sigle) : remplacé uniquement s'il forme un mot entier, pour
            ne pas abîmer les mots qui le contiennent ("AZ" ne touche pas "Amazon").
        Casse ignorée ; une occurrence tout en minuscules est remplacée en minuscules.
        -WholeWordTerms : termes toujours traités en mot entier, quelle que soit leur
        longueur (le nom du profil client, qui peut être un mot courant).
    #>
    param([string]$Html, [string[]]$Terms, [string[]]$WholeWordTerms, [string]$Replacement = "X")
    if ([string]::IsNullOrEmpty($Html)) { return $Html }
    # Variante encodée HTML (ex. "A&B" -> "A&amp;B"), puis les plus longs d'abord :
    # "MOPAZ" doit être traité avant "AZ".
    $variants = @{}
    foreach ($entry in @(@($Terms | ForEach-Object { @{ T = $_; W = $false } }) + @($WholeWordTerms | ForEach-Object { @{ T = $_; W = $true } }))) {
        if ($null -eq $entry -or [string]::IsNullOrWhiteSpace([string]$entry.T)) { continue }
        $t = ([string]$entry.T).Trim()
        if ($t.Length -lt 2) { continue }
        foreach ($v in @($t, (ConvertTo-HtmlSafe $t))) {
            $key = $v.ToLower()
            $whole = ($entry.W -or $v.Length -lt 4)
            # Un même terme déclaré des deux façons : la règle la plus large l'emporte.
            if ($variants.ContainsKey($key)) { $whole = ($whole -and $variants[$key].W) }
            $variants[$key] = [PSCustomObject]@{ T = $v; W = $whole }
        }
    }
    if ($variants.Count -eq 0) { return $Html }
    $ordered = @($variants.Values | Sort-Object -Property { $_.T.Length } -Descending)

    $lower = $Replacement.ToLowerInvariant()
    $evaluator = [System.Text.RegularExpressions.MatchEvaluator]{
        param($m)
        $v = $m.Value
        if ($v -ceq $v.ToLowerInvariant() -and $v -cne $v.ToUpperInvariant()) { return $lower }
        return $Replacement
    }.GetNewClosure()

    $regexes = foreach ($o in $ordered) {
        $esc = [regex]::Escape($o.T)
        $pattern = if ($o.W) { "(?<![\p{L}\p{N}])$esc(?![\p{L}\p{N}])" } else { $esc }
        New-Object System.Text.RegularExpressions.Regex($pattern, ([System.Text.RegularExpressions.RegexOptions]'IgnoreCase, CultureInvariant'))
    }

    $parts = [regex]::Split($Html, '(?is)(<style\b.*?</style>|<script\b.*?</script>)')
    $sb = New-Object System.Text.StringBuilder ($Html.Length)
    $total = 0
    # [V2.2] Remplacement limité aux chaînes JSON (entre guillemets doubles) des blocs de données
    $jsonCount = @{ N = 0 }
    $jsonLiteral = [System.Text.RegularExpressions.MatchEvaluator]{
        param($lit)
        $text = $lit.Value
        foreach ($rx in $regexes) {
            $jsonCount.N += $rx.Matches($text).Count
            $text = $rx.Replace($text, $evaluator)
        }
        return $text
    }.GetNewClosure()
    foreach ($part in $parts) {
        if ($part -match '^(?is)<script\b[^>]*\btype\s*=\s*[''"]application/json[''"]') {
            [void]$sb.Append([regex]::Replace($part, '"(?:[^"\\]|\\.)*"', $jsonLiteral))
            continue
        }
        if ($part -match '^(?is)<(style|script)\b') { [void]$sb.Append($part); continue }
        $seg = $part
        foreach ($rx in $regexes) {
            $total += $rx.Matches($seg).Count
            $seg = $rx.Replace($seg, $evaluator)
        }
        [void]$sb.Append($seg)
    }
    $total += $jsonCount.N
    Write-Log "Anonymisation : $total occurrence(s) de terme(s) sensible(s) remplacée(s) par '$Replacement'." -Level INFO
    return $sb.ToString()
}

function Export-AnonymizationMap {
    <#
        Table de correspondance pseudonyme -> valeur réelle. CONFIDENTIELLE : elle annule
        l'anonymisation et ne doit jamais accompagner le rapport.
    #>
    param([string]$Path, [string]$ClientName, [string]$ClientLabel, [string[]]$Terms)
    $rows = New-Object System.Collections.Generic.List[psobject]
    if ($ClientName -and $ClientLabel) {
        $rows.Add([PSCustomObject]@{ Type = 'Client'; Libelle = $ClientLabel; ValeurReelle = $ClientName })
    }
    foreach ($t in @($Terms)) {
        if (-not [string]::IsNullOrWhiteSpace($t)) { $rows.Add([PSCustomObject]@{ Type = 'Terme masque'; Libelle = $ClientLabel; ValeurReelle = $t }) }
    }
    if ($script:AnonNameMap) {
        foreach ($v in @($script:AnonNameMap.Values | Sort-Object Alias)) { $rows.Add([PSCustomObject]@{ Type = 'Poste'; Libelle = $v.Alias; ValeurReelle = $v.Real }) }
    }
    if ($script:AnonUpnMap) {
        foreach ($v in @($script:AnonUpnMap.Values | Sort-Object Alias)) { $rows.Add([PSCustomObject]@{ Type = 'Utilisateur'; Libelle = $v.Alias; ValeurReelle = $v.Real }) }
    }
    if ($script:AppAnonList) {
        foreach ($e in $script:AppAnonList) { $rows.Add([PSCustomObject]@{ Type = 'Application'; Libelle = $e.Label; ValeurReelle = $e.Real }) }
    }
    if ($script:ProcAnonList) {
        foreach ($e in $script:ProcAnonList) { $rows.Add([PSCustomObject]@{ Type = 'Executable'; Libelle = $e.Label; ValeurReelle = $e.Real }) }
    }
    if ($rows.Count -eq 0) { return $null }
    $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Log "Table de correspondance d'anonymisation écrite (CONFIDENTIELLE) : $Path ($($rows.Count) ligne(s))." -Level WARN
    return $Path
}

function Get-AnonymizationTerms {
    <# Termes à masquer mémorisés pour un client (chaîne "terme1;terme2"). #>
    param([string]$ClientName, [string]$Path = $AnonTermsFile)
    if ([string]::IsNullOrWhiteSpace($ClientName) -or -not (Test-Path $Path)) { return "" }
    try {
        $data = Get-Content -Path $Path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
        if ($data) {
            $p = $data.PSObject.Properties[$ClientName]
            if ($p) { return [string]$p.Value }
        }
    } catch {
        Write-Log "Lecture des termes d'anonymisation impossible : $($_.Exception.Message)" -Level WARN
    }
    return ""
}

function Save-AnonymizationTerms {
    <# Mémorise (ou efface si vide) les termes à masquer pour un client. #>
    param([string]$ClientName, [string]$Terms, [string]$Path = $AnonTermsFile)
    if ([string]::IsNullOrWhiteSpace($ClientName)) { return }
    try {
        $table = [ordered]@{}
        if (Test-Path $Path) {
            $data = Get-Content -Path $Path -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
            if ($data) { foreach ($p in $data.PSObject.Properties) { $table[$p.Name] = [string]$p.Value } }
        }
        $clean = (@(Split-AnonymizationTerms $Terms) -join ';')
        if ([string]::IsNullOrWhiteSpace($clean)) {
            if (-not $table.Contains($ClientName)) { return }
            $table.Remove($ClientName)
        } else {
            if ($table.Contains($ClientName) -and $table[$ClientName] -eq $clean) { return }
            $table[$ClientName] = $clean
        }
        $folder = Split-Path -Path $Path -Parent
        if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
        [System.IO.File]::WriteAllText($Path, ($table | ConvertTo-Json), (New-Object System.Text.UTF8Encoding($false)))
    } catch {
        Write-Log "Enregistrement des termes d'anonymisation impossible : $($_.Exception.Message)" -Level WARN
    }
}

function ConvertTo-ImportDouble {
    <# "75,5" (FR), "75.5", "-1" -> [double] ; vide / "N/A" -> $null. #>
    param($Value)
    $s = ([string]$Value).Trim() -replace '\s', ''
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $s = $s -replace ',', '.'
    $out = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$out)) { return $out }
    return $null
}

function ConvertTo-StorageBytes {
    <#
        Convertit une valeur de stockage en octets. Graph fournit des octets ; les
        exports du portail Intune fournissent des MÉGAOCTETS. Heuristique : une
        valeur >= 10 000 000 est déjà en octets (aucun export MB n'atteint 10 To),
        en-dessous elle est traitée comme des Mo.
    #>
    param($Value)
    $d = ConvertTo-ImportDouble $Value
    if ($null -eq $d -or $d -lt 0) { return $null }
    if ($d -ge 10000000) { return [long]$d }
    return [long]($d * 1MB)
}

function ConvertTo-SafeDateString {
    <#
        Neutralise un piège fondamental : Invoke-RestMethod utilise ConvertFrom-Json en
        interne, qui détecte automatiquement les valeurs JSON qui RESSEMBLENT à une date
        ISO 8601 et les convertit en objet [datetime] .NET — jamais conservées en texte.
        Un [datetime] .NET restitué plus tard via un cast [string] (implicite ou explicite,
        et il y en a partout dans ce script : logs, exports, comparaisons...) est alors
        reformaté selon la CULTURE RÉGIONALE COURANTE du poste qui exécute le script à cet
        instant précis — un format différent selon la machine, le compte de service ou le
        contexte de tâche planifiée, sans lien garanti avec la locale d'affichage de Windows.
        Résultat sans ce garde-fou : des dizaines de milliers de postes réellement
        synchronisés récemment peuvent se retrouver affichés comme inactifs depuis des mois,
        jour et mois ayant été intervertis silencieusement quelque part dans le pipeline.

        Cette fonction fige la valeur en ISO 8601 (format "o", round-trip, non ambigu et
        indépendant de toute culture) DÈS la réception de la réponse Graph, avant que quoi
        que ce soit d'autre n'y touche. Les valeurs déjà tex­tuelles (import CSV/JSON,
        exports antérieurs) sont laissées telles quelles : ConvertTo-ImportDate sait déjà
        les analyser explicitement en dd/MM/yyyy.
    #>
    param($Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [datetime]) { return $Value.ToString("o") }
    return [string]$Value
}

function Repair-GraphDateTimeFields {
    <#
        Applique ConvertTo-SafeDateString à une liste de champs nommés, sur toute une
        collection d'objets renvoyés par Graph — à appeler juste après chaque appel API
        qui peut renvoyer des champs de date, avant toute autre utilisation.
    #>
    param([array]$Items, [string[]]$FieldNames)
    foreach ($it in $Items) {
        if ($null -eq $it) { continue }
        foreach ($f in $FieldNames) {
            $prop = $it.PSObject.Properties[$f]
            if ($prop) { $prop.Value = ConvertTo-SafeDateString $prop.Value }
        }
    }
}

function ConvertTo-ImportDate {
    <# Date Graph ISO ou date FR d'export ("12/03/2026 09:12") -> [datetime] ; sinon $null. #>
    param($Value)
    $s = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $dt  = [datetime]::MinValue
    foreach ($f in @('dd/MM/yyyy HH:mm:ss', 'dd/MM/yyyy HH:mm', 'dd/MM/yyyy')) {
        if ([datetime]::TryParseExact($s, $f, $inv, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt }
    }
    if ([datetime]::TryParse($s, $inv, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt }
    if ([datetime]::TryParse($s, [ref]$dt)) { return $dt }
    return $null
}

function Get-EndpointAnalyticsScores {
    <# Scores par poste (Endpoint Analytics). Best-effort : liste vide si le service est inactif. #>
    param([Parameter(Mandatory = $true)][string]$AccessToken, [scriptblock]$ProgressCallback)
    try {
        $url = "https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsDeviceScores?`$top=999"
        return @(Get-GraphPagedResults -Url $url -AccessToken $AccessToken -ProgressCallback $ProgressCallback)
    } catch {
        Write-Log "Scores Endpoint Analytics indisponibles : $($_.Exception.Message)" -Level WARN
        Write-Log "  L'Analyse des points de terminaison doit être activée dans Intune pour renvoyer des données." -Level WARN
        return @()
    }
}

function Get-StartupPerformance {
    <# Performance de démarrage par poste (Endpoint Analytics). Best-effort également. #>
    param([Parameter(Mandatory = $true)][string]$AccessToken, [scriptblock]$ProgressCallback)
    try {
        $url = "https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsDevicePerformance?`$top=999"
        return @(Get-GraphPagedResults -Url $url -AccessToken $AccessToken -ProgressCallback $ProgressCallback)
    } catch {
        Write-Log "Performances de démarrage Endpoint Analytics indisponibles : $($_.Exception.Message)" -Level WARN
        return @()
    }
}

function Import-DeviceScoresFile {
    <#
        Import des scores par poste : export "Device scores" du portail (Rapports >
        Analyse des points de terminaison), JSON Graph, ou fichier produit par ce script.
        Objets renvoyés aux noms de propriétés Graph.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $rows = Get-ImportRows -Path $Path -Label "scores de santé"
    if (@($rows).Count -eq 0) { return @() }
    $map  = New-ColumnMap -Rows $rows
    $list = New-Object System.Collections.Generic.List[psobject]
    foreach ($r in $rows) {
        $name = [string](Get-MappedValue $r $map @('deviceName','device name','nom de l''appareil','hostname','computername','machine','name'))
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $list.Add([PSCustomObject]@{
            deviceName              = $name.Trim()
            endpointAnalyticsScore  = ConvertTo-ImportDouble (Get-MappedValue $r $map @('endpointAnalyticsScore','endpoint analytics score','score global','device score','score'))
            startupPerformanceScore = ConvertTo-ImportDouble (Get-MappedValue $r $map @('startupPerformanceScore','startup performance score','score de demarrage'))
            appReliabilityScore     = ConvertTo-ImportDouble (Get-MappedValue $r $map @('appReliabilityScore','app reliability score','fiabilite des applications'))
            batteryHealthScore      = ConvertTo-ImportDouble (Get-MappedValue $r $map @('batteryHealthScore','battery health score','sante de la batterie','batterie'))
            workFromAnywhereScore   = ConvertTo-ImportDouble (Get-MappedValue $r $map @('workFromAnywhereScore','work from anywhere score','travail de n''importe ou'))
        })
    }
    Write-Log "Import scores de santé : $($list.Count) poste(s)." -Level OK
    return $list.ToArray()
}

function Import-StartupPerfFile {
    <#
        Import des performances de démarrage : export "Startup performance" du portail
        ou JSON Graph. Les temps peuvent être fournis en millisecondes ou en secondes :
        une valeur < 1000 est traitée comme des secondes (aucun démarrage réel ne dure
        moins d'une seconde), au-delà comme des millisecondes.
    #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $rows = Get-ImportRows -Path $Path -Label "performances de démarrage"
    if (@($rows).Count -eq 0) { return @() }
    $map  = New-ColumnMap -Rows $rows

    $toMs = {
        param($v)
        $d = ConvertTo-ImportDouble $v
        if ($null -eq $d -or $d -lt 0) { return $null }
        if ($d -lt 1000) { return [int]($d * 1000) }
        return [int]$d
    }

    $list = New-Object System.Collections.Generic.List[psobject]
    foreach ($r in $rows) {
        $name = [string](Get-MappedValue $r $map @('deviceName','device name','nom de l''appareil','hostname','computername','machine','name'))
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $bsod = ConvertTo-ImportDouble (Get-MappedValue $r $map @('blueScreenCount','blue screen count','ecrans bleus','bsod'))
        $rst  = ConvertTo-ImportDouble (Get-MappedValue $r $map @('restartCount','restart count','redemarrages','restarts'))
        $list.Add([PSCustomObject]@{
            deviceName       = $name.Trim()
            coreBootTimeInMs = & $toMs (Get-MappedValue $r $map @('coreBootTimeInMs','core boot time in ms','core boot time','boot time','temps de demarrage'))
            coreLoginTimeInMs = & $toMs (Get-MappedValue $r $map @('coreLoginTimeInMs','core login time in ms','core login time','login time','temps de connexion'))
            bootScore        = ConvertTo-ImportDouble (Get-MappedValue $r $map @('bootScore','boot score','score de demarrage'))
            loginScore       = ConvertTo-ImportDouble (Get-MappedValue $r $map @('loginScore','login score','score de connexion'))
            blueScreenCount  = if ($null -ne $bsod) { [int]$bsod } else { $null }
            restartCount     = if ($null -ne $rst)  { [int]$rst }  else { $null }
            diskType         = [string](Get-MappedValue $r $map @('diskType','disk type','type de disque'))
        })
    }
    Write-Log "Import performances de démarrage : $($list.Count) poste(s)." -Level OK
    return $list.ToArray()
}

# ========================================
# PAGE 4 (SUITE) - SÉCURITÉ AVANCÉE, MISES À JOUR & FIABILITÉ DES APPLICATIONS
# ========================================
#
# Chaque vérification est en mode "meilleur effort" : une erreur ou un tenant sans la
# licence/fonctionnalité correspondante (Analyse des points de terminaison, Update Rings...)
# se traduit par une liste vide et un message dans le journal — jamais par l'arrêt du script.
# Chaque vérification est aussi individuellement activable/désactivable via
# $script:RemediationChecks : si elle est décochée, la fonction n'est même pas appelée.

function Get-BitLockerEncryptionStates {
    <#
        État de chiffrement BitLocker de tout le parc en UN SEUL appel (liste, pas de requête
        par poste). Renvoie les objets bruts Graph ; le rapprochement avec les appareils se
        fait par nom (deviceName) dans Build-RemediationData.
    #>
    param([Parameter(Mandatory = $true)][string]$AccessToken, [scriptblock]$ProgressCallback)
    try {
        if ($ProgressCallback) { & $ProgressCallback "Récupération de l'état BitLocker du parc..." }
        $url = "https://graph.microsoft.com/beta/deviceManagement/managedDeviceEncryptionStates?`$top=999"
        return @(Get-GraphPagedResults -Url $url -AccessToken $AccessToken -ProgressCallback $ProgressCallback)
    } catch {
        Write-Log "État BitLocker indisponible : $($_.Exception.Message)" -Level WARN
        return @()
    }
}

function Get-DefenderProtectionStates {
    <#
        État Defender (protection temps réel, ancienneté des signatures...) : PAS de liste
        globale côté Graph, une requête par poste — regroupées via le même mécanisme $batch
        déjà utilisé pour le détail de conformité (mêmes retries 429/5xx).
    #>
    param([Parameter(Mandatory = $true)][array]$DeviceIds, [Parameter(Mandatory = $true)][string]$AccessToken, [scriptblock]$ProgressCallback)
    if ($DeviceIds.Count -eq 0) { return @{} }
    try {
        if ($ProgressCallback) { & $ProgressCallback "Récupération de l'état Defender ($($DeviceIds.Count) appareils)..." }
        $requests = @()
        foreach ($id in $DeviceIds) {
            $requests += @{ id = $id; method = "GET"; url = "/deviceManagement/managedDevices/$id/windowsProtectionState" }
        }
        $responses = Invoke-GraphBatch -Requests $requests -AccessToken $AccessToken -ProgressCallback $ProgressCallback
        $byDevice = @{}
        foreach ($r in $responses) {
            # Un poste non-Windows ou sans agent Defender renvoie une erreur (404/400) : normal, on l'ignore.
            if ($r.status -eq 200 -and $r.body) {
                Repair-GraphDateTimeFields -Items @($r.body) -FieldNames @('lastReportedDateTime', 'signatureUpdateDateTime')
                $byDevice[$r.id] = $r.body
            }
        }
        return $byDevice
    } catch {
        Write-Log "État Defender indisponible : $($_.Exception.Message)" -Level WARN
        return @{}
    }
}

function Get-AppReliabilityData {
    <#
        Fiabilité des applications (Endpoint Analytics "App Health") : vue d'ensemble par
        application (pour identifier les applications à surveiller côté parc), et détail par
        poste (pour rattacher un plantage à un appareil précis).
    #>
    param([Parameter(Mandatory = $true)][string]$AccessToken, [scriptblock]$ProgressCallback)
    $result = [PSCustomObject]@{ ByApp = @(); ByDevice = @() }
    try {
        if ($ProgressCallback) { & $ProgressCallback "Récupération de la fiabilité des applications (Endpoint Analytics)..." }
        $urlApp = "https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsAppHealthApplicationPerformance?`$top=999"
        $result.ByApp = @(Get-GraphPagedResults -Url $urlApp -AccessToken $AccessToken -ProgressCallback $ProgressCallback)
    } catch {
        Write-Log "Fiabilité des applications (vue globale) indisponible : $($_.Exception.Message)" -Level WARN
    }
    try {
        $urlDev = "https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsAppHealthDevicePerformanceDetails?`$top=999"
        $result.ByDevice = @(Get-GraphPagedResults -Url $urlDev -AccessToken $AccessToken -ProgressCallback $ProgressCallback)
        Repair-GraphDateTimeFields -Items $result.ByDevice -FieldNames @('eventDateTime')
    } catch {
        Write-Log "Fiabilité des applications (détail par poste) indisponible : $($_.Exception.Message)" -Level WARN
    }
    return $result
}

function Get-DeviceStartupHistory {
    <#
        Historique de démarrage (Endpoint Analytics) : sert d'ESTIMATION du temps écoulé
        depuis le dernier redémarrage connu. Donnée agrégée à un rythme quotidien côté
        Microsoft : ce n'est pas un compteur d'uptime en direct, d'où le libellé "estimation"
        conservé jusque dans le rapport final.
    #>
    param([Parameter(Mandatory = $true)][string]$AccessToken, [scriptblock]$ProgressCallback)
    try {
        if ($ProgressCallback) { & $ProgressCallback "Récupération de l'historique de démarrage (estimation de l'uptime)..." }
        $url = "https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsDeviceStartupHistory?`$top=999"
        $rows = @(Get-GraphPagedResults -Url $url -AccessToken $AccessToken -ProgressCallback $ProgressCallback)
        Repair-GraphDateTimeFields -Items $rows -FieldNames @('startupDateTime', 'eventDateTime', 'lastBootUpTime')
        return $rows
    } catch {
        Write-Log "Historique de démarrage indisponible : $($_.Exception.Message)" -Level WARN
        return @()
    }
}

function Get-WindowsUpdateFailures {
    <#
        Échecs de mise à jour qualité, via le mécanisme générique de rapports Intune
        (job asynchrone : on le crée, on l'interroge jusqu'à complétion, on télécharge).
        PRÉREQUIS : les postes doivent être rattachés à une stratégie "Anneaux de mise à
        jour Windows" (Update Rings) — sans quoi ce rapport est vide, ce qui n'est pas un
        dysfonctionnement du script.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [scriptblock]$ProgressCallback,
        [int]$MaxWaitSeconds = 90
    )
    try {
        if ($ProgressCallback) { & $ProgressCallback "Lancement du rapport des échecs de mise à jour Windows..." }
        $body = @{ reportName = "QualityUpdateDeviceStatusByPolicy"; format = "json" } | ConvertTo-Json
        $job  = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs" `
                    -Headers @{ Authorization = "Bearer $AccessToken"; "Content-Type" = "application/json" } -Body $body -ErrorAction Stop

        $waited = 0
        $status = $job
        while ($status.status -ne "completed" -and $waited -lt $MaxWaitSeconds) {
            Start-Sleep -Seconds 5
            $waited += 5
            if ($ProgressCallback) { & $ProgressCallback "Génération du rapport de mises à jour en cours ($waited s)..." }
            $status = Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$($job.id)')" `
                        -Headers @{ Authorization = "Bearer $AccessToken" } -ErrorAction Stop
        }
        if ($status.status -ne "completed" -or -not $status.url) {
            Write-Log "Rapport des mises à jour Windows non abouti après $waited s (statut '$($status.status)') : ignoré pour cette génération." -Level WARN
            return @()
        }

        $zipPath = Join-Path $env:TEMP ("wu-report-{0}.zip" -f ([guid]::NewGuid().ToString().Substring(0,8)))
        Invoke-WebRequest -Uri $status.url -OutFile $zipPath -ErrorAction Stop
        $extractFolder = Join-Path $env:TEMP ("wu-report-{0}" -f ([guid]::NewGuid().ToString().Substring(0,8)))
        Expand-Archive -Path $zipPath -DestinationPath $extractFolder -Force
        $jsonFile = Get-ChildItem -Path $extractFolder -Filter "*.json" -Recurse | Select-Object -First 1
        $rows = @()
        if ($jsonFile) {
            # Deux formes possibles selon le rapport : tableau d'objets, ou { Schema, Values }
            # avec des lignes positionnelles. ConvertFrom-IntuneReportPayload normalise.
            $rows = @(ConvertFrom-IntuneReportPayload -Payload (Get-Content -Path $jsonFile.FullName -Raw | ConvertFrom-Json))
        }
        Remove-Item -Path $zipPath, $extractFolder -Recurse -Force -ErrorAction SilentlyContinue
        return $rows
    } catch {
        Write-Log "Rapport des mises à jour Windows indisponible : $($_.Exception.Message)" -Level WARN
        Write-Log "  Vérifiez que les postes sont rattachés à une stratégie 'Anneaux de mise à jour Windows'." -Level WARN
        return @()
    }
}

function Get-BatteryHealthDetails {
    <#
        Santé fine des batteries (Endpoint Analytics > Santé de la batterie) : capacité
        maximale restante, autonomie estimée, âge. Le batteryHealthScore déjà collecté par
        Get-EndpointAnalyticsScores est un indice composite de 0 à 100 ; ce sont ces
        trois grandeurs-ci qui permettent de DÉCIDER un remplacement et de le justifier
        auprès d'un utilisateur ou d'un acheteur.

        PRÉREQUIS : Endpoint Analytics activée, et parc comportant des appareils portables
        (un parc 100 % fixe renvoie légitimement une liste vide).
    #>
    param([Parameter(Mandatory = $true)][string]$AccessToken, [scriptblock]$ProgressCallback)
    try {
        if ($ProgressCallback) { & $ProgressCallback "Récupération de la santé des batteries (Endpoint Analytics)..." }
        $url = "https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsBatteryHealthDevicePerformance?`$top=999"
        return @(Get-GraphPagedResults -Url $url -AccessToken $AccessToken -ProgressCallback $ProgressCallback -Label "Santé des batteries")
    } catch {
        Write-Log "Santé des batteries indisponible : $($_.Exception.Message)" -Level WARN
        return @()
    }
}

function Get-ConfigurationProfileErrors {
    <#
        Postes en ERREUR ou en CONFLIT d'application d'un profil de configuration — la
        panne la plus silencieuse d'Intune : le poste reste "conforme", se synchronise
        normalement, et n'applique pourtant pas le paramétrage attendu.

        Couvre les DEUX familles de profils, qu'il serait trompeur de ne traiter qu'à moitié :
          * deviceConfigurations  : profils historiques (modèles) ;
          * configurationPolicies : catalogue de paramètres (Settings Catalog).

        STRATÉGIE : interroger les statuts PAR PROFIL (quelques dizaines d'appels) et non
        par appareil (un appel par poste, soit des milliers pour la même information).
    #>
    param(
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [scriptblock]$ProgressCallback,
        [int]$MaxProfiles = 300
    )
    $configProfiles = New-Object System.Collections.Generic.List[psobject]

    try {
        $classic = @(Get-GraphPagedResults -Url "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?`$select=id,displayName&`$top=200" `
                        -AccessToken $AccessToken -ProgressCallback $ProgressCallback -Label "Profils de configuration")
        foreach ($p in $classic) {
            [void]$configProfiles.Add([PSCustomObject]@{ Id = [string]$p.id; Name = [string]$p.displayName; Segment = "deviceConfigurations" })
        }
    } catch {
        Write-Log "Liste des profils de configuration indisponible : $($_.Exception.Message)" -Level WARN
    }

    try {
        $catalog = @(Get-GraphPagedResults -Url "https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?`$select=id,name&`$top=200" `
                        -AccessToken $AccessToken -ProgressCallback $ProgressCallback -Label "Catalogue de paramètres")
        foreach ($p in $catalog) {
            [void]$configProfiles.Add([PSCustomObject]@{ Id = [string]$p.id; Name = [string]$p.name; Segment = "configurationPolicies" })
        }
    } catch {
        Write-Log "Liste des stratégies du catalogue de paramètres indisponible : $($_.Exception.Message)" -Level WARN
    }

    $profileList = @($configProfiles)
    if ($profileList.Count -eq 0) {
        Write-Log "Aucun profil de configuration exploitable : vérification 'erreurs de profil' sans objet sur ce tenant." -Level INFO
        return @()
    }
    if ($profileList.Count -gt $MaxProfiles) {
        Write-Log "Profils de configuration : $($profileList.Count) profils détectés, analyse limitée aux $MaxProfiles premiers (garde-fou de durée, ajustable via `$script:MaxConfigProfilesAnalyzed)." -Level WARN
        $profileList = $profileList[0..($MaxProfiles - 1)]
    }

    # On ne demande PAS de $select ni de $filter sur deviceStatuses : leur prise en charge
    # varie d'un tenant et d'un type de profil à l'autre, et un rejet en 400 ferait perdre
    # tout le lot. Le tri "erreur/conflit" se fait donc localement, sur des listes déjà
    # bornées par profil.
    $requests    = New-Object System.Collections.Generic.List[psobject]
    $profileById = @{}
    $n = 0
    foreach ($p in $profileList) {
        $n++
        $requestId = "cfg$n"
        $profileById[$requestId] = $p
        [void]$requests.Add(@{ id = $requestId; method = "GET"; url = "/deviceManagement/$($p.Segment)/$($p.Id)/deviceStatuses?`$top=999" })
    }

    if ($ProgressCallback) { & $ProgressCallback "Analyse de l'application des profils de configuration ($($profileList.Count) profil(s))..." }
    $responses = @()
    try {
        $responses = Invoke-GraphBatch -Requests $requests.ToArray() -AccessToken $AccessToken `
                        -ProgressCallback $ProgressCallback -Label "Profils de configuration"
    } catch {
        Write-Log "Statuts des profils de configuration indisponibles : $($_.Exception.Message)" -Level WARN
        return @()
    }

    $failures  = New-Object System.Collections.Generic.List[psobject]
    $truncated = 0
    foreach ($r in $responses) {
        $status = 0
        try { $status = [int]$r.status } catch { $status = 0 }
        if ($status -ne 200 -or -not $r.body) { continue }
        $profileInfo = $profileById[[string]$r.id]
        if (-not $profileInfo) { continue }
        if ($r.body.'@odata.nextLink') { $truncated++ }
        foreach ($st in @($r.body.value)) {
            $state = ([string]$st.status).Trim().ToLower()
            if ($state -ne 'error' -and $state -ne 'conflict') { continue }
            [void]$failures.Add([PSCustomObject]@{
                DeviceName   = [string]$st.deviceDisplayName
                UserName     = [string]$st.userName
                ProfileName  = $profileInfo.Name
                Status       = $state
                LastReported = [string]$st.lastReportedDateTime
            })
        }
    }
    if ($truncated -gt 0) {
        Write-Log "Profils de configuration : $truncated profil(s) comptent plus de 999 postes affectés — décompte partiel pour ceux-là." -Level WARN
    }
    Write-Log "Profils de configuration : $($failures.Count) application(s) en erreur ou en conflit sur $($profileList.Count) profil(s) analysé(s)." -Level INFO
    return $failures.ToArray()
}

function Invoke-DeviceRemoteAction {
    <#
        Déclenche une action Graph ciblée sur un poste précis. PAS appelée par la génération
        du rapport (celui-ci reste 100 % lecture) : cette fonction est destinée à être copiée-
        collée par l'administrateur depuis le bouton correspondant du rapport, dans une session
        PowerShell déjà connectée (variable $AccessToken déjà présente dans cette session).

        Permission requise, nettement plus sensible que le reste du script (lecture seule) :
        DeviceManagementManagedDevices.PrivilegedOperations.All — à n'accorder qu'à une
        application dédiée à cet usage, jamais à celle qui génère le rapport en routine.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$DeviceId,
        [Parameter(Mandatory = $true)][ValidateSet('Sync', 'Reboot', 'Remediate')][string]$Action,
        [string]$RemediationScriptId,
        [Parameter(Mandatory = $true)][string]$AccessToken
    )
    $base = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId"
    $uri  = switch ($Action) {
        'Sync'      { "$base/syncDevice" }
        'Reboot'    { "$base/rebootNow" }
        'Remediate' { "$base/initiateOnDemandProactiveRemediation" }
    }
    $params = @{ Method = 'POST'; Uri = $uri; Headers = @{ Authorization = "Bearer $AccessToken" }; ErrorAction = 'Stop' }
    if ($Action -eq 'Remediate') {
        if ([string]::IsNullOrWhiteSpace($RemediationScriptId)) { throw "RemediationScriptId requis pour l'action Remediate (id du deviceHealthScript à exécuter)." }
        $params.Body        = (@{ scriptPolicyId = $RemediationScriptId } | ConvertTo-Json)
        $params.ContentType = "application/json"
    }
    Invoke-RestMethod @params  # 204 No Content attendu : rien à lire dans la réponse
    Write-Log "Action '$Action' déclenchée sur l'appareil $DeviceId." -Level OK
}

function Import-BitLockerFile {
    <# Import du statut BitLocker : colonnes Device name + État de chiffrement (+ motif). #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $rows = Get-ImportRows -Path $Path -Label "état BitLocker"
    if (@($rows).Count -eq 0) { return @() }
    $map = New-ColumnMap -Rows $rows
    $list = New-Object System.Collections.Generic.List[psobject]
    foreach ($r in $rows) {
        $name = [string](Get-MappedValue $r $map @('deviceName','device name','nom de l''appareil','hostname'))
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $list.Add([PSCustomObject]@{
            deviceName           = $name.Trim()
            encryptionState      = [string](Get-MappedValue $r $map @('encryptionState','encryption state','etat de chiffrement','chiffrement'))
            advancedBitLockerStates = [string](Get-MappedValue $r $map @('advancedBitLockerStates','advanced bitlocker state','motif','reason','detail'))
        })
    }
    Write-Log "Import BitLocker : $($list.Count) poste(s)." -Level OK
    return $list.ToArray()
}

function Import-DefenderFile {
    <# Import Defender : colonnes Device name + Protection temps réel + Signatures en retard. #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $rows = Get-ImportRows -Path $Path -Label "état Defender"
    if (@($rows).Count -eq 0) { return @{} }
    $map = New-ColumnMap -Rows $rows
    $byDevice = @{}
    foreach ($r in $rows) {
        $name = [string](Get-MappedValue $r $map @('deviceName','device name','nom de l''appareil','hostname'))
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        $rt  = [string](Get-MappedValue $r $map @('realTimeProtectionEnabled','real-time protection','protection temps reel'))
        $sig = [string](Get-MappedValue $r $map @('signatureUpdateOverdue','signature update overdue','signatures en retard'))
        $byDevice[(Get-NormalizedHeader $name)] = [PSCustomObject]@{
            deviceName              = $name.Trim()
            realTimeProtectionEnabled = ($rt -match '^(true|1|oui|yes)$')
            signatureUpdateOverdue   = ($sig -match '^(true|1|oui|yes)$')
        }
    }
    Write-Log "Import Defender : $($byDevice.Count) poste(s)." -Level OK
    return $byDevice
}

function Import-AppReliabilityFile {
    <# Import fiabilité applicative : une ligne par couple poste/application en échec. #>
    param([Parameter(Mandatory = $true)][string]$Path)
    $rows = Get-ImportRows -Path $Path -Label "fiabilité des applications"
    if (@($rows).Count -eq 0) { return @() }
    $map = New-ColumnMap -Rows $rows
    $list = New-Object System.Collections.Generic.List[psobject]
    foreach ($r in $rows) {
        $name = [string](Get-MappedValue $r $map @('deviceName','device name','nom de l''appareil'))
        $app  = [string](Get-MappedValue $r $map @('appDisplayName','application name','nom de l''application'))
        if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($app)) { continue }
        $list.Add([PSCustomObject]@{
            deviceName    = $name.Trim()
            appDisplayName = $app.Trim()
            appCrashCount  = ConvertTo-ImportInt (Get-MappedValue $r $map @('appCrashCount','crash count','nombre de plantages'))
        })
    }
    Write-Log "Import fiabilité applicative : $($list.Count) ligne(s)." -Level OK
    return $list.ToArray()
}

function Build-RemediationData {
    <#
        Croise appareils + scores + performances (jointure par nom de poste, normalisée)
        et en déduit, par appareil : les indicateurs de santé, une sévérité globale
        (crit / warn / ok) et des ACTIONS DE REMÉDIATION regroupées façon "motifs de
        non-conformité" — un accordéon par action, avec la liste des postes concernés.
    #>
    param(
        [array]$Devices,
        [array]$Scores,
        [array]$Performance,
        [array]$BitLocker,
        [hashtable]$Defender,
        $AppReliability,
        [array]$StartupHistory,
        [array]$ConfigProfileErrors,
        [array]$BatteryHealth,
        [array]$UpdateFailures,
        [hashtable]$ComplianceReasons,
        [bool]$AnonymizeData = $false,
        [bool]$AnonymizeApps = $false
    )
    $C = $script:RemediationChecks
    if (-not $C) { $C = @{} }
    # Vérification activée par défaut si absente de la table (rétro-compatibilité).
    function Test-CheckEnabled([string]$Name) { if ($C.Contains($Name)) { return [bool]$C[$Name] } else { return $true } }

    $T = $RemediationThresholds
    $scoreByName = @{}
    foreach ($s in @($Scores)) {
        if ($null -eq $s) { continue }
        $k = Get-NormalizedHeader ([string]$s.deviceName)
        if ($k -and -not $scoreByName.ContainsKey($k)) { $scoreByName[$k] = $s }
    }
    $perfByName = @{}
    foreach ($p in @($Performance)) {
        if ($null -eq $p) { continue }
        $k = Get-NormalizedHeader ([string]$p.deviceName)
        if ($k -and -not $perfByName.ContainsKey($k)) { $perfByName[$k] = $p }
    }
    $bitlockerByName = @{}
    foreach ($b in @($BitLocker)) {
        if ($null -eq $b) { continue }
        $k = Get-NormalizedHeader ([string]$b.deviceName)
        if ($k -and -not $bitlockerByName.ContainsKey($k)) { $bitlockerByName[$k] = $b }
    }
    $defenderByName = @{}
    if ($Defender) {
        foreach ($k in $Defender.Keys) {
            $d0 = $Defender[$k]
            $nk = Get-NormalizedHeader ([string]$d0.managedDeviceName)
            if (-not $nk) { $nk = Get-NormalizedHeader ([string]$k) }
            if ($nk -and -not $defenderByName.ContainsKey($nk)) { $defenderByName[$nk] = $d0 }
        }
    }
    # Fiabilité applicative : détail par poste regroupé -> pour chaque poste, l'application
    # qui plante le plus souvent (celle qui mérite l'action, pas la liste complète).
    $appCrashByName = @{}
    if ($AppReliability -and $AppReliability.ByDevice) {
        $grouped = @{}
        foreach ($e in @($AppReliability.ByDevice)) {
            if ("$($e.eventType)" -notmatch '(?i)crash|hang') { continue }
            $dn = [string](Get-PropCI $e @('deviceDisplayName','deviceName'))
            $k  = Get-NormalizedHeader $dn
            if (-not $k) { continue }
            $app = [string]$e.appDisplayName
            if (-not $grouped.ContainsKey($k)) { $grouped[$k] = @{} }
            if (-not $grouped[$k].ContainsKey($app)) { $grouped[$k][$app] = 0 }
            $grouped[$k][$app]++
        }
        foreach ($k in $grouped.Keys) {
            $top = $grouped[$k].GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
            if ($top) {
                $crashName = if ($AnonymizeApps) { Get-AnonymizedLabel -RealValue ([string]$top.Key) -Kind Proc } else { $top.Key }
                $appCrashByName[$k] = [PSCustomObject]@{ AppName = $crashName; Count = $top.Value }
            }
        }
    }
    $startupByName = @{}
    foreach ($h in @($StartupHistory)) {
        if ($null -eq $h) { continue }
        $dn = [string](Get-PropCI $h @('deviceName','deviceDisplayName'))
        $k  = Get-NormalizedHeader $dn
        $dt = ConvertTo-ImportDate ([string](Get-PropCI $h @('startupDateTime','eventDateTime','lastBootUpTime')))
        if (-not $k -or -not $dt) { continue }
        # Le plus récent événement de démarrage connu pour ce poste.
        if (-not $startupByName.ContainsKey($k) -or $dt -gt $startupByName[$k]) { $startupByName[$k] = $dt }
    }

    # ----- Profils de configuration en erreur / conflit : regroupés par poste -----
    # Un même poste peut échouer sur plusieurs profils : on garde le décompte ET les noms
    # (au-delà de trois, la liste devient illisible dans une cellule de tableau, on tronque).
    $configErrByName = @{}
    foreach ($e in @($ConfigProfileErrors)) {
        if ($null -eq $e) { continue }
        $k = Get-NormalizedHeader ([string]$e.DeviceName)
        if (-not $k) { continue }
        if (-not $configErrByName.ContainsKey($k)) {
            $configErrByName[$k] = [PSCustomObject]@{ Count = 0; Profiles = (New-Object System.Collections.Generic.List[string]); HasConflict = $false }
        }
        $entry = $configErrByName[$k]
        $entry.Count++
        if ($entry.Profiles.Count -lt 6 -and $e.ProfileName) { [void]$entry.Profiles.Add([string]$e.ProfileName) }
        if (([string]$e.Status) -eq 'conflict') { $entry.HasConflict = $true }
    }

    # ----- Santé fine des batteries -----
    $batteryByName = @{}
    foreach ($b in @($BatteryHealth)) {
        if ($null -eq $b) { continue }
        $k = Get-NormalizedHeader ([string](Get-PropCI $b @('deviceName','deviceDisplayName')))
        if ($k -and -not $batteryByName.ContainsKey($k)) { $batteryByName[$k] = $b }
    }

    # ----- Échecs de mise à jour qualité Windows -----
    # Les colonnes du rapport Intune varient selon la version du service : on cherche le nom
    # de poste et l'état d'agrégation sous plusieurs graphies plutôt que d'en figer une.
    $updateFailByName = @{}
    foreach ($u in @($UpdateFailures)) {
        if ($null -eq $u) { continue }
        $state = [string](Get-PropCI $u @('AggregateState','aggregateState','CurrentDeviceUpdateStatus','UpdateStatus','status'))
        if ($state -notmatch '(?i)error|fail|cancel') { continue }
        $k = Get-NormalizedHeader ([string](Get-PropCI $u @('DeviceName','deviceName','Device','deviceDisplayName')))
        if (-not $k -or $updateFailByName.ContainsKey($k)) { continue }
        $updateFailByName[$k] = [PSCustomObject]@{
            State   = $state
            Message = [string](Get-PropCI $u @('LatestAlertMessage','latestAlertMessage','ErrorCode','errorCode'))
        }
    }

    $now     = Get-Date
    $rows    = New-Object System.Collections.Generic.List[psobject]
    $actions = [ordered]@{}
    $sevRank = @{ ok = 0; warn = 1; crit = 2 }

    # Sources d'appareils : les managedDevices d'abord, puis les postes présents
    # uniquement dans les fichiers/flux de scores ou de performance (cas d'un import
    # partiel : les alertes score/démarrage restent produites).
    #
    # IMPORTANT : un même nom de poste peut correspondre à PLUSIEURS enregistrements
    # Intune distincts (id différents) — réimagé, ré-enrôlé, remplacement matériel
    # gardant le même hostname, ancien enregistrement jamais nettoyé après un Autopilot
    # reset... C'est très fréquent sur un grand parc. En garder arbitrairement "le
    # premier rencontré" (ordre de pagination Graph, qui n'a AUCUN rapport avec la
    # fraîcheur des données) peut faire remonter la date de synchro d'un enregistrement
    # fantôme à la place de celle du poste réellement actif portant le même nom — d'où
    # de faux positifs "inactif depuis N jours" en nombre. On garde donc, pour chaque
    # nom, l'enregistrement dont la synchronisation est la PLUS RÉCENTE.
    $byName = [ordered]@{}
    $dupCount = 0
    foreach ($d in @($Devices)) {
        if ($null -eq $d) { continue }
        $k = Get-NormalizedHeader ([string]$d.deviceName)
        if (-not $k) { continue }
        if (-not $byName.Contains($k)) {
            $byName[$k] = $d
            continue
        }
        $dupCount++
        $existingDate = ConvertTo-ImportDate ([string]$byName[$k].lastSyncDateTime)
        $candidateDate = ConvertTo-ImportDate ([string]$d.lastSyncDateTime)
        if ($null -ne $candidateDate -and ($null -eq $existingDate -or $candidateDate -gt $existingDate)) {
            $byName[$k] = $d
        }
    }
    if ($dupCount -gt 0) {
        Write-Log "Remédiation : $dupCount enregistrement(s) en doublon de nom de poste rencontré(s) — l'enregistrement le plus récemment synchronisé a été conservé pour chacun. Ces doublons méritent un nettoyage côté Intune." -Level WARN
    }

    $seen  = New-Object System.Collections.Generic.HashSet[string]
    $units = New-Object System.Collections.Generic.List[psobject]
    foreach ($k in $byName.Keys) {
        [void]$seen.Add($k)
        $units.Add([PSCustomObject]@{ Key = $k; Device = $byName[$k] })
    }
    foreach ($k in @($scoreByName.Keys) + @($perfByName.Keys)) {
        if ($seen.Add($k)) { $units.Add([PSCustomObject]@{ Key = $k; Device = $null }) }
    }

    foreach ($u in $units) {
        $d  = $u.Device
        $sc = if ($scoreByName.ContainsKey($u.Key)) { $scoreByName[$u.Key] } else { $null }
        $pf = if ($perfByName.ContainsKey($u.Key))  { $perfByName[$u.Key] }  else { $null }

        $name = ""
        if ($d)          { $name = [string]$d.deviceName }
        elseif ($sc)     { $name = [string]$sc.deviceName }
        elseif ($pf)     { $name = [string]$pf.deviceName }
        $upn  = if ($d) { [string]$d.userPrincipalName } else { "" }
        $os   = if ($d) { [string]$d.operatingSystem }   else { "" }
        $osv  = if ($d) { [string]$d.osVersion }         else { "" }
        $sync = if ($d) { [string]$d.lastSyncDateTime }  else { "" }
        $compState = if ($d) { ([string]$d.complianceState).Trim() } else { "" }

        # ----- Espace disque -----
        $freeGB = $null; $totGB = $null; $freePct = $null
        if ($d) {
            $fb = $null; $tb = $null
            if ($null -ne $d.freeStorageSpaceInBytes  -and "$($d.freeStorageSpaceInBytes)"  -match '^\d+$') { $fb = [long]$d.freeStorageSpaceInBytes }
            if ($null -ne $d.totalStorageSpaceInBytes -and "$($d.totalStorageSpaceInBytes)" -match '^\d+$') { $tb = [long]$d.totalStorageSpaceInBytes }
            if ($null -ne $fb) { $freeGB = [Math]::Round($fb / 1GB, 1) }
            if ($null -ne $tb -and $tb -gt 0) {
                $totGB = [Math]::Round($tb / 1GB, 1)
                if ($null -ne $fb) { $freePct = [int][Math]::Round(100.0 * $fb / $tb) }
            }
        }

        # ----- Inactivité -----
        $daysSince = $null
        $syncDate  = ConvertTo-ImportDate $sync
        if ($syncDate) { $daysSince = [int][Math]::Floor(($now - $syncDate).TotalDays) }

        # ----- Scores & démarrage -----
        $analytics = $null; $battery = $null
        if ($sc) {
            $analytics = ConvertTo-ImportDouble $sc.endpointAnalyticsScore
            $battery   = ConvertTo-ImportDouble $sc.batteryHealthScore
            if ($null -ne $analytics -and $analytics -lt 0) { $analytics = $null }   # -1 = non mesuré
            if ($null -ne $battery   -and $battery   -lt 0) { $battery   = $null }
        }
        $bootSec = $null; $bootScore = $null; $bsod = $null; $restarts = $null; $diskType = ""
        if ($pf) {
            $ms = ConvertTo-ImportDouble $pf.coreBootTimeInMs
            if ($null -ne $ms -and $ms -gt 0) { $bootSec = [int][Math]::Round($ms / 1000) }
            $bootScore = ConvertTo-ImportDouble $pf.bootScore
            if ($null -ne $pf.blueScreenCount -and "$($pf.blueScreenCount)" -match '^\d+$') { $bsod = [int]$pf.blueScreenCount }
            if ($null -ne $pf.restartCount    -and "$($pf.restartCount)"    -match '^\d+$') { $restarts = [int]$pf.restartCount }
            $diskType = ([string]$pf.diskType).Trim()
        }

        # ----- Règles d'alerte (chacune gouvernée par $script:RemediationChecks) -----
        $issues = @()
        if ((Test-CheckEnabled 'Disk') -and $null -ne $freePct) {
            $fgTxt = "$("$freeGB" -replace '\.', ',') Go libres sur $("$totGB" -replace '\.', ',') Go ($freePct %)"
            if ($freePct -lt $T.DiskFreePctCritical -or ($null -ne $freeGB -and $freeGB -lt $T.DiskFreeGbCritical)) {
                $issues += @{ Label = "Libérer de l'espace disque (moins de $($T.DiskFreePctCritical) % ou $($T.DiskFreeGbCritical) Go libres)"; Sev = 'crit'; Detail = $fgTxt }
            } elseif ($freePct -lt $T.DiskFreePctWarning) {
                $issues += @{ Label = "Espace disque à surveiller (moins de $($T.DiskFreePctWarning) % libres)"; Sev = 'warn'; Detail = $fgTxt }
            }
        }
        if ((Test-CheckEnabled 'Inactivity') -and $null -ne $daysSince) {
            if ($daysSince -ge $T.StaleDaysCritical) {
                $issues += @{ Label = "Appareil inactif depuis plus de $($T.StaleDaysCritical) jours — reprendre contact / retirer d'Intune"; Sev = 'crit'; Detail = "$daysSince jours sans synchronisation" }
            } elseif ($daysSince -ge $T.StaleDaysWarning) {
                $issues += @{ Label = "Appareil sans synchronisation depuis plus de $($T.StaleDaysWarning) jours"; Sev = 'warn'; Detail = "$daysSince jours sans synchronisation" }
            }
        }
        if (Test-CheckEnabled 'Bsod') {
            if ($null -ne $bsod -and $bsod -gt 0) {
                $sev = if ($bsod -ge $T.BsodCritical) { 'crit' } else { 'warn' }
                $det = "$bsod écran(s) bleu(s) sur 14 jours"
                if ($null -ne $restarts) { $det += ", $restarts redémarrage(s)" }
                $issues += @{ Label = "Écrans bleus (BSOD) récents — analyser pilotes et matériel"; Sev = $sev; Detail = $det }
            } elseif ($null -ne $restarts -and $restarts -ge $T.RestartsHigh) {
                $issues += @{ Label = "Redémarrages anormalement fréquents"; Sev = 'warn'; Detail = "$restarts redémarrage(s) sur 14 jours" }
            }
        }
        if ((Test-CheckEnabled 'Boot') -and $null -ne $bootSec -and $bootSec -ge $T.BootSlowSeconds) {
            $det = "Démarrage en $bootSec s"
            if ($null -ne $bootScore) { $det += " (score $([int]$bootScore))" }
            $issues += @{ Label = "Démarrage lent (plus de $($T.BootSlowSeconds) s) — alléger le démarrage"; Sev = 'warn'; Detail = $det }
        }
        if ((Test-CheckEnabled 'Boot') -and $diskType -match '(?i)hdd') {
            $det = "Disque mécanique (HDD)"
            if ($null -ne $bootSec) { $det += " — démarrage $bootSec s" }
            $issues += @{ Label = "Disque mécanique (HDD) — envisager un remplacement SSD"; Sev = 'warn'; Detail = $det }
        }
        if ((Test-CheckEnabled 'Battery') -and $null -ne $battery -and $battery -lt $T.BatteryPoor) {
            $issues += @{ Label = "Batterie dégradée (score sous $($T.BatteryPoor)) — prévoir un remplacement"; Sev = 'warn'; Detail = "Score batterie $([int]$battery)" }
        }
        if ((Test-CheckEnabled 'EaScore') -and $null -ne $analytics -and $analytics -lt $T.ScoreLow) {
            $issues += @{ Label = "Score Endpoint Analytics faible (sous $($T.ScoreLow))"; Sev = 'warn'; Detail = "Score global $([int]$analytics)" }
        }

        # ----- BitLocker -----
        # advancedBitLockerStates est un enum "flags" (valeurs cumulables) : Graph le sérialise
        # en TEXTE, avec plusieurs indicateurs séparés par une virgule quand plusieurs
        # s'appliquent à la fois (ex. "tpmNotReady, loggedOnUserNonAdmin") — jamais un tableau
        # JSON. Tous les indicateurs ne signalent pas un problème : "loggedOnUserNonAdmin"
        # (utilisateur non-administrateur) est une BONNE pratique de sécurité, pas un défaut ;
        # le confondre avec une vraie erreur de chiffrement a produit une écrasante majorité
        # de faux positifs sur ce point. Seuls les indicateurs listés ci-dessous sont retenus.
        $bl = if ($bitlockerByName.ContainsKey($u.Key)) { $bitlockerByName[$u.Key] } else { $null }
        if ((Test-CheckEnabled 'BitLocker') -and $bl) {
            $encState = [string]$bl.encryptionState
            $rawFlags = [string]$bl.advancedBitLockerStates
            $flags    = @($rawFlags -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -ne 'success' })

            # Vraie panne de chiffrement : la clé de récupération ne remonte pas, le TPM ne
            # protège pas le volume, ou le disque n'est tout simplement pas chiffré.
            $criticalFlags = @('osVolumeUnprotected', 'recoveryKeyBackupFailed', 'fixedDriveNotEncrypted', 'tpmNotAvailable', 'tpmNotReady')
            # Écart de politique ou aléa réseau ponctuel : à vérifier, sans urgence immédiate.
            $warningFlags  = @('osVolumeTpmRequired', 'osVolumeTpmOnlyRequired', 'osVolumeTpmPinRequired',
                               'osVolumeTpmStartupKeyRequired', 'osVolumeTpmPinStartupKeyRequired',
                               'osVolumeEncryptionMethodMismatch', 'fixedDriveEncryptionMethodMismatch',
                               'networkError', 'windowsRecoveryEnvironmentNotConfigured')
            # Ni l'un ni l'autre : informationnel (consentement utilisateur) ou carrément une
            # bonne pratique (non-admin) — jamais remonté comme un problème.
            $ignoredFlags  = @('noUserConsent', 'loggedOnUserNonAdmin')

            $relevantFlags = @($flags | Where-Object { $_ -notin $ignoredFlags })

            if ($encState -eq 'notEncrypted') {
                $issues += @{ Label = "BitLocker désactivé — chiffrer le poste"; Sev = 'crit'; Detail = "Volume système non chiffré" }
            } elseif ($relevantFlags.Count -gt 0) {
                $hasCritical = @($relevantFlags | Where-Object { $_ -in $criticalFlags })
                if ($hasCritical.Count -gt 0) {
                    $issues += @{ Label = "BitLocker en échec réel — protection du volume compromise"; Sev = 'crit'; Detail = ($relevantFlags -join ', ') }
                } else {
                    $issues += @{ Label = "BitLocker : écart de stratégie à vérifier (non bloquant)"; Sev = 'warn'; Detail = ($relevantFlags -join ', ') }
                }
            }
        }

        # ----- Defender (antivirus) -----
        $dv = if ($defenderByName.ContainsKey($u.Key)) { $defenderByName[$u.Key] } else { $null }
        if ((Test-CheckEnabled 'Defender') -and $dv) {
            if ($dv.realTimeProtectionEnabled -eq $false) {
                $issues += @{ Label = "Protection en temps réel Defender désactivée"; Sev = 'crit'; Detail = "realTimeProtectionEnabled = false" }
            }
            if ($dv.signatureUpdateOverdue -eq $true) {
                $sigDet = "Signatures en retard"
                if ($dv.lastReportedDateTime) { $sigDet += " (dernier rapport : $(Format-SyncDate $dv.lastReportedDateTime))" }
                $issues += @{ Label = "Signatures antivirus en retard (plus de $($T.SignatureStaleDays) jours)"; Sev = 'warn'; Detail = $sigDet }
            }
        }

        # ----- Fiabilité applicative -----
        $crashInfo = if ($appCrashByName.ContainsKey($u.Key)) { $appCrashByName[$u.Key] } else { $null }
        if ((Test-CheckEnabled 'AppReliability') -and $crashInfo -and $crashInfo.Count -ge $T.AppCrashWarning) {
            $issues += @{ Label = "Application qui plante fréquemment — envisager réinstallation/mise à jour"; Sev = 'warn'; Detail = "$($crashInfo.AppName) : $($crashInfo.Count) plantage(s)" }
        }

        # ----- Uptime (estimation via dernier démarrage connu) -----
        $lastBoot = if ($startupByName.ContainsKey($u.Key)) { $startupByName[$u.Key] } else { $null }
        $uptimeDays = if ($lastBoot) { [int][Math]::Floor(($now - $lastBoot).TotalDays) } else { $null }
        if ((Test-CheckEnabled 'Uptime') -and $null -ne $uptimeDays -and $uptimeDays -ge $T.UptimeWarningDays) {
            $issues += @{ Label = "Poste non redémarré depuis longtemps (estimation) — planifier un redémarrage"; Sev = 'warn'; Detail = "~$uptimeDays jour(s) depuis le dernier démarrage détecté" }
        }

        # ----- Conformité aux stratégies de conformité Intune -----
        # L'onglet 1 dit COMBIEN de postes sont non conformes et POURQUOI ; ici on ramène
        # l'information au niveau du poste, à côté de ses autres défauts, pour qu'un
        # technicien n'ait pas à recouper deux onglets avant d'agir. Le motif précis est
        # repris de l'analyse de conformité quand elle a été faite dans la même génération.
        if ((Test-CheckEnabled 'Compliance') -and $compState) {
            $reasonDetail = $null
            if ($ComplianceReasons -and $ComplianceReasons.ContainsKey($u.Key)) { $reasonDetail = [string]$ComplianceReasons[$u.Key] }
            switch -Regex ($compState) {
                '(?i)^noncompliant$' {
                    $det = if ($reasonDetail) { $reasonDetail } else { "État Intune : non conforme (motif non collecté dans cette génération)" }
                    $issues += @{ Label = "Non conforme aux stratégies de conformité Intune"; Sev = 'crit'; Detail = $det }
                }
                '(?i)^inGracePeriod$' {
                    $det = if ($reasonDetail) { "$reasonDetail (période de grâce)" } else { "État Intune : période de grâce avant bascule en non conforme" }
                    $issues += @{ Label = "En période de grâce — à traiter avant bascule en non conforme"; Sev = 'warn'; Detail = $det }
                }
                '(?i)^(error|conflict)$' {
                    $issues += @{ Label = "Évaluation de conformité en erreur ou en conflit — stratégie non concluante"; Sev = 'warn'; Detail = "État Intune : $compState" }
                }
            }
        }

        # ----- Profils de configuration en erreur / conflit -----
        $cfgErr = if ($configErrByName.ContainsKey($u.Key)) { $configErrByName[$u.Key] } else { $null }
        if ((Test-CheckEnabled 'ConfigProfile') -and $cfgErr -and $cfgErr.Count -gt 0) {
            $names = @($cfgErr.Profiles) | Select-Object -First 3
            $det   = "$($cfgErr.Count) profil(s) en échec"
            if ($names.Count -gt 0) { $det += " : " + ($names -join ' ; ') }
            if ($cfgErr.Count -gt $names.Count) { $det += " ..." }
            if ($cfgErr.Count -ge $T.ConfigErrorsCritical) {
                $issues += @{ Label = "Profils de configuration en échec ($($T.ConfigErrorsCritical) ou plus) — paramétrage non appliqué"; Sev = 'crit'; Detail = $det }
            } elseif ($cfgErr.HasConflict) {
                $issues += @{ Label = "Conflit entre profils de configuration — deux stratégies se contredisent"; Sev = 'warn'; Detail = $det }
            } else {
                $issues += @{ Label = "Profil de configuration en erreur — paramétrage non appliqué"; Sev = 'warn'; Detail = $det }
            }
        }

        # ----- Batterie : capacité, âge, autonomie -----
        $bh = if ($batteryByName.ContainsKey($u.Key)) { $batteryByName[$u.Key] } else { $null }
        $battCapacity = $null; $battAgeDays = $null; $battRuntimeMin = $null
        if ($bh) {
            $battCapacity   = ConvertTo-ImportDouble (Get-PropCI $bh @('maxCapacityPercentage','maxCapacityPercent'))
            $battAgeDays    = ConvertTo-ImportDouble (Get-PropCI $bh @('batteryAgeInDays'))
            $battRuntimeMin = ConvertTo-ImportDouble (Get-PropCI $bh @('estimatedRuntimeInMinutes'))
            # Endpoint Analytics code "non mesuré" par une valeur négative, pas par un vide.
            if ($null -ne $battCapacity   -and $battCapacity   -le 0) { $battCapacity   = $null }
            if ($null -ne $battAgeDays    -and $battAgeDays    -lt 0) { $battAgeDays    = $null }
            if ($null -ne $battRuntimeMin -and $battRuntimeMin -le 0) { $battRuntimeMin = $null }
        }
        if ((Test-CheckEnabled 'BatteryDetail') -and $null -ne $battCapacity -and $battCapacity -lt $T.BatteryCapacityPoor) {
            $det = "Capacité maximale restante : $([int]$battCapacity) %"
            if ($null -ne $battAgeDays)    { $det += " — batterie âgée de ~$([int][Math]::Round($battAgeDays / 365.0, 0)) an(s)" }
            if ($null -ne $battRuntimeMin) { $det += " — autonomie estimée $([int]$battRuntimeMin) min" }
            if ($battCapacity -lt $T.BatteryCapacityCrit) {
                $issues += @{ Label = "Batterie hors d'usage (capacité sous $($T.BatteryCapacityCrit) %) — remplacement à planifier"; Sev = 'crit'; Detail = $det }
            } else {
                $issues += @{ Label = "Batterie usée (capacité sous $($T.BatteryCapacityPoor) %) — remplacement à prévoir"; Sev = 'warn'; Detail = $det }
            }
        } elseif ((Test-CheckEnabled 'BatteryDetail') -and $null -ne $battRuntimeMin -and $battRuntimeMin -lt $T.BatteryRuntimeLowMin -and $null -ne $battCapacity) {
            $issues += @{ Label = "Autonomie insuffisante (moins de $($T.BatteryRuntimeLowMin) min) — poste devenu sédentaire"; Sev = 'warn'; Detail = "Autonomie estimée $([int]$battRuntimeMin) min pour $([int]$battCapacity) % de capacité" }
        }

        # ----- Mises à jour qualité Windows en échec -----
        $upd = if ($updateFailByName.ContainsKey($u.Key)) { $updateFailByName[$u.Key] } else { $null }
        if ((Test-CheckEnabled 'WindowsUpdate') -and $upd) {
            $det = "État : $($upd.State)"
            if ($upd.Message) { $det += " — $($upd.Message)" }
            $issues += @{ Label = "Mise à jour qualité Windows en échec — poste privé de correctifs de sécurité"; Sev = 'crit'; Detail = $det }
        }

        # ----- Sévérité globale du poste -----
        $sevMax = 0
        foreach ($i in $issues) { if ($sevRank[$i.Sev] -gt $sevMax) { $sevMax = $sevRank[$i.Sev] } }
        $severity = @('ok', 'warn', 'crit')[$sevMax]

        # ----- Anonymisation cohérente entre onglets -----
        if ($AnonymizeData) {
            $anon = Get-AnonymizedIdentity -RealName $name -RealUpn $upn
            $name = $anon.Name
            $upn  = $anon.Upn
        }

        foreach ($i in $issues) {
            if (-not $actions.Contains($i.Label)) {
                $actions[$i.Label] = @{ Severity = $i.Sev; Count = 0; Devices = @() }
            }
            $a = $actions[$i.Label]
            $a.Count++
            if ($sevRank[$i.Sev] -gt $sevRank[$a.Severity]) { $a.Severity = $i.Sev }
            $a.Devices += [PSCustomObject]@{
                DeviceName        = $name
                UserPrincipalName = $upn
                OperatingSystem   = $os
                Detail            = $i.Detail
                LastSyncDateTime  = $sync
            }
        }

        $rows.Add([PSCustomObject]@{
            DeviceName        = $name
            UserPrincipalName = $upn
            OperatingSystem   = $os
            OSVersion         = $osv
            FreeGB            = $freeGB
            TotalGB           = $totGB
            FreePct           = $freePct
            AnalyticsScore    = $analytics
            BootSeconds       = $bootSec
            DiskType          = $diskType
            BlueScreens       = $bsod
            Restarts          = $restarts
            BatteryScore      = $battery
            LastSyncDateTime  = $sync
            DaysSinceSync     = $daysSince
            BitLockerState    = if ($bl) { [string]$bl.encryptionState } else { $null }
            DefenderRealTime  = if ($dv) { $dv.realTimeProtectionEnabled } else { $null }
            DefenderSigStale  = if ($dv) { $dv.signatureUpdateOverdue }   else { $null }
            CrashApp          = if ($crashInfo) { $crashInfo.AppName }  else { $null }
            CrashCount        = if ($crashInfo) { $crashInfo.Count }    else { $null }
            UptimeDays        = $uptimeDays
            ComplianceState   = $compState
            ConfigErrorCount  = if ($cfgErr) { $cfgErr.Count } else { $null }
            ConfigErrorNames  = if ($cfgErr) { (@($cfgErr.Profiles) -join ' ; ') } else { $null }
            BatteryCapacity   = $battCapacity
            BatteryAgeDays    = $battAgeDays
            BatteryRuntimeMin = $battRuntimeMin
            UpdateFailure     = if ($upd) { $upd.State } else { $null }
            DeviceId          = if ($d) { [string]$d.id } else { "" }
            Severity          = $severity
        })
    }

    $sorted = @($rows | Sort-Object -Property @{ e = { $sevRank[$_.Severity] }; Descending = $true }, @{ e = { $_.DeviceName } })
    $crit = @($sorted | Where-Object { $_.Severity -eq 'crit' }).Count
    $warn = @($sorted | Where-Object { $_.Severity -eq 'warn' }).Count
    $okN  = $sorted.Count - $crit - $warn

    $avgScore = $null
    $withScore = @($sorted | Where-Object { $null -ne $_.AnalyticsScore })
    if ($withScore.Count -gt 0) {
        $avgScore = [int][Math]::Round((($withScore | Measure-Object -Property AnalyticsScore -Average).Average))
    }

    return [PSCustomObject]@{
        Rows      = $sorted
        Actions   = $actions
        CritCount = $crit
        WarnCount = $warn
        OkCount   = $okN
        AvgScore  = $avgScore
    }
}

function Build-RemediationDeviceTableHtml {
    <# Tableau des postes d'une action : colonne "Détail" propre à la remédiation. #>
    param([array]$Devices)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='tbl-scroll'><table class='tbl'><thead><tr><th scope='col'>Poste</th><th scope='col'>Utilisateur</th><th scope='col'>OS</th><th scope='col'>D&eacute;tail</th><th scope='col'>Derni&egrave;re synchro</th></tr></thead><tbody>")
    foreach ($d in $Devices) {
        $dn = ConvertTo-HtmlSafe ([string]$d.DeviceName)
        $up = ConvertTo-HtmlSafe ([string]$d.UserPrincipalName)
        $os = ConvertTo-HtmlSafe ([string]$d.OperatingSystem)
        $dt = ConvertTo-HtmlSafe ([string]$d.Detail)
        $ls = ConvertTo-HtmlSafe (Format-SyncDate $d.LastSyncDateTime)
        # Le détail entre dans le champ de recherche : chercher « BitLocker » ou un nom de
        # profil doit ramener les postes concernés, pas seulement les intitulés d'action.
        $search = ConvertTo-HtmlSafe (("$($d.DeviceName) $($d.UserPrincipalName) $($d.Detail)").ToLower())
        [void]$sb.Append("<tr data-search='$search'><td class='cell-strong'>$dn</td><td>$up</td><td>$os</td><td>$dt</td><td class='muted'>$ls</td></tr>")
    }
    [void]$sb.Append("</tbody></table></div>")
    return $sb.ToString()
}

function Build-RemediationAccordionHtml {
    <# Un accordéon par action recommandée, trié par sévérité puis nombre de postes. #>
    param($Actions)
    if (-not $Actions -or $Actions.Count -eq 0) {
        return "<div class='empty'>Aucune action de rem&eacute;diation n&eacute;cessaire &mdash; tous les indicateurs collect&eacute;s sont dans les seuils.</div>"
    }
    $rank = @{ crit = 2; warn = 1; ok = 0 }
    $sorted = $Actions.GetEnumerator() | Sort-Object -Property @{ e = { $rank[$_.Value.Severity] }; Descending = $true }, @{ e = { $_.Value.Count }; Descending = $true }
    $sb = New-Object System.Text.StringBuilder
    foreach ($entry in $sorted) {
        $label   = ConvertTo-HtmlSafe $entry.Key
        $info    = $entry.Value
        $variant = if ($info.Severity -eq 'crit') { 'red' } else { 'amber' }
        $slug    = ConvertTo-Slug $entry.Key
        $search  = ConvertTo-HtmlSafe (([string]$entry.Key).ToLower())
        $expBtn  = Build-AccordionExportButton -Label $entry.Key
        [void]$sb.Append("<details class='acc' data-search='$search' data-export-name='action-$slug' data-action-label=`"$label`"><summary><span class='dot dot-$variant'></span><span class='acc-title'>$label</span><span class='badge badge-$variant'>$($info.Count) poste(s)</span>$expBtn<span class='chev'>&rsaquo;</span></summary><div class='acc-body'>")
        [void]$sb.Append((Build-RemediationDeviceTableHtml -Devices $info.Devices))
        [void]$sb.Append("</div></details>")
    }
    return $sb.ToString()
}

function Build-RemediationRowsHtml {
    <# Lignes du tableau global de santé (recherche, filtre par état, tri par colonne). #>
    param([array]$Rows)
    $sb = New-Object System.Text.StringBuilder
    foreach ($r in $Rows) {
        $nm  = ConvertTo-HtmlSafe ([string]$r.DeviceName)
        $up  = ConvertTo-HtmlSafe ([string]$r.UserPrincipalName)
        # Le champ de recherche d'une ligne couvre l'identité DU POSTE et ses ALERTES : on
        # doit pouvoir taper « batterie », « profil » ou « non conforme » et voir la liste
        # se réduire aux postes concernés, sans passer par les accordéons.
        $searchBits = New-Object System.Collections.Generic.List[string]
        [void]$searchBits.Add([string]$r.DeviceName)
        [void]$searchBits.Add([string]$r.UserPrincipalName)
        [void]$searchBits.Add([string]$r.OperatingSystem)
        [void]$searchBits.Add([string]$r.OSVersion)
        if ($r.ComplianceState)  { [void]$searchBits.Add([string]$r.ComplianceState) }
        if ($r.ConfigErrorNames) { [void]$searchBits.Add("profil configuration " + [string]$r.ConfigErrorNames) }
        if ($null -ne $r.BatteryCapacity) { [void]$searchBits.Add("batterie") }
        if ($r.UpdateFailure)    { [void]$searchBits.Add("mise a jour " + [string]$r.UpdateFailure) }
        if ($r.CrashApp)         { [void]$searchBits.Add([string]$r.CrashApp) }
        if ($r.Severity -eq 'crit') { [void]$searchBits.Add("critique") } elseif ($r.Severity -eq 'warn') { [void]$searchBits.Add("a surveiller") }
        $search = ConvertTo-HtmlSafe ((($searchBits -join ' ')).ToLower())

        if ($null -ne $r.FreePct) {
            $fg = "$($r.FreeGB)" -replace '\.', ','
            $tg = "$($r.TotalGB)" -replace '\.', ','
            $disk = "$fg / $tg Go ($($r.FreePct)&nbsp;%)"
        } else { $disk = "&mdash;" }

        $score = if ($null -ne $r.AnalyticsScore) { "$([int]$r.AnalyticsScore)" }  else { "&mdash;" }
        $boot  = if ($null -ne $r.BootSeconds)    { "$($r.BootSeconds) s" }        else { "&mdash;" }
        $dtype = if ($r.DiskType)                 { ConvertTo-HtmlSafe ($r.DiskType.ToUpper()) } else { "&mdash;" }
        $bsod  = if ($null -ne $r.BlueScreens)    { "$($r.BlueScreens)" }          else { "&mdash;" }
        # Batterie : le score composite seul ne se discute pas avec un utilisateur ; la
        # capacité restante, si, et c'est elle qui justifie un bon de commande.
        if ($null -ne $r.BatteryCapacity) {
            $battClass = if ($r.BatteryCapacity -lt 50) { 'tag-red' } elseif ($r.BatteryCapacity -lt 70) { 'tag-amber' } else { 'tag-ok' }
            $batt = "<span class='tag $battClass'>$([int]$r.BatteryCapacity)&nbsp;%</span>"
            if ($null -ne $r.BatteryScore) { $batt += " <span class='muted'>score $([int]$r.BatteryScore)</span>" }
        } elseif ($null -ne $r.BatteryScore) {
            $batt = "$([int]$r.BatteryScore)"
        } else {
            $batt = "&mdash;"
        }
        $sync  = ConvertTo-HtmlSafe (Format-SyncDate $r.LastSyncDateTime)
        if ($null -ne $r.DaysSinceSync) { $sync += " <span class='muted'>($($r.DaysSinceSync)&nbsp;j)</span>" }

        # ----- Sécurité (BitLocker + Defender combinés dans une seule cellule) -----
        $secBits = @()
        if ($r.BitLockerState -eq 'notEncrypted') { $secBits += "<span class='tag tag-red'>BitLocker&nbsp;off</span>" }
        elseif ($r.BitLockerState) { $secBits += "<span class='tag tag-ok'>BitLocker&nbsp;OK</span>" }
        if ($r.DefenderRealTime -eq $false) { $secBits += "<span class='tag tag-red'>Defender&nbsp;off</span>" }
        if ($r.DefenderSigStale -eq $true)  { $secBits += "<span class='tag tag-amber'>Signatures&nbsp;p&eacute;rim&eacute;es</span>" }
        $secu = if ($secBits.Count -gt 0) { $secBits -join ' ' } else { "&mdash;" }

        # ----- Conformité Intune, profils de configuration et mises à jour -----
        # Regroupés en une cellule : ce sont les trois façons dont un poste peut être
        # « géré mais pas conforme à ce qu'on attend de lui ».
        $confBits = @()
        switch -Regex ([string]$r.ComplianceState) {
            '(?i)^compliant$'        { $confBits += "<span class='tag tag-ok'>Conforme</span>" }
            '(?i)^noncompliant$'     { $confBits += "<span class='tag tag-red'>Non conforme</span>" }
            '(?i)^inGracePeriod$'    { $confBits += "<span class='tag tag-amber'>P&eacute;riode de gr&acirc;ce</span>" }
            '(?i)^(error|conflict)$' { $confBits += "<span class='tag tag-amber'>&Eacute;valuation en erreur</span>" }
        }
        if ($null -ne $r.ConfigErrorCount -and $r.ConfigErrorCount -gt 0) {
            $cfgTitle = ConvertTo-HtmlSafe ([string]$r.ConfigErrorNames)
            $cfgClass = if ($r.ConfigErrorCount -ge 3) { 'tag-red' } else { 'tag-amber' }
            $confBits += "<span class='tag $cfgClass' title=`"$cfgTitle`">$($r.ConfigErrorCount)&nbsp;profil(s)&nbsp;KO</span>"
        }
        if ($r.UpdateFailure) { $confBits += "<span class='tag tag-red'>MAJ&nbsp;en&nbsp;&eacute;chec</span>" }
        $conf = if ($confBits.Count -gt 0) { $confBits -join ' ' } else { "&mdash;" }

        # ----- Fiabilité applicative -----
        $reli = if ($r.CrashApp) { "$(ConvertTo-HtmlSafe $r.CrashApp) <span class='muted'>($($r.CrashCount))</span>" } else { "&mdash;" }

        # ----- Uptime (estimation) -----
        $uptime = if ($null -ne $r.UptimeDays) { "~$($r.UptimeDays)&nbsp;j" } else { "&mdash;" }

        $badge = switch ($r.Severity) {
            'crit' { "<span class='badge badge-red'>Critique</span>" }
            'warn' { "<span class='badge badge-amber'>&Agrave; surveiller</span>" }
            default { "<span class='badge badge-ok'>OK</span>" }
        }

        # ----- Actions (copie de commande PowerShell prête à coller — voir note en bas de page) -----
        $actions = "&mdash;"
        if ($r.DeviceId) {
            $devIdJs = ($r.DeviceId -replace "'", "\\'")
            $nameJs  = ($r.DeviceName -replace "'", "\\'")
            $actions = "<div class='act-row'>" +
                "<button class='act-btn' onclick=`"copyDeviceAction('$devIdJs','Sync','$nameJs',this)`" title='Copier la commande de synchronisation'>Sync</button>" +
                "<button class='act-btn' onclick=`"copyDeviceAction('$devIdJs','Reboot','$nameJs',this)`" title='Copier la commande de red&eacute;marrage'>Reboot</button>" +
                "<button class='act-btn' onclick=`"copyDeviceAction('$devIdJs','Remediate','$nameJs',this)`" title='Copier la commande de rem&eacute;diation cibl&eacute;e'>Rem&eacute;dier</button>" +
                "</div>"
        }

        [void]$sb.Append("<tr data-status='$($r.Severity)' data-search='$search'><td class='cell-strong'>$nm</td><td>$up</td><td class='mono'>$disk</td><td class='mono'>$score</td><td class='mono'>$boot</td><td>$dtype</td><td class='mono'>$bsod</td><td>$batt</td><td>$secu</td><td>$conf</td><td>$reli</td><td class='mono'>$uptime</td><td class='muted'>$sync</td><td>$badge</td><td>$actions</td></tr>")
    }
    return $sb.ToString()
}

# ========================================
# INDICATEURS DE PARC (KPI DE L'EN-TÊTE DU DASHBOARD)
# ========================================
#
# Ces indicateurs répondent à la question que se pose un responsable de parc AVANT
# d'ouvrir le détail : « est-ce que ça va, et sinon, de combien de postes parle-t-on ? »
# Ils sont calculés sur le PÉRIMÈTRE RÉELLEMENT ANALYSÉ (après exclusion des machines
# virtuelles le cas échéant), afin que le pourcentage affiché soit celui du parc dont
# l'administrateur a la charge, et pas d'un ensemble théorique.

# Correspondance numéro de build -> version commerciale. Un administrateur raisonne en
# « 23H2 », pas en « 22631 » ; l'inverse est vrai pour Graph. Cette table fait le pont.
$script:WindowsReleaseMap = @{
    26200 = "Windows 11 25H2"
    26100 = "Windows 11 24H2"
    22631 = "Windows 11 23H2"
    22621 = "Windows 11 22H2"
    22000 = "Windows 11 21H2"
    19045 = "Windows 10 22H2"
    19044 = "Windows 10 21H2"
    19043 = "Windows 10 21H1"
    19042 = "Windows 10 20H2"
    19041 = "Windows 10 2004"
    18363 = "Windows 10 1909"
    18362 = "Windows 10 1903"
    17763 = "Windows 10 1809"
}

function Get-WindowsBuildNumber {
    <# "10.0.22631.4317" -> 22631. $null si la chaîne n'a pas la forme attendue. #>
    param([string]$OsVersion)
    if ([string]::IsNullOrWhiteSpace($OsVersion)) { return $null }
    $parts = ([string]$OsVersion).Trim().Split('.')
    if ($parts.Count -lt 3) { return $null }
    $build = 0
    if ([int]::TryParse($parts[2], [ref]$build) -and $build -gt 0) { return $build }
    return $null
}

function Get-OsFamilyLabel {
    <#
        Famille d'OS lisible. Le point sensible est la distinction Windows 10 / Windows 11 :
        Graph renvoie "Windows" pour les deux et une osVersion en 10.0.x — c'est le NUMÉRO
        DE BUILD, et lui seul, qui les sépare (seuil 22000). Se fier au libellé
        operatingSystem, comme on le voit souvent, range tout le parc en "Windows 10".
    #>
    param([string]$OperatingSystem, [string]$OsVersion)
    $os = ([string]$OperatingSystem).Trim()
    if ([string]::IsNullOrWhiteSpace($os)) { return "Non renseigné" }
    if ($os -match '(?i)^windows') {
        if ($os -match '(?i)phone|mobile') { return "Windows Phone" }
        $build = Get-WindowsBuildNumber -OsVersion $OsVersion
        if ($null -eq $build)   { return "Windows (version inconnue)" }
        if ($build -ge 22000)   { return "Windows 11" }
        if ($build -ge 10240)   { return "Windows 10" }
        return "Windows (antérieur à 10)"
    }
    if ($os -match '(?i)^mac')            { return "macOS" }
    if ($os -match '(?i)^ipad')           { return "iPadOS" }
    if ($os -match '(?i)^ios')            { return "iOS" }
    if ($os -match '(?i)^android')        { return "Android" }
    if ($os -match '(?i)^linux|^ubuntu')  { return "Linux" }
    if ($os -match '(?i)chrome')          { return "ChromeOS" }
    return $os
}

function Get-WindowsReleaseLabel {
    <# Version commerciale ("Windows 11 23H2") ou, à défaut, le build brut. #>
    param([string]$OperatingSystem, [string]$OsVersion)
    if (([string]$OperatingSystem) -notmatch '(?i)^windows') { return $null }
    $build = Get-WindowsBuildNumber -OsVersion $OsVersion
    if ($null -eq $build) { return "Windows (version inconnue)" }
    if ($script:WindowsReleaseMap.ContainsKey($build)) { return $script:WindowsReleaseMap[$build] }
    if ($build -ge 22000) { return "Windows 11 (build $build)" }
    if ($build -ge 10240) { return "Windows 10 (build $build)" }
    return "Windows (build $build)"
}

function Get-FleetKpiSummary {
    <#
        Agrège en un seul passage les indicateurs de tête du rapport. Un seul parcours du
        parc : sur 50 000 appareils, enchaîner huit Where-Object successifs coûte huit
        parcours complets pour un résultat identique.

        $StaleDays : seuil d'inactivité de l'indicateur principal (30 jours par défaut,
        aligné sur RemediationThresholds.StaleDaysWarning).
    #>
    param(
        [array]$Devices,
        [int]$StaleDays = 30,
        [int]$StaleCriticalDays = 90,
        [array]$BitLockerStates,
        $RemediationData
    )
    $now = Get-Date

    $total = 0; $compliant = 0; $nonCompliant = 0; $grace = 0; $errorState = 0; $otherState = 0
    $stale = 0; $staleCritical = 0; $neverSynced = 0
    $osFamilies      = @{}
    $windowsReleases = @{}
    $win10 = 0; $win11 = 0; $windowsTotal = 0

    foreach ($d in @($Devices)) {
        if ($null -eq $d) { continue }
        $total++

        switch -Regex (([string]$d.complianceState).Trim()) {
            '(?i)^compliant$'      { $compliant++ }
            '(?i)^noncompliant$'   { $nonCompliant++ }
            '(?i)^inGracePeriod$'  { $grace++ }
            '(?i)^(error|conflict)$' { $errorState++ }
            default                { $otherState++ }
        }

        $syncDate = ConvertTo-ImportDate ([string]$d.lastSyncDateTime)
        if ($null -eq $syncDate) {
            $neverSynced++
        } else {
            $days = ($now - $syncDate).TotalDays
            if ($days -ge $StaleCriticalDays) { $staleCritical++ }
            if ($days -ge $StaleDays)         { $stale++ }
        }

        $family = Get-OsFamilyLabel -OperatingSystem ([string]$d.operatingSystem) -OsVersion ([string]$d.osVersion)
        if (-not $osFamilies.ContainsKey($family)) { $osFamilies[$family] = 0 }
        $osFamilies[$family]++

        if (([string]$d.operatingSystem) -match '(?i)^windows' -and ([string]$d.operatingSystem) -notmatch '(?i)phone|mobile') {
            $windowsTotal++
            if     ($family -eq "Windows 11") { $win11++ }
            elseif ($family -eq "Windows 10") { $win10++ }
            $release = Get-WindowsReleaseLabel -OperatingSystem ([string]$d.operatingSystem) -OsVersion ([string]$d.osVersion)
            if ($release) {
                if (-not $windowsReleases.ContainsKey($release)) { $windowsReleases[$release] = 0 }
                $windowsReleases[$release]++
            }
        }
    }

    # Chiffrement : le parc chiffré n'est pas « le nombre d'états BitLocker remontés » mais
    # la part des postes analysés dont le volume système est effectivement chiffré.
    $encrypted = $null
    if ($BitLockerStates -and @($BitLockerStates).Count -gt 0) {
        $encryptedNames = New-Object System.Collections.Generic.HashSet[string]
        foreach ($b in @($BitLockerStates)) {
            if (([string]$b.encryptionState) -match '(?i)^encrypted$') {
                $k = Get-NormalizedHeader ([string]$b.deviceName)
                if ($k) { [void]$encryptedNames.Add($k) }
            }
        }
        $encrypted = $encryptedNames.Count
    }

    $complianceRate = if ($total -gt 0) { [int][Math]::Round(100.0 * $compliant / $total) } else { 0 }
    $win11Share     = if ($windowsTotal -gt 0) { [int][Math]::Round(100.0 * $win11 / $windowsTotal) } else { 0 }
    $encryptionRate = if ($null -ne $encrypted -and $total -gt 0) { [int][Math]::Round(100.0 * [Math]::Min($encrypted, $total) / $total) } else { $null }

    # Postes dont l'OS n'est plus soutenu : Windows 10 après le 14/10/2025 (hors LTSC/ESU).
    $endOfSupport = 0
    if ((Get-Date) -ge $script:Windows10EndOfSupport) { $endOfSupport = $win10 }

    # Tri décroissant des versions Windows : la plus répandue en tête, c'est celle sur
    # laquelle porte la décision de déploiement.
    $releasesSorted = [ordered]@{}
    foreach ($entry in ($windowsReleases.GetEnumerator() | Sort-Object -Property Value -Descending)) {
        $releasesSorted[$entry.Key] = $entry.Value
    }
    $familiesSorted = [ordered]@{}
    foreach ($entry in ($osFamilies.GetEnumerator() | Sort-Object -Property Value -Descending)) {
        $familiesSorted[$entry.Key] = $entry.Value
    }

    return [PSCustomObject]@{
        Total             = $total
        Compliant         = $compliant
        NonCompliant      = $nonCompliant
        GracePeriod       = $grace
        ErrorState        = $errorState
        OtherState        = $otherState
        ComplianceRate    = $complianceRate
        StaleDays         = $StaleDays
        StaleCount        = $stale
        StaleCriticalDays = $StaleCriticalDays
        StaleCritical     = $staleCritical
        NeverSynced       = $neverSynced
        OsFamilies        = $familiesSorted
        WindowsReleases   = $releasesSorted
        Windows10         = $win10
        Windows11         = $win11
        WindowsTotal      = $windowsTotal
        Windows11Share    = $win11Share
        EndOfSupport      = $endOfSupport
        EncryptedCount    = $encrypted
        EncryptionRate    = $encryptionRate
        CritActions       = if ($RemediationData) { [int]$RemediationData.CritCount } else { $null }
        WarnActions       = if ($RemediationData) { [int]$RemediationData.WarnCount } else { $null }
        AvgAnalyticsScore = if ($RemediationData) { $RemediationData.AvgScore } else { $null }
    }
}

# ========================================
# CONSTRUCTION DU RAPPORT HTML
# Design sur mesure, fichier 100% autonome (aucune dépendance, consultable hors ligne)
# ========================================

function Format-SyncDate {
    <#
        Affiche une date de synchronisation. Réutilise EXACTEMENT le même analyseur
        que le calcul du nombre de jours écoulés (ConvertTo-ImportDate), plutôt qu'un
        cast [datetime] "nu" : ce dernier interprète les dates ambiguës (jour <= 12,
        ex. "04/08/2026") selon la culture régionale courante du poste qui exécute le
        script, qui peut différer de l'analyse dd/MM/yyyy explicite utilisée pour les
        jours. Résultat sans ce correctif : la date AFFICHÉE peut sembler très récente
        (jour et mois inversés) alors que le nombre de jours, lui, est calculé
        correctement — un appareil réellement inactif depuis des mois s'affiche alors
        avec une date du jour même, ce qui n'a aucun sens pour qui lit le rapport.
    #>
    param($Value)
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return "—" }
    $dt = ConvertTo-ImportDate $Value
    if ($null -eq $dt) { return [string]$Value }
    try { return $dt.ToLocalTime().ToString("dd/MM/yyyy HH:mm") } catch { return [string]$Value }
}

function Get-PropCI {
    # Récupère une propriété quelle que soit sa casse (DeviceName vs deviceName),
    # pour réutiliser le même moteur de tableau avec les objets normalisés ET les objets Graph bruts.
    # $Object peut être $null (ex. @($null) produit un tableau à un élément en PowerShell,
    # jamais un tableau vide) : on renvoie alors simplement une chaîne vide, sans erreur.
    param($Object, [string[]]$Names)
    if ($null -eq $Object) { return "" }
    foreach ($n in $Names) {
        $p = $Object.PSObject.Properties[$n]
        if ($p -and $null -ne $p.Value -and "$($p.Value)" -ne "") { return [string]$p.Value }
    }
    return ""
}

function ConvertTo-Slug {
    <# Libellé -> fragment de nom de fichier sûr, sans accent ni caractère interdit.
       Sert à nommer les exports CSV d'après l'action ou le motif exporté. #>
    param([string]$Text, [int]$MaxLength = 60)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "extrait" }
    $normalized = [string]$Text
    try {
        $decomposed = $normalized.Normalize([System.Text.NormalizationForm]::FormD)
        $builder = New-Object System.Text.StringBuilder
        foreach ($ch in $decomposed.ToCharArray()) {
            if ([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark) {
                [void]$builder.Append($ch)
            }
        }
        $normalized = $builder.ToString()
    } catch { }
    $normalized = ($normalized -replace '[^A-Za-z0-9]+', '-').Trim('-')
    if ($normalized.Length -gt $MaxLength) { $normalized = $normalized.Substring(0, $MaxLength).Trim('-') }
    if ([string]::IsNullOrWhiteSpace($normalized)) { return "extrait" }
    return $normalized
}

function Build-AccordionExportButton {
    <# Bouton d'export CSV placé DANS le <summary> d'un accordéon. Le clic ne doit pas
       déplier/replier l'accordéon : le gestionnaire JS neutralise l'événement natif. #>
    param([string]$Label)
    $safe = ConvertTo-HtmlSafe $Label
    return "<button type='button' class='acc-exp' onclick='exportAccordion(this, event)' title='Exporter cette liste au format CSV' aria-label='Exporter la liste : $safe'>&#11015; CSV</button>"
}

function Build-DeviceTableHtml {
    param([array]$Devices, [bool]$IncludeSync = $true)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='tbl-scroll'><table class='tbl'><thead><tr><th scope='col'>Poste</th><th scope='col'>Utilisateur</th><th scope='col'>OS</th><th scope='col'>Version OS</th>")
    if ($IncludeSync) { [void]$sb.Append("<th scope='col'>Derni&egrave;re synchro</th>") }
    [void]$sb.Append("</tr></thead><tbody>")
    foreach ($d in $Devices) {
        $rawName = [string](Get-PropCI $d @('DeviceName','deviceName'))
        $rawUpn  = [string](Get-PropCI $d @('UserPrincipalName','userPrincipalName'))
        $dn = ConvertTo-HtmlSafe $rawName
        $up = ConvertTo-HtmlSafe $rawUpn
        $os = ConvertTo-HtmlSafe (Get-PropCI $d @('OperatingSystem','operatingSystem'))
        $ov = ConvertTo-HtmlSafe (Get-PropCI $d @('OSVersion','osVersion'))
        # data-search : socle du filtre instantané côté navigateur, calculé une fois ici
        # plutôt que reconstruit à chaque frappe depuis le texte du DOM.
        $search = ConvertTo-HtmlSafe (("$rawName $rawUpn").ToLower())
        [void]$sb.Append("<tr data-search='$search'><td class='cell-strong'>$dn</td><td>$up</td><td>$os</td><td class='mono'>$ov</td>")
        if ($IncludeSync) {
            $ls = ConvertTo-HtmlSafe (Format-SyncDate (Get-PropCI $d @('LastSyncDateTime','lastSyncDateTime')))
            [void]$sb.Append("<td class='muted'>$ls</td>")
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</tbody></table></div>")
    return $sb.ToString()
}

function Build-ReasonAccordionHtml {
    param($Breakdown, [ValidateSet('red','amber')][string]$Variant)
    if (-not $Breakdown -or $Breakdown.Count -eq 0) {
        return "<div class='empty'>Aucun appareil dans cette cat&eacute;gorie.</div>"
    }
    $sorted = $Breakdown.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending
    $sb = New-Object System.Text.StringBuilder
    foreach ($entry in $sorted) {
        $label = ConvertTo-HtmlSafe $entry.Key
        $info  = $entry.Value
        $slug   = ConvertTo-Slug $entry.Key
        $search = ConvertTo-HtmlSafe (([string]$entry.Key).ToLower())
        $expBtn = Build-AccordionExportButton -Label $entry.Key
        [void]$sb.Append("<details class='acc' data-search='$search' data-export-name='motif-$slug'><summary><span class='dot dot-$Variant'></span><span class='acc-title'>$label</span><span class='badge badge-$Variant'>$($info.Count) poste(s)</span>$expBtn<span class='chev'>&rsaquo;</span></summary><div class='acc-body'>")
        [void]$sb.Append((Build-DeviceTableHtml -Devices $info.Devices -IncludeSync $true))
        [void]$sb.Append("</div></details>")
    }
    return $sb.ToString()
}

function ConvertTo-JsonLiteral {
    <#
        [V2.2] Chaîne -> littéral JSON entre guillemets, pour le bloc de données du rapport.
        "<" est échappé (<) : une valeur ne peut jamais refermer le bloc <script> qui
        porte les données. Accents et "&" restent tels quels : le masquage des termes
        sensibles (Protect-HtmlSensitiveTerms) doit pouvoir les reconnaître.
    #>
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '""' }
    $s = $Value.Replace('\', '\\').Replace('"', '\"')
    # Remplacement par délégué seulement si nécessaire (rare) : convertir le bloc de code en
    # MatchEvaluator à chaque appel coûtait ~15 s pour 30 000 valeurs
    if ($s -match '[\x00-\x1f<\u2028\u2029]') {
        $s = [regex]::Replace($s, '[\x00-\x1f<\u2028\u2029]', { param($m) '\u{0:x4}' -f [int][char]$m.Value })
    }
    return '"' + $s + '"'
}

function Build-AppAccordionHtml {
    <#
        Un <details>/<summary> natif par application (déplier/replier sans framework),
        avec attribut data-search pour le filtre instantané côté navigateur.

        [V2.2] Les postes ne sont plus écrits en tableau HTML (une ligne par poste et par
        application : des dizaines de Mo de DOM dès que les listes sont complètes) mais dans
        UN bloc de données JSON compact, lu par le navigateur :
          * chaque poste n'y figure qu'une fois ("dev") ; chaque application ne porte que
            les numéros de ses postes ("apps") — 7 423 postes x 50 applications tiennent
            en ~2 Mo au lieu de ~40 Mo de HTML ;
          * la liste s'affiche à l'ouverture de l'application, 20 postes par page, avec une
            recherche (poste, utilisateur, OS) et un export CSV de la liste complète ou filtrée.
        Renvoie le HTML des accordéons suivi du bloc de données.
    #>
    param($Apps, $DevicesByApp, $DetailedIds, [int]$TopNDetailed, $TruncatedApps)
    $sb       = New-Object System.Text.StringBuilder
    $detailed = New-Object System.Collections.Generic.HashSet[string]
    foreach ($id in @($DetailedIds)) { if ($id) { [void]$detailed.Add([string]$id) } }

    # Table des postes (dédoublonnés) et listes de numéros par application
    $devIndex = @{}
    $devJson  = New-Object System.Collections.Generic.List[string]
    $appsJson = New-Object System.Collections.Generic.List[string]
    $key      = 0

    foreach ($app in $Apps) {
        $name      = ConvertTo-HtmlSafe $app.displayName
        $publisher = ConvertTo-HtmlSafe $app.publisher
        $version   = ConvertTo-HtmlSafe $app.version
        $count     = [int]$app.deviceCount
        $search    = ConvertTo-HtmlSafe (("$($app.displayName) $($app.publisher)").ToLower())
        $verChip   = if ($version) { "<span class='chip mono'>v$version</span>" } else { "" }
        $pubHtml   = if ($publisher) { "<span class='app-pub'>$publisher</span>" } else { "<span class='app-pub muted-i'>&Eacute;diteur inconnu</span>" }

        $appSlug = ConvertTo-Slug ([string]$app.displayName)
        $expBtn  = Build-AccordionExportButton -Label ([string]$app.displayName)

        $devices = @()
        if ($detailed.Contains([string]$app.id) -and $DevicesByApp -and $DevicesByApp.ContainsKey($app.id)) { $devices = @($DevicesByApp[$app.id]) }

        if ($devices.Count -gt 0) {
            $key++
            $idx = New-Object System.Collections.Generic.List[int]
            foreach ($d in $devices) {
                if ($null -eq $d) { continue }
                # Boucle la plus sollicitée (un passage par couple application/poste : jusqu'à un
                # million sur un grand parc) : accès direct aux propriétés (insensibles à la casse
                # en PowerShell, sans Get-PropCI) et clé poste + utilisateur ; l'OS et sa version
                # ne sont lus qu'à la première rencontre du poste.
                $dn = [string]$d.deviceName
                $up = [string]$d.userPrincipalName
                $k  = $dn + "`t" + $up
                $i  = $devIndex[$k]
                if ($null -eq $i) {
                    $i = $devJson.Count
                    $devIndex[$k] = $i
                    $devJson.Add('[' + (ConvertTo-JsonLiteral $dn) + ',' + (ConvertTo-JsonLiteral $up) + ',' + (ConvertTo-JsonLiteral ([string]$d.operatingSystem)) + ',' + (ConvertTo-JsonLiteral ([string]$d.osVersion)) + ']')
                }
                $idx.Add($i)
            }
            $appsJson.Add('"' + $key + '":[' + ($idx -join ',') + ']')

            [void]$sb.Append("<details class='acc app' data-search='$search' data-export-name='application-$appSlug' data-app='$key' data-count='$count'><summary><span class='acc-title'>$name</span>$pubHtml$verChip<span class='badge badge-blue'>$count poste(s)</span>$expBtn<span class='chev'>&rsaquo;</span></summary><div class='acc-body'>")
            if ($TruncatedApps -and $TruncatedApps.ContainsKey($app.id)) {
                $why = ConvertTo-HtmlSafe ([string]$TruncatedApps[$app.id])
                if ($TruncatedApps[$app.id] -is [bool]) { $why = "liste limit&eacute;e &agrave; la premi&egrave;re page" }
                [void]$sb.Append("<div class='note note-warn'>&#9888; Liste incompl&egrave;te&nbsp;: $($idx.Count) poste(s) collect&eacute;(s) sur $count annonc&eacute;(s) par Intune ($why). Relancez la g&eacute;n&eacute;ration pour compl&eacute;ter la liste.</div>")
            }
            [void]$sb.Append("<div class='devpager' data-app='$key'><div class='note note-info'>Liste des postes&nbsp;: d&eacute;pliez l'application pour l'afficher (JavaScript requis).</div></div>")
        } else {
            [void]$sb.Append("<details class='acc app' data-search='$search' data-export-name='application-$appSlug'><summary><span class='acc-title'>$name</span>$pubHtml$verChip<span class='badge badge-blue'>$count poste(s)</span>$expBtn<span class='chev'>&rsaquo;</span></summary><div class='acc-body'>")
            if ($detailed.Contains([string]$app.id)) {
                [void]$sb.Append("<div class='note note-info'>Aucun poste associ&eacute; &agrave; cette application dans les donn&eacute;es collect&eacute;es.</div>")
            } else {
                [void]$sb.Append("<div class='note note-info'>D&eacute;tail non charg&eacute;&nbsp;: la collecte a &eacute;t&eacute; limit&eacute;e aux $TopNDetailed applications les plus r&eacute;pandues. Cochez &laquo;&nbsp;Toutes les applications&nbsp;&raquo; dans l'outil (onglet Rapports &amp; Versions) pour les d&eacute;tailler toutes (ou d&eacute;tail absent du fichier import&eacute;).</div>")
            }
        }
        [void]$sb.Append("</div></details>")
    }

    # Bloc de données : guillemets simples sur la balise (le masquage des termes sensibles ne
    # traite que les chaînes JSON entre guillemets doubles, jamais la balise elle-même)
    if ($appsJson.Count -gt 0) {
        [void]$sb.Append("<script type='application/json' id='appDeviceData'>")
        [void]$sb.Append('{"cols":["Poste","Utilisateur","OS","Version OS"],"dev":[')
        [void]$sb.Append(($devJson -join ','))
        [void]$sb.Append('],"apps":{')
        [void]$sb.Append(($appsJson -join ','))
        [void]$sb.Append('}}</script>')
    }
    return $sb.ToString()
}

function Build-InventoryRowsHtml {
    param([array]$Inventory)
    $sb = New-Object System.Text.StringBuilder
    $ok = 0; $check = 0; $unknown = 0
    foreach ($item in $Inventory) {
        $statusKey = "unknown"; $statusLabel = "Non v&eacute;rifiable"
        if     ($item.Status -like "A jour*")     { $statusKey = "ok";    $statusLabel = "&Agrave; jour";      $ok++ }
        elseif ($item.Status -like "A v*")        { $statusKey = "check"; $statusLabel = "&Agrave; v&eacute;rifier"; $check++ }
        else                                      { $unknown++ }
        $nm  = ConvertTo-HtmlSafe $item.DisplayName
        $pub = ConvertTo-HtmlSafe $item.Publisher
        $cur = ConvertTo-HtmlSafe $item.CurrentVersion
        $lat = ConvertTo-HtmlSafe $item.LatestPublicVersion
        $rd  = ConvertTo-HtmlSafe $item.ReleaseDate
        $src = ConvertTo-HtmlSafe $item.Source
        $search = ConvertTo-HtmlSafe (("$($item.DisplayName) $($item.Publisher)").ToLower())
        [void]$sb.Append("<tr data-status='$statusKey' data-search='$search'><td class='cell-strong'>$nm</td><td class='muted'>$pub</td><td class='mono'>$cur</td><td class='mono'>$lat</td><td class='muted'>$rd</td><td><span class='tag'>$src</span></td><td><span class='badge badge-$statusKey'>$statusLabel</span></td></tr>")
    }
    return [PSCustomObject]@{ RowsHtml = $sb.ToString(); CountOk = $ok; CountCheck = $check; CountUnknown = $unknown }
}

function Build-KpiCardHtml {
    <# Une carte du bandeau d'indicateurs : icône, valeur, libellé, ligne secondaire et
       barre de proportion facultative. Toutes les cartes partagent exactement la même
       structure — c'est ce qui rend le bandeau lisible d'un coup d'œil. #>
    param(
        [string]$IconClass, [string]$Icon,
        [string]$Value, [string]$ValueClass = "",
        [string]$Label, [string]$Foot = "", [string]$BarHtml = ""
    )
    $valueClassAttr = if ([string]::IsNullOrWhiteSpace($ValueClass)) { "" } else { " $ValueClass" }
    $footHtml = if ([string]::IsNullOrWhiteSpace($Foot)) { "" } else { "<div class='kpi-foot'>$Foot</div>" }
    return "<div class='kpi'><div class='kpi-ic $IconClass' aria-hidden='true'>$Icon</div><div class='kpi-body'><div class='kpi-num$valueClassAttr'>$Value</div><div class='kpi-lb'>$Label</div>$BarHtml$footHtml</div></div>"
}

function Build-KpiBandHtml {
    <#
        Bandeau d'indicateurs affiché SOUS l'en-tête et AU-DESSUS des onglets : il porte sur
        l'ensemble du parc, pas sur un onglet en particulier, et reste donc visible quel que
        soit l'onglet consulté à l'ouverture.

        Une carte n'apparaît que si la donnée correspondante a réellement été collectée :
        afficher « 0 % chiffré » pour un parc dont l'état BitLocker n'a pas été demandé
        serait un contresens, et le pire des indicateurs est celui auquel on ne peut pas
        se fier.
    #>
    param($Kpi, [bool]$HasCompliance, [bool]$HasRemediation, [bool]$HasBitLocker)
    if (-not $Kpi -or $Kpi.Total -eq 0) { return "" }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<section class='kpi-band' aria-label='Indicateurs cl&eacute;s du parc'><div class='kpis'>")

    # ---- Parc analysé ----
    $osTop = ""
    if ($Kpi.OsFamilies.Count -gt 0) {
        $firstFamily = @($Kpi.OsFamilies.Keys)[0]
        $osTop = "Majoritairement " + (ConvertTo-HtmlSafe $firstFamily)
    }
    [void]$sb.Append((Build-KpiCardHtml -IconClass "ic-slate" -Icon "&#128421;" -Value "$($Kpi.Total)" -Label "Appareils analys&eacute;s" -Foot $osTop))

    # ---- Conformité globale ----
    if ($HasCompliance) {
        $rate  = $Kpi.ComplianceRate
        $tone  = if ($rate -ge 95) { "t-green" } elseif ($rate -ge 80) { "t-amber" } else { "t-red" }
        $color = if ($rate -ge 95) { "#10b981" } elseif ($rate -ge 80) { "#f59e0b" } else { "#ef4444" }
        $bar   = "<div class='kpi-bar' role='img' aria-label='$rate pour cent du parc conforme'><i style='width:$rate%;background:$color'></i></div>"
        $foot  = "$($Kpi.Compliant) conformes &middot; $($Kpi.NonCompliant) non conformes &middot; $($Kpi.GracePeriod) en gr&acirc;ce"
        [void]$sb.Append((Build-KpiCardHtml -IconClass "ic-green" -Icon "&#128737;" -Value "$rate&nbsp;%" -ValueClass $tone -Label "Conformit&eacute; globale du parc" -Foot $foot -BarHtml $bar))
    }

    # ---- Inactivité ----
    $staleTone = if ($Kpi.StaleCount -eq 0) { "t-green" } elseif ($Kpi.StaleCritical -gt 0) { "t-red" } else { "t-amber" }
    $staleFoot = "dont $($Kpi.StaleCritical) sans contact depuis plus de $($Kpi.StaleCriticalDays) jours"
    if ($Kpi.NeverSynced -gt 0) { $staleFoot += " &middot; $($Kpi.NeverSynced) sans date de synchro" }
    [void]$sb.Append((Build-KpiCardHtml -IconClass "ic-amber" -Icon "&#128268;" -Value "$($Kpi.StaleCount)" -ValueClass $staleTone -Label "Inactifs depuis plus de $($Kpi.StaleDays) jours" -Foot $staleFoot))

    # ---- Windows 11 / Windows 10 ----
    if ($Kpi.WindowsTotal -gt 0) {
        $share    = $Kpi.Windows11Share
        $share10  = if ($Kpi.WindowsTotal -gt 0) { [int][Math]::Round(100.0 * $Kpi.Windows10 / $Kpi.WindowsTotal) } else { 0 }
        $shareOth = [Math]::Max(0, 100 - $share - $share10)
        $bar = "<div class='kpi-bar' role='img' aria-label='$($Kpi.Windows11) postes Windows 11 et $($Kpi.Windows10) postes Windows 10'>" +
               "<i style='width:$share%;background:#5B2C8F' title='Windows 11'></i>" +
               "<i style='width:$share10%;background:#f59e0b' title='Windows 10'></i>" +
               "<i style='width:$shareOth%;background:#cbd5e1' title='Autres versions'></i></div>"
        $foot = "$($Kpi.Windows11) sous Windows&nbsp;11 &middot; $($Kpi.Windows10) sous Windows&nbsp;10"
        $tone = if ($share -ge 90) { "t-green" } elseif ($share -ge 50) { "" } else { "t-amber" }
        [void]$sb.Append((Build-KpiCardHtml -IconClass "ic-blue" -Icon "&#128187;" -Value "$share&nbsp;%" -ValueClass $tone -Label "Parc Windows migr&eacute; en 11" -Foot $foot -BarHtml $bar))
    }

    # ---- Fin de support ----
    if ($Kpi.EndOfSupport -gt 0) {
        $eosDate = $script:Windows10EndOfSupport.ToString("dd/MM/yyyy")
        [void]$sb.Append((Build-KpiCardHtml -IconClass "ic-red" -Icon "&#9888;" -Value "$($Kpi.EndOfSupport)" -ValueClass "t-red" -Label "Postes sans support de s&eacute;curit&eacute;" -Foot "Windows&nbsp;10, hors support depuis le $eosDate (sauf LTSC ou ESU)"))
    }

    # ---- Chiffrement ----
    if ($HasBitLocker -and $null -ne $Kpi.EncryptionRate) {
        $encRate  = $Kpi.EncryptionRate
        $encTone  = if ($encRate -ge 98) { "t-green" } elseif ($encRate -ge 85) { "t-amber" } else { "t-red" }
        $encColor = if ($encRate -ge 98) { "#10b981" } elseif ($encRate -ge 85) { "#f59e0b" } else { "#ef4444" }
        $encBar   = "<div class='kpi-bar' role='img' aria-label='$encRate pour cent du parc chiffr&eacute;'><i style='width:$encRate%;background:$encColor'></i></div>"
        [void]$sb.Append((Build-KpiCardHtml -IconClass "ic-blue" -Icon "&#128274;" -Value "$encRate&nbsp;%" -ValueClass $encTone -Label "Volumes syst&egrave;me chiffr&eacute;s" -Foot "$($Kpi.EncryptedCount) poste(s) confirm&eacute;(s) chiffr&eacute;(s)" -BarHtml $encBar))
    }

    # ---- Actions critiques ----
    if ($HasRemediation -and $null -ne $Kpi.CritActions) {
        $critTone = if ($Kpi.CritActions -eq 0) { "t-green" } else { "t-red" }
        $scoreFoot = if ($null -ne $Kpi.AvgAnalyticsScore) { "Score Endpoint Analytics moyen&nbsp;: $($Kpi.AvgAnalyticsScore)/100" } else { "$($Kpi.WarnActions) poste(s) &agrave; surveiller" }
        [void]$sb.Append((Build-KpiCardHtml -IconClass "ic-red" -Icon "&#128680;" -Value "$($Kpi.CritActions)" -ValueClass $critTone -Label "Postes en &eacute;tat critique" -Foot $scoreFoot))
    }

    [void]$sb.Append("</div></section>")
    return $sb.ToString()
}

function Get-OsReleaseColor {
    <# Couleur d'une version d'OS dans la barre de répartition. Le code couleur porte un
       sens : violet pour ce qui est soutenu (Windows 11), ambre/rouge pour ce qui ne l'est
       plus (Windows 10), gris pour le reste. Ce n'est pas une palette décorative. #>
    param([string]$Release, [int]$Index)
    $win11 = @('#5B2C8F', '#7B3FA8', '#9463c0', '#b28ed6')
    $win10 = @('#d97706', '#f59e0b', '#fbbf24', '#fcd34d')
    $other = @('#64748b', '#94a3b8', '#cbd5e1')
    if ($Release -match '(?i)windows 11') { return $win11[$Index % $win11.Count] }
    if ($Release -match '(?i)windows 10') { return $win10[$Index % $win10.Count] }
    return $other[$Index % $other.Count]
}

function Build-OsMixCardHtml {
    <#
        Carte « répartition des versions d'OS » : une barre empilée pour la vue d'ensemble,
        puis le détail chiffré version par version, avec la mention explicite du support.
        La barre seule ne suffit pas — c'est le tableau qui permet de préparer une campagne
        de mise à niveau.
    #>
    param($Kpi)
    if (-not $Kpi -or $Kpi.Total -eq 0) { return "" }

    $releases = $Kpi.WindowsReleases
    $families = $Kpi.OsFamilies
    if (($releases.Count -eq 0) -and ($families.Count -eq 0)) { return "" }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<div class='card'><h2><span class='h-dot violet'></span>R&eacute;partition des versions d'OS</h2>")
    [void]$sb.Append("<p class='sub'>La distinction Windows&nbsp;10 / Windows&nbsp;11 repose sur le num&eacute;ro de build (seuil 22000) et non sur le libell&eacute; d'OS renvoy&eacute; par Intune, identique pour les deux.</p>")

    # ----- Barre empilée des versions Windows -----
    if ($releases.Count -gt 0 -and $Kpi.WindowsTotal -gt 0) {
        [void]$sb.Append("<div class='osbar' role='img' aria-label='R&eacute;partition des versions de Windows sur le parc'>")
        $index = 0
        foreach ($key in $releases.Keys) {
            $count   = [int]$releases[$key]
            $percent = [Math]::Round(100.0 * $count / $Kpi.WindowsTotal, 2)
            $color   = Get-OsReleaseColor -Release $key -Index $index
            $safeKey = ConvertTo-HtmlSafe $key
            $pctText = ([Math]::Round($percent)).ToString()
            [void]$sb.Append("<i style='width:$($percent.ToString([System.Globalization.CultureInfo]::InvariantCulture))%;background:$color' title=`"$safeKey : $count poste(s) ($pctText %)`"></i>")
            $index++
        }
        [void]$sb.Append("</div>")

        [void]$sb.Append("<div class='tbl-scroll'><table class='tbl' id='osTable'><thead><tr><th scope='col'>Version</th><th scope='col'>Postes</th><th scope='col'>Part du parc Windows</th><th scope='col'>Support de s&eacute;curit&eacute;</th></tr></thead><tbody id='osBody'>")
        $index = 0
        foreach ($key in $releases.Keys) {
            $count   = [int]$releases[$key]
            $percent = [int][Math]::Round(100.0 * $count / $Kpi.WindowsTotal)
            $color   = Get-OsReleaseColor -Release $key -Index $index
            $safeKey = ConvertTo-HtmlSafe $key
            if ($key -match '(?i)windows 10' -and (Get-Date) -ge $script:Windows10EndOfSupport) {
                $support = "<span class='badge badge-red'>Hors support</span>"
            } elseif ($key -match '(?i)windows 11 21H2|windows 11 22H2') {
                $support = "<span class='badge badge-amber'>&Agrave; v&eacute;rifier selon l'&eacute;dition</span>"
            } elseif ($key -match '(?i)inconnue|build') {
                $support = "<span class='badge badge-unknown'>Ind&eacute;termin&eacute;</span>"
            } else {
                $support = "<span class='badge badge-ok'>Soutenu</span>"
            }
            $search = ConvertTo-HtmlSafe ($key.ToLower())
            [void]$sb.Append("<tr data-search='$search'><td class='cell-strong'><span class='lg-dot' style='display:inline-block;vertical-align:middle;margin-right:8px;background:$color'></span>$safeKey</td><td class='mono'>$count</td><td class='mono'>$percent&nbsp;%</td><td>$support</td></tr>")
            $index++
        }
        [void]$sb.Append("</tbody></table></div>")
        [void]$sb.Append("<div class='toolbar' style='margin-top:12px;margin-bottom:0'><button type='button' class='btn' onclick=`"exportTableById('osTable','versions-os', false)`">&#11015; Exporter les versions (CSV)</button></div>")
    }

    # ----- Familles d'OS (parc mixte) -----
    if ($families.Count -gt 1) {
        [void]$sb.Append("<p class='sub' style='margin-top:20px'>Familles d'OS pr&eacute;sentes sur le p&eacute;rim&egrave;tre analys&eacute;&nbsp;:</p><div class='chipbar'>")
        foreach ($key in $families.Keys) {
            $safeKey = ConvertTo-HtmlSafe $key
            [void]$sb.Append("<span class='chip'>$safeKey&nbsp;: <b>$($families[$key])</b></span>")
        }
        [void]$sb.Append("</div>")
    }

    [void]$sb.Append("</div>")
    return $sb.ToString()
}

function Build-DashboardHtml {
    param(
        [string]$ClientName, [string]$CompanyName,
        [string]$ContactPerson, [string]$ContactEmail, [string]$ContactPhone,
        [bool]$ShowContactInfo, [bool]$AnonymizeData,
        [int]$TotalDevices,
        [array]$NonCompliantDevices, [array]$GracePeriodDevices,
        $NonCompliantBreakdown, $GracePeriodBreakdown,
        $DiscoveredApps, [array]$TenantAppInventory,
        [int]$TopNDetailed,
        [string]$DataSourceLabel = "API Microsoft Graph",
        # Onglets autorisés dans le rapport (au moins un doit rester actif)
        [bool]$IncludeCompliance = $true,
        [bool]$IncludeDiscoveredApps = $true,
        [bool]$IncludeInventory = $true,
        [bool]$IncludeRemediation = $false,
        $RemediationData = $null,
        $KpiSummary = $null,
        [bool]$HasBitLockerData = $false
    )

    $ic = [System.Globalization.CultureInfo]::InvariantCulture
    $GeneratedAt = (Get-Date).ToString("dd/MM/yyyy 'à' HH:mm")

    # ---------- Statistiques de conformité ----------
    $nc   = $NonCompliantDevices.Count
    $gp   = $GracePeriodDevices.Count
    $okC  = [Math]::Max(0, $TotalDevices - $nc - $gp)
    $rate = if ($TotalDevices -gt 0) { [Math]::Round(100.0 * $okC / $TotalDevices) } else { 100 }
    $pOk  = if ($TotalDevices -gt 0) { [Math]::Round(100.0 * $okC / $TotalDevices) } else { 0 }
    $pNc  = if ($TotalDevices -gt 0) { [Math]::Round(100.0 * $nc  / $TotalDevices) } else { 0 }
    $pGp  = if ($TotalDevices -gt 0) { [Math]::Round(100.0 * $gp  / $TotalDevices) } else { 0 }

    # ---------- Donut SVG (aucune librairie : calcul des arcs en PowerShell) ----------
    $Circ  = 2 * [Math]::PI * 78
    $segOk = if ($TotalDevices -gt 0) { $Circ * $okC / $TotalDevices } else { $Circ }
    $segNc = if ($TotalDevices -gt 0) { $Circ * $nc  / $TotalDevices } else { 0 }
    $segGp = if ($TotalDevices -gt 0) { $Circ * $gp  / $TotalDevices } else { 0 }
    $sC  = ([Math]::Round($Circ, 2)).ToString($ic)
    $sOk = ([Math]::Round($segOk, 2)).ToString($ic)
    $sNc = ([Math]::Round($segNc, 2)).ToString($ic)
    $sGp = ([Math]::Round($segGp, 2)).ToString($ic)
    $oNc = ([Math]::Round(-1 * $segOk, 2)).ToString($ic)
    $oGp = ([Math]::Round(-1 * ($segOk + $segNc), 2)).ToString($ic)

    $ClientSafe  = ConvertTo-HtmlSafe $ClientName
    $CompanySafe = ConvertTo-HtmlSafe $CompanyName

    # ---------- Bandeau d'indicateurs de parc et répartition des OS ----------
    # Calculés ici, en amont de l'assemblage des onglets : le bandeau vit hors des onglets
    # (il porte sur tout le parc) et la carte OS est injectée dans l'onglet Conformité.
    $kpiBandHtml = ""
    $osMixHtml   = ""
    if ($KpiSummary) {
        $hasRemediationKpi = ($IncludeRemediation -and $null -ne $RemediationData)
        $kpiBandHtml = Build-KpiBandHtml -Kpi $KpiSummary -HasCompliance $IncludeCompliance `
                            -HasRemediation $hasRemediationKpi -HasBitLocker $HasBitLockerData
        if ($IncludeCompliance) { $osMixHtml = Build-OsMixCardHtml -Kpi $KpiSummary }
    }

    # Préfixe des fichiers CSV produits par les boutons d'export du rapport. Il est lu
    # côté navigateur depuis un attribut de <body> : le JavaScript reste ainsi une
    # constante littérale, sans interpolation PowerShell à l'intérieur.
    $ReportPrefix = ConvertTo-HtmlSafe ("Intune-" + (ConvertTo-Slug $ClientName))

    # ---------- Fragments d'en-tête ----------
    if ($ShowContactInfo) {
        $cp = ConvertTo-HtmlSafe $ContactPerson
        $ce = ConvertTo-HtmlSafe $ContactEmail
        $ct = ConvertTo-HtmlSafe $ContactPhone
        $contactHtml = "<div class='hdr-right'><div class='lbl'>CONTACT</div><div>$cp</div><div>$ce</div><div>$ct</div><div class='hdr-date'>G&eacute;n&eacute;r&eacute; le $GeneratedAt</div></div>"
    } else {
        $contactHtml = "<div class='hdr-right'><div class='hdr-date'>G&eacute;n&eacute;r&eacute; le $GeneratedAt</div></div>"
    }
    $anonChip   = if ($AnonymizeData) { "<span class='hchip warn'>&#128274; Donn&eacute;es anonymis&eacute;es</span>" } else { "" }
    $devChip    = if ($IncludeCompliance) { "<span class='hchip'>&#128421; $TotalDevices appareil(s) analys&eacute;(s)</span>" } else { "" }
    $srcChip    = "<span class='hchip'>&#128228; Source&nbsp;: " + (ConvertTo-HtmlSafe $DataSourceLabel) + "</span>"
    $anonFooter = if ($AnonymizeData) { " &middot; Donn&eacute;es anonymis&eacute;es." } else { "" }

    # ---------- Onglet 1 : sections conformité ----------
    if ($nc -eq 0 -and $gp -eq 0) {
        $complianceBody = "<div class='success-card'><div class='big'>&#127881;</div><p>Aucun appareil non conforme ni en p&eacute;riode de gr&acirc;ce &mdash; l'int&eacute;gralit&eacute; du parc analys&eacute; est conforme.</p></div>"
    } else {
        $ncSection = Build-ReasonAccordionHtml -Breakdown $NonCompliantBreakdown -Variant red
        $gpSection = Build-ReasonAccordionHtml -Breakdown $GracePeriodBreakdown -Variant amber
        $complianceBody = @"
<div class="card">
  <h2><span class="h-dot red"></span>Appareils Non-Compliant <span class="badge badge-red">$nc</span></h2>
  <p class="sub">Regroup&eacute;s par motif de non-conformit&eacute; &mdash; cliquez sur un motif pour afficher les postes concern&eacute;s. Chaque motif s'exporte s&eacute;par&eacute;ment.</p>
  <div class="toolbar">
    <button type="button" class="btn btn-export" onclick="exportAllAccordions('ncList','non-conformes','Motif de non-conformite',this)">&#11015; Exporter tous les motifs (CSV)</button>
  </div>
  <div id="ncList">
  $ncSection
  </div>
</div>
<div class="card">
  <h2><span class="h-dot amber"></span>Appareils In Grace Period <span class="badge badge-amber">$gp</span></h2>
  <p class="sub">Appareils encore en p&eacute;riode de gr&acirc;ce avant passage Non-Compliant &mdash; regroup&eacute;s par motif.</p>
  <div class="toolbar">
    <button type="button" class="btn btn-export" onclick="exportAllAccordions('gpList','periode-de-grace','Motif de non-conformite',this)">&#11015; Exporter tous les motifs (CSV)</button>
  </div>
  <div id="gpList">
  $gpSection
  </div>
</div>
"@
    }

    # ---------- Onglet 2 : applications découvertes ----------
    $appsTotal = $DiscoveredApps.AllApps.Count
    $detailedN = $DiscoveredApps.DetailedIds.Count
    # [V2.2] Listes de postes complètes : le texte indique pour combien d'applications
    $appsDetailText = if ($detailedN -ge $appsTotal -and $appsTotal -gt 0) { "La liste nominative compl&egrave;te des postes est disponible pour <b>toutes</b> les applications." } `
                      else { "La liste nominative compl&egrave;te des postes est disponible pour les <b>$detailedN</b> applications les plus r&eacute;pandues." }
    $appsHtml  = Build-AppAccordionHtml -Apps $DiscoveredApps.AllApps -DevicesByApp $DiscoveredApps.DevicesByApp -DetailedIds $DiscoveredApps.DetailedIds -TopNDetailed $TopNDetailed -TruncatedApps $DiscoveredApps.TruncatedApps

    # ---------- Onglet 3 : inventaire ----------
    $invBuild = Build-InventoryRowsHtml -Inventory $TenantAppInventory
    $invTotal = $TenantAppInventory.Count

    $tabPillCompliance = if (($nc + $gp) -gt 0) { "<span class='tab-pill'>$($nc + $gp)</span>" } else { "<span class='tab-pill gray'>0</span>" }

    # ---------- CSS (statique, aucune interpolation) ----------
    $css = @'
*{box-sizing:border-box}
html{scroll-behavior:smooth}
body{margin:0;font-family:'Segoe UI',system-ui,-apple-system,sans-serif;background:#eef1f6;color:#0f172a;font-size:14px;line-height:1.5}
.hidden{display:none!important}
.muted{color:#64748b}.muted-i{color:#94a3b8;font-style:italic}
.mono{font-family:'Cascadia Code',Consolas,'Courier New',monospace;font-size:12.5px}
.cell-strong{font-weight:600;color:#1e293b}

/* ===== En-tête ===== */
.hdr{background:linear-gradient(135deg,#452069 0%,#5B2C8F 45%,#7B3FA8 100%);color:#fff;padding:38px 46px 34px}
.hdr-in{max-width:1380px;margin:0 auto;display:flex;justify-content:space-between;gap:28px;align-items:flex-start;flex-wrap:wrap}
.hdr-brand{font-size:33px;font-weight:800;letter-spacing:4px;text-transform:uppercase}
.hdr-sub{margin-top:6px;font-size:14.5px;opacity:.85}
.hdr-chips{margin-top:16px;display:flex;gap:10px;flex-wrap:wrap}
.hchip{background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.28);padding:6px 14px;border-radius:999px;font-size:12.5px;font-weight:600}
.hchip.warn{background:rgba(255,193,7,.22);border-color:rgba(255,193,7,.55)}
.hdr-right{text-align:right;font-size:13px;opacity:.96}
.hdr-right .lbl{font-size:11px;letter-spacing:1.5px;opacity:.75;font-weight:700;margin-bottom:6px}
.hdr-right div+div{margin-top:3px}
.hdr-date{margin-top:12px!important;opacity:.72}

/* ===== Onglets ===== */
.tabs{position:sticky;top:0;z-index:50;background:#fff;box-shadow:0 1px 0 #e2e8f0,0 6px 18px rgba(15,23,42,.05)}
.tabs-in{max-width:1380px;margin:0 auto;display:flex;gap:6px;padding:0 46px;overflow-x:auto}
.tab-btn{appearance:none;background:none;border:0;border-bottom:3px solid transparent;padding:17px 20px 14px;font:inherit;font-weight:600;color:#64748b;cursor:pointer;display:flex;align-items:center;gap:9px;white-space:nowrap;transition:color .15s}
.tab-btn:hover{color:#334155;background:#f8fafc}
.tab-btn.active{color:#5B2C8F;border-bottom-color:#5B2C8F}
.tab-pill{background:#fee2e2;color:#b91c1c;border-radius:999px;padding:2px 9px;font-size:11.5px;font-weight:700}
.tab-pill.gray{background:#f1f5f9;color:#475569}

/* ===== Mise en page ===== */
main{max-width:1380px;margin:0 auto;padding:28px 46px 70px}
.tab-panel{display:none}
.tab-panel.active{display:block;animation:fadeIn .25s ease}
@keyframes fadeIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}

.card{background:#fff;border-radius:16px;box-shadow:0 1px 2px rgba(15,23,42,.05),0 6px 22px rgba(15,23,42,.06);padding:26px 28px;margin-top:22px}
.card h2{margin:0 0 4px;font-size:16.5px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}
.card .sub{margin:0 0 18px;font-size:13px;color:#64748b}
.h-dot{width:10px;height:10px;border-radius:3px;display:inline-block}
.h-dot.red{background:#ef4444}.h-dot.amber{background:#f59e0b}.h-dot.blue{background:#3b82f6}.h-dot.violet{background:#7c3aed}

/* ===== KPI ===== */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(215px,1fr));gap:18px;margin-top:22px}
.kpi{background:#fff;border-radius:16px;box-shadow:0 1px 2px rgba(15,23,42,.05),0 6px 22px rgba(15,23,42,.06);padding:20px 22px;display:flex;align-items:center;gap:16px}
.kpi-ic{width:46px;height:46px;border-radius:13px;display:flex;align-items:center;justify-content:center;font-size:21px;flex-shrink:0}
.kpi-num{font-size:29px;font-weight:800;line-height:1.1}
.kpi-lb{font-size:11.5px;color:#64748b;font-weight:700;text-transform:uppercase;letter-spacing:.5px;margin-top:2px}
.ic-slate{background:#f1f5f9}.ic-green{background:#d1fae5}.ic-red{background:#fee2e2}.ic-amber{background:#fef3c7}.ic-blue{background:#dbeafe}
.t-green{color:#059669}.t-red{color:#dc2626}.t-amber{color:#d97706}

/* ===== Bloc donut + légende ===== */
.duo{display:grid;grid-template-columns:270px 1fr;gap:36px;align-items:center}
.donut-num{font-size:34px;font-weight:800;fill:#0f172a}
.donut-sub{font-size:12.5px;fill:#64748b;font-weight:600}
.legend{display:flex;flex-direction:column;gap:14px;max-width:520px}
.lg-row{display:flex;align-items:center;gap:12px;font-size:14px}
.lg-dot{width:13px;height:13px;border-radius:4px;flex-shrink:0}
.lg-nm{flex:1;font-weight:600;color:#334155}
.lg-ct{font-weight:700}
.lg-pc{color:#94a3b8;font-size:12.5px;min-width:46px;text-align:right}
.bar{height:7px;border-radius:99px;background:#eef2f7;overflow:hidden;margin-top:6px}
.bar i{display:block;height:100%;border-radius:99px}

/* ===== Accordéons ===== */
.acc{border:1px solid #e5e9f0;border-radius:13px;margin-top:12px;overflow:hidden;background:#fff;transition:box-shadow .15s}
.acc[open]{box-shadow:0 4px 16px rgba(15,23,42,.07)}
.acc>summary{list-style:none;cursor:pointer;display:flex;align-items:center;gap:14px;padding:15px 20px;background:#fbfcfe;user-select:none}
.acc>summary::-webkit-details-marker{display:none}
.acc>summary:hover{background:#f3f6fb}
.acc-title{font-weight:600;color:#1e293b;flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.acc-body{border-top:1px solid #edf0f5;padding:16px 20px 20px}
.chev{color:#94a3b8;font-size:22px;line-height:1;transition:transform .18s ease;flex-shrink:0}
.acc[open]>summary .chev{transform:rotate(90deg)}
.dot{width:9px;height:9px;border-radius:99px;flex-shrink:0}
.dot-red{background:#ef4444}.dot-amber{background:#f59e0b}
.app-pub{color:#64748b;font-size:12.5px;max-width:230px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.chip{background:#f1f5f9;color:#475569;border-radius:8px;padding:3px 9px;font-size:11.5px;font-weight:600;white-space:nowrap}

/* ===== Badges & notes ===== */
.badge{border-radius:999px;padding:4px 12px;font-size:12px;font-weight:700;white-space:nowrap}
.badge-red{background:#fee2e2;color:#b91c1c}
.badge-amber{background:#fef3c7;color:#b45309}
.badge-blue{background:#dbeafe;color:#1d4ed8}
.badge-ok{background:#d1fae5;color:#047857}
.tag{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600;white-space:nowrap;margin:1px}
.tag-red{background:#fee2e2;color:#b91c1c}
.tag-amber{background:#fef3c7;color:#b45309}
.tag-ok{background:#d1fae5;color:#047857}
.act-row{display:flex;gap:5px;flex-wrap:wrap}
.act-btn{border:1px solid #d8dce6;background:#fff;color:#334155;border-radius:6px;padding:4px 9px;font-size:11px;font-weight:600;cursor:pointer;white-space:nowrap;transition:background .12s,border-color .12s}
.act-btn:hover{background:#f1f5f9;border-color:#94a3b8}
.act-btn.copied{background:#d1fae5;border-color:#10b981;color:#047857}
.badge-check{background:#fef3c7;color:#b45309}
.badge-unknown{background:#f1f5f9;color:#64748b}
.tag{background:#f8fafc;border:1px solid #e2e8f0;border-radius:7px;padding:2px 8px;font-size:11.5px;color:#475569;white-space:nowrap}
.note{border-radius:10px;padding:11px 15px;font-size:12.8px;margin-bottom:13px}
.note-warn{background:#fffbeb;border:1px solid #fde68a;color:#92400e}
.note-info{background:#f8fafc;border:1px dashed #cbd5e1;color:#64748b;margin-bottom:0}
.empty{padding:36px;text-align:center;color:#94a3b8;font-style:italic}
.success-card{background:linear-gradient(135deg,#ecfdf5,#f0fdf4);border:1px solid #a7f3d0;border-radius:16px;padding:34px;text-align:center;margin-top:22px}
.success-card .big{font-size:38px}
.success-card p{margin:10px 0 0;color:#065f46;font-weight:600;font-size:15px}

/* ===== Barres d'outils ===== */
.toolbar{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:8px}
.search{flex:1;min-width:230px;max-width:430px;padding:10px 15px;border:1.5px solid #dfe4ec;border-radius:11px;font:inherit;font-size:13.5px;background:#fbfcfe;transition:border-color .15s,box-shadow .15s}
.search:focus{outline:none;border-color:#7B3FA8;box-shadow:0 0 0 3.5px rgba(123,63,168,.13);background:#fff}
.btn{appearance:none;border:1.5px solid #dfe4ec;background:#fff;border-radius:10px;padding:9px 15px;font:inherit;font-size:12.8px;font-weight:600;color:#475569;cursor:pointer;transition:all .13s}
.btn:hover{background:#f8fafc;border-color:#c7cfdc}
.chipbar{display:flex;gap:7px;flex-wrap:wrap}
.fchip{appearance:none;border:1.5px solid #dfe4ec;background:#fff;border-radius:999px;padding:7px 15px;font:inherit;font-size:12.3px;font-weight:600;color:#64748b;cursor:pointer;transition:all .13s}
.fchip:hover{border-color:#b9c3d3}
.fchip.active{background:#5B2C8F;border-color:#5B2C8F;color:#fff}
.count-info{font-size:12.8px;color:#64748b;margin-left:auto}
.count-info b{color:#1e293b}

/* ===== Tableaux ===== */
.tbl-scroll{overflow:auto;max-height:430px;border:1px solid #e9edf3;border-radius:12px}
.tbl-scroll.tall{max-height:660px}
.tbl{width:100%;border-collapse:collapse;font-size:13px}
.tbl thead th{position:sticky;top:0;background:#f6f8fb;color:#475569;font-size:11.3px;text-transform:uppercase;letter-spacing:.05em;text-align:left;padding:12px 16px;border-bottom:1px solid #e2e8f0;z-index:2}
.tbl thead th.sortable{cursor:pointer;user-select:none}
.tbl thead th.sortable:hover{color:#5B2C8F}
.tbl thead th .arr{margin-left:5px;font-size:9px;color:#7B3FA8}
.tbl tbody td{padding:11px 16px;border-bottom:1px solid #f0f3f8;vertical-align:middle}
.tbl tbody tr:last-child td{border-bottom:0}
.tbl tbody tr:hover td{background:#fafbfe}

footer{text-align:center;color:#94a3b8;font-size:12px;padding:30px 20px 44px}
/* ===== Bandeau KPI global (sous l'en-tete, au-dessus des onglets) ===== */
.kpi-band{max-width:1380px;margin:0 auto;padding:24px 46px 4px}
.kpi-band .kpis{margin-top:0}
.kpi-body{min-width:0;flex:1}
.kpi-foot{font-size:11.5px;color:#64748b;margin-top:7px;line-height:1.45}
.kpi-bar{display:flex;height:6px;border-radius:99px;overflow:hidden;background:#eef2f7;margin-top:9px}
.kpi-bar i{display:block;height:100%}

/* ===== Repartition des versions d'OS ===== */
.osbar{display:flex;height:16px;border-radius:99px;overflow:hidden;background:#eef2f7;margin:4px 0 20px}
.osbar i{display:block;height:100%;transition:filter .15s}
.osbar i:hover{filter:brightness(1.12)}

/* ===== Recherche globale ===== */
.gbar{max-width:1380px;margin:0 auto;display:flex;align-items:center;gap:12px;padding:0 46px 12px;flex-wrap:wrap}
.gsearch-wrap{position:relative;flex:1;min-width:260px;max-width:560px}
.gsearch{width:100%;padding:10px 38px 10px 40px;border:1.5px solid #dfe4ec;border-radius:11px;font:inherit;font-size:13.5px;background:#fbfcfe;transition:border-color .15s,box-shadow .15s}
.gsearch:focus{outline:none;border-color:#7B3FA8;box-shadow:0 0 0 3.5px rgba(123,63,168,.13);background:#fff}
.gsearch-ic{position:absolute;left:13px;top:50%;transform:translateY(-50%);font-size:14px;opacity:.55;pointer-events:none}
.gsearch-x{position:absolute;right:8px;top:50%;transform:translateY(-50%);appearance:none;border:0;background:#e9edf4;color:#475569;width:22px;height:22px;border-radius:50%;font-size:12px;line-height:1;cursor:pointer;display:flex;align-items:center;justify-content:center}
.gsearch-x:hover{background:#d8dfea;color:#1e293b}
.gsearch-hint{font-size:12.5px;color:#64748b;font-weight:600}
.gsearch-hint.gsearch-empty{color:#b91c1c}
.gsearch-kbd{font-size:11px;color:#94a3b8;margin-left:auto}
.gsearch-kbd kbd{background:#f1f5f9;border:1px solid #e2e8f0;border-bottom-width:2px;border-radius:5px;padding:1px 5px;font-family:inherit;font-size:10.5px;color:#475569}
.tab-hit{background:#ede4f7;color:#5B2C8F;border-radius:999px;padding:2px 8px;font-size:11px;font-weight:700;margin-left:2px}
.tab-btn.no-hit{opacity:.45}
.tab-btn.no-hit .tab-hit{background:#f1f5f9;color:#94a3b8}

/* ===== Boutons d'export ===== */
.acc-exp{appearance:none;border:1px solid #dfe4ec;background:#fff;color:#5B2C8F;border-radius:7px;padding:4px 10px;font:inherit;font-size:11px;font-weight:700;cursor:pointer;white-space:nowrap;flex-shrink:0;transition:background .12s,border-color .12s}
.acc-exp:hover{background:#f5f0fb;border-color:#b9a1d4}
.acc-exp.copied{background:#d1fae5;border-color:#10b981;color:#047857}
.btn.copied{background:#d1fae5;border-color:#10b981;color:#047857}
.btn-export{color:#5B2C8F;border-color:#dcd0ec}
.btn-export:hover{background:#f5f0fb;border-color:#b9a1d4}

/* ===== Accessibilite ===== */
.sr-only{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0 0 0 0);white-space:nowrap;border:0}
.skip-link{position:absolute;left:-9999px;top:0;z-index:200;background:#5B2C8F;color:#fff;padding:12px 20px;border-radius:0 0 10px 0;font-weight:700;text-decoration:none}
.skip-link:focus{left:0}
:focus-visible{outline:3px solid #7B3FA8;outline-offset:2px;border-radius:4px}
.tab-btn:focus-visible{outline-offset:-3px}
.acc>summary:focus-visible{outline:3px solid #7B3FA8;outline-offset:-3px}
@media(prefers-reduced-motion:reduce){
*,*::before,*::after{animation-duration:.001ms!important;animation-iteration-count:1!important;transition-duration:.001ms!important;scroll-behavior:auto!important}
}

/* ===== Impression / export PDF ===== */
@media print{
body{background:#fff;font-size:11px}
.tabs,.gbar,.toolbar,.act-row,.acc-exp,.skip-link{display:none!important}
.tab-panel{display:block!important;page-break-before:always}
.tab-panel:first-of-type{page-break-before:avoid}
.card,.kpi{box-shadow:none;border:1px solid #d8dee9}
.tbl-scroll,.tbl-scroll.tall{max-height:none;overflow:visible}
.acc{page-break-inside:avoid}
details.acc{open:open}
.hdr{background:#452069!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}
}

@media(max-width:920px){
.hdr{padding:26px 20px}
.tabs-in{padding:0 12px}
.gbar{padding:0 12px 10px}
.kpi-band{padding:16px 16px 0}
main{padding:18px 16px 60px}
.hdr-right{text-align:left}
.duo{grid-template-columns:1fr;justify-items:center}
.card{padding:20px 18px}
.count-info{margin-left:0}
.gsearch-kbd{display:none}
.gsearch-wrap{max-width:none}
}
/* [V2.2] Postes par application : pagination (20 postes par page), recherche, export */
.pg-bar{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px}
.pg-bar .search{max-width:360px;padding:8px 13px;font-size:13px}
.pg-info{font-size:12.5px;color:#64748b;margin-left:auto}
.pg-info b{color:#1e293b}
.tbl-scroll.pg-scroll{max-height:none}
.pg-nav{display:flex;align-items:center;justify-content:center;gap:6px;margin-top:11px}
.pg-btn{appearance:none;border:1.5px solid #dfe4ec;background:#fff;border-radius:8px;min-width:36px;padding:6px 10px;font:inherit;font-size:13px;font-weight:700;color:#475569;cursor:pointer;transition:all .13s}
.pg-btn:hover:not([disabled]){background:#f5f0fb;border-color:#b9a1d4;color:#5B2C8F}
.pg-btn:focus-visible{outline:3px solid #7B3FA8;outline-offset:1px}
.pg-btn[disabled]{opacity:.4;cursor:default}
.pg-page{font-size:12.8px;font-weight:600;color:#475569;padding:0 10px;min-width:110px;text-align:center}
.tbl tbody td.pg-empty{padding:22px;text-align:center;color:#94a3b8;font-style:italic}
@media (max-width:720px){.pg-info{margin-left:0;width:100%}}
@media print{.pg-bar,.pg-nav{display:none!important}}
'@

    # ---------- JavaScript (statique, vanilla ES5, aucune dépendance) ----------
    $js = @'
/* =====================================================================
   Rapport Intune - moteur d'interface.
   Vanilla, aucune dependance, aucun acces reseau : le fichier doit rester
   consultable hors ligne, ouvert depuis un partage ou une piece jointe.

   Trois principes :
     1. L'index de recherche est construit UNE fois au chargement. Relire le
        DOM a chaque frappe coute cher des que le parc depasse le millier de
        lignes ; ici on ne manipule plus que des chaines deja normalisees.
     2. La recherche est insensible aux accents. Un administrateur tape
        "remediation", pas "remediation" avec les accents au bon endroit.
     3. Filtre global et filtres locaux (par onglet) se composent au lieu de
        s'ecraser : une recherche globale restreint la liste, le filtre
        "Critiques" la restreint encore.
   ===================================================================== */
(function () {
  'use strict';

  var DEBOUNCE_MS = 120;
  var state = { global: '', apps: '', inv: '', rem: '', invStatus: 'all', remStatus: 'all' };
  var panels = [], allRows = [], allAccs = [];
  var reportPrefix = 'rapport-intune';
  var timer = null;
  var sortState = {};

  function byId(id) { return document.getElementById(id); }
  function qsa(selector, root) { return Array.prototype.slice.call((root || document).querySelectorAll(selector)); }

  function norm(value) {
    var s = (value === null || value === undefined) ? '' : String(value);
    s = s.toLowerCase();
    if (s.normalize) { s = s.normalize('NFD').replace(/[\u0300-\u036f]/g, ''); }
    return s.replace(/\s+/g, ' ').trim();
  }

  function closestAcc(node) {
    if (node.closest) { return node.closest('details.acc'); }
    var current = node;
    while (current && current !== document.body) {
      if (current.tagName === 'DETAILS' && current.className.indexOf('acc') > -1) { return current; }
      current = current.parentNode;
    }
    return null;
  }

  /* ---------------------------------------------------------------
     Index de recherche
     --------------------------------------------------------------- */
  function buildIndex() {
    allRows = [];
    allAccs = [];
    panels = qsa('.tab-panel').map(function (panelEl) {
      var panel = { id: panelEl.id, el: panelEl, rows: [], accs: [] };

      qsa('details.acc', panelEl).forEach(function (accEl) {
        var summary = accEl.querySelector('summary');
        var acc = {
          el: accEl,
          panel: panel,
          isApp: accEl.className.indexOf('app') > -1,
          hay: norm((accEl.getAttribute('data-search') || '') + ' ' + (summary ? summary.textContent : '')),
          rows: [],
          selfHit: true,
          visible: true,
          autoOpened: false,
          dataKey: accEl.getAttribute('data-app'),
          pager: null,
          pagerGlobal: ''
        };
        accEl.__acc = acc;
        panel.accs.push(acc);
        allAccs.push(acc);
      });

      qsa('tbody tr', panelEl).forEach(function (trEl) {
        var accEl = closestAcc(trEl);
        var host = accEl ? accEl.__acc : null;
        var row = {
          el: trEl,
          panel: panel,
          acc: host,
          bodyId: trEl.parentNode ? (trEl.parentNode.id || '') : '',
          status: trEl.getAttribute('data-status') || '',
          hay: norm(trEl.getAttribute('data-search') || trEl.textContent),
          visible: true
        };
        if (host) { host.rows.push(row); }
        panel.rows.push(row);
        allRows.push(row);
      });

      return panel;
    });
  }

  /* ---------------------------------------------------------------
     Filtrage
     --------------------------------------------------------------- */
  function localRowMatch(row) {
    if (row.bodyId === 'invBody') {
      if (state.invStatus !== 'all' && row.status !== state.invStatus) { return false; }
      if (state.inv && row.hay.indexOf(state.inv) === -1) { return false; }
      return true;
    }
    if (row.bodyId === 'remBody') {
      if (state.remStatus !== 'all' && row.status !== state.remStatus) { return false; }
      if (state.rem && row.hay.indexOf(state.rem) === -1) { return false; }
      return true;
    }
    return true;
  }

  function localAccMatch(acc) {
    if (acc.isApp && state.apps) { return acc.hay.indexOf(state.apps) > -1; }
    return true;
  }

  function apply() {
    var query = state.global;

    /* Un accordeon dont le LIBELLE correspond montre toutes ses lignes : chercher
       "batterie" doit ouvrir l'action batterie complete, pas seulement les postes
       dont le nom contiendrait le mot. */
    allAccs.forEach(function (acc) {
      acc.selfHit = !query || acc.hay.indexOf(query) > -1;
    });

    allRows.forEach(function (row) {
      var globalHit = !query || row.hay.indexOf(query) > -1 || (row.acc && row.acc.selfHit);
      row.visible = globalHit && localRowMatch(row);
      row.el.classList.toggle('hidden', !row.visible);
    });

    allAccs.forEach(function (acc) {
      var childHit = false;
      if (acc.dataKey !== null) {
        /* [V2.2] Postes lus dans les donnees : la recherche globale trouve un poste
           meme s'il n'est pas sur la page affichee ; la liste ouverte se filtre sur
           la recherche (sauf si c'est l'application elle-meme qui correspond). */
        acc.pagerGlobal = (query && !acc.selfHit) ? query : '';
        if (acc.pagerGlobal) { childHit = appHasDeviceMatch(acc.dataKey, query); }
        if (acc.pager) { acc.pager.setGlobal(acc.pagerGlobal); }
      }
      for (var i = 0; i < acc.rows.length; i++) {
        if (acc.rows[i].visible) { childHit = true; break; }
      }
      var show = (acc.selfHit || childHit) && localAccMatch(acc);
      acc.visible = show;
      acc.el.classList.toggle('hidden', !show);

      if (query) {
        /* Deplie ce que la recherche a trouve a l'interieur : sinon le resultat
           existe mais reste invisible sous un accordeon ferme. */
        if (show && childHit && !acc.el.open) { acc.el.open = true; acc.autoOpened = true; }
      } else if (acc.autoOpened) {
        acc.el.open = false;
        acc.autoOpened = false;
      }
    });

    updateCounters();
  }

  function schedule() {
    if (timer) { clearTimeout(timer); }
    timer = setTimeout(function () { timer = null; apply(); }, DEBOUNCE_MS);
  }

  /* ---------------------------------------------------------------
     Compteurs et indicateurs de resultats
     --------------------------------------------------------------- */
  function setText(id, value) {
    var el = byId(id);
    if (el) { el.textContent = value; }
  }

  function panelHits(panel) {
    var appAccs = panel.accs.filter(function (a) { return a.isApp; });
    var n = 0;
    if (appAccs.length > 0) {
      appAccs.forEach(function (a) { if (a.visible) { n++; } });
      return n;
    }
    panel.rows.forEach(function (r) { if (r.visible) { n++; } });
    if (n === 0) { panel.accs.forEach(function (a) { if (a.visible && a.selfHit) { n++; } }); }
    return n;
  }

  function updateCounters() {
    var appVisible = 0, invVisible = 0, remVisible = 0;
    allAccs.forEach(function (a) { if (a.isApp && a.visible) { appVisible++; } });
    allRows.forEach(function (r) {
      if (!r.visible) { return; }
      if (r.bodyId === 'invBody') { invVisible++; }
      else if (r.bodyId === 'remBody') { remVisible++; }
    });
    setText('appCount', appVisible);
    setText('invCount', invVisible);
    setText('remCount', remVisible);

    var query = state.global;
    var totalHits = 0, tabsWithHits = 0;
    panels.forEach(function (panel) {
      var hits = panelHits(panel);
      if (query) {
        totalHits += hits;
        if (hits > 0) { tabsWithHits++; }
      }
      var btn = document.querySelector('.tab-btn[data-tab="' + panel.id + '"]');
      if (!btn) { return; }
      var badge = btn.querySelector('.tab-hit');
      if (!query) {
        if (badge) { badge.parentNode.removeChild(badge); }
        btn.classList.remove('no-hit');
        return;
      }
      if (!badge) {
        badge = document.createElement('span');
        badge.className = 'tab-hit';
        btn.appendChild(badge);
      }
      badge.textContent = hits;
      btn.classList.toggle('no-hit', hits === 0);
    });

    var hint = byId('gHint');
    if (hint) {
      if (!query) {
        hint.textContent = '';
        hint.classList.remove('gsearch-empty');
      } else if (totalHits === 0) {
        hint.textContent = 'Aucun resultat';
        hint.classList.add('gsearch-empty');
      } else {
        hint.textContent = totalHits + ' resultat(s) dans ' + tabsWithHits + ' onglet(s)';
        hint.classList.remove('gsearch-empty');
      }
    }
    var clearBtn = byId('gClear');
    if (clearBtn) { clearBtn.classList.toggle('hidden', !query); }
  }

  /* ---------------------------------------------------------------
     API appelee par les gestionnaires inline du document
     --------------------------------------------------------------- */
  window.globalSearchInput = function (el) { state.global = norm(el.value); schedule(); };

  window.clearGlobalSearch = function () {
    var el = byId('gSearch');
    if (el) { el.value = ''; el.focus(); }
    state.global = '';
    apply();
  };

  window.filterApps = function () { var el = byId('appSearch'); state.apps = el ? norm(el.value) : ''; schedule(); };
  window.filterInv  = function () { var el = byId('invSearch'); state.inv  = el ? norm(el.value) : ''; schedule(); };
  window.filterRem  = function () { var el = byId('remSearch'); state.rem  = el ? norm(el.value) : ''; schedule(); };

  function setChips(containerSelector, activeBtn) {
    qsa(containerSelector + ' .fchip').forEach(function (chip) {
      var on = (chip === activeBtn);
      chip.classList.toggle('active', on);
      chip.setAttribute('aria-pressed', on ? 'true' : 'false');
    });
  }

  window.setInvStatus = function (btn) { state.invStatus = btn.getAttribute('data-f'); setChips('#invChips', btn); apply(); };
  window.setRemStatus = function (btn) { state.remStatus = btn.getAttribute('data-f'); setChips('#remChips', btn); apply(); };

  window.setAppsOpen = function (open) {
    allAccs.forEach(function (acc) {
      if (acc.isApp && acc.visible) { acc.el.open = open; acc.autoOpened = false; }
    });
  };

  window.showTab = function (btn) {
    var id = btn.getAttribute('data-tab');
    qsa('.tab-btn').forEach(function (b) {
      var active = (b === btn);
      b.classList.toggle('active', active);
      b.setAttribute('aria-selected', active ? 'true' : 'false');
      b.setAttribute('tabindex', active ? '0' : '-1');
    });
    qsa('.tab-panel').forEach(function (p) {
      p.classList.toggle('active', p.id === id);
    });
    window.scrollTo(0, 0);
  };

  /* ---------------------------------------------------------------
     Tri des tableaux
     --------------------------------------------------------------- */
  function cellText(tr, index) {
    var cell = tr.children[index];
    return cell ? cell.textContent.replace(/\u00a0/g, ' ').trim() : '';
  }

  function sortTable(tableId, bodyId, th) {
    var table = byId(tableId), body = byId(bodyId);
    if (!table || !body) { return; }
    var index = parseInt(th.getAttribute('data-col'), 10);
    if (isNaN(index)) { return; }

    var st = sortState[tableId] || { col: -1, dir: 1 };
    st.dir = (st.col === index) ? -st.dir : 1;
    st.col = index;
    sortState[tableId] = st;

    var rows = qsa('tr', body);
    /* Tri "naturel" : 10 apres 9, "12 Go" apres "9 Go", et insensible aux accents. */
    rows.sort(function (a, b) {
      return cellText(a, index).localeCompare(cellText(b, index), 'fr', { numeric: true, sensitivity: 'base' }) * st.dir;
    });
    /* Reinsertion via un fragment : une seule ecriture dans le DOM au lieu d'une par ligne. */
    var fragment = document.createDocumentFragment();
    rows.forEach(function (row) { fragment.appendChild(row); });
    body.appendChild(fragment);

    qsa('th', table).forEach(function (header) {
      var arrow = header.querySelector('.arr');
      if (arrow) { arrow.textContent = ''; }
      header.removeAttribute('aria-sort');
    });
    var mine = th.querySelector('.arr');
    if (mine) { mine.textContent = (st.dir === 1) ? '\u25B2' : '\u25BC'; }
    th.setAttribute('aria-sort', (st.dir === 1) ? 'ascending' : 'descending');
  }

  window.sortInv = function (th) { sortTable('invTable', 'invBody', th); };
  window.sortRem = function (th) { sortTable('remTable', 'remBody', th); };

  /* ---------------------------------------------------------------
     Export CSV
     --------------------------------------------------------------- */
  function slug(text) {
    var s = norm(text).replace(/[^a-z0-9]+/g, '-').replace(/^-+/, '').replace(/-+$/, '');
    return (s || 'export').substring(0, 80);
  }

  function stamp() {
    var d = new Date();
    function pad(n) { return (n < 10 ? '0' : '') + n; }
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate()) + '-' + pad(d.getHours()) + pad(d.getMinutes());
  }

  function fileBase(name) { return slug(reportPrefix) + '_' + slug(name) + '_' + stamp(); }

  function csvCell(value) {
    var s = (value === null || value === undefined) ? '' : String(value);
    s = s.replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim();
    return '"' + s.replace(/"/g, '""') + '"';
  }

  function download(filename, text, mime) {
    var blob = new Blob([text], { type: mime });
    if (window.navigator && window.navigator.msSaveBlob) { window.navigator.msSaveBlob(blob, filename); return; }
    var url = URL.createObjectURL(blob);
    var link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.rel = 'noopener';
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    setTimeout(function () { document.body.removeChild(link); URL.revokeObjectURL(url); }, 500);
  }

  /* CSV pense pour Excel francophone : BOM UTF-8 (accents corrects a l'ouverture) et
     directive "sep=;" (Excel n'utilise pas la virgule dans les locales FR). Un outil qui
     ignore la directive lit simplement une premiere ligne "sep=;", sans casse. */
  function downloadCsv(baseName, matrix) {
    var lines = matrix.map(function (row) { return row.map(csvCell).join(';'); });
    var content = '\ufeff' + 'sep=;\r\n' + lines.join('\r\n') + '\r\n';
    download(baseName + '.csv', content, 'text/csv;charset=utf-8;');
  }

  function cleanHeader(th) {
    return (th.getAttribute('data-export') || th.textContent).replace(/[\u25B2\u25BC]/g, '').replace(/\s+/g, ' ').trim();
  }

  function tableToMatrix(table, visibleOnly) {
    var skip = {};
    var header = [];
    qsa('thead th', table).forEach(function (th, i) {
      if (th.className.indexOf('no-export') > -1) { skip[i] = true; return; }
      header.push(cleanHeader(th));
    });
    var matrix = [header];
    qsa('tbody tr', table).forEach(function (tr) {
      if (visibleOnly && tr.className.indexOf('hidden') > -1) { return; }
      var cells = [];
      for (var i = 0; i < tr.children.length; i++) {
        if (skip[i]) { continue; }
        cells.push(tr.children[i].textContent);
      }
      matrix.push(cells);
    });
    return matrix;
  }

  function flash(btn, message) {
    if (!btn) { return; }
    if (btn.getAttribute('data-label') === null) { btn.setAttribute('data-label', btn.innerHTML); }
    btn.innerHTML = message;
    btn.classList.add('copied');
    setTimeout(function () {
      var original = btn.getAttribute('data-label');
      if (original !== null) { btn.innerHTML = original; btn.removeAttribute('data-label'); }
      btn.classList.remove('copied');
    }, 1800);
  }

  window.exportAccordion = function (btn, event) {
    /* Le bouton vit dans un <summary> : sans neutralisation, le clic replierait
       l'accordeon en meme temps qu'il lance l'export. */
    if (event) { event.preventDefault(); event.stopPropagation(); }
    var acc = closestAcc(btn);
    if (!acc) { return; }
    /* [V2.2] Application : TOUS ses postes, lus dans les donnees (le tableau affiche
       n'en montre que 20) */
    var key = acc.getAttribute('data-app');
    if (key !== null) {
      var list = appIndexes(key);
      if (list.length === 0) { flash(btn, 'Rien a exporter'); return; }
      downloadCsv(fileBase(acc.getAttribute('data-export-name') || 'application'), devicesMatrix(list));
      flash(btn, '\u2713 ' + list.length + ' poste(s)');
      return;
    }
    var table = acc.querySelector('table.tbl');
    if (!table) { flash(btn, 'Rien a exporter'); return; }
    downloadCsv(fileBase(acc.getAttribute('data-export-name') || 'extrait'), tableToMatrix(table, false));
    flash(btn, '\u2713 Exporte');
  };

  window.exportTableById = function (tableId, name, visibleOnly, btn) {
    var table = byId(tableId);
    if (!table) { return; }
    var matrix = tableToMatrix(table, !!visibleOnly);
    if (matrix.length < 2) { flash(btn, 'Rien a exporter'); return; }
    downloadCsv(fileBase(name), matrix);
    flash(btn, '\u2713 Exporte');
  };

  /* Export "a plat" de tous les accordeons d'une section : une ligne par poste, avec
     l'intitule de l'action (ou du motif) en premiere colonne. C'est le format qu'on
     ouvre dans un tableur pour repartir le travail entre techniciens. */
  window.exportAllAccordions = function (containerId, name, firstColumnLabel, btn) {
    var container = byId(containerId);
    if (!container) { return; }
    var matrix = null;
    qsa('details.acc', container).forEach(function (accEl) {
      if (accEl.className.indexOf('hidden') > -1) { return; }
      var titleEl = accEl.querySelector('.acc-title');
      var label = titleEl ? titleEl.textContent.replace(/\s+/g, ' ').trim() : '';
      var key = accEl.getAttribute('data-app');
      var sub;
      if (key !== null) {
        /* [V2.2] Application : postes lus dans les donnees, filtres d'affichage compris */
        sub = devicesMatrix(appExportIndexes(accEl.__acc, key));
      } else {
        var table = accEl.querySelector('table.tbl');
        if (!table) { return; }
        sub = tableToMatrix(table, true);
      }
      if (sub.length < 2) { return; }
      if (!matrix) { matrix = [[firstColumnLabel].concat(sub[0])]; }
      for (var i = 1; i < sub.length; i++) { matrix.push([label].concat(sub[i])); }
    });
    if (!matrix || matrix.length < 2) { flash(btn, 'Rien a exporter'); return; }
    downloadCsv(fileBase(name), matrix);
    flash(btn, '\u2713 Exporte');
  };

  /* ---------------------------------------------------------------
     [V2.2] Postes par application : liste complete, 20 postes par page.
     Les postes sont lus dans le bloc JSON #appDeviceData (chaque poste une
     seule fois, chaque application = les numeros de ses postes) et ne sont
     rendus que pour l'application ouverte, page par page : le DOM reste
     leger meme avec des listes de plusieurs milliers de postes.
     --------------------------------------------------------------- */
  var PAGE_SIZE = 20;
  var appData = null;
  var devHay = [];
  var devMatch = { q: null, flags: null };

  function loadAppData() {
    if (appData) { return appData; }
    appData = { cols: ['Poste', 'Utilisateur', 'OS', 'Version OS'], dev: [], apps: {} };
    var holder = byId('appDeviceData');
    if (holder) {
      try {
        var parsed = JSON.parse(holder.textContent || '{}');
        if (parsed && parsed.dev && parsed.apps) { appData = parsed; }
      } catch (e) { /* donnees illisibles : listes vides, le reste du rapport fonctionne */ }
    }
    devHay = new Array(appData.dev.length);
    return appData;
  }

  function deviceHay(i) {
    var h = devHay[i];
    if (h === undefined) { h = norm(appData.dev[i].join(' ')); devHay[i] = h; }
    return h;
  }

  function appIndexes(key) { var d = loadAppData(); return d.apps[key] || []; }

  /* Postes correspondant a la recherche globale : calcules une fois par requete
     (quelques milliers de postes), puis simple lecture pour chaque application. */
  function deviceMatchFlags(query) {
    if (devMatch.q === query) { return devMatch.flags; }
    var d = loadAppData();
    var flags = new Uint8Array(d.dev.length);
    for (var i = 0; i < d.dev.length; i++) { if (deviceHay(i).indexOf(query) > -1) { flags[i] = 1; } }
    devMatch = { q: query, flags: flags };
    return flags;
  }

  function appHasDeviceMatch(key, query) {
    var flags = deviceMatchFlags(query), list = appIndexes(key);
    for (var i = 0; i < list.length; i++) { if (flags[list[i]]) { return true; } }
    return false;
  }

  function filterIndexes(all, local, global) {
    if (!local && !global) { return all; }
    var out = [];
    for (var i = 0; i < all.length; i++) {
      var h = deviceHay(all[i]);
      if ((!local || h.indexOf(local) > -1) && (!global || h.indexOf(global) > -1)) { out.push(all[i]); }
    }
    return out;
  }

  function devicesMatrix(indexes) {
    var d = loadAppData();
    var matrix = [(d.cols || ['Poste', 'Utilisateur', 'OS', 'Version OS']).slice()];
    for (var i = 0; i < indexes.length; i++) { matrix.push(d.dev[indexes[i]]); }
    return matrix;
  }

  function mk(tag, cls, text) {
    var node = document.createElement(tag);
    if (cls) { node.className = cls; }
    if (text !== undefined && text !== null) { node.textContent = text; }
    return node;
  }

  function Pager(acc, host) {
    var self = this;
    this.acc = acc;
    this.all = appIndexes(acc.dataKey);
    this.view = this.all;
    this.page = 0;
    this.local = '';
    this.global = acc.pagerGlobal || '';
    this.timer = null;
    var cols = loadAppData().cols || ['Poste', 'Utilisateur', 'OS', 'Version OS'];

    host.innerHTML = '';
    var bar = mk('div', 'pg-bar');
    this.input = mk('input', 'search pg-q');
    this.input.type = 'search';
    this.input.placeholder = '🔍  Rechercher un poste, un utilisateur, un OS...';
    this.input.setAttribute('aria-label', 'Rechercher dans les postes de cette application');
    this.input.addEventListener('input', function () {
      if (self.timer) { clearTimeout(self.timer); }
      self.timer = setTimeout(function () { self.timer = null; self.local = norm(self.input.value); self.filter(); }, DEBOUNCE_MS);
    });
    this.info = mk('span', 'pg-info');
    this.info.setAttribute('role', 'status');
    this.info.setAttribute('aria-live', 'polite');
    var exportBtn = mk('button', 'btn btn-export', '⬇ Exporter la liste (CSV)');
    exportBtn.type = 'button';
    exportBtn.title = 'Exporte tous les postes de la liste, recherche comprise (pas seulement la page affichee)';
    exportBtn.addEventListener('click', function () { self.exportCsv(exportBtn); });
    bar.appendChild(this.input);
    bar.appendChild(exportBtn);
    bar.appendChild(this.info);

    var scroll = mk('div', 'tbl-scroll pg-scroll');
    var table = mk('table', 'tbl');
    var thead = mk('thead'), headRow = mk('tr');
    cols.forEach(function (c) { var th = mk('th', null, c); th.setAttribute('scope', 'col'); headRow.appendChild(th); });
    thead.appendChild(headRow);
    this.body = mk('tbody');
    table.appendChild(thead);
    table.appendChild(this.body);
    scroll.appendChild(table);

    var nav = mk('div', 'pg-nav');
    nav.setAttribute('role', 'navigation');
    nav.setAttribute('aria-label', 'Pagination des postes');
    function navBtn(label, title, fn) {
      var b = mk('button', 'pg-btn', label);
      b.type = 'button';
      b.title = title;
      b.setAttribute('aria-label', title);
      b.addEventListener('click', fn);
      nav.appendChild(b);
      return b;
    }
    this.bFirst  = navBtn('«', 'Premiere page', function () { self.go(0); });
    this.bPrev   = navBtn('‹', 'Page precedente', function () { self.go(self.page - 1); });
    this.pageLbl = mk('span', 'pg-page');
    nav.appendChild(this.pageLbl);
    this.bNext   = navBtn('›', 'Page suivante', function () { self.go(self.page + 1); });
    this.bLast   = navBtn('»', 'Derniere page', function () { self.go(self.pages() - 1); });

    host.appendChild(bar);
    host.appendChild(scroll);
    host.appendChild(nav);
    this.filter();
  }

  Pager.prototype.pages = function () { return Math.max(1, Math.ceil(this.view.length / PAGE_SIZE)); };

  Pager.prototype.setGlobal = function (query) {
    if (query === this.global) { return; }
    this.global = query;
    this.filter();
  };

  Pager.prototype.filter = function () {
    this.view = filterIndexes(this.all, this.local, this.global);
    this.page = 0;
    this.render();
  };

  Pager.prototype.go = function (page) {
    this.page = Math.min(Math.max(0, page), this.pages() - 1);
    this.render();
  };

  Pager.prototype.render = function () {
    var d = loadAppData(), total = this.view.length, pages = this.pages();
    if (this.page > pages - 1) { this.page = pages - 1; }
    var start = this.page * PAGE_SIZE, end = Math.min(total, start + PAGE_SIZE);
    var frag = document.createDocumentFragment();
    for (var i = start; i < end; i++) {
      var dev = d.dev[this.view[i]], tr = mk('tr');
      tr.appendChild(mk('td', 'cell-strong', dev[0]));
      tr.appendChild(mk('td', null, dev[1]));
      tr.appendChild(mk('td', null, dev[2]));
      tr.appendChild(mk('td', 'mono', dev[3]));
      frag.appendChild(tr);
    }
    if (total === 0) {
      var emptyRow = mk('tr'), cell = mk('td', 'pg-empty', 'Aucun poste ne correspond a la recherche.');
      cell.colSpan = 4;
      emptyRow.appendChild(cell);
      frag.appendChild(emptyRow);
    }
    this.body.innerHTML = '';
    this.body.appendChild(frag);

    /* Texte construit en noeuds (pas d'innerHTML) : aucune valeur collectee n'est interpretee */
    var filtered = (total !== this.all.length);
    this.info.innerHTML = '';
    if (total === 0) {
      this.info.appendChild(document.createTextNode('0 poste'));
    } else {
      this.info.appendChild(mk('b', null, (start + 1) + '–' + end));
      this.info.appendChild(document.createTextNode(' sur '));
      this.info.appendChild(mk('b', null, String(total)));
      this.info.appendChild(document.createTextNode(' poste(s)'));
    }
    if (filtered) {
      this.info.appendChild(document.createTextNode(' · filtre sur ' + this.all.length));
    } else {
      var announced = parseInt(this.acc.el.getAttribute('data-count'), 10);
      if (!isNaN(announced) && announced !== this.all.length) {
        this.info.appendChild(document.createTextNode(' · Intune en annonce ' + announced));
      }
    }
    this.pageLbl.textContent = 'Page ' + (total === 0 ? 0 : this.page + 1) + ' / ' + (total === 0 ? 0 : pages);
    this.bFirst.disabled = this.bPrev.disabled = (this.page === 0);
    this.bLast.disabled = this.bNext.disabled = (this.page >= pages - 1);
  };

  Pager.prototype.exportCsv = function (btn) {
    if (this.view.length === 0) { flash(btn, 'Rien a exporter'); return; }
    var filtered = (this.view.length !== this.all.length);
    var name = (this.acc.el.getAttribute('data-export-name') || 'application') + (filtered ? '-filtre' : '');
    downloadCsv(fileBase(name), devicesMatrix(this.view));
    flash(btn, '✓ ' + this.view.length + ' poste(s) exporte(s)');
  };

  function ensurePager(acc) {
    if (!acc || acc.dataKey === null || acc.pager) { return; }
    var host = acc.el.querySelector('.devpager');
    if (!host) { return; }
    acc.pager = new Pager(acc, host);
  }

  /* Liste a exporter pour une application : celle du tableau ouvert (filtres compris),
     sinon la liste complete restreinte par la recherche globale eventuelle. */
  function appExportIndexes(acc, key) {
    if (acc && acc.pager) { return acc.pager.view; }
    return filterIndexes(appIndexes(key), '', (acc && acc.pagerGlobal) || '');
  }

  function initDevicePagers() {
    allAccs.forEach(function (acc) {
      if (acc.dataKey === null) { return; }
      acc.el.addEventListener('toggle', function () { if (acc.el.open) { ensurePager(acc); } });
      if (acc.el.open) { ensurePager(acc); }
    });
  }

  /* ---------------------------------------------------------------
     Copie d'une commande PowerShell ciblant un poste.
     Ce rapport est un fichier statique : il n'a ni jeton ni acces reseau au
     tenant, il ne peut donc pas declencher l'action lui-meme. Voir la
     fonction Invoke-DeviceRemoteAction du script generateur.
     --------------------------------------------------------------- */
  window.copyDeviceAction = function (deviceId, action, deviceName, btn) {
    var cmd;
    if (action === 'Remediate') {
      cmd = "Invoke-DeviceRemoteAction -DeviceId '" + deviceId + "' -Action Remediate " +
            "-RemediationScriptId '<ID_DU_SCRIPT_DE_REMEDIATION>' -AccessToken $AccessToken  # poste : " + deviceName;
    } else {
      cmd = "Invoke-DeviceRemoteAction -DeviceId '" + deviceId + "' -Action " + action +
            " -AccessToken $AccessToken  # poste : " + deviceName;
    }
    var done = function () { flash(btn, 'Copie !'); };
    var fallback = function () {
      /* navigator.clipboard est indisponible sur un fichier local (file://) selon le
         navigateur : repli sur execCommand via une zone de texte temporaire. */
      var area = document.createElement('textarea');
      area.value = cmd;
      area.style.position = 'fixed';
      area.style.opacity = '0';
      document.body.appendChild(area);
      area.focus();
      area.select();
      try { document.execCommand('copy'); done(); }
      catch (e) { window.prompt('Copie automatique indisponible - copiez manuellement :', cmd); }
      document.body.removeChild(area);
    };
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(cmd).then(done, fallback);
    } else {
      fallback();
    }
  };

  /* ---------------------------------------------------------------
     Accessibilite : navigation clavier des onglets, raccourcis
     --------------------------------------------------------------- */
  function initTabKeyboard() {
    var list = byId('tabList');
    if (!list) { return; }
    list.addEventListener('keydown', function (e) {
      var keys = ['ArrowLeft', 'ArrowRight', 'Home', 'End'];
      if (keys.indexOf(e.key) === -1) { return; }
      var buttons = qsa('.tab-btn', list);
      if (buttons.length === 0) { return; }
      var current = buttons.indexOf(document.activeElement);
      if (current === -1) { current = 0; }
      var next = current;
      if (e.key === 'ArrowLeft')  { next = (current - 1 + buttons.length) % buttons.length; }
      if (e.key === 'ArrowRight') { next = (current + 1) % buttons.length; }
      if (e.key === 'Home')       { next = 0; }
      if (e.key === 'End')        { next = buttons.length - 1; }
      e.preventDefault();
      buttons[next].focus();
      window.showTab(buttons[next]);
    });
  }

  function initShortcuts() {
    document.addEventListener('keydown', function (e) {
      var search = byId('gSearch');
      if (!search) { return; }
      var active = document.activeElement;
      var typing = active && /^(input|textarea|select)$/i.test(active.tagName);
      if (((e.ctrlKey || e.metaKey) && (e.key === 'k' || e.key === 'K')) || (e.key === '/' && !typing)) {
        e.preventDefault();
        search.focus();
        search.select();
        return;
      }
      if (e.key === 'Escape' && active === search) { window.clearGlobalSearch(); }
    });
  }

  function init() {
    var prefix = document.body ? document.body.getAttribute('data-report-prefix') : null;
    if (prefix) { reportPrefix = prefix; }
    buildIndex();
    initDevicePagers();
    initTabKeyboard();
    initShortcuts();
    apply();
  }

  if (document.readyState === 'loading') { document.addEventListener('DOMContentLoaded', init); }
  else { init(); }
})();
'@

    # ---------- Assemblage final ----------
    # ---------- Onglet 4 : remédiation & santé des postes ----------
    $rmRows  = @(); $rmCrit = 0; $rmWarn = 0; $rmOk = 0; $rmScoreTxt = "&mdash;"
    $remAccordion = ""; $remRowsHtml = ""
    if ($RemediationData) {
        $rmRows = @($RemediationData.Rows)
        $rmCrit = $RemediationData.CritCount
        $rmWarn = $RemediationData.WarnCount
        $rmOk   = $RemediationData.OkCount
        if ($null -ne $RemediationData.AvgScore) { $rmScoreTxt = "$($RemediationData.AvgScore)" }
        $remAccordion = Build-RemediationAccordionHtml -Actions $RemediationData.Actions
        $remRowsHtml  = Build-RemediationRowsHtml -Rows $rmRows
    } else {
        $remAccordion = "<div class='empty'>Aucune donn&eacute;e de sant&eacute; collect&eacute;e.</div>"
    }
    $rmTotal = $rmRows.Count
    $tabPillRem = if (($rmCrit + $rmWarn) -gt 0) { "<span class='tab-pill'>$($rmCrit + $rmWarn)</span>" } else { "<span class='tab-pill gray'>0</span>" }

    # ---------- Onglets du rapport : assemblés à la carte ----------
    # Chaque onglet n'est produit que s'il a été autorisé dans l'outil. Le premier onglet
    # retenu devient l'onglet actif à l'ouverture, quel qu'il soit.
    # Les compteurs Total / Conformes / Non conformes figurent desormais dans le bandeau
    # KPI global, visible quel que soit l'onglet : les repeter ici n'apporterait rien et
    # eloignerait le graphique du haut de page. L'onglet s'ouvre donc directement sur la
    # vue d'ensemble, puis la repartition des versions d'OS, puis le detail par motif.
    $panelT1 = @"
  <div class="card">
    <h2><span class="h-dot violet"></span>Vue d'ensemble de la conformit&eacute;</h2>
    <p class="sub">R&eacute;partition des $TotalDevices appareils du p&eacute;rim&egrave;tre analys&eacute;.</p>
    <div class="duo">
      <div style="justify-self:center">
<svg viewBox="0 0 200 200" width="230" height="230" role="img" aria-label="R&eacute;partition de conformit&eacute;">
  <g transform="rotate(-90 100 100)">
    <circle cx="100" cy="100" r="78" fill="none" stroke="#edf0f6" stroke-width="24"/>
    <circle cx="100" cy="100" r="78" fill="none" stroke="#10b981" stroke-width="24" stroke-dasharray="$sOk $sC" stroke-dashoffset="0"/>
    <circle cx="100" cy="100" r="78" fill="none" stroke="#ef4444" stroke-width="24" stroke-dasharray="$sNc $sC" stroke-dashoffset="$oNc"/>
    <circle cx="100" cy="100" r="78" fill="none" stroke="#f59e0b" stroke-width="24" stroke-dasharray="$sGp $sC" stroke-dashoffset="$oGp"/>
  </g>
  <text x="100" y="97" text-anchor="middle" class="donut-num">$rate%</text>
  <text x="100" y="120" text-anchor="middle" class="donut-sub">conformes</text>
</svg>
      </div>
      <div class="legend">
        <div><div class="lg-row"><span class="lg-dot" style="background:#10b981"></span><span class="lg-nm">Conformes</span><span class="lg-ct t-green">$okC</span><span class="lg-pc">$pOk&nbsp;%</span></div><div class="bar"><i style="width:$pOk%;background:#10b981"></i></div></div>
        <div><div class="lg-row"><span class="lg-dot" style="background:#ef4444"></span><span class="lg-nm">Non-Compliant</span><span class="lg-ct t-red">$nc</span><span class="lg-pc">$pNc&nbsp;%</span></div><div class="bar"><i style="width:$pNc%;background:#ef4444"></i></div></div>
        <div><div class="lg-row"><span class="lg-dot" style="background:#f59e0b"></span><span class="lg-nm">In Grace Period</span><span class="lg-ct t-amber">$gp</span><span class="lg-pc">$pGp&nbsp;%</span></div><div class="bar"><i style="width:$pGp%;background:#f59e0b"></i></div></div>
      </div>
    </div>
  </div>

  $osMixHtml

  $complianceBody
"@

    $panelT2 = @"
  <div class="card">
    <h2><span class="h-dot blue"></span>Applications d&eacute;couvertes sur le parc</h2>
    <p class="sub">$appsTotal application(s) d&eacute;tect&eacute;e(s), tri&eacute;es par nombre de postes d&eacute;croissant. $appsDetailText D&eacute;pliez une application&nbsp;: ses postes s'affichent par tranches de 20, avec une recherche et un export CSV de la liste compl&egrave;te.</p>
    <div class="toolbar">
      <input id="appSearch" class="search" type="search" placeholder="&#128269;  Rechercher une application ou un &eacute;diteur..." oninput="filterApps()">
      <button type="button" class="btn" onclick="setAppsOpen(true)">Tout d&eacute;plier</button>
      <button type="button" class="btn" onclick="setAppsOpen(false)">Tout replier</button>
      <button type="button" class="btn btn-export" onclick="exportAllAccordions('appList','applications-decouvertes','Application',this)">&#11015; Exporter (CSV)</button>
      <span class="count-info" role="status" aria-live="polite"><b id="appCount">$appsTotal</b> / $appsTotal affich&eacute;e(s)</span>
    </div>
    <div id="appList">
$appsHtml
    </div>
  </div>
"@

    $panelT3 = @"
  <div class="kpis">
    <div class="kpi"><div class="kpi-ic ic-blue">&#128451;</div><div><div class="kpi-num">$invTotal</div><div class="kpi-lb">Applications Intune</div></div></div>
    <div class="kpi"><div class="kpi-ic ic-green">&#10004;</div><div><div class="kpi-num t-green">$($invBuild.CountOk)</div><div class="kpi-lb">&Agrave; jour</div></div></div>
    <div class="kpi"><div class="kpi-ic ic-amber">&#128270;</div><div><div class="kpi-num t-amber">$($invBuild.CountCheck)</div><div class="kpi-lb">&Agrave; v&eacute;rifier</div></div></div>
    <div class="kpi"><div class="kpi-ic ic-slate">&#10067;</div><div><div class="kpi-num">$($invBuild.CountUnknown)</div><div class="kpi-lb">Non v&eacute;rifiables</div></div></div>
  </div>

  <div class="card">
    <h2><span class="h-dot blue"></span>Inventaire des applications du tenant &amp; audit des versions</h2>
    <p class="sub">&laquo;&nbsp;Derni&egrave;re version publique&nbsp;&raquo;&nbsp;: meilleur effort via fichier de correspondance et/ou winget &mdash; r&eacute;sultat indicatif, aucune source universelle n'existant pour cette donn&eacute;e. Un mod&egrave;le CSV pr&eacute;-rempli des applications non r&eacute;solues est g&eacute;n&eacute;r&eacute; &agrave; c&ocirc;t&eacute; du rapport. Cliquez sur un en-t&ecirc;te de colonne pour trier.</p>
    <div class="toolbar">
      <input id="invSearch" class="search" type="search" placeholder="&#128269;  Rechercher une application..." oninput="filterInv()">
      <div class="chipbar" id="invChips">
        <button class="fchip active" data-f="all" onclick="setInvStatus(this)">Toutes</button>
        <button class="fchip" data-f="ok" onclick="setInvStatus(this)">&Agrave; jour</button>
        <button class="fchip" data-f="check" onclick="setInvStatus(this)">&Agrave; v&eacute;rifier</button>
        <button class="fchip" data-f="unknown" onclick="setInvStatus(this)">Non v&eacute;rifiables</button>
      </div>
      <button type="button" class="btn btn-export" onclick="exportTableById('invTable','inventaire-applications',true,this)">&#11015; Exporter la vue (CSV)</button>
      <span class="count-info" role="status" aria-live="polite"><b id="invCount">$invTotal</b> / $invTotal affich&eacute;e(s)</span>
    </div>
    <div class="tbl-scroll tall">
      <table class="tbl" id="invTable">
        <thead><tr>
          <th scope="col" class="sortable" data-col="0" data-export="Application" onclick="sortInv(this)">Application<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="1" data-export="Editeur" onclick="sortInv(this)">&Eacute;diteur<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="2" data-export="Version Intune" onclick="sortInv(this)">Version Intune<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="3" data-export="Derniere version publique" onclick="sortInv(this)">Derni&egrave;re version publique<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="4" data-export="Date de sortie" onclick="sortInv(this)">Date de sortie<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="5" data-export="Source" onclick="sortInv(this)">Source<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="6" data-export="Statut" onclick="sortInv(this)">Statut<span class="arr"></span></th>
        </tr></thead>
        <tbody id="invBody">
$($invBuild.RowsHtml)
        </tbody>
      </table>
    </div>
  </div>
"@

    $panelT4 = @"
  <div class="kpis">
    <div class="kpi"><div class="kpi-ic ic-slate" aria-hidden="true">&#128421;</div><div class="kpi-body"><div class="kpi-num">$rmTotal</div><div class="kpi-lb">Postes analys&eacute;s</div></div></div>
    <div class="kpi"><div class="kpi-ic ic-red" aria-hidden="true">&#128680;</div><div class="kpi-body"><div class="kpi-num t-red">$rmCrit</div><div class="kpi-lb">Postes en &eacute;tat critique</div></div></div>
    <div class="kpi"><div class="kpi-ic ic-amber" aria-hidden="true">&#128270;</div><div class="kpi-body"><div class="kpi-num t-amber">$rmWarn</div><div class="kpi-lb">&Agrave; surveiller</div></div></div>
    <div class="kpi"><div class="kpi-ic ic-green" aria-hidden="true">&#128200;</div><div class="kpi-body"><div class="kpi-num t-green">$rmScoreTxt</div><div class="kpi-lb">Score moyen Endpoint Analytics</div></div></div>
  </div>

  <div class="card">
    <h2><span class="h-dot red"></span>Actions de rem&eacute;diation recommand&eacute;es</h2>
    <p class="sub">Postes regroup&eacute;s par action &agrave; mener, de la plus critique &agrave; la plus b&eacute;nigne &mdash; cliquez sur une action pour afficher les postes concern&eacute;s. Chaque action s'exporte s&eacute;par&eacute;ment, ou toutes ensemble &agrave; plat (une ligne par poste, l'action en premi&egrave;re colonne). Seuils ajustables en t&ecirc;te de script.</p>
    <div class="toolbar">
      <button type="button" class="btn btn-export" onclick="exportAllAccordions('remActions','plan-de-remediation','Action a mener',this)">&#11015; Exporter le plan complet (CSV)</button>
    </div>
    <div id="remActions">
    $remAccordion
    </div>
  </div>

  <div class="card">
    <h2><span class="h-dot blue"></span>Sant&eacute; d&eacute;taill&eacute;e des postes</h2>
    <p class="sub">Espace disque, score Endpoint Analytics, d&eacute;marrage, &eacute;crans bleus, batterie (score et capacit&eacute; restante), s&eacute;curit&eacute; (BitLocker/Defender), conformit&eacute; Intune et profils de configuration, fiabilit&eacute; applicative, uptime estim&eacute; et derni&egrave;re synchronisation. Cliquez sur un en-t&ecirc;te de colonne pour trier. L'export reprend exactement la vue filtr&eacute;e &agrave; l'&eacute;cran.</p>
    <div class="toolbar">
      <input id="remSearch" class="search" type="search" placeholder="&#128269;  Rechercher un poste ou un utilisateur..." oninput="filterRem()">
      <div class="chipbar" id="remChips">
        <button class="fchip active" data-f="all" onclick="setRemStatus(this)">Tous</button>
        <button class="fchip" data-f="crit" onclick="setRemStatus(this)">Critiques</button>
        <button class="fchip" data-f="warn" onclick="setRemStatus(this)">&Agrave; surveiller</button>
        <button class="fchip" data-f="ok" onclick="setRemStatus(this)">OK</button>
      </div>
      <button type="button" class="btn btn-export" onclick="exportTableById('remTable','sante-des-postes',true,this)">&#11015; Exporter la vue (CSV)</button>
      <span class="count-info" role="status" aria-live="polite"><b id="remCount">$rmTotal</b> / $rmTotal affich&eacute;(s)</span>
    </div>
    <div class="tbl-scroll tall">
      <table class="tbl" id="remTable">
        <thead><tr>
          <th scope="col" class="sortable" data-col="0" data-export="Poste" onclick="sortRem(this)">Poste<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="1" data-export="Utilisateur" onclick="sortRem(this)">Utilisateur<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="2" data-export="Espace libre" onclick="sortRem(this)">Espace libre<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="3" data-export="Score Endpoint Analytics" onclick="sortRem(this)">Score<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="4" data-export="Demarrage" onclick="sortRem(this)">D&eacute;marrage<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="5" data-export="Type de disque" onclick="sortRem(this)">Disque<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="6" data-export="Ecrans bleus" onclick="sortRem(this)">BSOD<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="7" data-export="Batterie" onclick="sortRem(this)">Batterie<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="8" data-export="Securite" onclick="sortRem(this)">S&eacute;curit&eacute;<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="9" data-export="Conformite et profils" onclick="sortRem(this)">Conformit&eacute;<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="10" data-export="Fiabilite applicative" onclick="sortRem(this)">Fiabilit&eacute; appli<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="11" data-export="Uptime estime" onclick="sortRem(this)">Uptime<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="12" data-export="Derniere synchro" onclick="sortRem(this)">Synchro<span class="arr"></span></th>
          <th scope="col" class="sortable" data-col="13" data-export="Etat" onclick="sortRem(this)">&Eacute;tat<span class="arr"></span></th>
          <th scope="col" class="no-export">Actions</th>
        </tr></thead>
        <tbody id="remBody">
$remRowsHtml
        </tbody>
      </table>
    </div>
    <p class="sub" style="margin-top:10px">&#128203; Les boutons <b>Sync / Reboot / Rem&eacute;dier</b> copient une commande PowerShell pr&ecirc;te &agrave; coller dans une session d&eacute;j&agrave; connect&eacute;e &agrave; Microsoft Graph (voir <code>Invoke-DeviceRemoteAction</code>) &mdash; ce rapport &eacute;tant un fichier statique, il ne peut pas d&eacute;clencher d'action lui-m&ecirc;me.</p>
  </div>
"@

    $tabDefs = @()
    if ($IncludeCompliance)     { $tabDefs += @{ Id = 't1'; Label = "&#128737; Conformit&eacute; $tabPillCompliance"; Body = $panelT1 } }
    if ($IncludeDiscoveredApps) { $tabDefs += @{ Id = 't2'; Label = "&#128230; Applications d&eacute;couvertes <span class='tab-pill gray'>$appsTotal</span>"; Body = $panelT2 } }
    if ($IncludeInventory)      { $tabDefs += @{ Id = 't3'; Label = "&#128451; Inventaire &amp; Versions <span class='tab-pill gray'>$invTotal</span>"; Body = $panelT3 } }
    if ($IncludeRemediation)    { $tabDefs += @{ Id = 't4'; Label = "&#128736; Rem&eacute;diation $tabPillRem"; Body = $panelT4 } }

    $navHtml    = ""
    $panelsHtml = ""
    $isFirst    = $true
    foreach ($t in $tabDefs) {
        $act      = if ($isFirst) { " active" } else { "" }
        $selected = if ($isFirst) { "true" } else { "false" }
        $tabIndex = if ($isFirst) { "0" } else { "-1" }
        # Motif ARIA "tablist" complet : sans aria-selected ni gestion du focus, un lecteur
        # d'ecran annonce quatre boutons quelconques au lieu d'un jeu d'onglets navigable.
        $navHtml    += '<button type="button" role="tab" id="tab-' + $t.Id + '" aria-controls="' + $t.Id + '" aria-selected="' + $selected + '" tabindex="' + $tabIndex + '" class="tab-btn' + $act + '" data-tab="' + $t.Id + '" onclick="showTab(this)">' + $t.Label + '</button>' + "`n  "
        $panelsHtml += '<section id="' + $t.Id + '" role="tabpanel" aria-labelledby="tab-' + $t.Id + '" tabindex="0" class="tab-panel' + $act + '">' + "`n" + $t.Body + "`n</section>`n`n"
        $isFirst = $false
    }

    $html = @"
<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Rapport Intune &mdash; $ClientSafe</title>
<style>
$css
</style>
</head>
<body data-report-prefix="$ReportPrefix">

<a class="skip-link" href="#mainContent">Aller au contenu principal</a>

<header class="hdr">
  <div class="hdr-in">
    <div>
      <div class="hdr-brand">$CompanySafe</div>
      <div class="hdr-sub">Rapport de parc Intune &mdash; Conformit&eacute; &middot; Applications &middot; Inventaire</div>
      <div class="hdr-chips">
        <span class="hchip">&#128100; Client&nbsp;: $ClientSafe</span>
        $devChip
        $srcChip
        $anonChip
      </div>
    </div>
    $contactHtml
  </div>
</header>

$kpiBandHtml

<nav class="tabs" aria-label="Navigation du rapport">
  <div class="tabs-in" id="tabList" role="tablist" aria-label="Sections du rapport">
  $navHtml
  </div>
  <div class="gbar">
    <div class="gsearch-wrap">
      <span class="gsearch-ic" aria-hidden="true">&#128269;</span>
      <input id="gSearch" class="gsearch" type="search" autocomplete="off" spellcheck="false"
             placeholder="Rechercher un poste, un utilisateur, une application, une alerte..."
             aria-label="Recherche dans l'ensemble du rapport"
             oninput="globalSearchInput(this)">
      <button type="button" id="gClear" class="gsearch-x hidden" onclick="clearGlobalSearch()" aria-label="Effacer la recherche">&#10005;</button>
    </div>
    <span id="gHint" class="gsearch-hint" role="status" aria-live="polite"></span>
    <span class="gsearch-kbd" aria-hidden="true"><kbd>/</kbd> pour rechercher &middot; <kbd>&Eacute;chap</kbd> pour effacer</span>
  </div>
</nav>

<main id="mainContent">

$panelsHtml

</main>

<footer>
  Rapport g&eacute;n&eacute;r&eacute; le $GeneratedAt &mdash; document HTML autonome (consultable hors ligne, aucune d&eacute;pendance externe).$anonFooter
</footer>

<script>
$js
</script>
</body>
</html>
"@

    return $html
}

# ========================================
# GÉNÉRATION DU DASHBOARD (ORCHESTRATION)
# ========================================

function Get-CollectedItems {
    <# Lecture défensive du résultat d'une collecte (parallèle ou séquentielle) : une source
       absente, vide ou en erreur rend un tableau vide, jamais $null — le reste de la chaîne
       n'a ainsi jamais à distinguer « pas collecté » de « collecté vide ». #>
    param([hashtable]$Bag, [string]$Name)
    if ($Bag -and $Bag.ContainsKey($Name) -and $Bag[$Name]) { return @($Bag[$Name].Items) }
    return @()
}

function Generate-Dashboard {
    param([bool]$OpenAfterGeneration = $true)

    # ===== SOURCE DES DONNÉES (onglet "Source des données") =====
    #   API     : tout via Microsoft Graph (comportement historique)
    #   IMPORT  : aucune connexion au tenant, tout provient des fichiers désignés
    #   HYBRIDE : connexion à Graph, mais chaque jeu de données disposant d'un fichier
    #             est lu depuis ce fichier plutôt que via l'API
    $SourceMode = "API"
    if     ($rbSourceImport.Checked) { $SourceMode = "IMPORT" }
    elseif ($rbSourceHybrid.Checked) { $SourceMode = "HYBRIDE" }

    # ===== ONGLETS AUTORISÉS DANS LE RAPPORT =====
    # Un onglet décoché n'est ni produit... ni collecté : la génération est d'autant
    # plus rapide (aucun appel Graph inutile) que le rapport est ciblé.
    $IncCompliance = $chkTabCompliance.Checked
    $IncDiscovered = $chkTabDiscovered.Checked
    $IncInventory  = $chkTabInventory.Checked
    $IncRemediation = $chkTabRemediation.Checked
    if (-not ($IncCompliance -or $IncDiscovered -or $IncInventory -or $IncRemediation)) {
        Show-ErrorMessage "Aucun onglet n'est autorisé pour le rapport.`n`nCochez au moins un onglet dans « Options générales »."
        return
    }

    $ImpDevices    = $txtImpDevices.Text.Trim()
    $ImpCompliance = $txtImpCompliance.Text.Trim()
    $ImpDiscovered = $txtImpDiscovered.Text.Trim()
    $ImpAppDevices = $txtImpAppDevices.Text.Trim()
    $ImpInventory  = $txtImpInventory.Text.Trim()
    $ImpScores     = $txtImpScores.Text.Trim()
    $ImpPerf       = $txtImpPerf.Text.Trim()

    if ($SourceMode -ne "API") {
        $providedFiles = @($ImpDevices, $ImpCompliance, $ImpDiscovered, $ImpAppDevices, $ImpInventory, $ImpScores, $ImpPerf) |
                         Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        if ($providedFiles.Count -eq 0) {
            Show-ErrorMessage "Mode '$SourceMode' sélectionné, mais aucun fichier d'import n'est renseigné.`n`nIndiquez au moins un fichier dans l'onglet « Source des données », ou revenez au mode « API Microsoft Graph »."
            return
        }
        foreach ($f in $providedFiles) {
            if (-not (Test-Path $f)) { Show-ErrorMessage "Fichier d'import introuvable :`n$f"; return }
        }
    }

    # Quels jeux de données proviennent d'un fichier ?
    $DevicesFromFile    = ($SourceMode -ne "API") -and (-not [string]::IsNullOrWhiteSpace($ImpDevices))
    $ComplianceFromFile = ($SourceMode -ne "API") -and (-not [string]::IsNullOrWhiteSpace($ImpCompliance))
    $DiscoveredFromFile = ($SourceMode -ne "API") -and ((-not [string]::IsNullOrWhiteSpace($ImpDiscovered)) -or (-not [string]::IsNullOrWhiteSpace($ImpAppDevices)))
    $InventoryFromFile  = ($SourceMode -ne "API") -and (-not [string]::IsNullOrWhiteSpace($ImpInventory))
    $ScoresFromFile     = ($SourceMode -ne "API") -and (-not [string]::IsNullOrWhiteSpace($ImpScores))
    $PerfFromFile       = ($SourceMode -ne "API") -and (-not [string]::IsNullOrWhiteSpace($ImpPerf))
    $UseGraph           = ($SourceMode -ne "IMPORT")

    $selectedClient = $cmbClients.Text.Trim()
    $hasClient      = -not ($selectedClient -eq "-- Sélectionnez un client --" -or [string]::IsNullOrWhiteSpace($selectedClient))

    $TenantId = ""; $ClientId = ""; $ClientSecret = ""

    if ($UseGraph) {
        if (-not $hasClient) { Show-ErrorMessage "Veuillez sélectionner un client !"; return }
        $config = Get-ClientConfig -ClientName $selectedClient
        if ($null -eq $config) {
            Show-ErrorMessage "Impossible de charger la configuration du client !`n`nAssurez-vous que le fichier a été généré avec le même script (même clé AES)."
            return
        }
        $ClientName   = $config.ClientName
        $TenantId     = $config.TenantId
        $ClientId     = $config.ClientId
        $ClientSecret = $config.ClientSecret
    } else {
        # Mode 100 % import : aucune configuration client n'est nécessaire, le nom
        # ne sert qu'à titrer le rapport et à nommer le fichier de sortie.
        if ($hasClient) {
            $ClientName = $selectedClient
        } else {
            $ClientName = $txtImportClientName.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($ClientName)) { $ClientName = "Import" }
        }
    }

    $lblStatus.Text = "Génération du dashboard pour $ClientName..."
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $form.Refresh()

    $Progress = { param($msg) $lblStatus.Text = $msg; $form.Refresh(); Write-Log $msg }

    # Référentiel de pseudonymisation remis à zéro à chaque génération : un même appareil
    # doit porter le même pseudonyme dans l'onglet Conformité et dans l'onglet Remédiation,
    # mais jamais d'une génération/d'un client à l'autre.
    Reset-AnonymizationMaps

    Write-Log "================================================================" -Level INFO
    Write-Log "NOUVELLE EXECUTION - Client : $ClientName - OpenAfterGeneration : $OpenAfterGeneration" -Level INFO
    Write-Log "Source des donnees : $SourceMode (appareils=$(if($DevicesFromFile){'fichier'}else{'API'}) / conformite=$(if($ComplianceFromFile){'fichier'}else{'API'}) / appsDecouvertes=$(if($DiscoveredFromFile){'fichier'}else{'API'}) / inventaire=$(if($InventoryFromFile){'fichier'}else{'API'}))" -Level INFO
    if ($SourceMode -ne "API") {
        Write-Log "Fichiers : appareils='$ImpDevices' / conformite='$ImpCompliance' / appsDecouvertes='$ImpDiscovered' / postesParApp='$ImpAppDevices' / inventaire='$ImpInventory' / scores='$ImpScores' / demarrage='$ImpPerf'" -Level INFO
    }
    Write-Log "Onglets du rapport : Conformite=$IncCompliance / AppsDecouvertes=$IncDiscovered / Inventaire=$IncInventory / Remediation=$IncRemediation" -Level INFO
    Write-Log "Options : ExcludeVM=$($chkExcludeVM.Checked) / Anonymize=$($chkAnonymize.Checked) / AnonClient=$($chkAnonClient.Checked) / AnonApps=$($chkAnonApps.Checked) / TopNApps=$(if ($chkAllApps.Checked) { 'toutes' } else { $numTopApps.Value }) / UseWinget=$($chkUseWinget.Checked) / UseGitHubPkgs=$($chkUseGitHubPkgs.Checked) / JetonGitHub=$(if($txtGhToken.Text.Trim()){'oui'}else{'non'}) / OverrideCsv='$($txtOverrideCsv.Text)'" -Level INFO

    try {
        # ===== CONNEXION GRAPH (uniquement si au moins un jeu de données vient de l'API) =====
        # (Aucun module externe requis : authentification et appels via Invoke-RestMethod natif,
        #  rendu HTML sur mesure -> démarrage instantané, plus aucune installation de module.)
        $AccessToken = $null
        if ($UseGraph) {
            & $Progress "Connexion à Microsoft Graph..."

            # Garde-fou de session : tout le script passe par Invoke-RestMethod. Si une
            # FONCTION du même nom a été définie dans la session (harnais de test
            # dot-sourcé, module tiers, profil PowerShell), elle masque le cmdlet natif
            # et TOUS les appels partent dans le vide — y compris l'authentification, qui
            # « réussit » alors en renvoyant un jeton vide. Le symptôme apparaît ensuite
            # 40 lignes plus loin sous la forme d'une erreur de liaison de paramètre
            # incompréhensible. Autant le dire ici, franchement.
            $shadowed = @()
            foreach ($native in 'Invoke-RestMethod', 'Invoke-WebRequest') {
                $cmd = Get-Command -Name $native -ErrorAction SilentlyContinue
                if ($cmd -and $cmd.CommandType -ne 'Cmdlet') { $shadowed += $native }
            }
            if ($shadowed.Count -gt 0) {
                throw ("$($shadowed -join ' et ') masqué(s) dans cette session par une fonction du même nom : " +
                       "aucun appel réseau réel ne peut aboutir. Corrigez sans fermer la fenêtre avec :" +
                       [Environment]::NewLine + [Environment]::NewLine +
                       "    Remove-Item function:$($shadowed -join ', function:') -ErrorAction SilentlyContinue" +
                       [Environment]::NewLine + [Environment]::NewLine +
                       "(cause habituelle : un harnais de test chargé par dot-sourcing plus tôt dans cette session)")
            }

            $Scope   = "https://graph.microsoft.com/.default"
            $AuthUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
            $Body    = @{ client_id = $ClientId; scope = $Scope; client_secret = $ClientSecret; grant_type = "client_credentials" }
            $Connection  = Invoke-RestMethod -Method POST -Uri $AuthUrl -Body $Body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
            $AccessToken = $Connection.access_token

            # Un jeton vide n'est PAS une réussite : sans ce contrôle, l'erreur ne se
            # manifeste qu'au premier appel Graph, loin de sa cause.
            if ([string]::IsNullOrWhiteSpace($AccessToken)) {
                throw ("Microsoft Graph a répondu sans jeton d'accès (champ access_token vide). " +
                       "Vérifiez l'identifiant de tenant, l'ID d'application et le secret client de ce profil.")
            }
            Write-Log "Jeton d'accès Graph obtenu (validité $($Connection.expires_in) s)." -Level OK
        } else {
            & $Progress "Mode import : aucune connexion au tenant."
            Write-Log "Mode IMPORT : aucune connexion à Microsoft Graph n'est établie." -Level INFO
        }

        # ===== OPTIONS UI =====
        $ExcludeVM      = $chkExcludeVM.Checked
        $AnonymizeData  = $chkAnonymize.Checked
        $AnonymizeClient = $chkAnonClient.Checked
        $AnonymizeApps   = $chkAnonApps.Checked
        $AnonTerms       = if ($AnonymizeClient) { @(Split-AnonymizationTerms $txtAnonTerms.Text) } else { @() }
        $AnyAnonymization = ($AnonymizeData -or $AnonymizeClient -or $AnonymizeApps)
        # Nom affiché dans le rapport (en-tête, titre, préfixe des exports, nom du fichier).
        # $ClientName reste le vrai nom pour tout le reste (configuration, fichiers internes).
        $ReportClientName = if ($AnonymizeClient) { $AnonClientLabel } else { $ClientName }
        Save-AnonymizationTerms -ClientName $ClientName -Terms $txtAnonTerms.Text
        # [V2.3] 0 = toutes les applications (case cochée par défaut)
        $TopNDetailed   = if ($chkAllApps.Checked) { 0 } else { [int]$numTopApps.Value }
        $UseWinget      = $chkUseWinget.Checked
        $OverrideCsv    = $txtOverrideCsv.Text.Trim()
        $UseGitHubPkgs  = $chkUseGitHubPkgs.Checked
        # Pseudonymes remis à zéro à chaque génération (cohérents entre les onglets 1, 2 et 4).
        Reset-AnonymizationMaps
        $GitHubToken    = $txtGhToken.Text.Trim()
        # Le jeton saisi est conservé pour les exécutions suivantes (ou effacé s'il est vidé).
        Save-GitHubToken -Token $GitHubToken

        # ===== PAGE 1 : APPAREILS NON-COMPLIANT / GRACE PERIOD =====
        # NOTE : on récupère TOUS les managedDevices puis on filtre côté client sur
        # ComplianceState (un $filter OData combiné "... or ..." peut être rejeté en 400
        # selon le tenant). OPTIMISATION : $select limite la réponse aux 9 champs utiles
        # au lieu des ~60 propriétés de chaque appareil -> payload et temps de transfert
        # fortement réduits, et charge moindre côté service (moins de throttling).
        if (-not ($IncCompliance -or $IncRemediation)) {
            Write-Log "Onglets Conformité et Remédiation désactivés : aucune collecte d'appareils." -Level INFO
            $AllManagedDevicesRaw = @()
        } elseif ($DevicesFromFile) {
            & $Progress "Import des appareils depuis $([System.IO.Path]::GetFileName($ImpDevices))..."
            $AllManagedDevicesRaw = @(Import-ManagedDevicesFile -Path $ImpDevices)
        } elseif ($UseGraph) {
            & $Progress "Récupération de l'ensemble des appareils gérés..."
            $AllManagedDevicesUrl = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$select=id,deviceName,userPrincipalName,operatingSystem,osVersion,complianceState,lastSyncDateTime,manufacturer,model,freeStorageSpaceInBytes,totalStorageSpaceInBytes&`$top=999"
            $AllManagedDevicesRaw = Get-GraphPagedResults -Url $AllManagedDevicesUrl -AccessToken $AccessToken -ProgressCallback $Progress
            # CRITIQUE : fige lastSyncDateTime en ISO 8601 avant que quoi que ce soit d'autre
            # n'y touche (voir ConvertTo-SafeDateString) — sans quoi Invoke-RestMethod ayant
            # déjà converti ce champ en [datetime] .NET, tout cast [string] ultérieur (il y en
            # a dans tout le script) le reformate selon la culture régionale du poste
            # d'exécution, de façon ambiguë et imprévisible.
            Repair-GraphDateTimeFields -Items $AllManagedDevicesRaw -FieldNames @('lastSyncDateTime')
        } else {
            # Mode import sans fichier d'appareils : l'onglet 1 du rapport reste vide.
            Write-Log "Aucune source pour les appareils : l'onglet 1 du rapport sera vide." -Level WARN
            $AllManagedDevicesRaw = @()
        }

        # Déduplication par id dès la source (co-gestion / doublons éventuels),
        # pour des compteurs exacts partout (total, non conformes, grace period).
        $AllManagedDevicesRaw = @($AllManagedDevicesRaw | Sort-Object -Property id -Unique)

        # Le filtre VM (optionnel) s'applique à TOUT le périmètre : le taux de conformité
        # du donut est ainsi calculé sur une base cohérente.
        if ($ExcludeVM) {
            $AllManagedDevicesRaw = @($AllManagedDevicesRaw | Where-Object {
                -not (
                    ($_.manufacturer -match 'VMware|innotek|VirtualBox|Parallels|QEMU') -or
                    ($_.model        -match 'Virtual|VMware|VirtualBox|Hyper-V|Parallels|QEMU|KVM')
                )
            })
        }

        $TotalDevices = $AllManagedDevicesRaw.Count
        Write-Log "Périmètre d'analyse : $TotalDevices appareil(s) (ExcludeVM=$ExcludeVM)."

        # Normalisation en objets PascalCase (cohérent avec le reste du script)
        $NonCompliantAndGrace = $AllManagedDevicesRaw |
            Where-Object { $_.complianceState -eq "noncompliant" -or $_.complianceState -eq "inGracePeriod" } |
            ForEach-Object {
                [PSCustomObject]@{
                    Id                = $_.id
                    DeviceName        = $_.deviceName
                    UserPrincipalName = $_.userPrincipalName
                    OperatingSystem   = $_.operatingSystem
                    OSVersion         = $_.osVersion
                    ComplianceState   = $_.complianceState
                    LastSyncDateTime  = $_.lastSyncDateTime
                    # Motif éventuellement porté par le fichier d'import lui-même (colonne
                    # "Motif"/"Reason"), utilisé à défaut de fichier de détail de conformité.
                    ComplianceReason  = $_.complianceReason
                    ImportReasons     = @()
                }
            }

        $NonCompliantDevices = @($NonCompliantAndGrace | Where-Object { $_.ComplianceState -eq "noncompliant" })
        $GracePeriodDevices  = @($NonCompliantAndGrace | Where-Object { $_.ComplianceState -eq "inGracePeriod" })

        # ----- Origine du DÉTAIL des motifs de non-conformité -----
        # Règle : si la liste des appareils vient d'un fichier, le détail vient lui aussi
        # d'un fichier — sauf en mode HYBRIDE lorsque l'export contient de vrais
        # identifiants Intune (GUID) : l'API peut alors compléter le détail règle par règle.
        $UseImportBreakdown = (-not $UseGraph)
        if ($DevicesFromFile -or $ComplianceFromFile) {
            $UseImportBreakdown = $true
            if (-not $ComplianceFromFile -and $AccessToken -and (Test-DevicesHaveGraphIds -Devices $NonCompliantAndGrace)) {
                $UseImportBreakdown = $false
                Write-Log "Appareils importés porteurs d'identifiants Intune valides : le détail des règles en échec sera récupéré via l'API." -Level INFO
            }
        }

        if ($UseImportBreakdown) {
            $ImportedReasons = $null
            if ($ComplianceFromFile) {
                & $Progress "Import du détail de conformité..."
                $ImportedReasons = Import-ComplianceDetailFile -Path $ImpCompliance
            } else {
                Write-Log "Aucun fichier de détail de conformité : les appareils seront regroupés sous un motif générique (ou sur la colonne 'Motif' du fichier des appareils)." -Level WARN
            }
            # AVANT anonymisation : la correspondance se fait sur le nom réel du poste.
            Add-ImportedComplianceReasons -Devices $NonCompliantDevices -ReasonsByDevice $ImportedReasons
            Add-ImportedComplianceReasons -Devices $GracePeriodDevices  -ReasonsByDevice $ImportedReasons
        }

        # Anonymisation optionnelle de la Page 1 (l'ancien rendu ne masquait que la Page 2) :
        # mutation EN PLACE pour conserver l'Id (requis par l'analyse des motifs) et garantir
        # les mêmes libellés anonymisés dans les tableaux par motif.
        if ($AnonymizeData) {
            foreach ($d in (@($NonCompliantDevices) + @($GracePeriodDevices))) {
                $anon = Get-AnonymizedIdentity -RealName $d.DeviceName -RealUpn $d.UserPrincipalName
                $d.DeviceName        = $anon.Name
                $d.UserPrincipalName = $anon.Upn
            }
        }

        if ($UseImportBreakdown) {
            $NonCompliantBreakdown = Get-ComplianceBreakdownFromImport -Devices $NonCompliantDevices
            $GracePeriodBreakdown  = Get-ComplianceBreakdownFromImport -Devices $GracePeriodDevices
        } else {
            $NonCompliantBreakdown = Get-ComplianceBreakdown -Devices $NonCompliantDevices -AccessToken $AccessToken -ProgressCallback $Progress
            $GracePeriodBreakdown  = Get-ComplianceBreakdown -Devices $GracePeriodDevices  -AccessToken $AccessToken -ProgressCallback $Progress
        }

        # Courte pause de respiration : la Page 1 vient d'envoyer un grand nombre d'appels
        # $batch sur le même service Intune (managedDevices/compliancePolicyStates). On laisse
        # retomber un éventuel throttling avant d'attaquer detectedApps (Page 2), qui est connu
        # pour être plus sensible au débit (cf. limite réelle constatée bien en-deçà des 600
        # req/min documentées sur cet endpoint).
        # Les pauses anti-throttling n'ont de sens que si la section suivante interroge l'API :
        # une génération 100 % hors ligne est ainsi immédiate.
        if ($UseGraph -and $IncDiscovered -and $IncCompliance -and -not $DiscoveredFromFile -and -not $UseImportBreakdown) {
            & $Progress "Pause avant la Page 2 (limitation du débit Intune constatée sur ce tenant)..."
            Start-ResponsiveSleep -Seconds 25 -Message "Pause avant la Page 2 (limitation du débit Intune)"
        }

        # ===== PAGE 2 : DISCOVERED APPS =====
        if (-not $IncDiscovered) {
            Write-Log "Onglet Applications découvertes désactivé : aucune collecte." -Level INFO
            $DiscoveredAppsReport = [PSCustomObject]@{ AllApps = @(); DevicesByApp = @{}; DetailedIds = @(); TruncatedApps = @{} }
        } elseif ($DiscoveredFromFile) {
            & $Progress "Import des applications découvertes..."
            $DiscoveredAppsReport = Get-DiscoveredAppsReportFromImport -AppsPath $ImpDiscovered -DevicesPath $ImpAppDevices `
                -TopNDetailed $TopNDetailed -AnonymizeData $AnonymizeData -ProgressCallback $Progress
        } elseif ($UseGraph) {
            & $Progress "Récupération du rapport Discovered Apps..."
            $DiscoveredAppsReport = Get-DiscoveredAppsReport -AccessToken $AccessToken -TopNDetailed $TopNDetailed -AnonymizeData $AnonymizeData -ProgressCallback $Progress
        } else {
            # Mode import sans fichier d'applications découvertes : l'onglet reste vide.
            Write-Log "Aucune source pour les applications découvertes : l'onglet 2 du rapport sera vide." -Level WARN
            $DiscoveredAppsReport = [PSCustomObject]@{ AllApps = @(); DevicesByApp = @{}; DetailedIds = @(); TruncatedApps = @{} }
        }

        if ($UseGraph -and $IncInventory -and $IncDiscovered -and -not $InventoryFromFile -and -not $DiscoveredFromFile) {
            & $Progress "Pause courte avant la Page 3 (limitation du débit Intune)..."
            Start-ResponsiveSleep -Seconds 15 -Message "Pause avant la Page 3 (limitation du débit Intune)"
        }

        # ===== PAGE 3 : INVENTAIRE & VERSIONS =====
        # NOTE : la recherche des dernières versions publiques (CSV manuel, API éditeurs,
        # winget, Chocolatey) s'applique de la même façon, que l'inventaire vienne de
        # l'API ou d'un fichier importé.
        if (-not $IncInventory) {
            Write-Log "Onglet Inventaire & versions désactivé : aucune collecte ni recherche de versions." -Level INFO
            $TenantAppInventory = @()
        } elseif ($InventoryFromFile) {
            & $Progress "Import de l'inventaire des applications..."
            $ImportedInventoryApps = Import-MobileAppsFile -Path $ImpInventory
            $TenantAppInventory = Get-TenantAppInventory -PreloadedApps $ImportedInventoryApps -OverrideCsvPath $OverrideCsv `
                -UseOnlineSources $UseWinget -UseGitHubPkgs $UseGitHubPkgs -GitHubToken $GitHubToken -ProgressCallback $Progress
        } elseif ($UseGraph) {
            & $Progress "Récupération de l'inventaire des applications du tenant..."
            $TenantAppInventory = Get-TenantAppInventory -AccessToken $AccessToken -OverrideCsvPath $OverrideCsv `
                -UseOnlineSources $UseWinget -UseGitHubPkgs $UseGitHubPkgs -GitHubToken $GitHubToken -ProgressCallback $Progress
        } else {
            Write-Log "Aucune source pour l'inventaire applicatif : l'onglet 3 du rapport sera vide." -Level WARN
            $TenantAppInventory = @()
        }
        # Conservé pour la fenêtre "Versions manuelles" : elle propose alors la liste
        # réelle du tenant (y compris les applications masquées, pour pouvoir les réafficher),
        # pré-remplie avec ce qui a été résolu en ligne.
        $script:LastInventory = $TenantAppInventory

        # Applications décochées dans la fenêtre "Versions manuelles" : retirées du rapport
        # seulement, jamais du référentiel ni de l'éditeur.
        if ($TenantAppInventory.Count -gt 0) {
            $hiddenApps = Get-ExcludedAppNames
            if ($hiddenApps.Count -gt 0) {
                $before = $TenantAppInventory.Count
                $TenantAppInventory = @($TenantAppInventory | Where-Object { -not $hiddenApps.ContainsKey("$($_.DisplayName)".Trim().ToLower()) })
                $removed = $before - $TenantAppInventory.Count
                if ($removed -gt 0) { Write-Log "$removed application(s) masquée(s) retirée(s) de l'onglet Inventaire du rapport." -Level INFO }
            }
        }

        if ($UseGraph -and $IncRemediation -and -not $ScoresFromFile -and -not $PerfFromFile -and ($IncInventory -or $IncDiscovered -or $IncCompliance)) {
            & $Progress "Pause avant la Page 4 (limitation du débit Intune)..."
            Start-ResponsiveSleep -Seconds 10 -Message "Pause avant la Page 4 (limitation du débit Intune)"
        }

        # ===== PAGE 4 : REMÉDIATION & SANTÉ DES POSTES =====
        # Réutilise l'ensemble des appareils déjà collectés pour la Page 1 (tous états de
        # conformité confondus, dédupliqués et filtrés VM) : aucun appel supplémentaire
        # n'est nécessaire pour l'espace disque et l'inactivité. Chaque vérification
        # supplémentaire (BitLocker, Defender, mises à jour, fiabilité applicative, uptime)
        # est indépendamment activable/désactivable dans l'onglet "Remédiation avancée" :
        # décochée, elle n'est ni collectée ni affichée.
        # Déclarées hors du bloc conditionnel : le bandeau d'indicateurs les consulte même
        # lorsque l'onglet Remédiation est désactivé, et une variable jamais affectée
        # produirait ici un indicateur silencieusement faux plutôt qu'une absence assumée.
        $RemediationData = $null
        $BitLockerStates = @()
        $HasBitLockerData = $false
        if (-not $IncRemediation) {
            Write-Log "Onglet Remédiation désactivé : aucune collecte de scores ni de performances." -Level INFO
        } else {
            # Synchronise les cases à cocher de l'onglet "Remédiation avancée" avec le
            # référentiel lu par Build-RemediationData.
            foreach ($k in @($script:ChkRemChecks.Keys)) {
                $script:RemediationChecks[$k] = $script:ChkRemChecks[$k].Checked
            }
            $RC = $script:RemediationChecks

            # ----- Chemins d'import éventuels (modes Import / Mixte) -----
            $ImpBitLockerPath = if ($txtImpBitLocker) { $txtImpBitLocker.Text.Trim() } else { "" }
            $ImpDefenderPath  = if ($txtImpDefender)  { $txtImpDefender.Text.Trim() }  else { "" }
            $ImpAppReliabPath = if ($txtImpAppReliab) { $txtImpAppReliab.Text.Trim() } else { "" }

            # ================================================================
            # COLLECTE DES DONNÉES DE SANTÉ
            # ----------------------------------------------------------------
            # Une source part à l'API si, et seulement si : la vérification est cochée, le
            # mode de collecte autorise l'API, et aucun fichier d'import ne la couvre déjà.
            # Toutes celles qui remplissent ces conditions sont de simples lectures de
            # listes, indépendantes les unes des autres : on les lance DE FRONT plutôt qu'à
            # la file. C'est le principal gain de temps de la génération, sans changer d'un
            # iota les données produites.
            # ================================================================
            $useApiFor = @{
                Scores      = ($UseGraph -and -not $ScoresFromFile)
                Performance = ($UseGraph -and -not $PerfFromFile)
                BitLocker   = ($UseGraph -and $RC.BitLocker -and (($SourceMode -eq "API") -or [string]::IsNullOrWhiteSpace($ImpBitLockerPath)))
                Battery     = ($UseGraph -and $RC.BatteryDetail)
                AppReliab   = ($UseGraph -and $RC.AppReliability -and (($SourceMode -eq "API") -or [string]::IsNullOrWhiteSpace($ImpAppReliabPath)))
                Startup     = ($UseGraph -and $RC.Uptime)
            }

            $graphBase    = "https://graph.microsoft.com/beta/deviceManagement"
            $parallelJobs = @{}
            if ($useApiFor.Scores)      { $parallelJobs["Scores Endpoint Analytics"]      = "$graphBase/userExperienceAnalyticsDeviceScores?`$top=999" }
            if ($useApiFor.Performance) { $parallelJobs["Performances de démarrage"]      = "$graphBase/userExperienceAnalyticsDevicePerformance?`$top=999" }
            if ($useApiFor.BitLocker)   { $parallelJobs["État BitLocker"]                 = "$graphBase/managedDeviceEncryptionStates?`$top=999" }
            if ($useApiFor.Battery)     { $parallelJobs["Santé des batteries"]            = "$graphBase/userExperienceAnalyticsBatteryHealthDevicePerformance?`$top=999" }
            if ($useApiFor.AppReliab)   { $parallelJobs["Fiabilité applicative (postes)"] = "$graphBase/userExperienceAnalyticsAppHealthDevicePerformanceDetails?`$top=999" }
            if ($useApiFor.AppReliab)   { $parallelJobs["Fiabilité applicative (apps)"]   = "$graphBase/userExperienceAnalyticsAppHealthApplicationPerformance?`$top=999" }
            if ($useApiFor.Startup)     { $parallelJobs["Historique de démarrage"]        = "$graphBase/userExperienceAnalyticsDeviceStartupHistory?`$top=999" }

            $collected = @{}
            if ($parallelJobs.Count -gt 0) {
                if ($script:UseParallelCollection -and $parallelJobs.Count -gt 1) {
                    & $Progress "Collecte des données de santé ($($parallelJobs.Count) jeux de données en parallèle)..."
                    $collected = Invoke-ParallelGraphCollections -Jobs $parallelJobs -AccessToken $AccessToken `
                                    -MaxConcurrency $script:MaxParallelCollections -ProgressCallback $Progress
                } else {
                    foreach ($jobName in @($parallelJobs.Keys)) {
                        try {
                            & $Progress "Récupération : $jobName..."
                            $items = @(Get-GraphPagedResults -Url $parallelJobs[$jobName] -AccessToken $AccessToken -ProgressCallback $Progress -Label $jobName)
                            $collected[$jobName] = [PSCustomObject]@{ Items = $items.ToArray(); Error = $null }
                        } catch {
                            Write-Log "Collecte '$jobName' impossible : $($_.Exception.Message)" -Level WARN
                            $collected[$jobName] = [PSCustomObject]@{ Items = @(); Error = $_.Exception.Message }
                        }
                    }
                }
            }

            # ----- Scores Endpoint Analytics -----
            if ($ScoresFromFile) {
                & $Progress "Import des scores de santé (Endpoint Analytics)..."
                $DeviceScores = Import-DeviceScoresFile -Path $ImpScores
            } else {
                $DeviceScores = Get-CollectedItems -Bag $collected -Name "Scores Endpoint Analytics"
            }

            # ----- Performances de démarrage -----
            if ($PerfFromFile) {
                & $Progress "Import des performances de démarrage..."
                $StartupPerf = Import-StartupPerfFile -Path $ImpPerf
            } else {
                $StartupPerf = Get-CollectedItems -Bag $collected -Name "Performances de démarrage"
            }

            # ----- BitLocker -----
            if (-not $RC.BitLocker) {
                $BitLockerStates = @()
            } elseif (($SourceMode -ne "API") -and -not [string]::IsNullOrWhiteSpace($ImpBitLockerPath)) {
                & $Progress "Import de l'état BitLocker..."
                $BitLockerStates = Import-BitLockerFile -Path $ImpBitLockerPath
            } else {
                $BitLockerStates = Get-CollectedItems -Bag $collected -Name "État BitLocker"
            }

            # ----- Santé fine des batteries -----
            $BatteryHealth = if ($RC.BatteryDetail) { Get-CollectedItems -Bag $collected -Name "Santé des batteries" } else { @() }
            if ($RC.BatteryDetail -and $UseGraph -and $BatteryHealth.Count -eq 0) {
                Write-Log "Aucune donnée de santé de batterie : Endpoint Analytics inactive, ou parc sans appareil portable." -Level INFO
            }

            # ----- Fiabilité applicative -----
            if (-not $RC.AppReliability) {
                $AppReliabilityData = [PSCustomObject]@{ ByApp = @(); ByDevice = @() }
            } elseif (($SourceMode -ne "API") -and -not [string]::IsNullOrWhiteSpace($ImpAppReliabPath)) {
                & $Progress "Import de la fiabilité applicative..."
                $imported = Import-AppReliabilityFile -Path $ImpAppReliabPath
                # Reconstruit une entrée par plantage compté (Build-RemediationData compte les
                # occurrences dans ByDevice) : plus simple et plus lisible qu'une expansion en une
                # seule expression, et sans dépendre de l'ordre d'évaluation d'un pipeline imbriqué.
                $expanded = New-Object System.Collections.Generic.List[psobject]
                foreach ($row in $imported) {
                    $n = [Math]::Max(1, [int]$row.appCrashCount)
                    for ($i = 0; $i -lt $n; $i++) {
                        $expanded.Add([PSCustomObject]@{
                            deviceDisplayName = $row.deviceName
                            eventType         = "crash"
                            appDisplayName    = $row.appDisplayName
                        })
                    }
                }
                $AppReliabilityData = [PSCustomObject]@{ ByApp = @(); ByDevice = $expanded.ToArray() }
            } else {
                $byDevice = Get-CollectedItems -Bag $collected -Name "Fiabilité applicative (postes)"
                # Les dates renvoyées par Invoke-RestMethod sont déjà des [datetime] .NET :
                # on les fige en ISO 8601 avant que le moindre cast [string] ne les reformate
                # selon la culture régionale du poste d'exécution (voir ConvertTo-SafeDateString).
                Repair-GraphDateTimeFields -Items $byDevice -FieldNames @('eventDateTime')
                $AppReliabilityData = [PSCustomObject]@{
                    ByApp    = (Get-CollectedItems -Bag $collected -Name "Fiabilité applicative (apps)")
                    ByDevice = $byDevice
                }
            }

            # ----- Defender : pas de liste globale côté Graph, une requête par poste -----
            # Reste hors de la collecte parallèle : c'est un appel $batch, déjà parallélisé
            # par lots de 20 côté service.
            if (-not $RC.Defender) {
                $DefenderStates = @{}
            } elseif (($SourceMode -ne "API") -and -not [string]::IsNullOrWhiteSpace($ImpDefenderPath)) {
                & $Progress "Import de l'état Defender..."
                $DefenderStates = Import-DefenderFile -Path $ImpDefenderPath
            } elseif ($UseGraph) {
                $deviceIdsForDefender = @($AllManagedDevicesRaw | Where-Object { $_.id } | Select-Object -ExpandProperty id)
                $DefenderStates = Get-DefenderProtectionStates -DeviceIds $deviceIdsForDefender -AccessToken $AccessToken -ProgressCallback $Progress
            } else {
                $DefenderStates = @{}
            }

            # ----- Profils de configuration en erreur / conflit (API seule) -----
            $ConfigProfileErrors = if ($RC.ConfigProfile -and $UseGraph) {
                Get-ConfigurationProfileErrors -AccessToken $AccessToken -ProgressCallback $Progress -MaxProfiles $script:MaxConfigProfilesAnalyzed
            } else { @() }

            # ----- Mises à jour Windows (API seule : rapport asynchrone Intune) -----
            $WindowsUpdateFailures = if ($RC.WindowsUpdate -and $UseGraph) {
                Get-WindowsUpdateFailures -AccessToken $AccessToken -ProgressCallback $Progress
            } else { @() }
            if ($RC.WindowsUpdate -and $WindowsUpdateFailures.Count -eq 0 -and $UseGraph) {
                Write-Log "Aucune donnée de mise à jour Windows exploitable (rapport vide ou Update Rings non configurés)." -Level INFO
            }

            # ----- Uptime (historique de démarrage) -----
            $StartupHistory = if ($RC.Uptime) { Get-CollectedItems -Bag $collected -Name "Historique de démarrage" } else { @() }
            if ($StartupHistory.Count -gt 0) {
                Repair-GraphDateTimeFields -Items $StartupHistory -FieldNames @('startupDateTime', 'eventDateTime', 'lastBootUpTime')
            }

            # ----- Motifs de non-conformité, rapatriés au niveau du poste -----
            # L'onglet Conformité raisonne « par motif » ; l'onglet Remédiation raisonne
            # « par poste ». On inverse ici l'index déjà calculé pour l'onglet 1 : un
            # technicien voit le motif exact à côté des autres défauts du poste, sans avoir
            # à recouper deux onglets. Aucun appel supplémentaire n'est nécessaire.
            $ComplianceReasonsByDevice = @{}
            foreach ($breakdown in @($NonCompliantBreakdown, $GracePeriodBreakdown)) {
                if (-not $breakdown) { continue }
                foreach ($reasonEntry in $breakdown.GetEnumerator()) {
                    foreach ($dev in @($reasonEntry.Value.Devices)) {
                        # Les postes de l'onglet 1 sont déjà pseudonymisés : la jointure avec
                        # l'onglet 4 se fait sur le nom RÉEL, sinon aucun motif n'est retrouvé.
                        $key = Get-NormalizedHeader (Get-RealDeviceName ([string](Get-PropCI $dev @('DeviceName','deviceName'))))
                        if (-not $key) { continue }
                        if (-not $ComplianceReasonsByDevice.ContainsKey($key)) {
                            $ComplianceReasonsByDevice[$key] = [string]$reasonEntry.Key
                        } elseif ($ComplianceReasonsByDevice[$key] -notlike "*$($reasonEntry.Key)*") {
                            $ComplianceReasonsByDevice[$key] += " ; " + [string]$reasonEntry.Key
                        }
                    }
                }
            }

            & $Progress "Analyse de la santé des postes..."
            $RemediationData = Build-RemediationData -Devices $AllManagedDevicesRaw -Scores $DeviceScores -Performance $StartupPerf `
                -BitLocker $BitLockerStates -Defender $DefenderStates -AppReliability $AppReliabilityData `
                -StartupHistory $StartupHistory -ConfigProfileErrors $ConfigProfileErrors `
                -BatteryHealth $BatteryHealth -UpdateFailures $WindowsUpdateFailures `
                -ComplianceReasons $ComplianceReasonsByDevice -AnonymizeData $AnonymizeData -AnonymizeApps $AnonymizeApps
            $rmLevel = if ($RemediationData.CritCount -gt 0) { "WARN" } else { "OK" }
            Write-Log ("Rémédiation : $($RemediationData.Rows.Count) poste(s) analysé(s) - Critiques=$($RemediationData.CritCount) / À surveiller=$($RemediationData.WarnCount) / OK=$($RemediationData.OkCount)") -Level $rmLevel
            if ($DeviceScores.Count -eq 0 -and $StartupPerf.Count -eq 0 -and $UseGraph -and -not $ScoresFromFile -and -not $PerfFromFile) {
                Write-Log "Aucun score Endpoint Analytics ni performance de démarrage disponible : l'onglet Remédiation n'affichera que l'espace disque et l'inactivité." -Level WARN
            }
        }

        # ===== INDICATEURS DE PARC (bandeau de tête du rapport) =====
        # Un seul parcours du parc pour tous les indicateurs : conformité, inactivité,
        # répartition des versions d'OS, chiffrement. Le seuil d'inactivité affiché est
        # celui des seuils de remédiation, pour qu'un même poste ne soit pas « inactif »
        # dans le bandeau et « à surveiller » dans le tableau selon deux règles différentes.
        $HasBitLockerData = (@($BitLockerStates).Count -gt 0)
        $KpiSummary = $null
        if ($TotalDevices -gt 0) {
            & $Progress "Calcul des indicateurs de parc..."
            $KpiSummary = Get-FleetKpiSummary -Devices $AllManagedDevicesRaw `
                -StaleDays $RemediationThresholds.StaleDaysWarning `
                -StaleCriticalDays $RemediationThresholds.StaleDaysCritical `
                -BitLockerStates $BitLockerStates -RemediationData $RemediationData
            Write-Log ("Indicateurs de parc : conformité $($KpiSummary.ComplianceRate) % - $($KpiSummary.StaleCount) inactif(s) depuis plus de $($KpiSummary.StaleDays) j - Windows 11 : $($KpiSummary.Windows11) / Windows 10 : $($KpiSummary.Windows10)") -Level INFO
            if ($KpiSummary.EndOfSupport -gt 0) {
                Write-Log "$($KpiSummary.EndOfSupport) poste(s) sous Windows 10 ne reçoivent plus de correctifs de sécurité (fin de support du $($script:Windows10EndOfSupport.ToString('dd/MM/yyyy')))." -Level WARN
            }
        }

        # ===== GÉNÉRATION HTML =====
        & $Progress "Génération du fichier HTML..."
        $Timestamp        = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
        $AnonymizedSuffix = if ($AnyAnonymization) { "_ANONYMIZED" } else { "" }

        # ----- Export des données collectées (fichiers CSV réimportables) -----
        # Permet de rejouer plus tard la même génération hors ligne, d'archiver l'état du
        # tenant à une date donnée, ou de transmettre les données sans donner d'accès Graph.
        $ExportedFolder = $null
        if ($chkExportRaw.Checked -and $SourceMode -ne "IMPORT") {
            try {
                & $Progress "Export des données collectées (format réimportable)..."
                $ExportedFolder = Export-RawCollectedData -ClientName $ClientName -Timestamp $Timestamp `
                    -Devices $AllManagedDevicesRaw -NonCompliantBreakdown $NonCompliantBreakdown `
                    -GracePeriodBreakdown $GracePeriodBreakdown -DiscoveredApps $DiscoveredAppsReport `
                    -Inventory $TenantAppInventory -Scores $DeviceScores -Performance $StartupPerf
            } catch {
                # Un échec d'export ne doit jamais empêcher la production du rapport.
                Write-Log "Export des données collectées impossible : $($_.Exception.Message)" -Level WARN
            }
        }

        $SourceLabel = switch ($SourceMode) {
            "IMPORT"  { "Import de fichiers (hors ligne)" }
            "HYBRIDE" { "API Microsoft Graph + fichiers import&eacute;s" }
            default   { "API Microsoft Graph" }
        }
        $ReportFileName   = "Intune-Dashboard-Conformite-Apps_${ReportClientName}${AnonymizedSuffix}_${Timestamp}.html"

        $ShowContactInfo = $chkShowContact.Checked
        $CompanyName     = if ([string]::IsNullOrWhiteSpace($txtCompanyName.Text))   { $DefaultCompanyName }   else { $txtCompanyName.Text }
        $ContactPerson   = if ([string]::IsNullOrWhiteSpace($txtContactPerson.Text)) { $DefaultContactPerson } else { $txtContactPerson.Text }
        $ContactEmail    = if ([string]::IsNullOrWhiteSpace($txtContactEmail.Text))  { $DefaultContactEmail }  else { $txtContactEmail.Text }
        $ContactPhone    = if ([string]::IsNullOrWhiteSpace($txtContactPhone.Text))  { $DefaultContactPhone }  else { $txtContactPhone.Text }

        # ----- Modèle CSV de correspondance : pré-rempli avec les noms EXACTS du tenant -----
        # L'utilisateur complète LatestVersion (et ReleaseDate s'il le souhaite) pour les
        # applications qui l'intéressent, puis pointe l'option "Fichier CSV" vers ce fichier :
        # la correspondance est garantie puisque les noms proviennent du tenant lui-même.
        $TemplateCsvPath = $null
        $Unresolved = @($TenantAppInventory | Where-Object { $_.LatestPublicVersion -eq "Non trouvé" })
        if ($Unresolved.Count -gt 0) {
            $TemplateCsvPath = Join-Path $OutputFolder ("Modele-Versions_{0}_{1}.csv" -f $ClientName, $Timestamp)
            $Unresolved | Sort-Object DisplayName | Select-Object `
                @{ n = 'AppName';           e = { $_.DisplayName } },
                @{ n = 'LatestVersion';     e = { '' } },
                @{ n = 'ReleaseDate';       e = { '' } },
                @{ n = 'InfoVersionIntune'; e = { $_.CurrentVersion } },
                @{ n = 'InfoEditeur';       e = { $_.Publisher } } |
                Export-Csv -Path $TemplateCsvPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
            Write-Log "Modèle de correspondance versions généré : $TemplateCsvPath ($($Unresolved.Count) application(s) à compléter)." -Level OK
        }

        # ----- Anonymisation des applications (après le modèle CSV, qui garde les vrais noms) -----
        $InventoryForReport  = $TenantAppInventory
        $DiscoveredForReport = $DiscoveredAppsReport
        if ($AnonymizeApps) {
            & $Progress "Anonymisation des noms d'applications..."
            if (@($TenantAppInventory).Count -gt 0) { $InventoryForReport = @(ConvertTo-AnonymizedInventory -Inventory $TenantAppInventory) }
            $DiscoveredForReport = ConvertTo-AnonymizedDiscoveredApps -Report $DiscoveredAppsReport
            Write-Log "Anonymisation : $($script:AppAnonList.Count) application(s) et $($script:ProcAnonList.Count) exécutable(s) renommé(s)." -Level INFO
        }

        $ReportPath  = Join-Path $OutputFolder $ReportFileName
        $HtmlContent = Build-DashboardHtml `
            -ClientName $ReportClientName -CompanyName $CompanyName `
            -ContactPerson $ContactPerson -ContactEmail $ContactEmail -ContactPhone $ContactPhone `
            -ShowContactInfo $ShowContactInfo -AnonymizeData $AnyAnonymization `
            -TotalDevices $TotalDevices `
            -NonCompliantDevices $NonCompliantDevices -GracePeriodDevices $GracePeriodDevices `
            -NonCompliantBreakdown $NonCompliantBreakdown -GracePeriodBreakdown $GracePeriodBreakdown `
            -DiscoveredApps $DiscoveredForReport -TenantAppInventory $InventoryForReport `
            -TopNDetailed $TopNDetailed -DataSourceLabel $SourceLabel `
            -IncludeCompliance $IncCompliance -IncludeDiscoveredApps $IncDiscovered -IncludeInventory $IncInventory `
            -IncludeRemediation $IncRemediation -RemediationData $RemediationData `
            -KpiSummary $KpiSummary -HasBitLockerData $HasBitLockerData

        # ----- Masquage du nom du client et des termes associés, sur le HTML final -----
        $AnonMapPath = $null
        if ($AnonymizeClient) {
            & $Progress "Masquage du nom du client et des termes sensibles..."
            # Le nom par défaut du mode import ("Import") n'est pas un nom de client : le
            # masquer abîmerait le libellé de source du rapport.
            $clientTerms = if ($ClientName -and $ClientName -ne "Import") { @($ClientName) } else { @() }
            $HtmlContent = Protect-HtmlSensitiveTerms -Html $HtmlContent -Terms $AnonTerms `
                -WholeWordTerms $clientTerms -Replacement $AnonClientLabel
        }
        if ($AnyAnonymization) {
            try {
                $AnonMapPath = Join-Path $OutputFolder ("CONFIDENTIEL_Correspondance-Anonymisation_{0}_{1}.csv" -f $ClientName, $Timestamp)
                $AnonMapPath = Export-AnonymizationMap -Path $AnonMapPath -ClientName $(if ($AnonymizeClient) { $ClientName } else { "" }) `
                    -ClientLabel $AnonClientLabel -Terms $AnonTerms
            } catch {
                Write-Log "Table de correspondance d'anonymisation impossible à écrire : $($_.Exception.Message)" -Level WARN
                $AnonMapPath = $null
            }
        }

        [System.IO.File]::WriteAllText($ReportPath, $HtmlContent, (New-Object System.Text.UTF8Encoding($true)))
        Write-Log "Fichier HTML écrit : $ReportPath ($([Math]::Round((Get-Item $ReportPath).Length / 1KB)) Ko)" -Level OK

        if ($OpenAfterGeneration) { Start-Process $ReportPath }

        $lblStatus.Text = "Dashboard généré avec succès."; $lblStatus.ForeColor = [System.Drawing.Color]::Green
        Write-Log "SUCCES - Fichier : $OutputFolder\$ReportFileName" -Level OK
        Write-Log "Résumé : Total=$TotalDevices / NonCompliant=$($NonCompliantDevices.Count) / GracePeriod=$($GracePeriodDevices.Count) / DiscoveredApps=$($DiscoveredAppsReport.AllApps.Count) / TenantApps=$($TenantAppInventory.Count)" -Level OK
        $message = "Dashboard généré avec succès !`n`nFichier : $OutputFolder\$ReportFileName"
        if ($OpenAfterGeneration) { $message += "`n`nLe rapport s'ouvre automatiquement dans votre navigateur." }
        else                      { $message += "`n`nLe rapport est disponible dans le dossier de sortie." }
        if ($TemplateCsvPath) { $message += "`n`nModèle de versions à compléter (à repointer via l'option 'Fichier CSV') :`n$TemplateCsvPath" }
        if ($ExportedFolder)  { $message += "`n`nDonnées collectées exportées (réimportables via l'onglet « Source des données ») :`n$ExportedFolder" }
        if ($AnonMapPath) {
            $message += "`n`nTable de correspondance d'anonymisation (CONFIDENTIELLE - ne jamais la transmettre avec le rapport) :`n$AnonMapPath"
        }
        if ($AnyAnonymization -and ($TemplateCsvPath -or $ExportedFolder)) {
            $message += "`n`nAttention : le modèle de versions et l'export des données brutes contiennent les données réelles."
        }
        Show-InfoMessage $message

    } catch {
        $lblStatus.Text = "Erreur lors de la génération"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
        Write-Log "ECHEC - $($_.Exception.Message)" -Level ERROR
        Write-Log "Ligne : $($_.InvocationInfo.ScriptLineNumber) / Trace : $($_.ScriptStackTrace)" -Level ERROR
        Show-ErrorMessage "Erreur lors de la génération du dashboard :`n`n$($_.Exception.Message)`n`n(Détail complet consigné dans $LogFile)"
    }
}

# ========================================
# RÉFÉRENTIEL DES VERSIONS MANUELLES
# ========================================
#
# PRINCIPE : un fichier unique et persistant sert de mémoire des versions que les
# catalogues en ligne ne savent pas fournir (pilotes constructeurs, applications métier).
# Il est chargé AUTOMATIQUEMENT à chaque génération, sans rien avoir à sélectionner, et
# se remplit via la fenêtre d'édition intégrée plutôt qu'à la main dans Excel.
#
# Il reste un simple CSV (séparateur ';') : il peut être versionné, partagé entre
# techniciens ou modifié dans Excel si vous préférez.

function Get-ExcludedAppNames {
    <#
        Applications à ne pas faire figurer dans l'onglet "Inventaire & versions" du
        rapport. Renvoie une table de hachage indexée sur le nom en minuscules.
    #>
    param([string]$Path = $ExcludedAppsFile)
    $set = @{}
    if (-not (Test-Path $Path)) { return $set }
    foreach ($delim in @(';', ',')) {
        try {
            $data = @(Import-Csv -Path $Path -Delimiter $delim -ErrorAction Stop)
            if ($data.Count -gt 0 -and ($data[0].PSObject.Properties.Name -contains 'AppName')) {
                foreach ($row in $data) {
                    $k = "$($row.AppName)".Trim().ToLower()
                    if ($k) { $set[$k] = "$($row.AppName)".Trim() }
                }
                break
            }
        } catch { }
    }
    if ($set.Count -gt 0) { Write-Log "Applications masquées du rapport : $($set.Count) entrée(s) chargée(s)." -Level INFO }
    return $set
}

function Save-ExcludedAppNames {
    <# Réécrit la liste des applications masquées (vide -> le fichier est supprimé). #>
    param([string[]]$Names, [string]$Path = $ExcludedAppsFile)
    $clean = @($Names | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    if ($clean.Count -eq 0) {
        if (Test-Path $Path) { Remove-Item $Path -Force }
        Write-Log "Aucune application masquée : liste vidée." -Level INFO
        return 0
    }
    $clean | ForEach-Object { [PSCustomObject]@{ AppName = $_ } } |
        Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Log "Applications masquées du rapport : $($clean.Count) enregistrée(s) dans $Path" -Level OK
    return $clean.Count
}

function Get-ManualVersionEntries {
    <#
        Lit le référentiel manuel et renvoie une table de hachage indexée sur le nom
        d'application en minuscules. Séparateur ';' d'abord, ',' en secours.
    #>
    param([string]$Path = $ManualVersionsFile)

    $entries = @{}
    if (-not $Path -or -not (Test-Path $Path)) { return $entries }

    foreach ($delim in @(';', ',')) {
        try {
            $rows = @(Import-Csv -Path $Path -Delimiter $delim)
            if ($rows.Count -eq 0) { continue }
            # Un fichier lu avec le mauvais séparateur ne produit qu'une seule colonne :
            # on le détecte et on tente l'autre séparateur.
            if (-not ($rows[0].PSObject.Properties.Name -contains 'AppName')) { continue }
            foreach ($row in $rows) {
                if ([string]::IsNullOrWhiteSpace($row.AppName)) { continue }
                $entries[$row.AppName.Trim().ToLower()] = [PSCustomObject]@{
                    AppName       = $row.AppName.Trim()
                    LatestVersion = "$($row.LatestVersion)".Trim()
                    ReleaseDate   = "$($row.ReleaseDate)".Trim()
                    Publisher     = "$($row.InfoEditeur)".Trim()
                }
            }
            if ($entries.Count -gt 0) { break }
        } catch {
            Write-Log "Lecture du référentiel manuel impossible (séparateur '$delim') : $($_.Exception.Message)" -Level WARN
        }
    }
    return $entries
}

function Save-ManualVersionEntries {
    <#
        Écrit le référentiel manuel. Seules les lignes réellement renseignées sont
        conservées : le fichier ne se remplit pas de lignes vides au fil des générations.
    #>
    param([Parameter(Mandatory = $true)]$Entries, [string]$Path = $ManualVersionsFile)

    $folder = Split-Path $Path -Parent
    if ($folder -and -not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }

    $rows = @()
    foreach ($e in $Entries) {
        if ([string]::IsNullOrWhiteSpace($e.AppName)) { continue }
        if ([string]::IsNullOrWhiteSpace($e.LatestVersion)) { continue }
        $rows += [PSCustomObject]@{
            AppName       = $e.AppName
            LatestVersion = $e.LatestVersion
            ReleaseDate   = $e.ReleaseDate
            InfoEditeur   = $e.Publisher
        }
    }

    if ($rows.Count -eq 0) {
        if (Test-Path $Path) { Remove-Item $Path -Force }
        Write-Log "Référentiel manuel vidé (aucune version renseignée)." -Level INFO
        return 0
    }

    $rows | Sort-Object AppName | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Log "Référentiel manuel enregistré : $($rows.Count) version(s) dans $Path" -Level OK
    return $rows.Count
}

function Get-VersionEditorRows {
    <#
        Construit la liste d'applications proposée dans l'éditeur, par ordre de préférence :
          1. le fichier d'inventaire importé, dès qu'il est désigné dans l'onglet
             "Source des données" (la liste correspond ainsi toujours au fichier
             actuellement sélectionné, sans qu'aucun rapport n'ait à être généré) ;
          2. l'inventaire de la dernière génération / analyse de cette session ;
          3. à défaut, le modèle CSV le plus récent produit à côté d'un rapport ;
          4. à défaut, le référentiel manuel existant seul.
        Dans tous les cas, les versions déjà saisies manuellement sont pré-remplies, et ce
        qui a été trouvé en ligne lors d'une génération/analyse est rappelé pour information.
    #>
    $manual = Get-ManualVersionEntries
    $rows   = @()
    $origin = ""

    # Index de ce qui a été résolu lors de la dernière génération ou analyse de versions :
    # sert à enrichir la liste quelle que soit sa provenance.
    $found = @{}
    if ($script:LastInventory -and $script:LastInventory.Count -gt 0) {
        foreach ($item in $script:LastInventory) {
            $k = "$($item.DisplayName)".Trim().ToLower()
            if ($k) { $found[$k] = $item }
        }
    }

    # ----- Source 1 : fichier d'inventaire importé -----
    $importPath = ""
    if ($txtImpInventory -and $rbSourceApi -and (-not $rbSourceApi.Checked)) {
        $importPath = $txtImpInventory.Text.Trim()
    }
    if (-not [string]::IsNullOrWhiteSpace($importPath) -and (Test-Path $importPath)) {
        try {
            $imported = @(Import-MobileAppsFile -Path $importPath)
            if ($imported.Count -gt 0) {
                $origin = "fichier importé $([System.IO.Path]::GetFileName($importPath)) ($($imported.Count) applications)"
                foreach ($app in $imported) {
                    $key = "$($app.displayName)".Trim().ToLower()
                    $m   = if ($manual.ContainsKey($key)) { $manual[$key] } else { $null }
                    $f   = if ($found.ContainsKey($key))  { $found[$key] }  else { $null }
                    $cur = if ([string]::IsNullOrWhiteSpace($app.displayVersion)) { "N/A" } else { "$($app.displayVersion)" }
                    $rows += [PSCustomObject]@{
                        AppName        = $app.displayName
                        Publisher      = $app.publisher
                        CurrentVersion = $cur
                        FoundVersion   = if ($f) { $f.LatestPublicVersion } else { "Non recherché" }
                        FoundSource    = if ($f) { $f.Source }             else { "" }
                        LatestVersion  = if ($m) { $m.LatestVersion } else { "" }
                        ReleaseDate    = if ($m) { $m.ReleaseDate }   else { "" }
                    }
                }
            }
        } catch {
            Write-Log "Lecture du fichier d'inventaire importé impossible pour l'éditeur : $($_.Exception.Message)" -Level WARN
        }
    }

    # ----- Source 2 : inventaire de la dernière génération / analyse -----
    if ($rows.Count -eq 0 -and $script:LastInventory -and $script:LastInventory.Count -gt 0) {
        $origin = "dernière génération ($($script:LastInventory.Count) applications)"
        foreach ($item in $script:LastInventory) {
            $key = "$($item.DisplayName)".Trim().ToLower()
            $m   = if ($manual.ContainsKey($key)) { $manual[$key] } else { $null }
            $rows += [PSCustomObject]@{
                AppName        = $item.DisplayName
                Publisher      = $item.Publisher
                CurrentVersion = $item.CurrentVersion
                FoundVersion   = $item.LatestPublicVersion
                FoundSource    = $item.Source
                LatestVersion  = if ($m) { $m.LatestVersion } else { "" }
                ReleaseDate    = if ($m) { $m.ReleaseDate }   else { "" }
            }
        }
    }

    # ----- Source 3 : dernier modèle CSV produit à côté d'un rapport -----
    if ($rows.Count -eq 0) {
        $template = Get-ChildItem -Path $OutputFolder -Filter "Modele-Versions_*.csv" -ErrorAction SilentlyContinue |
                    Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($template) {
            $origin = "modèle $($template.Name)"
            foreach ($delim in @(';', ',')) {
                $data = @()
                try { $data = @(Import-Csv -Path $template.FullName -Delimiter $delim) } catch { }
                if ($data.Count -gt 0 -and ($data[0].PSObject.Properties.Name -contains 'AppName')) {
                    foreach ($row in $data) {
                        $key = "$($row.AppName)".Trim().ToLower()
                        $m   = if ($manual.ContainsKey($key)) { $manual[$key] } else { $null }
                        $rows += [PSCustomObject]@{
                            AppName        = $row.AppName
                            Publisher      = "$($row.InfoEditeur)"
                            CurrentVersion = "$($row.InfoVersionIntune)"
                            FoundVersion   = "Non trouvé"
                            FoundSource    = ""
                            LatestVersion  = if ($m) { $m.LatestVersion } else { "$($row.LatestVersion)" }
                            ReleaseDate    = if ($m) { $m.ReleaseDate }   else { "$($row.ReleaseDate)" }
                        }
                    }
                    break
                }
            }
        }
    }

    # Entrées du référentiel manuel absentes de la source retenue (saisies lors d'une
    # session précédente, sur un autre tenant ou un autre fichier). Elles sont conservées
    # dans la liste — sinon un simple enregistrement les effacerait — mais marquées
    # FromReferential afin que l'éditeur les masque par défaut : la liste affichée
    # correspond alors exactement au fichier / au tenant analysé.
    $known = @{}
    foreach ($r in $rows) { $known["$($r.AppName)".Trim().ToLower()] = $true }
    foreach ($k in $manual.Keys) {
        if ($known.ContainsKey($k)) { continue }
        $m = $manual[$k]
        $rows += [PSCustomObject]@{
            AppName         = $m.AppName
            Publisher       = $m.Publisher
            CurrentVersion  = "hors liste"
            FoundVersion    = ""
            FoundSource     = "référentiel manuel"
            LatestVersion   = $m.LatestVersion
            ReleaseDate     = $m.ReleaseDate
            FromReferential = $true
        }
    }
    if (-not $origin -and $rows.Count -gt 0) { $origin = "référentiel manuel" }

    return [PSCustomObject]@{ Rows = @($rows | Sort-Object AppName); Origin = $origin }
}

function Show-VersionEditor {
    <#
        Fenêtre d'édition des versions manuelles : un tableau où seules les colonnes
        "Version publique" et "Date de sortie" sont modifiables, avec recherche et
        filtre sur les applications non résolues. À l'enregistrement, seules les lignes
        renseignées sont écrites dans le référentiel.
    #>
    $data = Get-VersionEditorRows
    if ($data.Rows.Count -eq 0) {
        Show-InfoMessage ("Aucune application à afficher pour le moment.`n`n" +
                          "Deux façons d'obtenir la liste :`n" +
                          "  - désignez un fichier d'inventaire dans l'onglet « Source des données » " +
                          "(les applications de ce fichier sont alors proposées immédiatement, sans générer de rapport) ;`n" +
                          "  - ou lancez une génération : la liste du tenant sera proposée ici, " +
                          "pré-remplie avec ce qui a été trouvé en ligne.`n`n" +
                          "Référentiel : $ManualVersionsFile")
        return
    }

    # Les entrées héritées du référentiel (absentes de la source) sont comptées à part :
    # elles ne doivent pas gonfler le total annoncé ni polluer la liste par défaut.
    $orphanRows  = @($data.Rows | Where-Object { $_.FromReferential })
    $orphanCount = $orphanRows.Count
    $mainCount   = $data.Rows.Count - $orphanCount

    $editor = New-Object System.Windows.Forms.Form
    $editor.Text            = "Versions manuelles - $mainCount application(s)"
    $editor.ClientSize      = New-Object System.Drawing.Size(1080, 660)
    $editor.StartPosition   = "CenterParent"
    $editor.BackColor       = ConvertTo-UIColor $Theme.FormBack
    $editor.Font            = New-Object System.Drawing.Font($Theme.FontFamily, 9)
    $editor.MinimizeBox     = $false
    $editor.Add_Load({ Enable-ModernWindowCorners -TargetForm $editor })

    # --- En-tête assorti à la fenêtre principale ---
    $hdr = New-Object System.Windows.Forms.Panel
    $hdr.Location = New-Object System.Drawing.Point(0, 0)
    $hdr.Size     = New-Object System.Drawing.Size(1080, 74)
    $hdr.Add_Paint({
        param($sender, $e)
        $rect  = $sender.ClientRectangle
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, (ConvertTo-UIColor $Theme.HeaderColor1), (ConvertTo-UIColor $Theme.HeaderColor2), 25)
        $e.Graphics.FillRectangle($brush, $rect); $brush.Dispose()
    })
    $editor.Controls.Add($hdr)

    $hTitle = New-Object System.Windows.Forms.Label
    $hTitle.Location  = New-Object System.Drawing.Point(28, 14)
    $hTitle.Size      = New-Object System.Drawing.Size(1020, 28)
    $hTitle.Text      = "Saisie des versions non trouvées automatiquement"
    $hTitle.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 13, [System.Drawing.FontStyle]::Bold)
    $hTitle.ForeColor = [System.Drawing.Color]::White
    $hTitle.BackColor = [System.Drawing.Color]::Transparent
    $hdr.Controls.Add($hTitle)

    $hSub = New-Object System.Windows.Forms.Label
    $hSub.Location  = New-Object System.Drawing.Point(30, 44)
    $hSub.Size      = New-Object System.Drawing.Size(1020, 22)
    $hSub.Text      = "Source : $($data.Origin)   -   Référentiel : $ManualVersionsFile"
    $hSub.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.5)
    $hSub.ForeColor = [System.Drawing.Color]::FromArgb(220, 255, 255, 255)
    $hSub.BackColor = [System.Drawing.Color]::Transparent
    $hdr.Controls.Add($hSub)

    # --- Barre d'outils ---
    $lblSearch = New-Object System.Windows.Forms.Label
    $lblSearch.Location  = New-Object System.Drawing.Point(24, 92)
    $lblSearch.Size      = New-Object System.Drawing.Size(70, 24)
    $lblSearch.Text      = "Rechercher"
    $lblSearch.ForeColor = ConvertTo-UIColor $Theme.TextMuted
    $editor.Controls.Add($lblSearch)

    $txtSearch = New-Object System.Windows.Forms.TextBox
    $txtSearch.Location    = New-Object System.Drawing.Point(100, 89)
    $txtSearch.Size        = New-Object System.Drawing.Size(300, 26)
    $txtSearch.BorderStyle = "FixedSingle"
    $editor.Controls.Add($txtSearch)

    $chkOnlyMissing = New-Object System.Windows.Forms.CheckBox
    $chkOnlyMissing.Location  = New-Object System.Drawing.Point(420, 90)
    $chkOnlyMissing.Size      = New-Object System.Drawing.Size(360, 24)
    $chkOnlyMissing.Text      = "Uniquement les applications sans version trouvée"
    $chkOnlyMissing.Checked   = $true
    $chkOnlyMissing.ForeColor = ConvertTo-UIColor $Theme.TextMain
    $editor.Controls.Add($chkOnlyMissing)

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Location  = New-Object System.Drawing.Point(800, 92)
    $lblCount.Size      = New-Object System.Drawing.Size(256, 24)
    $lblCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lblCount.ForeColor = ConvertTo-UIColor $Theme.TextMuted
    $editor.Controls.Add($lblCount)

    $chkShowOrphans = New-Object System.Windows.Forms.CheckBox
    $chkShowOrphans.Location  = New-Object System.Drawing.Point(24, 118)
    $chkShowOrphans.Size      = New-Object System.Drawing.Size(1030, 24)
    $chkShowOrphans.Text      = "Afficher aussi $orphanCount entrée(s) déjà présente(s) dans le référentiel mais absente(s) de cette liste (pour les corriger ou les supprimer)"
    $chkShowOrphans.Checked   = $false
    $chkShowOrphans.Visible   = ($orphanCount -gt 0)
    $chkShowOrphans.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.5)
    $chkShowOrphans.ForeColor = ConvertTo-UIColor $Theme.TextMuted
    $editor.Controls.Add($chkShowOrphans)

    # --- Tableau ---
    $gridTop    = if ($orphanCount -gt 0) { 148 } else { 124 }
    $gridHeight = 586 - $gridTop

    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location              = New-Object System.Drawing.Point(24, $gridTop)
    $grid.Size                  = New-Object System.Drawing.Size(1032, $gridHeight)
    $grid.BackgroundColor       = ConvertTo-UIColor $Theme.CardBack
    $grid.BorderStyle           = "None"
    $grid.AllowUserToAddRows    = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AllowUserToResizeRows = $false
    $grid.RowHeadersVisible     = $false
    $grid.SelectionMode         = [System.Windows.Forms.DataGridViewSelectionMode]::CellSelect
    $grid.EditMode              = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
    $grid.GridColor             = ConvertTo-UIColor $Theme.CardBorder
    $grid.EnableHeadersVisualStyles = $false
    $grid.ColumnHeadersDefaultCellStyle.BackColor = ConvertTo-UIColor $Theme.FormBack
    $grid.ColumnHeadersDefaultCellStyle.ForeColor = ConvertTo-UIColor $Theme.TextMuted
    $grid.ColumnHeadersDefaultCellStyle.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.5, [System.Drawing.FontStyle]::Bold)
    $grid.ColumnHeadersHeight   = 34
    $grid.RowTemplate.Height    = 26
    $editor.Controls.Add($grid)

    # Colonne à cocher : décocher une application la retire de l'onglet
    # "Inventaire & versions" du rapport, sans supprimer sa version saisie.
    $colShow = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colShow.Name       = "Show"
    $colShow.HeaderText = "Afficher"
    $colShow.Width      = 62
    $colShow.ReadOnly   = $false
    $colShow.ToolTipText = "Décochez pour ne pas faire figurer cette application dans le rapport."
    [void]$grid.Columns.Add($colShow)

    $cols = @(
        @{ Name = "AppName";        Header = "Application";        Width = 258; ReadOnly = $true  },
        @{ Name = "Publisher";      Header = "Éditeur";            Width = 130; ReadOnly = $true  },
        @{ Name = "CurrentVersion"; Header = "Version Intune";     Width = 120; ReadOnly = $true  },
        @{ Name = "FoundVersion";   Header = "Trouvée en ligne";   Width = 130; ReadOnly = $true  },
        @{ Name = "LatestVersion";  Header = "VERSION À SAISIR";   Width = 150; ReadOnly = $false },
        @{ Name = "ReleaseDate";    Header = "Date (facultatif)";  Width = 120; ReadOnly = $false }
    )
    foreach ($col in $cols) {
        $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $c.Name       = $col.Name
        $c.HeaderText = $col.Header
        $c.Width      = $col.Width
        $c.ReadOnly   = $col.ReadOnly
        if ($col.ReadOnly) {
            $c.DefaultCellStyle.BackColor = ConvertTo-UIColor $Theme.FormBack
            $c.DefaultCellStyle.ForeColor = ConvertTo-UIColor $Theme.TextMuted
        } else {
            $c.DefaultCellStyle.BackColor = [System.Drawing.Color]::White
            $c.DefaultCellStyle.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
        }
        [void]$grid.Columns.Add($c)
    }

    # Les saisies sont conservées hors du tableau : filtrer ne fait donc jamais
    # perdre ce qui a été tapé, même si la ligne disparaît de l'affichage.
    $excluded = Get-ExcludedAppNames
    $script:EditorValues = @{}
    foreach ($r in $data.Rows) {
        $script:EditorValues["$($r.AppName)"] = [PSCustomObject]@{
            LatestVersion = $r.LatestVersion
            ReleaseDate   = $r.ReleaseDate
            Show          = (-not $excluded.ContainsKey("$($r.AppName)".Trim().ToLower()))
        }
    }

    $refresh = {
        $q    = $txtSearch.Text.Trim().ToLower()
        $only = $chkOnlyMissing.Checked
        $grid.Rows.Clear()
        $shown = 0
        foreach ($r in $data.Rows) {
            # Entrées héritées du référentiel : masquées sauf demande explicite. Elles
            # restent dans $data.Rows, donc l'enregistrement ne les perd jamais.
            if ($r.FromReferential -and -not $chkShowOrphans.Checked) { continue }
            if ($only) {
                $found = "$($r.FoundVersion)"
                # "Non recherché" = application issue d'un import non encore analysé :
                # elle doit rester visible dans le filtre "sans version trouvée".
                if ($found -and $found -notin @("Non trouvé", "Non recherché", "")) { continue }
            }
            if ($q -and ("$($r.AppName) $($r.Publisher)".ToLower() -notlike "*$q*")) { continue }
            $v = $script:EditorValues["$($r.AppName)"]
            [void]$grid.Rows.Add(@($v.Show, $r.AppName, $r.Publisher, $r.CurrentVersion, $r.FoundVersion, $v.LatestVersion, $v.ReleaseDate))
            $shown++
        }
        $filled = @($script:EditorValues.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.LatestVersion) }).Count
        $hidden = @($script:EditorValues.Values | Where-Object { -not $_.Show }).Count
        $lblCount.Text = "$shown ligne(s) - $filled saisie(s) - $hidden masquée(s)"
    }

    # Sans cela, une case à cocher ne déclenche CellValueChanged qu'après avoir quitté
    # la cellule : le compteur et la mémorisation paraîtraient décalés d'un clic.
    $grid.Add_CurrentCellDirtyStateChanged({
        param($sender, $e)
        if ($sender.IsCurrentCellDirty -and $sender.CurrentCell -is [System.Windows.Forms.DataGridViewCheckBoxCell]) {
            $sender.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit)
        }
    })

    $grid.Add_CellValueChanged({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        $name = "$($sender.Rows[$e.RowIndex].Cells['AppName'].Value)"
        if (-not $name -or -not $script:EditorValues.ContainsKey($name)) { return }
        $script:EditorValues[$name].LatestVersion = "$($sender.Rows[$e.RowIndex].Cells['LatestVersion'].Value)".Trim()
        $script:EditorValues[$name].ReleaseDate   = "$($sender.Rows[$e.RowIndex].Cells['ReleaseDate'].Value)".Trim()
        $script:EditorValues[$name].Show          = [bool]$sender.Rows[$e.RowIndex].Cells['Show'].Value
        $filled = @($script:EditorValues.Values | Where-Object { -not [string]::IsNullOrWhiteSpace($_.LatestVersion) }).Count
        $hidden = @($script:EditorValues.Values | Where-Object { -not $_.Show }).Count
        $lblCount.Text = ($lblCount.Text -replace '\d+ saisie\(s\) - \d+ masquée\(s\)', "$filled saisie(s) - $hidden masquée(s)")
    })

    $txtSearch.Add_TextChanged($refresh)
    $chkOnlyMissing.Add_CheckedChanged($refresh)
    $chkShowOrphans.Add_CheckedChanged($refresh)

    # --- Boutons ---
    $btnSave = New-ThemedButton -Text "Enregistrer" -X 828 -Y 604 -W 228 -H 40 -BackHex $Theme.Success -HoverHex $Theme.SuccessHover
    $btnSave.Add_Click({
        # Valide l'édition en cours : sans cela, la dernière cellule saisie serait perdue.
        try { $grid.EndEdit() | Out-Null } catch { }
        $toSave = @()
        foreach ($r in $data.Rows) {
            $v = $script:EditorValues["$($r.AppName)"]
            $toSave += [PSCustomObject]@{
                AppName       = $r.AppName
                LatestVersion = $v.LatestVersion
                ReleaseDate   = $v.ReleaseDate
                Publisher     = $r.Publisher
            }
        }
        $hiddenNames = @()
        foreach ($r in $data.Rows) {
            $v = $script:EditorValues["$($r.AppName)"]
            if ($v -and -not $v.Show) { $hiddenNames += $r.AppName }
        }
        try {
            $n = Save-ManualVersionEntries -Entries $toSave
            $h = Save-ExcludedAppNames -Names $hiddenNames
            $msg = "$n version(s) enregistrée(s) dans le référentiel.`n`n$ManualVersionsFile`n`nElles seront utilisées automatiquement à la prochaine génération."
            if ($h -gt 0) { $msg += "`n`n$h application(s) décochée(s) : elles ne figureront pas dans l'onglet « Inventaire & versions » du rapport." }
            Show-InfoMessage $msg
            $editor.Close()
        } catch {
            Show-ErrorMessage "Enregistrement impossible : $($_.Exception.Message)"
        }
    })
    $editor.Controls.Add($btnSave)

    $btnCancel = New-ThemedButton -Text "Fermer sans enregistrer" -X 620 -Y 604 -W 196 -H 40 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.TextMain -FontSize 9 -Ghost
    $btnCancel.Add_Click({ $editor.Close() })
    $editor.Controls.Add($btnCancel)

    # --- Partage du référentiel entre collègues ---
    $btnImport = New-ThemedButton -Text "Importer..." -X 24 -Y 604 -W 140 -H 40 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.TextMain -FontSize 9 -Ghost
    $btnImport.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
        $dlg.Title  = "Importer un référentiel de versions partagé"
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $imported = Get-ManualVersionEntries -Path $dlg.FileName
        if ($imported.Count -eq 0) {
            Show-ErrorMessage "Aucune version exploitable dans ce fichier.`n`nColonnes attendues (séparateur ';') : AppName;LatestVersion;ReleaseDate"
            return
        }
        # Correspondance insensible à la casse avec les applications déjà listées
        $ciMap = @{}
        foreach ($name in @($script:EditorValues.Keys)) { $ciMap[$name.ToLower()] = $name }
        $updated = 0; $added = 0
        foreach ($k in $imported.Keys) {
            $e = $imported[$k]
            if ($ciMap.ContainsKey($k)) {
                $target = $ciMap[$k]
                $script:EditorValues[$target].LatestVersion = $e.LatestVersion
                $script:EditorValues[$target].ReleaseDate   = $e.ReleaseDate
                $updated++
            } else {
                $data.Rows = @($data.Rows) + [PSCustomObject]@{
                    AppName        = $e.AppName
                    Publisher      = $e.Publisher
                    CurrentVersion = ""
                    FoundVersion   = ""
                    FoundSource    = "importé"
                    LatestVersion  = $e.LatestVersion
                    ReleaseDate    = $e.ReleaseDate
                }
                $script:EditorValues[$e.AppName] = [PSCustomObject]@{
                    LatestVersion = $e.LatestVersion
                    ReleaseDate   = $e.ReleaseDate
                    Show          = $true
                }
                $added++
            }
        }
        & $refresh
        Show-InfoMessage "Import terminé : $updated version(s) appliquée(s) aux applications existantes, $added application(s) ajoutée(s).`n`nCliquez sur Enregistrer pour les conserver dans votre référentiel."
    })
    $editor.Controls.Add($btnImport)

    $btnExport = New-ThemedButton -Text "Exporter..." -X 174 -Y 604 -W 140 -H 40 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.TextMain -FontSize 9 -Ghost
    $btnExport.Add_Click({
        try { $grid.EndEdit() | Out-Null } catch { }
        $toExport = @()
        foreach ($r in $data.Rows) {
            $v = $script:EditorValues["$($r.AppName)"]
            if ($v -and -not [string]::IsNullOrWhiteSpace($v.LatestVersion)) {
                $toExport += [PSCustomObject]@{
                    AppName       = $r.AppName
                    LatestVersion = $v.LatestVersion
                    ReleaseDate   = $v.ReleaseDate
                    Publisher     = $r.Publisher
                }
            }
        }
        if ($toExport.Count -eq 0) {
            Show-InfoMessage "Aucune version saisie à exporter pour le moment.`n`nSaisissez d'abord des versions dans la colonne « VERSION À SAISIR »."
            return
        }
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter   = "Fichiers CSV (*.csv)|*.csv"
        $dlg.Title    = "Exporter le référentiel pour le partager"
        $dlg.FileName = "versions-manuelles_partage_" + (Get-Date -Format "yyyyMMdd") + ".csv"
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            $n = Save-ManualVersionEntries -Entries $toExport -Path $dlg.FileName
            Show-InfoMessage "$n version(s) exportée(s) vers :`n$($dlg.FileName)`n`nVos collègues peuvent l'importer via « Importer... », ou le déposer directement en C:\temp\versions-manuelles.csv sur leur poste."
        } catch {
            Show-ErrorMessage "Export impossible : $($_.Exception.Message)"
        }
    })
    $editor.Controls.Add($btnExport)

    & $refresh
    [void]$editor.ShowDialog()
}

# ========================================
# INTERFACE GRAPHIQUE
# (apparence entièrement pilotée par le bloc $Theme en tête de script)
# ========================================

$form = New-Object System.Windows.Forms.Form
$form.Text            = $Theme.WindowTitle
$form.ClientSize      = New-Object System.Drawing.Size(980, 800)
$form.StartPosition   = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox     = $false
$form.BackColor       = ConvertTo-UIColor $Theme.FormBack
$form.Font            = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$form.Add_Load({ Enable-ModernWindowCorners -TargetForm $form })

# Icône dessinée (petit graphique à barres aux couleurs du thème)
$iconBitmap   = New-Object System.Drawing.Bitmap(32, 32)
$iconGraphics = [System.Drawing.Graphics]::FromImage($iconBitmap)
$iconGraphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
$iconGraphics.Clear((ConvertTo-UIColor $Theme.Accent))
$whiteBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)
$iconGraphics.FillRectangle($whiteBrush, 6, 16, 4, 12)
$iconGraphics.FillRectangle($whiteBrush, 14, 10, 4, 18)
$iconGraphics.FillRectangle($whiteBrush, 22, 6, 4, 22)
$whiteBrush.Dispose(); $iconGraphics.Dispose()
$form.Icon = [System.Drawing.Icon]::FromHandle($iconBitmap.GetHicon())

# ----- En-tête dégradé (façon dashboard) -----
$headerPanel          = New-Object System.Windows.Forms.Panel
$headerPanel.Location = New-Object System.Drawing.Point(0, 0)
$headerPanel.Size     = New-Object System.Drawing.Size(980, 96)
$headerPanel.Add_Paint({
    param($sender, $e)
    $rect  = $sender.ClientRectangle
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush($rect, (ConvertTo-UIColor $Theme.HeaderColor1), (ConvertTo-UIColor $Theme.HeaderColor2), 25)
    $e.Graphics.FillRectangle($brush, $rect)
    $brush.Dispose()
})
$form.Controls.Add($headerPanel)

$lblTitle           = New-Object System.Windows.Forms.Label
$lblTitle.Location  = New-Object System.Drawing.Point(32, 20)
$lblTitle.Size      = New-Object System.Drawing.Size(916, 34)
$lblTitle.Text      = $Theme.HeaderTitle
$lblTitle.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = [System.Drawing.Color]::White
$lblTitle.BackColor = [System.Drawing.Color]::Transparent
$headerPanel.Controls.Add($lblTitle)

$lblSubtitle           = New-Object System.Windows.Forms.Label
$lblSubtitle.Location  = New-Object System.Drawing.Point(34, 57)
$lblSubtitle.Size      = New-Object System.Drawing.Size(914, 24)
$lblSubtitle.Text      = $Theme.HeaderSubtitle
$lblSubtitle.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9.5)
$lblSubtitle.ForeColor = [System.Drawing.Color]::FromArgb(225, 255, 255, 255)
$lblSubtitle.BackColor = [System.Drawing.Color]::Transparent
$headerPanel.Controls.Add($lblSubtitle)

# ----- Onglets (dessin personnalisé : onglet actif en couleur, souligné) -----
$tabControl          = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(24, 112)
$tabControl.Size     = New-Object System.Drawing.Size(932, 574)
$tabControl.DrawMode = [System.Windows.Forms.TabDrawMode]::OwnerDrawFixed
$tabControl.SizeMode = "Fixed"
$tabControl.ItemSize = New-Object System.Drawing.Size(210, 42)
$tabControl.Add_DrawItem({
    param($sender, $e)
    $g = $e.Graphics
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $isSelected = ($sender.SelectedIndex -eq $e.Index)
    $rect = $e.Bounds
    $backBrush = New-Object System.Drawing.SolidBrush((ConvertTo-UIColor $(if ($isSelected) { $Theme.CardBack } else { $Theme.FormBack })))
    $g.FillRectangle($backBrush, $rect); $backBrush.Dispose()
    $style = if ($isSelected) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $font  = New-Object System.Drawing.Font($Theme.FontFamily, 9.75, $style)
    $foreBrush = New-Object System.Drawing.SolidBrush((ConvertTo-UIColor $(if ($isSelected) { $Theme.Accent } else { $Theme.TextMuted })))
    $sf = New-Object System.Drawing.StringFormat
    $sf.Alignment = "Center"; $sf.LineAlignment = "Center"
    $rectF = New-Object System.Drawing.RectangleF($rect.X, $rect.Y, $rect.Width, $rect.Height)
    $g.DrawString($sender.TabPages[$e.Index].Text, $font, $foreBrush, $rectF, $sf)
    $foreBrush.Dispose(); $font.Dispose(); $sf.Dispose()
    if ($isSelected) {
        $barBrush = New-Object System.Drawing.SolidBrush((ConvertTo-UIColor $Theme.Accent))
        $g.FillRectangle($barBrush, $rect.X + 16, $rect.Bottom - 4, $rect.Width - 32, 3)
        $barBrush.Dispose()
    }
})
$form.Controls.Add($tabControl)

# ========================================
# ONGLET 1 : CLIENT & OPTIONS GÉNÉRALES
# ========================================

$tabConfig           = New-Object System.Windows.Forms.TabPage
$tabConfig.Text      = "Client & Options"
$tabConfig.BackColor = ConvertTo-UIColor $Theme.FormBack
$tabControl.Controls.Add($tabConfig)

$lblClient           = New-Object System.Windows.Forms.Label
$lblClient.Location  = New-Object System.Drawing.Point(28, 24)
$lblClient.Size      = New-Object System.Drawing.Size(140, 26)
$lblClient.Text      = "Client Intune :"
$lblClient.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 10, [System.Drawing.FontStyle]::Bold)
$lblClient.ForeColor = ConvertTo-UIColor $Theme.TextMain
$lblClient.BackColor = [System.Drawing.Color]::Transparent
$tabConfig.Controls.Add($lblClient)

$cmbClients               = New-Object System.Windows.Forms.ComboBox
$cmbClients.Location      = New-Object System.Drawing.Point(172, 21)
$cmbClients.Size          = New-Object System.Drawing.Size(724, 26)
$cmbClients.DropDownStyle = "DropDownList"
$cmbClients.FlatStyle     = "Flat"
$cmbClients.Font          = New-Object System.Drawing.Font($Theme.FontFamily, 10)
$tabConfig.Controls.Add($cmbClients)

$cardOptions = New-CardPanel -Title "Options générales" -X 24 -Y 62 -W 880 -H 108
$tabConfig.Controls.Add($cardOptions)

$chkExcludeVM           = New-Object System.Windows.Forms.CheckBox
$chkExcludeVM.Location  = New-Object System.Drawing.Point(22, 46)
$chkExcludeVM.Size      = New-Object System.Drawing.Size(830, 24)
$chkExcludeVM.Text      = "Exclure les machines virtuelles de l'analyse de conformité"
$chkExcludeVM.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$chkExcludeVM.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardOptions.Controls.Add($chkExcludeVM)

$lblTabs           = New-Object System.Windows.Forms.Label
$lblTabs.Location  = New-Object System.Drawing.Point(22, 76)
$lblTabs.Size      = New-Object System.Drawing.Size(196, 24)
$lblTabs.Text      = "Onglets inclus dans le rapport :"
$lblTabs.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
$lblTabs.ForeColor = ConvertTo-UIColor $Theme.TextMain
$lblTabs.BackColor = [System.Drawing.Color]::Transparent
$cardOptions.Controls.Add($lblTabs)

$tabToggles = @(
    @{ Var = "chkTabCompliance";  Text = "1. Conformité";    X = 226; W = 118 }
    @{ Var = "chkTabDiscovered";  Text = "2. Apps découvertes"; X = 350; W = 168 }
    @{ Var = "chkTabInventory";   Text = "3. Inventaire";     X = 524; W = 140 }
    @{ Var = "chkTabRemediation"; Text = "4. Remédiation";    X = 670; W = 160 }
)
foreach ($tg in $tabToggles) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Location  = New-Object System.Drawing.Point($tg.X, 74)
    $c.Size      = New-Object System.Drawing.Size($tg.W, 24)
    $c.Text      = $tg.Text
    $c.Checked   = $true
    $c.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
    $c.ForeColor = ConvertTo-UIColor $Theme.TextMain
    $cardOptions.Controls.Add($c)
    Set-Variable -Name $tg.Var -Value $c -Scope Script
}

# ----- Anonymisation du rapport -----
$cardAnon = New-CardPanel -Title "Anonymisation du rapport" -X 24 -Y 180 -W 880 -H 136
$tabConfig.Controls.Add($cardAnon)
$tipAnon = New-Object System.Windows.Forms.ToolTip
$tipAnon.AutoPopDelay = 20000

$chkAnonymize           = New-Object System.Windows.Forms.CheckBox
$chkAnonymize.Location  = New-Object System.Drawing.Point(22, 44)
$chkAnonymize.Size      = New-Object System.Drawing.Size(400, 24)
$chkAnonymize.Text      = "Noms de postes et d'utilisateurs"
$chkAnonymize.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$chkAnonymize.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardAnon.Controls.Add($chkAnonymize)
$tipAnon.SetToolTip($chkAnonymize, "Remplacés par Poste-xxxxxxxx / User-xxxxxxxx, identiques d'un onglet à l'autre.")

$chkAnonClient           = New-Object System.Windows.Forms.CheckBox
$chkAnonClient.Location  = New-Object System.Drawing.Point(440, 44)
$chkAnonClient.Size      = New-Object System.Drawing.Size(416, 24)
$chkAnonClient.Text      = "Nom du client et termes ci-dessous (remplacés par « $AnonClientLabel »)"
$chkAnonClient.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$chkAnonClient.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardAnon.Controls.Add($chkAnonClient)
$tipAnon.SetToolTip($chkAnonClient, "Masque le nom du profil client et chaque terme listé partout dans le rapport : en-tête, titre, stratégies, éditeurs, noms de paquets, nom du fichier.")

$chkAnonApps           = New-Object System.Windows.Forms.CheckBox
$chkAnonApps.Location  = New-Object System.Drawing.Point(22, 72)
$chkAnonApps.Size      = New-Object System.Drawing.Size(834, 24)
$chkAnonApps.Text      = "Noms d'applications (apps_1, apps_2...) et d'exécutables en échec (proc_1, proc_2...)"
$chkAnonApps.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$chkAnonApps.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardAnon.Controls.Add($chkAnonApps)
$tipAnon.SetToolTip($chkAnonApps, "Onglets Applications découvertes, Inventaire et Remédiation. Une même application garde le même numéro dans tous les onglets.")

$lblAnonTerms           = New-Object System.Windows.Forms.Label
$lblAnonTerms.Location  = New-Object System.Drawing.Point(22, 103)
$lblAnonTerms.Size      = New-Object System.Drawing.Size(190, 24)
$lblAnonTerms.Text      = "Autres termes (séparés par ;) :"
$lblAnonTerms.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$lblAnonTerms.ForeColor = ConvertTo-UIColor $Theme.TextMuted
$lblAnonTerms.BackColor = [System.Drawing.Color]::Transparent
$cardAnon.Controls.Add($lblAnonTerms)

$txtAnonTerms             = New-Object System.Windows.Forms.TextBox
$txtAnonTerms.Location    = New-Object System.Drawing.Point(214, 100)
$txtAnonTerms.Size        = New-Object System.Drawing.Size(642, 25)
$txtAnonTerms.Font        = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$txtAnonTerms.BorderStyle = "FixedSingle"
$txtAnonTerms.Enabled     = $false
$cardAnon.Controls.Add($txtAnonTerms)
$tipAnon.SetToolTip($txtAnonTerms, ("Exemple : Allianz;AZ;MOPAZ`n" +
    "4 caractères ou plus : remplacé partout, même au milieu d'un mot.`n" +
    "Moins de 4 caractères (sigle) : remplacé seulement s'il forme un mot entier.`n" +
    "Liste mémorisée pour chaque client."))

$chkAnonClient.Add_CheckedChanged({ $txtAnonTerms.Enabled = $chkAnonClient.Checked })
# Termes mémorisés rechargés à chaque changement de client (jamais ceux d'un autre client).
$cmbClients.Add_SelectedIndexChanged({
    if ($cmbClients.SelectedIndex -le 0) { $txtAnonTerms.Text = ""; return }
    $txtAnonTerms.Text = Get-AnonymizationTerms -ClientName $cmbClients.Text.Trim()
})

$cardContact = New-CardPanel -Title "Coordonnées affichées sur le rapport" -X 24 -Y 326 -W 880 -H 190
$tabConfig.Controls.Add($cardContact)

$chkShowContact           = New-Object System.Windows.Forms.CheckBox
$chkShowContact.Location  = New-Object System.Drawing.Point(22, 44)
$chkShowContact.Size      = New-Object System.Drawing.Size(420, 24)
$chkShowContact.Text      = "Afficher les coordonnées de contact"
$chkShowContact.Checked   = $true
$chkShowContact.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
$chkShowContact.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardContact.Controls.Add($chkShowContact)

$contactRows = @(
    @{ Label = "Nom entreprise :"; Var = "txtCompanyName";   Default = $DefaultCompanyName;   Y = 74  }
    @{ Label = "Contact :";        Var = "txtContactPerson"; Default = $DefaultContactPerson; Y = 102 }
    @{ Label = "Email :";          Var = "txtContactEmail";  Default = $DefaultContactEmail;  Y = 130 }
    @{ Label = "Téléphone :";      Var = "txtContactPhone";  Default = $DefaultContactPhone;  Y = 158 }
)
foreach ($row in $contactRows) {
    $l = New-Object System.Windows.Forms.Label
    $l.Location  = New-Object System.Drawing.Point(22, $row.Y)
    $l.Size      = New-Object System.Drawing.Size(150, 24)
    $l.Text      = $row.Label
    $l.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
    $l.ForeColor = ConvertTo-UIColor $Theme.TextMuted
    $l.BackColor = [System.Drawing.Color]::Transparent
    $cardContact.Controls.Add($l)

    $t = New-Object System.Windows.Forms.TextBox
    $t.Location    = New-Object System.Drawing.Point(180, ($row.Y - 2))
    $t.Size        = New-Object System.Drawing.Size(676, 25)
    $t.Text        = $row.Default
    $t.Font        = New-Object System.Drawing.Font($Theme.FontFamily, 9)
    $t.BorderStyle = "FixedSingle"
    $cardContact.Controls.Add($t)
    Set-Variable -Name $row.Var -Value $t -Scope Script
}

# ========================================
# ONGLET 2 : SOURCE DES DONNÉES (API / IMPORT DE FICHIERS / MIXTE)
# ========================================

$tabSource           = New-Object System.Windows.Forms.TabPage
$tabSource.Text      = "Source des données"
$tabSource.BackColor = ConvertTo-UIColor $Theme.FormBack
$tabControl.Controls.Add($tabSource)

$tipImport = New-Object System.Windows.Forms.ToolTip
$tipImport.AutoPopDelay = 20000
$tipImport.InitialDelay = 350

$cardSource = New-CardPanel -Title "Mode de collecte des données" -X 24 -Y 24 -W 880 -H 158
$tabSource.Controls.Add($cardSource)

$rbSourceApi           = New-Object System.Windows.Forms.RadioButton
$rbSourceApi.Location  = New-Object System.Drawing.Point(22, 44)
$rbSourceApi.Size      = New-Object System.Drawing.Size(830, 24)
$rbSourceApi.Text      = "API Microsoft Graph — collecte en ligne sur le tenant du client (mode historique)"
$rbSourceApi.Checked   = $true
$rbSourceApi.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$rbSourceApi.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardSource.Controls.Add($rbSourceApi)

$rbSourceImport           = New-Object System.Windows.Forms.RadioButton
$rbSourceImport.Location  = New-Object System.Drawing.Point(22, 70)
$rbSourceImport.Size      = New-Object System.Drawing.Size(830, 24)
$rbSourceImport.Text      = "Import de fichiers — aucune connexion au tenant : tout provient des fichiers ci-dessous"
$rbSourceImport.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$rbSourceImport.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardSource.Controls.Add($rbSourceImport)

$rbSourceHybrid           = New-Object System.Windows.Forms.RadioButton
$rbSourceHybrid.Location  = New-Object System.Drawing.Point(22, 96)
$rbSourceHybrid.Size      = New-Object System.Drawing.Size(830, 24)
$rbSourceHybrid.Text      = "Mixte — API pour ce qui n'a pas de fichier, fichier prioritaire pour le reste"
$rbSourceHybrid.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$rbSourceHybrid.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardSource.Controls.Add($rbSourceHybrid)

$lblImportClient           = New-Object System.Windows.Forms.Label
$lblImportClient.Location  = New-Object System.Drawing.Point(22, 126)
$lblImportClient.Size      = New-Object System.Drawing.Size(196, 24)
$lblImportClient.Text      = "Nom du client (mode import) :"
$lblImportClient.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$lblImportClient.ForeColor = ConvertTo-UIColor $Theme.TextMuted
$lblImportClient.BackColor = [System.Drawing.Color]::Transparent
$cardSource.Controls.Add($lblImportClient)

$txtImportClientName             = New-Object System.Windows.Forms.TextBox
$txtImportClientName.Location    = New-Object System.Drawing.Point(224, 124)
$txtImportClientName.Size        = New-Object System.Drawing.Size(230, 25)
$txtImportClientName.Font        = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$txtImportClientName.BorderStyle = "FixedSingle"
$cardSource.Controls.Add($txtImportClientName)
$tipImport.SetToolTip($txtImportClientName, "Sert uniquement à titrer le rapport : aucune configuration .clientconfig n'est nécessaire en mode import. Ignoré si un client est sélectionné dans l'onglet précédent.")

$chkExportRaw           = New-Object System.Windows.Forms.CheckBox
$chkExportRaw.Location  = New-Object System.Drawing.Point(478, 124)
$chkExportRaw.Size      = New-Object System.Drawing.Size(378, 26)
$chkExportRaw.Text      = "Exporter les données collectées (CSV réimportables)"
$chkExportRaw.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$chkExportRaw.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardSource.Controls.Add($chkExportRaw)
$tipImport.SetToolTip($chkExportRaw, "À la fin d'une collecte API, écrit les données dans $ImportFolder au format attendu par l'import : la même génération peut ensuite être rejouée hors ligne, sans accès au tenant.")

$cardImportFiles = New-CardPanel -Title "Fichiers d'import (CSV, TSV ou JSON — tous facultatifs)" -X 24 -Y 194 -W 880 -H 316
$tabSource.Controls.Add($cardImportFiles)

$importRows = @(
    @{ Label = "Appareils gérés :";           TxtVar = "txtImpDevices";    BtnVar = "btnImpDevices";    Y = 42
       Hint  = "Onglet 1 du rapport. Export « Tous les appareils » du portail Intune, ou JSON Graph /deviceManagement/managedDevices.`nColonnes utiles : Device name, Primary UPN, OS, OS version, Compliance, Last check-in, Manufacturer, Model." }
    @{ Label = "Détail de conformité :";      TxtVar = "txtImpCompliance"; BtnVar = "btnImpCompliance"; Y = 72
       Hint  = "Une ligne par couple appareil / règle en échec, pour regrouper les postes par motif.`nColonnes utiles : Device name + (Motif OU Setting OU Policy name). Sans ce fichier, les postes sont regroupés sous un motif générique." }
    @{ Label = "Applications découvertes :";  TxtVar = "txtImpDiscovered"; BtnVar = "btnImpDiscovered"; Y = 102
       Hint  = "Onglet 2 du rapport. Export « Applications découvertes » (une ligne par application).`nColonnes utiles : Application name, Version, Publisher, Devices (nombre de postes)." }
    @{ Label = "Postes par application :";    TxtVar = "txtImpAppDevices"; BtnVar = "btnImpAppDevices"; Y = 132
       Hint  = "Détail déplié sous chaque application : une ligne par couple application / poste.`nColonnes utiles : Application name, Device name, UPN, OS. Peut être utilisé seul : le nombre de postes est alors calculé." }
    @{ Label = "Inventaire des applications :"; TxtVar = "txtImpInventory"; BtnVar = "btnImpInventory"; Y = 162
       Hint  = "Onglet 3 du rapport. Export « Toutes les applications » du portail, ou JSON Graph /deviceAppManagement/mobileApps.`nColonnes utiles : Name, Publisher, Version. L'audit des versions publiques s'applique ensuite normalement." }
    @{ Label = "Scores de santé (EA) :";       TxtVar = "txtImpScores";     BtnVar = "btnImpScores";     Y = 192
       Hint  = "Onglet 4 du rapport. Export « Device scores » (Rapports > Analyse des points de terminaison), ou JSON Graph userExperienceAnalyticsDeviceScores.`nColonnes utiles : Device name, Endpoint analytics score, Battery health score. Sans ce fichier, l'API est interrogée automatiquement (meilleur effort)." }
    @{ Label = "Performances de démarrage :"; TxtVar = "txtImpPerf";      BtnVar = "btnImpPerf";      Y = 222
       Hint  = "Onglet 4 du rapport. Export « Startup performance », ou JSON Graph userExperienceAnalyticsDevicePerformance.`nColonnes utiles : Device name, Core boot time, Blue screen count, Restart count. Sans ce fichier, l'API est interrogée automatiquement (meilleur effort)." }
)

foreach ($row in $importRows) {
    $l = New-Object System.Windows.Forms.Label
    $l.Location  = New-Object System.Drawing.Point(22, ($row.Y + 3))
    $l.Size      = New-Object System.Drawing.Size(228, 22)
    $l.Text      = $row.Label
    $l.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
    $l.ForeColor = ConvertTo-UIColor $Theme.TextMain
    $l.BackColor = [System.Drawing.Color]::Transparent
    $cardImportFiles.Controls.Add($l)

    $t = New-Object System.Windows.Forms.TextBox
    $t.Location    = New-Object System.Drawing.Point(256, $row.Y)
    $t.Size        = New-Object System.Drawing.Size(470, 25)
    $t.Font        = New-Object System.Drawing.Font($Theme.FontFamily, 8.75)
    $t.BorderStyle = "FixedSingle"
    $cardImportFiles.Controls.Add($t)
    $tipImport.SetToolTip($t, $row.Hint)

    $b = New-ThemedButton -Text "Parcourir..." -X 736 -Y ($row.Y - 1) -W 118 -H 28 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.TextMain -FontSize 8.5 -Ghost
    $b.Tag = $t
    $b.Add_Click({
        param($s, $e)
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Exports Intune (*.csv;*.json;*.tsv;*.txt)|*.csv;*.json;*.tsv;*.txt|Tous les fichiers (*.*)|*.*"
        if (Test-Path $ImportFolder) { $dlg.InitialDirectory = $ImportFolder }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $s.Tag.Text = $dlg.FileName }
    })
    $cardImportFiles.Controls.Add($b)

    Set-Variable -Name $row.TxtVar -Value $t -Scope Script
    Set-Variable -Name $row.BtnVar -Value $b -Scope Script
}

$lblImportNote           = New-Object System.Windows.Forms.Label
$lblImportNote.Location  = New-Object System.Drawing.Point(22, 256)
$lblImportNote.Size      = New-Object System.Drawing.Size(660, 30)
$lblImportNote.Text      = "Les noms de colonnes sont reconnus en français comme en anglais (casse et accents ignorés). Survolez un champ pour voir les colonnes attendues."
$lblImportNote.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.5, [System.Drawing.FontStyle]::Italic)
$lblImportNote.ForeColor = ConvertTo-UIColor $Theme.TextMuted
$lblImportNote.BackColor = [System.Drawing.Color]::Transparent
$cardImportFiles.Controls.Add($lblImportNote)

$btnImportTemplates = New-ThemedButton -Text "Créer des modèles..." -X 700 -Y 254 -W 154 -H 30 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.Accent -FontSize 8.5 -Ghost
$btnImportTemplates.Add_Click({
    try {
        $folder = New-ImportTemplates
        Show-InfoMessage "Modèles de fichiers d'import créés dans :`n$folder`n`nComplétez-les (Excel), puis désignez-les ci-dessus."
        Start-Process -FilePath "explorer.exe" -ArgumentList "`"$folder`""
    } catch {
        Show-ErrorMessage "Création des modèles impossible : $($_.Exception.Message)"
    }
})
$cardImportFiles.Controls.Add($btnImportTemplates)

$rbSourceApi.Add_CheckedChanged({ Update-SourceModeUi })
$rbSourceImport.Add_CheckedChanged({ Update-SourceModeUi })
$rbSourceHybrid.Add_CheckedChanged({ Update-SourceModeUi })
Update-SourceModeUi

# ========================================
# ONGLET 3 : RAPPORTS (Discovered Apps + Inventaire/Versions)
# ========================================

$tabReports           = New-Object System.Windows.Forms.TabPage
$tabReports.Text      = "Rapports & Versions"
$tabReports.BackColor = ConvertTo-UIColor $Theme.FormBack
$tabControl.Controls.Add($tabReports)

$cardDiscovered = New-CardPanel -Title "Applications découvertes (Onglet 2 du rapport)" -X 24 -Y 24 -W 880 -H 118
$tabReports.Controls.Add($cardDiscovered)

$lblTopApps           = New-Object System.Windows.Forms.Label
$lblTopApps.Location  = New-Object System.Drawing.Point(22, 46)
$lblTopApps.Size      = New-Object System.Drawing.Size(490, 52)
$lblTopApps.Text      = "Applications dont la liste des postes est chargée :`nToutes par défaut ; limitez aux N plus répandues seulement si la collecte est trop longue."
$lblTopApps.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$lblTopApps.ForeColor = ConvertTo-UIColor $Theme.TextMain
$lblTopApps.BackColor = [System.Drawing.Color]::Transparent
$cardDiscovered.Controls.Add($lblTopApps)

# [V2.3] « Toutes les applications » coché par défaut : la liste des postes de CHAQUE
# application est chargée. Le nombre ne sert plus qu'à limiter volontairement la collecte.
$chkAllApps           = New-Object System.Windows.Forms.CheckBox
$chkAllApps.Location  = New-Object System.Drawing.Point(530, 53)
$chkAllApps.Size      = New-Object System.Drawing.Size(190, 24)
$chkAllApps.Text      = "Toutes les applications"
$chkAllApps.Checked   = $true
$chkAllApps.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9, [System.Drawing.FontStyle]::Bold)
$chkAllApps.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardDiscovered.Controls.Add($chkAllApps)

$numTopApps             = New-Object System.Windows.Forms.NumericUpDown
$numTopApps.Location    = New-Object System.Drawing.Point(730, 52)
$numTopApps.Size        = New-Object System.Drawing.Size(90, 26)
$numTopApps.Minimum     = 1
$numTopApps.Maximum     = 5000
$numTopApps.Value       = 50
$numTopApps.Increment   = 10
$numTopApps.Font        = New-Object System.Drawing.Font($Theme.FontFamily, 10)
$numTopApps.BorderStyle = "FixedSingle"
$numTopApps.Enabled     = $false
$cardDiscovered.Controls.Add($numTopApps)

$chkAllApps.Add_CheckedChanged({ $numTopApps.Enabled = -not $chkAllApps.Checked })
$tipTopApps = New-Object System.Windows.Forms.ToolTip
$tipTopApps.AutoPopDelay = 20000
$tipTopApps.InitialDelay = 350
$tipTopApps.SetToolTip($chkAllApps, ("Charge la liste complète des postes de chaque application découverte." + [Environment]::NewLine +
    "Coût : une requête par application et une par tranche de 999 postes, envoyées par lots de 20." + [Environment]::NewLine +
    "Décochez pour ne détailler que les N applications les plus répandues (collecte plus courte)."))
$tipTopApps.SetToolTip($numTopApps, "Nombre d'applications détaillées quand « Toutes les applications » est décoché (les plus répandues d'abord).")

$cardVersions = New-CardPanel -Title "Inventaire & audit des versions (Onglet 3 du rapport)" -X 24 -Y 156 -W 880 -H 270
$tabReports.Controls.Add($cardVersions)

$chkUseWinget           = New-Object System.Windows.Forms.CheckBox
$chkUseWinget.Location  = New-Object System.Drawing.Point(22, 46)
$chkUseWinget.Size      = New-Object System.Drawing.Size(830, 40)
$chkUseWinget.Text      = "Rechercher les dernières versions publiques en ligne (API éditeurs, winget, winget-pkgs, Chocolatey). Décochez pour n'utiliser que le fichier CSV."
$chkUseWinget.Checked   = $true
$chkUseWinget.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$chkUseWinget.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardVersions.Controls.Add($chkUseWinget)

$chkUseGitHubPkgs           = New-Object System.Windows.Forms.CheckBox
$chkUseGitHubPkgs.Location  = New-Object System.Drawing.Point(22, 90)
$chkUseGitHubPkgs.Size      = New-Object System.Drawing.Size(414, 24)
$chkUseGitHubPkgs.Text      = "Interroger le catalogue winget-pkgs sur GitHub"
$chkUseGitHubPkgs.Checked   = $true
$chkUseGitHubPkgs.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$chkUseGitHubPkgs.ForeColor = ConvertTo-UIColor $Theme.TextMain
$cardVersions.Controls.Add($chkUseGitHubPkgs)

$tipVersions = New-Object System.Windows.Forms.ToolTip
$tipVersions.AutoPopDelay = 20000
$tipVersions.InitialDelay = 350
$tipVersions.SetToolTip($chkUseGitHubPkgs, ("Lit le même catalogue que winget, mais directement sur https://github.com/microsoft/winget-pkgs via api.github.com." + [Environment]::NewLine +
    "Utile quand winget est absent de la session (compte admin, Windows Server) ou que son CDN est bloqué par le proxy."))

$lblGhToken           = New-Object System.Windows.Forms.Label
$lblGhToken.Location  = New-Object System.Drawing.Point(446, 92)
$lblGhToken.Size      = New-Object System.Drawing.Size(112, 22)
$lblGhToken.Text      = "Jeton GitHub :"
$lblGhToken.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.75)
$lblGhToken.ForeColor = ConvertTo-UIColor $Theme.TextMuted
$lblGhToken.BackColor = [System.Drawing.Color]::Transparent
$cardVersions.Controls.Add($lblGhToken)

$txtGhToken             = New-Object System.Windows.Forms.TextBox
$txtGhToken.Location    = New-Object System.Drawing.Point(560, 89)
$txtGhToken.Size        = New-Object System.Drawing.Size(292, 24)
$txtGhToken.Font        = New-Object System.Drawing.Font($Theme.FontFamily, 8.75)
$txtGhToken.BorderStyle = "FixedSingle"
$txtGhToken.UseSystemPasswordChar = $true
$txtGhToken.Text        = Get-GitHubToken
$cardVersions.Controls.Add($txtGhToken)
$tipVersions.SetToolTip($txtGhToken, ("Facultatif. Sans jeton, l'API GitHub est limitée à 60 requêtes/heure (environ 25 applications)." + [Environment]::NewLine +
    "Un jeton d'accès personnel SANS AUCUNE PORTÉE suffit (le dépôt est public) et porte la limite à 5 000/heure." + [Environment]::NewLine +
    "Il est conservé en clair dans " + $GitHubTokenFile + " : n'utilisez pas un jeton disposant de droits d'écriture."))

$lblOverrideCsv           = New-Object System.Windows.Forms.Label
$lblOverrideCsv.Location  = New-Object System.Drawing.Point(22, 120)
$lblOverrideCsv.Size      = New-Object System.Drawing.Size(830, 32)
$lblOverrideCsv.Text      = "Fichier CSV de correspondance (facultatif). Laissez vide : le référentiel des versions manuelles est repris automatiquement. Utilisez le bouton ci-dessous pour le remplir."
$lblOverrideCsv.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.75)
$lblOverrideCsv.ForeColor = ConvertTo-UIColor $Theme.TextMuted
$lblOverrideCsv.BackColor = [System.Drawing.Color]::Transparent
$cardVersions.Controls.Add($lblOverrideCsv)

$txtOverrideCsv             = New-Object System.Windows.Forms.TextBox
$txtOverrideCsv.Location    = New-Object System.Drawing.Point(22, 156)
$txtOverrideCsv.Size        = New-Object System.Drawing.Size(706, 25)
$txtOverrideCsv.Font        = New-Object System.Drawing.Font($Theme.FontFamily, 9)
$txtOverrideCsv.BorderStyle = "FixedSingle"
$cardVersions.Controls.Add($txtOverrideCsv)

$btnBrowseCsv = New-ThemedButton -Text "Parcourir..." -X 740 -Y 153 -W 112 -H 30 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.TextMain -FontSize 8.75 -Ghost
$btnBrowseCsv.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtOverrideCsv.Text = $dlg.FileName
    }
})
$cardVersions.Controls.Add($btnBrowseCsv)

$btnTestWinget = New-ThemedButton -Text "Tester la recherche de versions" -X 22 -Y 194 -W 232 -H 32 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.Accent -FontSize 8.75 -Ghost
$btnTestWinget.Add_Click({
    $old = $btnTestWinget.Text
    $btnTestWinget.Text = "Test en cours..."
    $btnTestWinget.Enabled = $false
    $form.Refresh()
    try {
        $diag = Test-VersionResolution
        Write-Log "Diagnostic de resolution des versions :`n$diag" -Level INFO
        Show-InfoMessage $diag
    } catch {
        Show-ErrorMessage "Le diagnostic a echoue : $($_.Exception.Message)"
    } finally {
        $btnTestWinget.Text = $old
        $btnTestWinget.Enabled = $true
    }
})
$cardVersions.Controls.Add($btnTestWinget)

$btnManualVersions = New-ThemedButton -Text "Versions manuelles..." -X 266 -Y 194 -W 180 -H 32 -BackHex $Theme.Accent -HoverHex $Theme.AccentHover -FontSize 8.75
$btnManualVersions.Add_Click({
    try { Show-VersionEditor }
    catch { Show-ErrorMessage "Ouverture de l'editeur impossible : $($_.Exception.Message)" }
})
$cardVersions.Controls.Add($btnManualVersions)

$btnAnalyzeImport = New-ThemedButton -Text "Analyser les versions du fichier importé..." -X 22 -Y 232 -W 424 -H 32 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.Accent -FontSize 8.75 -Ghost
$btnAnalyzeImport.Add_Click({
    # Recherche les dernières versions publiques pour les applications du fichier
    # d'inventaire importé, PUIS ouvre l'éditeur pré-rempli — sans générer de rapport.
    $path = ""
    if ($txtImpInventory) { $path = $txtImpInventory.Text.Trim() }
    if ([string]::IsNullOrWhiteSpace($path)) {
        Show-InfoMessage ("Aucun fichier d'inventaire n'est renseigné.`n`n" +
                          "Renseignez « Inventaire des applications » dans l'onglet « Source des données » " +
                          "(mode Import ou Mixte), puis relancez l'analyse.")
        return
    }
    if (-not (Test-Path $path)) { Show-ErrorMessage "Fichier introuvable :`n$path"; return }

    $old = $btnAnalyzeImport.Text
    $btnAnalyzeImport.Text    = "Analyse en cours..."
    $btnAnalyzeImport.Enabled = $false
    $form.Refresh()
    try {
        $apps = @(Import-MobileAppsFile -Path $path)
        if ($apps.Count -eq 0) {
            Show-ErrorMessage ("Aucune application n'a pu être lue dans ce fichier.`n`n" +
                               "Vérifiez qu'il contient une colonne de nom d'application (Name / displayName / Nom).")
            return
        }
        $prog = { param($msg) $lblStatus.Text = $msg; $form.Refresh(); Write-Log $msg }
        Save-GitHubToken -Token $txtGhToken.Text.Trim()
        $inv  = Get-TenantAppInventory -PreloadedApps $apps -OverrideCsvPath $txtOverrideCsv.Text.Trim() `
                    -UseOnlineSources $chkUseWinget.Checked -UseGitHubPkgs $chkUseGitHubPkgs.Checked `
                    -GitHubToken $txtGhToken.Text.Trim() -ProgressCallback $prog
        # Mémorisé pour l'éditeur ET pour une génération ultérieure dans la même session.
        $script:LastInventory = $inv
        $resolved = @($inv | Where-Object { $_.LatestPublicVersion -ne "Non trouvé" }).Count
        $lblStatus.Text = "Analyse terminée : $resolved version(s) trouvée(s) sur $($inv.Count)."
        $lblStatus.ForeColor = [System.Drawing.Color]::Green
        Show-VersionEditor
    } catch {
        Write-Log "Analyse du fichier importé impossible : $($_.Exception.Message)" -Level ERROR
        Show-ErrorMessage "Analyse impossible :`n`n$($_.Exception.Message)"
    } finally {
        $btnAnalyzeImport.Text    = $old
        $btnAnalyzeImport.Enabled = $true
    }
})
$cardVersions.Controls.Add($btnAnalyzeImport)

$lblVersionNote           = New-Object System.Windows.Forms.Label
$lblVersionNote.Location  = New-Object System.Drawing.Point(458, 190)
$lblVersionNote.Size      = New-Object System.Drawing.Size(394, 40)
$lblVersionNote.Text      = "Ces options s'appliquent aussi à un inventaire importé.`nNote : aucune API publique universelle n'existe pour la dernière version d'un logiciel — résultat indicatif."
$lblVersionNote.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.5, [System.Drawing.FontStyle]::Italic)
$lblVersionNote.ForeColor = [System.Drawing.Color]::FromArgb(180, 120, 0)
$lblVersionNote.BackColor = [System.Drawing.Color]::Transparent
$cardVersions.Controls.Add($lblVersionNote)

# ========================================
# ONGLET 4 : REMÉDIATION AVANCÉE (vérifications, imports dédiés, actions)
# ========================================

$tabRemAdv           = New-Object System.Windows.Forms.TabPage
$tabRemAdv.Text      = "Remédiation avancée"
$tabRemAdv.BackColor = ConvertTo-UIColor $Theme.FormBack
$tabControl.Controls.Add($tabRemAdv)

$tipRemAdv = New-Object System.Windows.Forms.ToolTip
$tipRemAdv.AutoPopDelay = 25000
$tipRemAdv.InitialDelay = 350

$cardRemChecks = New-CardPanel -Title "Vérifications incluses dans l'onglet Remédiation (Onglet 4 du rapport)" -X 24 -Y 24 -W 880 -H 190
$tabRemAdv.Controls.Add($cardRemChecks)

$lblRemChecksNote           = New-Object System.Windows.Forms.Label
$lblRemChecksNote.Location  = New-Object System.Drawing.Point(22, 42)
$lblRemChecksNote.Size      = New-Object System.Drawing.Size(836, 32)
$lblRemChecksNote.Text      = "Une vérification décochée n'est ni collectée ni affichée : aucun appel superflu, aucune ligne trompeuse. Survolez une case pour voir son seuil."
$lblRemChecksNote.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.75, [System.Drawing.FontStyle]::Italic)
$lblRemChecksNote.ForeColor = ConvertTo-UIColor $Theme.TextMuted
$lblRemChecksNote.BackColor = [System.Drawing.Color]::Transparent
$cardRemChecks.Controls.Add($lblRemChecksNote)

# Quatre colonnes plutôt que trois : les vérifications sont passées de 11 à 14 et la carte
# ne peut pas grandir en hauteur sans déborder de l'onglet.
$remCheckDefs = @(
    @{ Key = "Disk";           Text = "Espace disque";              Col = 0; Row = 0
       Hint = "Alerte sous $($RemediationThresholds.DiskFreePctWarning) % d'espace libre, critique sous $($RemediationThresholds.DiskFreePctCritical) % ou $($RemediationThresholds.DiskFreeGbCritical) Go." }
    @{ Key = "Inactivity";     Text = "Inactivité (synchro)";        Col = 0; Row = 1
       Hint = "Postes sans synchronisation depuis plus de $($RemediationThresholds.StaleDaysWarning) jours ; critique au-delà de $($RemediationThresholds.StaleDaysCritical) jours." }
    @{ Key = "Boot";           Text = "Démarrage lent / HDD";        Col = 0; Row = 2
       Hint = "Démarrage au-delà de $($RemediationThresholds.BootSlowSeconds) s, et postes encore équipés d'un disque mécanique." }
    @{ Key = "Bsod";           Text = "Écrans bleus / redémarrages"; Col = 0; Row = 3
       Hint = "Écrans bleus sur 14 jours (critique à partir de $($RemediationThresholds.BsodCritical)) et redémarrages anormalement fréquents." }
    @{ Key = "Battery";        Text = "Batterie (score)";            Col = 1; Row = 0
       Hint = "Score de santé composite Endpoint Analytics, alerte sous $($RemediationThresholds.BatteryPoor)." }
    @{ Key = "BatteryDetail";  Text = "Batterie (capacité, âge)";    Col = 1; Row = 1
       Hint = "Capacité maximale restante, âge et autonomie estimée. Alerte sous $($RemediationThresholds.BatteryCapacityPoor) % de capacité, critique sous $($RemediationThresholds.BatteryCapacityCrit) %. Ce sont ces chiffres qui justifient un remplacement auprès d'un utilisateur ou d'un acheteur." }
    @{ Key = "EaScore";        Text = "Score Endpoint Analytics";    Col = 1; Row = 2
       Hint = "Score global du poste, alerte sous $($RemediationThresholds.ScoreLow)." }
    @{ Key = "Uptime";         Text = "Uptime (estimation)";         Col = 1; Row = 3
       Hint = "Temps écoulé depuis le dernier démarrage connu, au-delà de $($RemediationThresholds.UptimeWarningDays) jours. Donnée agrégée quotidiennement par Microsoft : c'est une estimation, pas un compteur en direct." }
    @{ Key = "BitLocker";      Text = "BitLocker";                   Col = 2; Row = 0
       Hint = "Chiffrement du volume système et véritables échecs de protection. Les indicateurs informatifs (consentement, utilisateur non-administrateur) ne sont volontairement pas remontés comme des défauts." }
    @{ Key = "Defender";       Text = "Antivirus (Defender)";        Col = 2; Row = 1
       Hint = "Protection en temps réel désactivée et signatures en retard de plus de $($RemediationThresholds.SignatureStaleDays) jours." }
    @{ Key = "WindowsUpdate";  Text = "Mises à jour Windows";        Col = 2; Row = 2
       Hint = "Échecs de mise à jour qualité. Nécessite que les postes soient rattachés à une stratégie « Anneaux de mise à jour Windows »." }
    @{ Key = "AppReliability"; Text = "Fiabilité des applications";  Col = 2; Row = 3
       Hint = "Application qui plante le plus souvent sur chaque poste, à partir de $($RemediationThresholds.AppCrashWarning) plantages." }
    @{ Key = "Compliance";     Text = "Conformité Intune";           Col = 3; Row = 0
       Hint = "Reprend l'état de conformité du poste au niveau du poste, avec le motif exact issu de l'onglet Conformité : plus besoin de recouper deux onglets avant d'agir. Aucun appel réseau supplémentaire." }
    @{ Key = "ConfigProfile";  Text = "Profils de configuration";    Col = 3; Row = 1
       Hint = "Postes en erreur ou en conflit d'application d'un profil (modèles ET catalogue de paramètres). C'est la panne la plus silencieuse d'Intune : le poste reste conforme, se synchronise, et n'applique pourtant pas le paramétrage attendu. Coût : quelques dizaines d'appels supplémentaires." }
)
$script:ChkRemChecks = @{}
foreach ($cd in $remCheckDefs) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Location  = New-Object System.Drawing.Point((22 + $cd.Col * 214), (78 + $cd.Row * 26))
    $c.Size      = New-Object System.Drawing.Size(206, 24)
    $c.Text      = $cd.Text
    $c.Checked   = $true
    $c.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.75)
    $c.ForeColor = ConvertTo-UIColor $Theme.TextMain
    $cardRemChecks.Controls.Add($c)
    if ($cd.Hint) { $tipRemAdv.SetToolTip($c, $cd.Hint) }
    $script:ChkRemChecks[$cd.Key] = $c
}

$cardRemImports = New-CardPanel -Title "Fichiers d'import dédiés (facultatifs — sinon, meilleur effort via l'API)" -X 24 -Y 222 -W 880 -H 168
$tabRemAdv.Controls.Add($cardRemImports)

$remImportRows = @(
    @{ Label = "État BitLocker :";           TxtVar = "txtImpBitLocker"; BtnVar = "btnImpBitLocker"; Y = 42
       Hint  = "Colonnes utiles : Device name, Encryption state (encrypted/notEncrypted), motif (advancedBitLockerStates)." }
    @{ Label = "État Defender :";            TxtVar = "txtImpDefender";  BtnVar = "btnImpDefender";  Y = 72
       Hint  = "Colonnes utiles : Device name, Real-time protection (true/false), Signature update overdue (true/false)." }
    @{ Label = "Fiabilité applicative :";     TxtVar = "txtImpAppReliab"; BtnVar = "btnImpAppReliab"; Y = 102
       Hint  = "Une ligne par couple poste/application en échec. Colonnes utiles : Device name, Application name, Crash count." }
)
foreach ($row in $remImportRows) {
    $l = New-Object System.Windows.Forms.Label
    $l.Location  = New-Object System.Drawing.Point(22, ($row.Y + 3))
    $l.Size      = New-Object System.Drawing.Size(196, 22)
    $l.Text      = $row.Label
    $l.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9)
    $l.ForeColor = ConvertTo-UIColor $Theme.TextMain
    $l.BackColor = [System.Drawing.Color]::Transparent
    $cardRemImports.Controls.Add($l)

    $t = New-Object System.Windows.Forms.TextBox
    $t.Location    = New-Object System.Drawing.Point(224, $row.Y)
    $t.Size        = New-Object System.Drawing.Size(502, 25)
    $t.Font        = New-Object System.Drawing.Font($Theme.FontFamily, 8.75)
    $t.BorderStyle = "FixedSingle"
    $cardRemImports.Controls.Add($t)
    $tipRemAdv.SetToolTip($t, $row.Hint)

    $b = New-ThemedButton -Text "Parcourir..." -X 736 -Y ($row.Y - 1) -W 118 -H 28 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.TextMain -FontSize 8.5 -Ghost
    $b.Tag = $t
    $b.Add_Click({
        param($s, $e)
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Exports Intune (*.csv;*.json;*.tsv;*.txt)|*.csv;*.json;*.tsv;*.txt|Tous les fichiers (*.*)|*.*"
        if (Test-Path $ImportFolder) { $dlg.InitialDirectory = $ImportFolder }
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $s.Tag.Text = $dlg.FileName }
    })
    $cardRemImports.Controls.Add($b)
    Set-Variable -Name $row.TxtVar -Value $t -Scope Script
    Set-Variable -Name $row.BtnVar -Value $b -Scope Script
}

$lblRemImportNote           = New-Object System.Windows.Forms.Label
$lblRemImportNote.Location  = New-Object System.Drawing.Point(22, 134)
$lblRemImportNote.Size      = New-Object System.Drawing.Size(836, 26)
$lblRemImportNote.Text      = "Mises à jour Windows, profils de configuration, batteries et uptime restent en API seule (meilleur effort). Sans fichier ici, ces vérifications tentent l'API si le mode de collecte le permet."
$lblRemImportNote.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.5, [System.Drawing.FontStyle]::Italic)
$lblRemImportNote.ForeColor = ConvertTo-UIColor $Theme.TextMuted
$lblRemImportNote.BackColor = [System.Drawing.Color]::Transparent
$cardRemImports.Controls.Add($lblRemImportNote)

$cardRemActions = New-CardPanel -Title "Actions directes depuis le rapport" -X 24 -Y 398 -W 880 -H 150
$tabRemAdv.Controls.Add($cardRemActions)

$lblRemActionsNote           = New-Object System.Windows.Forms.Label
$lblRemActionsNote.Location  = New-Object System.Drawing.Point(22, 42)
$lblRemActionsNote.Size      = New-Object System.Drawing.Size(836, 96)
$lblRemActionsNote.Text      = (
    "Le rapport reste un fichier HTML statique (consultable hors ligne) : il ne peut pas appeler Microsoft Graph " +
    "lui-même. Les boutons Sync / Reboot / Remédier, en face de chaque poste de l'onglet 4, copient dans le " +
    "presse-papiers une commande PowerShell prête à coller — à exécuter dans une session déjà connectée " +
    "(fonction Invoke-DeviceRemoteAction, incluse dans ce script).`n`n" +
    "Permission requise pour ces actions : DeviceManagementManagedDevices.PrivilegedOperations.All — nettement " +
    "plus sensible que le reste du script (lecture seule). Réservez-la à une application dédiée à cet usage."
)
$lblRemActionsNote.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 8.75)
$lblRemActionsNote.ForeColor = ConvertTo-UIColor $Theme.TextMain
$lblRemActionsNote.BackColor = [System.Drawing.Color]::Transparent
$cardRemActions.Controls.Add($lblRemActionsNote)

# ========================================
# BOUTONS DE GÉNÉRATION & BARRE DE STATUT
# ========================================

$btnGenerateOnly = New-ThemedButton -Text "Générer le Dashboard" -X 262 -Y 702 -W 210 -H 46 -BackHex $Theme.Success -HoverHex $Theme.SuccessHover
$btnGenerateOnly.Add_Click({ Generate-Dashboard -OpenAfterGeneration $false })
$form.Controls.Add($btnGenerateOnly)

$btnGenerateOpen = New-ThemedButton -Text "Générer et Ouvrir" -X 488 -Y 702 -W 230 -H 46 -BackHex $Theme.Accent -HoverHex $Theme.AccentHover
$btnGenerateOpen.Add_Click({ Generate-Dashboard -OpenAfterGeneration $true })
$form.Controls.Add($btnGenerateOpen)

$lblStatus           = New-Object System.Windows.Forms.Label
$lblStatus.Location  = New-Object System.Drawing.Point(24, 762)
$lblStatus.Size      = New-Object System.Drawing.Size(776, 26)
$lblStatus.Text      = "Sélectionnez un client et configurez votre dashboard"
$lblStatus.Font      = New-Object System.Drawing.Font($Theme.FontFamily, 9, [System.Drawing.FontStyle]::Italic)
$lblStatus.ForeColor = ConvertTo-UIColor $Theme.TextMuted
$lblStatus.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
$lblStatus.BackColor = [System.Drawing.Color]::Transparent
$form.Controls.Add($lblStatus)

$btnViewLog = New-ThemedButton -Text "Voir le journal" -X 816 -Y 758 -W 140 -H 32 -BackHex $Theme.CardBack -HoverHex $Theme.GhostHover -ForeHex $Theme.TextMain -FontSize 8.75 -Ghost
$btnViewLog.Add_Click({
    try {
        if (-not (Test-Path $LogFile)) {
            Show-InfoMessage "Aucun journal n'existe encore ($LogFile).`n`nLancez d'abord une génération."
            return
        }
        Start-Process -FilePath "notepad.exe" -ArgumentList "`"$LogFile`""
    } catch {
        Show-ErrorMessage "Impossible d'ouvrir le journal : $($_.Exception.Message)`n`nChemin : $LogFile"
    }
})
$form.Controls.Add($btnViewLog)

# ========================================
# LANCEMENT
# ========================================

Write-Log "Application démarrée." -Level INFO
Load-ClientConfigs
[void]$form.ShowDialog()