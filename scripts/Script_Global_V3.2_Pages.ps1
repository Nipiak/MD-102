# ============================================================
# Script : Intune Global Dashboard Generator - v3.2 (UI premium + navigation par pages)
# Description : Génère un dashboard HTML interactif pour visualiser 
#               l'état des appareils Intune (conformité, chiffrement, 
#               applications, update rings, hardware)
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
# Auteur : ECONOCOM
# ============================================================

# ===== IMPORTS ET ASSEMBLIES =====
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ===== CONFIGURATION GLOBALE =====
$ConfigFolder = "C:\temp\clients-id"
$OutputFolder = "C:\temp"

# ===== [MODIF v3] SEUIL D'INACTIVITÉ - SECTION "NON ENCRYPTED" =====
# Nombre de jours sans synchronisation Intune au-delà duquel un poste est
# considéré comme inactif dans la section "Non Encrypted" du dashboard.
# Utilisé par le filtre (case à cocher) et le compteur dynamique.
$NonEncryptedInactiveDays = 30

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
    param([string]$Icon, [string]$Title, [string]$Text = "", [ValidateSet('success', 'neutral')][string]$Tone = 'neutral')
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
.ix-empty__title { font-size: 16px; font-weight: 600; color: var(--ix-ink); }
.ix-empty__text { max-width: 60ch; font-size: 13px; color: var(--ix-muted); }

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

function Anonymize-DeviceData {
    param([Parameter(Mandatory=$true)]$DeviceList)
    return $DeviceList | ForEach-Object {
        [PSCustomObject]@{
            DeviceName              = "Poste-" + ([guid]::NewGuid().ToString().Substring(0, 8))
            UserPrincipalName       = "User-"  + ([guid]::NewGuid().ToString().Substring(0, 8))
            OperatingSystem         = $_.OperatingSystem
            Manufacturer            = $_.Manufacturer
            Model                   = $_.Model
            OSVersion               = $_.OSVersion
            ComplianceState         = $_.ComplianceState
            IsEncrypted             = $_.IsEncrypted
            LastSyncDateTime        = $_.LastSyncDateTime
            EnrolledDateTime        = $_.EnrolledDateTime
            FreeStorageSpaceInBytes = $_.FreeStorageSpaceInBytes
            # [MODIF v3] Conservation des champs d'analyse "Non Encrypted".
            # Ils valent $null pour les listes qui ne les possèdent pas :
            # sans impact sur les autres tables (colonnes non sélectionnées).
            RootCause               = $_.RootCause
            DaysInactive            = $_.DaysInactive
        }
    }
}

function Get-GraphPagedResults {
    param([string]$Url, [string]$AccessToken)
    $results = @(); $next = $Url
    while ($next) {
        # [MODIF v3] Gestion basique du throttling Graph (HTTP 429) : nécessaire
        # car l'analyse Root Cause génère un appel API par poste non chiffré.
        try {
            $resp = Invoke-RestMethod -Method GET -Uri $next -Headers @{ Authorization = "Bearer $AccessToken" }
        } catch {
            $statusCode = 0
            try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
            if ($statusCode -eq 429) {
                $retryAfter = 5
                try { $retryAfter = [int]$_.Exception.Response.Headers['Retry-After'] } catch { }
                if (-not $retryAfter -or $retryAfter -le 0) { $retryAfter = 5 }
                Start-Sleep -Seconds $retryAfter
                continue   # on retente la même URL après la pause
            }
            throw
        }
        if ($resp.value) { $results += $resp.value }
        $next = $resp.'@odata.nextLink'
    }
    return $results
}

# ========================================
# [MODIF v3] ANALYSE DES CAUSES DE NON-CONFORMITÉ (ROOT CAUSE)
# ========================================

# [MODIF v3] Traduit un paramètre de conformité Intune (settingName brut,
# ex: "Windows10CompliancePolicy.BitLockerEnabled") en libellé lisible
# pour la colonne "RootCause" du dashboard.
function Get-FriendlyComplianceReason {
    param([string]$Setting, [string]$SettingName)

    # On privilégie settingName (ex: "BitLockerEnabled"), sinon le chemin complet
    $raw = if (-not [string]::IsNullOrWhiteSpace($SettingName)) { $SettingName } else { $Setting }
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }

    switch -Regex ($raw) {
        'BitLocker'                                  { return "Chiffrement BitLocker désactivé" }
        'StorageRequireEncryption|RequireEncryption' { return "Chiffrement du stockage non conforme" }
        'Firewall'                                   { return "Pare-feu (Firewall) inactif" }
        'OsMinimumVersion'                           { return "Version OS obsolète (< minimum requis)" }
        'OsMaximumVersion'                           { return "Version OS non autorisée (> maximum)" }
        'ValidOperatingSystemBuildRanges'            { return "Build OS hors plage autorisée" }
        'Defender|AntiVirus|AntiSpyware'             { return "Antivirus / Defender non conforme" }
        'RtpEnabled'                                 { return "Protection temps réel désactivée" }
        'SignatureOutOfDate'                         { return "Signatures antivirus obsolètes" }
        'SecureBoot'                                 { return "Secure Boot désactivé" }
        'Tpm'                                        { return "TPM requis absent ou désactivé" }
        'CodeIntegrity'                              { return "Intégrité du code non conforme" }
        'Password'                                   { return "Stratégie de mot de passe non conforme" }
        'DeviceThreatProtection'                     { return "Niveau de menace appareil trop élevé" }
        'RequireRemainContact'                       { return "Perte de contact Intune (poste inactif)" }
        'RequireDeviceCompliancePolicyAssigned'      { return "Aucune stratégie de conformité assignée" }
        'RequireUserExistence'                       { return "Utilisateur inscrit introuvable" }
        'Jailbroken|Rooted'                          { return "Appareil jailbreaké / rooté" }
        default                                      { return $raw }   # libellé brut si paramètre inconnu
    }
}

# [MODIF v3] Récupère, pour un appareil donné, la liste des raisons précises de
# non-conformité via l'API Graph :
#   1) /managedDevices/{id}/deviceCompliancePolicyStates          (état par stratégie)
#   2) .../deviceCompliancePolicyStates/{policyId}/settingStates  (détail par paramètre)
# Endpoint beta requis pour obtenir le détail settingStates de manière fiable.
# Retourne un tableau de libellés, causes liées au chiffrement en premier.
function Get-DeviceNonComplianceReasons {
    param(
        [Parameter(Mandatory)][string]$DeviceId,
        [Parameter(Mandatory)][string]$AccessToken
    )
    $reasons = [System.Collections.Generic.List[string]]::new()
    try {
        # 1) États des stratégies de conformité de l'appareil
        $polUri       = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/deviceCompliancePolicyStates"
        $policyStates = Get-GraphPagedResults -Url $polUri -AccessToken $AccessToken

        foreach ($pol in $policyStates) {
            # On ne détaille que les stratégies en écart (nonCompliant / error / conflict)
            if ($pol.state -notin @('nonCompliant', 'error', 'conflict')) { continue }

            # 2) Détail des paramètres (settingStates) de la stratégie en écart
            $settingStates = @()
            try {
                $ssUri         = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/deviceCompliancePolicyStates/$($pol.id)/settingStates"
                $settingStates = Get-GraphPagedResults -Url $ssUri -AccessToken $AccessToken
            } catch { }

            $badSettings = @($settingStates | Where-Object { $_.state -in @('nonCompliant', 'error', 'conflict') })

            if ($badSettings.Count -gt 0) {
                foreach ($s in $badSettings) {
                    $label = Get-FriendlyComplianceReason -Setting $s.setting -SettingName $s.settingName
                    if ($label -and -not $reasons.Contains($label)) { [void]$reasons.Add($label) }
                }
            } else {
                # Pas de granularité disponible : on remonte au moins la stratégie fautive
                $label = "Non conforme : $($pol.displayName)"
                if (-not $reasons.Contains($label)) { [void]$reasons.Add($label) }
            }
        }
    } catch {
        # Une erreur API sur un poste ne doit pas interrompre la génération du dashboard
    }

    # Priorité d'affichage : causes liées au chiffrement en premier
    return @($reasons | Sort-Object { if ($_ -match 'BitLocker|Chiffrement') { 0 } else { 1 } })
}

# ========================================
# INSTALLATION DES MODULES
# ========================================

function Install-RequiredModules {
    $lblStatus.Text = "🔍 Vérification des modules PowerShell..."; $form.Refresh()
    $missingModules = @()
    if (-not (Get-Module -Name PSWriteHTML    -ListAvailable)) { $missingModules += "PSWriteHTML" }
    if (-not (Get-Module -Name Microsoft.Graph -ListAvailable)) { $missingModules += "Microsoft.Graph" }
    if ($missingModules.Count -eq 0) { return $true }

    $lblStatus.Text = "⚠ Modules manquants détectés"; $lblStatus.ForeColor = [System.Drawing.Color]::Red; $form.Refresh()
    $moduleList = $missingModules -join ", "
    $message = "Modules PowerShell manquants : $moduleList`n`nCes modules sont nécessaires pour générer le dashboard.`n`nVoulez-vous les installer maintenant ?`n`nNote : L'installation peut prendre quelques minutes."
    $result = [System.Windows.Forms.MessageBox]::Show($message, "Installation des modules requis", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        $lblStatus.Text = "📦 Installation des modules en cours..."; $lblStatus.ForeColor = [System.Drawing.Color]::Orange; $form.Refresh()
        try {
            foreach ($module in $missingModules) {
                $lblStatus.Text = "📦 Installation de $module... (patientez)"; $form.Refresh()
                Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser -ErrorAction Stop
                $lblStatus.Text = "✓ $module installé avec succès"; $lblStatus.ForeColor = [System.Drawing.Color]::Green; $form.Refresh()
                Start-Sleep -Seconds 1
            }
            Show-InfoMessage "Modules installés avec succès !`n`nVous pouvez maintenant générer le dashboard."
            $lblStatus.Text = "✓ Modules installés - Prêt à générer"; $lblStatus.ForeColor = [System.Drawing.Color]::Green
            return $true
        } catch {
            $lblStatus.Text = "❌ Erreur lors de l'installation"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
            Show-ErrorMessage "Erreur lors de l'installation :`n`n$($_.Exception.Message)`n`nInstallez manuellement :`nInstall-Module -Name $moduleList -Force -AllowClobber -Scope CurrentUser"
            return $false
        }
    } else {
        $lblStatus.Text = "❌ Installation annulée"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
        Show-InfoMessage "Installation annulée.`n`nPour installer manuellement :`nInstall-Module -Name $($missingModules -join ',') -Force -AllowClobber -Scope CurrentUser"
        return $false
    }
}

# ========================================
# GÉNÉRATION DU DASHBOARD
# ========================================

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

    $lblStatus.Text = "⏳ Génération du dashboard pour $ClientName..."
    $lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
    $form.Refresh()

    try {
        if (-not (Install-RequiredModules)) { return }

        $lblStatus.Text = "📚 Chargement des modules PowerShell..."; $form.Refresh()
        try {
            Import-Module PSWriteHTML    -ErrorAction Stop
            Import-Module Microsoft.Graph -ErrorAction Stop
        } catch {
            $lblStatus.Text = "❌ Erreur de chargement des modules"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
            Show-ErrorMessage "Erreur lors du chargement des modules :`n`n$($_.Exception.Message)"; return
        }

        # ===== CONNEXION GRAPH =====
        $lblStatus.Text = "🔐 Connexion à Microsoft Graph..."; $form.Refresh()
        $Scope   = "https://graph.microsoft.com/.default"
        $AuthUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $Body    = @{ client_id = $ClientId; scope = $Scope; client_secret = $ClientSecret; grant_type = "client_credentials" }
        $Connection  = Invoke-RestMethod -Method POST -Uri $AuthUrl -Body $Body -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
        $AccessToken = $Connection.access_token
        $SecureToken = ConvertTo-SecureString -String $AccessToken -AsPlainText -Force
        Connect-MgGraph -AccessToken $SecureToken -NoWelcome

        # ===== RÉCUPÉRATION DES APPAREILS =====
        $lblStatus.Text = "💻 Récupération des appareils gérés..."; $form.Refresh()
        $ManagedDevices = Get-MgDeviceManagementManagedDevice -All
        $ExcludeVM      = $chkExcludeVM.Checked
        $AnonymizeData  = $chkAnonymize.Checked

        $DevicesScope = if ($ExcludeVM) {
            $ManagedDevices | Where-Object {
                -not (
                    ($_.Manufacturer -match 'VMware|innotek|VirtualBox|Parallels|QEMU') -or
                    ($_.Model        -match 'Virtual|VMware|VirtualBox|Hyper-V|Parallels|QEMU|KVM')
                )
            }
        } else { $ManagedDevices }

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

        $ShowLowStorage = $chkLowStorage.Checked

        # ===== APPLICATIONS =====
        if ($chkApplications.Checked) {
            $lblStatus.Text = "📱 Récupération des applications..."; $form.Refresh()
            $ManagedApps = Get-MgDeviceAppManagementMobileApp -All
        }

        # ===== CLASSIFICATION PAR PLATEFORME =====
        $lblStatus.Text = "🔍 Analyse des types d'appareils..."; $form.Refresh()
        $WindowsDevices = @($DevicesScope | Where-Object { $_.OperatingSystem -like "Windows*" })
        $iOSDevices     = @($DevicesScope | Where-Object { $_.OperatingSystem -like "iOS*" })
        $AndroidDevices = @($DevicesScope | Where-Object { $_.OperatingSystem -like "Android*" })
        $MacDevices     = @($DevicesScope | Where-Object { $_.OperatingSystem -like "macOS*" -or $_.OperatingSystem -like "Mac*" })

        # ===== CONFORMITÉ =====
        if ($chkCompliance.Checked) {
            $lblStatus.Text = "✓ Analyse de la conformité..."; $form.Refresh()
            $CompliantDevices    = @($DevicesScope | Where-Object ComplianceState -EQ "compliant")
            $NoncompliantDevices = @($DevicesScope | Where-Object ComplianceState -EQ "noncompliant")
        }

        # ===== APPAREILS INACTIFS (paliers dynamiques) =====
        $InactiveDevicesByThreshold = @{}
        if ($InactiveThresholds.Count -gt 0) {
            $lblStatus.Text = "⏰ Détection des appareils inactifs..."; $form.Refresh()
            foreach ($days in $InactiveThresholds) {
                $cutoff = (Get-Date).AddDays(-$days)
                $InactiveDevicesByThreshold["$days"] = @($DevicesScope | Where-Object { $_.LastSyncDateTime -lt $cutoff })
            }
        }

        # ===== LOW STORAGE =====
        if ($ShowLowStorage) {
            $MinimumFreeSpace  = 100
            $LowStorageDevices = @($DevicesScope | Where-Object { ($_.FreeStorageSpaceInBytes / 1GB) -lt $MinimumFreeSpace })
        }

        # ===== CHIFFREMENT =====
        if ($chkEncryption.Checked) {
            $lblStatus.Text = "🔒 Vérification du chiffrement..."; $form.Refresh()
            $EncryptedDevices  = @($DevicesScope | Where-Object IsEncrypted -EQ $true)
            $UnecryptedDevices = @($DevicesScope | Where-Object IsEncrypted -EQ $false)

            # ===== [MODIF v3] ENRICHISSEMENT "NON ENCRYPTED" : ROOT CAUSE + INACTIVITÉ =====
            # Pour chaque poste non chiffré :
            #  - calcul du nombre de jours d'inactivité (lastSyncDateTime, déjà présent
            #    dans les objets renvoyés par Get-MgDeviceManagementManagedDevice)
            #  - récupération des raisons exactes de non-conformité via l'API Graph
            #    (deviceCompliancePolicyStates + settingStates)
            $NonEncryptedEnriched = @()
            $idx    = 0
            $nowUtc = (Get-Date).ToUniversalTime()   # lastSyncDateTime est en UTC

            foreach ($dev in $UnecryptedDevices) {
                $idx++
                if ($idx -eq 1 -or $idx % 5 -eq 0 -or $idx -eq $UnecryptedDevices.Count) {
                    $lblStatus.Text = "🔎 Analyse des causes racines (Non Encrypted) : $idx / $($UnecryptedDevices.Count)..."
                    $form.Refresh()
                }

                # --- Jours d'inactivité (null = jamais synchronisé => traité comme inactif) ---
                $daysInactive = if ($dev.LastSyncDateTime) {
                    [int][math]::Floor(($nowUtc - $dev.LastSyncDateTime).TotalDays)
                } else { $null }

                # --- Raisons précises de non-conformité (appels Graph) ---
                $reasons = Get-DeviceNonComplianceReasons -DeviceId $dev.Id -AccessToken $AccessToken

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

                $NonEncryptedEnriched += [PSCustomObject]@{
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
                }
            }

            # Tri : postes actifs en premier (les null = jamais vus, à la fin)
            $NonEncryptedEnriched = @($NonEncryptedEnriched | Sort-Object @{ Expression = { if ($null -eq $_.DaysInactive) { 999999 } else { $_.DaysInactive } } })
        }

        # ===== INVENTAIRE =====
        $ManagedDevicesTable = $DevicesScope |
            Select-Object DeviceName, UserPrincipalName, OperatingSystem, Manufacturer, Model, OSVersion, ComplianceState, IsEncrypted, LastSyncDateTime, EnrolledDateTime |
            Sort-Object -Descending ComplianceState

        # ===== ANONYMISATION =====
        if ($AnonymizeData) {
            $lblStatus.Text = "🔐 Anonymisation des données sensibles..."; $form.Refresh()
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
            if ($ShowLowStorage)          { $LowStorageDevices  = Anonymize-DeviceData -DeviceList $LowStorageDevices }
        }

        # ===== APPLICATIONS - TABLE =====
        if ($chkApplications.Checked) {
            $ManagedAppTable = $ManagedApps |
                Select-Object DisplayName, Publisher, Id, CreatedDateTime |
                Sort-Object -Descending CreatedDateTime
        }

        # ===== HARDWARE =====
        if ($chkHardware.Checked -and $WindowsDevices.Count -gt 0) {
            $OSVersions = $WindowsDevices | Group-Object OSVersion | Sort-Object Count -Descending | Select-Object Name, Count -First 10
            $Model      = $WindowsDevices | Group-Object { "$($_.Manufacturer), $($_.Model)" } | Sort-Object Count -Descending | Select-Object Name, Count -First 10
        }

        # ===== ÉCHECS APPS =====
        if ($chkApplications.Checked) {
            $lblStatus.Text = "📱 Analyse des échecs d'installation..."; $form.Refresh()
            $FailedAppsAll = @()
            try {
                $FailedAppsUri = "https://graph.microsoft.com/beta/deviceManagement/reports/getFailedMobileAppsReport"
                $FailedBody    = @{ top = 5000; orderBy = @("FailedDeviceCount desc") } | ConvertTo-Json
                $FailedResp    = Invoke-RestMethod -Method POST -Uri $FailedAppsUri -Headers @{ Authorization = "Bearer $AccessToken" } -Body $FailedBody -ContentType "application/json" -ErrorAction Stop
                if ($FailedResp.Values) {
                    $schema = $FailedResp.Schema
                    foreach ($row in $FailedResp.Values) {
                        $obj = [ordered]@{}
                        for ($i = 0; $i -lt $schema.Count; $i++) { $obj[$schema[$i].Column] = $row[$i] }
                        $FailedAppsAll += [PSCustomObject]$obj
                    }
                }
            } catch {}

            $FailedAppsTop10       = $FailedAppsAll | Sort-Object -Property FailedDeviceCount -Descending | Select-Object -First 10
            $LatestFailedAppsTop10 = @()
            $ManagedAppsSorted     = $ManagedApps | Sort-Object -Property CreatedDateTime -Descending
            foreach ($app in $ManagedAppsSorted) {
                $failed = $FailedAppsAll | Where-Object { $_.DisplayName -eq $app.DisplayName } | Select-Object -First 1
                if ($failed) {
                    $LatestFailedAppsTop10 += [PSCustomObject]@{
                        DisplayName       = $failed.DisplayName
                        FailedDeviceCount = $failed.FailedDeviceCount
                        CreatedDateTime   = $app.CreatedDateTime
                    }
                }
                if ($LatestFailedAppsTop10.Count -ge 10) { break }
            }
            $AppsForFailedDonut = if ($LatestFailedAppsTop10.Count -gt 0) { $LatestFailedAppsTop10 } else { $FailedAppsTop10 }
        }

        # ===== UPDATE RINGS =====
        if ($chkUpdateRings.Checked) {
            $lblStatus.Text = "🔄 Analyse des Windows Update Rings..."; $form.Refresh()
            $UpdateRingsSummary = @(); $UpdateRingsDevices = @()
            try {
                $allConfigs         = Get-MgDeviceManagementDeviceConfiguration -All
                $UpdateRingPolicies = $allConfigs | Where-Object {
                    $_.ODataType -like "*windowsUpdateForBusinessConfiguration*" -or
                    $_.AdditionalProperties.'@odata.type' -like "*windowsUpdateForBusinessConfiguration*"
                }
                foreach ($pol in $UpdateRingPolicies) {
                    $statusUri      = "https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations/$($pol.Id)/deviceStatuses?`$top=1000"
                    $devStatuses    = Get-GraphPagedResults -Url $statusUri -AccessToken $AccessToken
                    if (-not $devStatuses -or $devStatuses.Count -eq 0) { continue }
                    $devUserStatuses = $devStatuses | Where-Object {
                        $_.UserName -and $_.UserName.Trim() -ne "" -and $_.UserName -ne "System account" -and $_.UserName -match "@"
                    }
                    if ($devUserStatuses.Count -eq 0) { continue }
                    $UpdateRingsSummary += [PSCustomObject]@{
                        RingName      = $pol.DisplayName
                        DeviceCount   = $devUserStatuses.Count
                        Succeeded     = ($devUserStatuses | Where-Object Status -eq "compliant").Count
                        Error         = ($devUserStatuses | Where-Object Status -eq "error").Count
                        Conflict      = ($devUserStatuses | Where-Object Status -eq "conflict").Count
                        NotApplicable = ($devUserStatuses | Where-Object Status -eq "notApplicable").Count
                        InProgress    = ($devUserStatuses | Where-Object Status -eq "inProgress").Count
                    }
                    foreach ($ds in $devUserStatuses) {
                        $UpdateRingsDevices += [PSCustomObject]@{
                            RingName     = $pol.DisplayName
                            DeviceName   = $ds.DeviceDisplayName
                            UserName     = $ds.UserName
                            Status       = $ds.Status
                            LastReported = $ds.LastReportedDateTime
                        }
                    }
                }
                if ($AnonymizeData -and $UpdateRingsDevices.Count -gt 0) {
                    $UpdateRingsDevices = $UpdateRingsDevices | ForEach-Object {
                        [PSCustomObject]@{
                            RingName     = $_.RingName
                            DeviceName   = "Poste-" + ([guid]::NewGuid().ToString().Substring(0, 8))
                            UserName     = "User-"  + ([guid]::NewGuid().ToString().Substring(0, 8))
                            Status       = $_.Status
                            LastReported = $_.LastReported
                        }
                    }
                }
                $UpdateRingsSummary = $UpdateRingsSummary | Sort-Object RingName
                $UpdateRingsDevices = $UpdateRingsDevices | Sort-Object RingName, DeviceName
            } catch {}
        }

        # ===== GÉNÉRATION HTML =====
        $lblStatus.Text = "📄 Génération du fichier HTML..."; $form.Refresh()
        $Timestamp        = (Get-Date).ToString("yyyy-MM-dd_HHmmss")
        $AnonymizedSuffix = if ($AnonymizeData) { "_ANONYMIZED" } else { "" }
        $ReportFileName   = "Intune-Dashboard_${ClientName}${AnonymizedSuffix}_${Timestamp}.html"

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

        # --- [v3.2] Cartes inactivité + stockage (déplacées vers la page Optimisation du parc) ---
        $op = [System.Text.StringBuilder]::new()
        foreach ($days in $InactiveThresholds) {
            $count = $InactiveDevicesByThreshold["$days"].Count
            [void]$op.Append((New-IxStatCard -Label "Inactifs $days+ jours" -Value $count -Icon 'clock' -Color $InactiveColorMap["$days"] -Caption "<b>$(Get-IxPercent $count $TotalDevices) %</b> sans synchronisation"))
        }
        if ($ShowLowStorage) {
            [void]$op.Append((New-IxStatCard -Label 'Low storage' -Value $LowStorageDevices.Count -Icon 'hard-drive' -Color $Colors.Warning -Caption 'Espace libre &lt; 100 GB'))
        }
        $OptimHtml = if ($op.Length -gt 0) { "<div class=`"ix-root ix-overview`">$($op.ToString())</div>" } else { "" }

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

        # ============================================================
        # [UI v3.2] PAGES VIRTUELLES
        # Une page n'est créée (et son onglet affiché) que si au moins une
        # des sections qu'elle regroupe est cochée dans l'interface.
        # ============================================================
        $HasSecurity = $chkCompliance.Checked -or $chkEncryption.Checked
        $HasDeploy   = $chkUpdateRings.Checked -or $chkApplications.Checked
        $HasOptim    = ($InactiveThresholds.Count -gt 0) -or $ShowLowStorage

        $IxPages = [System.Collections.Generic.List[object]]::new()
        $IxPages.Add([PSCustomObject]@{ Id = 'vue-ensemble'; Icon = 'layout'; Label = "Vue d'ensemble"; Hint = 'Taille et composition du parc : plateformes, modèles et versions de Windows' })
        if ($HasSecurity) { $IxPages.Add([PSCustomObject]@{ Id = 'securite';     Icon = 'shield';  Label = 'Sécurité & conformité';       Hint = 'Conformité Intune et chiffrement BitLocker des appareils' }) }
        if ($HasDeploy)   { $IxPages.Add([PSCustomObject]@{ Id = 'deploiement';  Icon = 'package'; Label = 'Mises à jour & applications'; Hint = 'Windows Update Rings et déploiement des applications' }) }
        if ($HasOptim)    { $IxPages.Add([PSCustomObject]@{ Id = 'optimisation'; Icon = 'gauge';   Label = 'Optimisation du parc';        Hint = 'Appareils inactifs et espace disque disponible' }) }

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
            # PAGE 4 : OPTIMISATION DU PARC — inactivité et stockage
            # ============================================================
            if ($HasOptim) {
                Get-IxPageStart -Page $IxPage['optimisation']

                # [v3.2] Cartes inactivité / stockage (auparavant dans "Devices Overview")
                New-HTMLSection -HeaderText "Inactive Devices & Storage" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.Primary -CanCollapse {
                    New-HTMLPanel {
                        New-HTMLText -Text $OptimHtml -FontSize 1
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

                if ($ShowLowStorage) {
                    New-HTMLSection -HeaderText "Devices with Low Storage (< 100 GB)" -HeaderTextSize 14 -HeaderBackGroundColor $Colors.DetailTables -CanCollapse -Collapsed {
                        New-HTMLTable -DataTable ($LowStorageDevices | Select-Object DeviceName, UserPrincipalName, OperatingSystem, Manufacturer, Model, OSVersion, ComplianceState, IsEncrypted, LastSyncDateTime, EnrolledDateTime, @{Name="FreeSpaceGB";Expression={[math]::Round($_.FreeStorageSpaceInBytes / 1GB, 2)}} | Sort-Object FreeSpaceGB) -Filtering -PagingLength 50
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
        Show-InfoMessage $message

    } catch {
        $lblStatus.Text = "❌ Erreur lors de la génération"; $lblStatus.ForeColor = [System.Drawing.Color]::Red
        Show-ErrorMessage "Erreur lors de la génération du dashboard :`n`n$($_.Exception.Message)"
    }
}

# ========================================
# INTERFACE GRAPHIQUE
# ========================================

$form = New-Object System.Windows.Forms.Form
$form.Text            = "Intune Dashboard v3.2"
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
$lblTitle.Text      = "Intune Dashboard  —  v3.2"
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
$grpMainSections.Size     = New-Object System.Drawing.Size(430, 185)
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

# Sections supplémentaires (vides - déplacées dans Affichage)
$grpExtraSections          = New-Object System.Windows.Forms.GroupBox
$grpExtraSections.Location = New-Object System.Drawing.Point(480, 48)
$grpExtraSections.Size     = New-Object System.Drawing.Size(430, 185)
$grpExtraSections.Text     = " Info "
$grpExtraSections.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabContent.Controls.Add($grpExtraSections)

$lblExtraInfo           = New-Object System.Windows.Forms.Label
$lblExtraInfo.Location  = New-Object System.Drawing.Point(20, 35)
$lblExtraInfo.Size      = New-Object System.Drawing.Size(395, 60)
$lblExtraInfo.Text      = "Les options d'affichage des plateformes, des appareils inactifs et du Low Storage se configurent dans l'onglet « Affichage »."
$lblExtraInfo.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblExtraInfo.ForeColor = [System.Drawing.Color]::FromArgb(0, 120, 212)
$grpExtraSections.Controls.Add($lblExtraInfo)

# Boutons tout sélectionner / tout désélectionner
$btnSelectAll           = New-Object System.Windows.Forms.Button
$btnSelectAll.Location  = New-Object System.Drawing.Point(30, 250)
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
    $chkHardware.Checked     = $true
})
$tabContent.Controls.Add($btnSelectAll)

$btnDeselectAll           = New-Object System.Windows.Forms.Button
$btnDeselectAll.Location  = New-Object System.Drawing.Point(245, 250)
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
    $chkHardware.Checked     = $false
})
$tabContent.Controls.Add($btnDeselectAll)

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
$lblAffichageInfo.Text      = "Choisissez quelles données afficher dans la section 'Devices Overview'"
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

# --- Groupe Low Storage ---
$grpLowStorage          = New-Object System.Windows.Forms.GroupBox
$grpLowStorage.Location = New-Object System.Drawing.Point(30, 190)
$grpLowStorage.Size     = New-Object System.Drawing.Size(430, 65)
$grpLowStorage.Text     = " Low Storage "
$grpLowStorage.Font     = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
$tabAffichage.Controls.Add($grpLowStorage)

$chkLowStorage          = New-Object System.Windows.Forms.CheckBox
$chkLowStorage.Location = New-Object System.Drawing.Point(20, 28); $chkLowStorage.Size = New-Object System.Drawing.Size(395, 25)
$chkLowStorage.Text     = "💾  Afficher les appareils avec peu d'espace disque (< 100 GB)"
$chkLowStorage.Checked  = $true; $chkLowStorage.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$grpLowStorage.Controls.Add($chkLowStorage)

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