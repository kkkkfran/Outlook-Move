#requires -Version 5.1
param(
    [switch]$SelfTest,
    [switch]$RunJob,

    [string]$DataFilePath,
    [string]$SourceStoreName,
    [string]$SourceFolderPath,
    [switch]$IncludeSubfolders,
    [switch]$RemoveEmptySourceFolders,
    [string]$DestinationParentPath,
    [string]$BackupName,
    [int]$Year,
    [switch]$Execute,
    [string]$LogPath
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = "Stop"

$scriptBasePath = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }
if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $LogPath = Join-Path $scriptBasePath "respaldo-correos-log.csv"
}
$configPath = Join-Path $scriptBasePath "respaldo-correos-config.json"

$script:OlMailItemClass = 43
$script:OlMailItemType = 0
$script:LogRows = New-Object System.Collections.Generic.List[object]
$script:CandidateCount = 0
$script:TotalCandidateCount = 0
$script:ProcessedCandidateCount = 0
$script:MovedCount = 0
$script:ErrorCount = 0
$script:ScannedFolderCount = 0
$script:SkippedFolderCount = 0
$script:DestinationFolderCreatedCount = 0
$script:DestinationFolderReusedCount = 0
$script:RemovedSourceFolderCount = 0
$script:KeptSourceFolderCount = 0
$script:PreferredSourceFolderPath = ""
$script:PreferredDestinationParentPath = ""

function Get-OutlookSession {
    $outlook = $null
    $errors = New-Object System.Collections.Generic.List[string]

    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
    }
    catch {
        $errors.Add("No se encontró una instancia activa: $($_.Exception.Message)") | Out-Null
    }

    if ($null -eq $outlook) {
        try {
            $outlook = New-Object -ComObject Outlook.Application
        }
        catch {
            $errors.Add("No se pudo iniciar Outlook por COM: $($_.Exception.Message)") | Out-Null
        }
    }

    if ($null -eq $outlook) {
        $details = $errors -join " | "
        throw "No se pudo conectar con Outlook clásico. Abra Outlook Classic manualmente, espere a que cargue el buzón, cierre cualquier ventana de perfil/actualización, y ejecute esta herramienta con los mismos permisos que Outlook. Detalle técnico: $details"
    }

    $namespace = $outlook.GetNamespace("MAPI")
    try {
        $namespace.Logon($null, $null, $false, $false)
    }
    catch { }

    [pscustomobject]@{
        Application = $outlook
        Namespace   = $namespace
    }
}

function Test-OutlookComConnection {
    try {
        $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        if ($null -ne $outlook) {
            return $true
        }
    }
    catch { }

    return $false
}

function Get-OutlookProcessStatusText {
    $processes = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    if ($processes.Count -eq 0) {
        return "No hay procesos OUTLOOK.EXE abiertos."
    }

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($process in $processes) {
        $title = ""
        try { $title = [string]$process.MainWindowTitle } catch { }
        if ([string]::IsNullOrWhiteSpace($title)) {
            $title = "(sin ventana principal)"
        }
        $lines.Add(("OUTLOOK.EXE PID {0}: {1}" -f $process.Id, $title)) | Out-Null
    }

    $lines -join [Environment]::NewLine
}

function Find-OutlookExecutablePath {
    $registryPaths = @(
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
        "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
        "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE"
    )

    foreach ($path in $registryPaths) {
        try {
            $value = (Get-ItemProperty -LiteralPath $path -ErrorAction Stop)."(default)"
            if (-not [string]::IsNullOrWhiteSpace($value) -and (Test-Path -LiteralPath $value)) {
                return $value
            }
        }
        catch { }
    }

    $commonPaths = @(
        "$env:ProgramFiles\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\OUTLOOK.EXE",
        "$env:ProgramFiles\Microsoft Office\Office16\OUTLOOK.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\Office16\OUTLOOK.EXE"
    )

    foreach ($path in $commonPaths) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            return $path
        }
    }

    return "outlook.exe"
}

function Start-OutlookClassicAndWait {
    param([int]$TimeoutSeconds = 45)

    $outlookPath = Find-OutlookExecutablePath
    Start-Process -FilePath $outlookPath | Out-Null

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds 2
        if (Test-OutlookComConnection) {
            return $true
        }
    }

    return $false
}

function Open-OutlookDataFile {
    param(
        $Namespace,
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "No se encontró el archivo PST: $Path"
    }

    $resolvedPath = (Resolve-Path -LiteralPath $Path).Path
    foreach ($store in $Namespace.Stores) {
        try {
            $storePath = [string]$store.FilePath
            if (-not [string]::IsNullOrWhiteSpace($storePath) -and $storePath -ieq $resolvedPath) {
                Write-Host "PST ya abierto: $resolvedPath"
                return
            }
        }
        catch { }
    }

    try {
        $Namespace.AddStoreEx($resolvedPath, 3)
    }
    catch {
        $Namespace.AddStore($resolvedPath)
    }

    Write-Host "PST abierto: $resolvedPath"
}

function Test-DestinationStoreAvailable {
    param(
        $Namespace,
        [string]$Path
    )

    $parts = @(Split-OutlookPath -Path $Path)
    if ($parts.Count -lt 1) {
        return $false
    }

    $store = Find-Store -Namespace $Namespace -StoreName $parts[0]
    $null -ne $store
}

function Split-OutlookPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return @()
    }

    $normalized = $Path.Replace("/", "\").Trim().Trim("\")
    @($normalized -split "\\" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
}

function Find-Store {
    param(
        $Namespace,
        [string]$StoreName
    )

    foreach ($store in $Namespace.Stores) {
        if ($store.DisplayName -ieq $StoreName) {
            return $store
        }
    }

    return $null
}

function Find-ChildFolder {
    param(
        $ParentFolder,
        [string]$Name
    )

    foreach ($folder in $ParentFolder.Folders) {
        if ($folder.Name -ieq $Name) {
            return $folder
        }
    }

    return $null
}

function Resolve-OutlookFolderPath {
    param(
        $Namespace,
        [string]$Path,
        [switch]$Create
    )

    $parts = @(Split-OutlookPath -Path $Path)
    if ($parts.Count -lt 1) {
        throw "La ruta destino está vacía."
    }

    $store = Find-Store -Namespace $Namespace -StoreName $parts[0]
    if ($null -eq $store) {
        throw "No se encontró el almacén de Outlook '$($parts[0])'. Presione 'Cargar Outlook' en la ventana y elija una ruta de la lista."
    }

    $current = Get-StoreRootFolder -Namespace $Namespace -Store $store
    for ($i = 1; $i -lt $parts.Count; $i++) {
        $next = Find-ChildFolder -ParentFolder $current -Name $parts[$i]
        if ($null -eq $next) {
            if (-not $Create) {
                throw "No se encontró la carpeta '$($parts[$i])' dentro de '$($current.FolderPath)'."
            }
            $next = $current.Folders.Add($parts[$i])
        }
        $current = $next
    }

    [pscustomobject]@{
        Store  = $store
        Folder = $current
        Parts  = $parts
    }
}

function Get-StoreRootFolder {
    param(
        $Namespace,
        $Store
    )

    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $root = $null
        try {
            $root = $Store.GetRootFolder()
        }
        catch { }

        if ($null -eq $root) {
            try {
                $root = $Namespace.Folders.Item([string]$Store.DisplayName)
            }
            catch { }
        }

        if ($null -ne $root) {
            return $root
        }

        Start-Sleep -Milliseconds 500
    }

    throw "Outlook no entregó la carpeta raíz para '$($Store.DisplayName)'. Cierre y vuelva a abrir el PST o presione 'Reparar Outlook'."
}

function Get-DefaultSourceStore {
    param(
        $Namespace,
        $DestinationStore
    )

    if (-not [string]::IsNullOrWhiteSpace($SourceStoreName)) {
        $store = Find-Store -Namespace $Namespace -StoreName $SourceStoreName
        if ($null -eq $store) {
            throw "No se encontró el buzón origen '$SourceStoreName'."
        }
        return $store
    }

    try {
        if ($null -ne $Namespace.DefaultStore) {
            return $Namespace.DefaultStore
        }
    }
    catch { }

    foreach ($store in $Namespace.Stores) {
        if ($store.StoreID -ne $DestinationStore.StoreID) {
            return $store
        }
    }

    foreach ($store in $Namespace.Stores) {
        return $store
    }

    throw "No hay buzones/almacenes de Outlook disponibles."
}

function Get-OrCreateBackupFolder {
    param(
        $ParentFolder,
        [string]$Name
    )

    if ($null -eq $ParentFolder) {
        throw "No se pudo resolver la carpeta destino. Vuelva a presionar 'Cargar Outlook' y seleccione el PST en el árbol."
    }

    if ($ParentFolder.Name -ieq $Name) {
        return $ParentFolder
    }

    $existing = Find-ChildFolder -ParentFolder $ParentFolder -Name $Name
    if ($null -ne $existing) {
        return $existing
    }

    $folders = $null
    try {
        $folders = $ParentFolder.Folders
    }
    catch { }

    if ($null -eq $folders) {
        throw "La carpeta seleccionada no permite crear subcarpetas por COM. Seleccione una carpeta dentro del PST, por ejemplo 'Bandeja de entrada', o vuelva a cargar Outlook."
    }

    try {
        return $folders.Add($Name, $script:OlMailItemType)
    }
    catch {
        return $folders.Add($Name)
    }
}

function Get-NormalizedFolderPath {
    param($Folder)

    ([string]$Folder.FolderPath).Replace("/", "\").TrimEnd("\").ToLowerInvariant()
}

function Test-FolderIsUnder {
    param(
        $Folder,
        $ParentFolder
    )

    $folderPath = Get-NormalizedFolderPath -Folder $Folder
    $parentPath = Get-NormalizedFolderPath -Folder $ParentFolder

    $folderPath -eq $parentPath -or $folderPath.StartsWith($parentPath + "\")
}

function Test-ExcludedFolderName {
    param([string]$Name)

    $excluded = @(
        "Deleted Items",
        "Elementos eliminados",
        "Junk Email",
        "Correo no deseado",
        "Correo electronico no deseado",
        "Drafts",
        "Borradores",
        "Outbox",
        "Bandeja de salida",
        "Sync Issues",
        "Problemas de sincronizacion",
        "RSS Feeds",
        "Fuentes RSS"
    )

    foreach ($item in $excluded) {
        if ($Name -ieq $item) {
            return $true
        }
    }

    return $false
}

function Get-MailDate {
    param($Item)

    try {
        $received = [datetime]$Item.ReceivedTime
        if ($received.Year -gt 1900) {
            return $received
        }
    }
    catch { }

    try {
        $sent = [datetime]$Item.SentOn
        if ($sent.Year -gt 1900) {
            return $sent
        }
    }
    catch { }

    try {
        $created = [datetime]$Item.CreationTime
        if ($created.Year -gt 1900) {
            return $created
        }
    }
    catch { }

    return $null
}

function Get-MailSnapshot {
    param($Item)

    $subject = ""
    $mailDate = ""
    $entryId = ""

    try { $subject = [string]$Item.Subject } catch { }
    try {
        $date = Get-MailDate -Item $Item
        if ($null -ne $date) {
            $mailDate = $date.ToString("s")
        }
    }
    catch { }
    try { $entryId = [string]$Item.EntryID } catch { }

    [pscustomobject]@{
        Subject = $subject
        Date    = $mailDate
        EntryID = $entryId
    }
}

function Add-LogRow {
    param(
        [string]$Action,
        [string]$SourcePath,
        [string]$DestinationPath,
        $Snapshot,
        [string]$Message
    )

    $script:LogRows.Add([pscustomobject]@{
        Time        = (Get-Date).ToString("s")
        Action      = $Action
        Source      = $SourcePath
        Destination = $DestinationPath
        Subject     = $Snapshot.Subject
        Date        = $Snapshot.Date
        EntryID     = $Snapshot.EntryID
        Message     = $Message
    }) | Out-Null
}

function Get-MailFolders {
    param(
        $Folder,
        $BackupFolder,
        [bool]$IncludeSubfolders = $true,
        [bool]$IsSelectedRoot = $true,
        [bool]$SkipExcludedNames = $true
    )

    if (Test-FolderIsUnder -Folder $Folder -ParentFolder $BackupFolder) {
        $script:SkippedFolderCount++
        return
    }

    if ($SkipExcludedNames -and (-not $IsSelectedRoot) -and (Test-ExcludedFolderName -Name $Folder.Name)) {
        $script:SkippedFolderCount++
        return
    }

    $isMailFolder = $true
    try {
        $isMailFolder = ([int]$Folder.DefaultItemType -eq $script:OlMailItemType)
    }
    catch {
        $isMailFolder = $true
    }

    if ($isMailFolder) {
        $Folder
    }

    if (-not $IncludeSubfolders) {
        return
    }

    foreach ($child in $Folder.Folders) {
        foreach ($entry in Get-MailFolders -Folder $child -BackupFolder $BackupFolder -IncludeSubfolders $true -IsSelectedRoot $false -SkipExcludedNames $SkipExcludedNames) {
            $entry
        }
    }
}

function Get-SourceFolderTree {
    param(
        $Folder,
        [bool]$IncludeSubfolders = $true
    )

    $Folder

    if (-not $IncludeSubfolders) {
        return
    }

    foreach ($child in $Folder.Folders) {
        foreach ($entry in Get-SourceFolderTree -Folder $child -IncludeSubfolders $true) {
            $entry
        }
    }
}

function Test-OutlookMailFolder {
    param($Folder)

    try {
        return ([int]$Folder.DefaultItemType -eq $script:OlMailItemType)
    }
    catch {
        return $true
    }
}

function Get-RelativeFolderParts {
    param(
        $RootFolder,
        $Folder
    )

    $rootParts = @(Split-OutlookPath -Path ([string]$RootFolder.FolderPath))
    $folderParts = @(Split-OutlookPath -Path ([string]$Folder.FolderPath))

    if ($folderParts.Count -lt $rootParts.Count) {
        throw "La carpeta '$($Folder.FolderPath)' no está dentro de '$($RootFolder.FolderPath)'."
    }

    for ($i = 0; $i -lt $rootParts.Count; $i++) {
        if ($folderParts[$i] -ine $rootParts[$i]) {
            throw "La carpeta '$($Folder.FolderPath)' no está dentro de '$($RootFolder.FolderPath)'."
        }
    }

    if ($folderParts.Count -eq $rootParts.Count) {
        return @()
    }

    @($folderParts[$rootParts.Count..($folderParts.Count - 1)])
}

function Join-OutlookPath {
    param(
        [string]$BasePath,
        [string[]]$Parts
    )

    $cleanBasePath = $BasePath.Replace("/", "\").TrimEnd("\")
    if ($null -eq $Parts -or $Parts.Count -eq 0) {
        return $cleanBasePath
    }

    $suffix = (@($Parts) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) -join "\"
    if ([string]::IsNullOrWhiteSpace($suffix)) {
        return $cleanBasePath
    }

    "$cleanBasePath\$suffix"
}

function Get-OrCreateChildFolderPath {
    param(
        $ParentFolder,
        [string[]]$Parts,
        [string]$PathPrefix = ""
    )

    $current = $ParentFolder
    $currentPath = if ([string]::IsNullOrWhiteSpace($PathPrefix)) { [string]$ParentFolder.FolderPath } else { $PathPrefix }
    foreach ($part in @($Parts)) {
        if ([string]::IsNullOrWhiteSpace($part)) {
            continue
        }

        $nextPath = Join-OutlookPath -BasePath $currentPath -Parts @($part)
        $existing = Find-ChildFolder -ParentFolder $current -Name $part
        if ($null -ne $existing) {
            $script:DestinationFolderReusedCount++
            Write-Host ("Carpeta destino existente: {0}" -f $nextPath)
            $current = $existing
        }
        else {
            $folders = $null
            try {
                $folders = $current.Folders
            }
            catch { }

            if ($null -eq $folders) {
                throw "La carpeta destino '$currentPath' no permite crear subcarpetas."
            }

            try {
                $current = $folders.Add($part, $script:OlMailItemType)
            }
            catch {
                $current = $folders.Add($part)
            }
            $script:DestinationFolderCreatedCount++
            Write-Host ("Carpeta destino creada: {0}" -f $nextPath)
        }

        $currentPath = $nextPath
    }

    $current
}

function Test-FolderIsEmpty {
    param($Folder)

    $itemCount = 0
    $childCount = 0
    try { $itemCount = [int]$Folder.Items.Count } catch { $itemCount = 1 }
    try { $childCount = [int]$Folder.Folders.Count } catch { $childCount = 1 }

    $itemCount -eq 0 -and $childCount -eq 0
}

function Remove-EmptySourceFolders {
    param(
        [System.Collections.Generic.List[object]]$FolderPlans,
        $SelectedRootFolder
    )

    for ($i = $FolderPlans.Count - 1; $i -ge 0; $i--) {
        $plan = $FolderPlans[$i]
        $folder = $plan.SourceFolder
        $sourcePath = ""
        try { $sourcePath = [string]$folder.FolderPath } catch { $sourcePath = [string]$plan.SourcePath }

        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            continue
        }

        try {
            if (Test-FolderIsEmpty -Folder $folder) {
                $folder.Delete()
                $script:RemovedSourceFolderCount++
                $emptySnapshot = [pscustomobject]@{ Subject = ""; Date = ""; EntryID = "" }
                Add-LogRow -Action "FolderRemoved" -SourcePath $sourcePath -DestinationPath $plan.DestinationPath -Snapshot $emptySnapshot -Message "Carpeta origen vacía eliminada"
                Write-Host ("Carpeta origen vacía eliminada: {0}" -f $sourcePath)
            }
            else {
                $script:KeptSourceFolderCount++
                Write-Host ("Carpeta origen conservada con contenido: {0}" -f $sourcePath)
            }
        }
        catch {
            $script:ErrorCount++
            $emptySnapshot = [pscustomobject]@{ Subject = ""; Date = ""; EntryID = "" }
            Add-LogRow -Action "Error" -SourcePath $sourcePath -DestinationPath $plan.DestinationPath -Snapshot $emptySnapshot -Message $_.Exception.Message
            Write-Host ("No se pudo quitar carpeta origen '{0}': {1}" -f $sourcePath, $_.Exception.Message)
        }
    }
}

function Test-MailItemMatchesYear {
    param(
        $Item,
        [int]$TargetYear
    )

    try {
        if ([int]$Item.Class -ne $script:OlMailItemClass) {
            return $false
        }
    }
    catch {
        return $false
    }

    $mailDate = Get-MailDate -Item $Item
    if ($null -eq $mailDate) {
        return $false
    }

    [int]$mailDate.Year -eq $TargetYear
}

function Count-YearMailFromFolder {
    param(
        $SourceFolder,
        [string]$DestinationPath,
        [int]$TargetYear
    )

    $sourcePath = [string]$SourceFolder.FolderPath
    if (-not (Test-OutlookMailFolder -Folder $SourceFolder)) {
        Write-Host ("Saltando carpeta no-correo: {0}" -f $sourcePath)
        return 0
    }

    try {
        $items = $SourceFolder.Items
        $count = [int]$items.Count
    }
    catch {
        $script:ErrorCount++
        $emptySnapshot = [pscustomobject]@{ Subject = ""; Date = ""; EntryID = "" }
        Add-LogRow -Action "Error" -SourcePath $sourcePath -DestinationPath $DestinationPath -Snapshot $emptySnapshot -Message $_.Exception.Message
        return 0
    }

    if ($count -eq 0) {
        return 0
    }

    Write-Host ("Contando {0} ({1} items)" -f $sourcePath, $count)
    $matches = 0
    for ($i = $count; $i -ge 1; $i--) {
        $item = $null
        try {
            $item = $items.Item($i)
        }
        catch {
            $script:ErrorCount++
            continue
        }

        if ($null -eq $item) {
            continue
        }

        if (Test-MailItemMatchesYear -Item $item -TargetYear $TargetYear) {
            $matches++
            if (($matches % 50) -eq 0) {
                Write-Host ("  Encontrados en esta carpeta: {0}" -f $matches)
            }
        }
    }

    if ($matches -gt 0) {
        Write-Host ("  Total encontrados en carpeta: {0}" -f $matches)
    }

    return $matches
}

function Move-YearMailFromFolder {
    param(
        $SourceFolder,
        $DestinationFolder,
        [string]$DestinationPath,
        [int]$TargetYear,
        [bool]$DoMove,
        [int]$TotalCandidates = 0
    )

    $script:ScannedFolderCount++
    $sourcePath = [string]$SourceFolder.FolderPath
    if (-not (Test-OutlookMailFolder -Folder $SourceFolder)) {
        Write-Host ("Saltando carpeta no-correo: {0}" -f $sourcePath)
        return
    }

    try {
        $items = $SourceFolder.Items
        $count = [int]$items.Count
    }
    catch {
        $script:ErrorCount++
        $emptySnapshot = [pscustomobject]@{ Subject = ""; Date = ""; EntryID = "" }
        Add-LogRow -Action "Error" -SourcePath $sourcePath -DestinationPath $DestinationPath -Snapshot $emptySnapshot -Message $_.Exception.Message
        return
    }

    if ($count -eq 0) {
        return
    }

    Write-Host ("Revisando {0} ({1} items)" -f $sourcePath, $count)

    for ($i = $count; $i -ge 1; $i--) {
        $item = $null
        try {
            $item = $items.Item($i)
        }
        catch {
            $script:ErrorCount++
            continue
        }

        if ($null -eq $item) {
            continue
        }

        if (-not (Test-MailItemMatchesYear -Item $item -TargetYear $TargetYear)) {
            continue
        }

        $snapshot = Get-MailSnapshot -Item $item

        if ($DoMove) {
            $script:ProcessedCandidateCount++
            try {
                if ($null -eq $DestinationFolder) {
                    throw "No se resolvió la carpeta destino '$DestinationPath'."
                }
                $null = $item.Move($DestinationFolder)
                $script:MovedCount++
                Add-LogRow -Action "Moved" -SourcePath $sourcePath -DestinationPath $DestinationPath -Snapshot $snapshot -Message "OK"
            }
            catch {
                $script:ErrorCount++
                Add-LogRow -Action "Error" -SourcePath $sourcePath -DestinationPath $DestinationPath -Snapshot $snapshot -Message $_.Exception.Message
            }

            $processed = $script:ProcessedCandidateCount
            $remaining = [Math]::Max(0, $TotalCandidates - $processed)
            $percent = if ($TotalCandidates -gt 0) { [int][Math]::Floor(($processed / [double]$TotalCandidates) * 100) } else { 100 }
            if ($TotalCandidates -le 20 -or ($processed % 10) -eq 0 -or $processed -eq $TotalCandidates) {
                Write-Host ("Progreso: {0}/{1} procesados ({2}%). Movidos: {3}. Errores: {4}. Faltan: {5}." -f $processed, $TotalCandidates, $percent, $script:MovedCount, $script:ErrorCount, $remaining)
            }
        }
        else {
            $script:CandidateCount++
            Add-LogRow -Action "WouldMove" -SourcePath $sourcePath -DestinationPath $DestinationPath -Snapshot $snapshot -Message "Simulation"
        }
    }
}

function Invoke-BackupJob {
    if ($Year -lt 1900 -or $Year -gt 3000) {
        throw "Indique un año válido."
    }
    if ([string]::IsNullOrWhiteSpace($BackupName)) {
        throw "Indique el nombre del respaldo. Ejemplo: Jose Sevilla - $Year"
    }
    if ([string]::IsNullOrWhiteSpace($DestinationParentPath)) {
        throw "Indique donde dejar el respaldo."
    }

    $session = Get-OutlookSession
    try {
        Open-OutlookDataFile -Namespace $session.Namespace -Path $DataFilePath
    }
    catch {
        if (Test-DestinationStoreAvailable -Namespace $session.Namespace -Path $DestinationParentPath) {
            Write-Host "Aviso: Outlook no permitió abrir el PST otra vez, pero el destino ya está disponible. Se continúa."
            Write-Host ("Detalle PST: {0}" -f $_.Exception.Message)
        }
        else {
            throw
        }
    }

    $destination = Resolve-OutlookFolderPath -Namespace $session.Namespace -Path $DestinationParentPath -Create
    $backupFolder = Get-OrCreateBackupFolder -ParentFolder $destination.Folder -Name $BackupName

    $sourceFolder = $null
    $sourceStore = $null
    $includeChildren = $true
    if (-not [string]::IsNullOrWhiteSpace($SourceFolderPath)) {
        $resolvedSource = Resolve-OutlookFolderPath -Namespace $session.Namespace -Path $SourceFolderPath
        $sourceStore = $resolvedSource.Store
        $sourceFolder = $resolvedSource.Folder
        $includeChildren = $true
    }
    else {
        $sourceStore = Get-DefaultSourceStore -Namespace $session.Namespace -DestinationStore $destination.Store
        $sourceFolder = $sourceStore.GetRootFolder()
        $includeChildren = $true
    }

    if ((Get-NormalizedFolderPath -Folder $sourceFolder) -eq (Get-NormalizedFolderPath -Folder $backupFolder)) {
        throw "La carpeta origen y la carpeta destino son la misma. Seleccione una carpeta destino distinta."
    }

    if ($sourceStore.StoreID -eq $destination.Store.StoreID) {
        Write-Host "Aviso: origen y destino están en el mismo buzón/almacén. Esto organiza correos, pero no libera espacio real del buzón."
    }

    $mode = if ($Execute) { "EJECUCION REAL" } else { "SIMULACION" }

    Write-Host "Modo: $mode"
    Write-Host "Año: $Year"
    Write-Host "Origen: $($sourceStore.DisplayName)"
    Write-Host "Carpeta origen: $($sourceFolder.FolderPath)"
    Write-Host ("Incluye subcarpetas y mantiene estructura: {0}" -f ($(if ($includeChildren) { "Si" } else { "No" })))
    Write-Host ("Quita carpetas origen vacías: {0}" -f ($(if ($RemoveEmptySourceFolders) { "Si" } else { "No" })))
    Write-Host "Destino: $($backupFolder.FolderPath)"
    Write-Host ""

    $destinationStructureRootFolder = $backupFolder
    $destinationStructureRootPath = [string]$backupFolder.FolderPath
    if (-not [string]::IsNullOrWhiteSpace($SourceFolderPath)) {
        $sourceRootName = [string]$sourceFolder.Name
        if ([string]::IsNullOrWhiteSpace($sourceRootName)) {
            $sourceRootParts = @(Split-OutlookPath -Path $SourceFolderPath)
            if ($sourceRootParts.Count -gt 0) {
                $sourceRootName = $sourceRootParts[$sourceRootParts.Count - 1]
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($sourceRootName) -and $backupFolder.Name -ine $sourceRootName) {
            $destinationStructureRootPath = Join-OutlookPath -BasePath ([string]$backupFolder.FolderPath) -Parts @($sourceRootName)
            if ($Execute) {
                $destinationStructureRootFolder = Get-OrCreateChildFolderPath -ParentFolder $backupFolder -Parts @($sourceRootName) -PathPrefix ([string]$backupFolder.FolderPath)
            }
        }

        Write-Host "Carpeta raíz en destino: $destinationStructureRootPath"
        Write-Host ""
    }

    $folders = New-Object System.Collections.Generic.List[object]

    if (-not [string]::IsNullOrWhiteSpace($SourceFolderPath)) {
        foreach ($treeFolder in Get-SourceFolderTree -Folder $sourceFolder -IncludeSubfolders $true) {
            if (Test-FolderIsUnder -Folder $treeFolder -ParentFolder $backupFolder) {
                $script:SkippedFolderCount++
                continue
            }
            $folders.Add($treeFolder) | Out-Null
        }
    }
    else {
        foreach ($folder in $sourceFolder.Folders) {
            foreach ($mailFolder in Get-MailFolders -Folder $folder -BackupFolder $backupFolder -IncludeSubfolders $true -IsSelectedRoot $false -SkipExcludedNames $true) {
                $folders.Add($mailFolder) | Out-Null
            }
        }
    }

    Write-Host ("Carpetas a revisar: {0}" -f $folders.Count)

    $folderPlans = New-Object System.Collections.Generic.List[object]
    foreach ($folder in $folders) {
        $relativeParts = @(if (-not [string]::IsNullOrWhiteSpace($SourceFolderPath)) {
            Get-RelativeFolderParts -RootFolder $sourceFolder -Folder $folder
        }
        else {
            [string]$folder.Name
        })

        $destinationPath = Join-OutlookPath -BasePath $destinationStructureRootPath -Parts $relativeParts
        $destinationFolder = $null
        if ($Execute) {
            $destinationFolder = Get-OrCreateChildFolderPath -ParentFolder $destinationStructureRootFolder -Parts $relativeParts -PathPrefix $destinationStructureRootPath
        }

        $relativeText = if ($relativeParts.Count -gt 0) { $relativeParts -join "\" } else { "." }
        Write-Host ("Mapa: {0} -> {1}" -f $relativeText, $destinationPath)
        $folderPlans.Add([pscustomobject]@{
            SourceFolder      = $folder
            SourcePath        = [string]$folder.FolderPath
            DestinationFolder = $destinationFolder
            DestinationPath   = $destinationPath
            RelativePath      = $relativeText
        }) | Out-Null
    }

    if ($Execute) {
        Write-Host "Contando correos que coinciden antes de mover..."
        foreach ($plan in $folderPlans) {
            $script:TotalCandidateCount += Count-YearMailFromFolder -SourceFolder $plan.SourceFolder -DestinationPath $plan.DestinationPath -TargetYear $Year
        }
        Write-Host ("Total a mover: {0}" -f $script:TotalCandidateCount)
        Write-Host ""
    }

    foreach ($plan in $folderPlans) {
        Move-YearMailFromFolder -SourceFolder $plan.SourceFolder -DestinationFolder $plan.DestinationFolder -DestinationPath $plan.DestinationPath -TargetYear $Year -DoMove ([bool]$Execute) -TotalCandidates $script:TotalCandidateCount
    }

    if ($Execute -and $RemoveEmptySourceFolders) {
        Write-Host ""
        Write-Host "Quitando carpetas origen que quedaron vacías..."
        Remove-EmptySourceFolders -FolderPlans $folderPlans -SelectedRootFolder $sourceFolder
    }

    $logDir = Split-Path -Path $LogPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($logDir) -and -not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    $script:LogRows | Export-Csv -Path $LogPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Resumen"
    Write-Host ("  Carpetas revisadas: {0}" -f $script:ScannedFolderCount)
    Write-Host ("  Carpetas omitidas: {0}" -f $script:SkippedFolderCount)
    Write-Host ("  Carpetas destino creadas: {0}" -f $script:DestinationFolderCreatedCount)
    Write-Host ("  Carpetas destino existentes/usadas: {0}" -f $script:DestinationFolderReusedCount)
    Write-Host ("  Carpetas origen vacías quitadas: {0}" -f $script:RemovedSourceFolderCount)
    Write-Host ("  Carpetas origen conservadas: {0}" -f $script:KeptSourceFolderCount)
    $summaryCandidateCount = if ($Execute) { $script:TotalCandidateCount } else { $script:CandidateCount }
    Write-Host ("  Correos del año {0}: {1}" -f $Year, $summaryCandidateCount)
    Write-Host ("  Correos movidos: {0}" -f $script:MovedCount)
    Write-Host ("  Errores: {0}" -f $script:ErrorCount)
    Write-Host ("  Log: {0}" -f $LogPath)

    if (-not $Execute) {
        Write-Host ""
        Write-Host "No se movió nada. Revise el log y luego use 'Mover correos'."
    }
}

if ($RunJob) {
    Invoke-BackupJob
    return
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

function New-UiFont {
    param(
        [float]$Size = 9.0,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )

    New-Object System.Drawing.Font("Segoe UI", $Size, $Style)
}

function ConvertTo-CommandLineArgument {
    param([string]$Value)

    if ($null -eq $Value) {
        return '""'
    }

    '"' + ($Value -replace '"', '\"') + '"'
}

function Add-OutputLine {
    param([string]$Text)

    $outputBox.AppendText($Text + [Environment]::NewLine)
    $outputBox.SelectionStart = $outputBox.TextLength
    $outputBox.ScrollToCaret()
}

function Set-RunningState {
    param([bool]$Running)

    $progress.Visible = $Running
    $progress.Style = if ($Running) { [System.Windows.Forms.ProgressBarStyle]::Marquee } else { [System.Windows.Forms.ProgressBarStyle]::Blocks }
    if ($Running) {
        $progress.Value = 0
        $progressLabel.Text = "Preparando..."
    }
    else {
        $progressLabel.Text = ""
    }
    foreach ($button in @($openPstButton, $refreshButton, $repairButton, $diagnoseButton, $simulateButton, $moveButton, $openLogButton)) {
        $button.Enabled = -not $Running
    }
    foreach ($control in @($yearNumeric, $backupNameTextBox, $pstPathTextBox, $sourceStoreComboBox, $sourceTree, $destinationTree, $removeEmptySourceFoldersCheckBox)) {
        $control.Enabled = -not $Running
    }
}

function Update-ProgressFromOutputLine {
    param([string]$Text)

    if ($Text -match '^Total a mover:\s+(\d+)') {
        $total = [int]$Matches[1]
        $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
        $progress.Value = 0
        $progressLabel.Text = "Total a mover: $total"
        return
    }

    if ($Text -match '^Progreso:\s+(\d+)/(\d+)\s+procesados\s+\((\d+)%\)\.\s+Movidos:\s+(\d+)\.\s+Errores:\s+(\d+)\.\s+Faltan:\s+(\d+)\.') {
        $processed = [int]$Matches[1]
        $total = [int]$Matches[2]
        $percent = [Math]::Min(100, [Math]::Max(0, [int]$Matches[3]))
        $moved = [int]$Matches[4]
        $errors = [int]$Matches[5]
        $remaining = [int]$Matches[6]

        $progress.Style = [System.Windows.Forms.ProgressBarStyle]::Blocks
        $progress.Value = $percent
        $progressLabel.Text = ("{0}/{1} procesados | Movidos: {2} | Faltan: {3} | Errores: {4}" -f $processed, $total, $moved, $remaining, $errors)
    }
}

function Show-OutlookDiagnostics {
    Add-OutputLine ""
    Add-OutputLine "Diagnostico Outlook"
    Add-OutputLine (Get-OutlookProcessStatusText)
    if (Test-OutlookComConnection) {
        Add-OutputLine "COM Outlook: disponible."
    }
    else {
        Add-OutputLine "COM Outlook: no disponible."
    }
}

function Repair-OutlookConnection {
    $message = "Esto intentará cerrar Outlook Classic y abrirlo de nuevo para reparar la conexión COM.`r`n`r`nGuarda correos/borradores abiertos antes de continuar.`r`n`r`n¿Quieres continuar?"
    $choice = [System.Windows.Forms.MessageBox]::Show($message, "Reparar conexion Outlook", "YesNo", "Warning")
    if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
        return
    }

    Set-RunningState -Running $true
    try {
        Add-OutputLine "Reparando conexion con Outlook..."
        Show-OutlookDiagnostics

        $processes = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
        foreach ($process in $processes) {
            try {
                Add-OutputLine ("Cerrando Outlook PID {0}..." -f $process.Id)
                if ($process.MainWindowHandle -ne 0) {
                    $null = $process.CloseMainWindow()
                }
            }
            catch {
                Add-OutputLine ("No se pudo solicitar cierre de PID {0}: {1}" -f $process.Id, $_.Exception.Message)
            }
        }

        $deadline = (Get-Date).AddSeconds(20)
        while ((Get-Date) -lt $deadline -and @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0) {
            Start-Sleep -Seconds 1
        }

        $remaining = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
        if ($remaining.Count -gt 0) {
            $force = [System.Windows.Forms.MessageBox]::Show("Outlook no se cerró solo. ¿Quieres forzar el cierre de OUTLOOK.EXE?", "Outlook sigue abierto", "YesNo", "Warning")
            if ($force -eq [System.Windows.Forms.DialogResult]::Yes) {
                foreach ($process in $remaining) {
                    try {
                        Add-OutputLine ("Forzando cierre de Outlook PID {0}..." -f $process.Id)
                        $process.Kill()
                    }
                    catch {
                        Add-OutputLine ("No se pudo forzar PID {0}: {1}" -f $process.Id, $_.Exception.Message)
                    }
                }
                Start-Sleep -Seconds 3
            }
            else {
                Add-OutputLine "Reparacion cancelada porque Outlook sigue abierto."
                return
            }
        }

        Add-OutputLine "Abriendo Outlook Classic..."
        if (Start-OutlookClassicAndWait -TimeoutSeconds 60) {
            Add-OutputLine "Outlook respondio por COM."
            Refresh-OutlookTree
        }
        else {
            Add-OutputLine "Outlook se abrió, pero COM aún no responde. Revisa si hay ventanas de perfil, contraseña o inicio de sesión."
            Show-OutlookDiagnostics
        }
    }
    catch {
        Add-OutputLine "Error reparando Outlook: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message, "Error reparando Outlook", "OK", "Error") | Out-Null
    }
    finally {
        Set-RunningState -Running $false
    }
}

function Add-FolderNode {
    param(
        [System.Windows.Forms.TreeNode]$ParentNode,
        $Folder,
        [string]$Path,
        [int]$Depth,
        [int]$MaxDepth
    )

    if ($Depth -gt $MaxDepth) {
        return
    }

    foreach ($child in $Folder.Folders) {
        $childPath = "{0}\{1}" -f $Path, $child.Name
        $node = New-Object System.Windows.Forms.TreeNode([string]$child.Name)
        $node.Tag = $childPath
        $ParentNode.Nodes.Add($node) | Out-Null
        Add-FolderNode -ParentNode $node -Folder $child -Path $childPath -Depth ($Depth + 1) -MaxDepth $MaxDepth
    }
}

function Find-TreeNodeByTag {
    param(
        [System.Windows.Forms.TreeNodeCollection]$Nodes,
        [string]$Tag
    )

    if ([string]::IsNullOrWhiteSpace($Tag)) {
        return $null
    }

    foreach ($node in $Nodes) {
        if ([string]$node.Tag -ieq $Tag) {
            return $node
        }

        $found = Find-TreeNodeByTag -Nodes $node.Nodes -Tag $Tag
        if ($null -ne $found) {
            return $found
        }
    }

    return $null
}

function Select-TreeNodeByTag {
    param(
        [System.Windows.Forms.TreeView]$Tree,
        [string]$Tag
    )

    $node = Find-TreeNodeByTag -Nodes $Tree.Nodes -Tag $Tag
    if ($null -ne $node) {
        $Tree.SelectedNode = $node
        $node.EnsureVisible()
    }

    return $node
}

function Refresh-OutlookTree {
    Set-RunningState -Running $true
    try {
        Add-OutputLine "Cargando Outlook..."
        $session = Get-OutlookSession
        $pstOpenError = $null
        try {
            Open-OutlookDataFile -Namespace $session.Namespace -Path $pstPathTextBox.Text.Trim()
        }
        catch {
            $pstOpenError = $_.Exception.Message
            Add-OutputLine ("Aviso: Outlook no permitió abrir el PST desde archivo: {0}" -f $pstOpenError)
            Add-OutputLine "Se cargarán igual los buzones y PST que ya estén abiertos en Outlook."
        }

        $sourceStoreComboBox.Items.Clear()
        $sourceTree.Nodes.Clear()
        $destinationTree.Nodes.Clear()
        $preferredSourceNode = $null
        $preferredDestinationNode = $null
        $resolvedPstPath = ""
        if (-not [string]::IsNullOrWhiteSpace($pstPathTextBox.Text) -and (Test-Path -LiteralPath $pstPathTextBox.Text)) {
            $resolvedPstPath = (Resolve-Path -LiteralPath $pstPathTextBox.Text).Path
        }

        foreach ($store in $session.Namespace.Stores) {
            $storeName = [string]$store.DisplayName
            if ([string]::IsNullOrWhiteSpace($storeName)) {
                continue
            }

            $sourceStoreComboBox.Items.Add($storeName) | Out-Null

            $storeRootFolder = Get-StoreRootFolder -Namespace $session.Namespace -Store $store

            $sourceRootNode = New-Object System.Windows.Forms.TreeNode($storeName)
            $sourceRootNode.Tag = $storeName
            $sourceTree.Nodes.Add($sourceRootNode) | Out-Null
            Add-FolderNode -ParentNode $sourceRootNode -Folder $storeRootFolder -Path $storeName -Depth 1 -MaxDepth 6
            $sourceRootNode.Expand()

            $destinationRootNode = New-Object System.Windows.Forms.TreeNode($storeName)
            $destinationRootNode.Tag = $storeName
            $destinationTree.Nodes.Add($destinationRootNode) | Out-Null
            Add-FolderNode -ParentNode $destinationRootNode -Folder $storeRootFolder -Path $storeName -Depth 1 -MaxDepth 6
            $destinationRootNode.Expand()

            if (-not [string]::IsNullOrWhiteSpace($script:PreferredSourceFolderPath) -and $storeName -ieq (@(Split-OutlookPath -Path $script:PreferredSourceFolderPath)[0])) {
                $preferredSourceNode = Select-TreeNodeByTag -Tree $sourceTree -Tag $script:PreferredSourceFolderPath
            }

            if (-not [string]::IsNullOrWhiteSpace($resolvedPstPath)) {
                try {
                    $storePath = [string]$store.FilePath
                    if (-not [string]::IsNullOrWhiteSpace($storePath) -and $storePath -ieq $resolvedPstPath) {
                        $preferredDestinationNode = $destinationRootNode
                    }
                }
                catch { }
            }
        }

        if ($sourceStoreComboBox.Items.Count -gt 0 -and $sourceStoreComboBox.SelectedIndex -lt 0) {
            $sourceStoreComboBox.SelectedIndex = 0
        }

        if ($null -eq $preferredSourceNode -and -not [string]::IsNullOrWhiteSpace($script:PreferredSourceFolderPath)) {
            $preferredSourceNode = Select-TreeNodeByTag -Tree $sourceTree -Tag $script:PreferredSourceFolderPath
        }

        if ($null -ne $preferredSourceNode) {
            $sourceTree.SelectedNode = $preferredSourceNode
            Add-OutputLine ("Origen seleccionado: {0}" -f $preferredSourceNode.Tag)
        }

        if ($null -ne $preferredDestinationNode) {
            $destinationTree.SelectedNode = $preferredDestinationNode
            $destinationTree.Focus()
            Add-OutputLine ("Destino seleccionado automáticamente: {0}" -f $preferredDestinationNode.Tag)
        }
        elseif (-not [string]::IsNullOrWhiteSpace($script:PreferredDestinationParentPath)) {
            $preferredDestinationNode = Select-TreeNodeByTag -Tree $destinationTree -Tag $script:PreferredDestinationParentPath
            if ($null -ne $preferredDestinationNode) {
                Add-OutputLine ("Destino seleccionado: {0}" -f $preferredDestinationNode.Tag)
            }
        }
        elseif ($null -ne $pstOpenError -and -not [string]::IsNullOrWhiteSpace($resolvedPstPath)) {
            Add-OutputLine "No encontré ese PST en el árbol. Abre el PST manualmente en Outlook: Archivo > Abrir y exportar > Abrir archivo de datos de Outlook."
            Add-OutputLine "Luego vuelve a presionar 'Cargar Outlook'."
        }

        Update-SourcePreview
        Update-FinalDestinationPreview
        Add-OutputLine "Outlook cargado. Selecciona carpeta origen y carpeta base destino."
    }
    catch {
        Add-OutputLine "Error: $($_.Exception.Message)"
        Show-OutlookDiagnostics
        [System.Windows.Forms.MessageBox]::Show("No se pudo conectar con Outlook. Abre Outlook Classic y presiona 'Reparar Outlook' o 'Diagnostico'.", "Error cargando Outlook", "OK", "Error") | Out-Null
    }
    finally {
        Set-RunningState -Running $false
    }
}

function Save-Configuration {
    $selectedSourcePath = ""
    if ($null -ne $sourceTree.SelectedNode) {
        $selectedSourcePath = [string]$sourceTree.SelectedNode.Tag
    }

    $selectedPath = ""
    if ($null -ne $destinationTree.SelectedNode) {
        $selectedPath = [string]$destinationTree.SelectedNode.Tag
    }

    $config = [pscustomobject]@{
        ConfigVersion         = 2
        Year                  = [int]$yearNumeric.Value
        BackupName            = $backupNameTextBox.Text.Trim()
        DataFilePath          = $pstPathTextBox.Text.Trim()
        SourceStoreName       = [string]$sourceStoreComboBox.Text
        SourceFolderPath      = $selectedSourcePath
        IncludeSubfolders     = [bool]$includeSubfoldersCheckBox.Checked
        RemoveEmptySourceFolders = [bool]$removeEmptySourceFoldersCheckBox.Checked
        DestinationParentPath = $selectedPath
    }

    $config | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $configPath -Encoding UTF8
}

function Load-Configuration {
    if (-not (Test-Path -LiteralPath $configPath)) {
        return
    }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
        if ($null -ne $config.Year) { $yearNumeric.Value = [decimal]$config.Year }
        if ($null -ne $config.BackupName) { $backupNameTextBox.Text = [string]$config.BackupName }
        if ($null -ne $config.DataFilePath) { $pstPathTextBox.Text = [string]$config.DataFilePath }
        if ($null -ne $config.SourceStoreName) { $sourceStoreComboBox.Text = [string]$config.SourceStoreName }
        if ($null -ne $config.SourceFolderPath) { $script:PreferredSourceFolderPath = [string]$config.SourceFolderPath }
        if ($null -ne $config.DestinationParentPath) { $script:PreferredDestinationParentPath = [string]$config.DestinationParentPath }
        $includeSubfoldersCheckBox.Checked = $true
        if ($null -ne $config.RemoveEmptySourceFolders) { $removeEmptySourceFoldersCheckBox.Checked = [bool]$config.RemoveEmptySourceFolders }
    }
    catch {
        Add-OutputLine "No se pudo cargar la configuracion guardada: $($_.Exception.Message)"
    }
}

function Get-SelectedSourcePath {
    if ($null -eq $sourceTree.SelectedNode) {
        return ""
    }

    [string]$sourceTree.SelectedNode.Tag
}

function Get-SourcePathText {
    $sourcePath = Get-SelectedSourcePath
    if ([string]::IsNullOrWhiteSpace($sourcePath)) {
        return "Origen: selecciona una carpeta"
    }

    $mode = if ($includeSubfoldersCheckBox.Checked) { "con subcarpetas y estructura" } else { "solo esta carpeta" }
    "Origen: $sourcePath ($mode)"
}

function Update-SourcePreview {
    if ($null -eq $sourcePathLabel) {
        return
    }

    $sourcePathLabel.Text = Get-SourcePathText
}

function Get-SelectedDestinationPath {
    if ($null -eq $destinationTree.SelectedNode) {
        return ""
    }

    [string]$destinationTree.SelectedNode.Tag
}

function Get-FinalDestinationPathText {
    $basePath = Get-SelectedDestinationPath
    $backupName = $backupNameTextBox.Text.Trim()
    $sourcePath = Get-SelectedSourcePath

    if ([string]::IsNullOrWhiteSpace($basePath)) {
        return "Destino final: selecciona el PST o carpeta base"
    }

    if ([string]::IsNullOrWhiteSpace($backupName)) {
        return "Destino final: escribe el nombre de la carpeta destino"
    }

    $baseParts = @(Split-OutlookPath -Path $basePath)
    $baseName = if ($baseParts.Count -gt 0) { $baseParts[$baseParts.Count - 1] } else { $basePath }
    $backupPath = if ($baseName -ieq $backupName) { $basePath } else { "$basePath\$backupName" }

    $sourceParts = @(Split-OutlookPath -Path $sourcePath)
    if ($sourceParts.Count -eq 0) {
        return "Destino final: $backupPath"
    }

    $sourceName = $sourceParts[$sourceParts.Count - 1]
    $backupParts = @(Split-OutlookPath -Path $backupPath)
    $backupLeafName = if ($backupParts.Count -gt 0) { $backupParts[$backupParts.Count - 1] } else { $backupPath }
    if ($backupLeafName -ieq $sourceName) {
        return "Destino final: $backupPath"
    }

    "Destino final: $backupPath\$sourceName"
}

function Update-FinalDestinationPreview {
    if ($null -eq $finalDestinationLabel) {
        return
    }

    $finalDestinationLabel.Text = Get-FinalDestinationPathText
}

function Validate-UiInputs {
    if ([string]::IsNullOrWhiteSpace($backupNameTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Escribe el nombre de la carpeta destino. Ejemplo: Bandeja de Entrada", "Falta nombre", "OK", "Warning") | Out-Null
        return $false
    }

    if ([string]::IsNullOrWhiteSpace((Get-SelectedSourcePath))) {
        [System.Windows.Forms.MessageBox]::Show("Selecciona en el árbol la carpeta origen que quieres mover. Ejemplo: Bandeja de entrada.", "Falta origen", "OK", "Warning") | Out-Null
        return $false
    }

    if ([string]::IsNullOrWhiteSpace((Get-SelectedDestinationPath))) {
        [System.Windows.Forms.MessageBox]::Show("Selecciona en el árbol el PST o carpeta base destino. La carpeta final se llamará igual que 'Carpeta destino'.", "Falta destino", "OK", "Warning") | Out-Null
        return $false
    }

    if ([string]::IsNullOrWhiteSpace($sourceStoreComboBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Selecciona el buzón origen.", "Falta origen", "OK", "Warning") | Out-Null
        return $false
    }

    return $true
}

function Start-JobProcess {
    param([bool]$DoMove)

    if (-not (Validate-UiInputs)) {
        return
    }

    if ($DoMove) {
        $message = "Esto moverá correos reales del año $([int]$yearNumeric.Value).`r`n`r`nOrigen: $(Get-SelectedSourcePath)`r`nDestino: $(Get-FinalDestinationPathText)`r`nSubcarpetas: $(if ($includeSubfoldersCheckBox.Checked) { "se copiará la estructura" } else { "solo carpeta seleccionada" })`r`nCarpetas origen vacías: $(if ($removeEmptySourceFoldersCheckBox.Checked) { "se quitarán después de mover" } else { "se conservarán" })`r`n`r`n¿Ejecutaste una simulación y quieres continuar?"
        $choice = [System.Windows.Forms.MessageBox]::Show($message, "Confirmar movimiento real", "YesNo", "Warning")
        if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) {
            return
        }
    }

    Save-Configuration

    $args = New-Object System.Collections.Generic.List[string]
    $args.Add("-NoProfile") | Out-Null
    $args.Add("-ExecutionPolicy") | Out-Null
    $args.Add("Bypass") | Out-Null
    $args.Add("-STA") | Out-Null
    $args.Add("-File") | Out-Null
    $args.Add($PSCommandPath) | Out-Null
    $args.Add("-RunJob") | Out-Null
    $args.Add("-Year") | Out-Null
    $args.Add([string][int]$yearNumeric.Value) | Out-Null
    $args.Add("-BackupName") | Out-Null
    $args.Add($backupNameTextBox.Text.Trim()) | Out-Null
    $args.Add("-DestinationParentPath") | Out-Null
    $args.Add((Get-SelectedDestinationPath)) | Out-Null
    $args.Add("-SourceStoreName") | Out-Null
    $args.Add([string]$sourceStoreComboBox.Text) | Out-Null
    $args.Add("-SourceFolderPath") | Out-Null
    $args.Add((Get-SelectedSourcePath)) | Out-Null
    $args.Add("-LogPath") | Out-Null
    $args.Add($LogPath) | Out-Null

    if ($includeSubfoldersCheckBox.Checked) {
        $args.Add("-IncludeSubfolders") | Out-Null
    }

    if ($removeEmptySourceFoldersCheckBox.Checked) {
        $args.Add("-RemoveEmptySourceFolders") | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($pstPathTextBox.Text)) {
        $args.Add("-DataFilePath") | Out-Null
        $args.Add($pstPathTextBox.Text.Trim()) | Out-Null
    }

    if ($DoMove) {
        $args.Add("-Execute") | Out-Null
    }

    Set-RunningState -Running $true
    $outputBox.Clear()
    Add-OutputLine ($(if ($DoMove) { "== Movimiento real ==" } else { "== Simulacion ==" }))
    Add-OutputLine ""

    $stdoutPath = Join-Path $env:TEMP ("respaldo-correos-{0}.out.log" -f [guid]::NewGuid().ToString("N"))
    $stderrPath = Join-Path $env:TEMP ("respaldo-correos-{0}.err.log" -f [guid]::NewGuid().ToString("N"))
    $powerShellExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    $argumentLine = (($args.ToArray() | ForEach-Object { ConvertTo-CommandLineArgument $_ }) -join " ")
    $cmdArguments = '/d /c ""{0}" {1} > "{2}" 2> "{3}""' -f $powerShellExe, $argumentLine, $stdoutPath, $stderrPath

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $env:ComSpec
        $psi.Arguments = $cmdArguments
        $psi.UseShellExecute = $false
        $psi.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $null = $process.Start()

        Add-OutputLine "Proceso iniciado. El avance aparecera aqui en vivo..."

        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 700
        $timer.Add_Tick({
            $readNewLines = {
                param(
                    [string]$Path,
                    [string]$OffsetVariableName,
                    [bool]$IsError
                )

                if (-not (Test-Path -LiteralPath $Path)) {
                    return
                }

                $text = Get-Content -LiteralPath $Path -Raw -ErrorAction SilentlyContinue
                if ([string]::IsNullOrEmpty($text)) {
                    return
                }

                $offset = [int](Get-Variable -Name $OffsetVariableName -Scope Script -ValueOnly)
                if ($text.Length -le $offset) {
                    return
                }

                $newText = $text.Substring($offset)
                Set-Variable -Name $OffsetVariableName -Scope Script -Value $text.Length

                $newText.TrimEnd() -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
                    $line = [string]$_
                    if ($IsError) {
                        Add-OutputLine ("Error: {0}" -f $line)
                    }
                    else {
                        Add-OutputLine $line
                        Update-ProgressFromOutputLine $line
                    }
                }
            }

            try {
                & $readNewLines $script:CurrentStdOutPath "CurrentStdOutOffset" $false
                & $readNewLines $script:CurrentStdErrPath "CurrentStdErrOffset" $true

                if ($null -eq $script:CurrentProcess -or -not $script:CurrentProcess.HasExited) {
                    return
                }

                $script:CurrentTimer.Stop()
                & $readNewLines $script:CurrentStdOutPath "CurrentStdOutOffset" $false
                & $readNewLines $script:CurrentStdErrPath "CurrentStdErrOffset" $true

                Add-OutputLine ""
                Add-OutputLine ("Finalizado con codigo {0}" -f $script:CurrentProcess.ExitCode)
                if ($script:CurrentProcess.ExitCode -eq 0 -and $progress.Style -eq [System.Windows.Forms.ProgressBarStyle]::Blocks -and $progress.Value -gt 0) {
                    $progress.Value = 100
                }
            }
            catch {
                Add-OutputLine ("Error leyendo avance: {0}" -f $_.Exception.Message)
            }
            finally {
                if ($null -ne $script:CurrentProcess -and $script:CurrentProcess.HasExited) {
                    Set-RunningState -Running $false
                    try { $script:CurrentProcess.Dispose() } catch { }
                    try { $script:CurrentTimer.Dispose() } catch { }
                    try { Remove-Item -LiteralPath $script:CurrentStdOutPath, $script:CurrentStdErrPath -Force -ErrorAction SilentlyContinue } catch { }
                    $script:CurrentProcess = $null
                    $script:CurrentTimer = $null
                    $script:CurrentStdOutPath = ""
                    $script:CurrentStdErrPath = ""
                    $script:CurrentStdOutOffset = 0
                    $script:CurrentStdErrOffset = 0
                }
            }
        })

        $script:CurrentProcess = $process
        $script:CurrentTimer = $timer
        $script:CurrentStdOutPath = $stdoutPath
        $script:CurrentStdErrPath = $stderrPath
        $script:CurrentStdOutOffset = 0
        $script:CurrentStdErrOffset = 0
        $timer.Start()
    }
    catch {
        Set-RunningState -Running $false
        Add-OutputLine "Error iniciando proceso: $($_.Exception.Message)"
    }
}

function Invoke-UiAction {
    param(
        [scriptblock]$Action,
        [string]$Title = "Error"
    )

    try {
        & $Action
    }
    catch {
        $message = $_.Exception.Message
        try {
            Add-OutputLine ("Error: {0}" -f $message)
        }
        catch { }
        try {
            Set-RunningState -Running $false
        }
        catch { }
        [System.Windows.Forms.MessageBox]::Show($message, $Title, "OK", "Error") | Out-Null
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Respaldo de correos por año"
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1220, 820)
$form.MinimumSize = New-Object System.Drawing.Size(1020, 700)
$form.Font = New-UiFont
$form.BackColor = [System.Drawing.Color]::FromArgb(246, 248, 250)

$headerPanel = New-Object System.Windows.Forms.Panel
$headerPanel.Dock = "Top"
$headerPanel.Height = 82
$headerPanel.BackColor = [System.Drawing.Color]::FromArgb(28, 43, 58)

$titleLabel = New-Object System.Windows.Forms.Label
$titleLabel.Text = "Respaldo de correos por año"
$titleLabel.ForeColor = [System.Drawing.Color]::White
$titleLabel.Font = New-UiFont -Size 18 -Style Bold
$titleLabel.Location = New-Object System.Drawing.Point(24, 14)
$titleLabel.Size = New-Object System.Drawing.Size(620, 32)

$subtitleLabel = New-Object System.Windows.Forms.Label
$subtitleLabel.Text = "Elige año, carpeta origen y destino; primero simula y luego mueve con progreso visible."
$subtitleLabel.ForeColor = [System.Drawing.Color]::FromArgb(215, 225, 235)
$subtitleLabel.Location = New-Object System.Drawing.Point(26, 49)
$subtitleLabel.Size = New-Object System.Drawing.Size(840, 22)

$headerPanel.Controls.AddRange(@($titleLabel, $subtitleLabel))

$mainPanel = New-Object System.Windows.Forms.Panel
$mainPanel.Dock = "Fill"
$mainPanel.Padding = New-Object System.Windows.Forms.Padding(22)

$yearLabel = New-Object System.Windows.Forms.Label
$yearLabel.Text = "Año"
$yearLabel.Location = New-Object System.Drawing.Point(24, 20)
$yearLabel.Size = New-Object System.Drawing.Size(80, 22)
$yearLabel.Font = New-UiFont -Style Bold

$yearNumeric = New-Object System.Windows.Forms.NumericUpDown
$yearNumeric.Location = New-Object System.Drawing.Point(24, 45)
$yearNumeric.Size = New-Object System.Drawing.Size(90, 26)
$yearNumeric.Minimum = 1900
$yearNumeric.Maximum = 3000
$yearNumeric.Value = [Math]::Max(1900, (Get-Date).Year - 1)

$backupNameLabel = New-Object System.Windows.Forms.Label
$backupNameLabel.Text = "Carpeta destino"
$backupNameLabel.Location = New-Object System.Drawing.Point(140, 20)
$backupNameLabel.Size = New-Object System.Drawing.Size(180, 22)
$backupNameLabel.Font = New-UiFont -Style Bold

$backupNameTextBox = New-Object System.Windows.Forms.TextBox
$backupNameTextBox.Location = New-Object System.Drawing.Point(140, 45)
$backupNameTextBox.Size = New-Object System.Drawing.Size(290, 26)
$backupNameTextBox.Text = "Respaldo - $([int]$yearNumeric.Value)"

$pstLabel = New-Object System.Windows.Forms.Label
$pstLabel.Text = "Archivo PST destino"
$pstLabel.Location = New-Object System.Drawing.Point(455, 20)
$pstLabel.Size = New-Object System.Drawing.Size(180, 22)
$pstLabel.Font = New-UiFont -Style Bold

$pstPathTextBox = New-Object System.Windows.Forms.TextBox
$pstPathTextBox.Location = New-Object System.Drawing.Point(455, 45)
$pstPathTextBox.Size = New-Object System.Drawing.Size(225, 26)
$pstPathTextBox.Anchor = "Top,Left"

$openPstButton = New-Object System.Windows.Forms.Button
$openPstButton.Text = "Abrir PST"
$openPstButton.Location = New-Object System.Drawing.Point(695, 43)
$openPstButton.Size = New-Object System.Drawing.Size(105, 30)
$openPstButton.Anchor = "Top,Left"

$sourceLabel = New-Object System.Windows.Forms.Label
$sourceLabel.Text = "Buzon origen"
$sourceLabel.Location = New-Object System.Drawing.Point(24, 90)
$sourceLabel.Size = New-Object System.Drawing.Size(140, 22)
$sourceLabel.Font = New-UiFont -Style Bold

$sourceStoreComboBox = New-Object System.Windows.Forms.ComboBox
$sourceStoreComboBox.Location = New-Object System.Drawing.Point(24, 115)
$sourceStoreComboBox.Size = New-Object System.Drawing.Size(330, 26)
$sourceStoreComboBox.DropDownStyle = "DropDown"

$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Cargar Outlook"
$refreshButton.Location = New-Object System.Drawing.Point(375, 113)
$refreshButton.Size = New-Object System.Drawing.Size(130, 30)

$diagnoseButton = New-Object System.Windows.Forms.Button
$diagnoseButton.Text = "Diagnostico"
$diagnoseButton.Location = New-Object System.Drawing.Point(515, 113)
$diagnoseButton.Size = New-Object System.Drawing.Size(120, 30)

$repairButton = New-Object System.Windows.Forms.Button
$repairButton.Text = "Reparar Outlook"
$repairButton.Location = New-Object System.Drawing.Point(645, 113)
$repairButton.Size = New-Object System.Drawing.Size(145, 30)

$simulateButton = New-Object System.Windows.Forms.Button
$simulateButton.Text = "Simular"
$simulateButton.Location = New-Object System.Drawing.Point(24, 160)
$simulateButton.Size = New-Object System.Drawing.Size(140, 38)
$simulateButton.BackColor = [System.Drawing.Color]::FromArgb(232, 244, 255)

$moveButton = New-Object System.Windows.Forms.Button
$moveButton.Text = "Mover correos"
$moveButton.Location = New-Object System.Drawing.Point(180, 160)
$moveButton.Size = New-Object System.Drawing.Size(150, 38)
$moveButton.BackColor = [System.Drawing.Color]::FromArgb(255, 235, 235)

$openLogButton = New-Object System.Windows.Forms.Button
$openLogButton.Text = "Abrir log"
$openLogButton.Location = New-Object System.Drawing.Point(346, 160)
$openLogButton.Size = New-Object System.Drawing.Size(120, 38)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(485, 170)
$progress.Size = New-Object System.Drawing.Size(330, 18)
$progress.Anchor = "Top,Left"
$progress.Visible = $false

$progressLabel = New-Object System.Windows.Forms.Label
$progressLabel.Text = ""
$progressLabel.Location = New-Object System.Drawing.Point(830, 168)
$progressLabel.Size = New-Object System.Drawing.Size(330, 22)
$progressLabel.Anchor = "Top,Left,Right"
$progressLabel.ForeColor = [System.Drawing.Color]::FromArgb(55, 70, 85)

$sourceFolderLabel = New-Object System.Windows.Forms.Label
$sourceFolderLabel.Text = "Carpeta origen"
$sourceFolderLabel.Location = New-Object System.Drawing.Point(24, 220)
$sourceFolderLabel.Size = New-Object System.Drawing.Size(160, 22)
$sourceFolderLabel.Font = New-UiFont -Style Bold

$sourcePathLabel = New-Object System.Windows.Forms.Label
$sourcePathLabel.Text = "Origen: selecciona una carpeta"
$sourcePathLabel.Location = New-Object System.Drawing.Point(145, 221)
$sourcePathLabel.Size = New-Object System.Drawing.Size(220, 22)
$sourcePathLabel.ForeColor = [System.Drawing.Color]::FromArgb(35, 85, 130)
$sourcePathLabel.Font = New-UiFont -Style Bold

$sourceTree = New-Object System.Windows.Forms.TreeView
$sourceTree.Location = New-Object System.Drawing.Point(24, 245)
$sourceTree.Size = New-Object System.Drawing.Size(330, 330)
$sourceTree.Anchor = "Top,Bottom,Left"
$sourceTree.HideSelection = $false

$includeSubfoldersCheckBox = New-Object System.Windows.Forms.CheckBox
$includeSubfoldersCheckBox.Text = "Incluir subcarpetas y mantener estructura"
$includeSubfoldersCheckBox.Location = New-Object System.Drawing.Point(24, 585)
$includeSubfoldersCheckBox.Size = New-Object System.Drawing.Size(280, 24)
$includeSubfoldersCheckBox.Checked = $true
$includeSubfoldersCheckBox.Enabled = $false

$removeEmptySourceFoldersCheckBox = New-Object System.Windows.Forms.CheckBox
$removeEmptySourceFoldersCheckBox.Text = "Quitar carpetas origen si quedan vacías"
$removeEmptySourceFoldersCheckBox.Location = New-Object System.Drawing.Point(24, 615)
$removeEmptySourceFoldersCheckBox.Size = New-Object System.Drawing.Size(280, 24)
$removeEmptySourceFoldersCheckBox.Checked = $true

$destinationLabel = New-Object System.Windows.Forms.Label
$destinationLabel.Text = "PST o carpeta base destino"
$destinationLabel.Location = New-Object System.Drawing.Point(380, 220)
$destinationLabel.Size = New-Object System.Drawing.Size(240, 22)
$destinationLabel.Font = New-UiFont -Style Bold

$finalDestinationLabel = New-Object System.Windows.Forms.Label
$finalDestinationLabel.Text = "Destino final: selecciona el PST o carpeta base"
$finalDestinationLabel.Location = New-Object System.Drawing.Point(380, 585)
$finalDestinationLabel.Size = New-Object System.Drawing.Size(780, 22)
$finalDestinationLabel.Anchor = "Top,Left,Right"
$finalDestinationLabel.ForeColor = [System.Drawing.Color]::FromArgb(35, 85, 130)
$finalDestinationLabel.Font = New-UiFont -Style Bold

$destinationTree = New-Object System.Windows.Forms.TreeView
$destinationTree.Location = New-Object System.Drawing.Point(380, 245)
$destinationTree.Size = New-Object System.Drawing.Size(330, 330)
$destinationTree.Anchor = "Top,Bottom,Left"
$destinationTree.HideSelection = $false

$outputLabel = New-Object System.Windows.Forms.Label
$outputLabel.Text = "Salida"
$outputLabel.Location = New-Object System.Drawing.Point(735, 220)
$outputLabel.Size = New-Object System.Drawing.Size(100, 22)
$outputLabel.Font = New-UiFont -Style Bold

$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Location = New-Object System.Drawing.Point(735, 245)
$outputBox.Size = New-Object System.Drawing.Size(425, 330)
$outputBox.Anchor = "Top,Bottom,Left,Right"
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Both"
$outputBox.WordWrap = $false
$outputBox.ReadOnly = $true
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 9)
$outputBox.BackColor = [System.Drawing.Color]::White

$hintLabel = New-Object System.Windows.Forms.Label
$hintLabel.Text = "Al mover, se usan/crean las carpetas destino equivalentes. Las carpetas origen solo se quitan si quedan vacías."
$hintLabel.Location = New-Object System.Drawing.Point(24, 620)
$hintLabel.Size = New-Object System.Drawing.Size(1080, 24)
$hintLabel.Anchor = "Bottom,Left,Right"
$hintLabel.ForeColor = [System.Drawing.Color]::FromArgb(85, 92, 100)

$mainLayout = New-Object System.Windows.Forms.TableLayoutPanel
$mainLayout.Dock = "Fill"
$mainLayout.ColumnCount = 1
$mainLayout.RowCount = 4
$mainLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$mainLayout.Padding = New-Object System.Windows.Forms.Padding(0)
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 74))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 84))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$mainLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null

$topLayout = New-Object System.Windows.Forms.TableLayoutPanel
$topLayout.Dock = "Fill"
$topLayout.ColumnCount = 4
$topLayout.RowCount = 1
$topLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$topLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 100))) | Out-Null
$topLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 42))) | Out-Null
$topLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 58))) | Out-Null
$topLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 124))) | Out-Null

$yearPanel = New-Object System.Windows.Forms.Panel
$yearPanel.Dock = "Fill"
$yearLabel.Dock = "Top"
$yearLabel.Height = 22
$yearNumeric.Dock = "Top"
$yearPanel.Controls.Add($yearNumeric)
$yearPanel.Controls.Add($yearLabel)

$backupPanel = New-Object System.Windows.Forms.Panel
$backupPanel.Dock = "Fill"
$backupPanel.Margin = New-Object System.Windows.Forms.Padding(10, 0, 10, 0)
$backupNameLabel.Dock = "Top"
$backupNameLabel.Height = 22
$backupNameTextBox.Dock = "Top"
$backupPanel.Controls.Add($backupNameTextBox)
$backupPanel.Controls.Add($backupNameLabel)

$pstPanel = New-Object System.Windows.Forms.Panel
$pstPanel.Dock = "Fill"
$pstPanel.Margin = New-Object System.Windows.Forms.Padding(10, 0, 10, 0)
$pstLabel.Dock = "Top"
$pstLabel.Height = 22
$pstPathTextBox.Dock = "Top"
$pstPanel.Controls.Add($pstPathTextBox)
$pstPanel.Controls.Add($pstLabel)

$openPstPanel = New-Object System.Windows.Forms.Panel
$openPstPanel.Dock = "Fill"
$openPstButton.Dock = "Top"
$openPstButton.Height = 30
$openPstButton.Margin = New-Object System.Windows.Forms.Padding(0)
$openPstPanel.Padding = New-Object System.Windows.Forms.Padding(0, 22, 0, 0)
$openPstPanel.Controls.Add($openPstButton)

$topLayout.Controls.Add($yearPanel, 0, 0)
$topLayout.Controls.Add($backupPanel, 1, 0)
$topLayout.Controls.Add($pstPanel, 2, 0)
$topLayout.Controls.Add($openPstPanel, 3, 0)

$commandLayout = New-Object System.Windows.Forms.TableLayoutPanel
$commandLayout.Dock = "Fill"
$commandLayout.ColumnCount = 2
$commandLayout.RowCount = 2
$commandLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$commandLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 34))) | Out-Null
$commandLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 40))) | Out-Null
$commandLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 360))) | Out-Null
$commandLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$sourceLabel.Dock = "Fill"
$sourceStoreComboBox.Dock = "Fill"
$sourceStoreComboBox.Margin = New-Object System.Windows.Forms.Padding(0, 2, 10, 6)

$outlookButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$outlookButtonsPanel.Dock = "Fill"
$outlookButtonsPanel.FlowDirection = "LeftToRight"
$outlookButtonsPanel.WrapContents = $false
$outlookButtonsPanel.Margin = New-Object System.Windows.Forms.Padding(0)
foreach ($button in @($refreshButton, $diagnoseButton, $repairButton)) {
    $button.Width = 130
    $button.Height = 30
    $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 4)
    $outlookButtonsPanel.Controls.Add($button)
}

$actionStatusLayout = New-Object System.Windows.Forms.TableLayoutPanel
$actionStatusLayout.Dock = "Fill"
$actionStatusLayout.ColumnCount = 2
$actionStatusLayout.RowCount = 1
$actionStatusLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$actionStatusLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 430))) | Out-Null
$actionStatusLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

$actionButtonsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$actionButtonsPanel.Dock = "Fill"
$actionButtonsPanel.FlowDirection = "LeftToRight"
$actionButtonsPanel.WrapContents = $false
$actionButtonsPanel.Margin = New-Object System.Windows.Forms.Padding(0)
foreach ($button in @($simulateButton, $moveButton, $openLogButton)) {
    $button.Width = 130
    $button.Height = 34
    $button.Margin = New-Object System.Windows.Forms.Padding(0, 0, 12, 0)
    $actionButtonsPanel.Controls.Add($button)
}

$progressLayout = New-Object System.Windows.Forms.TableLayoutPanel
$progressLayout.Dock = "Fill"
$progressLayout.RowCount = 2
$progressLayout.ColumnCount = 1
$progressLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$progressLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 18))) | Out-Null
$progressLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 20))) | Out-Null
$progress.Dock = "Fill"
$progress.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 2)
$progressLabel.Dock = "Fill"
$progressLabel.AutoEllipsis = $true
$progressLabel.Margin = New-Object System.Windows.Forms.Padding(0)
$progressLayout.Controls.Add($progress, 0, 0)
$progressLayout.Controls.Add($progressLabel, 0, 1)

$actionStatusLayout.Controls.Add($actionButtonsPanel, 0, 0)
$actionStatusLayout.Controls.Add($progressLayout, 1, 0)

$commandLayout.Controls.Add($sourceLabel, 0, 0)
$commandLayout.Controls.Add($outlookButtonsPanel, 1, 0)
$commandLayout.Controls.Add($sourceStoreComboBox, 0, 1)
$commandLayout.Controls.Add($actionStatusLayout, 1, 1)

$contentLayout = New-Object System.Windows.Forms.TableLayoutPanel
$contentLayout.Dock = "Fill"
$contentLayout.ColumnCount = 3
$contentLayout.RowCount = 1
$contentLayout.Margin = New-Object System.Windows.Forms.Padding(0)
$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 31))) | Out-Null
$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 31))) | Out-Null
$contentLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 38))) | Out-Null

$sourceGroup = New-Object System.Windows.Forms.GroupBox
$sourceGroup.Text = "Carpeta origen"
$sourceGroup.Dock = "Fill"
$sourceGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 12, 0)
$sourceGroup.Padding = New-Object System.Windows.Forms.Padding(8)
$sourceInnerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$sourceInnerLayout.Dock = "Fill"
$sourceInnerLayout.ColumnCount = 1
$sourceInnerLayout.RowCount = 4
$sourceInnerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 28))) | Out-Null
$sourceInnerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$sourceInnerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) | Out-Null
$sourceInnerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) | Out-Null
$sourcePathLabel.Dock = "Fill"
$sourcePathLabel.AutoEllipsis = $true
$sourceTree.Dock = "Fill"
$includeSubfoldersCheckBox.Dock = "Fill"
$removeEmptySourceFoldersCheckBox.Dock = "Fill"
$sourceInnerLayout.Controls.Add($sourcePathLabel, 0, 0)
$sourceInnerLayout.Controls.Add($sourceTree, 0, 1)
$sourceInnerLayout.Controls.Add($includeSubfoldersCheckBox, 0, 2)
$sourceInnerLayout.Controls.Add($removeEmptySourceFoldersCheckBox, 0, 3)
$sourceGroup.Controls.Add($sourceInnerLayout)

$destinationGroup = New-Object System.Windows.Forms.GroupBox
$destinationGroup.Text = "PST o carpeta base destino"
$destinationGroup.Dock = "Fill"
$destinationGroup.Margin = New-Object System.Windows.Forms.Padding(0, 0, 12, 0)
$destinationGroup.Padding = New-Object System.Windows.Forms.Padding(8)
$destinationInnerLayout = New-Object System.Windows.Forms.TableLayoutPanel
$destinationInnerLayout.Dock = "Fill"
$destinationInnerLayout.ColumnCount = 1
$destinationInnerLayout.RowCount = 2
$destinationInnerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$destinationInnerLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 30))) | Out-Null
$destinationTree.Dock = "Fill"
$finalDestinationLabel.Dock = "Fill"
$finalDestinationLabel.AutoEllipsis = $true
$destinationInnerLayout.Controls.Add($destinationTree, 0, 0)
$destinationInnerLayout.Controls.Add($finalDestinationLabel, 0, 1)
$destinationGroup.Controls.Add($destinationInnerLayout)

$outputGroup = New-Object System.Windows.Forms.GroupBox
$outputGroup.Text = "Salida"
$outputGroup.Dock = "Fill"
$outputGroup.Margin = New-Object System.Windows.Forms.Padding(0)
$outputGroup.Padding = New-Object System.Windows.Forms.Padding(8)
$outputBox.Dock = "Fill"
$outputGroup.Controls.Add($outputBox)

$contentLayout.Controls.Add($sourceGroup, 0, 0)
$contentLayout.Controls.Add($destinationGroup, 1, 0)
$contentLayout.Controls.Add($outputGroup, 2, 0)

$hintLabel.Dock = "Fill"
$hintLabel.AutoEllipsis = $true

$mainLayout.Controls.Add($topLayout, 0, 0)
$mainLayout.Controls.Add($commandLayout, 0, 1)
$mainLayout.Controls.Add($contentLayout, 0, 2)
$mainLayout.Controls.Add($hintLabel, 0, 3)
$mainPanel.Controls.Add($mainLayout)

$form.Controls.Add($mainPanel)
$form.Controls.Add($headerPanel)

$openPstButton.Add_Click({ Invoke-UiAction -Title "Error abriendo PST" -Action {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Seleccionar archivo PST de respaldo"
    $dialog.Filter = "Archivos Outlook (*.pst)|*.pst|Todos los archivos (*.*)|*.*"
    $dialog.CheckFileExists = $true
    if ($dialog.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) {
        $pstPathTextBox.Text = $dialog.FileName
        Refresh-OutlookTree
    }
} })

$refreshButton.Add_Click({ Invoke-UiAction -Title "Error cargando Outlook" -Action { Refresh-OutlookTree } })
$diagnoseButton.Add_Click({ Invoke-UiAction -Title "Error en diagnostico" -Action { Show-OutlookDiagnostics } })
$repairButton.Add_Click({ Invoke-UiAction -Title "Error reparando Outlook" -Action { Repair-OutlookConnection } })
$simulateButton.Add_Click({ Invoke-UiAction -Title "Error simulando" -Action { Start-JobProcess -DoMove $false } })
$moveButton.Add_Click({ Invoke-UiAction -Title "Error moviendo correos" -Action { Start-JobProcess -DoMove $true } })
$sourceTree.Add_AfterSelect({
    Update-SourcePreview
    Update-FinalDestinationPreview
    $parts = @(Split-OutlookPath -Path (Get-SelectedSourcePath))
    if ($parts.Count -gt 0) {
        $sourceStoreComboBox.Text = $parts[0]
    }
})
$includeSubfoldersCheckBox.Add_CheckedChanged({ Update-SourcePreview })
$destinationTree.Add_AfterSelect({ Update-FinalDestinationPreview })
$backupNameTextBox.Add_TextChanged({ Update-FinalDestinationPreview })
$openLogButton.Add_Click({ Invoke-UiAction -Title "Error abriendo log" -Action {
    if (Test-Path -LiteralPath $LogPath) {
        Start-Process -FilePath $LogPath
    }
    else {
        [System.Windows.Forms.MessageBox]::Show("Todavía no existe el log. Ejecuta una simulación primero.", "Log no encontrado", "OK", "Information") | Out-Null
    }
} })

$yearNumeric.Add_ValueChanged({
    if ([string]::IsNullOrWhiteSpace($backupNameTextBox.Text) -or $backupNameTextBox.Text -match '\d{4}$') {
        $prefix = $backupNameTextBox.Text -replace '\s*-\s*\d{4}$', ''
        if ([string]::IsNullOrWhiteSpace($prefix)) {
            $prefix = "Respaldo"
        }
        $backupNameTextBox.Text = "$prefix - $([int]$yearNumeric.Value)"
    }
})

$form.Add_Shown({
    Load-Configuration
    Update-SourcePreview
    Update-FinalDestinationPreview
    Add-OutputLine "Paso 1: abre el PST o carga Outlook."
    Add-OutputLine "Paso 2: selecciona carpeta origen y carpeta base destino."
    Add-OutputLine "Paso 3: escribe la carpeta destino, simula y luego mueve."
})

if ($SelfTest) {
    Write-Host "OK UI initialized"
    return
}

[void][System.Windows.Forms.Application]::Run($form)
