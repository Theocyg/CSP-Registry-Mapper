if (-not (Get-Module -Name ImportExcel -ListAvailable)) {
    Write-Host "Le module ImportExcel n'est pas installé. Installation en cours..."
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force
    }
    catch {
        Write-Error "Impossible d'installer le module ImportExcel. Veuillez l'installer manuellement avec 'Install-Module -Name ImportExcel'"
        exit 1
    }
}

function Parse-CSPPath {
    param (
        [string]$CSPPath
    )
    
    $cleanPath = $CSPPath -replace '^\.\/', ''
    
    $parts = $cleanPath -split '/'
    
    $area = $null
    $category = $null
    $setting = $null
    $possibleAreas = @()
    
    if ($parts.Count -ge 6 -and $parts[0] -eq "Device" -and $parts[1] -eq "Vendor" -and $parts[2] -eq "MSFT" -and $parts[3] -eq "Policy") {
        if ($parts[4] -eq "Config") {
            $area = $parts[5].ToLower()
            
            $possibleAreas += $area
            
            switch ($area) {
                "system" { $possibleAreas += "systemservices" }
                "systemservices" { $possibleAreas += "system" }
                "devicelock" { $possibleAreas += "devicelockdown" }
                "devicelockdown" { $possibleAreas += "devicelock" }
                "update" { $possibleAreas += "windowsupdate" }
                "windowsupdate" { $possibleAreas += "update" }
            }
            
            if ($parts.Count -ge 7) {
                $setting = $parts[6]
            }
        }
    }
    
    return @{
        PossibleAreas = $possibleAreas
        Setting       = $setting
        IsValid       = ($area -ne $null)
    }
}

function Find-PolicyInRegistry {
    param (
        [array]$PossibleAreas,
        [string]$Setting
    )
    
    $results = @()
    
    function Search-RegistryRecursively {
        param (
            [string]$BasePath,
            [string]$Setting,
            [int]$MaxDepth = 2, # Profondeur de recherche maximale
            [int]$CurrentDepth = 0
        )
        
        $localResults = @()
        
        if ($CurrentDepth -gt $MaxDepth) {
            return $localResults
        }
        
        if (-not (Test-Path $BasePath -ErrorAction SilentlyContinue)) {
            return $localResults
        }
        
        if (Test-Path "$BasePath\$Setting" -ErrorAction SilentlyContinue) {
            $localResults += [PSCustomObject]@{
                Path      = "$BasePath\$Setting"
                Exists    = $true
                MatchType = "Exact"
            }
        }
        
        $potentialMatches = Get-ChildItem -Path $BasePath -ErrorAction SilentlyContinue | 
        Where-Object { $_.PSChildName -like "$Setting*" }
        
        foreach ($match in $potentialMatches) {
            if ($match.PSChildName -ne $Setting) {
                $localResults += [PSCustomObject]@{
                    Path       = $match.PSPath
                    Exists     = $true
                    MatchType  = "Préfixe"
                    ActualName = $match.PSChildName
                }
            }
        }
        
        $values = Get-ItemProperty -Path $BasePath -ErrorAction SilentlyContinue
        if ($values -ne $null) {
            foreach ($prop in $values.PSObject.Properties) {
                if ($prop.Name -eq $Setting) {
                    $localResults += [PSCustomObject]@{
                        Path      = "$BasePath"
                        Exists    = $true
                        MatchType = "Exact"
                        ValueName = $Setting
                        Value     = $prop.Value
                    }
                }
                elseif ($prop.Name -like "$Setting*") {
                    $localResults += [PSCustomObject]@{
                        Path      = "$BasePath"
                        Exists    = $true
                        MatchType = "Préfixe"
                        ValueName = $prop.Name
                        Value     = $prop.Value
                    }
                }
                elseif ($prop.Name -eq $Setting -or $prop.Name -like "$Setting*") {
                    $localResults += [PSCustomObject]@{
                        Path      = "$BasePath"
                        Exists    = $true
                        MatchType = "Insensible à la casse"
                        ValueName = $prop.Name
                        Value     = $prop.Value
                    }
                }
            }
        }
        

        if ($CurrentDepth -lt $MaxDepth) {
            Get-ChildItem -Path $BasePath -ErrorAction SilentlyContinue | ForEach-Object {
                $subResults = Search-RegistryRecursively -BasePath $_.PSPath -Setting $Setting -MaxDepth $MaxDepth -CurrentDepth ($CurrentDepth + 1)
                $localResults += $subResults
            }
        }
        
        return $localResults
    }
    
    foreach ($area in $PossibleAreas) {

        $basePaths = @(
            "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\$area",
            "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\device\$area",
            "HKLM:\SOFTWARE\Microsoft\PolicyManager\configured\device\$area",
            "HKLM:\SOFTWARE\Policies\Microsoft\$area"
        )
        
       
        $providerPaths = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\PolicyManager\providers" -ErrorAction SilentlyContinue | 
        ForEach-Object { Join-Path $_.PSPath "device\$area" }
        $basePaths += $providerPaths
        
    
        $additionalPaths = @(
            "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\MDM\$area",
            "HKLM:\SYSTEM\CurrentControlSet\Services\$area"
        )
        $basePaths += $additionalPaths
        
        foreach ($basePath in $basePaths) {
            $pathResults = Search-RegistryRecursively -BasePath $basePath -Setting $Setting
            $results += $pathResults
        }
    }
    
  
    $genericPaths = @(
        "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device",
        "HKLM:\SOFTWARE\Microsoft\PolicyManager\default\device",
        "HKLM:\SOFTWARE\Microsoft\PolicyManager\configured\device",
        "HKLM:\SOFTWARE\Policies\Microsoft"
    )
    
    foreach ($basePath in $genericPaths) {
        if (Test-Path $basePath -ErrorAction SilentlyContinue) {
            $subKeys = Get-ChildItem -Path $basePath -ErrorAction SilentlyContinue
            foreach ($key in $subKeys) {
                $pathResults = Search-RegistryRecursively -BasePath $key.PSPath -Setting $Setting -MaxDepth 1
                $results += $pathResults
            }
        }
    }
    
  
    if ($results.Count -eq 0) {
        $results += [PSCustomObject]@{
            Path      = "Non trouvé dans le registre"
            Exists    = $false
            MatchType = "N/A"
        }
    }
    
    return $results
}

function Analyze-CSPPaths {
    param (
        [string]$ExcelFilePath,
        [string]$WorksheetName = $null,
        [string]$CSPColumnName = $null
    )
    
    
    if (-not (Test-Path $ExcelFilePath)) {
        Write-Error "Le fichier Excel spécifié n'existe pas: $ExcelFilePath"
        return
    }
    
    try {
        
        $excelData = Import-Excel -Path $ExcelFilePath -WorksheetName $WorksheetName
        
       
        if (-not $CSPColumnName) {
            $potentialColumns = $excelData[0].PSObject.Properties.Name | Where-Object {
                $excelData[0].$_ -match '(Device|./Device)/Vendor/MSFT/Policy'
            }
            
            if ($potentialColumns.Count -gt 0) {
                $CSPColumnName = $potentialColumns[0]
                Write-Host "Utilisation automatique de la colonne: $CSPColumnName"
            }
            else {
                $CSPColumnName = $excelData[0].PSObject.Properties.Name[0]
                Write-Host "Aucune colonne CSP trouvée, utilisation de la première colonne: $CSPColumnName"
            }
        }
        
        
        $results = @()
        
        
        foreach ($row in $excelData) {
            $cspPath = $row.$CSPColumnName
            
           
            if ([string]::IsNullOrWhiteSpace($cspPath)) {
                continue
            }
            
            Write-Host "Analyse du chemin CSP: $cspPath"
            $parsedPath = Parse-CSPPath -CSPPath $cspPath
            
            if ($parsedPath.IsValid) {
                $registryResults = Find-PolicyInRegistry -PossibleAreas $parsedPath.PossibleAreas -Setting $parsedPath.Setting
                
                foreach ($regResult in $registryResults) {
                    $results += [PSCustomObject]@{
                        CSPPath       = $cspPath
                        Setting       = $parsedPath.Setting
                        RegistryPath  = $regResult.Path
                        Exists        = $regResult.Exists
                        CSPPolicyName = $regResult.ValueName
                        Value         = $regResult.Value
                    }
                }
            }
            else {
                $results += [PSCustomObject]@{
                    CSPPath       = $cspPath
                    Setting       = "Format non valide"
                    RegistryPath  = "N/A"
                    Exists        = $false
                    CSPPolicyName = "N/A"
                    Value         = "N/A"
                }
            }
        }
        
        
        $outputPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($ExcelFilePath), 
            [System.IO.Path]::GetFileNameWithoutExtension($ExcelFilePath) + "_ResultatsRegistre.xlsx")
        
        $results | Export-Excel -Path $outputPath -WorksheetName "Résultats" -AutoSize -TableName "RésultatsCSP"
        
        Write-Host "Analyse terminée. Les résultats ont été enregistrés dans: $outputPath"
        
        return $results
    }
    catch {
        Write-Error "Une erreur s'est produite lors de l'analyse: $_"
    }
}

$excelPath = Read-Host "Entrez le chemin complet du fichier Excel contenant les chemins CSP"


$worksheetName = Read-Host "Entrez le nom de la feuille Excel (laissez vide pour utiliser la première feuille)"
if ([string]::IsNullOrWhiteSpace($worksheetName)) {
    $worksheetName = $null
}


$columnName = Read-Host "Entrez le nom de la colonne contenant les chemins CSP (laissez vide pour détection automatique)"
if ([string]::IsNullOrWhiteSpace($columnName)) {
    $columnName = $null
}


Analyze-CSPPaths -ExcelFilePath $excelPath -WorksheetName $worksheetName -CSPColumnName $columnName