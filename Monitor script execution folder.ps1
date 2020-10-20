<#
    .SYNOPSIS
        Allow end users to execute scripts by placing input files in a folder

    .DESCRIPTION
        This script is designed to be triggered by a scheduled task. It will
        execute the correct script based on the folder where the user stores
        the the input file.

    .PARAMETER ScriptMapping
        Defines the link between the script to execute and the folder where
        the input file is stored. In case a drop folder is a child folder
        of the folder defined in ScriptMapping it will still execute the 
        correct script and no extra hash key is required in ScriptMapping.

        'parent folder or drop folder' = @{
            script = 'the ps1 file to execute when files are in the drop folder'
            defaultParameters = 'default values for the script'
                - scriptName = 'is a mandatory parameter to generate unique 
                                script names for echt script when multiple input 
                                files were found for the same script'
        }

    .PARAMETER DropFolder
        Path of the folder where the input files will be stored. this folder
        will be searched for .json files and will trigger script execution for
        every single file.

    .PARAMETER Archive
        When the archive switch is used the input file will be moved to the 
        folder 'Archive' within the drop folder. Any error in the input file
        will generate an error file next to he input file in the archive folder.
 #>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String[]]$DropFolder,
    [String]$ScriptName = 'Monitor script execution folder (BNL)',
    [HashTable]$ScriptMapping = @{
        'T:\Input\Scripts\Get AD group members'       = @{
            script            = 'T:\Prod\AD Reports\AD Group members\AD Group members.ps1'
            defaultParameters = @{
                ScriptName = 'AD group members (BNL)'
            }
        }
        'T:\Input\Scripts\Get matrix AD object names' = @{
            script            = 'T:\Prod\Permission matrix\Get matrix AD object names\Get matrix AD object names.ps1'
            defaultParameters = @{
                ScriptName = 'Matrix AD object names (BNL)'
            }
        }
    },
    [String]$LogFolder = "\\$env:COMPUTERNAME\Log",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN,
    [Switch]$Archive
)

Begin {
    Try {
        Function Get-ScriptSettingsHC {
            Param (
                [parameter(Mandatory)]
                [String]$folderName
            )

            $keyMatches = $ScriptMapping.GetEnumerator() | Where-Object {
                ($folderName -eq $_.Key ) -or
                ($folderName -like "$($_.Key)\*" )
            }

            if ($keyMatches.count -eq 0) {
                throw "Drop folder '$folderName' has no matching script settings"
            }
            if ($keyMatches.count -ge 2) {
                throw "Drop folder '$folderName' has multiple matching script settings"
            }

            $keyMatches.Value
        }

        $null = Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        $LogParams = @{
            LogFolder    = New-FolderHC -Path $LogFolder -ChildPath "Monitor\Monitor script execution folder\$ScriptName"
            Name         = $ScriptName
            Date         = 'ScriptStartTime'
            NoFormatting = $true
        }
        $LogFile = New-LogFileNameHC @LogParams
        #endregion

        $psBuildInParameters = [System.Management.Automation.PSCmdlet]::CommonParameters
        $psBuildInParameters += [System.Management.Automation.PSCmdlet]::OptionalCommonParameters

        #region Test ScriptMapping
        foreach ($item in $ScriptMapping.GetEnumerator()) {
            $folder = $item.Key
            $script = $item.Value.script
            $defaultParameters = $item.Value.defaultParameters

            if (-not (Test-Path $folder -PathType Container)) {
                throw "Folder '$folder' not found"
            }

            if (-not $item.Value) {
                throw "Folder '$folder' is missing settings"
            }

            if (-not $script) {
                throw "Folder '$folder' is missing the property 'script'"
            }

            if (-not $defaultParameters) {
                throw "Folder '$folder' is missing the property 'defaultParameters'"
            }

            if (-not (Test-Path $script -PathType Leaf)) {
                throw "Script '$script' for folder '$folder' not found"
            }

            #region Get the script parameters
            $scriptParameters = (Get-Command $script).Parameters.GetEnumerator() | 
            Where-Object { $psBuildInParameters -notContains $_.Key }

            $item.Value['scriptParameters'] = @{
                getCommand    = $scriptParameters
                nameList      = $($scriptParameters.Key)
                userInfoList  = foreach ($p in $scriptParameters.GetEnumerator()) {
                    'Name: {0} Type: {1} Mandatory: {2} ' -f $p.Value.Name, $p.Value.ParameterType, $p.Value.Attributes.Mandatory
                }
                defaultValues = Get-DefaultParameterValuesHC -Path $script
            }

            $scriptParametersNameList = $item.Value['scriptParameters'].nameList
            #endregion

            #region Test ScriptName is mandatory
            if (-not $defaultParameters['ScriptName']) {
                throw "Parameter 'ScriptName' is missing and is mandatory for every script. We need to be able to hand over a unique 'ScriptName' to the script so it can create a unique log folder and event viewer log based on the 'ScriptName'."
            }
            #endregion

            #region Test that the default parameters are in the script
            $($defaultParameters.Keys) | 
            Where-Object { $scriptParametersNameList -notContains $_ } | 
            ForEach-Object {
                throw "Default parameter '$_' does not exist in script '$script'. Only the following parameters are known: $($defaultParameters.Keys)"
            }
            #endregion
        }
        #endregion

        #region Test dropFolder
        foreach ($d in $DropFolder) {
            if (-not (Test-Path -LiteralPath $d -PathType Container)) {
                throw "Drop folder '$d' not found"
            }
            
            $null = Get-ScriptSettingsHC -folderName $d
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $jobsStarted = @()

        Write-Verbose 'Load modules for job scope'
        $LoadModules = {
            Get-ChildItem ($env:PSModulePath -split ';') -EA Ignore |
            Where-Object Name -Like 'Toolbox*' | Import-Module
        }

        foreach ($folder in $dropFolder) {
            $sameScriptNameCount = 0
            Write-Verbose "Check drop folder '$folder'"

            $scriptSettings = Get-ScriptSettingsHC -folderName $folder

            $defaultParameters = $scriptSettings.defaultParameters
            Write-Verbose "Default parameters '$($defaultParameters.Keys)'"

            $scriptParametersNameList = $scriptSettings.scriptParameters.NameList
            Write-Verbose "Script scriptParametersNameList '$scriptParametersNameList'"

            $scriptParametersDefaultValues = $scriptSettings.scriptParameters.defaultValues
            Write-Verbose "Script scriptParametersNameList '$scriptParametersDefaultValues'"

            foreach (
                $inputFile in 
                (Get-ChildItem -LiteralPath $folder -File -Filter '*.json')
            ) {
                $job = @{
                    inputFile      = $inputFile
                    scriptSettings = $scriptSettings
                    argumentList   = $null
                    owner          = $null
                    archiveDir     = $null
                    job            = $null
                }
                
                Write-Verbose "Found input file '$inputFile'"
                $fileContent = $inputFile | Get-Content -Raw

                #region Copy input file to log folder
                Write-Verbose "Copy input file to log folder '$($LogParams.LogFolder)'"
                Copy-Item -Path $inputFile.FullName -Destination "$LogFile - $($inputFile.Directory.Name) - $($inputFile.Name)"
                #endregion

                $job.owner = $inputFile.GetAccessControl().Owner -replace "$env:USERDOMAIN\\"

                if ($Archive) {
                    Write-Verbose 'Move to archive folder'
                    $job.archiveDir = New-FolderHC -Path $folder -ChildPath Archive
                    $inputFile | Move-Item -Destination $job.archiveDir -Force -EA Stop
                }

                #region Test valid .json file and allowed user parameters
                try {
                    $userParameters = $fileContent | ConvertFrom-Json
                    Write-Verbose "User parameters '$userParameters'"    

                    $userParameters.PSObject.Properties.Name | Where-Object {
                        $scriptParametersNameList -notContains $_
                    } | ForEach-Object {
                        $invalidParameter = $_
                        throw "The parameter '$invalidParameter' is not accepted by script '$($scriptSettings.script)'."
                    }
                }
                catch {
                    Write-Warning "Invalid .json input file: $_"

                    #region Create error file
                    $errorFileMessage = [ordered]@{
                        errorMessage      = $_.Exception.Message
                        scriptFile        = $scriptSettings.script
                        scriptParameters  = $scriptSettings.scriptParameters.userInfoList
                        startJobArguments = $startJobArgumentList
                        dropFolder        = $folder
                    } | ConvertTo-Json -Depth 5 | Format-JsonHC

                    $logFileFullName = "$LogFile - $($inputFile.Directory.Name) - $($inputFile.BaseName) - ERROR.json"
                    
                    $errorFileMessage | Out-File $logFileFullName -Encoding utf8 -Force -EA Ignore
                    #endregion

                    #region Send mail
                    $mailParams = @{
                        To          = $ScriptAdmin 
                        Subject     = "FAILURE - $($inputFile.Directory.Name)"
                        Priority    = 'High' 
                        Message     = "Script '<b>$($inputFile.Directory.Name)</b>' failed with error:
                <p>Invalid parameter file: $_</p>
                <p><i>* Check the attachment for details</i></p>"
                        Header      = $ScriptName 
                        Attachments = $logFileFullName
                    }
                    Send-MailHC @mailParams
                    #endregion

                    if ($Archive) {
                        Write-Verbose 'Create error file in archive folder'
                        $errorFile = "$($job.archiveDir.FullName)\$($inputFile.BaseName) - ERROR.json"
                        $errorFileMessage | Out-File $errorFile -Encoding utf8 -Force
                    }

                    Write-EventLog @EventWarnParams -Message $errorFileMessage
                    Continue
                }
                #endregion

                #region Set script name to be unique for each job
                if ($userParameters.ScriptName) {
                    $userParameters.ScriptName = $null
                }
                if ($sameScriptNameCount -ne 0) {
                    $defaultParameters['ScriptName'] = $defaultParameters['ScriptName'] + " $sameScriptNameCount"
                }
                #endregion

                #region Build ordered ArgumentList for Start-Job
                $startJobArgumentList = @()

                foreach ($p in $scriptParametersNameList) {
                    $value = $null
                    if ($scriptParametersDefaultValues[$p]) {
                        $value = $scriptParametersDefaultValues[$p]
                    }
                    if ($defaultParameters.$p) {
                        $value = $defaultParameters.$p
                    }
                    if ($userParameters.$p) {
                        $value = $userParameters.$p
                    }
                    Write-Verbose "Parameter name '$p' value '$value'"
                    $startJobArgumentList += , $value
                }

                $job.argumentList = $startJobArgumentList
                Write-Verbose "Start-Job ArgumentList '$startJobArgumentList'"
                #endregion

                #region Start job
                Write-Verbose 'Start job'

                Write-EventLog @EventOutParams -Message (
                    "Launch script:`n" +
                    "`n- InputFile:`t" + $inputFile.FullName + 
                    "`n- Script:`t`t" + $scriptSettings.script + 
                    "`n- ArgumentList:`t" + $startJobArgumentList)
                $StartJobParams = @{
                    Name                 = $defaultParameters['ScriptName']
                    InitializationScript = $LoadModules
                    LiteralPath          = $scriptSettings.script
                    ArgumentList         = $startJobArgumentList
                }
                $job.job = Start-Job @StartJobParams
                #endregion

                $jobsStarted += $job
                
                $sameScriptNameCount++
            }
        }

        if ($jobsStarted) {
            #region send mail to admin
            $mailParams = @{
                To      = $ScriptAdmin 
                Subject = "$($jobsStarted.Count) script started"
                Message = "<p>Scripts started:</p>
                <table>
                <tr><th>Script name</th><th>Input file</th><th>Owner</th></tr>
                $($jobsStarted.ForEach({
                    "<tr><td style=``"text-align: center``">{0}</td><td>{1}</td><td>{2}</td></tr>" -f $_.job.Name, $_.inputFile.Name, $_.owner
                }))
                </table>"
                Header  = $ScriptName
            }
            Send-MailHC @mailParams
            #endregion

            Write-Verbose 'Wait 5 seconds for initial job launch'
            Start-Sleep -Seconds 5

            foreach ($j in $jobsStarted) {
                $jobError = $null
                Write-Verbose "Job '$($j.job.Name)' status '$($j.job.State)'"

                # Missing mandatory parameters set the state to 'Blocked'
                if ($j.Job.State -eq 'Blocked') {
                    $jobError = "Job status 'Blocked', have you provided all mandatory parameters?"
                }
                else {
                    $j.job | Wait-Job
                    $null = $j.Job | Receive-Job -ErrorVariable 'jobError'
                    if ($jobError) { $jobError = $jobError.Exception.Message }
                }
                
                if ($jobError) {
                    Write-Warning "Job '$($j.job.Name)' with status '$($j.job.State)' has error: $jobError"

                    #region Create error file
                    $errorFileMessage = [ordered]@{
                        errorMessage      = $jobError
                        jobState          = $j.job.State
                        inputFile         = $j.inputFile.FullName
                        scriptFile        = $j.scriptSettings.script
                        scriptParameters  = $j.scriptSettings.scriptParameters.userInfoList
                        startJobArguments = $j.argumentList
                    } | ConvertTo-Json -Depth 5 | Format-JsonHC
                    
                    $logFileFullName = "$LogFile - $($j.inputFile.Directory.Name) - $($j.inputFile.BaseName) - ERROR.json"
                                        
                    $errorFileMessage | Out-File $logFileFullName -Encoding utf8 -Force -EA Ignore
                    #endregion
                                        
                    #region Send mail
                    $mailParams = @{
                        To          = $ScriptAdmin 
                        Subject     = "FAILURE - $($j.inputFile.Directory.Name)"
                        Priority    = 'High' 
                        Message     = "Script '<b>$($j.inputFile.Directory.Name)</b>' failed with error:
                        <p>$jobError</p>
                        <p><i>* Check the attachment for details</i></p>"
                        Header      = $ScriptName 
                        Attachments = $logFileFullName
                    }
                    Send-MailHC @mailParams
                    #endregion

                    if ($Archive) {
                        Write-Verbose 'Create error file in archive folder'
                        $errorFile = "$($j.archiveDir.FullName)\$($j.inputFile.BaseName) - ERROR.json"
                        $errorFileMessage | Out-File $errorFile -Encoding utf8 -Force
                    }

                    Write-EventLog @EventWarnParams -Message $errorFileMessage
                }
                else {
                    Write-Verbose 'Job has no parameter failure, resume other jobs'
                }

                $j.Job | Remove-Job -Force
                Continue
            }
            #endregion

            Write-Verbose 'All jobs finished'
        }
        else {
            Write-Verbose 'No jobs started'
        }

        Write-Verbose 'Script done'
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Get-Job | Remove-Job -Force
        Write-EventLog @EventEndParams
    }
}