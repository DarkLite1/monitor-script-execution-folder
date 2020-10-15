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
        folder 'Archive' withing the drop folder. Any error in the input file
        will generate an error file next to he input file in the archive folder.

        In case the archive switch is not used and input error is detected a 
        mail will be send to the admin.
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
    [String]$ScriptAdmin = $env:ScriptAdmin,
    [Switch]$Archive
)

Begin {
    Try {
        Function Get-FunctionDefaultParameterHC {
            <#
            .SYNOPSIS
                Get the the default parameter values set in a script of function
            
            .EXAMPLE
                 Get-FunctionDefaultParameter -Path Get-Something
            
            .PARAMETER Path
                Function name or path to the script file
            #>
        
            [CmdletBinding()]
            [OutputType([hashtable])]
            Param (
                [Parameter(Mandatory)]
                [string]$Path
            )
            try {
                $ast = (Get-Command $Path).ScriptBlock.Ast
                
                $selectParams = @{
                    Property = @{ 
                        Name       = 'Name'; 
                        Expression = { $_.Name.VariablePath.UserPath } 
                    },
                    @{ 
                        Name       = 'Value'; 
                        Expression = { $_.DefaultValue.Extent.Text -replace "`"|'" }
                    }
                }
                
                $result = @{ }
        
                @($ast.FindAll( { $args[0] -is [System.Management.Automation.Language.ParameterAst] }, $true) | 
                    Where-Object { $_.DefaultValue } | 
                    Select-Object @selectParams).foreach( { 
                        $result[$_.Name] = $ExecutionContext.InvokeCommand.ExpandString($_.Value)
                    })
                $result
            }
            catch {
                throw "Failed retrieving the default parameter values: $_"
            }
        }

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

        #region Test logFolder
        if (-not (Test-Path $LogFolder -PathType Container)) {
            throw "Log folder '$LogFolder' not found"
        }
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
                defaultValues = Get-FunctionDefaultParameterHC -Path $script
            }

            $scriptParametersNameList = $item.Value['scriptParameters'].nameList
            #endregion

            #region Test ScriptName is not NULL
            if (-not $defaultParameters['ScriptName']) {
                throw "Parameter 'ScriptName' is missing and is mandatory for every script. We need to be able to hand over a unique 'ScriptName' to the script so ti can create a unique log folder and event viewer log based on the 'ScriptName'."
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
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
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
                Write-Verbose "Found input file '$inputFile'"
                $fileContent = $inputFile | Get-Content -Raw

                if ($Archive) {
                    Write-Verbose 'Move to archive folder'
                    $archiveDir = New-FolderHC -Path $folder -ChildPath Archive
                    $inputFile | Move-Item  -Destination $archiveDir -Force -EA Stop
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

                    $errorMessage = "Invalid json input file '$($inputFile.Name)'`r`nError:$_`r`nScript parameters:`r`n$($scriptSettings.scriptParameters.userInfoList -join `"`r`n`")" 

                    if ($Archive) {
                        Write-Verbose 'Create error file in archive folder'
                        $errorFile = "$($archiveDir.FullName)\$($inputFile.BaseName) - ERROR.txt"
                        $errorMessage | Out-File $errorFile -Encoding utf8 -Force
                    }
                    else {
                        Write-Verbose 'Send mail to admin'
                        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $errorMessage -Header $ScriptName
                    }

                    Write-EventLog @EventWarnParams -Message $errorMessage
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

                Write-Verbose "Start-Job ArgumentList '$startJobArgumentList'"
                #endregion

                #region Start job
                Write-Verbose 'Start job'

                Write-EventLog @EventOutParams -Message (
                    "Launch script:`n`n" +
                    "- InputFile:`t" + $inputFile.FullName + "`n" +
                    "- Script:`t`t" + $scriptSettings.script + "`n" +
                    "- ArgumentList:`t" + $startJobArgumentList)
                $StartJobParams = @{
                    Name                 = $defaultParameters['ScriptName']
                    InitializationScript = $LoadModules
                    LiteralPath          = $scriptSettings.script
                    ArgumentList         = $startJobArgumentList
                }
                $job = Start-Job @StartJobParams
                #endregion

                $jobsStarted += @{
                    archiveDir     = $archiveDir
                    inputFile      = $inputFile
                    scriptSettings = $scriptSettings
                    job            = $job
                }
                $sameScriptNameCount++
            }
        }

        if ($jobsStarted) {
            #region Test combined parameters were accepted
            Write-Verbose 'Wait 5 seconds for initial job launch'
            Start-Sleep -Seconds 5

            foreach ($j in $jobsStarted.Where( { $_.job.State -match '^Blocked$|^Failed$' })) {
                # We only check parameter validation errors, 
                # all other errors are handled by the child scripts
                Write-Verbose "Job '$($j.job.Name)' has status '$($j.job.State)'"

                if ($j.Job.State -eq 'Blocked') {
                    $parameterError = 'Have you provided all mandatory parameters?'
                }
                else {
                    $null = $j.Job | Receive-Job -ErrorVariable 'prob'
                    $parameterError = ( $prob | Where-Object { 
                            $_.FullyQualifiedErrorId -Like '*Parameter*' }).Exception.Message
                }
                
                if ($parameterError) {
                    Write-Warning 'Invalid input file parameters'

                    $errorMessage = "Invalid input file '$($j.inputFile.Name)'`r`n`r`nParameter error: $parameterError`r`n`r`nScript parameters:`r`n$($j.scriptSettings.scriptParameters.userInfoList -join `"`r`n`")" 

                    if ($Archive) {
                        Write-Verbose 'Create error file in archive folder'
                        $errorFile = "$($j.archiveDir.FullName)\$($j.inputFile.BaseName) - ERROR.txt"
                        $errorMessage | Out-File $errorFile -Encoding utf8 -Force
                    }
                    else {
                        Write-Verbose 'Send mail to admin'
                        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $errorMessage -Header $ScriptName
                    }

                    Write-EventLog @EventWarnParams -Message $errorMessage
                }
                else {
                    Write-Verbose 'Job has no parameter failure, resume other jobs'
                }

                $j.Job | Remove-Job -Force
                Continue
            }
            #endregion

            Write-Verbose 'Wait for all jobs to finish'
            Get-Job | Wait-Job | Receive-Job
            Write-Verbose 'All jobs finished'
        }
        else {
            Write-Verbose 'No jobs started'
        }

        Write-Verbose 'Script done'
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Exit 1
    }
    Finally {
        Get-Job | Remove-Job -Force
        Write-EventLog @EventEndParams
    }
}