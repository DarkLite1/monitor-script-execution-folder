#Requires -Module Assert, Pester

BeforeAll {
    $StartJobCommand = Get-Command Start-Job

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and 
        ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }
    
    $testScriptToExecute = (New-Item -Path "TestDrive:\scripts\scriptA.ps1" -Force -ItemType File -EA Ignore).FullName
    $testScriptFolder = (New-Item -Path "TestDrive:\input\scriptA" -ItemType Directory -Force -EA Ignore).FullName
    $testScriptParameters = @{
        ScriptName   = 'Get printers (BNL)'
        PrinterColor = 'red'
        PrinterName  = 'PRINTER1'
    }
    $testScriptSettings = @{ 
        script            = $testScriptToExecute
        defaultParameters = $testScriptParameters
    }
    
    @"
    Param (
        [Parameter(Mandatory)]
        [String]`$PrinterName,
        [Parameter(Mandatory)]
        [String]`$PrinterColor,
        [String]`$ScriptName,
        [String]`$PaperSize = 'A4'
    )
"@ | Out-File -FilePath $testScriptToExecute -Encoding utf8 -Force

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName    = 'Test'
        DropFolder    = (New-Item -Path "TestDrive:\input\scriptA\Get printers" -ItemType Directory -Force -EA Ignore).FullName
        ScriptMapping = @{ $testScriptFolder = $testScriptSettings }
        Archive       = $true
        LogFolder     = "TestDrive:\Log"
    }    

    Mock Send-MailHC
    Mock Start-Job -MockWith {
        & $StartJobCommand -Scriptblock { 1 } -Name 'Get printers (BNL)'
    }
    Mock Write-EventLog
}
Describe 'error handling' {    
    Context 'mandatory parameters' {
        It '<Name>' -TestCases @{ Name = 'DropFolder' } {
            (Get-Command $testScript).Parameters[$Name].Attributes.Mandatory |
            Should -BeTrue
        }
    }
    Context 'the logFolder' {
        It 'cannot be created' {
            $clonedParams = $testParams.Clone()
            $clonedParams.LogFolder = 'xx:\NotExisting'
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Failed creating the log folder 'xx:\NotExisting'*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
    }
    Context 'the dropFolder' {
        It 'should exist' {
            $clonedParams = $testParams.Clone()
            $clonedParams.DropFolder = @('NotExistingDropFolder')
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Drop folder 'NotExistingDropFolder' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It 'needs to be within a script folder' {
            $testDropFolder = (New-Item -Path "TestDrive:\input\ScriptWithoutScriptMapping\Get printers" -ItemType Container -Force -EA Ignore).FullName
            $clonedParams = $testParams.Clone()
            $clonedParams.DropFolder = $testDropFolder

            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Drop folder '$testDropFolder' has no matching script settings*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It 'can be the same as the script folder' {
            $clonedParams = $testParams.Clone()
            $clonedParams.DropFolder = $testScriptFolder

            . $testScript @clonedParams
    
            Should -Invoke -Not Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Drop folder*")
            }
        }
    }
    Context 'ScriptMapping' {
        Context 'the folder' {
            It 'should exist' {
                $clonedParams = $testParams.Clone()
                $clonedParams.ScriptMapping = @{
                    "TestDrive:\NotExistingFolder" = $testScriptSettings 
                }
                . $testScript @clonedParams
        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*Folder 'TestDrive:\NotExistingFolder' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
            }
            It 'should have settings' {
                $clonedParams = $testParams.Clone()
                $testFolderFullName = (New-Item -Path "TestDrive:\input\scriptA" -ItemType Container -Force -EA Ignore).FullName
                $clonedParams.ScriptMapping = @{
                    $testFolderFullName = $null
                }
                . $testScript @clonedParams
        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*Folder '$testFolderFullName' is missing settings*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
            }
            It "should have the property '<Name>'" -TestCases @(
                @{ Name = 'script' }
                @{ Name = 'defaultParameters' }
            ) {
                $clonedScriptSettings = $testScriptSettings.Clone()
                $clonedScriptSettings.Remove($Name)
                
                $clonedParams = $testParams.Clone()
                $clonedParams.ScriptMapping = @{ 
                    $testScriptFolder = $clonedScriptSettings 
                }
              
                . $testScript @clonedParams
        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*Folder '$testScriptFolder' is missing the property '$Name'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
            }
        }
        Context 'the script' {
            It 'should exist' {
                $clonedScriptSettings = $testScriptSettings.Clone()
                $clonedScriptSettings.script = "TestDrive:\doesNotExist.ps1" 

                $clonedParams = $testParams.Clone()
                $clonedParams.ScriptMapping = @{ 
                    $testScriptFolder = $clonedScriptSettings 
                }

                . $testScript @clonedParams
        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*Script 'TestDrive:\doesNotExist.ps1' for folder '$testScriptFolder' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
            } 
        }
        Context 'the defaultParameters' {
            It 'scriptName is mandatory' {
                $clonedScriptSettings = $testScriptSettings.Clone()
                $clonedScriptSettings.defaultParameters = @{
                    PrinterName  = 'PRINTER1'
                    PrinterColor = 'red'
                    ScriptName   = $null
                }

                $clonedParams = $testParams.Clone()
                $clonedParams.ScriptMapping = @{ 
                    $testScriptFolder = $clonedScriptSettings 
                }

                . $testScript @clonedParams
        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*Parameter 'ScriptName' is missing and is mandatory for every script. We need to be able to hand over a unique 'ScriptName' to the script so it can create a unique log folder and event viewer log based on the 'ScriptName'.*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
            }
            It 'should exist in the script' {
                $clonedScriptSettings = $testScriptSettings.Clone()
                $clonedScriptSettings.defaultParameters = @{
                    PrinterName      = 'PRINTER1'
                    PrinterColor     = 'red'
                    ScriptName       = 'Get printer (BNL)'
                    UnknownParameter = 'Not allowed'
                }

                $clonedParams = $testParams.Clone()
                $clonedParams.ScriptMapping = @{ 
                    $testScriptFolder = $clonedScriptSettings 
                }

                . $testScript @clonedParams
        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*Default parameter 'UnknownParameter' does not exist in script '$testScriptToExecute'*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
            }
        }
    }
}
Describe 'a valid user input file found in a drop folder' {
    BeforeAll {
        $testInputFile = (Join-Path $testParams.dropFolder 'inputFile.json')
     
        @{  PrinterName = "MyCustomPrinter" } | 
        ConvertTo-Json | Out-File $testInputFile -Encoding utf8

        $clonedParams = $testParams.Clone()
        $clonedParams.LogFolder = "$($testParams.LogFolder)\Monitor\Monitor script execution folder\$($testParams.ScriptName)"
        . $testScript @clonedParams
    }
    It 'is detected because it has a .json file extension' {
        $inputFile.FullName | Should -Be $testInputFile
    }
    It 'is moved to the archive folder within the drop folder' {
        "$($testParams.DropFolder)\inputFile.json" | Should -Not -Exist
        "$($testParams.DropFolder)\Archive\inputFile.json" | Should -Exist
    }
    It 'is matched with the correct script in the ScriptMapping hashtable' {
        $scriptSettings.script | Should -BeExactly $testScriptToExecute
    }
    It 'is converted to user specific parameters' {
        $userParameters.PrinterName | Should -BeExactly 'MyCustomPrinter'
        $userParameters.PrinterColor | Should -BeNullOrEmpty
    }
    It 'should overwrite the default parameters with the user specific parameters' {
        $startJobArgumentList | Should -Contain 'MyCustomPrinter'
    }
    It 'should keep the default parameters when the user did not set it' {
        $startJobArgumentList | Should -Contain 'red'
    }
    It 'should invoke Start-Job with the parameters in the correct order' {
        Should -Invoke Start-Job -Exactly 1 -Scope Describe -ParameterFilter {
            ($LiteralPath -eq $testScriptToExecute) -and
            ($ArgumentList[0] -eq 'MyCustomPrinter') -and
            ($ArgumentList[1] -eq 'red') -and
            ($ArgumentList[2] -eq 'Get printers (BNL)') -and
            ($ArgumentList[3] -eq 'A4') # default parameter in the script is copied
        }
    }
    It 'should send an email to the admin' {
        Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
            ($To -eq $ScriptAdmin) -and 
            ($Subject -eq '1 script started') -and 
            ($Message -like "*Script name*")
        }
    }
    Context 'logging' {
        It 'the log folder is created' {            
            $clonedParams.LogFolder | Should -Exist
        }
        It 'a copy of the input file is stored in the log folder' {
            $testLogFile = Get-ChildItem $clonedParams.LogFolder -Recurse -File
            
            $testLogFile | Should -Not -BeNullOrEmpty
            $testLogFile.Name | Should -BeLike '*- Get printers - inputFile.json'
        }
    }
}
Describe 'when the archive switch is not used' {
    Context 'and the input file is correct' {
        BeforeAll {
            $testInputFile = (Join-Path $testParams.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter" } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8
        
            $clonedParams = $testParams.Clone()
            $clonedParams.Remove('Archive')

            . $testScript @clonedParams
        }
        It 'no Archive folder is created' {
            "$($testParams.DropFolder)\Archive" | Should -Not -Exist
        }
        It 'the input file is not moved to the archive folder' {
            $testInputFile | Should -Exist
        }
        It 'Start-Job is called' {
            Should -Invoke Start-Job -Scope Context
        }
    }
    Context 'and the input file is not correct' {
        BeforeAll {
            $clonedParams = $testParams.Clone()
            $clonedParams.Remove('Archive')

            $testInputFile = (Join-Path $testParams.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter"; UnknownParameter = 'Oops' } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8
    
            . $testScript @clonedParams
        }
        It 'no Archive folder is created' {
            "$($testParams.DropFolder)\Archive" | Should -Not -Exist
        }
        It 'the input file is not moved to the archive folder' {
            $testInputFile | Should -Exist
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers')
                ($Message -like "*parameter 'UnknownParameter' is not accepted*")
            }
        }
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
    }
}
Describe 'when the drop folder is the same as the folder in ScriptMapping' {
    It 'the script is also correctly executed' {
        $testInputFile = (Join-Path $testScriptFolder 'inputFile.json')
     
        @{  PrinterName = "MyCustomPrinter" } | 
        ConvertTo-Json | Out-File $testInputFile -Encoding utf8

        $clonedParams = $testParams.Clone()
        $clonedParams.DropFolder = $testScriptFolder

        . $testScript @clonedParams

        Should -Invoke Start-Job -Exactly 1 -Scope Describe -ParameterFilter {
            ($LiteralPath -eq $testScriptToExecute) -and
            ($ArgumentList[0] -eq 'MyCustomPrinter') -and
            ($ArgumentList[1] -eq 'red')
        }
    }
}
Describe 'multiple input files for the same script' {
    BeforeAll {
        $testInputFile1 = (Join-Path $testParams.dropFolder 'inputFile1.json')
        $testInputFile2 = (Join-Path $testParams.dropFolder 'inputFile2.json')
     
        @{ PrinterName = "Printer1" } | 
        ConvertTo-Json | Out-File $testInputFile1 -Encoding utf8
        @{  PrinterName = "Printer2" } | 
        ConvertTo-Json | Out-File $testInputFile2 -Encoding utf8
      
        . $testScript @testParams
    }
    It 'should invoke Start-Job with a unique script name for a unique log folder creation by the child script' {
        Should -Invoke Start-Job -Exactly 1 -Scope Describe -ParameterFilter {
            ($LiteralPath -eq $testScriptToExecute) -and
            ($ArgumentList[0] -eq 'Printer1') -and
            ($ArgumentList[1] -eq 'red') -and
            ($ArgumentList[2] -eq 'Get printers (BNL)')
        }
        Should -Invoke Start-Job -Exactly 1 -Scope Describe -ParameterFilter {
            ($LiteralPath -eq $testScriptToExecute) -and
            ($ArgumentList[0] -eq 'Printer2') -and
            ($ArgumentList[1] -eq 'red') -and
            ($ArgumentList[2] -eq 'Get printers (BNL) 1')
        }
    }
}
Describe 'when an input file is incorrect because' {
    Context 'it has another extension than .json' {
        BeforeAll {
            Mock ConvertFrom-Json

            $testInputFileExclude = (Join-Path $testParams.dropFolder 'file.txt')
         
            (New-Item -Path $testInputFileExclude -Force -ItemType File -EA Ignore).FullName
    
            . $testScript @testParams
        }
        It 'it is ignored and left in the drop folder' {
            $testInputFileExclude | Should -Exist
            Should -Not -Invoke ConvertFrom-Json -Scope Context
        }
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
        It 'Send-MailHC is not called' {
            Should -Not -Invoke Send-MailHC -Exactly 1 -Scope Context
        }
    }
    Context 'it is not a valid .json file' {
        BeforeAll {
            $testInputFile = (Join-Path $testParams.dropFolder 'inputFile.json')
     
            (New-Item -Path $testInputFile -Force -ItemType File -EA Ignore).FullName

            "NotJsonFormat ;!= " | Out-File $testInputFile -Encoding utf8

            $clonedParams = $testParams.Clone()
            $clonedParams.LogFolder = "$($testParams.LogFolder)\test"
            . $testScript @clonedParams
        }
        It 'it is moved to the archive folder in the drop folder' {
            "$($testParams.DropFolder)\inputFile.json" | Should -Not -Exist
            "$($testParams.DropFolder)\Archive\inputFile.json" | Should -Exist
        }
        It 'an error file is created in the archive folder' {
            "$($testParams.DropFolder)\Archive\inputFile - ERROR.json" | 
            Should -Exist
        }
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers')
                ($Message -like "*Invalid json input file*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFile = Get-ChildItem $clonedParams.LogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $clonedParams.LogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[1].Name | Should -BeLike '*- Get printers - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[0].Name | Should -BeLike '*- Get printers - inputFile - ERROR.json'
            }
        }
    }
    Context 'the user used a parameter that is not available in the scriptMapping table' {
        BeforeAll {
            $testInputFile = (Join-Path $testParams.dropFolder 'inputFile.json')
     
            @{  
                PrinterName      = "MyCustomPrinter"
                UnknownParameter = 'Oops' 
            } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8
    
            $clonedParams = $testParams.Clone()
            $clonedParams.LogFolder = "$($testParams.LogFolder)\Test2"
            . $testScript @clonedParams -EA Ignore
        }
        It 'it is moved to the archive folder in the drop folder' {
            "$($testParams.DropFolder)\inputFile.json" | Should -Not -Exist
            "$($testParams.DropFolder)\Archive\inputFile.json" | Should -Exist
        }
        It 'an error file is created in the archive folder' {
            "$($testParams.DropFolder)\Archive\inputFile - ERROR.json" | 
            Should -Exist
        }
        It 'the error file contains the incorrect parameter name' {
            $actual = Get-Content "$($testParams.DropFolder)\Archive\inputFile - ERROR.json" -Raw | ConvertFrom-Json
            $actual.errorMessage | Should -BeLike "*parameter 'UnknownParameter' is not accepted*"
        } 
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers')
                ($Message -like "*parameter 'UnknownParameter' is not accepted*")
            }
        } 
        Context 'logging' {
            BeforeAll {
                $testLogFile = Get-ChildItem $clonedParams.LogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $clonedParams.LogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[1].Name | Should -BeLike '*- Get printers - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[0].Name | Should -BeLike '*- Get printers - inputFile - ERROR.json'
            }
        }
    }
}
Describe 'when Start-Job fails' {
    Context 'because of a missing mandatory parameter' {
        BeforeAll {
            $testInputFile = (Join-Path $testParams.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter" } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8

            Mock Start-Job {
                & $StartJobCommand -Scriptblock { 
                    Param (
                        [parameter(Mandatory)]
                        [int]$Number,
                        [parameter(Mandatory)]
                        [String]$Name
                    )
                    # parameters name is missing
                } -ArgumentList 1
            }
            . $testScript @testParams -EA SilentlyContinue
        }
        It 'the input file is moved to the archive folder' {
            "$($testParams.DropFolder)\inputFile.json" | Should -Not -Exist
            "$($testParams.DropFolder)\Archive\inputFile.json" | Should -Exist
        }
        It 'an error file is created in the archive folder' {
            "$($testParams.DropFolder)\Archive\inputFile - ERROR.json" | 
            Should -Exist
        }
        It 'the error file contains the error message' {
            $actual = Get-Content "$($testParams.DropFolder)\Archive\inputFile - ERROR.json" -Raw | ConvertFrom-Json
            $actual.errorMessage | Should -BeLike "*Job status 'Blocked', have you provided all mandatory parameters?*"
        } -Tag test
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers')
                ($Message -like "*Job status 'Blocked', have you provided all mandatory parameters?*")
            }
        } 
    }
    Context 'because of parameter validation issues' {
        BeforeAll {
            $testInputFile = (Join-Path $testParams.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter" } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8

            Mock Start-Job {
                & $StartJobCommand -Scriptblock { 
                    Param (
                        [parameter(Mandatory)]
                        [int]$IncorrectParameters
                    )
                    # parameters not matching
                } -ArgumentList 'string'
            }
            . $testScript @testParams -EA SilentlyContinue
        }
        It 'the input file is moved to the archive folder' {
            "$($testParams.DropFolder)\inputFile.json" | Should -Not -Exist
            "$($testParams.DropFolder)\Archive\inputFile.json" | Should -Exist
        }
        It 'an error file is created in the archive folder' {
            "$($testParams.DropFolder)\Archive\inputFile - ERROR.json" | 
            Should -Exist
        }
        It 'the error file contains the error message' {
            $actual = Get-Content "$($testParams.DropFolder)\Archive\inputFile - ERROR.json" -Raw | ConvertFrom-Json
            $actual.errorMessage | Should -BeLike "*Cannot process argument transformation on parameter 'IncorrectParameters'*"
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers')
                ($Message -like "*Cannot process argument transformation on parameter 'IncorrectParameters'*")
            }
        } 
    }
    Context 'because of errors in the execution script' {
        BeforeAll {
            $testInputFile = (Join-Path $testParams.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter" } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8

            Mock Start-Job {
                & $StartJobCommand -Scriptblock { 
                    throw 'Failure in script'
                }
            }

            . $testScript @testParams -EA SilentlyContinue
        }
        It 'the input file is moved to the archive folder' {
            "$($testParams.DropFolder)\inputFile.json" | Should -Not -Exist
            "$($testParams.DropFolder)\Archive\inputFile.json" | Should -Exist
        }
        It 'an error file is created in the archive folder' {
            "$($testParams.DropFolder)\Archive\inputFile - ERROR.json" | 
            Should -Exist
        }
        It 'the error file contains the error message' {
            $actual = Get-Content "$($testParams.DropFolder)\Archive\inputFile - ERROR.json" -Raw | ConvertFrom-Json
            $actual.errorMessage | Should -BeLike "*Failure in script*"
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers')
                ($Message -like "*Failure in script*")
            }
        } 
    } 
}