#Requires -Module Assert, Pester

BeforeAll {
    $StartJobCommand = Get-Command Start-Job

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

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

    $Params = @{
        ScriptName    = 'Test'
        DropFolder    = (New-Item -Path "TestDrive:\input\scriptA\Weekly" -ItemType Directory -Force -EA Ignore).FullName
        LogFolder     = (New-Item -Path "TestDrive:\Log" -ItemType Directory -EA Ignore).FullName
        ScriptMapping = @{ $testScriptFolder = $testScriptSettings }
        Archive       = $true
    }    

    Mock Send-MailHC
    Mock Start-Job
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
        It 'should exist' {
            $clonedParams = $Params.Clone()
            $clonedParams.LogFolder = 'NotExistingLogFolder'
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Log folder 'NotExistingLogFolder' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
    }
    Context 'the dropFolder' {
        It 'should exist' {
            $clonedParams = $Params.Clone()
            $clonedParams.DropFolder = @('NotExistingDropFolder')
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Drop folder 'NotExistingDropFolder' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It 'needs to be within a script folder' {
            $testDropFolder = (New-Item -Path "TestDrive:\input\ScriptWithoutScriptMapping\Weekly" -ItemType Container -Force -EA Ignore).FullName
            $clonedParams = $Params.Clone()
            $clonedParams.DropFolder = $testDropFolder

            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Drop folder '$testDropFolder' has no matching script settings*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It 'can be the same as the script folder' {
            $clonedParams = $Params.Clone()
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
                $clonedParams = $Params.Clone()
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
                $clonedParams = $Params.Clone()
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
                
                $clonedParams = $Params.Clone()
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

                $clonedParams = $Params.Clone()
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

                $clonedParams = $Params.Clone()
                $clonedParams.ScriptMapping = @{ 
                    $testScriptFolder = $clonedScriptSettings 
                }

                . $testScript @clonedParams
        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*Parameter 'ScriptName' is missing and is mandatory for every script. We need to be able to hand over a unique 'ScriptName' to the script so ti can create a unique log folder and event viewer log based on the 'ScriptName'.*")
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

                $clonedParams = $Params.Clone()
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
        $testInputFile = (Join-Path $Params.dropFolder 'inputFile.json')
     
        @{  PrinterName = "MyCustomPrinter" } | 
        ConvertTo-Json | Out-File $testInputFile -Encoding utf8

        . $testScript @Params
    }
    It 'is detected because it has a .json file extension' {
        $inputFile.FullName | Should -Be $testInputFile
    }
    It 'is moved to the archive folder within the drop folder' {
        "$($Params.DropFolder)\inputFile.json" | Should -Not -Exist
        "$($Params.DropFolder)\Archive\inputFile.json" | Should -Exist
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
        Should -invoke Start-Job -Exactly 1 -Scope Describe -ParameterFilter {
            ($LiteralPath -eq $testScriptToExecute) -and
            ($ArgumentList[0] -eq 'MyCustomPrinter') -and
            ($ArgumentList[1] -eq 'red') -and
            ($ArgumentList[2] -eq 'Get printers (BNL)') -and
            ($ArgumentList[3] -eq 'A4') # default parameter in the script is copied
        }
    }
}

Describe 'when the archive switch is not used' {
    Context 'and the input file is correct' {
        BeforeAll {
            $testInputFile = (Join-Path $Params.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter" } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8
        
            $clonedParams = $Params.Clone()
            $clonedParams.Remove('Archive')

            . $testScript @clonedParams
        }
        It 'no Archive folder is created' {
            "$($Params.DropFolder)\Archive" | Should -Not -Exist
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
            $clonedParams = $Params.Clone()
            $clonedParams.Remove('Archive')

            $testInputFile = (Join-Path $Params.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter"; UnknownParameter = 'Oops' } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8
    
            . $testScript @clonedParams
        }
        It 'no Archive folder is created' {
            "$($Params.DropFolder)\Archive" | Should -Not -Exist
        }
        It 'the input file is not moved to the archive folder' {
            $testInputFile | Should -Exist
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                (&$MailAdminParams) -and 
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

        $clonedParams = $Params.Clone()
        $clonedParams.DropFolder = $testScriptFolder

        . $testScript @clonedParams

        Should -invoke Start-Job -Exactly 1 -Scope Describe -ParameterFilter {
            ($LiteralPath -eq $testScriptToExecute) -and
            ($ArgumentList[0] -eq 'MyCustomPrinter') -and
            ($ArgumentList[1] -eq 'red')
        }
    }
}

Describe 'multiple input files for the same script' {
    BeforeAll {
        $testInputFile1 = (Join-Path $Params.dropFolder 'inputFile1.json')
        $testInputFile2 = (Join-Path $Params.dropFolder 'inputFile2.json')
     
        @{ PrinterName = "Printer1" } | 
        ConvertTo-Json | Out-File $testInputFile1 -Encoding utf8
        @{  PrinterName = "Printer2" } | 
        ConvertTo-Json | Out-File $testInputFile2 -Encoding utf8
      
        . $testScript @Params
    }
    It 'should invoke Start-Job with a unique script name for a unique log folder creation by the child script' {
        Should -invoke Start-Job -Exactly 1 -Scope Describe -ParameterFilter {
            ($LiteralPath -eq $testScriptToExecute) -and
            ($ArgumentList[0] -eq 'Printer1') -and
            ($ArgumentList[1] -eq 'red') -and
            ($ArgumentList[2] -eq 'Get printers (BNL)')
        }
        Should -invoke Start-Job -Exactly 1 -Scope Describe -ParameterFilter {
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

            $testInputFileExclude = (Join-Path $Params.dropFolder 'file.txt')
         
            (New-Item -Path $testInputFileExclude -Force -ItemType File -EA Ignore).FullName
    
            . $testScript @Params
        }
        It 'it is ignored and left in the drop folder' {
            $testInputFileExclude | Should -Exist
            Should -Not -Invoke ConvertFrom-Json -Scope Context
        }
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
    }
    Context 'it is not a valid .json file' {
        BeforeAll {
            $testInputFile = (Join-Path $Params.dropFolder 'inputFile.json')
     
            (New-Item -Path $testInputFile -Force -ItemType File -EA Ignore).FullName

            "NotJsonFormat ;!= " | Out-File $testInputFile -Encoding utf8

            . $testScript @Params
        }
        It 'it is moved to the archive folder in the drop folder' {
            "$($Params.DropFolder)\inputFile.json" | Should -Not -Exist
            "$($Params.DropFolder)\Archive\inputFile.json" | Should -Exist
        }
        It 'an error file is created in the archive folder' {
            "$($Params.DropFolder)\Archive\inputFile - ERROR.txt" | 
            Should -Exist
        }
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
    }
    Context 'the user used a parameter that is not available in the scriptMapping table' {
        BeforeAll {
            $testInputFile = (Join-Path $Params.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter"; UnknownParameter = 'Oops' } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8
    
            . $testScript @Params
        }
        It 'it is moved to the archive folder in the drop folder' {
            "$($Params.DropFolder)\inputFile.json" | Should -Not -Exist
            "$($Params.DropFolder)\Archive\inputFile.json" | Should -Exist
        }
        It 'an error file is created in the archive folder' {
            "$($Params.DropFolder)\Archive\inputFile - ERROR.txt" | 
            Should -Exist
        }
        It 'the error file contains the incorrect parameter name' {
            Get-Content "$($Params.DropFolder)\Archive\inputFile - ERROR.txt" -Raw | Should -BeLike "*parameter 'UnknownParameter' is not accepted*"
        }
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
    }
    Context 'when Start-Job fails because of an incorrect parameter in the input file' {
        BeforeAll {
            $testInputFile = (Join-Path $Params.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter" } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8
    
            $clonedScriptSettings = $testScriptSettings.Clone()
            $clonedScriptSettings.PrinterColor = $null
            
            $clonedParams = $Params.Clone()
            $clonedParams.ScriptMapping = @{ 
                $testScriptFolder = $clonedScriptSettings 
            }

            Mock Start-Job {
                & $StartJobCommand -Scriptblock { 
                    Param (
                        [parameter(Mandatory)]
                        [int]$validParameter
                    )
                } -ArgumentList 'string'
            }
            . $testScript @Params
        }
        It 'the input file is moved to the archive folder' {
            "$($Params.DropFolder)\inputFile.json" | Should -Not -Exist
            "$($Params.DropFolder)\Archive\inputFile.json" | Should -Exist
        }
        It 'an error file is created in the archive folder' {
            "$($Params.DropFolder)\Archive\inputFile - ERROR.txt" | 
            Should -Exist
        } -Tag test
        It 'the error file contains the script parameters that are allowed' {
            Get-Content "$($Params.DropFolder)\Archive\inputFile - ERROR.txt" -Raw | Should -BeLike "*Invalid input file 'inputFile.json'*validParameter*"
        }
    } 
}

Describe 'when the informAdmin switch is used' {
    Context 'an email is send to the admin when' {
        It 'an incorrect user input file is used' {
            $testInputFile = (Join-Path $Params.dropFolder 'inputFile.json')
     
            (New-Item -Path $testInputFile -Force -ItemType File -EA Ignore).FullName

            "NotJsonFormat ;!= " | Out-File $testInputFile -Encoding utf8

            . $testScript @Params -InformAdmin

            Should -Not -Invoke Start-Job
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Invalid json input file*")
            }
        }  -tag test
        It 'a script has been launched with a valid user input file' {
            $testInputFile = (Join-Path $Params.dropFolder 'inputFile.json')
     
            @{  PrinterName = "MyCustomPrinter" } | 
            ConvertTo-Json | Out-File $testInputFile -Encoding utf8
    
            . $testScript @Params -InformAdmin

            Should -Invoke Start-Job
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Subject -eq '1 script started') -and 
                ($Message -like "*Script name*")
            }
        }
    }
}