
# This will search for Social Security Numbers
function Get-SSN {
    Get-ChildItem  -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string "[0-9]{9}" , "[0-9]{3}[-| ][0-9]{2}[-| ][0-9]{4}"
}
# This will search for Credit Card data: Discover, MasterCard, Visa
function Get-CCards {
    Get-ChildItem  -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string "[456][0-9]{3}[-| ][0-9]{4}[-| ][0-9]{4}[-| ][0-9]{4}" 
    Get-ChildItem  -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string "[456][0-9]{15}"
}


#American Express
function Get-Amex{
    Get-childitem -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string "3[47][0-9]{4}[-| ][0-9]{6}[-| ][0-9]{5}"
    Get-childitem -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string "3[47][0-9]{2}[-| ][0-9]{4}[-| ][0-9]{4}[-| ][0-9]{3}"
    Get-ChildItem -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string "3[47][0-9]{13}","3[47][0-9]{2}[-| ][0-9]{6}[-| ][0-9]{5}"
}

#Array of strings to match
$LiteralStrings = @("account" , "medical" , "driver" , "patient" , "maiden" , "birth" , "password" , "username", "social", "credit", "passport")

#Function that finds indicator of PII
function Find-Indicators{
    foreach($n in $LiteralStrings){
        Get-childitem -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string -AllMatches $n
    }

            $files    = Get-Childitem $path .\* -Force -Include *.docx,*.doc -Recurse | Where-Object { !($_.psiscontainer) }

            # Loop through all *.doc files in the $path directory
            Foreach ($file In $files){
                $application = New-Object -com word.application
                $application.visible = $False
                $document = $application.documents.open($file.FullName,$false,$true)
                $range = $document.content
                foreach($n in $LiteralStrings){
                    If($document.content.text -like "*$n*"){ 
                        Write-Host "[+] Located File with possible indicator"
                        Write-Host "[+]" $file
                        Write-Host "[+] Indicator =" $n
                }
            }
                $document.close()
                $application.quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) 
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($application) 
        }
    }




#Searches directories and sub directories for PII in the name of the file
function Find-Files{
    
    foreach($n in $LiteralStrings){
            Get-childitem -Recurse -Path .\* -Force | Where-Object { !$PsIsContainer -and [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -match $n }
    }

}
function Get-PDF{


}