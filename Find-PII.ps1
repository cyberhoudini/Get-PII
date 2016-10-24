
#Strings of indicators array
$LiteralStrings = @("account" , "medical" , "driver" , "patient" , "maiden" , "birth" , "password" , "username", "social", "credit", "passport")
#Social Security Numbers regex array
$ssn = @("[0-9]{9}" , "[0-9]{3}[-| ][0-9]{2}[-| ][0-9]{4}")
#crdit card numbers regex array
$creditnumbers = @("[456][0-9]{3}[-| ][0-9]{4}[-| ][0-9]{4}[-| ][0-9]{4}" , "[456][0-9]{15}" , "3[47][0-9]{4}[-| ][0-9]{6}[-| ][0-9]{5}" , "3[47][0-9]{2}[-| ][0-9]{4}[-| ][0-9]{4}[-| ][0-9]{3}" , "3[47][0-9]{13}" , "3[47][0-9]{2}[-| ][0-9]{6}[-| ][0-9]{5}")

#Function that finds indicator of PII
function Find-Indicators{
        #function that parses through text files
        function textfiles{
            foreach($n in $LiteralStrings){
                Get-childitem -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string -AllMatches "$n"

            }
            foreach ($n in $creditnumbers){
                Get-ChildItem -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string $n
            }
            foreach ($n in $ssn){
                $results = Get-ChildItem -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string $n
            }
            
        
        }
        #function to parse through word files
        function wordfiles{
                $files = Get-Childitem $path .\* -Force -Include *.docx,*.doc -Recurse | Where-Object { !($_.psiscontainer) }

                # Loop through all *.doc files in the $path directory
                Foreach ($file In $files){
                    $application = New-Object -com word.application
                    $application.visible = $False
                    $document = $application.documents.open($file.FullName,$false,$true)
                    $range = $document.content
                    foreach($n in $LiteralStrings){
                        If($document.content.text -like "*$n*"){ 
                        Write-Host "[+][+] Located File With Possible Indicator [+][+]" -foregroundcolor "yellow"
                        Write-Host "[+]" $file -foregroundcolor "green"
                        Write-Host "[+] Indicator =" $n -foregroundcolor "red"
                    }
                }
                foreach ($x in $ssn){   
                    if($document.content.text -match "$x"){
                    Write-Host "[+][+] Found Possible Social Security Number [+][+]" -foregroundcolor "yellow"
                    Write-Host "[+]" $file -foregroundcolor "green" 
                    }
                }
                foreach ($x in $creditnumbers){   
                    if($document.content.text -match "$x"){
                    Write-Host "[+][+] Found Possible Credit Card Number [+][+]" -foregroundcolor "yellow"
                    Write-Host "[+]" $file -foregroundcolor "green" 
                    }
                }
                    $document.close()
                    $application.quit()
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) 
                    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($application) 
            }
        }


        function exelfiles{
            $files = Get-Childitem $path .\* -Force -Include *.xls,*.xlsm,*.xlsx -Recurse | Where-Object { !($_.psiscontainer) }
            foreach ($file in $files){
                $application = New-Object -com excel.application
                $application.visible = $false
                $document = $application.workbooks.open($file.FullName,$false,$true)
                $range = $document.content
                foreach ($n in $LiteralStrings){
                    if([bool]$application.Cells.Find("*$n*")){
                        Write-Host "[+][+] Located File With Possible Indicator [+][+]" -foregroundcolor "yellow"
                        Write-Host "[+]" $file -foregroundcolor "green"
                        Write-Host "[+] Indicator =" $n -foregroundcolor "red"
                    }
                }

                foreach ($x in $ssn){   
                    if($application.value.compareto("$x")){
                    Write-Host "[+][+] Found Possible Social Security Numbers [+][+]" -foregroundcolor "yellow"
                    Write-Host "[+]" $file -foregroundcolor "green" 
                    }
                }
                foreach ($x in $creditnumbers){   
                if($application.value.compareto("$x")){
                Write-Host "[+][+] Found Possible Credit Card Numbers [+][+]" -foregroundcolor "yellow" 
                Write-Host "[+]" $file -foregroundcolor "green"
                }
            }
                
                $document.close()
                $application.quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($document) 
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($application)

            }
        }
    

textfiles
wordfiles
exelfiles
}




#Searches directories and sub directories for PII in the name of the file
function Find-Files{
    
    foreach($n in $LiteralStrings){
            Get-childitem -Recurse -Path .\* -Force | Where-Object { !$PsIsContainer -and [System.IO.Path]::GetFileNameWithoutExtension($_.Name) -match $n }
    }

}
function Get-PDF{


}