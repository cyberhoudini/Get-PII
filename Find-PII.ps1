

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
$LiteralStrings = @("account" , "medical" , "driver" , "patient" , "maiden" , "birth" , "password" , "username")

#Function that finds indicator of PII
function Find-Indicators{
    foreach($n in $LiteralStrings){
        Get-childitem -rec | ?{ findstr.exe /mprc:. $_.FullName } | select-string -AllMatches $n
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

