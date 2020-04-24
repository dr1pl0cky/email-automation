######################################
########   Email Automation   ########
########     Version 1.0      ########
######################################

######################################
# NOTE: run `Instal-Module PSExcel`  #
# fist as it is a dependency in      #
# Powershell as admin                #
######################################

################################
# NOTE: Run this script inside #
# same folder as data please   #
################################


#import needed modules
Import-Module PSExcel

Function Init-Auto {
    #get working directory
    $pwd1 = pwd

    #look for excel sheets
    $excel_file = Get-ChildItem -Path $pwd1 -Include *.xlsx -File -Recurse -ErrorAction SilentlyContinue -Name

    #export xlsx to csv for easy working
    Import-XLSX -Path $pwd1\$excel_file | Export-Csv "output.csv" -NoTypeInformation

}


Function Get-Vars {
    #define headers for CSV
    $csv_headers = @("First","Last","Date","Annual","Dues","Email")

    #get csv
    $csv_file = Get-ChildItem -Path $pwd1 -Include *.csv -File -Recurse -ErrorAction SilentlyContinue -Name

    #imports CSV with headers and make var
    $person_list = Import-Csv $pwd1\$csv_file -Header $csv_headers
    
    #loop through records, this is needed so i can set each one as a variable.
    foreach ($person in $person_list) {
        Write-Host "$($person.First), $($person.Last), $($person.Date), $($person.Annual), $($person.Dues), $($person.Email)"
        
        $person_email = $($person.Email.ToString())
        $person_name = $($person.First.ToString()) + " " + $($person.Last.ToString())
        $person_dues = $pwd1.ToString() + "\" + $($person.Dues.ToString())
        $person_annual = $pwd1.ToString() + "\" + $($person.Annual.ToString())
        $person_date = $($person.Date.ToString())
        $person_dues_name = $($person.Dues.ToString())
        $person_annual_name = $($person.Annual.ToString())

        $user = "some-email-for-here@gmail.com"
        $pass = ConvertTo-SecureString -String "MYPASSWORD" -AsPlainText -Force
        $cred = New-Object System.Management.Automation.PSCredential $user, $pass
        

        $To = $person_email
        $From = "ME <some-email-for-here@gmail.com>"
        $Attachment = $person_annual, $person_dues
        $Subject = "Email Subject"
        $Body = "Hello there: $person_name

We are emailing you as you owe us money,

Attached are your annual report and your dues document named: $($person.Dues.ToString()) and $($person.Annual.ToString())

The Date due is: $person_date

All the best,

HOA people."

        $SmtpServer = "smtp.gmail.com"
        $Port = "587"

        Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -SmtpServer $SmtpServer -port $Port -UseSsl -Credential $cred -Attachments $Attachment


    }
}

function Clean-Up {
    #just cleans up csv
    $csv_file = Get-ChildItem -Path $pwd1 -Include *.csv -File -Recurse -ErrorAction SilentlyContinue -Name

    rm $csv_file

}
Init-Auto
Get-Vars

Clean-Up
