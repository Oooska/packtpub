function New-PacktPubWebSession {
    param(
        [parameter(Mandatory=$true)]
        [string]$email,
        
        [parameter(Mandatory=$true)]
        [string]$password
    )

    $page = Invoke-WebRequest https://www.packtpub.com/packt/offers/free-learning -SessionVariable sv
    $form = getLoginForm -page $page -email $email -password $password

    $loggedInPage = Invoke-WebRequest https://www.packtpub.com/ -Method Post -Body $form -WebSession $sv

    return $sv
}


#Get-PacktPubFreeBook logs into the specified PacktPub account, claims the free daily book, and, if $download=$true, downloads
#the book in the specified format to $savedir.
function Get-PacktPubClaimFreeBook {
    param (
        [parameter()]
        $ws=(New-PacktPubWebSession)
    )

    #Get Page, parse login form and claim URL
    $freeBookPage = Invoke-WebRequest https://www.packtpub.com/packt/offers/free-learning -WebSession $ws
    $title = getFreeBookTitle $freeBookPage
    $claimURL = getFreeBookClaimURL $freeBookPage
    $bookID = getFreeBookID $claimURL
      
    #Claim free book - throws 404 error if not logged in successfully
    $claimBookPage = Invoke-WebRequest $claimURL -WebSession $ws
    

    if($claimBookPage){
        return @{
            ws = $ws
            id = $bookID
            title = $title
        }
    }
    
    throw "Error claiming book `"$title`" ($bookID)"
}

function Get-PacktPubDownloadBooks {
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true)]
        [hashtable]$books,

        [parameter()] 
        [ValidateSet("pdf", "mobi", "epub")]
        [string]$format="pdf",

        [parameter()]
        [string]$saveDir #=$env:USERPROFILE+"\Documents\ebooks"
    )

    begin {
        if($saveDir){
            Set-Location $saveDir
        }
    }

    process {
        foreach($book in $books){
            $bookID = $book.id
            $bookTitle = $book.title
            Invoke-WebRequest "https://www.packtpub.com/ebook_download/$bookID/$format" -WebSession $book.ws -OutFile "$bookTitle ($bookID).$format"
        }
    }
}


function Register-PacktPubFreeBookJob {
    #Registers a job to call Get-PacktPubFreeBook at the specified time
    #By default, the task will wake your computer if it is hibernating or sleeping in order to accomplish this task

    #You can delete the task in taskschd.msc or Unregister-ScheduledTask
    param(
        [parameter(Mandatory=$true)]
        [string]$email,
        
        [parameter(Mandatory=$true)]
        [string]$password,

        [parameter()]
        [bool]$download=$true,

        [parameter()]
        [ValidateSet("pdf", "mobi", "epub")]
        [string]$format="pdf",

        [parameter()]
        [string]$saveDir=$env:USERPROFILE+"\Documents\ebooks",

        [parameter()]
        [string]$time="8:00 PM",

        [parameter()]
        [bool]$WakeToRun=$true
    )

    #Write the neccesary settings out to a file for the job to read back later
    writeSettings -email $email -password $password -download $download -format $format -saveDir $saveDir

    #Setup trigger and scheduled job options
    $dailyTrigger = New-JobTrigger -Daily -At $time
    if($WakeToRun){
        $options = New-ScheduledJobOption -StartIfOnBattery -ContinueIfGoingOnBattery -RunElevated -WakeToRun 
    } else {
        $options = New-ScheduledJobOption -StartIfOnBattery -ContinueIfGoingOnBattery -RunElevated 
    }

    #Schedule the job
    Register-ScheduledJob -Name PacktPubFreeBook -Trigger $dailyTrigger -ScheduledJobOption $options -RunNow -ScriptBlock {
        $s = Read-PacktPubSettings
        $book = Get-PacktPubClaimFreeBook (New-PacktPubWebSession $s.email $s.password)

        #-download $s.download  -format $s.format -saveDir $s.saveDir
        if($s.download){
            $book | Get-PacktPubDownloadBooks -format $s.format -saveDir $s.saveDir
        }
    }
}



#getLoginForm returns a hash table representing the form that must be
#POSTed to packtpub.
function getLoginForm($page, $email, $password){
    #Find login form,
    #Then check login form to get the form_build_id value
    $form = $page.Forms | where {$_.Id -eq 'packt-user-login-form'}
    $formBuildID = $form.Fields.Keys | where {$_ -like 'form-*'}

    if(!$formBuildID){
        throw "Unable to parse login form"
    }

    return @{
        email = $email
        password = $password
        op = "Login"
        form_id = "packt_user_login_form"
        form_build_id = $formBuildID
    }
}

#getFreeBookClaimURL Parses the dotd page for the claim URL
#URL is in the form https://www.packtpub.com/freelearning-claim/8064/21478
function getFreeBookClaimURL($dotdpage){
    $a = $dotdpage.Links | where { $_.href -like '/freelearning-claim/*' }
    if(!$a){
        throw "Unable to parse claim URL"
    }

    return "https://www.packtpub.com" + $a.href
}

#getFreeBookID parses the dotd page for the ID of the book that is to be claimed
function getFreeBookID($dotdClaimURL){
    #URL in form https://www.packtpub.com/freelearning-claim/8064/21478
    #Book ID is the first number after freelearning-claim (8064)
    return ($dotdClaimURL -split "/")[4]
}

#getFreeBookTitle parses the dotd page for the title of the book that is to be claimed
function getFreeBookTitle($dotdpage){
    return ($dotdpage.ParsedHtml.getElementsByTagName("div") | where{ $_.className -eq 'dotd-title'}).textContent.Trim()
}


#writeSettings saves the settings required by a scheduled job to %appdata%/packtpub.dat
#TODO: Encrypt information (particularly password)
function writeSettings($email, $password, $download, $format, $saveDir){
    "$email`t$password`t$download`t$format`t$saveDir" | Set-Content -Path $env:APPDATA\packtpub.dat
}

#Read-PacktPubSettings reads the settings from %appdata%/packtpub.dat
#and returns an associative array
function Read-PacktPubSettings(){
    $s = (Get-Content -Path $env:APPDATA\packtpub.dat).Split("`t")

    return @{
        email = $s[0]
        password = $s[1]
        download = [System.Convert]::ToBoolean($s[2])
        format = $s[3]
        saveDir = $s[4]
    }
}

Export-ModuleMember -Function New-PacktPubWebSession
Export-ModuleMember -Function Get-PacktPubClaimFreeBook
Export-ModuleMember -Function Get-PacktPubDownloadBooks
Export-ModuleMember -Function Register-PacktPubFreeBookJob
Export-ModuleMember -Function Read-PacktPubSettings