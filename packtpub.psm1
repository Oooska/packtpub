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
        Import-Module $env:USERPROFILE+"\Documents\WindowsPowerShell\Modules\packtpub.psm1"
        $s = Read-PacktPubSettings
        Get-PacktPubFreeBook -email $s.email -password $s.password -download $s.download  -format $s.format -saveDir $s.saveDir
    }
}

#Get-PacktPubFreeBook logs into the specified PacktPub account, claims the free daily book, and, if $download=$true, downloads
#the book in the specified format to $savedir.
function Get-PacktPubFreeBook {
    param (
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
        [string]$saveDir=$env:USERPROFILE+"\Documents\ebooks"
    )


    #Get Page, parse login form and claim URL
    $page = Invoke-WebRequest https://www.packtpub.com/packt/offers/free-learning -SessionVariable sv
    $form = getLoginForm -page $page -email $email -password $password
    $claimURL = getFreeBookClaimURL $page

    #Login
    $loggedInPage = Invoke-WebRequest https://www.packtpub.com/ -Method Post -Body $form -WebSession $sv

    #Claim free book - throws 404 error if not logged in successfully
    $claimBookPage = Invoke-WebRequest $claimURL -WebSession $sv

    if($claimBookPage){
        "Successfully claimed " + (getFreeBookTitle $page)
    } else {
        throw "Error claiming book" + (getFreeBookTitle $page)
        return 1
    }

    #Download the book in the specified format
    if($download){
        if($saveDir){
            Set-Location $saveDir
        }
        $bookID = getFreeBookID $claimURL
        Invoke-WebRequest ("https://www.packtpub.com/ebook_download/" + $bookID + "/" + $format) -WebSession $sv -OutFile ((getFreeBookTitle $page)+"."+ $format)
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
    "Download: "
    $download
    "{0}`t{1}`t{2}`t{3}`t{4}" -f $email,$password,$download,$format,$saveDir |`
      Set-Content -Path $env:APPDATA\packtpub.dat
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

Export-ModuleMember -Function Register-PacktPubFreeBookJob
Export-ModuleMember -Function Get-PacktPubFreeBook
Export-ModuleMember -Function Read-PacktPubSettings

<#
function Get-AllPacktPubBooks {
    [CmdletBinding()]
    param(
        [parameter(Mandatory=$true)]
        [string]$email,

        [parameter(Mandatory=$true)]
        [string]$password,

        [parameter(Mandatory=$true)]
        [string]$folder,

        [parameter(Mandatory=$true)]
        [string]$format #pdf, epub, mobi
    )


}


$email = "junk@oooska.com"
$password = "..."

$page = Invoke-WebRequest https://www.packtpub.com/packt/offers/free-learning -SessionVariable sv
$form = get-loginForm -page $page -email $email -password $password
$loggedInPage = Invoke-WebRequest https://www.packtpub.com/ -Method Post -Body $form -WebSession $sv
$format = "pdf"
$listpage = Invoke-WebRequest https://www.packtpub.com/account/my-ebooks -WebSession $sv 

$list = $listpage.ParsedHtml.getElementsByTagName("div") | where { $_.className -like "product-line*" }
$list[0]

Invoke-WebRequest https://www.packtpub.com/ebook_download/8064/pdf -WebSession $sv -OutFile "book.pdf"
#>