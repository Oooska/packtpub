<#
    .SYNOPSIS

    Logs into PacktPub with the supplied email address and password.

    .DESCRIPTION

    New-PacktPubWebSession takes the supplied username and password and attempts to log into PacktPub.com.
    If it is successful, it returns a WebSession object that contains the login cookie. 

    If it is not successful, it throws an error containining the inner text from the 'messages error' class on the returned page.

    .PARAMETER email

    The email address of the PacktPub account

    .PARAMETER password

    The password of the PacktPub account

    .EXAMPLE

    $ws = New-PacktPubWebSession "email@somewhere.com" "mysupersecretpassword"

    .NOTES

    You will not need to call this function yourself. It is exported for use by Register-PacktPubFreeBookJob. 
#>
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
    $errs = $loggedInPage.ParsedHtml.body.getElementsByClassName('messages error')

    if ($errs.length -eq 0){
        return $sv
    }

    throw "Unable to log into PacktPub. Error: "+$errs[0].innerText
}

<#
    .SYNOPSIS

    Returns a list of all books on the specified PacktPub account.

    .DESCRIPTION

    Get-PacktPubBooks returns a list of books out to the pipeline. The list consists of a hash table containing the title
    of the book, the PacktPub ID of the book, and the WebSession value that allows access to the book.

    .PARAMETER ws
    The websession value to use when accessing the PacktPub servers. If it is not supplied, a prompt will request the email address 
    password from the user.

    .EXAMPLE

    $books = Get-PacktPubBooks
#>
function Get-PacktPubBooks {
    param (
        [parameter()]
        $ws=(New-PacktPubWebSession)
    )

    process {
        $booksPage = Invoke-WebRequest https://www.packtpub.com/account/my-ebooks -WebSession $ws
        $booksElemArray = $booksPage.ParsedHtml.body.getElementsByClassName('product-line')
        
        foreach($elem in $booksElemArray){
            $title = $elem.getAttribute('title')
            $id = $elem.getAttribute('nid')
            if($title -ne $null -and $id -ne $null){
                $title = $title.Substring(0, $title.length-8) #remove " [ebook]" from end
                Write-Output (Select-BookObject -id $id -title $title -ws $ws)
            }
        }
    }
}

<#
    .SYNOPSIS

    Claims the free book of the day on the specified account

    .DESCRIPTION

    Get-PacktPubFreeBook logs into the specified PacktPub account and claims the free daily book 

    .PARAMETER ws
    The websession value to use when accessing the PacktPub servers. If it is not supplied, a prompt will request the email address 
    password from the user.

    .EXAMPLE

    Get-PacktPubClaimFreeBook
    
    To claim and download the book: 
    Get-PacktPubDownloadBooks | Get-PacktPubDownloadBooks
#>
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
        return Select-BookObject $bookID $title $ws
    }
    
    throw "Error claiming book `"$title`" ($bookID)"
}

<#
    .SYNOPSIS

    Downloads the books specified by Get-PacktPubBooks or Get-PacktPubClaimFreeBook

    .DESCRIPTION

    Get-PacktPubDownloadBooks downloads the books that are passed in via the pipeline. Books can be 
    downloaded as PDF, mobi, or epub, and saves the files to the current working directory unless
    saveDir is specified. The filename  will consist of the title of the book, and the book id in parenthesis. 
    If the title of the book has characters that are not valid for a filename (e.g. a colon ':'), they are stripped.

    .PARAMETER books

    The book information supplied by Get-PacktPubBooks or Get-PacktPubClaimFreeBook

    .PARAMETER format

    The format to download the books in. Currently supported values are PDF, mobi and epub.

    .PARAMETER saveDir

    The directory to save the book(s) to.

    .EXAMPLE

    To download the claimed book: 
    Get-PacktPubDownloadBooks | Get-PacktPubDownloadBooks

    To download all the books on your account:
    Get-PacktPubBooks | Get-PacktPubDownloadBooks
#>
function Get-PacktPubDownloadBooks {
    param (
        [parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true)]
        $books,

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
            $bookTitle = (removeInvalidFilenameChars $book.title)
            Invoke-WebRequest "https://www.packtpub.com/ebook_download/$bookID/$format" -WebSession $book.ws -OutFile "$bookTitle ($bookID).$format"
        }
    }
}

<#
    .SYNOPSIS

    Creates a task to automatically acquire the free book of the day and download it at a specific time.

    .DESCRIPTION

    Register-PacktPubFreeBookJob creates a re-occuring task to acquire the free book, and then optionally may 
    download the book in the specified format. Credentials are saved as plaintext in a file in %appdata%. 

    The task will run when this function is first called, and then at 8PM local time there after.

    By default, the computer will wake if it is sleeping or hibernating. Set WakeToRun to $false to override this feature.

    If you want to cancel the task, you can use Get-ScheduledJob to view the list of powershell jobs, and Unregister-ScheduledJob to remove it

    .PARAMETER email

    The email address to use to log into packtpub.

    .PARAMETER password

    The password to use to log into packtpub.

    .PARAMETER download

    A bool indicating whether the task should download the book after adding it to the account.

    .PARAMETER format

    The format to download the books in. Currently supported values are PDF, mobi and epub.

    .PARAMETER saveDir

    The directory to save the book to. Default location is "Documents\ebooks" for the current user.

    .PARAMETER time

    The time to run the scheduled task. By default this runs at 8PM.

    .PARAMETER WakeToRun

    A boolean switch indicating whether the task should wake the computer from sleep/hibernate. The default is true.

#>
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

#Select-BookObject takes the id, title, and websession required to download a specific book and returns
#a table-formatted object to be used with out-gridview or similar
function Select-BookObject {
    param(
        [parameter()]
        $id,

        [parameter()]
        $title,

        [parameter()]
        $ws
    )

    return Select-Object -InputObject @{id = $id; title = $title; ws = $ws}`
            -Property @{Label="id"; Expression={$_.id}},`
                      @{Label="title"; Expression={$_.title}},`
                      @{Label="ws"; Expression={$_.ws}}
}

#removeInvalidChars takes a string, and removes characters that can't
#be used in a filename
function removeInvalidFilenameChars {
    param(
        [parameter()]
        $filename
    )
   $arrInvalidChars = [System.IO.Path]::GetInvalidFileNameChars() 
   
   #'[' and ']' are not illegal characters, but Invoke-WebRequest doesn't like them
   return $filename.Split($arrInvalidChars) -join "" -replace '[\[]','(' -replace '[\]]',')'  #Probably a more efficient way to do this
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
Export-ModuleMember -Function Get-PacktPubBooks
Export-ModuleMember -Function Get-PacktPubClaimFreeBook
Export-ModuleMember -Function Get-PacktPubDownloadBooks
Export-ModuleMember -Function Register-PacktPubFreeBookJob
Export-ModuleMember -Function Read-PacktPubSettings