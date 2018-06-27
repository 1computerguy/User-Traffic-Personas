<#
.SYNOPSIS
    This is a PowerShell script to be used with all of the dot sourced scripts to simulate
    real user activity in a Windows environment.

.DESCRIPTION
    This brings together several other scripts in an attempt to simulate real user activity.
    The activity involves functions like opening/closing documents, printing documents, sending/
    receiving email, creating MS Office documents, and surfing the "web" through Internet Explorer.
    The script uses multhreaded operations in order to allow script functions to run simultaneously
    similar to how a real user might use his/her computer.  Each user function has a built-in stop
    method that stops that particular function at some point to simulate the user no longer
    performing that function (this prevents a single user from printing or creating hundreds of
    documents in the course of a short time frame).  Each user will only be logged in for a random
    duration between 30 min and 2 hours at which point they will be logged out and another user
    automatically logged into the system.

.PARAMETER 
    None

.EXAMPLE
    Script is intended to be run on login
        .\User-Activity.ps1

.NOTES
    Revision History
        05/01/2018 : Bryan Scarbrough - Created
#>

####
##
##  From Trim-Length.ps1
##
####
function Trim-Length {
    param (
        [parameter(Mandatory=$True,ValueFromPipeline=$True)] [string] $Str
      , [parameter(Mandatory=$True,Position=1)] [int] $Length
    )
        $Str[0..($Length-1)] -join ""
}


####
##
##  From Get-RandText.ps1
##
####
# Get-RandText function to generate randomly selected text from the Enron email_corpus
# by default
function Get-RandText ($message_type,
                        $message_part,
                        $company_replace
                        ) {
    
    # Using system.io.file::readalllines as opposed to get-content for speed and memory consumption
    $count = ([System.IO.File]::ReadAllLines("C:\scripts\$message_type.txt")).count
    
    # Go through loop at least once in order to grab email subject or body depending on 
    # $message_type variable defined above
    $StopLoop = $false
    do {
        try {
            if ($message_part -eq "body") {
                # Grab text from randomly selected path and replace any "Enron" reference with
                # $company_replace data to customize message content
                [string]$message = (get-content (([System.IO.File]::ReadAllLines("C:\scripts\$message_type.txt"))[(get-random -minimum 1 -maximum $count)])  | select-object -skip 7 | ? {$_ -ne ""}) | % { $_ -replace '[Ee][Nn][Rr][Oo][Nn]',$company_replace } -ErrorAction SilentlyContinue
            
            # If $message_part is "subject" then only get the subject line of a randomly selected
            # message
            } elseif ($message_part -eq "subject") {
                [string]$message = (get-content (([System.IO.File]::ReadAllLines("C:\scripts\$message_type.txt"))[(get-random -minimum 1 -maximum $count)])  | select-object) | ? { $_ -match 'Subject:' } | select -First 1 | % { $_ -replace 'Subject:', '' }
            }
            
            # If data is obtained without error then exit the loop
            $StopLoop = $true
        }
        # Not worried about error, just want to continue until finished
        catch {
        }
    } while ($StopLoop -eq $false)
    
    return $message
}

# Example usage - uncomment below to test
#Get-RandText -message_type "internal" -message_part "subject" -company_replace "US Army"
#Get-RandText "internal" "body" "US_Army"


####
##
##  From Send-Email.ps1
##
####
# Get DNS MX Record for associated domain
function Get-MX ($lookup) {
    [string]$record = nslookup -type=mx $lookup
    $mx = $record.split() | select-string -pattern $lookup
    return $mx[1]
}

# Get accounts to send emails to
function Get-Acct ($internal,
                    $external,
                    $max,
                    $internal_acct="C:\scripts\internal_users.txt",
                    $external_acct="C:\scripts\external_users.txt"
                    ) {

    $recipient = New-Object System.Collections.Generic.List[System.Object]
    [System.Collections.ArrayList]$internal_recipients = get-content $internal_acct
    [System.Collections.ArrayList]$external_recipients = get-content $external_acct

    # Domains to use for sending email (all accounts are in same domain for this script)
    $domain = "3bct4id.army.mil"
    
    # Make sure that at least one domain selection is made
    # DEFAULT is internal
    if (!($internal -Or $external)) {
        $internal = 1
    }

    # If both domains (external and internal) are selected, then distribute maximum
    # recipients between the two domains
    $int = 0
    $ext = 0
    if (($internal -And $external) -And $max -eq 1) {
        $max = 2
        $int = get-random -minimum 1 -maximum $max
        $ext = $max - $int
    } elseif ($internal -And $external) {
        $int = get-random -minimum 1 -maximum $max
        $ext = $max - $int
    } elseif ($internal -And !$external) {
        $int = $max
    } elseif ($external -And !$internal) {
        $ext = $max
    }
    
    # If sending to an internal email address then randomly select a user from the 
    # $internal_recipients text file
    $username = $(Get-WMIObject -class Win32_ComputerSystem | select UserName).username.split("\")[1]
    if ($internal) {
        $internal_recipients.remove($username)
        1..$int | % {
            $name = $internal_recipients | get-random
            $recipient.add("$name@$domain")
            $internal_recipients.remove($name)
        }
    }
    
    # If sending to an external email address then randomly select a user from the
    # $external_recipients text file
    if ($external) {
        1..$ext | % {
            $name = $external_recipients | get-random
            #$ext_domain = $target_domains
            $recipient.add("$name@$domain")
            $external_recipients.remove($name)
        }
    }

    return $recipient
}

# Get a random file name from the c:\scripts\documents (default) folder to attach to email
function Get-Attachment ($num_attach,
                        $attachments_dir = "c:\scripts\documents"
                        ) {
    $attachments = New-Object System.Collections.Generic.List[System.Object]
    # Command to create attachments.txt files
    # get-childitem .\documents -recurse | % { $_.fullname } >> attachments.txt
    [System.Collections.ArrayList]$attachment_list = get-childitem $attachments_dir -recurse | % { $_.fullname }
    if (!$num_attach -eq 0) {
        1..$num_attach | % {
            $attach = $attachment_list | get-random
            $attachments.add($attach)
            $attachment_list.remove($attach)
        }
    }
    
    return $attachments
}

function Send-Email ([string]$in_usr,
                   [string]$in_or_out = "internal"
                   ) {
    
    $num_recipients = 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 2, 2, 3, 3, 4, 4, 5
    $num_attachments = 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 2
    $attach_files = Get-Attachment ($num_attachments | get-random)
    $domain = "3bct4id.army.mil"
    $from_domain = "exch1-3bct4id.3bct4id.army.mil"
    
    # If no attachments then do not declare them in the Mail_info object
    if ($attach_files) {
        $Mail_info = @{
            To = Get-Acct (get-random -maximum 2) (get-random -maximum 2) ($num_recipients | get-random)
            From = "$in_usr@$domain"
            Subject = (Get-RandText "$in_or_out" "subject" "US Army")
            Body = (Get-RandText "$in_or_out" "body" "US_Army")
            Attachments = $attach_files
            SmtpServer = $from_domain
        }
    } else {
        $Mail_info = @{
            To = Get-Acct (get-random -maximum 2) (get-random -maximum 2) ($num_recipients | get-random)
            From = "$in_usr@$domain"
            Subject = (Get-RandText "$in_or_out" "subject" "US Army")
            Body = (Get-RandText "$in_or_out" "body" "US_Army")
            SmtpServer = $from_domain
        }
    }
    
    # Send email using @mail_info data
    Send-MailMessage @Mail_info -usessl
}

# Example Usage - uncomment below to test
# Send-Email "adam.mcfarland" "3bct4id.army.mil" "internal"
# Send-Email "adamcfly" "facebook.com" "external"


####
##
##  From Open-Print-Docs.ps1
##
####
function Open-Print-Docs ($doc, $print=0) {

    if ($print) {
        
        # Open document and print.
        # TO DO: Figure out how to print already opened document (current setup will
        # open document and print then immediately close document).
        Start-Process -FilePath $doc -Verb Print -PassThru | % { sleep 30; $_ } | kill
    } else {
    
        # Open document and wait 60-120 min before closing
        Start-Process -FilePath $doc -PassThru | % { sleep (get-random -minimum 3600 -maximum 7200); $_ } | kill    
    }

}


####
##
##  From Create-Office.ps1
##
####
function Create-Office ($doc_type,
                        $data_file = $null
                        ) {
                        
    Add-type -AssemblyName office
    
    ## CREATE WORD DOCUMENT ##
    if ($doc_type -cmatch "doc") {
        ## UNCOMMENT THE BELOW LINES FOR TESTING ##
        ## $Date = (Get-Date -Format yyyyMMddHHmmss).toString()
        ## Add-type -AssemblyName office
        ## $folder = [Environment]::GetFolderPath("MyDocuments")
        
        # Create MS Word object instance
        $Word = New-Object -ComObject Word.Application
        $Word.Visible = $True
        $WDocument = $Word.Documents.Add()
        $Selection = $Word.Selection
        
        # Select the number of paragraphs and the number of "lines" per paragraph
        # These values are used in the foreach loops below
        $num_of_paragraphs = get-random -minimum 2 -maximum 8
        $num_of_lines = get-random -minimum 1 -maximum 3
        $num_of_images = get-random -minimum 0 -maximum 5

        # Iterate over the number of lines and paragraphs
        0..$num_of_paragraphs | % {
            $Selection.TypeParagraph()
            0..$num_of_lines | % {
            
                # Write text to the file
                $Selection.TypeText((Get-RandText "internal" "body" "US Army"))
                
                # Sleep between 120-300 seconds to add realistic delay
                sleep (get-random -minimum 120 -maximum 300)
            }
            # Insert images between lines of text
            if ($num_of_images -gt 0) {
                $Selection.InlineShapes.AddPicture((Get-Random (Get-ChildItem "C:\scripts\images" | % { $_.FullName })))
                $num_of_images--
            }
        }
        # Save and close the Word Document
        Save-and-Close("doc")
    }
    
    ## CREATE EXCEL SPREADSHEET ##
    if ($doc_type -cmatch "xls") {
        ## UNCOMMENT THE BELOW LINES FOR TESTING ##
        ## $Date = (Get-Date -Format yyyyMMddHHmmss).toString()
        ## Add-type -AssemblyName office
        ## $folder = [Environment]::GetFolderPath("MyDocuments")
        
        # Create MS Excel object instance
        $Excel = New-Object -ComObject Excel.Application
        $Excel.Visible = $True
        $Workbook = $Excel.Workbooks.Add()
        $Sheet = $Workbook.WorkSheets.item("sheet1")
        $Sheet.activate()
        
        # Determine maximum number of rows and columns for spreadsheet
        $max_rows = get-random -minimum 10 -maximum 100
        $max_columns = get-random -minimum 5 -maximum 15
        
        # Iterate over rows and columns and split message text into an array
        # Then insert each value of the array into a row in the spreadsheet
        1..($max_rows+1) | % {
            $data = (Get-RandText "internal" "body" "US Army").split()
            $data = $data | ?{$_}
            $row = $_
            1..($max_columns+1) | % {
                $Sheet.Cells.Item($row,$_) = $data[$_-1]
            }
            
            # Sleep between 30 seconds and 3 minutes to add realistic delay
            sleep (get-random -minimum 20 -maximum 180)
        }
        # Save and close the Excel Workbook
        Save-and-Close("xls")
    }
    
    ## CREATE POWERPOINT PRESENTATION ##
    if ($doc_type -cmatch "ppt") {
        ## UNCOMMENT THE BELOW LINES FOR PPT CREATION TESTING ##
        ## $Date = (Get-Date -Format yyyyMMddHHmmss).toString()
        ## Add-type -AssemblyName office
        ## $folder = [Environment]::GetFolderPath("MyDocuments")
        
        # Check data_file value and if null get random file from c:\scripts\csv-files
        if ($data_file -eq $null) {
            $InputFile = get-childitem 'C:\scripts\csv-files' | % {$_.fullname} | get-random
        } else {
            $InputFile = $data_file
        }
        
        # Import and parse CSV file to get values to use for slide Title and Slide Body
        $CSV_file = Import-Csv -Path "$InputFile"
        $Title = ($CSV_file | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name')[1]
        $Body =  ($CSV_file | Get-member -MemberType 'NoteProperty' | Select-Object -ExpandProperty 'Name')[0]
        
        # Create True/False values for PPT
        $MSTrue=[Microsoft.Office.Core.MsoTriState]::msoTrue
        $MsFalse=[Microsoft.Office.Core.MsoTriState]::msoFalse
        
        # Create MS PowerPoint object instance
        $PowerPoint = New-Object -ComObject Powerpoint.Application
        $PowerPoint.Visible = $MSTrue
        $Presentation = $PowerPoint.Presentations.Add()
        $SlideType = [microsoft.office.interop.powerpoint.ppSlideLayout] -as [type]
        
        # Define slide layout types to use for presentation
        $slide_type_title = $SlideType::ppLayoutTitle
        $slide_type_chart = $SlideType::ppLayoutChart
        $slide_type_text = $SlideType::ppLayoutText
        
        # Create presentation and add first slide as title type using
        # CSV Title and Body values from above
        $Slide = $Presentation.slides.Add(1,$slide_type_title)
        $Slide.Shapes.Title.TextFrame.TextRange.Text = $Title
        $Slide.shapes.item(2).TextFrame.TextRange.Text = $Body
        $Slide.BackgroundStyle = 11
        
        # Increment slide number, then iterate over CSV object creating
        # slides and populating them with the CSV file data and random
        # images from the c:\scripts\images folder
        $slide_num = 2
        $CSV_file | ForEach-Object {
            # If get-random returns a value greater than 7, then an image slide is
            # added to the presentation, otherwise create text only slides
            if ((get-random 10) -gt 7) {
                $Image = get-childitem 'c:\scripts\images\' | % {$_.fullname} | get-random
                $Slide = $Presentation.Slides.Add($slide_num, $slide_type_chart)
                $Slide.BackgroundStyle = 11
                $Slide.Shapes.title.TextFrame.TextRange.Text = $_.$Title
                $Slide.Shapes.AddPicture($Image, $MSFalse, $MSTrue, 200, 400)
                
                # Sleep from 10-100 seconds to add user realism to slide creation
                sleep (get-random -minimum 10 -maximum 100)
            } else {
                $Slide = $Presentation.slides.Add($slide_num, $slide_type_text)
                $Slide.BackgroundStyle = 11
                $Slide.Shapes.title.TextFrame.TextRange.Text = $_.$Title
                $Slide.Shapes.item(2).TextFrame.TextRange.Text = $_.$Body
                
                # Sleep from 300-1200 seconds to add user realism to slide creation
                sleep (get-random -minimum 300 -maximum 1200)
            }
            
            # Increment slide number for next slide to add
            $slide_num ++
        }
        # Save and close the PowerPoint Presentation
        Save-and-Close("ppt")
    }
}

#Create-Office "doc"
#Create-Office "xls"
#Create-Office "ppt"


function Save-and-Close ($doc_type,
                        $folder = [Environment]::GetFolderPath("MyDocuments"),
                        $filename = (Get-RandText "C:\scripts\internal" "subject" "US Army" | Trim-length 8)
                        ) {

    $Date = (Get-Date -Format yyyyMMddHHmmss).toString()
    $FullName = "$FileName - $Date.$doc_type"
    $Output = "$Folder\$FullName"

    if ($doc_type -eq "doc") {
        # Create variables for filename and folder location and save document.
        # Once saved, then exit all instances.
        $Word_Doc = [Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application')

        $Word_Doc.Documents | ? { $_.Name -like 'Document1' -or $_.Name -like 'Document2' -or $_.Name -like 'Document3'} | % { $_.SaveAs([ref]$Output,[ref]$SaveFormat::wdFormatDocument); $_.Close() }
        $Word_Doc.Quit()
        $Word_Doc = $null
    }

    if ($doc_type -eq "xls") {
        # Create variables for filename and folder location and save document.
        # Once saved, then exit all instances.
        $Excel_Doc = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')

        $Excel_Doc.ActiveWorkbook | ? { $_.Name -like 'Book1' -or $_.Name -like 'Book2' -or $_.Name -like 'Book3' } | % { $_.SaveAs($Output); $_.Close() }
        $Excel_Doc.Quit()
        $Excel_Doc = $null
    }

    if ($doc_type -eq "ppt") {
        # Create variables for filename and folder location and save document.
        # Once saved, then exit all instances.
        $PowerPoint_Doc = [Runtime.Interopservices.Marshal]::GetActiveObject('PowerPoint.Application')

        $PowerPoint_Doc.Presentations | ? { $_.Name -like 'Presentation1' -or $_.Name -like 'Presentation2' -or $_.Name -like 'Presentation3' } | % { $_.SaveAs($Output); $_.Close() }
        $PowerPoint_Doc.Quit()
        $PowerPoint_Doc = $null
    }
    
    # This condition is used to close office documents not created by the Create-Office script, and just opened from the files or attachments
    # directory.  Since they are already named and not called the generic names above for new documents, I want to save them and close the 
    # applications safely so there are no lingering issues opening and using MS Office due to unsaved or corrupted documents.
    if ($doc_type -eq "all") {
        # Save and Close MS Word
        $Word_Doc = [Runtime.Interopservices.Marshal]::GetActiveObject('Word.Application')
        $Word_Doc.Documents.Save()
        $Word_Doc.Documents.Close()
        $Word_Doc.Quit()

        # Save and Close MS Excel
        $Excel_Doc = [Runtime.Interopservices.Marshal]::GetActiveObject('Excel.Application')
        $Excel_Doc.ActiveWorkbook.Save()
        $Excel_Doc.ActiveWorkbook.Close()
        $Excel_Doc.Quit()

        # Save and Close MS PowerPoint
        $PowerPoint_Doc = [Runtime.Interopservices.Marshal]::GetActiveObject('PowerPoint.Application')
        $PowerPoint_Doc.ActivePresentation.Save()
        $PowerPoint_Doc.ActivePresentation.Close()
        $PowerPoint_Doc.Quit()
    }
    # Perform garbage collection - required to completely stop some processes and
    # clear memory buffers appropriately.
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
}


####
##
##  From Add-IETab.ps1
##
####
## UNCOMMMENT BELOW LINE AND REMOVE FUNCTION DEFINITION TO USE AS STANDALONE SCRIPT
#param([STRING][Parameter(ValueFromPipeline=$TRUE)] $URL = "about:blank", [SWITCH] $Close, [SWITCH] $Passthru)

function Open-IE ($URL="about:blank",
                  $Close=0,
                  $Passthru=0
                  ) {

    # create object "Shell.Application" and get window list
    $oWindows = (New-Object -ComObject Shell.Application).Windows

    $IEexists = $FALSE
    if ($Passthru) { $TIMESTAMP = (Get-Date).ToString() }

    # only if window present
    if ($oWindows.Invoke().Count -gt 0)
    { # check every window
        foreach ($oWindow in $oWindows.Invoke())
        {
            # only windows of Internet Explorer
            if ($oWindow.Fullname -match "IEXPLORE.EXE")
            {
                if ($Close)
                { # close tab
                    # does Internet Explorer tab match this URL?
                    if ($oWindow.LocationURL -match $URL)
                    {
                        # URL found
                        $IEexists = $TRUE
                        # close tab
                        Write-Host "Closing tab $($oWindow.LocationURL)"
                        $oWindow.Quit()
                    }
                }
                else
                { # create tab
                    if (!$IEexists)
                    { # get COM object of existing Internet Explorer
                        $oIE = $oWindow
                        $IEexists = $TRUE
                        if (!$Passthru)
                        { break }
                    }
                    if ($Passthru)
                    { # mark window to recognize as existing window later
                        $oWindow.PutProperty($TIMESTAMP, $TIMESTAMP)
                    }
                }
            }
        }
    }

    if ($IEexists)
    { # existing Internet Explorer found
        if (!$Close)
        { # add tab
            Write-Host "Creating Internet Explorer tab with URL $URL."
            # navOpenInNewTab = 0x800
            # navOpenInBackgroundTab = 0x1000
            # navOpenNewForegroundTab = 0x10000
          # create new foreground tab
          $oIE.Navigate2($URL, 0x10000)
        }
    }
    else
    {
        if ($Close)
        { # no tab to close found
            Write-Host "No Internet Explorer tab with URL $URL found."
        }
        else
        { # existing Internet Explorer found, creating new instance
            Write-Host "Creating Internet Explorer instance with URL $URL."
            $oIE = New-Object -ComObject "InternetExplorer.Application"
        $oIE.Navigate2($URL)
        while ($oIE.Busy) { sleep -Milliseconds 50 }
        $oIE.visible = $TRUE
      }
    }

    if ($Passthru -and !$Close)
    { # to return the correct window handle we have to enumerate windows once again
        # and return the object of the tab we did not mark at the beginning of the script

        # give time to rebuild window list
        Sleep 1

        # create object "Shell.Application" and get window list
        $oWindows = (New-Object -ComObject Shell.Application).Windows

        # only if window present
        if ($oWindows.Invoke().Count -gt 0)
        { # check every window
            foreach ($oWindow in $oWindows.Invoke())
            {
                # only windows of Internet Explorer
                if ($oWindow.Fullname -match "IEXPLORE.EXE")
                { # check for mark
                    if ($oWindow.GetProperty($TIMESTAMP) -eq $TIMESTAMP)
                    { # remove mark
                        $oWindow.PutProperty($TIMESTAMP, $NULL)
                    }
                    else
                    { # no mark found, this has to be the new window
                        $oIE = $oWindow
                    }
                }
            }
        }
        # return COM object
        return $oIE
    }
}

# Get currently opened tab titles
function Get-Tabs () {
    [array]$tabs = @()
    (New-Object -ComObject Shell.Application).Windows() |
      ? { $_.FullName -like '*iexplore.exe' } |
      % {
         $Title = $_.Document.title
         if (!$Title) { $Title = $_.LocationName }
         $tab = New-Object -TypeName PSObject -Property @{Title = $Title}
         $tabs += $tab
      }
      return $tabs
}

# Close all open tabs
function Close-All () {
    $ShellWindows = (New-Object -ComObject Shell.Application).Windows()
    $ShellWindows | % {
        if ($_.FullName -like "*iexplore.exe")
        {
            $_.quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($_)
        }
    }
}

#Open-IE "www.google.com"
#Open-IE "www.microsoft.com"
#Open-IE "www.bing.com"

## Example to close all tabs
# Get-Tabs | % { Open-IE ($_.Title | Trim-length 8) 1 }
## Example to close all tabs with Google as title
# Get-Tabs | Open-IE (? { $_.Title -match "google" }) 1
## Example to close all tabs and ensure application closes properly
# Close-All

# RunspacePool for use with the multi-threaded operations below. Runspaces offer a
# significantly more efficient mechanism over Start-Job or Start-Process for multithreading.
#
Function Invoke-RunspacePool {
    <#
        .SYNOPSIS
        Creates a new runspace pool, executing the ThreadBlock in multiple threads.

        .EXAMPLE
        Invoke-RunspacePool $cmd $args
    #>

    [CmdletBinding()]
    Param(
        # Script block to execute in each thread.
        [Parameter(Mandatory=$True,
                   Position=1)]
        [System.Collections.Generic.List[System.Object]]$ThreadBlock,

        # Set of arguments to pass to the thread. $threadId will always be added to this.
        [Parameter(Mandatory=$False,
                   Position=2)]
        [hashtable]$ThreadParams,

        # Maximum number of threads. Default is the number of logical CPUs on the executing machine.
        #[Parameter(Mandatory=$False,
        #           Position=3)]
        #[int]$MaxThreads,

        # Custom Functions to pass into the RunspacePool
        [Parameter(Mandatory=$False,
                   Position=3)]
        [array]$Functions,
        
        # Garbage collector cleanup interval.
        [Parameter(Mandatory=$False)]
        [int]$CleanupInterval = 2,

        # Powershell modules to import into the RunspacePool.
        [Parameter(Mandatory=$False)]
        [String[]]$ImportModules,

        # Paths to modules to be imported into the RunspacePool.
        [Parameter(Mandatory=$False)]
        [String[]]$ImportModulesPath
    )

    #if (!$MaxThreads) {
    #    $MaxThreads = ((Get-WmiObject Win32_Processor) `
    #                       | Measure-Object -Sum -Property NumberOfLogicalProcessors).Sum
    #}

    $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    if ($ImportModules) {
        $ImportModules | % { $sessionState.ImportPSModule($_) }
    }

    if ($ImportModulesPath) {
        $ImportModulesPath | % { $sessionState.ImportPSModulesFromPath($_) }
    }

    $Functions | % {
        $CustomFunctions = Get-Content function:\$_
        $FunctionEntry = New-Object System.Management.Automation.Runspaces.SessionStateFunctionEntry -ArgumentList "$_", $CustomFunctions
        $sessionState.Commands.Add($FunctionEntry)
    }
    
    $pool = [RunspaceFactory]::CreateRunspacePool(1, $ThreadBlock.Count, $sessionState, $Host)

    $pool.ApartmentState  = "STA" # Single-threaded runspaces created
    #$pool.CleanupInterval = 2 * [timespan]::TicksPerMinute

    $pool.Open()

    $jobs      = New-Object 'Collections.Generic.List[System.IAsyncResult]'
    $pipelines = New-Object 'Collections.Generic.List[System.Management.Automation.PowerShell]'
    $handles   = New-Object 'Collections.Generic.List[System.Threading.WaitHandle]'

    $ThreadBlock | % {

        $pipeline = [powershell]::Create()
        $pipeline.RunspacePool = $pool
        $block = $_.block

        if ($_.args) {
            $_.args | % {
                $pipeline.AddScript($block).AddArgument($_)
            }
        } else {
            $pipeline.AddScript($block) | Out-Null
        }
        
        $params = @{ 'threadId' = ($ThreadBlock.IndexOf($_) + 1) }

        if ($ThreadParams) {
            $params += $ThreadParams
        }

        $pipeline.AddParameters($params) | Out-Null

        $pipelines.Add($pipeline)

        $job = $pipeline.BeginInvoke()
        $jobs.Add($job)

        $handles.Add($job.AsyncWaitHandle)
    }

    while ($pipelines.Count -gt 0) {

        $h = [System.Threading.WaitHandle]::WaitAny($handles)

        $handle   = $handles.Item($h)
        $job      = $jobs.Item($h)
        $pipeline = $pipelines.Item($h)

        $result = $pipeline.EndInvoke($job)

        ### Process results here
        if ($PSBoundParameters['Verbose'].IsPresent) { Write-Host "" }
        Write-Verbose "Pipeline state: $($pipeline.InvocationStateInfo.State)"
        if ($pipeline.HadErrors) {
            $pipeline.Streams.Error.ReadAll() | % { Write-Error $_ }
        }
        $result | % { Write-Verbose $_ }

        $handles.RemoveAt($h)
        $jobs.RemoveAt($h)
        $pipelines.RemoveAt($h)

        try {
            $handle.Dispose()
        } catch {
            break
        }
        $pipeline.Dispose()
    }

    $pool.Close()
}


####
##
##  From Get-Actions.ps1
##
####
function Get-Actions ( $open_doc,
                       $create_doc,
                       $pref_doc,
                       $print_doc,
                       $check_email,
                       $sr_email,
                       $web,
                       $application,
                       $url_persona
                       ) {

    # Setup some variables to use with the Invoke-RunspacePool function for multi-threaded operations
    $args = @{}
    $activities = New-Object System.Collections.ArrayList

    # Open and close office documents
    if ($open_doc -eq 1) {
    
        # Run as separate script thread
        #$wshell = New-Object -ComObject Wscript.Shell
        #$wshell.Popup("OPEN DOCS...",0,"Done",0x1)
        
        $open = {

            # Import Open-Print-Docs.ps1 script into job instances
            $open_docs = $true
            while ( $open_docs ) {
            
                # Check My Documents folder for any available created documents.  If available
                # then add them to the potential open queue based on the random count selected
                # in the if statement below.  Otherwise use the c:\scripts\attachments folder
                $Documents_dir = "C:\scripts\documents"
                $MyDocsFolder = [Environment]::GetFolderPath("MyDocuments")
                $MyDocsIsEmpty = (Get-Childitem $MyDocsFolder | Measure-Object).count
                if (($MyDocsIsEmpty -ne 0) -and ((get-random 10) -gt 7)) {
                    $doc = (Get-ChildItem $MyDocsFolder -recurse | % { $_.fullname }) | get-random
                } else {
                    $doc = (Get-Childitem $documents_dir -recurse | % { $_.fullname }) | get-random
                }

                # Sleep between 1-2 hrs between open jobs.  If random value is
                # between 4050 and 4020 then user will stop opening documents.
                $sleep_time = (get-random -minimum 3600 -maximum 7200)
                sleep $sleep_time
                Open-Print-Docs "$doc"                

                if (($sleep_time -lt 4050) -and ($sleep_time -gt 4020)) {
                   $open_docs = $false
                }
            }
        }
        $activities.add( @{block = $open; args = @()} )
    }
    
    # Create new Microsoft Office documents randomly selected from
    # the $doc_type array below
    if ($create_doc -eq 1) {
    
        #$wshell = New-Object -ComObject Wscript.Shell
        #$wshell.Popup("CREATE DOCS...",0,"Done",0x1)
        $create = {
            param ($pref)
            # Import Create-Office.ps1 for use within job instances
            $create_docs = $true
            while ($create_docs) {
                if ($pref -eq 0) {
                    #write "NO PREFERENCE"
                    $doc_type = "doc", "xls", "ppt"
                } elseif ($pref -eq "ppt") {
                    #write "PPT PREFERENCE"
                    $doc_type = "ppt", "ppt", "ppt", "doc", "xls"
                } elseif ($pref -eq "doc") {
                    #write "DOC PREFERENCE"
                    $doc_type = "doc", "doc", "doc", "ppt", "xls"
                } elseif ($pref -eq "xls") {
                    #write "XLS PREFERENCE"
                    $doc_type = "xls", "xls", "xls", "doc", "ppt"
                }

                # Sleep between 120-180 min between create jobs.  If random value is
                # between 5050 and 5020 then user will stop creating docs.
                $sleep_time = (get-random -minimum 7200 -maximum 10800)
                sleep $sleep_time
                                
                Create-Office ($doc_type | get-random)

                if (($sleep_time -lt 8050) -and ($sleep_time -gt 8020)) {
                   $create_docs = $false
                }
            }
        }
        $activities.add( @{block = $create; args = @($pref_doc)} )
    }
    
    # Print existing documents located either in the user's My Documents folder, or
    # from the "c:\scripts\attachments" folder - documents are randomly selected
    if ($print_doc -eq 1) {

        # Run as separate script thread
        #$wshell = New-Object -ComObject Wscript.Shell
        #$wshell.Popup("PRINT DOCS...",0,"Done",0x1)
        $print = {
        
            # Import Open-Print-Docs.ps1 script into job instances
            $print_docs = $true
            while ($print_docs) {
            
                # Check My Documents folder for any available created documents.  If available
                # then add them to the potential print queue based on the random count selected
                # in the if statement below.  Otherwise use the c:\scripts\attachments folder
                $Documents_dir = "C:\scripts\documents"
                $MyDocsFolder = [Environment]::GetFolderPath("MyDocuments")
                $MyDocsIsEmpty = (Get-Childitem $MyDocsFolder | Measure-Object).count
                if (($MyDocsIsEmpty -ne 0) -and ((get-random 10) -gt 7)) {
                    $doc = (Get-ChildItem $MyDocsFolder -recurse | % { $_.fullname }) | get-random
                } else {
                    $doc = (Get-Childitem $documents_dir -recurse | % { $_.fullname }) | get-random
                }

                # Sleep between 120-180 min between printing jobs.  If random value is
                # between 4050 and 4020 then user will stop printing.
                $sleep_time = (get-random -minimum 7200 -maximum 10800)
                sleep $sleep_time
                Open-Print-Docs "$doc" 1
                
                if (($sleep_time -lt 4050) -and ($sleep_time -gt 4020)) {
                   $print_docs = $false
                }
            }
        }
        $activities.add( @{block = $print; args = @()} )
    }
    
    # User only "checks" email which means they only open MS Outlook.
    # TO DO: Have user open new emails, select different folders and
    # move messages from one folder to another
    if ($check_email -eq 1) {
    
        # Run as separate script thread
        # Open MS Outlook to simulate user checking email
        #$wshell = New-Object -ComObject Wscript.Shell
        #$wshell.Popup("CHECK EMAIL...",0,"Done",0x1)
        $check = {
            $Outlook = New-Object -ComObject Outlook.Application
            $Namespace = $Outlook.GetNamespace("MAPI")
            $Folder = $Namespace.GetDefaultFolder("olFolderInbox")
            $Explorer = $Folder.GetExplorer()
            $Explorer.Display()

            $check = $true
            while ($check) {
                $sleep_time = (get-random -minimum 1800 -maximum 7200)
                $Namespace.SendAndReceive($true)
                sleep $sleep_time
            }
        }
        $activities.add( @{block = $check; args = @()} )
    }
    
    # User sends and receives email
    if ($sr_email -eq 1) {
        
        # Run as separate script thread
        # Use the Send-Email function to send email as currently logged in user
        #$wshell = New-Object -ComObject Wscript.Shell
        #$wshell.Popup("SEND RECEIVE EMAIL...",0,"Done",0x1)
        $sendreceive = {
        
            # Open MS Outlook to simulate user checking email
            $Outlook = New-Object -ComObject Outlook.Application
            $Namespace = $Outlook.GetNamespace("MAPI")
            $Folder = $Namespace.GetDefaultFolder("olFolderInbox")
            $Explorer = $Folder.GetExplorer()
            $Explorer.Display()
            
            $Sender_name = $(Get-WMIObject -class Win32_ComputerSystem | select UserName).username.split("\")[1]
            
            # Begin sending emails.  Use get-random to determine if email is sent internally
            # - through Exchange, or externally - through associated user's external username
            # and domain
            $sr = $true
            while ($sr) {
                # Sleep between 2.5 and 4 hrs between email jobs.  If random value is
                # between 4000 and 4100 then user will stop sending emails.
                $sleep_time = (get-random -minimum 9000 -maximum 14400)
                sleep $sleep_time
                
                Send-Email $Sender_name "internal"
                $Namespace.SendAndReceive($true)
                
                if (($sleep_time -lt 10100) -and ($sleep_time -gt 10000)) {
                   $sr = $false
                }
            }
        }
        $activities.add( @{block = $sendreceive; args = @()} )
    }
    
    # User will open Internet Explorer and open/close various websites in the browser
    if ($web -eq 1) {
    
        #$wshell = New-Object -ComObject Wscript.Shell
        #$wshell.Popup("SURF WEB...",0,"Done",0x1)
        $internet = {
            param ($urls_to_surf)
            # Import Trim-Length.ps1 and Add-IETab.ps1 scripts for use within
            # job instance
            $surf = $true
            while ($surf) {
                # If get-random returns value greater than 7 then browser tabs will be
                # enumerated and random tabs closed.  Otherwise new tabs will be opened.
                do {
                    $link = get-random (Import-Csv C:\scripts\urls.csv | % { $_.$urls_to_surf })
                } while ( $link -eq '' )
                
                if ((get-random 10) -lt 7) {
                    Open-IE $link
                } else {
                    Get-Tabs | Open-IE (? { $_.Title -match ($_.Title | Trim-Length 8) }) 1
                }
                
                # Wait from 20-60 min between web actions.  If sleep time is between
                # 2000 and 2100, then stop surfing and close Internet Explorer for
                # some time between 
                sleep (get-random -minimum 1200 -maximum 3600)
                if (($sleep_time -lt 2100) -and ($sleep_time -gt 2000)) {
                    Close-All
                    sleep $sleep_time
                } else {
                    sleep $sleep_time
                }
            }
        }
        $activities.add( @{block = $internet; args = @($url_persona)} )
    }

    if ($application -ne 0) {
    
        #$wshell = New-Object -ComObject Wscript.Shell
        #$wshell.Popup("OPEN PROGRAM...",0,"Done",0x1)
        $program = {
            start $application
        }
        $activities.add( @{block = $program; args = @()} )
    }
    
    $funcs = @("Trim-Length","Open-Print-Docs","Get-RandText","Create-Office","Send-Email","Get-Acct","Get-Attachment","Get-Tabs","Save-and-Close","Close-All","Get-MX","Open-IE")
    Invoke-RunspacePool $activities $args $funcs

    # Move Mouse to keep system from hybernating...
    # This will keep mouse moving for 7 hours, then move to save and shutdown
    Add-Type -AssemblyName System.Windows.Forms
    $shutdown_time = 0
    while ($shutdown_time -ne 2520 ) {
        $Pos = [System.Windows.Forms.Cursor]::Position
        $x = ($pos.X % 500) + 1
        $y = ($pos.y % 500) + 1
        [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
        sleep 10
        $shutdown_time += 1
    }

    # Sleep for 8hrs and allow "normal" user activity
    # sleep 28800
 
    # Close all applications to prepare for Logout
    # Close IE
    Close-All

    # Save and close all Office documents
    "doc", "ppt", "xls", "all" | % {
        Save-and-Close ($_)
    }

    # Logout at the end of the day
    # Uncomment the below line to logout instead of system shutdown
    #(Get-WmiObject -Class Win32_OperatingSystem).Win32Shutdown(0)
    
    # Shutdown the computer at the end of the user cycle
    # Comment the below line if you wish to logout instead of shutdown
    Stop-Computer
}

# Function to get the username of the current user and generate the persona ("normal" activity profile)
#   for the user.
function Get-Current-User () {
    $username = $(Get-WMIObject -class Win32_ComputerSystem | select UserName).username.split("\")[1]
    
    #
    #   Import CSV file to determine "normal" activity.  If user not in profile, then randomly
    #   select activity from below
    $persona = Import-Csv C:\scripts\personas.csv -Delimiter ","

    if ($persona | ? { $_.username -like $username }) {
        # Import persona data from CSV
        $persona | % {
            if (($persona | ? { $_.username -like $username }).username -like $username) {
                $open = ($persona | ? { $_.username -like $username }).opendocs
                $create = ($persona | ? { $_.username -like $username }).createdocs
                $pref_doc = ($persona | ? { $_.username -like $username }).preferreddoc
                $print = ($persona | ? { $_.username -like $username }).print
                $check = ($persona | ? { $_.username -like $username }).checkemail
                $sr = ($persona | ? { $_.username -like $username }).sendreceiveemail
                $surf_web = ($persona | ? { $_.username -like $username }).surfweb
                if (($persona | ? { $_.username -like $username }).application -like "cpof") {
                    #$application = "C:\scripts\cpof.bat"
                }
                $web_persona = ($persona | ? { $_.username -like $username }).urls
            }
        }
    } else {
        do {
            $open = (get-random 2)
            $create = (get-random 2)
            $pref_doc = "0"
            $print = (get-random 2)
            $check = (get-random 2)
            $sr = (get-random 2)
            $surf_web = (get-random 2)
            $application = (get-random 2)
            $web_persona = "general"

            # Make sure that Send/Receive and Check email are not selected at the same time
            # to prevent multiple MS Outlook instances from opening
            if ($sr) {
                $check = 0
            }
        } while ($open -eq 0 -and $create -eq 0 -and $print -eq 0 -and $check -eq 0 -and $sr -eq 0 -and $surf_web -eq 0 -and $application -eq 0)
    }
    # Call Get-Actions function to start script operations using random activity values above
    Get-Actions $open $create $pref_doc $print $check $sr $surf_web $application $web_persona
}

# Run all actions. Comment below to use functions as standalone operations in script.
if ( (((Get-WmiObject Win32_ComputerSystem).Domain) -ne "WORKGROUP") -or ((Get-WMIObject -class Win32_ComputerSystem | select UserName).username.split("\")[1]) -ne "bccsadministrator" ) {
    Get-Current-User
}
