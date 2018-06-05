<#

        .SYNOPSIS

        Created by: https://ingogegenwarth.wordpress.com/
        Version:    42 ("What do you get if you multiply six by nine?")
        Changed:    03.12.2017

        .LINK
        http://gsexdev.blogspot.com/
        https://ingogegenwarth.wordpress.com/2015/05/01/troubleshooting-calendar-items/
        https://ingogegenwarth.wordpress.com/2017/11/20/advanced-cal/

        .DESCRIPTION

        The purpose of the script is to for calendar items with a given subject,GlobalObjectID or CleanGlobalObjectID for given mailboxes

        .PARAMETER EmailAddress

        The e-mail address of the mailbox, which will be checked. The script accepts piped objects from Get-Mailbox or Get-Recipient

        .PARAMETER Credentials

        Credentials you want to use. If omitted current user context will be used.

        .PARAMETER Impersonate

        Use this switch, when you want to impersonate.

        .PARAMETER Subject

        The subject, which you want to search for.

        .PARAMETER StartDateLastModified

        If you want to limit the search for items modified after the given date.

        .PARAMETER EndDateLastModified

        If you want to limit the search for items modified before the given date.

        .PARAMETER GlobalObjectID

        Use GlobalObjectID in your search for items.

        .PARAMETER GlobalObjectID

        Use CleanGlobalObjectID in your search for items.

        .PARAMETER CalendarOnly

        By default the script will enumerate all folders under RecoverableItemsRoot and Calendar,Inbox and Sent Items. If you want to limit to folders with type "IPF.Appointment" use this switch.

        .PARAMETER Server

        By default the script tries to retrieve the EWS endpoint via Autodiscover. If you want to run the script against a specific server, just provide the name in this parameter. Not the URL!

        .PARAMETER AllFolders

        All folders within the mailbox will be search for a given criteria (e.g.:Subject,GlobalObjectID or CleanGlobalObjectID)

        .PARAMETER AllItemProps

        The full set of all properties for each item will be loaded

        .PARAMETER SortByDateTimeCreated

        The output will be sort by DateTimeCreated

        .PARAMETER DestinationID
        
        If you ant to have the ItemId converted provide the destination format. Valid formats are "EwsLegacyId","EwsId","EntryId","HexEntryId","StoreId","OwaId" based on https://msdn.microsoft.com/library/microsoft.exchange.webservices.data.idformat(v=exchg.80).aspx

        .PARAMETER TrustAnySSL

        Switch to trust any certificate.

        .PARAMETER DateFormat

        Here you can define the format for the timestamp LastModifiedTime. By default the current culture will be enumerated and the milliseconds appended. To have the same format as the CmdLets use "yyyyMMddThhmmssfff"

        .PARAMETER WebServicesDLL

        Path to the DLL

        .PARAMETER StartDate

        Filter by Start date of a single appointment. Note: Cannot be used to find recurring meetings!

        .PARAMETER EndDate

        Filter by End date of a single appointment. Note: Cannot be used to find recurring meetings!

        .PARAMETER StartDateTimeCreated

        When filter by Datetimecreated, all items created after this date are returned.

        .PARAMETER EndDateTimeCreated

        When filter by Datetimecreated, all items created before this date are returned.

        .PARAMETER UseLocalTime
        When this switch is used, DateTimeCreated and LastModifiedTime is converted to local time of the machine where the script is running on.

        .EXAMPLE

        #run the script against a single mailbox for a given CleanObjectID
        .\Get-CalendarItems.ps1 -EmailAddress trick.duck@adatum.com -CleanGlobalObjectID  040000008200E00074C5B7101A82E00800000000903A2068F779D0010000000000000000100000007900DF424FFB6C498B4B68E21CA9D455

        #same as before, but load all properties for the items
        .\Get-CalendarItems.ps1 -EmailAddress trick.duck@adatum.com -CleanGlobalObjectID  040000008200E00074C5B7101A82E00800000000903A2068F779D0010000000000000000100000007900DF424FFB6C498B4B68E21CA9D455 -AllItemProps

        #search only in folders with type IPM.Appointment
        .\Get-CalendarItems.ps1 -EmailAddress trick.duck@adatum.com -CleanGlobalObjectID  040000008200E00074C5B7101A82E00800000000903A2068F779D0010000000000000000100000007900DF424FFB6C498B4B68E21CA9D455 -AllItemProps -CalendarOnly

        #search in all folders for a specific time range
        .\Get-CalendarItems.ps1 -EmailAddress trick.duck@adatum.com -Subject "Bi-Weekly" -AllFolders -StartDateLastModified ([datetime]::Parse("04.04.2015")) -EndDateLastModified ([datetime]::Parse("05.04.2015"))

        .NOTES
#>

Param (
    [parameter(ValueFromPipelineByPropertyName=$true,Mandatory=$true, Position=0)]
    [Alias('PrimarySmtpAddress')]
    [string[]]$EmailAddress,

    [parameter( Mandatory=$false, Position=1)]
    [System.Management.Automation.PsCredential]
    $Credentials,

    [parameter( Mandatory=$false, Position=2)]
    [switch]$Impersonate,

    [parameter( Mandatory=$false, Position=3)]
    [String]$Subject,

    [parameter( Mandatory=$false, Position=4)]
    [DateTime]$StartDateLastModified,

    [parameter( Mandatory=$false, Position=5)]
    [DateTime]$EndDateLastModified,

    [parameter( Mandatory=$false, Position=6)]
    [string]$CleanGlobalObjectID,

    [parameter( Mandatory=$false, Position=7)]
    [string]$GlobalObjectID,

    [parameter( Mandatory=$false, Position=8)]
    [switch]$CalendarOnly,

    [parameter( Mandatory=$false, Position=9)]
    [string]$Server,

    [parameter( Mandatory=$false, Position=10)]
    [switch]$AllFolders,

    [parameter( Mandatory=$false, Position=11)]
    [switch]$AllItemProps,

    [parameter( Mandatory=$false, Position=12)]
    [switch]$SortByDateTimeCreated,

    [parameter( Mandatory=$false, Position=13)]
    [ValidateSet("EwsLegacyId","EwsId","EntryId","HexEntryId","StoreId","OwaId")]
    [string]$DestinationID,

    [parameter( Mandatory=$false, Position=13)]
    [switch]$TrustAnySSL,

    [parameter( Mandatory=$false, Position=14)]
    [string]$DateFormat='yyyyMMdd HHmmssfff',

    [parameter( Mandatory=$false, Position=15)]
    [ValidateScript({If (Test-Path $_ -PathType leaf){$True} Else {Throw "WebServices DLL could not be found!"}})]
    [string]$WebServicesDLL,

    [parameter( Mandatory=$false, Position=16)]
    [DateTime]$StartDate,

    [parameter( Mandatory=$false, Position=17)]
    [DateTime]$EndDate,

    [parameter( Mandatory=$false, Position=18)]
    [DateTime]$StartDateTimeCreated,

    [parameter( Mandatory=$false, Position=19)]
    [DateTime]$EndDateTimeCreated,

    [parameter( Mandatory=$false, Position=20)]
    [switch]$UseLocalTime

)

Begin {
    function BinToHex {
        param(
            [Parameter(
                    Position=0,
                    Mandatory=$true,
                ValueFromPipeline=$true)
            ]
        [Byte[]]$Bin)
        # assume pipeline input if we don't have an array (surely there must be a better way)
        if ($bin.Length -eq 1){$bin = @($input)}
        $return = -join ($Bin |  ForEach-Object -Process { "{0:X2}" -f $_ })
        Write-Output -InputObject $return
    }

    function HexToBin {
        param(
            [Parameter(
                    Position=0,
                    Mandatory=$true,
                ValueFromPipeline=$true)
            ]
        [string]$s)
        $return = @()
        for ($i = 0; $i -lt $s.Length ; $i += 2)
        {
            $return += [Byte]::Parse($s.Substring($i, 2), [System.Globalization.NumberStyles]::HexNumber)
        }

        Write-Output -InputObject $return
    }

    function ConvertToString($ipInputString){
        $Val1Text = ""
        for ($clInt=0;$clInt -lt $ipInputString.length;$clInt++){
            $Val1Text = $Val1Text + [Convert]::ToString([Convert]::ToChar([Convert]::ToInt32($ipInputString.Substring($clInt,2),16)))
            $clInt++
        }
        return $Val1Text
    }

    function Format-Recurrence {
        param(
            [object]$Item
        )
        Begin {
            [string]$RecurrencePattern= ''
        }
        Process {
            if ($null -ne $Item.Recurrence){
                #get list of properties
                [array]$Properties = $Item.Recurrence | Get-Member -MemberType Property | Select-Object -Property Name
                If ($null -ne $Properties){
                    foreach ($Property in $Properties){
                        $RecurrencePattern+= "$($Property.Name)=$($Item.Recurrence.$($Property.Name))|"
                    }
                }
            }
        }
        End {
            return $RecurrencePattern.Trim('|');
        }
    }

    function ConvertUTCTimeToTimeZone {
        param(
            $UTCTime,
            [string]$TargetZone
        )
        #$Time = [datetime]::Parse($UTCTime)
        $TimeZones = [System.TimeZoneInfo]::GetSystemTimeZones()
        $Id = $TimeZones | Where-Object -FilterScript {$_.DisplayName -eq $TargetZone}
        try {
            $returnValue = [System.TimeZoneInfo]::ConvertTimeFromUtc($UTCTime,$Id)
        }
        catch {
            Write-Verbose -Message "Could not calculate time from extended properties!"
        }
        return $returnValue
    }

    function ConvertFrom-ClientIntent
    {
        [CmdletBinding()]
        [Alias()]
        [OutputType([string])]
        Param
        (
            [Parameter(Mandatory=$true,
                       ValueFromPipelineByPropertyName=$false,
                       Position=0)]
            [int]$ClientIntentValue
        )

        Begin
        {
            [string]$RetunValue =  ''
            $ClientIntentHash = @{
                Manager                        = 1
                Delegate                       = 2
                DeletedWithNoResponse          = 4
                DeletedExceptionWithNoResponse = 8
                RespondedTentative             = 16
                RespondedAccept                = 32
                RespondedDecline               = 64
                ModifiedStartTime              = 128
                ModifiedEndTime                = 256
                ModifiedLocation               = 512
                RespondedExceptionDecline      = 1024
                Canceled                       = 2048
                ExceptionCanceled              = 4096
            }

        }
        Process
        {
            foreach ($Bit in ($ClientIntentHash.GetEnumerator() | Sort-Object -Property Value )){
                if (($ClientIntentValue -band $Bit.Value) -ne 0){
                    $RetunValue += $Bit.Key +'|'
                }
            }
        }
        End
        {
            Write-Verbose -Message "Bit mask:$([Convert]::ToString($ClientIntentValue,2))"
            return ($RetunValue.TrimEnd("|"))
        }
    }

    function ConvertFrom-ChangeHighlight
    {
        [CmdletBinding()]
        [Alias()]
        [OutputType([string])]
        Param
        (
            [Parameter(Mandatory=$true,
                       ValueFromPipelineByPropertyName=$false,
                       Position=0)]
            [int]$ChangeHighlightValue
        )

        Begin
        {
            #$Value= [Convert]::ToString($ChangeHighlightValue,2)
            [string]$RetunValue =  ''
            $ChangeHighlightHash = @{
                START        = 1
                END          = 2
                RECUR        = 4
                LOCATION     = 8
                SUBJECT      = 16
                REQATT       = 32
                OPTATT       = 64
                BODY         = 128
                RESPONSE     = 512
                ALLOWPROPOSE = 1024
            }

        }
        Process
        {
            foreach ($Bit in ($ChangeHighlightHash.GetEnumerator() | Sort-Object -Property Value )){
                if (($ChangeHighlightValue -band $Bit.Value) -ne 0){
                    $RetunValue += $Bit.Key +'|'
                }
            }
        }
        End
        {
            Write-Verbose -Message "Bit mask:$([Convert]::ToString($ChangeHighlightValue,2))"
            return ($RetunValue.TrimEnd("|"))
        }
    }

    #some checks for ambiguous parameters
    If (($Subject -and ($GlobalObjectID -or $CleanGlobalObjectID)) -or ($GlobalObjectID -and $CleanGlobalObjectID)) {
        Write-Warning "Ambiguous parameter combination! Either search for subject or CleanGlobalObjectID/GlobalObjectID!"
        Break
    }

    If (($StartDateLastModified -and ($StartDate -or $EndDate)) -or ($EndDateLastModified -and ($StartDate -or $EndDate))) {
        Write-Warning "Ambiguous parameter combination! Either search for modification or appointment dates!"
        Break
    }

    If (($StartDateLastModified -and ($StartDateTimeCreated -or $StartDateTimeCreated)) -or ($EndDateLastModified -and ($StartDateTimeCreated -or $StartDateTimeCreated))) {
        Write-Warning "Ambiguous parameter combination! Either search for modification or appointment dates!"
        Break
    }

    If (($StartDateTimeCreated -and ($StartDate -or $EndDate)) -or ($EndDateTimeCreated -and ($StartDate -or $EndDate))) {
        Write-Warning "Ambiguous parameter combination! Either search for modification or appointment dates!"
        Break
    }

    If (($StartDateTimeCreated -and ($StartDateLastModified -or $StartDateLastModified)) -or ($EndDateTimeCreated -and ($StartDateLastModified -or $StartDateLastModified))) {
        Write-Warning "Ambiguous parameter combination! Either search for modification or appointment dates!"
        Break
    }

    If ($CalendarOnly -and $AllFolders) {
        Write-Warning "Ambiguous parameter combination! Either search in CalendarOnly or in AllFolders!"
        Break
    }

    $objcol = @()
    #get culture for datetime formatting
    $culture=get-culture
    $TimeZones = [System.TimeZoneInfo]

    [string]$AlternateIDName= $Null
    Switch -wildcard ($DestinationID)
    {
        "EwsL*" {$AlternateIDName = "EwsLegacyId"}
        "EwsI*" {$AlternateIDName = "EwsId"}
        "Entr*" {$AlternateIDName = "EntryId"}
        "HexE*" {$AlternateIDName = "HexEntryId"}
        "Stor*" {$AlternateIDName = "StoreId"}
        "OwaI*" {$AlternateIDName = "OwaId"}
    }
}

Process {
    try {
        #$MailboxName = $EmailAddress
        ForEach ($MailboxName in $EmailAddress){
            [string]$RootFolder="MsgFolderRoot"
            if ($WebServicesDLL){
                try {
                    $EWSDLL = $WebServicesDLL
                    Import-Module -Name $EWSDLL
                }
                catch {
                    $Error[0].Exception
                    exit
                }
            }
            else {
                ## Load Managed API dll
                ###CHECK FOR EWS MANAGED API, IF PRESENT IMPORT THE HIGHEST VERSION EWS DLL, ELSE EXIT
                $EWSDLL = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Exchange\Web Services'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Install Directory') + "Microsoft.Exchange.WebServices.dll")
                if (Test-Path -Path $EWSDLL)
                {
                    Import-Module -Name $EWSDLL
                }
                else
                {
                    "$(get-date -format yyyyMMddHHmmss):"
                    "This script requires the EWS Managed API 1.2 or later."
                    "Please download and install the current version of the EWS Managed API from"
                    "http://go.microsoft.com/fwlink/?LinkId=255472"
                    ""
                    "Exiting Script."
                    exit
                }
            }

            ## Set Exchange Version
            $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
            ## Create Exchange Service Object
            $Service = New-Object -TypeName Microsoft.Exchange.WebServices.Data.ExchangeService -ArgumentList ($ExchangeVersion)
            #$service.PreAuthenticate = $true
            #set DateTimePrecision to get milliseconds
            $Service.DateTimePrecision=[Microsoft.Exchange.WebServices.Data.DateTimePrecision]::Milliseconds
            ## Set Credentials to use two options are available Option1 to use explict credentials or Option 2 use the Default (logged On) credentials
            If ($Credentials){
                #Credentials Option 1 using UPN for the windows Account
                #$psCred = Get-Credential
                $psCred = $Credentials
                $creds = New-Object -TypeName System.Net.NetworkCredential -ArgumentList ($psCred.UserName.ToString(),$psCred.GetNetworkCredential().password.ToString())
                $service.Credentials = $creds
                #$service.TraceEnabled = $true
            }
            Else {
                #Credentials Option 2
                $service.UseDefaultCredentials = $true
            }

            If ($TrustAnySSL){
                ## Choose to ignore any SSL Warning issues caused by Self Signed Certificates
                ## Code From http://poshcode.org/624
                ## Create a compilation environment
                $Provider=New-Object -TypeName Microsoft.CSharp.CSharpCodeProvider
                $Compiler=$Provider.CreateCompiler()
                $Params=New-Object -TypeName System.CodeDom.Compiler.CompilerParameters
                $Params.GenerateExecutable=$False
                $Params.GenerateInMemory=$True
                $Params.IncludeDebugInformation=$False
                $Params.ReferencedAssemblies.Add("System.DLL") | Out-Null
                $TASource=@'
namespace Local.ToolkitExtensions.Net.CertificatePolicy{
public class TrustAll : System.Net.ICertificatePolicy {
public TrustAll(){
}
public bool CheckValidationResult(System.Net.ServicePoint sp,
System.Security.Cryptography.X509Certificates.X509Certificate cert,
System.Net.WebRequest req, int problem){
return true;
}
}
}
'@
                $TAResults=$Provider.CompileAssemblyFromSource($Params,$TASource)
                $TAAssembly=$TAResults.CompiledAssembly
                ## We now create an instance of the TrustAll and attach it to the ServicePointManager
                $TrustAll=$TAAssembly.CreateInstance("Local.ToolkitExtensions.Net.CertificatePolicy.TrustAll")
                [System.Net.ServicePointManager]::CertificatePolicy=$TrustAll
                ## end code from http://poshcode.org/624
            }

            ## Set the URL of the CAS (Client Access Server) to use two options are availbe to use Autodiscover to find the CAS URL or Hardcode the CAS to use
            If ($Server){
                #CAS URL Option 2 Hardcoded
                $uri=[system.URI] "https://$server/ews/exchange.asmx"
                $service.Url = $uri
            }
            Else {
                #CAS URL Option 1 Autodiscover
                $service.AutodiscoverUrl($MailboxName,{$true})
                #"Using CAS Server : " + $Service.url
            }

            ## Optional section for Exchange Impersonation
            If ($Impersonate){
                $Service.ImpersonatedUserId = new-object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress, $MailboxName)
            }

            If ($CalendarOnly){
                #Search only for Calendarfolder
                $RootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$RootFolder,$MailboxName)
                #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling
                $FolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
                #Deep Transval will ensure all folders in the search path are returned
                $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep;
                #$FolderPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $FolderPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
                #Add Properties to the Property Set
                $FolderPropertySet.Add($PR_Folder_Path)
                $FolderPropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
                $FolderView.PropertySet = $FolderPropertySet
                $FolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Appointment")
                $FolderResult = $null
                $FolderResult = $Service.FindFolders($RootFolderId,$FolderSearchFilter,$FolderView)
                $AllFolderResult = $FolderResult
            }
            ElseIf ($AllFolders){
                #Search from MsgRoot
                $RootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$MailboxName)
                #$RootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Root,$MailboxName)
                #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling
                $FolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
                #Deep Transval will ensure all folders in the search path are returned
                $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                #$FolderPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $FolderPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
                #Add Properties to the Property Set
                $FolderPropertySet.Add($PR_Folder_Path)
                $FolderPropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
                $FolderView.PropertySet = $FolderPropertySet
                $PR_FOLDER_TYPE = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(13825,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
                #create search filter for folders
                $FolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PR_FOLDER_TYPE,"1")
                $FolderResult = $null
                $FolderResult = $Service.FindFolders($RootFolderId,$FolderSearchFilter,$FolderView)
                $AllFolderResult = $FolderResult
                #Bind to RecoverableItemsRoot
                $RecoverableItemsRoot = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsRoot,$MailboxName)
                $FolderResult = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$RecoverableItemsRoot,$FolderPropertySet)
                $AllFolderResult += $FolderResult   
                #find subfolders of RecoverableItemsRoot
                $RootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsRoot,$MailboxName)
                $FolderResult = $Service.FindFolders($RootFolderId,$FolderView)
                $AllFolderResult += $FolderResult
            }
            Else {
                #Define the FolderView used for Export should not be any larger then 1000 folders due to throttling
                $FolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)
                #Deep Transval will ensure all folders in the search path are returned
                $FolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                #$FolderPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                $FolderPropertySet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                $PR_Folder_Path = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(26293, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
                #Add Properties to the Property Set
                $FolderPropertySet.Add($PR_Folder_Path)
                $FolderPropertySet.Add([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName)
                $FolderView.PropertySet = $FolderPropertySet
                $FolderResult = $null
                $AllFolderResult = $null
                #Search only for Calendarfolder
                $RootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::$RootFolder,$MailboxName)
                $FolderSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::FolderClass, "IPF.Appointment")
                $FolderResult = $Service.FindFolders($RootFolderId,$FolderSearchFilter,$FolderView)
                $AllFolderResult += $FolderResult
                #Bind to Inbox
                $Inbox = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$MailboxName)
                $FolderResult = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$Inbox,$FolderPropertySet)
                $AllFolderResult += $FolderResult
                #Bind to SentItems
                $SentItems = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::SentItems,$MailboxName)
                $FolderResult = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$SentItems,$FolderPropertySet)
                $AllFolderResult += $FolderResult
                #Bind to RecoverableItemsRoot
                $RecoverableItemsRoot = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsRoot,$MailboxName)
                $FolderResult = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($Service,$RecoverableItemsRoot,$FolderPropertySet)
                $AllFolderResult += $FolderResult
                #find subfolders of RecoverableItemsRoot
                $RootFolderId = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::RecoverableItemsRoot,$MailboxName)
                $FolderResult = $Service.FindFolders($RootFolderId,$FolderView)
                $AllFolderResult += $FolderResult
            }

            #define propertyset for items
            $Client                        = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::CalendarAssistant,0x000B,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
            $Action                        = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::CalendarAssistant,0x0006,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
            $PidLidGlobalObjectId          = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting,0x0003,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
            $PidLidCleanGlobalObjectId     = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting,0x0023,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary)
            $PidLidAppointmentMessageClass = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting,0x0024,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
            $PidLidClientIntent            = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::CalendarAssistant,0x0015,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
            $PidLidIsException             = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting,0x000A,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Boolean)
            $PidLidAppointmentStartWhole   = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment,0x820D,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
            $PidLidAppointmentEndWhole     = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment,0x820E,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::SystemTime)
            $PidLidTimeZoneDescription     = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment,0x8234,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
            $PidLidChangeHighlight         = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Appointment,0x8204,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Integer)
            $PR_Creator_Name               = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3FF8,[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
            $UCMeetingSettingStr           = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings,"UCMeetingSetting",[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
            $OnlineMeetingConfLink         = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::PublicStrings,"OnlineMeetingConfLink",[Microsoft.Exchange.WebServices.Data.MapiPropertyType]::String)
            $ItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)

            If ($AllItemProps){
                $ItemPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
            }
            Else {
                $ItemPropset= new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
            }

            #add properties to be loaded
            $ItemPropset.Add($Client)
            $ItemPropset.Add($Action)
            $ItemPropset.Add($PidLidGlobalObjectId)
            $ItemPropset.Add($PidLidCleanGlobalObjectId)
            $ItemPropset.Add($PidLidClientIntent)
            $ItemPropset.Add($PR_Creator_Name)
            $ItemPropset.Add($PidLidChangeHighlight)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeSent)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::ItemClass)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedName)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::OptionalAttendees)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Organizer)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::RequiredAttendees)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::ICalUid)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::StartTimeZone)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::EndTimeZone)
            $ItemPropset.Add([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsRecurring)
            $ItemPropset.Add($PidLidAppointmentStartWhole)
            $ItemPropset.Add($PidLidAppointmentEndWhole)
            $ItemPropset.Add($PidLidIsException)
            $ItemPropset.Add($PidLidTimeZoneDescription)
            $ItemPropset.Add($UCMeetingSettingStr)
            #default searchfiltercollection
            $SearchFilterCollection = new-object  Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::And)
            $SearchFilterCollectionItemClass = new-object  Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::Or)
            $PR_MESSAGE_CLASS_Filter_Appointment = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Appointment")
            $PR_MESSAGE_CLASS_Filter_Schedule_Canceled = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Schedule.Meeting.Canceled")
            $PR_MESSAGE_CLASS_Filter_Schedule_Request = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Schedule.Meeting.Request")
            $PR_MESSAGE_CLASS_Filter_Schedule_Resp_Neg = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Schedule.Meeting.Resp.Neg")
            $PR_MESSAGE_CLASS_Filter_Schedule_Resp_Pos = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Schedule.Meeting.Resp.Pos")
            $PR_MESSAGE_CLASS_Filter_Schedule_Resp_Tent = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::ItemClass, "IPM.Schedule.Meeting.Resp.Tent")
            $PidLidAppointmentMessageClass_Filter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PidLidAppointmentMessageClass, "IPM.Appointment")
            $SearchFilterCollectionItemClass.Add($PR_MESSAGE_CLASS_Filter_Appointment)
            $SearchFilterCollectionItemClass.Add($PR_MESSAGE_CLASS_Filter_Schedule_Canceled)
            $SearchFilterCollectionItemClass.Add($PR_MESSAGE_CLASS_Filter_Schedule_Request)
            $SearchFilterCollectionItemClass.Add($PR_MESSAGE_CLASS_Filter_Schedule_Resp_Neg)
            $SearchFilterCollectionItemClass.Add($PR_MESSAGE_CLASS_Filter_Schedule_Resp_Pos)
            $SearchFilterCollectionItemClass.Add($PR_MESSAGE_CLASS_Filter_Schedule_Resp_Tent)
            $SearchFilterCollectionItemClass.Add($PidLidAppointmentMessageClass_Filter)
            $SearchFilterCollection.Add($SearchFilterCollectionItemClass)

            #search by subject
            If ($Subject){
                $SubjectFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+ContainsSubstring([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::Subject, "$Subject")
                $SearchFilterCollection.Add($SubjectFilter)
            }

            #search by CleanGlobalObjectID
            If($CleanGlobalObjectID){
                $CleanGlobalObjectID_Bin = HexToBin $CleanGlobalObjectID
                $CleanGOIDFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PidLidCleanGlobalObjectId,[System.Convert]::ToBase64String($CleanGlobalObjectID_Bin))
                $SearchFilterCollection.Add($CleanGOIDFilter)
            }

            #search by GlobalObjectID
            If($GlobalObjectID){
                $GlobalObjectID_Bin = HexToBin $GlobalObjectID
                $GOIDFilter = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($PidLidGlobalObjectId, [System.Convert]::ToBase64String($GlobalObjectID_Bin))
                $SearchFilterCollection.Add($GOIDFilter)
            }

            #search by date range
            #search by DateLastModified
            If ($StartDateLastModified -and $EndDateLastModified){
                $SearchFilterCollectionDateTime = new-object  Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::AND)
                $StartLastModified = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, $StartDateLastModified)
                $EndLastModified = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, $EndDateLastModified)
                $SearchFilterCollectionDateTime.Add($StartLastModified)
                $SearchFilterCollectionDateTime.Add($EndLastModified)
                $SearchFilterCollection.Add($SearchFilterCollectionDateTime)
            }
            Else {
                If ($StartDateLastModified){
                    $StartLastModified = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, $StartDateLastModified)
                    $SearchFilterCollection.Add($StartLastModified)
                }

                If ($EndDateLastModified){
                    $EndLastModified = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::LastModifiedTime, $EndDateLastModified)
                    $SearchFilterCollection.Add($EndLastModified)
                }
            }
            #search by Start/End date of appointment
            If ($StartDate -and $EndDate){
                $SearchFilterCollectionDateTime = new-object  Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::AND)
                $StartItem = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, $StartDate)
                $EndItem = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, $EndDate)
                $IsRecurring = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsRecurring, $true)
                $NoRecurring = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($IsRecurring)
                $SearchFilterCollectionDateTime.Add($StartItem)
                $SearchFilterCollectionDateTime.Add($EndItem)
                $SearchFilterCollection.Add($SearchFilterCollectionDateTime)
                $SearchFilterCollection.Add($NoRecurring)
            }
            Else {
                If ($StartDate){
                    $StartItem = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::Start, $StartDate)
                    $IsRecurring = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsRecurring, $true)
                    $NoRecurring = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($IsRecurring)
                    $SearchFilterCollection.Add($StartItem)
                    $SearchFilterCollection.Add($NoRecurring)
                }

                If ($EndDate){
                    $EndItem = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::End, $EndDate)
                    $IsRecurring = New-Object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.AppointmentSchema]::IsRecurring, $true)
                    $NoRecurring = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+Not($IsRecurring)
                    $SearchFilterCollection.Add($EndItem)
                    $SearchFilterCollection.Add($NoRecurring)
                }
            }
            #search by DateTimeCreated
            If($StartDateTimeCreated -and $EndDateTimeCreated){
                $SearchFilterCollectionDateTime = new-object  Microsoft.Exchange.WebServices.Data.SearchFilter+SearchFilterCollection([Microsoft.Exchange.WebServices.Data.LogicalOperator]::AND)
                $StartDateTimeCreatedItem = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $StartDateTimeCreated)
                $EndDateTimeCreatedItem = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $EndDateTimeCreated)
                $SearchFilterCollectionDateTime.Add($StartDateTimeCreatedItem)
                $SearchFilterCollectionDateTime.Add($EndDateTimeCreatedItem)
                $SearchFilterCollection.Add($SearchFilterCollectionDateTime)
            }
            Else{
                If($StartDateTimeCreated){
                    $StartDateTimeCreatedItem = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsGreaterThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $StartDateTimeCreated)
                    $SearchFilterCollection.Add($StartDateTimeCreatedItem)
                }
                If($EndDateTimeCreated){
                    $EndDateTimeCreatedItem = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsLessThanOrEqualTo([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeCreated, $EndDateTimeCreated)
                    $SearchFilterCollection.Add($EndDateTimeCreatedItem)
                }
            }

            #loop through folders
            do {
                #exclude Audit folder
                $AllFolderResult = $AllFolderResult | Where-Object -FilterScript {$_.Displayname -ne 'Audits'}
                [int]$i='1'
                ForEach ($Folder in $AllFolderResult){
                    #Write-Output "Working on  $($Folder.DisplayName)"
                    #show progress
                    Write-Progress `
                    -id 1 `
                    -Activity "Processing mailbox - $($MailboxName) with $($AllFolderResult.count) folders" `
                    -PercentComplete ( $i / $AllFolderResult.count * 100) `
                    -Status "Remaining folders: $($AllFolderResult.count - $i) processing folder: $($Folder.DisplayName)"
                    $foldpathval = $null
                    $fpath = $null
                    #Try to get the FolderPath Value and then covert it to a usable String
                    If ($Folder.TryGetProperty($PR_Folder_Path,[ref] $foldpathval)){
                        $binarry = [Text.Encoding]::UTF8.GetBytes($foldpathval)
                        $hexArr = $binarry | ForEach-Object { $_.ToString("X2") }
                        $hexString = $hexArr -join ''
                        $hexString = $hexString.Replace("EFBFBE", "5C")
                        $fpath = ConvertToString($hexString)
                    }
                    $FindItems = $null
                    [int]$p= '0'
                    do {
                        $FindItems = $Service.FindItems($Folder.Id,$SearchFilterCollection,$ItemView)
                        If ($FindItems.Items.Count -ge '1'){
                            [Void]$service.LoadPropertiesForItems($FindItems,$ItemPropset)
                            [int]$y='1'
                            [int]$z= '0'
                            If ($DestinationID){
                                #create list
                                Write-Verbose    "Create Generic.List for AlternateIDs"
                                $GenList = ('System.Collections.Generic.List'+'`'+'1') -as 'Type'
                                $GenList = $GenList.MakeGenericType('Microsoft.Exchange.WebServices.Data.AlternateIdBase' -as 'Type')
                                $AltIdBases = [Activator]::CreateInstance($GenList)
                                $FindItems.Items | ForEach{$AltIdBases.Add($(New-Object -TypeName Microsoft.Exchange.WebServices.Data.AlternateId -ArgumentList ('EWSId',$_.ID,$MailboxName)))}
                                $Converted= $Service.ConvertIds($AltIdBases,"$DestinationID")
                                If ($FindItems.Items.Count -eq 1){
                                    If ($Converted.Result -ne 'Success'){
                                        Write-Warning "1 item from folder $($Folder.DisplayName) couldn't be converted successful!"
                                    }
                                }
                                Else {
                                    If (($Converted.Result -ne 'Success').Count -gt 0){
                                        Write-Warning "$(($Converted.Result -ne 'Success').Count) items from folder $($Folder.DisplayName) couldn't be converted successful!"
                                    }
                                }
                                
                            }
                            ForEach ($Item in $FindItems.Items){
                                Write-Progress `
                                -id 2 `
                                -ParentId 1 `
                                -Activity "Processing item - $($Item.Subject)" `
                                -PercentComplete ( $p / $FindItems.TotalCount * 100)`
                                -Status "Total items: $($FindItems.TotalCount) remaining items: $($FindItems.TotalCount - $p) processing folder: $($Folder.DisplayName)"
                                $data = New-Object -TypeName PSObject
                                If ($DateFormat){
                                    If ($SortByDateTimeCreated){
                                        If ($UseLocalTime){
                                            $data | add-member -type NoteProperty -Name DateTimeCreatedLocal -Value $( Get-Date $([System.TimeZone]::CurrentTimeZone.ToLocalTime($Item.DateTimeCreated.ToUniversalTime())) -Format $($DateFormat))
                                        }
                                        Else {
                                            $data | add-member -type NoteProperty -Name DateTimeCreatedUTC -Value $( Get-Date $Item.DateTimeCreated.ToUniversalTime() -Format $($DateFormat))
                                        }
                                    }
                                    If ($UseLocalTime){
                                        $data | add-member -type NoteProperty -Name LastModifiedTimeLocal -Value $( Get-Date $([System.TimeZone]::CurrentTimeZone.ToLocalTime($Item.LastModifiedTime.ToUniversalTime())) -Format $($DateFormat))
                                    }
                                    Else{
                                        $data | add-member -type NoteProperty -Name LastModifiedTimeUTC -Value $( Get-Date $Item.LastModifiedTime.ToUniversalTime() -Format $($DateFormat))
                                    }
                                }
                                Else {
                                    If ($SortByDateTimeCreated){
                                        If ($UseLocalTime){
                                            $data | add-member -type NoteProperty -Name DateTimeCreatedLocal -Value [System.TimeZone]::CurrentTimeZone.ToLocalTime($( Get-Date $Item.DateTimeCreated.ToUniversalTime() -Format $($culture.DateTimeFormat.FullDateTimePattern.ToString().Replace(':ss',':ss.fff'))))
                                        }
                                        Else {
                                            $data | add-member -type NoteProperty -Name DateTimeCreatedUTC -Value $( Get-Date $Item.DateTimeCreated.ToUniversalTime() -Format $($culture.DateTimeFormat.FullDateTimePattern.ToString().Replace(':ss',':ss.fff')))
                                        }
                                    }
                                    If ($UseLocalTime){
                                        $data | add-member -type NoteProperty -Name LastModifiedTimeLocal -Value [System.TimeZone]::CurrentTimeZone.ToLocalTime($( Get-Date $Item.LastModifiedTime.ToUniversalTime() -Format $($culture.DateTimeFormat.FullDateTimePattern.ToString().Replace(':ss',':ss.fff'))))
                                    }
                                    Else{
                                        $data | add-member -type NoteProperty -Name LastModifiedTimeUTC -Value $( Get-Date $Item.LastModifiedTime.ToUniversalTime() -Format $($culture.DateTimeFormat.FullDateTimePattern.ToString().Replace(':ss',':ss.fff')))
                                    }
                                }
                                $data | add-member -type NoteProperty -Name Mailbox -Value $MailboxName
                                If ($AllItemProps){
                                    $data | add-member -type NoteProperty -Name Item -Value $Item
                                    $data | add-member -type NoteProperty -Name LastModifiedName -Value $Item.LastModifiedName
                                    $data | add-member -type NoteProperty -Name Creator -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.Tag -eq '16376'}){($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.Tag -eq '16376'}).Value})
                                    $data | add-member -type NoteProperty -Name Subject -Value $Item.Subject
                                    $data | add-member -type NoteProperty -Name FolderPath -Value $fpath
                                    $data | add-member -type NoteProperty -Name Client -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '11'}){($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '11'}).Value})
                                    $data | add-member -type NoteProperty -Name Action -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '6'}){($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '6'}).Value})
                                    $data | add-member -type NoteProperty -Name ItemClass -Value $Item.ItemClass

                                    If(($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}) -and ($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'})){
                                        Write-Verbose "Both properties exist. Will caculate start time..."
                                        #$data | add-member -type NoteProperty -Name Start -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}){ConvertUTCTimeToTimeZone -UTCTime $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}).Value -TargetZone $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}).Value} )
                                        If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}){
                                            $UTCTime = $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}).Value
                                            $TargetZone = $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}).Value
                                            $StartValue = ConvertUTCTimeToTimeZone -UTCTime $UTCTime -TargetZone $TargetZone
                                        }
                                        If ($null -eq $StartValue){
                                            $StartValue = $Item.Start
                                        }
                                        $data | add-member -type NoteProperty -Name Start -Value $StartValue
                                    }
                                    Else{
                                        If ($null -ne $Item.Start){
                                            $data | add-member -type NoteProperty -Name Start -Value $Item.Start
                                        }
                                        Else {
                                            $data | add-member -type NoteProperty -Name Start -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}){"UTC:"+$($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}).Value })
                                        }
                                    }

                                    If (($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}) -and ($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'})){
                                        Write-Verbose "Both properties exist. Will caculate end time..."
                                        #$data | add-member -type NoteProperty -Name End -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}){ConvertUTCTimeToTimeZone -UTCTime $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}).Value -TargetZone $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}).Value} )
                                        If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}){
                                            $UTCTime = $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}).Value
                                            $TargetZone = $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}).Value
                                            $EndValue = ConvertUTCTimeToTimeZone -UTCTime $UTCTime -TargetZone $TargetZone
                                        }
                                        If ($null -eq $EndValue){
                                            $EndValue = $Item.End
                                        }
                                        $data | add-member -type NoteProperty -Name End -Value $EndValue
                                    }
                                    Else {
                                        If ($null -ne $Item.End){
                                            $data | add-member -type NoteProperty -Name End -Value $Item.End
                                        }
                                        Else {
                                            $data | add-member -type NoteProperty -Name End -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}){"UTC:"+$($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}).Value })
                                        }
                                    }

                                    $data | add-member -type NoteProperty -Name IsRecurring -Value $Item.IsRecurring
                                    $data | add-member -type NoteProperty -Name IsException -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '10'}){($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '10'}).Value})
                                    $data | add-member -type NoteProperty -Name Recurrence -Value $(Format-Recurrence -Item $Item)
                                    $data | add-member -type NoteProperty -Name PidLidClientIntent -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '21'}){ConvertFrom-ClientIntent -ClientIntentValue ($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '21'}).Value})
                                    $data | add-member -type NoteProperty -Name PidLidChangeHighlight -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33284'}){ConvertFrom-ChangeHighlight -ChangeHighlightValue ($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33284'}).Value})
                                    $data | add-member -type NoteProperty -Name CleanGlobalObjectID -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '35'}){[System.BitConverter]::ToString(($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '35'}).Value) -Replace '-',''})
                                    $data | add-member -type NoteProperty -Name GlobalObjectID -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '3'}){[System.BitConverter]::ToString(($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '3'}).Value) -Replace '-',''})
                                    If ($DestinationID){
                                        $data | add-member -type NoteProperty -Name $AlternateIDName -Value $($Converted[$z].ConvertedId.UniqueId)
                                        $z++
                                    }
                                    $data | add-member -type NoteProperty -Name ModifiedOccurences -Value $(If($Item.ModifiedOccurrences){[string]::Join(",",$($Item.ModifiedOccurrences | Select-Object -Property Start,End )) -replace "@{","" -replace "}","" -replace ";","" -replace ",",";"}) 
                                    $data | add-member -type NoteProperty -Name DeletedOccurrences -Value $(If($Item.DeletedOccurrences){[string]::Join(";",$($Item.DeletedOccurrences| Select-Object -Property OriginalStart )) -replace "@{","" -replace "}",""})
                                    $data | add-member -type NoteProperty -Name TimeZone -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}){($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}).Value})
                                    $data | add-member -type NoteProperty -Name UCMeetingSetting -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.Name -eq 'UcMeetingSetting'}){($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.Name -eq 'UcMeetingSetting'}).Value})
                                    $objcol += $data
                                }
                                Else {
                                    $data | add-member -type NoteProperty -Name LastModifiedName -Value $Item.LastModifiedName
                                    $data | add-member -type NoteProperty -Name Subject -Value $Item.Subject
                                    $data | add-member -type NoteProperty -Name Client -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '11'}){($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '11'}).Value})
                                    $data | add-member -type NoteProperty -Name FolderPath -Value $fpath
                                    $data | add-member -type NoteProperty -Name Action -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '6'}){($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '6'}).Value})
                                    If(($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}) -and ($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'})){
                                        Write-Verbose "Both properties exist. Will caculate start time..."
                                        #$data | add-member -type NoteProperty -Name Start -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}){ConvertUTCTimeToTimeZone -UTCTime $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}).Value -TargetZone $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}).Value} )
                                        If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}){
                                            $UTCTime = $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}).Value
                                            $TargetZone = $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}).Value
                                            $StartValue = ConvertUTCTimeToTimeZone -UTCTime $UTCTime -TargetZone $TargetZone
                                        }
                                        If ($null -eq $StartValue){
                                            $StartValue = $Item.Start
                                        }
                                        $data | add-member -type NoteProperty -Name Start -Value $StartValue
                                    }
                                    Else{
                                        If ($null -ne $Item.Start){
                                            $data | add-member -type NoteProperty -Name Start -Value $Item.Start
                                        }
                                        Else {
                                            $data | add-member -type NoteProperty -Name Start -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}){"UTC:"+$($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33293'}).Value })
                                        }
                                    }

                                    If (($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}) -and ($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'})){
                                        Write-Verbose "Both properties exist. Will caculate end time..."
                                        #$data | add-member -type NoteProperty -Name End -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}){ConvertUTCTimeToTimeZone -UTCTime $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}).Value -TargetZone $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}).Value} )
                                        If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}){
                                            $UTCTime = $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}).Value
                                            $TargetZone = $($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33332'}).Value
                                            $EndValue = ConvertUTCTimeToTimeZone -UTCTime $UTCTime -TargetZone $TargetZone
                                        }
                                        If ($null -eq $EndValue){
                                            $EndValue = $Item.End
                                        }
                                        $data | add-member -type NoteProperty -Name End -Value $EndValue
                                    }
                                    Else {
                                        If ($null -ne $Item.End){
                                            $data | add-member -type NoteProperty -Name End -Value $Item.End
                                        }
                                        Else {
                                            $data | add-member -type NoteProperty -Name End -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}){"UTC:"+$($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '33294'}).Value })
                                        }
                                    }
                                    $data | add-member -type NoteProperty -Name ItemClass -Value $Item.ItemClass
                                    $data | add-member -type NoteProperty -Name CleanGlobalObjectID -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '35'}){[System.BitConverter]::ToString(($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '35'}).Value) -Replace '-',''})
                                    $data | add-member -type NoteProperty -Name Organizer -Value $Item.Organizer
                                    $data | add-member -type NoteProperty -Name RequiredAttendees -Value $(If($Item.RequiredAttendees.Count -gt '0'){[string]::Join(";",$($Item.RequiredAttendees| ForEach-Object -Process {$_.Address}))})
                                    $data | add-member -type NoteProperty -Name OptionalAttendees -Value $(If($Item.OptionalAttendees.Count -gt '0'){[string]::Join(";",$($Item.OptionalAttendees| ForEach-Object -Process {$_.Address}))})
                                    $data | add-member -type NoteProperty -Name IsRecurring -Value $Item.IsRecurring
                                    $data | add-member -type NoteProperty -Name DateTimeCreated -Value $Item.DateTimeCreated
                                    $data | add-member -type NoteProperty -Name DateTimeReceived -Value $Item.DateTimeReceived
                                    $data | add-member -type NoteProperty -Name DateTimeSent -Value $Item.DateTimeSent
                                    $data | add-member -type NoteProperty -Name FolderID -Value $Folder.Id.UniqueId
                                    $data | add-member -type NoteProperty -Name GlobalObjectID -Value $(If($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '3'}){[System.BitConverter]::ToString(($Item.ExtendedProperties | Where-Object -FilterScript {$_.PropertyDefinition.id -eq '3'}).Value) -Replace '-',''})
                                    If ($DestinationID){
                                        $data | add-member -type NoteProperty -Name $AlternateIDName -Value $($Converted[$z].ConvertedId.UniqueId)
                                        $z++
                                    }
                                    $objcol += $data
                                }
                                $y++
                                $p++
                            }
                        }
                        $ItemView.Offset = $FindItems.NextPageOffset
                    }while($FindItems.MoreAvailable)
                    Write-Progress -id 2 -ParentId 1 -Activity "Processing item - $($Item.Subject)" -Status "Ready" -Completed
                    $i++
                }
                $FolderView.Offset += $FolderResult.Folders.Count
                #end folder loop
            }while($FolderResult.MoreAvailable -eq $true)
            Write-Progress -Activity "Processing item - $($Item.Subject)" -Status "Ready" -Completed
            Write-Progress -Activity "Processing mailbox - $($MailboxName) with $($FolderResult.Folders.count) folders" -Status "Ready" -Completed
        }
    }
    catch{
        #create object
        $returnValue = New-Object -TypeName PSObject
        #get all properties from last error
        $ErrorProperties =$Error[0] | Get-Member -MemberType Property
        #add existing properties to object
        foreach ($Property in $ErrorProperties){
            if ($Property.Name -eq 'InvocationInfo'){
                $returnValue | Add-Member -Type NoteProperty -Name 'InvocationInfo' -Value $($Error[0].InvocationInfo.PositionMessage)
            }
            else {
                $returnValue | Add-Member -Type NoteProperty -Name $($Property.Name) -Value $($Error[0].$($Property.Name))
            }
        }
        #return object
        $returnValue
    }
}

End {
    If ($SortByDateTimeCreated){
        If ($UseLocalTime){
            $objcol | Sort-Object -Property DateTimeCreatedLocal
        }
        Else{
            $objcol | Sort-Object -Property DateTimeCreatedUTC
        }
    }
    Else {
        $objcol | Sort-Object -Property LastModifiedTimeUTC
    }
}