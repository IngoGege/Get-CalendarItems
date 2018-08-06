# Get-CalendarItems

Retrieves Calendar items of specified mailboxes for troubleshooting.

### Prerequisites

# Permissions: Either FullAccess or ApplicationImpersonation

# [Exchange Web Services Managed API](http://go.microsoft.com/fwlink/?LinkId=255472)

## Examples

#run the script against a single mailbox for a given CleanObjectID
```
.\Get-CalendarItems.ps1 -EmailAddress trick.duck@adatum.com -CleanGlobalObjectID  040000008200E00074C5B7101A82E00800000000903A2068F779D0010000000000000000100000007900DF424FFB6C498B4B68E21CA9D455
```

#same as before, but load all properties for the items
```
.\Get-CalendarItems.ps1 -EmailAddress trick.duck@adatum.com -CleanGlobalObjectID  040000008200E00074C5B7101A82E00800000000903A2068F779D0010000000000000000100000007900DF424FFB6C498B4B68E21CA9D455 -AllItemPropsThe script has multiple parameters. Only a few are shown in this example:
```

#search only in folders with type IPM.Appointment```
```
.\Get-CalendarItems.ps1 -EmailAddress trick.duck@adatum.com -CleanGlobalObjectID  040000008200E00074C5B7101A82E00800000000903A2068F779D0010000000000000000100000007900DF424FFB6C498B4B68E21CA9D455 -AllItemProps -CalendarOnly.\Get-CalendarItems.ps1 -EmailAddress manager@contoso.com,assistant@contoso.com -Impersonate -Credentials (Get-Credential) -Subject '1:1 with Rob' 
```
#search in all folders for a specific time range
```
.\Get-CalendarItems.ps1 -EmailAddress trick.duck@adatum.com -Subject "Bi-Weekly" -AllFolders -StartDateLastModified ([datetime]::Parse("04.04.2015")) -EndDateLastModified ([datetime]::Parse("05.04.2015"))### About
```

## Required Parameters

### -EmailAddress

The e-mail address of the mailbox, which will be checked. The script accepts piped objects from Get-Mailbox or Get-Recipient

## Optional Parameters

### -Credentials

Credentials you want to use. If omitted current user context will be used.For more information on this script, as well as usage and examples, see the related blog article on [The Clueless Guy](https://ingogegenwarth.wordpress.com/2015/05/01/troubleshooting-calendar-items/#Script).

### -Impersonate

Use this switch, when you want to impersonate.

### -Subject

The subject, which you want to search for.

### -StartDateLastModified

If you want to limit the search for items modified after the given date.

### -EndDateLastModified

If you want to limit the search for items modified before the given date.

### -GlobalObjectID

Use GlobalObjectID in your search for items.

### -GlobalObjectID

Use CleanGlobalObjectID in your search for items.

### -CalendarOnly

By default the script will enumerate all folders under RecoverableItemsRoot and Calendar,Inbox and Sent Items. If you want to limit to folders with type "IPF.Appointment" use this switch.

### -Server

By default the script tries to retrieve the EWS endpoint via Autodiscover. If you want to run the script against a specific server, just provide the name in this parameter. Not the URL!

### -AllFolders

All folders within the mailbox will be search for a given criteria (e.g.:Subject,GlobalObjectID or CleanGlobalObjectID)

### -AllItemProps

The full set of all properties for each item will be loaded

### -SortByDateTimeCreated

The output will be sort by DateTimeCreated

### -DestinationID

If you ant to have the ItemId converted provide the destination format. Valid formats are "EwsLegacyId","EwsId","EntryId","HexEntryId","StoreId","OwaId" based on https://msdn.microsoft.com/library/microsoft.exchange.webservices.data.idformat(v=exchg.80).aspx

### -TrustAnySSL

Switch to trust any certificate.

### -DateFormat

Here you can define the format for the timestamp LastModifiedTime. By default the current culture will be enumerated and the milliseconds appended. To have the same format as the CmdLets use "yyyyMMddThhmmssfff"

### -WebServicesDLL

Path to the DLL

### -StartDate

Filter by Start date of a single appointment. Note: Cannot be used to find recurring meetings!

### -EndDate

Filter by End date of a single appointment. Note: Cannot be used to find recurring meetings!

### -StartDateTimeCreated

When filter by Datetimecreated, all items created after this date are returned.

### -EndDateTimeCreated

When filter by Datetimecreated, all items created before this date are returned.

### -UseLocalTime

When this switch is used, DateTimeCreated and LastModifiedTime is converted to local time of the machine where the script is running on.

### -UseOAuth

Use OAuth for authentication.

### -UserPrincipalName

UserPrincipalName when using OAuth. This is optional.

### -ADALPath

Path to 

### -ClientId

ClientId when using OAuth.

### -ConnectionUri

ConnectionUri when using OAuth.

### -RedirectUri

RedirectUri when using OAuth.

### -PromptBehavior

PromptBehavior when using OAuth.

## Links

### [Troubleshooting calendar items] (https://ingogegenwarth.wordpress.com/2015/05/01/troubleshooting-calendar-items/)

### [Advanced troubleshooting calendar items] (https://ingogegenwarth.wordpress.com/2017/11/20/advanced-cal/)

### [EWS and OAuth] (https://ingogegenwarth.wordpress.com/2018/08/02/ews-and-oauth/)

## License

This project is licensed under the MIT License - see the LICENSE.md for details.



