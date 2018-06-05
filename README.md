# Get-CalendarItems

Retrieves Calendar items of specified mailboxes for troubleshooting.

### Prerequisites

* FullAccess or ApplicationImpersonation
* Microsoft Exchange Web Services (EWS) API

### Usage

The script has multiple parameters. Only a few are shown in this example:

```
.\Get-CalendarItems.ps1 -EmailAddress manager@contoso.com,assistant@contoso.com -Impersonate -Credentials (Get-Credential) -Subject '1:1 with Rob' 
```

### About

For more information on this script, as well as usage and examples, see
the related blog article on [The Clueless Guy](https://ingogegenwarth.wordpress.com/2015/05/01/troubleshooting-calendar-items/#Script).

## License

This project is licensed under the MIT License - see the LICENSE.md for details.