# Password Expire Notification #
A small command line tool that will send a “Your password is about to expire” email to all members of a domain group if their password will expire within a specified number of days.

I created this tool to help forgetful mobile users who loses access to their email or VPN if their password expires. 

## Download ##

- [PasswordExpireNotification.exe](https://raw.github.com/egil/PasswordExpireNotification/master/bin/Release/PasswordExpireNotification.exe) - only a sigle executable needed (32/64 bit)
- [NotificationTemplate.txt](https://raw.github.com/egil/PasswordExpireNotification/master/NotificationTemplate.txt) - a sample template text file.

## Requirements ##

- .NET framework 3.5 or higher
- A user account with sufficient permissions to query Active Directory and send emails through the specified SMTP server.

## Command line arguments ##

`-group` or `-g`: The security group that contains the users who should be notified when their user account password is about to expire. You can specify the group using one of the following identifiers:

 - Sam Account Name: The identity is a Security Account Manager (SAM) name.
 - Name: The identity is a name.
 - User Principal Name: The identity is a User Principal Name (UPN).
 - Distinguished Name: The identity is a Distinguished Name (DN).
 - Sid: The identity is a Security Identifier (SID) in Security Descriptor Definition Language (SDDL) format.
 - Guid: The identity is a Globally Unique Identifier (GUID).

`-max-password-age` or `-mpa`: The maximum password age of the domain.

`-notification-days` or `-nd`: The number of days before expiration to start notifying users.

`-message-template` or `-mt`: The path to the e-mail template text file used when sending out notifications. Path can be either relative or absolute.

*Replacement patterns that can be used in the template are: `{DaysLeft}`, `{HoursLeft}` and `{TotalHoursLeft}`.*

`-email-server` or `-es`: The email server name, either as an IP or hostname.

`-sender-email` or `-se`: The from address to use for sending out e-mail notifications.

`-subject` or `-s`: [Optional] The subject of the email. Default is *Your password is about to expire!*

*Replacement patterns that can be used in the subject are: `{DaysLeft}`, `{HoursLeft}` and `{TotalHoursLeft}`.*

`-test-email` or `te`: [Optional] Send all mails to the specified email address instead. Useful when testing templates, so you do not bother your users.

### Example: ###

**For testing:**

`PasswordExpireNotification.exe -g:"Password Expire Notification" -mpa:120 -nd:2 -mt:"NotificationTemplate.txt" -es:mx.example.com -se:administrator@example.dk -s:"Your password will expire in {TotalHoursLeft} hours"` **`-te:"administrator@example.com"`**

**For production with custom subject:**

`PasswordExpireNotification.exe -g:"Password Expire Notification" -mpa:120 -nd:2 -mt:"NotificationTemplate.txt" -es:mx.example.com -se:administrator@example.dk -s:"Your password will expire in {TotalHoursLeft} hours"`

**For production:**

`PasswordExpireNotification.exe -g:"Password Expire Notification" -mpa:120 -nd:2 -mt:"NotificationTemplate.txt" -es:mx.example.com -se:administrator@example.dk`

## Author information ##
My name is Egil Hansen and you can find more of my stuff here on GitHub or on my personal site http://egilhansen.com.

If you have any suggestions or find a bug, please add submit a new issue to the [issues](issues) page, or even better, for submit a patch.
