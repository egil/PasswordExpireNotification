using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.DirectoryServices.AccountManagement;
using System.Reflection;

namespace Assimilated.PasswordExpireNotification
{
    class Program
    {
        private static TimeSpan _mpa;
        private static int _nd;
        private static string _group;
        private static string _es;
        private static FileInfo _mtf;
        private static string _mt;
        private static string _se;
        private static string _subject;
        private static string _testEmail;

        internal class UserInfo
        {
            public string Name { get; set; }
            public string EmailAddress { get; set; }
            public int DaysLeft { get; set; }
            public int HoursLeft { get; set; }
            public int TotalHoursLeft { get; set; }
        }

        static void Main(string[] args)
        {
            Console.WriteLine("####################################");
            Console.WriteLine("    Password Expire Notification");
            Console.WriteLine("    Version: {0}", GetProductVersion());
            Console.WriteLine("####################################");
            Console.WriteLine();

            CheckAndExtratSuppliedArguments(new Arguments(args));
            
            var today = DateTime.Now;

            Console.WriteLine("Settings:\n");
            Console.WriteLine("  Group: {0}", _group);
            Console.WriteLine("  Maximum password age: {0}", _mpa.Days);
            Console.WriteLine("  Notification days: {0}", _nd);
            Console.WriteLine("  Message subject: {0}", _subject);
            Console.WriteLine("  Message template: {0}", _mtf.FullName);
            Console.WriteLine("  Sender email: {0}", _se);
            Console.WriteLine("  Email server: {0}", _es);
            if (!string.IsNullOrEmpty(_testEmail))
            {
                Console.WriteLine("  Test Email: {0}", _testEmail);
            }
            Console.WriteLine("  Date and time: {0}{1}", today, Environment.NewLine);

            var usersToEmail = new List<UserInfo>();

            try
            {
                using (var context = new PrincipalContext(ContextType.Domain))
                using (var group = GroupPrincipal.FindByIdentity(context, _group))
                {
                    if (group != null)
                    {
                        usersToEmail.AddRange(from UserPrincipal user in @group.GetMembers()
                                              let lastPasswordSet = user.LastPasswordSet
                                              where
                                                  lastPasswordSet != null && !user.PasswordNeverExpires &&
                                                  lastPasswordSet.HasValue
                                              let left = lastPasswordSet.Value.AddDays(_mpa.Days).Subtract(today)
                                              let daysLeft = left.Days
                                              let hoursLeft = left.Hours
                                              let totalHours = Convert.ToInt32(Math.Floor(left.TotalHours))                        
                                              where left.TotalDays < _nd
                                              select
                                                  new UserInfo
                                                      {
                                                          Name = user.DisplayName,
                                                          EmailAddress = string.IsNullOrEmpty(_testEmail) ? user.EmailAddress : _testEmail,
                                                          DaysLeft = daysLeft,
                                                          HoursLeft = hoursLeft,
                                                          TotalHoursLeft = totalHours
                                                      });
                    }
                    else
                    {
                        Console.WriteLine("No group found that matched the search parameter. Aborting.");
                        Environment.Exit(1);
                    }
                }
            }
            catch (MultipleMatchesException)
            {
                Console.WriteLine("Multiple groups was found. Aborting.");
                Environment.Exit(1);
            }

            Console.WriteLine("Found {0} users whoes password expires in {1} days or less.{2}", usersToEmail.Count, _nd, Environment.NewLine);

            if (usersToEmail.Count > 0)
            {
                // load the template text into memory
                try
                {
                    using (StreamReader sr = File.OpenText(_mtf.FullName))
                    {
                        _mt = sr.ReadToEnd();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Cound not read the email template file.");
                    Console.WriteLine(ex.Message);
                    Environment.Exit(1);
                }

                // email each user
                usersToEmail.ForEach(SendEmailToUser);
            }

            Console.WriteLine();
            Console.WriteLine("Done!");
        }

        private static void SendEmailToUser(UserInfo userInfo)
        {
            var body = ReplaceTokens(_mt, userInfo);
            var subject = ReplaceTokens(_subject, userInfo);

            using (var message = new MailMessage(_se, userInfo.EmailAddress, subject, body))
            {
                var client = new SmtpClient(_es);
                try
                {
                    Console.Write("Sending email to: {0} . . . ", userInfo.Name);
                    // Add credentials if the SMTP server requires them.
                    client.Credentials = CredentialCache.DefaultNetworkCredentials;
                    client.Send(message);
                    Console.Write("Sent!{0}", Environment.NewLine);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error while trying to send email.");
                    Console.WriteLine(ex.Message);
                    Environment.Exit(1);
                }                
            }
        }
        private static string ReplaceTokens(string text, UserInfo userInfo)
        {
            return text.Replace("{Name}", userInfo.Name)
              .Replace("{DaysLeft}", userInfo.DaysLeft.ToString(CultureInfo.InvariantCulture))
              .Replace("{HoursLeft}", userInfo.HoursLeft.ToString(CultureInfo.InvariantCulture))
              .Replace("{TotalHoursLeft}", userInfo.TotalHoursLeft.ToString(CultureInfo.InvariantCulture));
        }

        private static TimeSpan GetMaxPasswordAge()
        {
            using (Domain d = Domain.GetCurrentDomain())
            using (DirectoryEntry domain = d.GetDirectoryEntry())
            {
                var ds = new DirectorySearcher(domain, "(objectClass=*)", null, SearchScope.Base);
                var sr = ds.FindOne();
                var maxPwdAge = TimeSpan.MinValue;
                if (sr.Properties.Contains("maxPwdAge")) maxPwdAge = TimeSpan.FromTicks((long)sr.Properties["maxPwdAge"][0]);
                return maxPwdAge.Duration();
            }
        }

        static string GetProductVersion()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            FileVersionInfo fileVersionInfo = FileVersionInfo.GetVersionInfo(assembly.Location);
            return fileVersionInfo.ProductVersion;
        }

        static void CheckAndExtratSuppliedArguments(Arguments arguments)
        {
            if (arguments.Count == 0)
            {
                Console.WriteLine("  Missing required arguments.");
                Console.WriteLine("  See https://github.com/egil/PasswordExpireNotification for help.");
                Environment.Exit(0);
            }

            _group = arguments["group"] ?? arguments["g"];
            if (string.IsNullOrEmpty(_group))
            {
                Console.WriteLine("Missing argument: -group or -g{0}", Environment.NewLine);

                Console.WriteLine("Please specify which security group contains the users who should be notified when their user account password is about to expire.{0}", Environment.NewLine);

                Console.WriteLine("You can specify the group using one of the following identifiers:{0}", Environment.NewLine);

                Console.WriteLine(" - Sam Account Name: The identity is a Security Account Manager (SAM) name.");
                Console.WriteLine(" - Name: The identity is a name.");
                Console.WriteLine(" - User Principal Name: The identity is a User Principal Name (UPN).");
                Console.WriteLine(" - Distinguished Name: The identity is a Distinguished Name (DN).");
                Console.WriteLine(" - Sid: The identity is a Security Identifier (SID) in Security Descriptor Definition Language (SDDL) format.");
                Console.WriteLine(" - Guid: The identity is a Globally Unique Identifier (GUID).{0}", Environment.NewLine);
                Environment.Exit(0);
            }

            if ((string.IsNullOrEmpty(arguments["notification-days"]) && string.IsNullOrEmpty(arguments["nd"])) ||
                (!int.TryParse(arguments["notification-days"], out _nd) && !int.TryParse(arguments["nd"], out _nd)))
            {
                Console.WriteLine("Missing or invalid argument: -notification-days or -nd{0}", Environment.NewLine);
                Console.WriteLine("Please specify number of days before expiration to start notifying users.{0}", Environment.NewLine);
                Environment.Exit(0);
            }

            var mt = arguments["message-template"] ?? arguments["mt"];
            if (string.IsNullOrEmpty(mt) || !File.Exists(mt))
            {
                Console.WriteLine("Missing or invalid argument: -message-template or -mt{0}", Environment.NewLine);
                Console.WriteLine("Please specify the path to the e-mail template text file used when sending out notifiations.{0}", Environment.NewLine);
                Environment.Exit(0);
            }

            _es = arguments["email-server"] ?? arguments["es"];
            if (string.IsNullOrEmpty(_es))
            {
                Console.WriteLine("Missing or invalid argument: -email-server or -es{0}", Environment.NewLine);
                Console.WriteLine("Please specify email server name, either as an IP or hostname.{0}", Environment.NewLine);
                Environment.Exit(0);
            }

            _se = arguments["sender-email"] ?? arguments["se"];
            if (string.IsNullOrEmpty(_se))
            {
                Console.WriteLine("Missing or invalid argument: -sender-email or -se{0}", Environment.NewLine);
                Console.WriteLine("Please specify the from address to use for sending out e-mail notifications.{0}", Environment.NewLine);
                Environment.Exit(0);
            }

            _subject = arguments["subject"] ?? arguments["s"] ?? "Your password is about to expire!";
            
            // max password age
            var mpaInput = arguments["max-password-age"] ?? arguments["mpa"] ?? string.Empty;
            int mpa;
            if (int.TryParse(mpaInput, out mpa))
            {
                _mpa = TimeSpan.FromDays(mpa);
            }
            else
            {
                _mpa = GetMaxPasswordAge();
            }

            // force all emails to specific address
            _testEmail = arguments["test-email"] ?? arguments["te"];

            _mtf = new FileInfo(mt);
        }
    }
}
