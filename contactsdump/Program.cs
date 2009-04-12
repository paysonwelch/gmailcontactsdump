using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Contacts;
using System.IO;
using OutLook = Microsoft.Office.Interop.Outlook;

namespace contactsdump
{
    class Program
    {
        static void Main(string[] args)
        {
            string whatDoYouWant = @"
What do you want?
    1. Transfer from one gmail account to another
    2. Consolidation one gmail account
    3. Dump one gmail account's contacts to your local Outlook
Please type in the number and hit enter (default 1)";
            Console.WriteLine(whatDoYouWant);

            switch (Console.ReadLine())
            {
                case "1":
                    Transfer();
                    break;
                case "2":
                    Consolidate();
                    break;
                case "3":
                    DumpToOutlook();
                    break;
                default:
                    Transfer();
                    break;
            }
        }

        private static void DumpToOutlook()
        {
            ContactsService srcService = new ContactsService("src");
            string srcUsername = string.Empty;
            string srcPassword = string.Empty;
            Console.WriteLine("Type in source user name:");
            srcUsername = Console.ReadLine();
            Console.WriteLine("Type in source password:");
            Console.ForegroundColor = ConsoleColor.Black;
            srcPassword = Console.ReadLine();
            Console.ForegroundColor = ConsoleColor.Gray;

            srcService.Credentials = new Google.GData.Client.GDataCredentials(srcUsername, srcPassword);

            var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            query.NumberToRetrieve = 10000;
            var results = srcService.Query(query);

            Dictionary<string, List<ContactEntry>> workBench = new Dictionary<string, List<ContactEntry>>();
            foreach (ContactEntry r in results.Entries)
            {
                Console.WriteLine(r.Title.Text);
                var app = new OutLook.ApplicationClass();
                var contactFolder = app.Session.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
                var newContact = contactFolder.Items.Add(OutLook.OlItemType.olContactItem) as OutLook.ContactItem;

                newContact.FullName = r.Title.Text;
                //phone
                if (r.Phonenumbers.Count > 3)
                {
                    newContact.Home2TelephoneNumber = r.Phonenumbers[3].Value;
                }
                if (r.Phonenumbers.Count > 2)
                {
                    newContact.HomeTelephoneNumber = r.Phonenumbers[2].Value;
                }
                if (r.Phonenumbers.Count > 1)
                {
                    newContact.Business2TelephoneNumber = r.Phonenumbers[1].Value;
                }
                if (r.Phonenumbers.Count > 0)
                {
                    newContact.BusinessTelephoneNumber = r.Phonenumbers[0].Value;
                }
                //email
                if (r.Emails.Count > 2)
                {
                    newContact.Email3Address = r.Emails[2].Address;
                }
                if (r.Emails.Count > 1)
                {
                    newContact.Email2Address = r.Emails[1].Address;
                }
                if (r.Emails.Count > 0)
                {
                    newContact.Email1Address = r.Emails[0].Address;
                }

                newContact.Save();
            }
        }

        private static void Transfer()
        {
            ContactsService srcService = new ContactsService("src");
            ContactsService destService = new ContactsService("dest");
            string srcUsername = string.Empty;
            string srcPassword = string.Empty;
            string destUsername = string.Empty;
            string destPassword = string.Empty;
            Console.WriteLine("Type in source user name:");
            srcUsername = Console.ReadLine();
            Console.WriteLine("Type in source password:");
            Console.ForegroundColor = ConsoleColor.Black;
            srcPassword = Console.ReadLine();
            Console.ForegroundColor = ConsoleColor.Gray;

            Console.WriteLine("Type in destination user name:");
            destUsername = Console.ReadLine();
            Console.WriteLine("Type in destination password:");
            Console.ForegroundColor = ConsoleColor.Black;
            destPassword = Console.ReadLine();
            Console.ForegroundColor = ConsoleColor.Gray;

            srcService.Credentials = new Google.GData.Client.GDataCredentials(srcUsername, srcPassword);
            destService.Credentials = new Google.GData.Client.GDataCredentials(destUsername, destPassword);

            var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            query.NumberToRetrieve = 10000;
            var results = srcService.Query(query);
            //if (Directory.Exists(rootPath) == false)
            //{
            //    Directory.CreateDirectory(rootPath);
            //}
            foreach (ContactEntry r in results.Entries)
            {
                Console.WriteLine(r.Title.Text);
                r.Title.Text = r.Title.Text.Trim('"', ' ', '\'');
                r.GroupMembership.Clear();
                Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));
                var response = destService.Insert<ContactEntry>(feedUri, r);
            }
        }

        private static void Consolidate()
        {
            ContactsService srcService = new ContactsService("src");
            string srcUsername = string.Empty;
            string srcPassword = string.Empty;
            Console.WriteLine("Type in source user name:");
            srcUsername = Console.ReadLine();
            Console.WriteLine("Type in source password:");
            Console.ForegroundColor = ConsoleColor.Black;
            srcPassword = Console.ReadLine();
            Console.ForegroundColor = ConsoleColor.Gray;

            srcService.Credentials = new Google.GData.Client.GDataCredentials(srcUsername, srcPassword);

            var query = new ContactsQuery(ContactsQuery.CreateContactsUri("default"));
            query.NumberToRetrieve = 10000;
            var results = srcService.Query(query);

            Dictionary<string, List<ContactEntry>> workBench = new Dictionary<string, List<ContactEntry>>();
            foreach (ContactEntry r in results.Entries)
            {
                Console.WriteLine(r.Title.Text);
                r.Title.Text = r.Title.Text.Trim('"', ' ', '\'');
                r.GroupMembership.Clear();
                if (workBench.ContainsKey(r.Title.Text.ToLower()) == false)
                {
                    workBench[r.Title.Text.ToLower()] = new List<ContactEntry>();
                }

                workBench[r.Title.Text.ToLower()].Add(r);
            }

            Uri feedUri = new Uri(ContactsQuery.CreateContactsUri("default"));

            foreach (var key in workBench.Keys)
            {
                if (key.Trim() == string.Empty)
                {
                    //delete records without title.
                    //these are mostly autogenerated records. 
                    //not very useful if I don't know who they are.
                    foreach (var e in workBench[key])
                    {
                        srcService.Delete(e);
                    }
                    continue;
                }
                var baseRecord = workBench[key][0];

                //now they are the same guy. let's move them together to the first entry
                for (int i = 1; i < workBench[key].Count; i++)
                {
                    var r = workBench[key][i];
                    if (string.IsNullOrEmpty(baseRecord.Summary.Text))
                    {
                        baseRecord.Summary.Text = r.Summary.Text;
                    }
                    //emails
                    var newEmails = baseRecord.Emails.Union(r.Emails, new EmailComparer()).ToList();
                    baseRecord.Emails.Clear();
                    foreach (var e in newEmails)
                    {
                        e.Primary = false;
                        baseRecord.Emails.Add(e);
                    }
                    if (baseRecord.Emails.Count > 0)
                    {
                        baseRecord.Emails[0].Primary = true;
                    }
                    //phones
                    var newPhones = baseRecord.Phonenumbers.Union(r.Phonenumbers, new PhoneComparer()).ToList();
                    baseRecord.Phonenumbers.Clear();
                    foreach (var e in newPhones)
                    {
                        e.Primary = false;
                        baseRecord.Phonenumbers.Add(e);
                    }
                    if (baseRecord.Phonenumbers.Count > 0)
                    {
                        baseRecord.Phonenumbers[0].Primary = true;
                    }
                    //addresses
                    var newAddresses = baseRecord.PostalAddresses.Union(r.PostalAddresses).ToList();
                    baseRecord.PostalAddresses.Clear();
                    foreach (var e in newAddresses)
                    {
                        e.Primary = false;
                        baseRecord.PostalAddresses.Add(e);
                    }
                    if (baseRecord.PostalAddresses.Count > 0)
                    {
                        baseRecord.PostalAddresses[0].Primary = true;
                    }

                    //delete the rest
                    srcService.Delete(r);
                }
                //update base record
                var response = srcService.Update<ContactEntry>(baseRecord);
            }
        }
    }
}
