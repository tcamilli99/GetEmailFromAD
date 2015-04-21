using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.Security.AccessControl;
using System.Security.Principal;
using Microsoft.Win32;


namespace GetEmailFromAD
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.WriteLine("Getting email address from AD");
            String emailAddress = "";

            if (IsDomainAccount())
            {
                // PC is configured to log into a domain
                String account = Environment.UserName;

                if ((account.Length > 0))
                {
                    emailAddress = GetEmailFromAD(account);

                    if (emailAddress.Length == 0)
                    {
                        // Couldn't get email from AD, possible offline. Attempt to get cached info from registry
                        emailAddress = GetEmailFromRegistry(account);                       
                    }

                }
                else
                {
                    // Couldn't get the logged in user (running as system?)
                    // TODO:  Add logic to detect other users.  Could possibly walk the HKEY_USERS registry hive.
                }                
            }
            else
            {
                // PC is not configured to login to a domain, won't be able to get AD information
                Console.WriteLine("Error:  This PC is not a mmember of a domain");
            }

            // If email found, output, otherwise report error
            if (emailAddress.Length > 0)
                Console.WriteLine("Detected user email address: " + emailAddress);
            else
                Console.WriteLine("Could determine logged-on user's email address");

            return;
        }

        static bool IsDomainAccount()
        {
            // Simple check for domain.  Could possibly do something more sophisticated like checking for or connecting to a domain controller.
            // TODO:  Make sure this doesn't still pass when configured for workgroup.
            if (Environment.UserDomainName.Length > 0)
                return true;
            else
                return false;
        }

        static String GetEmailFromAD(string user)
        {
            // This will attempt to run a search of AD using DirectorySearcher.
            // Will return email address if sucessful or empty string if not.

            String emailAddress = "";
            DirectorySearcher search = new DirectorySearcher();

            search.Filter = "(&(objectClass=user)(anr=" + user + "))";
            search.PropertiesToLoad.Add("mail");

            try {
                SearchResult result = search.FindOne();
                emailAddress = result.Properties["mail"][0].ToString();
            }
            catch 
            {
                Console.WriteLine("Error searching AD");
            }

            return emailAddress;
        }

        static String GetEmailFromRegistry(string account)
        {
            // This will attempt to search the registry for AD information.
            // Will return email address if successful and empty string if not.

            String emailAddress = "";
            String dnNameString;

            // AD information is containied within a key that includes the user SID.  First get the SID for the current user.
            WindowsIdentity winID = WindowsIdentity.GetCurrent();
            String key = "HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows\\CurrentVersion\\Group Policy\\DataStore\\" + winID.Owner.Value + "\\0";

            dnNameString = Registry.GetValue(key, "DNName", "").ToString();

            if (dnNameString.Contains("CN="))
            {
                // Contains a CN element
                int indexStart = dnNameString.IndexOf("CN=") + 3;
                int indexEnd = dnNameString.IndexOf(',', indexStart);

                if (indexEnd == -1)
                {
                    // "CN" element was the only or at the end of the string, not sure if this will ever happend but covering bases
                    emailAddress = dnNameString.Substring(indexStart, dnNameString.Length - 3);
                }
                else
                {
                    emailAddress = dnNameString.Substring(indexStart, indexEnd - indexStart);
                }
            }
            else
            {
                //Registry didn't have cached info or we need to parse additional SIDs and/or enumerations within the SID key
                Console.WriteLine("Couldn't finded cached info in registry");
            }        

            return emailAddress;
        }

    }
}
