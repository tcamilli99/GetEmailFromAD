using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Security.AccessControl;
using System.Security.Principal;
using Microsoft.Win32;


namespace GetEmailFromAD
{
    class Program
    {
        static void Main(string[] args)
        {
                    
            String emailAddress = "";

            if (IsDomainAccount())
            {
                // PC is configured to log into a domain                
               emailAddress = GetEmailFromAD();

                if (emailAddress.Length > 0)
                    Console.WriteLine("Email Address from ActiveDirectory:  " + emailAddress);
                else
                {
                    // Couldn't get email from AD, possible offline. Attempt to get cached info from registry
                    emailAddress = GetEmailFromRegistry();

                    if (emailAddress.Length > 0)
                        Console.WriteLine("Email Address Cached in Registry:  " + emailAddress);
                }             
                           
            }
            else
            {
                // PC is not configured to login to a domain, won't be able to get AD information
                emailAddress = GetEmailFromMicrosoftAccount();

                if (emailAddress.Length > 0)
                    Console.WriteLine("Email Address from Microsoft Account:  " + emailAddress);

            }

            // If email found, output, otherwise report error
           if (emailAddress.Length == 0)
                Console.WriteLine("Could determine logged-on user's email address");

            return;
        }

        static bool IsDomainAccount()
        {
            // Simple check for domain.  Could possibly do something more sophisticated like checking for or connecting to a domain controller.
            // TODO:  Make sure this doesn't still pass when configured for workgroup.
            
            if ((Environment.UserDomainName.Length > 0) && (Environment.MachineName != Environment.UserDomainName))
                return true;
            else
                return false;

        }

        static String GetEmailFromAD()
        {
            // This will attempt to run a search of AD using DirectorySearcher.
            // Will return email address if sucessful or empty string if not.
            String emailAddress = "";

            //New way
            try
            {
                UserPrincipal user = UserPrincipal.Current;
                
                if (user != null)
                {
                    if (user.EmailAddress.Length > 0)
                        emailAddress = user.EmailAddress;
                }
            }
            catch
            {
                //Console.WriteLine("Error searching AD");
            }

            return emailAddress;
        }

        static String GetEmailFromRegistry()
        {
            // This will attempt to search the registry for AD information.
            // Will return email address if successful and empty string if not.

            String emailAddress = "";
            String dnNameString;
            
            // AD information is containied within a key that includes the user SID.  First get the SID for the current user.
            WindowsIdentity winID = WindowsIdentity.GetCurrent();
            String key = "HKEY_LOCAL_MACHINE\\Software\\Microsoft\\Windows\\CurrentVersion\\Group Policy\\DataStore\\" + winID.Owner.Value + "\\0";

            try
            {
                dnNameString = Registry.GetValue(key, "DNName", "").ToString();
            }
            catch
            {
                return "";
            }

            if (dnNameString.Contains("CN="))
            {
                // Contains a CN element
                int iStart = dnNameString.IndexOf("CN=") + 3;
                int iEnd = dnNameString.IndexOf(',', iStart);

                if (iEnd == -1)
                {
                    // "CN" element was the only or at the end of the string, not sure if this will ever happend but covering bases
                    emailAddress = dnNameString.Substring(iStart, dnNameString.Length - 3);
                }
                else
                {
                    emailAddress = dnNameString.Substring(iStart, iEnd - iStart);
                }
            }
            else
            {
                //Registry didn't have cached info or we need to parse additional SIDs and/or enumerations within the SID key
                Console.WriteLine("Couldn't finded cached info in registry");
            }        

            return emailAddress;
        }

        static String GetEmailFromMicrosoftAccount()
        {
            String emailAddress = "";

            WindowsIdentity winID = WindowsIdentity.GetCurrent();
            IdentityReferenceCollection idRefCollection = winID.Groups;
            String account;

            foreach (IdentityReference idRef in idRefCollection)
            {
                try
                {
                    account = idRef.Translate(typeof(NTAccount)).Value;
                    if(account.Contains("MicrosoftAccount"))
                    {
                        int iStart = @"MicrosoftAccount\\".Length - 1;
                        int iLength = account.Length - iStart;
                        emailAddress = account.Substring(iStart, iLength);
                        return emailAddress;
                    }                    
                }
                catch
                {
                    //Console.WriteLine("Found IdentityReference:" + idRef.Value);
                }
                
            }  

            return emailAddress;
        }

    }
}
