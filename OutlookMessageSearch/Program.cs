/* OutlookMessageSearch
 * 
 * This is a proof-of-concept to basically see if we can reproduce some of the functionality
 * from the defunct Todoist Outlook plugin.  This program will ultimately be registered as a 
 * protocol handler to open up direct links to Outlook messages
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;

namespace OutlookMessageSearch
{
    internal class Program
    {
        private static Application outlookInstance = new Application();
        private static NameSpace outlookNameSpace = outlookInstance.GetNamespace("mapi");
        
        static void Main(string[] args)
        {
            // Get the inbox folder
            Folder outlookFolder = (Folder)outlookNameSpace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            // Get a collection of items in the selected folder
            Items outlookItems = outlookFolder.Items;

            
            // Can be sorted by any property in the item type.  In this case, these are MailItem
            outlookItems.Sort("[ReceivedTime]",false);

            // Get the first (or last) message (for testing purposes to see if this works)
            MailItem outlookMessage = (MailItem)outlookItems.GetFirst();
            outlookMessage = (MailItem)outlookItems.GetNext();


            // Iterate through all MailItems in the folder looking for a specific EntryID
            foreach (Object outlookItem in outlookItems)
            {
                if (outlookItem is MailItem)
                {
                    MailItem mailItem = outlookItem as MailItem;
                    Console.WriteLine(mailItem.Subject);
                }
                else
                {
                    Console.WriteLine("SOMETHING THAT IS NOT AN EMAIL");
                }
            }

            // Code Snippet: Iterate through all Folders in my namespace
            Console.WriteLine("Listing all the folders now");
            Folders outlookFolders = outlookNameSpace.Folders;

            foreach(Folder selectedFolder in outlookFolders)
            {
                Console.WriteLine($"Searching through: {selectedFolder.Name}");

                foreach(Folder selectedSubFolder in selectedFolder.Folders)
                {
                    Console.WriteLine(selectedSubFolder.Name);
                }
            }

            // See if this worked
            //Console.WriteLine(outlookMessage.Subject);
            //Console.WriteLine(outlookMessage.EntryID);


            // Extract a Message ID from the active message
            const string PR_INTERNET_MESSAGE_ID_W_TAG = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
            PropertyAccessor oPropAccessor = outlookMessage.PropertyAccessor;
            string messageID = (string)oPropAccessor.GetProperty(PR_INTERNET_MESSAGE_ID_W_TAG);
            Console.WriteLine(messageID);


            Console.WriteLine("Attempting to open a specific Outlook message via the shell");

            string strCommandText = $"/select \"outlook:{outlookMessage.EntryID}\"";

            Console.WriteLine(strCommandText);

            System.Diagnostics.Process.Start("outlook.exe", strCommandText);
            Console.ReadKey();
        }
    }
}
