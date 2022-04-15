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
        
        // Get a list of the EntryIDs for each folder
        // We need this because entering a folder requires the EntryID
        private static List<string> getFolders()
        {
            List<string> namespaceFolders = new List<string>();
            
            Console.WriteLine("Listing all the folders now");
            Folders outlookFolders = outlookNameSpace.Folders;

            foreach (Folder selectedFolder in outlookFolders)
            {
                namespaceFolders.Add(selectedFolder.EntryID);
                
                //Console.WriteLine($"Searching through: {selectedFolder.Name}");

                // TODO: This does not go through subfolders.  Will need to adjust code later
                foreach (Folder selectedSubFolder in selectedFolder.Folders)
                {
                  //  Console.WriteLine($"{selectedSubFolder.Name}: Folder ID is: {selectedSubFolder.EntryID}");
                    namespaceFolders.Add(selectedSubFolder.EntryID);
                }
            }
            return namespaceFolders;
        }

        private static string getMessageID(MailItem currentMessage)
        {
            const string PR_INTERNET_MESSAGE_ID_W_TAG = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
            PropertyAccessor oPropAccessor = currentMessage.PropertyAccessor;
            
            string messageID = (string)oPropAccessor.GetProperty(PR_INTERNET_MESSAGE_ID_W_TAG);
            return messageID;
        }

        private static bool findMessageByID(string messageID, string folderEntryID)
        {
            // Select the active folder
            Folder outlookFolder = (Folder)outlookNameSpace.GetFolderFromID(folderEntryID);

            // Get a collection of items in the selected folder
            Items outlookItems = outlookFolder.Items;

            // Can be sorted by any property in the item type.  In this case, the item type should be
            // MailItem, and we are sorting by Received Time in from newest to oldest
            //outlookItems.Sort("[ReceivedTime]", true);

            Console.WriteLine($"Searching in {outlookFolder.Name}...");
            string filter = $"@SQL=\"http://schemas.microsoft.com/mapi/proptag/0x1035001F\" = '{messageID}'";
            MailItem foundMessage = outlookItems.Find(filter);

            if(foundMessage != null)
            {
                Console.WriteLine($"Subject name here is {foundMessage.Subject}");
                string strCommandText = $"/select \"outlook:{foundMessage.EntryID}\"";

                System.Diagnostics.Process.Start("outlook.exe", strCommandText);
                
                return true;
            }
            else
            {
                return false;
            }
            
            // MAY NOT NEED THE CODE BELOW IF THE CODE ABOVE IS ROBUST
            // Iterate through all MailItems in the folder looking for a specific EntryID
            //foreach (Object outlookItem in outlookItems)
            //{
            //    if (outlookItem is MailItem)
            //    {
            //        // convert object to a MailItem for further processing
            //        MailItem mailItem = outlookItem as MailItem;

            //        // extract the message id from the current message
            //        string currentMessageID = getMessageID(mailItem);
            //        //Console.WriteLine($"Message received on: {mailItem.ReceivedTime} from: {mailItem.SenderName}");

            //        // compare it to the message id we are looking for
            //        if(currentMessageID == messageID)
            //        {
            //            Console.WriteLine($"We found the message with the subject {mailItem.Subject}");

            //            Console.WriteLine("Attempting to open a specific Outlook message via the shell");

            //            string strCommandText = $"/select \"outlook:{mailItem.EntryID}\"";

            //            System.Diagnostics.Process.Start("outlook.exe", strCommandText);
                        
            //            return true;
            //        }
            //    }
            //}
            //return false;
        }

        // main function takes a command line argument which is the message ID
        // example being: <BN7PR05MB4322089E8B3FD6E287A05319C85B9@BN7PR05MB4322.namprd05.prod.outlook.com>
        static void Main(string[] args)
        {
            List<string> folderNames = getFolders();

            foreach(string name in folderNames)
            {
                //findMessageByID("<BN7PR05MB4322089E8B3FD6E287A05319C85B9@BN7PR05MB4322.namprd05.prod.outlook.com>", name);
                // if(findMessageByID(args[0], "000000006C5F7C2868207944B3CA43011342594901005AC540F88BEA6E4995B7DAB7B8B93FC300000000013C0000"))
                if (findMessageByID(args[0], name))
                        break;
            }

            // Explicitly release objects
            outlookInstance = null;
            outlookNameSpace = null;
        }
    }
}
