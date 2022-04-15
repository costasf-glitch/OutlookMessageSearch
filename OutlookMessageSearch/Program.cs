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

        private static List<string> namespaceFolders = new List<string>();

        // Utility method to decode the URI for the task (mostly because browser's will replace special
        // characters like the '<' and '>' found in the URI
        private static string base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        // Get a list of the EntryIDs for each folder
        // We need this because entering a folder requires the EntryID
        private static void getFolders()
        {
            Console.WriteLine("Listing all the folders now");
            Folders outlookFolders = outlookNameSpace.Folders;

            foreach (Folder selectedFolder in outlookFolders)
            {
                crawlFolders(selectedFolder);
            }
        }

        private static List <string> crawlFolders(Folder selectedFolder)
        {
            namespaceFolders.Add(selectedFolder.EntryID);

            // if the selected folder has subfolders, do some recursion through the function
            // to add it to the list
            if (selectedFolder.Folders.Count != 0)
            {
                foreach (Folder selectedSubFolder in selectedFolder.Folders)
                {
                    crawlFolders((Folder)selectedSubFolder);
                }
            }
            
            // This is just for me because I have a number of folders and it seems that
            // the most useful ones are at the end of the list.
            namespaceFolders.Reverse();

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

            // Look through each folder for the specific message
            Console.WriteLine($"Searching in {outlookFolder.Name}...");
            string filter = $"@SQL=\"http://schemas.microsoft.com/mapi/proptag/0x1035001F\" = '{messageID}'";
            MailItem foundMessage = outlookItems.Find(filter);

            if(foundMessage != null)
            {
                Console.WriteLine($"Found the message with subject: {foundMessage.Subject}");
                string strCommandText = $"/select \"outlook:{foundMessage.EntryID}\"";

                System.Diagnostics.Process.Start("outlook.exe", strCommandText);
                
                return true;
            }
            else
            {
                return false;
            }
        }

        // main function takes a command line argument which is the message ID
        // example being: <BN7PR05MB4322089E8B3FD6E287A05319C85B9@BN7PR05MB4322.namprd05.prod.outlook.com>
        static void Main(string[] args)
        {
            Console.WriteLine($"Searching for message with Base64 String: {args[0]}");
            string decodedMessageID = base64Decode(args[0].Substring(10));
            
            Console.WriteLine($"Searching for message with MessageID {decodedMessageID}");
            
            // Get a listing of all the folders within the Outlook Namespace
            getFolders();

            // Look for a specific MessageID in every folder that we listed
            foreach(string name in namespaceFolders)
            {

                if (findMessageByID(decodedMessageID, name))
                        break;
            }

            // Explicitly release the Outlook Interop objects...
            outlookInstance = null;
            outlookNameSpace = null;
        }
    }
}
