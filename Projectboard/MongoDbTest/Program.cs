using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MongoDbAccess;
using Microsoft.VisualBasic;


namespace MongoDbTest
{
    class Program
    {
        static void Main(string[] args)
        {
            MainAsync(args).GetAwaiter().GetResult();
        }


        static async Task MainAsync(string[] args) { 
            Request r = new Request("ds129342.mlab.com:29342", "pmichallenge", "pk", "test");

            //r.clearDocuments(); //removes all documents from db

            Console.WriteLine(@"Enter filepath of the document that you want to upload to the database (eg. 'D:\Philipp\Desktop\Datenblatt.pdf' ):");
            string filepath = @Console.ReadLine();//"D:\\Philipp\\Desktop\\Datenblatt.pdf";
            var id = await r.StoreDocumentToDBAsync(filepath, 1, "pk", new String[]{"pk"}, "desc");

            Console.WriteLine(@"Enter filepath where you want the downloaded document to be stored on the filesystem (eg. 'D:\Philipp\Desktop\blub.pdf' ):");
            String path = @Console.ReadLine();
            Console.WriteLine("document is stored at " + await r.retrieveDocumentFromDBByIdAsync(id, "pk", path));

            Console.WriteLine(@"Enter a filename for which you want to list all revisions that are in the database (eg. 'Datenblatt.pdf' ):");
            String name = @Console.ReadLine();
            r.FindAllRevisionsOfDocumentInDB(name);

            Console.WriteLine("Press enter to close terminal.");
            Console.ReadLine();
        }
    }
}
