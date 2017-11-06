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
            Request r = new Request("ds129342.mlab.com:29342", "pmichallenge", "pk", "test");

            //r.clearDocuments();
            
            // Console.WriteLine("Is documents collection empty?");
            // bool empty = r.collectionEmpty("documents.files");
            // Console.WriteLine(empty);

            string filepath = "D:\\Philipp\\Desktop\\Datenblatt.pdf";
            var id = r.StoreDocumentToDB(filepath, 1, "pk", new String[]{"pk"}, "desc");

            String path = "D:\\Philipp\\Desktop\\download.pdf";
            Console.WriteLine(r.retrieveDocumentFromDB(id, "pk", path));
            r.FindAllDocumentRevisionsInDB("Datenblatt.pdf");

            Console.ReadLine();
        }
    }
}
