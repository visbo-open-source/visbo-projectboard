using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;

using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Builders;
using System.Collections;
using MongoDB.Driver.Linq;

using ProjectBoardDefinitions;

namespace MongoDbAccess
{
    public class Request
    {
        public MongoClient Client { get; set; }
        public MongoServer Server { get; set; }
        public MongoDatabase Database { get; set; }
        public MongoCollection CollectionProjects { get; set; }
        public MongoCollection CollectionConstellations { get; set; }
        public MongoCollection CollectionDependencies { get; set; }

        public Request()
        {
            Client = new MongoClient("mongodb://localhost");
            Server = Client.GetServer();
            Database = Server.GetDatabase("projectboard");
            CollectionProjects = Database.GetCollection<clsProjektDB>("projects");
            CollectionConstellations = Database.GetCollection<clsConstellationDB>("constellations");
            CollectionDependencies = Database.GetCollection<clsDependenciesOfPDB>("dependencies");
        }

        public Request(string databaseName)
        {
            Client = new MongoClient("mongodb://localhost");
            Server = Client.GetServer();
            Database = Server.GetDatabase(databaseName);
            CollectionProjects = Database.GetCollection<clsProjektDB>("projects");
            CollectionConstellations = Database.GetCollection<clsConstellationDB>("constellations");
            CollectionDependencies = Database.GetCollection<clsDependenciesOfPDB>("dependencies");
        }

        public bool collectionEmpty(string name)
        {
            return Database.GetCollection<clsProjektDB>(name).Count() == 0;
        }

        public bool storeProjectToDB(clsProjekt projekt)
        {
            var projektDB = new clsProjektDB();
            projektDB.copyfrom(projekt);
            projektDB.Id = projektDB.name + "#" + projektDB.variantName + "#" + projektDB.timestamp.ToString();
            return CollectionProjects.Save(projektDB).Ok;      


            //projektDB.copyfrom(ref projekt); wenn von kopiert wird, muss das nicht per Ref übergeben werden 
            //neues Dokument speichern falls letztes Backup länger als 1 tag her, sonst aktuelle version überschreiben:
            //projektDB.timestamp = DateTime.Today;
            //projektDB.timestamp = DateTime.Now; projektDB.timestamp wird in copyfrom gesetzt ....
            //projektDB.Id = projektDB.name + "#" + projektDB.variantName + "#" + projektDB.timestamp.ToString("yyyy-MM-dd");
                           
        }


        public SortedList<string, clsProjekt> retrieveProjectsFromDB(string projectname, string variantName, DateTime zeitraumStart, DateTime zeitraumEnde, DateTime storedEarliest, DateTime storedLatest, bool onlyLatest)
        {
            var result = new SortedList<string, clsProjekt>();

            if (onlyLatest)
            {
                int startMonat = (int)DateAndTime.DateDiff(DateInterval.Month, Module1.StartofCalendar, zeitraumStart)+1;
                

                var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                            .Where(c => c.tfSpalte >= startMonat-Module1.maxProjektdauer && c.startDate <= zeitraumEnde)
                            .Select(c => c.name)
                            .Distinct();

                foreach (string name in prequery)
                {
                    var projektDB = CollectionProjects.AsQueryable<clsProjektDB>()
                                 .Where(c => c.name == name)
                                 .OrderBy(c => c.timestamp)
                                 .Last();
                    //TODO: rückumwandeln
                    
                    if (projektDB.tfSpalte + projektDB.Dauer >= startMonat  )
                    
                    {
                        var projekt = new clsProjekt();
                        projektDB.copyto(ref projekt);
                        string schluessel = projekt.name + '#' + projekt.variantName;
                        //result.Add(projekt.Id, projekt);
                        result.Add(schluessel, projekt);
                    }

                   
                }
            }

            else
            {
                //gegeben: Projektname, Backupzeitraum (also storedEarliest, storedLatest)
                var projects = from e in CollectionProjects.AsQueryable<clsProjektDB>()
                               where e.name == projectname
                               where e.variantName == variantName
                               where e.timestamp >= storedEarliest && e.timestamp <= storedLatest
                               select e;

                foreach (clsProjektDB p in projects)
                {
                    //TODO: rückumwandeln
                    var projekt = new clsProjekt();
                    p.copyto(ref projekt);
                    // wichtig: in p steht das timestamp in utc format, in projekt in localtime
                    string schluessel = projekt.timeStamp.ToString();
                    //result.Add(projekt.Id, projekt);
                    result.Add(schluessel, projekt);
                }
            }

            return result;
        }

        public SortedList<DateTime, clsProjekt> retrieveProjectHistoryFromDB(string projectname, string variantName, DateTime storedEarliest, DateTime storedLatest)
        {
            var result = new SortedList<DateTime, clsProjekt>();

           
                //gegeben: Projektname, Backupzeitraum (also storedEarliest, storedLatest)
                var projects = from e in CollectionProjects.AsQueryable<clsProjektDB>()
                               where e.name == projectname
                               where e.variantName == variantName
                               where e.timestamp >= storedEarliest && e.timestamp <= storedLatest
                               select e;

                foreach (clsProjektDB p in projects)
                {
                    //TODO: rückumwandeln
                    var projekt = new clsProjekt();
                    p.copyto(ref projekt);
                    //string schluessel = p.timestamp.ToString();
                    DateTime schluessel = p.timestamp;
                    //result.Add(projekt.Id, projekt);
                    result.Add(schluessel, projekt);
                }
            

            return result;
        }

        public bool storeConstellationToDB(clsConstellation c)
        {
            var cDB = new clsConstellationDB();
            cDB.copyfrom(ref c);
            cDB.Id = cDB.constellationName;
            return CollectionConstellations.Save(cDB).Ok;
        }

        public clsConstellations retrieveConstellationsFromDB()
        {
            var result = new clsConstellations();


            var constellationsDB = CollectionConstellations.AsQueryable<clsConstellationDB>()
                                 .Select(cDB => cDB);
            foreach (clsConstellationDB cDB in constellationsDB)
            {
                var c = new clsConstellation();
                cDB.copyto(ref c);
                result.Add(c);
              
            }
            
            

            return result;
        }

        public bool storeDependencyofPToDB(clsDependenciesOfP d)
        {
            var depDB = new clsDependenciesOfPDB();
            depDB.copyFrom(d);
            depDB.Id = depDB.projectName;
            return CollectionDependencies.Save(depDB).Ok;
        }

        public clsDependencies  retrieveDependenciesFromDB()
        {
            var result = new clsDependencies();


            var DependenciesOfPDB = CollectionDependencies.AsQueryable<clsDependenciesOfPDB>()
                                 .Select(depDB => depDB);
            foreach (clsDependenciesOfPDB depDB in DependenciesOfPDB)
            {
                var newDofP = new clsDependenciesOfP();
                depDB.copyTo(ref newDofP);
                result.Add(newDofP, true);
            }



            return result;
        }

    }
}
