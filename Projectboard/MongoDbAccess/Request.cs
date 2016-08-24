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
using MongoDbAccess.Properties;


namespace MongoDbAccess
{
    public class Request
    {
        public MongoClient Client { get; set; }

        // neu 3.0 
        protected static IMongoClient newClient;

        public MongoServer Server { get; set; }
        public MongoDatabase Database { get; set; }


        // neu 3.0 
        protected static IMongoDatabase newDatabase;


        public MongoCollection CollectionProjects { get; set; }

        // neu 3.0
        protected static IMongoCollection<clsProjektDB> newCollectionProjects { get; set; }
        
        public MongoCollection CollectionConstellations { get; set; }
        // neu 3.0 
        protected static IMongoCollection<clsConstellationDB> newCollectionConstellations { get; set; }

        public MongoCollection CollectionDependencies { get; set; }
        public MongoCollection CollectionFilter { get; set; }

        ////////public Request()
        ////////{
        ////////    var connectionString = "mongodb://localhost";
        ////////    /**var connectionString = "mongodb://ute:Mopsi@localhost"; Aufruf mit MongoDB mit Authentication */
        ////////    Client = new MongoClient(connectionString);
        ////////    Server = Client.GetServer();
        ////////    Database = Server.GetDatabase("projectboard");
        ////////    CollectionProjects = Database.GetCollection<clsProjektDB>("projects");
        ////////    CollectionConstellations = Database.GetCollection<clsConstellationDB>("constellations");
        ////////    CollectionDependencies = Database.GetCollection<clsDependenciesOfPDB>("dependencies");
        ////////    CollectionFilter = Database.GetCollection<clsFilterDB>("filters");
        ////////}

        public Request(string databaseURL, string databaseName, string username, string dbPasswort)
        {
            //var databaseURL = "localhost";
            if (String.IsNullOrEmpty(username) && String.IsNullOrEmpty(dbPasswort))
            {
                var connectionString = "mongodb://" + databaseURL;
                //var connectionString = "mongodb://@ds034198.mongolab.com:34198";
                Client = new MongoClient(connectionString);
                newClient = new MongoClient(connectionString);
            }
            else
            {

                var connectionString = "mongodb://" + username + ":" + dbPasswort + "@" + databaseURL + "/" + databaseName;  /*Aufruf mit MongoDB mit Authentication  */
                //var connectionString = "mongodb://" + username + ":" + dbPasswort + "@ds034198.mongolab.com:34198";
                Client = new MongoClient(connectionString);
            }
            
            Server = Client.GetServer();
            Database = Server.GetDatabase(databaseName);

            // neu 3.0 
            newDatabase = newClient.GetDatabase(databaseName);
            newCollectionProjects = newDatabase.GetCollection<clsProjektDB>("projects");

            CollectionProjects = Database.GetCollection<clsProjektDB>("projects");
            CollectionConstellations = Database.GetCollection<clsConstellationDB>("constellations");
            CollectionDependencies = Database.GetCollection<clsDependenciesOfPDB>("dependencies");
            CollectionFilter = Database.GetCollection<clsFilterDB>("filters");

        }

        public bool collectionEmpty(string name)
        {
            return Database.GetCollection<clsProjektDB>(name).Count() == 0;
        }

        /** prüft ob der Projektname schon vorhanden ist (ggf. inkl. VariantName)
         *  falls Variantname null ist oder leerer String wird nur der Projektname überprüft.
         */
        public bool projectNameAlreadyExists(string projectname, string variantname)
        {
            bool result;
            // in der Datenbank ist der Projektname abgespeichert als projectName#variantName, wenn es einen Varianten-Namen gibt
            // nur projectname , sonst (hat historische Gründe .. weil sonst nach Einführung der Varianten alle bisherigen Projekt-Namen in der Datenbank
            // Namen geändert werden müssten .. )

            string searchstr = Projekte.calcProjektKeyDB(projectname, variantname);
            result = CollectionProjects.AsQueryable<clsProjektDB>()
                    .Any(c => c.name == searchstr);

            
            return result;

            //if (variantname != null && variantname.Length > 0)
            //{
            //    string searchstr = Projekte.calcProjektKey(projectname, variantname);
            //    result = CollectionProjects.AsQueryable<clsProjektDB>()
            //            .Any(c => c.name == searchstr);
            //}
            //else
            //    result = CollectionProjects.AsQueryable<clsProjektDB>()
            //            .Any(c => c.name == projectname);
            
        }

        /** liest ein bestimmtes Projekt aus der DB (ggf. inkl. VariantName)
         *  falls Variantname null ist oder leerer String wird nur der Projektname überprüft.
         */
        public clsProjekt retrieveOneProjectfromDB(string projectname, string variantname)
        {
            var result = new clsProjektDB();
            string searchstr = Projekte.calcProjektKeyDB(projectname, variantname);
            result = CollectionProjects.AsQueryable<clsProjektDB>()
                    .Where(c => c.name == searchstr)
                    .Last();

            //if (variantname != null && variantname.Length > 0)
            //{
            //    string searchstr = Projekte.calcProjektKey(projectname, variantname);
            //    result = CollectionProjects.AsQueryable<clsProjektDB>()
            //            .Where(c => c.name == searchstr )
            //            .Last();
            //}
            //else
            //    result = CollectionProjects.AsQueryable<clsProjektDB>()
            //            .Where(c => c.name == projectname)
            //            .Last();
                       

            //TODO: rückumwandeln
            var projekt = new clsProjekt();
            result.copyto(ref projekt);
            return projekt;
        }

        /**
         * prüft die Verfügbarkeit der MongoDB
         */
        public bool pingMongoDb()
        {
            bool ping;
            try
            {
                Server.Ping();
                ping = true;
            }
            catch (Exception e)
            {
                ping = false;
            }
            return ping;
        }

        public bool storeProjectToDB(clsProjekt projekt)
        {
            try
            {
                var projektDB = new clsProjektDB();
                bool ergebnis;
                //string xx = "";
                projektDB.copyfrom(projekt);
                projektDB.Id = projektDB.name + "#" + projektDB.variantName + "#" + projektDB.timestamp.ToString();
                ergebnis = !CollectionProjects.Save(projektDB).HasLastErrorMessage;
                //xx = CollectionProjects.Save(projektDB).LastErrorMessage;
                //return !CollectionProjects.Save(projektDB).HasLastErrorMessage;    
                return ergebnis;
            }
            catch
            {
                return false;
            }
              


            //projektDB.copyfrom(ref projekt); wenn von kopiert wird, muss das nicht per Ref übergeben werden 
            //neues Dokument speichern falls letztes Backup länger als 1 tag her, sonst aktuelle version überschreiben:
            //projektDB.timestamp = DateTime.Today;
            //projektDB.timestamp = DateTime.Now; projektDB.timestamp wird in copyfrom gesetzt ....
            //projektDB.Id = projektDB.name + "#" + projektDB.variantName + "#" + projektDB.timestamp.ToString("yyyy-MM-dd");
                           
        }
        //************************************/
        public bool deleteProjectHistoryFromDB(string projectname, string variantName, DateTime storedEarliest, DateTime storedLatest)
        {
            
            storedLatest = storedLatest.ToUniversalTime();
            storedEarliest = storedEarliest.ToUniversalTime();
            string searchstr = Projekte.calcProjektKeyDB(projectname, variantName);
            

            var query = Query < clsProjektDB >
                //.Where(p => (p.name == projectname && p.variantName == variantName && p.timestamp >= storedEarliest && p.timestamp <= storedLatest));
                .Where(p => (p.name == searchstr && p.timestamp >= storedEarliest && p.timestamp <= storedLatest));

            var query2 = Query.And(
                    Query<clsProjektDB>.EQ(p => p.name, projectname),
                    Query<clsProjektDB>.EQ(p => p.variantName, variantName),
                    Query<clsProjektDB>.GTE(p => p.timestamp, storedEarliest),
                    Query<clsProjektDB>.LTE(p => p.timestamp, storedLatest)
                );
            
            return !CollectionProjects.Remove(query).HasLastErrorMessage;
        }

        //************************************/
        public bool deleteProjectTimestampFromDB(string projectname, string variantName, DateTime stored)
        {

            stored = stored.ToUniversalTime();
            string searchstr = Projekte.calcProjektKeyDB(projectname, variantName);


            var query = Query<clsProjektDB>
                        .Where(p => (p.name == searchstr && p.timestamp == stored));

            
            return !CollectionProjects.Remove(query).HasLastErrorMessage;
        }

        public SortedList<string, clsProjekt> retrieveProjectsFromDB(string projectname, string variantName, DateTime zeitraumStart, DateTime zeitraumEnde, DateTime storedEarliest, DateTime storedLatest, bool onlyLatest)
        {
            var result = new SortedList<string, clsProjekt>();
            
            // in der Datenbank sind die Zeiten als Universal time gespeichert .. 
            // deshalb muss hier umgerechnet werden 
            storedLatest = storedLatest.ToUniversalTime();
            storedEarliest = storedEarliest.ToUniversalTime();

            if (onlyLatest)
            {
                int startMonat = (int)DateAndTime.DateDiff(DateInterval.Month, Module1.StartofCalendar, zeitraumStart)+1;
                

                var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                            //.Where(c => c.tfSpalte >= startMonat-Module1.maxProjektdauer && c.startDate <= zeitraumEnde)
                            .Where(c => c.startDate <= zeitraumEnde && c.endDate >= zeitraumStart)
                            .Select(c => c.name)
                            .Distinct();

                
                foreach (string name in prequery)
                {
                    
                    //// Ergänzt 15.10.14
                    //// wieder zurückgenommen, weil jetzt in der Datenbank gespeichert wird, daß ein Projektname 
                    //// pName#vName ist, sofern es einen variantName gibt 

                    
                    //var prequeryV = CollectionProjects.AsQueryable<clsProjektDB>()
                    //    //.Where(c => c.tfSpalte >= startMonat-Module1.maxProjektdauer && c.startDate <= zeitraumEnde)
                    //    //    .Where(c => c.startDate <= zeitraumEnde && c.endDate >= zeitraumStart)
                    //        .Where(c => c.name == name)
                    //        .Select(c => c.variantName)
                    //        .Distinct();

                    //foreach (string vName in prequeryV)
                    //{
                    //    var projektDB0 = CollectionProjects.AsQueryable<clsProjektDB>()
                    //             .Where(c => c.name == name && c.variantName == vName)
                    //             .OrderBy(c => c.timestamp)
                    //             .Last();
                    //    //TODO: rückumwandeln

                    //    if (projektDB0.tfSpalte + projektDB0.Dauer >= startMonat)
                    //    {
                    //        var projekt = new clsProjekt();
                    //        projektDB0.copyto(ref projekt);
                    //        string schluessel = projekt.name + '#' + projekt.variantName;
                    //        //result.Add(projekt.Id, projekt);
                    //        result.Add(schluessel, projekt);
                    //    }
                    //}


                    //// Ende Ergänzung 15.10.14
                    
                    // Start alter Code vor Ergänzung 15.10 - der ist jetzt (ab 16.10.14) wieder gültig 
                    var projektDB = CollectionProjects.AsQueryable<clsProjektDB>()
                                 .Where(c => c.name == name)
                                 .OrderBy(c => c.timestamp)
                                 .Last();
                    //TODO: rückumwandeln

                    if (projektDB.tfSpalte + projektDB.Dauer >= startMonat)
                    {
                        var projekt = new clsProjekt();
                        projektDB.copyto(ref projekt);
                        string schluessel = projekt.name + '#' + projekt.variantName;
                        //result.Add(projekt.Id, projekt);
                        result.Add(schluessel, projekt);
                    }
                    
                    // Ende alter Code vor Ergänzung 15.10 - jetzt wieder der richtige Code
                   
                }
            }

            else
            {
                //gegeben: Projektname, Backupzeitraum (also storedEarliest, storedLatest)
                
                // in der Datenbank ist der Projektname als pName#vName gespeichert, wenn es eine Variante gibt
                // pName, sonst
                
                string searchstr = Projekte.calcProjektKeyDB(projectname, variantName); 


                //if (variantName != null && variantName.Length > 0)
                //   searchstr = Projekte.calcProjektKey(projectname, variantName);
                //else
                //    searchstr = projectname;

                
                var projects = from e in CollectionProjects.AsQueryable<clsProjektDB>()
                               where e.name == searchstr
                               // wird nicht mehr benötigt: where e.variantName == variantName
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


        public Collection retrieveVariantNamesFromDB(string projectName)
        {
            var result = new Collection();

            string trennzeichen = "#";
            string searchstr = string.Concat(projectName, trennzeichen);
                        
            //gegeben: Projektname, Backupzeitraum (also storedEarliest, storedLatest)
            //var projects = from e in CollectionProjects.AsQueryable<clsProjektDB>()
            //               where e.name.Contains(searchstr)
            //               select e.variantName
            //               .Distinct();


            var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                            .Where(c => c.name.Contains(searchstr))
                            .Select(c => c.variantName)
                            .Distinct();

            foreach (string vName in prequery)
            {
                result.Add(vName);
            }

            return result;
        }

        
        //
        // gibt die Projekthistorie innerhalb eines gegebenen Zeitraums zu einem gegebenen Projekt+Varianten-Namen zurück
        //
        public SortedList<DateTime, clsProjekt> retrieveProjectHistoryFromDB(string projectname, string variantName, DateTime storedEarliest, DateTime storedLatest)
        {
            var result = new SortedList<DateTime, clsProjekt>();

            storedLatest = storedLatest.ToUniversalTime();
            storedEarliest = storedEarliest.ToUniversalTime();

            // in der Datenbank ist der Projektname als pName#vName gespeichert, wenn es eine Variante gibt
            // pName, sonst
            
            string searchstr = Projekte.calcProjektKeyDB(projectname, variantName); 


            //if (variantName != null && variantName.Length > 0)
            //    searchstr = Projekte.calcProjektKey(projectname, variantName);
            //else
            //    searchstr = projectname;

            //gegeben: Projektname, Backupzeitraum (also storedEarliest, storedLatest)
            var projects = from e in CollectionProjects.AsQueryable<clsProjektDB>()
                               where e.name == searchstr
                               // wird nicht mehr benötigt where e.variantName == variantName
                               where e.timestamp >= storedEarliest && e.timestamp <= storedLatest
                               select e;

            foreach (clsProjektDB p in projects)
                {
                    //TODO: rückumwandeln
                    var projekt = new clsProjekt();
                    p.copyto(ref projekt);
                    
                    DateTime schluessel = projekt.timeStamp;
                    result.Add(schluessel, projekt);
                }
            

            return result;
        }

        public bool storeConstellationToDB(clsConstellation c)
        {
            var cDB = new clsConstellationDB();
            cDB.copyfrom(ref c);
            cDB.Id = cDB.constellationName;
            return !CollectionConstellations.Save(cDB).HasLastErrorMessage;
           
        }

        public bool removeConstellationFromDB(clsConstellation c)
        {
            //var cDB = new clsConstellationDB();
            //cDB.copyfrom(ref c);
            //cDB.Id = cDB.constellationName;
            var query = Query<clsConstellationDB>.EQ(e => e.Id, c.constellationName);
            return !CollectionConstellations.Remove(query).HasLastErrorMessage;
        }

        //
        // benennt alle Projekte mit Namen oldName um
        // aber nur, wenn der neue Name nicht schon in der Datenbank existiert 
        public async void renameProjectsInDB(string oldName, String newName)
        {
            if (projectNameAlreadyExists(newName, ""))
            {
                // return false;
            }
            
            {
                // erstmal das Projekt selber umbenennen 
                string oldFullName = Projekte.calcProjektKeyDB(oldName, "");
                string newFullName = Projekte.calcProjektKeyDB(newName, "");

                // neu 3.0 
                var filter = Builders<clsProjektDB>.Filter.Eq("name", oldFullName);
                var update = Builders<clsProjektDB>.Update
                                    .Set("name", newFullName);

                var ergebnis = await newCollectionProjects.UpdateManyAsync(filter, update);

                // jetzt 
                // alle Varianten des Projektes umbenennen 
                Collection listOfVariants = retrieveVariantNamesFromDB(oldName);

               
                foreach (string vName in listOfVariants )
                {
                    oldFullName = Projekte.calcProjektKeyDB(oldName, vName);
                    newFullName = Projekte.calcProjektKeyDB(newName, vName);

                    // neu 3.0 
                    filter = Builders<clsProjektDB>.Filter.Eq("name", oldFullName);
                    update = Builders<clsProjektDB>.Update
                                    .Set("name", newFullName);

                    ergebnis = await newCollectionProjects.UpdateManyAsync(filter, update);
                                       
                }
            }
            // return true;
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


        // * speichert Dependencies in DB 
        public bool storeDependencyofPToDB(clsDependenciesOfP d)
        {
            var depDB = new clsDependenciesOfPDB();
            depDB.copyFrom(d);
            depDB.Id = depDB.projectName;
            return !CollectionDependencies.Save(depDB).HasLastErrorMessage;
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

        /** liest einen bestimmten Filter aus der DB               */

        public clsFilter retrieveOneFilterfromDB(string filtername)
        {
            var result = new clsFilterDB();
            string searchstr = filtername;
            result = CollectionFilter.AsQueryable<clsFilterDB>()
                    .Where(c => c.name == searchstr)
                    .Last();
     
            //TODO: rückumwandeln
            var filter = new clsFilter();
            result.copyto(ref filter);
            return filter;
        }

        /** speichert einen Filter mit Namen 'name' in der Datenbank*/

        public bool storeFilterToDB(clsFilter filter, Boolean selfilter)
        {
            var filterDB = new clsFilterDB();
            filterDB.copyfrom( ref filter,  selfilter);
            filterDB.Id = filter.name;
            return !CollectionFilter.Save(filterDB).HasLastErrorMessage;
        }
        /** löscht einen bestimmten Filter aus der Datenbank */

        public bool removeFilterFromDB(clsFilterDB filter)
        {
            //var cDB = new clsConstellationDB();
            //cDB.copyfrom(ref c);
            //cDB.Id = cDB.constellationName;
           
            var query = Query<clsFilterDB>
                .Where(e => (e.name == filter.name));
            return !CollectionFilter.Remove(query).HasLastErrorMessage;
        }

        /** liest alle Filter aus der Datenbank */
        public SortedList<String, clsFilter> retrieveAllFilterFromDB(Boolean selfilter)
        {
            var result = new SortedList<String, clsFilter>();

            var filterDB = CollectionFilter.AsQueryable<clsFilterDB>()
                                 .Select(cDB => cDB);
            foreach (clsFilterDB cDB in filterDB)
            {
                if (selfilter == cDB.selFilter)
                {
                    var f = new clsFilter();
                    cDB.copyto(ref f);
                    result.Add(f.name, f);
                }
            }

            return result;
        }
    }
}
