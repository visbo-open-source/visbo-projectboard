using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic;

using MongoDB.Bson;
using MongoDB.Driver;
using MongoDB.Driver.Core.Authentication;
using MongoDB.Driver.Builders;
using System.Collections;
using MongoDB.Driver.Linq;

using ProjectBoardDefinitions;
using MongoDbAccess.Properties;


namespace MongoDbAccess
{
    public class Request
    {
        // alt 2.x
        //public MongoClient Client { get; set; }
        //public MongoServer Server { get; set; }
        //public MongoDatabase Database { get; set; }
        //public MongoCollection CollectionProjects { get; set; }
        //public MongoCollection CollectionConstellations { get; set; }
        //public MongoCollection CollectionDependencies { get; set; }
        //public MongoCollection CollectionFilter { get; set; }

              
        // neu 3.0 
        protected  IMongoClient Client;
        protected  IMongoDatabase Database;
        protected MongoServer Server;
        protected  IMongoCollection<clsProjektDB> CollectionProjects;
        protected  IMongoCollection<clsProjektDB> CollectionTrashProjects;
        protected  IMongoCollection<clsConstellationDB> CollectionConstellations;
        protected IMongoCollection<clsConstellationDB> CollectionTrashConstellations; 
        protected  IMongoCollection<clsDependenciesOfPDB> CollectionDependencies;
        protected  IMongoCollection<clsFilterDB> CollectionFilter;
        
        public Request(string databaseURL, string databaseName, string username, string dbPasswort)
        {
            //var databaseURL = "localhost";
            if (String.IsNullOrEmpty(username) && String.IsNullOrEmpty(dbPasswort))
            {

                //var connectionString = "mongodb://" + databaseURL + "?connectTimeoutMS=30&SocketTimeoutMS=10";
                var connectionString = "mongodb://" + databaseURL; 

                //var connectionString = "mongodb://@ds034198.mongolab.com:34198";
                Client = new MongoClient(connectionString);
            }
            else

            {

                // wird nicht mehr verwendet , führt ggf zu Problemen bei zu schnellem Timeout 
                // var connectionString = "mongodb://" + username + ":" + dbPasswort + "@" + databaseURL + "/" + databaseName + "?connectTimeoutMS=30&SocketTimeoutMS=10";  /*Aufruf mit MongoDB mit Authentication  */
                var connectionString = "mongodb://" + username + ":" + dbPasswort + "@" + databaseURL + "/" + databaseName;
                
                //var connectionString = "mongodb://" + username + ":" + dbPasswort + "@ds034198.mongolab.com:34198";
                Client = new MongoClient(connectionString);
                
            }
            
            //alt 2.x
            //Server = Client.GetServer();
            //Database = Server.GetDatabase(databaseName);
  
            // neu 3.0 
            Database = Client.GetDatabase(databaseName);
            
                      
            CollectionProjects = Database.GetCollection<clsProjektDB>("projects");
            CollectionTrashProjects = Database.GetCollection<clsProjektDB>("trashprojects");
            CollectionConstellations = Database.GetCollection<clsConstellationDB>("constellations");
            CollectionTrashConstellations = Database.GetCollection<clsConstellationDB>("trashconstellations");
            CollectionDependencies = Database.GetCollection<clsDependenciesOfPDB>("dependencies");
            CollectionFilter = Database.GetCollection<clsFilterDB>("filters");

        }

        public  bool createIndicesOnce()
        {
            try
            {
                // wenn ein Index bereits existiert, wird er nicht mehr erzeugt ... 
                var keys = Builders<clsProjektDB>.IndexKeys.Ascending("timestamp");
                var ergebnis = CollectionProjects.Indexes.CreateOne(keys);
                string test = ergebnis;
                
                keys = Builders<clsProjektDB>.IndexKeys.Ascending("name");
                ergebnis = CollectionProjects.Indexes.CreateOne(keys);
                
                keys = Builders<clsProjektDB>.IndexKeys.Ascending("variantName");
                ergebnis = CollectionProjects.Indexes.CreateOne(keys);
                
                keys = Builders<clsProjektDB>.IndexKeys.Ascending("startDate");
                ergebnis = CollectionProjects.Indexes.CreateOne(keys);
                
                keys = Builders<clsProjektDB>.IndexKeys.Ascending("endDate");
                ergebnis = CollectionProjects.Indexes.CreateOne(keys);
                return true;
            }
            catch
            {
                return false;
            }
           
        }

        public bool collectionEmpty(string name)
        {
            //return Database.GetCollection<clsProjektDB>(name).Count() == 0;
            long result;
            switch (name)
            {
                case "projects":
                    result = CollectionProjects.AsQueryable<clsProjektDB>().Count();
                    break;
                case "constellations":
                    result = CollectionConstellations.AsQueryable<clsConstellationDB>().Count();
                    break;
                case "dependencies":
                    result = CollectionDependencies.AsQueryable<clsDependenciesOfPDB>().Count();
                    break;
                case "filters":
                    result = CollectionFilter.AsQueryable<clsFilterDB>().Count();
                    break;
                default:
                    result = 0;
                    break;
            }
            
            return result == 0; 
        }

        /** prüft ob der Projektname schon vorhanden ist (ggf. inkl. VariantName)
         *  falls Variantname null ist oder leerer String wird nur der Projektname überprüft.
         */
        public bool projectNameAlreadyExists(string projectname, string variantname, DateTime storedAtorBefore)
        {
            bool result;

            try 
            {
          
            // in der Datenbank ist der Projektname abgespeichert als projectName#variantName, wenn es einen Varianten-Namen gibt
            // nur projectname , sonst (hat historische Gründe .. weil sonst nach Einführung der Varianten alle bisherigen Projekt-Namen in der Datenbank
            // Namen geändert werden müssten .. )

            if (storedAtorBefore == null)
            {
                //storedAtorBefore = DateTime.SpecifyKind(DateTime.Now, DateTimeKind.Utc);
                storedAtorBefore = DateTime.Now.AddDays(1).ToUniversalTime();
            }
            else
            {
                //storedAtorBefore = DateTime.SpecifyKind(storedAtorBefore, DateTimeKind.Utc);
                storedAtorBefore = storedAtorBefore.ToUniversalTime();
            }

            string searchstr = Projekte.calcProjektKeyDB(projectname, variantname);
            result = CollectionProjects.AsQueryable<clsProjektDB>()
                    .Any(c => (c.name == searchstr && c.timestamp <= storedAtorBefore));

            
            return result;
                  }
            catch
            {
                throw new ArgumentException("Zugriff wurde von der Datenbank verweigert");
            }
                                   
        }

        /** liest ein bestimmtes Projekt aus der DB (ggf. inkl. VariantName)
         *  falls Variantname null ist oder leerer String wird nur der Projektname überprüft.
         */
        public clsProjekt retrieveOneProjectfromDB(string projectname, string variantname, DateTime storedAtOrBefore)
        {
            var result = new clsProjektDB();
            string searchstr = Projekte.calcProjektKeyDB(projectname, variantname);

            if (storedAtOrBefore == null)
            {
                
                //storedAtOrBefore = DateTime.SpecifyKind(DateTime.Now, DateTimeKind.Utc);
                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime();
            }
            else
            {
                //storedAtOrBefore = DateTime.SpecifyKind(storedAtOrBefore, DateTimeKind.Utc); 
                storedAtOrBefore = storedAtOrBefore.ToUniversalTime();
            }
            
            //var tmpErgebnis = CollectionProjects.AsQueryable<clsProjektDB>()
            //        .Where(c => c.name == searchstr)
            //        .OrderBy(c => c.timestamp)
            //        .Last();

            //var tmpErgebnis = (from c in CollectionProjects.AsQueryable<clsProjektDB>()
            //        where c.name == searchstr
            //        orderby c.timestamp
            //        select c)
            //        .Last();

            var builder = Builders<clsProjektDB>.Filter;
            var filter = builder.Eq("name", searchstr) & builder.Lte("timestamp", storedAtOrBefore);
            var sort = Builders<clsProjektDB>.Sort.Ascending("timestamp");

            try
            {
                result = CollectionProjects.Find(filter).Sort(sort).ToList().Last();
            }
            catch 
            {
                result = null;
            }
                        
            //TODO: rückumwandeln
            if (result == null)
            {
                
                return null;
            }
            else
            {
                var projekt = new clsProjekt();
                result.copyto(ref projekt);
                return projekt;
            }
            
        }

        /**
         * prüft die Verfügbarkeit der MongoDB
         */
        public bool pingMongoDb()
        {
            bool ping;
            try
            {
                if (Client == null)
                    { ping = false; }
                else
                    { ping = true; }

            }
            catch (Exception)
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
                //bool ergebnis;
                //string xx = "";
                projektDB.copyfrom(projekt);
                projektDB.Id = projektDB.name + "#" + projektDB.variantName + "#" + projektDB.timestamp.ToString();

                CollectionProjects.InsertOne(projektDB);
                // alt 2.x
                //ergebnis = !CollectionProjects.Save(projektDB).HasLastErrorMessage;
                //return ergebnis
                //xx = CollectionProjects.Save(projektDB).LastErrorMessage;
                //return !CollectionProjects.Save(projektDB).HasLastErrorMessage;    
                return true;
            }
            catch (Exception)
            {
                return false;
            }
              
                                       
        }

        public bool storeProjectToTrash(clsProjektDB projektDB)
        {
            try
            {
                
                CollectionTrashProjects.InsertOne(projektDB);
                  
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        //************************************/
        // darf nicht mehr verwendet werden, weil damit keine Speicherung der TimeStamps im Papierkorb verbunden ist ... 
        ////public bool deleteProjectHistoryFromDB(string projectname, string variantName, DateTime storedEarliest, DateTime storedLatest)
        ////{

        ////    try
        ////    {
        ////        storedLatest = storedLatest.ToUniversalTime();
        ////        storedEarliest = storedEarliest.ToUniversalTime();
        ////        string searchstr = Projekte.calcProjektKeyDB(projectname, variantName);
                               

        ////        var dResult = CollectionProjects.DeleteMany<clsProjektDB>(p => (p.name == searchstr && p.timestamp >= storedEarliest && p.timestamp <= storedLatest));
        ////        if (dResult.DeletedCount > 0 )
        ////            { return true; }
        ////        else
        ////            { return false; }
                
        ////    }
        ////    catch (Exception)
        ////    {
                
        ////        return false;
        ////    }
            
        ////    // das folgende geht noch nicht , wer weiss warum ? 
        ////    ////CollectionProjects.DeleteMany<clsProjektDB>(query);
        ////    ////CollectionProjects.DeleteMany<clsProjektDB>(query);
            
            
        ////    // alt 2.x 
        ////    //var query = Query<clsProjektDB>
        ////    //         .Where(p => (p.name == searchstr && p.timestamp >= storedEarliest && p.timestamp <= storedLatest));
        ////    //return !CollectionProjects.Remove(query).HasLastErrorMessage;
        ////}

        //************************************/
        public bool deleteProjectTimestampFromDB(string projectname, string variantName, DateTime stored)
        {
            try
            {
                stored = stored.ToUniversalTime();
                string searchstr = Projekte.calcProjektKeyDB(projectname, variantName);


                var query = Query<clsProjektDB>
                            .Where(p => (p.name == searchstr && p.timestamp == stored));

                
                var sResult = CollectionProjects.Find<clsProjektDB>(p => (p.name == searchstr && p.timestamp == stored));
                
                if (sResult == null)
                {
                    return false;
                }
                else
                {
                    try
                    {
                        clsProjektDB projektDB = sResult.Single();
                        if (storeProjectToTrash(projektDB))
                        {
                            // jetzt wird erst gelöscht 
                            var dResult = CollectionProjects.DeleteOne<clsProjektDB>(p => (p.name == searchstr && p.timestamp == stored));

                            if (dResult.DeletedCount > 0)
                            { return true; }
                            else
                            { return false; }
                        }
                        else
                        {
                            return false;
                        }

                        
                    }
                    catch (Exception)
                    {
                        return false; 
                    }
                                      
                }
                
               
            }
            catch (Exception)
            {
                
                return false;
            }
            
            
            // alt: 2.x 
            //return !CollectionProjects.Remove(query).HasLastErrorMessage;
        }
        /// <summary>
        /// liest alle vorkommenden Namen ProjektName#VariantenName aus der Datenbank , die zum Zeitpunkt storedLatest auch in der Datenbank existiert haben 
        /// dabei wird ein übergebener Zeitraum berücksichtigt ... also nur Projekte, die auch im Zeitraum liegen ...
        /// </summary>
        /// <param name="zeitraumStart"></param>
        /// <param name="zeitraumEnde"></param>
        /// <param name="storedEarliest"></param>
        /// <param name="storedatOrBefore"></param>
        /// <returns></returns>
        public SortedList<string, string> retrieveProjectVariantNamesFromDB(DateTime zeitraumStart, DateTime zeitraumEnde, DateTime storedatOrBefore)
        {
            var result = new SortedList<string, string>();

            // in der Datenbank sind die Zeiten als Universal time gespeichert .. 
            // deshalb muss hier umgerechnet werden 
            storedatOrBefore = storedatOrBefore.ToUniversalTime();
            
            int startMonat = (int)DateAndTime.DateDiff(DateInterval.Month, Module1.StartofCalendar, zeitraumStart) + 1;
            
                
            var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                            .Where(c => c.startDate <= zeitraumEnde && c.endDate >= zeitraumStart && c.timestamp <= storedatOrBefore)
                            .Select(c => c.name)
                            .Distinct()
                            .ToList();

            foreach (string name in prequery)
                {
                                        
                    try
                    {

                        if  (result.ContainsKey (name))  
                        {
                            // nichts tun 
                        }
                        else
                        {
                            result.Add(name, name);
                        }
                        

                    }
                    catch (Exception)
                    {

                        // nichts tun ...
                    }


                }
          

            return result;
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
                

                //var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                //            //.Where(c => c.tfSpalte >= startMonat-Module1.maxProjektdauer && c.startDate <= zeitraumEnde)
                //            .Where(c => c.startDate <= zeitraumEnde && c.endDate >= zeitraumStart)
                //            .Select(c => c.name)
                //            .Distinct();

                var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                    //      .Where(c => c.startDate <= zeitraumEnde && c.endDate >= zeitraumStart )
                            .Where(c => c.startDate <= zeitraumEnde && c.endDate >= zeitraumStart && c.timestamp <= storedLatest )
                            .Select(c => c.name)
                            .Distinct()
                            .ToList();

                foreach (string name in prequery)
                
                {
                    // im Vergleich zum alten: es muss um toList ergänzt werden :
                    //clsProjektDB projektDB = CollectionProjects.AsQueryable<clsProjektDB>()
                    //             .Where(c => c.name == name)
                    //             .OrderBy(c => c.timestamp)
                    //             .ToList()
                    //             .Last();

                    var filter = Builders<clsProjektDB>.Filter.Eq("name", name);
                    var sort = Builders<clsProjektDB>.Sort.Ascending("timestamp");

                    try
                    {
                                                                        
                        clsProjektDB projektDB = CollectionProjects.Find(filter).Sort(sort).ToList().Last();
                        var projekt = new clsProjekt();
                        projektDB.copyto(ref projekt);

                        string schluessel = Projekte.calcProjektKey(projekt);

                        result.Add(schluessel, projekt);
                        
                    }
                    catch (Exception)
                    {
                        
                        // nichts tun ...
                    }
                    
                           
                }
            }

            else
            {
                //gegeben: Projektname, Backupzeitraum (also storedEarliest, storedLatest)
                
                // in der Datenbank ist der Projektname als pName#vName gespeichert, wenn es eine Variante gibt
                // pName, sonst
                
                string searchstr = Projekte.calcProjektKeyDB(projectname, variantName); 

                               
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
                            .OrderBy(c => c.variantName)
                            .Select(c => c.variantName)
                            .ToList()
                            .Distinct();

            foreach (string vName in prequery)
            {
                result.Add(vName);
            }

            return result;
        }

        // bringt alle vorkommenden TimeStamps zurück , in absteigender Sortierung
        public Collection retrieveZeitstempelFromDB()
        {
            var result = new Collection();


            var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                            .OrderByDescending (c => c.timestamp)
                            .Select(c => c.timestamp)
                            .ToList()
                            .Distinct();

            foreach (DateTime tStamp in prequery)
            {
                DateTime tmpStamp = tStamp.ToLocalTime();
                result.Add(tmpStamp);
            }

            return result;
        }


        
        // bringt für die angegebene Projekt-Variante alle Zeitstempel in absteigender Sortierung zurück 
        public Collection retrieveZeitstempelFromDB(string pvName)
        {
            var result = new Collection();


            var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                            .Where(c => c.name == pvName)
                            .OrderByDescending(c => c.timestamp)
                            .Select(c => c.timestamp)
                            .ToList()
                            .Distinct();

            foreach (DateTime tStamp in prequery)
            {
                DateTime tmpStamp = tStamp.ToLocalTime();
                result.Add(tmpStamp);
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

            var builder = Builders<clsProjektDB>.Filter;
            var filter = builder.Eq("name", searchstr) & builder.Lte("timestamp", storedLatest);
            var sort = Builders<clsProjektDB>.Sort.Ascending("timestamp");
            //var result = await collection.Find(filter).Sort(sort).ToListAsync();
            var projects = CollectionProjects.Find(filter).Sort(sort).ToList();

            //gegeben: Projektname, Backupzeitraum (also storedEarliest, storedLatest)
            //var projects = from e in CollectionProjects.AsQueryable<clsProjektDB>()
            //                   where e.name == searchstr
            //                   // wird nicht mehr benötigt where e.variantName == variantName
            //                   where e.timestamp >= storedEarliest && e.timestamp <= storedLatest
            //                   select e;

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

            try
            {
                var cDB = new clsConstellationDB();
                cDB.copyfrom(c);
                cDB.Id = cDB.constellationName;

                bool alreadyExisting = CollectionConstellations.AsQueryable<clsConstellationDB>()
                        .Any(p => p.constellationName == c.constellationName);

                if (alreadyExisting)
                {
                    var filter = Builders<clsConstellationDB>.Filter.Eq("constellationName", c.constellationName);
                    
                    var rResult = CollectionConstellations.ReplaceOne(filter, cDB);
                    if (rResult.ModifiedCount > 0)
                    { return true; }
                    else
                    { return false; }
                }
                else
                {
                    CollectionConstellations.InsertOne(cDB);
                    return true;
                }
                
            }
            catch (Exception)
            {

                return false;
            }
            

            // alt 2.x
            //return !CollectionConstellations.Save(cDB).HasLastErrorMessage;
           
        }

        public bool storeConstellationToTrash(clsConstellation c)
        {
            try
            {
                var cDB = new clsConstellationDB();
                cDB.copyfrom(c);
                cDB.Id = cDB.constellationName;

                bool alreadyExisting = CollectionTrashConstellations.AsQueryable<clsConstellationDB>()
                        .Any(p => p.constellationName == c.constellationName);

                if (alreadyExisting)
                {
                    var filter = Builders<clsConstellationDB>.Filter.Eq("constellationName", c.constellationName);

                    var rResult = CollectionTrashConstellations.ReplaceOne(filter, cDB);
                    if (rResult.ModifiedCount > 0)
                    { return true; }
                    else
                    { return false; }
                }
                else
                {
                    CollectionTrashConstellations.InsertOne(cDB);
                    return true;
                }

            }
            catch (Exception)
            {

                return false;
            }
        }
        public bool removeConstellationFromDB(clsConstellation c)
        {
            
          try 
	        {	        
		    // neu 3.0 

              if (storeConstellationToTrash (c))
              {
                  var dResult = CollectionConstellations.DeleteOne<clsConstellationDB>(p => (p.Id == c.constellationName));
                  if (dResult.DeletedCount > 0)
                  { return true; }
                  else
                  { return false; }
              }
              else
              {
                  return false;
              }
            
           
            
            // alt 2.x
            //var query = Query<clsConstellationDB>.EQ(e => e.Id, c.constellationName);
            //return !CollectionConstellations.Remove(query).HasLastErrorMessage;
	        }
	      catch (Exception)
	        {
              return false;		  
	        }
           
        }

        //
        // benennt alle Projekte mit Namen oldName um
        // aber nur, wenn der neue Name nicht schon in der Datenbank existiert 
        public bool renameProjectsInDB(string oldName, String newName)
        {
            if (projectNameAlreadyExists(newName, "", DateTime.Now))
            {
                return false;
            }
            
            {

                try
                {
                    string oldFullName;
                    string newFullName;
                    bool ok = true;
                    // erstmal das Projekt selber umbenennen , falls es in der () Variante überhaupt existiert ..
                    if (projectNameAlreadyExists(oldName, "", DateTime.Now))
                    {
                        oldFullName = Projekte.calcProjektKeyDB(oldName, "");
                        newFullName = Projekte.calcProjektKeyDB(newName, "");

                        // neu 3.0 
                        var filter = Builders<clsProjektDB>.Filter.Eq("name", oldFullName);
                        var update = Builders<clsProjektDB>.Update
                                            .Set("name", newFullName);

                        var uResult = CollectionProjects.UpdateMany(filter, update);
                        ok = (uResult.ModifiedCount > 0); 
                        
                    }
                    

                    // jetzt 
                    // alle Varianten des Projektes umbenennen , wenn immer noch ok 

                    if (ok)
                    {
                        Collection listOfVariants = retrieveVariantNamesFromDB(oldName);


                        foreach (string vName in listOfVariants)
                        {
                            oldFullName = Projekte.calcProjektKeyDB(oldName, vName);
                            newFullName = Projekte.calcProjektKeyDB(newName, vName);

                            // neu 3.0 
                            var filter = Builders<clsProjektDB>.Filter.Eq("name", oldFullName);
                            var update = Builders<clsProjektDB>.Update
                                            .Set("name", newFullName);

                            var uResult = CollectionProjects.UpdateMany(filter, update);
                            ok = ok & (uResult.ModifiedCount > 0); 
                            
                        }

                       // jetzt müssen die Constellations aktualisiert werden ...

                       var constellationsDB = CollectionConstellations.AsQueryable<clsConstellationDB>()
                                 .Select(cDB => cDB);

                       int zaehler = 0;
                       int gesamt = 0; 

                       foreach (clsConstellationDB cDB in constellationsDB)
                        {
                            var c = new clsConstellation();
                            cDB.copyto(ref c);
                            int a = c.renameProject(oldName, newName);

                           if (a>0)
                           {
                               clsConstellationDB chgcDB = new clsConstellationDB();
                               chgcDB.copyfrom(c);
                               // mit Id=null kann kein Replace gemacht werden  
                               chgcDB.Id = cDB.Id;

                               var filter = Builders<clsConstellationDB>.Filter.Eq("constellationName", chgcDB.constellationName);
                               var rResult = CollectionConstellations.ReplaceOne(filter, chgcDB);
                               //ok = ok & (rResult.ModifiedCount > 0);
                               ok = ok & rResult.IsAcknowledged;

                               zaehler = zaehler + 1;
                               gesamt = gesamt + a; 
                           }
                            

                        }
                       // Énde Aktualisierung Constellations ...


                       // dann müssen noch die Dependencies aktualisiert werden ...

                        if (ok)
                        { return true; }
                        else
                        { return false; }
                        
                    }
                    else
                    { return false;  }
                    
                }
                catch (Exception)
                {
                    
                    return false;
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

            try
            {
                var depDB = new clsDependenciesOfPDB();
                depDB.copyFrom(d);
                depDB.Id = depDB.projectName;

                bool alreadyExisting = CollectionDependencies.AsQueryable<clsDependenciesOfPDB>()
                        .Any(p => p.projectName == d.projectName);

                if (alreadyExisting)
                {
                    var filter = Builders<clsDependenciesOfPDB>.Filter.Eq("projectName", d.projectName);
                    var rResult = CollectionDependencies.ReplaceOne(filter, depDB);
                    if (rResult.ModifiedCount > 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    CollectionDependencies.InsertOne(depDB);
                    return true;
                }
                
            }
            catch (Exception)
            {

                return false;
            }


            // alt 2.x
            
            //var depDB = new clsDependenciesOfPDB();
            //depDB.copyFrom(d);
            //depDB.Id = depDB.projectName;
                        
            //return !CollectionDependencies.Save(depDB).HasLastErrorMessage;
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
            var tmpListe = CollectionFilter.AsQueryable<clsFilterDB>()
                    .Where(c => c.name == searchstr);

            int anzahl;
            anzahl = tmpListe.Count();

            int zaehler = 0;
            foreach (clsFilterDB p in tmpListe)
            {
                zaehler = zaehler + 1;
                if (zaehler == anzahl)
                {
                    result = p;
                }

            }
     
            //TODO: rückumwandeln
            var filter = new clsFilter();
            result.copyto(ref filter);
            return filter;
        }

        /** speichert einen Filter mit Namen 'name' in der Datenbank*/

        public bool storeFilterToDB(clsFilter ptFilter, Boolean selfilter)
        {

            try
            {
                var filterDB = new clsFilterDB();
                filterDB.copyfrom(ref ptFilter, selfilter);
                filterDB.Id = ptFilter.name;

                bool alreadyExisting = CollectionFilter.AsQueryable<clsFilterDB>()
                        .Any(p => p.name == ptFilter.name);

                if (alreadyExisting)
                {
                    var flt = Builders<clsFilterDB>.Filter.Eq("name", ptFilter.name);
                    var rResult = CollectionFilter.ReplaceOne(flt, filterDB);
                    if (rResult.ModifiedCount > 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    CollectionFilter.InsertOne(filterDB);
                    return true;
                }
                
            }
            catch (Exception)
            {

                return false;
            }

            // alt 2.x
            //var filterDB = new clsFilterDB();
            //filterDB.copyfrom( ref ptFilter,  selfilter);
            //filterDB.Id = ptFilter.name;
            //return !CollectionFilter.Save(filterDB).HasLastErrorMessage;
        }
        /** löscht einen bestimmten Filter aus der Datenbank */

        public bool removeFilterFromDB(clsFilter filter)
        {

            try
            {
                var dResult = CollectionFilter.DeleteOne<clsFilterDB>(p => (p.name == filter.name));
                if (dResult.DeletedCount > 0)
                {
                    return true;
                }
                else
                {
                    return false;
                }
                
            }
            catch (Exception)
            {

                return false;
            }
            
            // alt 2.x 
            //var query = Query<clsFilterDB>
            //    .Where(e => (e.name == filter.name));

            //return !CollectionFilter.Remove(query).HasLastErrorMessage;
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
