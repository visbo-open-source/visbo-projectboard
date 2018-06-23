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
using MongoDB.Driver.GridFS;
using System.IO;

namespace MongoDbAccess
{
    /// <summary>
    /// request , der an eine Datenbank gestellt wird
    /// </summary>
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
        /// <summary>
        /// Client ist der Datenbank Access Client
        /// </summary>
        protected  IMongoClient Client;
        /// <summary>
        /// 
        /// </summary>
        protected  IMongoDatabase Database;
        /// <summary>
        /// 
        /// </summary>
        protected MongoServer Server;
        /// <summary>
        /// die Collection, wo alle Projekte enthalten sind 
        /// </summary>
        protected IMongoCollection<clsProjektDB> CollectionProjects;
        /// <summary>
        /// die Rollen Definitionen
        /// </summary>
        protected IMongoCollection<clsRollenDefinitionDB> CollectionRoles;
        /// <summary>
        /// die Kostenart Definitionen
        /// </summary>
        protected IMongoCollection<clsKostenartDefinitionDB> CollectionCosts;
        /// <summary>
        /// vermerkt, welche Projekte für die laufenden Session bzw permanent von wem geschützt sind 
        /// </summary>
        protected IMongoCollection<clsWriteProtectionItemDB> CollectionWriteProtections;
        /// <summary>
        /// nimmt die gelöschten Projekte auf; erst wenn Sie hier gelöscht werden, sind sie komplett veroren ..
        /// </summary>
        protected IMongoCollection<clsProjektDB> CollectionTrashProjects;
        /// <summary>
        /// nimmt die Defintion der Portfolios auf 
        /// </summary>
        protected IMongoCollection<clsConstellationDB> CollectionConstellations;
        /// <summary>
        /// nimmt die glöschten Portfolio definitionen auf 
        /// </summary>
        protected IMongoCollection<clsConstellationDB> CollectionTrashConstellations; 
        /// <summary>
        /// enthält die Projekt-Abhängigkeiten
        /// </summary>
        protected IMongoCollection<clsDependenciesOfPDB> CollectionDependencies;
        /// <summary>
        /// die gespeicherten Selektion-Filter
        /// </summary>
        protected IMongoCollection<clsFilterDB> CollectionFilter;
        /// <summary>
        /// wird für das Speichern von Dokumenten benötigt 
        /// </summary>
        protected IGridFSBucket BucketDocuments;
        private String User;
        
        /// <summary>
        /// Verbindung mit der Datenbank aufbauen (mit Angabe von Username und Passwort
        /// </summary>
        /// <param name="databaseURL"></param>
        /// <param name="databaseName"></param>
        /// <param name="username"></param>
        /// <param name="dbPasswort"></param>
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

                var connectionString = "";
                
                if (Module1.awinSettings.DBWithSSL)
                {
                     connectionString = "mongodb://" + username + ":" + dbPasswort + "@" + databaseURL + "/" + databaseName + "?ssl=true";
                }
                else
                {
                     connectionString = "mongodb://" + username + ":" + dbPasswort + "@" + databaseURL + "/" + databaseName;
                }
                
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
            CollectionRoles = Database.GetCollection<clsRollenDefinitionDB>("roledefinitions");
            CollectionCosts = Database.GetCollection<clsKostenartDefinitionDB>("costdefinitions");
            CollectionWriteProtections = Database.GetCollection<clsWriteProtectionItemDB>("writeProtections");
            CollectionConstellations = Database.GetCollection<clsConstellationDB>("constellations");
            CollectionTrashConstellations = Database.GetCollection<clsConstellationDB>("trashconstellations");
            CollectionDependencies = Database.GetCollection<clsDependenciesOfPDB>("dependencies");
            CollectionFilter = Database.GetCollection<clsFilterDB>("filters");
            BucketDocuments = new GridFSBucket(Database, new GridFSBucketOptions{BucketName = "documents"});
            User = username;

        }
        
        ////public bool loginSucessful(string databaseName, string username, string dbPasswort)
        ////{
        ////    try
        ////    {
        ////        var credential = MongoCredential.CreateCredential(databaseName, username, dbPasswort);

        ////        // test tk für Authentification
        ////        try
        ////        {
        ////            //var credential = MongoCredential.CreateMongoCRCredential(databaseName, "tk", "test");
        ////            var settings = new MongoClientSettings
        ////            {
        ////                Credentials = new[] { credential }
        ////            };

        ////            var mongoClient = new MongoClient(settings);
                    
        ////        }
        ////        catch (Exception)
        ////        {

        ////            throw;
        ////        }

        ////        int berta=0;
        ////        int a = berta;
                             
        ////        return true;
        ////    }
        ////    catch
        ////    {
        ////        return false;
        ////    }
        ////}
        /// <summary>
        /// wichtige Indices für CollectionProjects und CollectionWriteProtections setzen
        /// </summary>
        /// <returns></returns>
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

                var keys2 = Builders<clsWriteProtectionItemDB>.IndexKeys.Ascending("pName").Ascending("vName").Ascending("type");
                var options = new CreateIndexOptions() { Unique = true };
                ergebnis = CollectionWriteProtections.Indexes.CreateOne(keys2, options);
             
                return true;
            }
            catch
            {
                return false;
            }
           
        }

        /// <summary>
        /// Abfrage, ob die Collection 'name' Inhalte hat
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
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
                case "roledefinitions":
                    result = CollectionRoles.AsQueryable<clsRollenDefinitionDB>().Count();
                    break;
                case "costdefinitions":
                    result = CollectionCosts.AsQueryable<clsKostenartDefinitionDB>().Count();
                    break;
                case "writeProtections":
                    result = CollectionWriteProtections.AsQueryable<clsWriteProtectionItemDB>().Count();
                    break;
                default:
                    result = 0;
                    break;
            }
            
            return result == 0; 
        }

      /// <summary>
      /// hier werden einmalig alle in der Datenbank vorhandenen Projekte/Variante in die CollectionWriteProtections eingetragen
      /// </summary>
      /// <param name="user"></param>
      /// <returns></returns>
      public int initWriteProtectionsOnce(string user)
        {
            int i = 0;
            int result = 0;

            try
            {
                if (collectionEmpty("writeProtections"))
                {
                   clsWriteProtectionItem wpItem = null;
                   clsWriteProtectionItemDB wpItemDB = new clsWriteProtectionItemDB();

                   var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                        .Where(c => c.name != null)
                        .Select(c => c.name)
                        .Distinct()
                        .ToList();

                   foreach (string dbKey in prequery )
                   {
                         string pname = Projekte.getPnameFromKey(dbKey);
                         string vname = Projekte.getVariantnameFromKey(dbKey);

                         wpItem = new clsWriteProtectionItem(Projekte.calcProjektKey(pname, vname), 0, user, false, false);
                         wpItemDB = new clsWriteProtectionItemDB();
                         wpItemDB.copyFrom(wpItem);
                         CollectionWriteProtections.InsertOne(wpItemDB);
                         i = i + 1;
                   }
                } 
                
                result = i;
                      
            }
            catch
            {
                return result;
            }

            return result;
        }


        /// <summary>
        /// prüft ob der Projektname schon vorhanden ist (ggf. inkl. VariantName)
        /// falls Variantname null ist oder leerer String wird nur der Projektname überprüft.
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="variantname"></param>
        /// <param name="storedAtorBefore"></param>
        /// <returns></returns>
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

    
        /// <summary>
        /// holt die erste beauftragte Version des Projects 
        /// immer mit Variant-Name = ""
        /// </summary>
        /// <param name="projectname"></param>
        /// <returns></returns>
        public clsProjekt retrieveFirstContractedPFromDB(string projectname)
        {
            var result = new clsProjektDB();
            string searchstr = Projekte.calcProjektKeyDB(projectname, "");

            
            var builder = Builders<clsProjektDB>.Filter;

            var filter = builder.Eq("name", searchstr) & builder.Eq("status", "beauftragt");
            // das folgende könnte auch gemacht werden 
            // var filter = builder.Eq(c => c.name, searchstr) & builder.Lte(c => c.timestamp, storedAtOrBefore);

            var sort = Builders<clsProjektDB>.Sort.Descending("timestamp");

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
                int a = projekt.dauerInDays;
                return projekt;
            }
            
        }

        /// <summary>
        /// liest ein bestimmtes Projekt aus der DB (ggf. inkl. VariantName), das zum angegebenen Zeitpunkt das aktuelle war
        /// falls Variantname null ist oder leerer String wird nur der Projektname überprüft.
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="variantname"></param>
        /// <param name="storedAtOrBefore"></param>
        /// <returns></returns>
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
            // das folgende könnte auch gemacht werden 
            // var filter = builder.Eq(c => c.name, searchstr) & builder.Lte(c => c.timestamp, storedAtOrBefore);

            

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
                int a = projekt.dauerInDays;
                return projekt;
            }
            
        }



        /// <summary>
        /// liest die angegebene Rollen Definition aus der Datenbank
        /// </summary>
        /// <param name="roleID"></param>
        /// <param name="storedAtOrBefore"></param>
        /// <returns></returns>
        public clsRollenDefinition retrieveOneRoleFromDB(int roleID,  DateTime storedAtOrBefore)
        {
            var result = new clsRollenDefinitionDB();
            
            if (storedAtOrBefore == null)
            {

                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime();
            }
            else
            {
                
                storedAtOrBefore = storedAtOrBefore.ToUniversalTime();
            }

            
            var builder = Builders<clsRollenDefinitionDB>.Filter;

            var filter = builder.Eq("uid", roleID) & builder.Lte("timestamp", storedAtOrBefore);

            var sort = Builders<clsRollenDefinitionDB>.Sort.Ascending("timestamp");

            try
            {
                result = CollectionRoles.Find(filter).Sort(sort).ToList().Last();
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
                var currentrole = new clsRollenDefinition();
                result.copyTo(ref currentrole);
                return currentrole;
            }

        }
        /// <summary>
        /// liest die Rollendefinitionen aus der Datenbank 
        /// </summary>
        /// <param name="storedAtOrBefore"></param>
        /// <returns></returns>
        public clsRollen retrieveRolesFromDB(DateTime storedAtOrBefore)
        {
            var result = new clsRollen();

            if (storedAtOrBefore == null)
            {

                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime();
            }
            else
            {

                storedAtOrBefore = storedAtOrBefore.ToUniversalTime();
            }

            var prequery = CollectionRoles.AsQueryable<clsRollenDefinitionDB>()
                            .Where(c => c.timestamp <= storedAtOrBefore)
                            .Select(c => c.uid)
                            .Distinct()
                            .ToList();

            foreach (int tmpUid in prequery)
            {

                clsRollenDefinition tmpRole = retrieveOneRoleFromDB(tmpUid, storedAtOrBefore);
                if (!result.get_containsUid(tmpRole.UID))
                {
                    result.Add(tmpRole);
                }

                
            }

            return result;
        }

        /// <summary>
        /// liest die Kostenartdefinitionen aus der Datenbank 
        /// </summary>
        /// <param name="storedAtOrBefore"></param>
        /// <returns></returns>
        public clsKostenarten retrieveCostsFromDB(DateTime storedAtOrBefore)
        {
            var result = new clsKostenarten();

            if (storedAtOrBefore == null)
            {

                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime();
            }
            else
            {

                storedAtOrBefore = storedAtOrBefore.ToUniversalTime();
            }

            var prequery = CollectionCosts.AsQueryable<clsKostenartDefinitionDB>()
                            .Where(c => c.timestamp <= storedAtOrBefore)
                            .Select(c => c.uid)
                            .Distinct()
                            .ToList();

            foreach (int tmpUid in prequery)
            {

                clsKostenartDefinition tmpCost = retrieveOneCostFromDB(tmpUid, storedAtOrBefore);
                if (!result.get_containsUid(tmpCost.UID))
                {
                    result.Add(tmpCost);
                }


            }

            return result;
        }

        /// <summary>
        /// liest die angegebene Kostenart aus der Datenbank 
        /// </summary>
        /// <param name="costID"></param>
        /// <param name="storedAtOrBefore"></param>
        /// <returns></returns>
        public clsKostenartDefinition retrieveOneCostFromDB(int costID, DateTime storedAtOrBefore)
        {
            var result = new clsKostenartDefinitionDB();

            if (storedAtOrBefore == null)
            {

                storedAtOrBefore = DateTime.Now.AddDays(1).ToUniversalTime();
            }
            else
            {

                storedAtOrBefore = storedAtOrBefore.ToUniversalTime();
            }


            var builder = Builders<clsKostenartDefinitionDB>.Filter;

            var filter = builder.Eq("uid", costID) & builder.Lte("timestamp", storedAtOrBefore);

            var sort = Builders<clsKostenartDefinitionDB>.Sort.Ascending("timestamp");

            try
            {
                result = CollectionCosts.Find(filter).Sort(sort).ToList().Last();
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
                var currentcost = new clsKostenartDefinition();
                result.copyTo(ref currentcost);
                return currentcost;
            }

        }


        /// <summary>
        /// holt von allen Projekt-Varianten in AlleProjekte die Write-Protections
        /// </summary>
        /// <param name="AlleProjekte"></param>
        /// <returns></returns>
        public SortedList<string, clsWriteProtectionItem> retrieveWriteProtectionsFromDB(clsProjekteAlle AlleProjekte)
        {
            // holt von allen Projekt-Varianten in AlleProjekte die Write-Protections

            var result = new SortedList<string, clsWriteProtectionItem>();
            var writeProtectDB = CollectionWriteProtections.AsQueryable<clsWriteProtectionItemDB>().Select(cDB => cDB);

            foreach (clsWriteProtectionItemDB cDB in writeProtectDB)
            {
                string pvName = Projekte.calcProjektKey(cDB.pName, cDB.vName);
                if (AlleProjekte.get_Containskey(pvName))
                {
                    var wpi = new clsWriteProtectionItem();
                    cDB.copyTo(ref wpi);
                    result.Add(wpi.pvName, wpi);
                }
            }
            
      
            return result;
        }

        /// <summary>
        /// liefert für den pName und vName das clsWriteProtectiomItem zurück
        /// wenn es das nch nicht gibt, dann Null 
        /// </summary>
        /// <param name="pName"></param>
        /// <param name="vName"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public clsWriteProtectionItem getWriteProtection(string pName, string vName, int type = 0)
        {
          
            clsWriteProtectionItemDB wpItemDB = null;
            clsWriteProtectionItem wpItem = new clsWriteProtectionItem();
              
            try
            {
               
                var filter = Builders<clsWriteProtectionItemDB>.Filter.Eq("pName", pName) &
                             Builders<clsWriteProtectionItemDB>.Filter.Eq("vName", vName) &
                             Builders<clsWriteProtectionItemDB>.Filter.Eq("type", type);
                wpItemDB = CollectionWriteProtections.Find(filter).ToList().Last();
                //var fresult = CollectionWriteProtections.Find(filter).ToList();
                
                wpItemDB.copyTo(ref wpItem);

                return wpItem;


            }
            catch (Exception)
            {

                return null;

            }
           
        }

        /// <summary>
        /// setzt für das entsprechende Item das Flag, dass es geschützt ist 
        /// gibt true zurück, wenn die Aktion erfolgreich war, false andernfalls
        /// </summary>
        /// <param name="wpItem"></param>
        /// <returns></returns>
        public bool setWriteProtection(clsWriteProtectionItem wpItem)
        {

            clsWriteProtectionItemDB wpItemDB = new clsWriteProtectionItemDB();

            try
            {

                //string[] separator = new string[] {"#"};
                //string[] tmpstr = wpItem.pvName.Split(separator,StringSplitOptions.None);
                //string searchstr = Projekte.calcProjektKeyDB(tmpstr[0], tmpstr[1]);

                string pName = Projekte.getPnameFromKey(wpItem.pvName);
                string vName = Projekte.getVariantnameFromKey(wpItem.pvName);

                string searchstr = Projekte.calcProjektKeyDB(pName,vName);

                bool projAlreadyExisting = CollectionProjects.AsQueryable<clsProjektDB>()
                         .Any(p => p.name == searchstr);

                if (projAlreadyExisting)
                {
                    // Projekt ist in der DB in CollectionProjects enthalten
                    // Schutz kann evt. durchgeführt werden

           

                    var filter = Builders<clsWriteProtectionItemDB>.Filter.Eq("pName", pName) &
                                 Builders<clsWriteProtectionItemDB>.Filter.Eq("vName", vName) &
                                 Builders<clsWriteProtectionItemDB>.Filter.Eq("type", wpItem.type);
                    //var sort = Builders<clsWriteProtectionItemDB>.Sort.Ascending("pName");


                    // jetzt soll ein Update / Insert gemacht werden; 
                    // es muss aber vorher sichergestellt sein, dass das Element verändert werden darf 
                    // gesucht werden muss das Element mit pvName=pvname und kennung = kennung 
                    // geschützt werden darf nur, wenn isProtected = false oder (isProtected = true und gleicher User) 
                    // Schutz aufheben nur, wenn isProtected = true und user = <user> oder user=<admin>


                    bool alreadyExisting = CollectionWriteProtections.AsQueryable<clsWriteProtectionItemDB>()
                                 .Any(wp => wp.pName == pName && wp.vName == vName && wp.type == wpItem.type);

               
                    if (alreadyExisting)
                    {

                        wpItemDB = CollectionWriteProtections.Find(filter).ToList().Last();
                        //var fresult = CollectionWriteProtections.Find(filter).ToList();

                        switch (wpItemDB.isProtected)
                        {
                            case true:

                                if (wpItemDB.userName == wpItem.userName)
                                {
                                    wpItemDB.copyFrom(wpItem);
                                    var r1Result = CollectionWriteProtections.ReplaceOne(filter, wpItemDB);
                                    return r1Result.IsAcknowledged;

                                }
                                else
                                {
                                    return false;
                                };

                            case false:

                                wpItemDB.copyFrom(wpItem);
                                var r2Result = CollectionWriteProtections.ReplaceOne(filter, wpItemDB);
                                return r2Result.IsAcknowledged;


                            default:

                                return false;

                        }
                    }
                    else
                    {
                        wpItemDB.copyFrom(wpItem);
                        CollectionWriteProtections.InsertOne(wpItemDB);
                        return true;
                    }

                }
                else
                {
                    //   Es existiert dieses Projekt/Variante noch gar nicht in der Datenbank in CollectionProjects 
                    //   kann also auch nicht geschützt werden 
                    return false;
                }

            }
            catch (Exception)
            {

                wpItemDB.copyFrom(wpItem);
                CollectionWriteProtections.InsertOne(wpItemDB);
                return false;
                
            }
        }
        /// <summary>
        /// überprüft, ob der User userName für das Projekt pvname vom Typ type 
        /// die Erlaubnis hat etwas zu verändern
        /// </summary>
        /// <param name="pName"></param>
        /// <param name="vName"></param>
        /// <param name="userName"></param>
        /// <param name="type"></param>
        /// <returns>true -  es darf geändert werden
        ///          false - es darf nicht geändert werden</returns>
        public bool checkChgPermission(string pName, string vName, string userName, int type = 0)
        {
            try
            {
                clsWriteProtectionItemDB wpItemDB = new clsWriteProtectionItemDB();

                var filter = Builders<clsWriteProtectionItemDB>.Filter.Eq("pName", pName) &
                             Builders<clsWriteProtectionItemDB>.Filter.Eq("vName", vName) &
                             Builders<clsWriteProtectionItemDB>.Filter.Eq("type", type);
                //var sort = Builders<clsWriteProtectionItemDB>.Sort.Ascending("pvName");

                bool alreadyExisting = CollectionWriteProtections.AsQueryable<clsWriteProtectionItemDB>()
                               .Any(wp => wp.pName == pName && wp.vName == vName && wp.type == type);

                if (alreadyExisting)
                {

                    wpItemDB = CollectionWriteProtections.Find(filter).ToList().Last();
                    //var fresult = CollectionWriteProtections.Find(filter).ToList();
                    if (wpItemDB.isProtected)
                    {
                        return (wpItemDB.userName == userName);   
                    }
                    else
                    {
                        return true;
                    }
                 
                }
                else
                {
                    return true;
                }
            }

            catch (Exception)
            {

                return false;

            }
        }
  

        /// <summary>
        /// löst von allen Projekt-Varianten des Users user die nonpermanent writeProtections
        /// </summary>
        /// <param name="user"></param>
        /// <returns></returns>
        public bool cancelWriteProtections(string user)
        {  
            // löst von allen Projekt-Varianten des Users user die nonpermanent writeProtections

            var filter = Builders<clsWriteProtectionItemDB>.Filter.Eq("userName", user) &
                         Builders<clsWriteProtectionItemDB>.Filter.Eq("permanent", false) &
                         Builders<clsWriteProtectionItemDB>.Filter.Eq("isProtected", true);

            var updatedef = Builders<clsWriteProtectionItemDB>.Update.Set("isProtected", false).Set("lastDateReleased", DateTime.UtcNow);
           
            var uresult = CollectionWriteProtections.UpdateMany(filter,updatedef);
            return uresult.IsAcknowledged;
        }


        /// <summary>
        /// setzt für alle Projekt-Varianten des Users user die temporär die writeProtections
        /// </summary>
        /// <param name="pName"></param>
        /// <param name="user"></param>
        /// <param name="set"></param>
        /// <returns></returns>
        public bool protectAllVariants(string pName, string user, bool set = true)
        {
            // löst von allen Projekt-Varianten des Users user die nonpermanent writeProtections

            var filter = Builders<clsWriteProtectionItemDB>.Filter.Eq("pName", pName) &
                         Builders<clsWriteProtectionItemDB>.Filter.Eq("type", 0) &
                         Builders<clsWriteProtectionItemDB>.Filter.Eq("userName", user);

            var updatedef = Builders<clsWriteProtectionItemDB>.Update.Set("isProtected", set).Set("lastDateSet", DateTime.UtcNow);

            var uresult = CollectionWriteProtections.UpdateMany(filter, updatedef);
            return (uresult.IsAcknowledged && uresult.ModifiedCount > 0);
        }
      
        /// <summary>
        ///  prüft die Verfügbarkeit der MongoDB
        /// </summary>
        /// <returns></returns>
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


        /// <summary>
        /// speichert ein einzelnes Projekt in der Datenbank
        /// Zeitstempel wird aus den Projekt-Infos genommen
        /// </summary>
        /// <param name="projekt"></param>
        /// <param name="userName"></param>
        /// <returns></returns>
        public bool storeProjectToDB(clsProjekt projekt, string userName)
        {
            try
            {
           
                if (checkChgPermission(projekt.name, projekt.variantName, userName))
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

                    // Projekt in der CollectionWriteProtections anlegen
                    
                    var filter = Builders<clsWriteProtectionItemDB>.Filter.Eq("pName", projekt.name) &
                                 Builders<clsWriteProtectionItemDB>.Filter.Eq("vName", projekt.variantName) &
                                 Builders<clsWriteProtectionItemDB>.Filter.Eq("type",0);
                    //var sort = Builders<clsWriteProtectionItemDB>.Sort.Ascending("pName");

                    bool alreadyExisting = CollectionWriteProtections.AsQueryable<clsWriteProtectionItemDB>()
                                   .Any(wp => wp.pName == projekt.name && wp.vName == projekt.variantName && wp.type == 0);


   
                    if (!alreadyExisting)
                    {
                        string pvName = Projekte.calcProjektKey(projekt);
                        clsWriteProtectionItem wpItem = new clsWriteProtectionItem(pvName, 0, userName, false, false);
                        clsWriteProtectionItemDB wpItemDB = new clsWriteProtectionItemDB();
                        wpItemDB.copyFrom(wpItem);
                        CollectionWriteProtections.InsertOne(wpItemDB);
                        
                        return true;
                    }
                    else
                    {
                        var updateFilter = Builders<clsWriteProtectionItemDB>.Filter.Eq("pName", projekt.name) &
                                 Builders<clsWriteProtectionItemDB>.Filter.Eq("vName", projekt.variantName) &
                                 Builders<clsWriteProtectionItemDB>.Filter.Eq("type",0) &
                                 Builders<clsWriteProtectionItemDB>.Filter.Eq("userName", userName) &
                                 Builders<clsWriteProtectionItemDB>.Filter.Eq("permanent", false) &
                                 Builders<clsWriteProtectionItemDB>.Filter.Eq("isProtected", true);

                        var updatedef = Builders<clsWriteProtectionItemDB>.Update.Set("isProtected", false).Set("lastDateReleased", DateTime.UtcNow);

                        var uresult = CollectionWriteProtections.UpdateOne(updateFilter, updatedef);
                        return uresult.IsAcknowledged;
                    }


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
        /// <summary>
        /// speichert ein Projekt in der Trash-Datenbank
        /// </summary>
        /// <param name="projektDB"></param>
        /// <returns></returns>
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

        
        /// <summary>
        /// löscht den angegebenen TimeStamp der Projekt-Variante aus der Datenbank 
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="variantName"></param>
        /// <param name="stored"></param>
        /// <param name="userName"></param>
        /// <returns></returns>
        public bool deleteProjectTimestampFromDB(string projectname, string variantName, DateTime stored, string userName)
        {
            try
            {
        
                if (checkChgPermission(projectname, variantName, userName))
                {
                    
                    stored = stored.ToUniversalTime();
                    string searchstr = Projekte.calcProjektKeyDB(projectname, variantName);   /* in der CollectionsProjects in der DB wird der Name des Projektes (wenn variantName = "") am Ende ohne # gespeichert */

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
                else
                {
                    return false;
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
                            .Where(c => c.startDate <= zeitraumEnde && c.endDate >= zeitraumStart && c.timestamp <= storedatOrBefore && c.projectType != 1 && c.projectType != 2)
                            .Select(c => c.name)
                            .Distinct()
                            .ToList();

            // tk 29.5.18
            // wurde eingeführt, weil in Datenbank wo noch kein isUnion Attribut steckt , sonst die leere Liste rauskommt ...
            // 
            if (prequery.Count == 0)
                {
                 prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                            .Where(c => c.startDate <= zeitraumEnde && c.endDate >= zeitraumStart && c.timestamp <= storedatOrBefore )
                            .Select(c => c.name)
                            .Distinct()
                            .ToList();
            }

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

        /// <summary>
        /// liest entweder alle Projekte im angegebenen Zeitraum 
        /// oder aber alle Timestamps der übergebenen Projektvariante im angegeben Zeitfenster
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="variantName"></param>
        /// <param name="zeitraumStart"></param>
        /// <param name="zeitraumEnde"></param>
        /// <param name="storedEarliest"></param>
        /// <param name="storedLatest"></param>
        /// <param name="onlyLatest"></param>
        /// <returns></returns>
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
                        // nötig, um die Dauer in Monaten zu aktualisieren 
                        int a = projekt.dauerInDays; 
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
                    int a = projekt.dauerInDays;
                    // wichtig: in p steht das timestamp in utc format, in projekt in localtime
                    string schluessel = projekt.timeStamp.ToString();
                    //result.Add(projekt.Id, projekt);
                    result.Add(schluessel, projekt);
                }
            }

            return result;
        }

        /// <summary>
        /// liefert alle Varianten Namen eines bestimmten Projektes zurück 
        /// </summary>
        /// <param name="projectName"></param>
        /// <returns></returns>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pNR"></param>
        /// <returns></returns>
        public Collection retrieveProjectNamesByPNRFromDB(string pNR)
        {
            var result = new Collection();

            
            //gegeben: Projektname, Backupzeitraum (also storedEarliest, storedLatest)
            //var projects = from e in CollectionProjects.AsQueryable<clsProjektDB>()
            //               where e.name.Contains(searchstr)
            //               select e.variantName
            //               .Distinct();


            var prequery = CollectionProjects.AsQueryable<clsProjektDB>()
                            .Where(c => c.kundenNummer == pNR)
                            .OrderBy(c => c.name)
                            .Select(c => c.name)
                            .ToList()
                            .Distinct();

            foreach (string vName in prequery)
            {
                result.Add(vName);
            }

            return result;
        }

        /// <summary>
        /// bringt alle in der Datenbank vorkommenden TimeStamps zurück , in absteigender Sortierung
        /// </summary>
        /// <returns></returns>
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


        
        /// <summary>
        /// bringt für die angegebene Projekt-Variante alle Zeitstempel in absteigender Sortierung zurück 
        /// </summary>
        /// <param name="pvName"></param>
        /// <returns></returns>
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

        /// <summary>
        /// gibt die Projekthistorie innerhalb eines gegebenen Zeitraums zu einem gegebenen Projekt+Varianten-Namen zurück
        /// </summary>
        /// <param name="projectname"></param>
        /// <param name="variantName"></param>
        /// <param name="storedEarliest"></param>
        /// <param name="storedLatest"></param>
        /// <returns></returns>
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

                    int a = projekt.dauerInDays;

                    DateTime schluessel = projekt.timeStamp;
                    result.Add(schluessel, projekt);
                }
            

            return result;
        }

        /// <summary>
        /// speichert eine Rolle in der Datenbank; 
        /// wenn insertNewDate = true: speichere eine neue Timestamp-Instanz 
        /// andernfalls wird die Rolle Replaced 
        /// </summary>
        /// <param name="role"></param>
        /// <param name="insertNewDate"></param>
        /// <param name="ts"></param>
        /// <returns></returns>
        public bool storeRoleDefinitionToDB(clsRollenDefinition role, bool insertNewDate, DateTime ts)
        {
            bool tmpResult = true;
            try
            {
                var roleDB = new clsRollenDefinitionDB();
                roleDB.copyFrom(role);

                if (insertNewDate)
                {
                    roleDB.timestamp = ts;
                    CollectionRoles.InsertOne(roleDB);
                }
                else
                {

                    var filter = Builders<clsRollenDefinitionDB>.Filter.Eq("uid", role.UID);
                    var sort = Builders<clsRollenDefinitionDB>.Sort.Ascending("timestamp");

                    try
                    {

                        if (CollectionRoles == null) 
                        {
                            CollectionRoles.InsertOne(roleDB);
                        }
                        else
                        {
                            try
                            {
                                clsRollenDefinitionDB tmpRole = CollectionRoles.Find(filter).Sort(sort).ToList().Last();
                                if (tmpRole == null)
                                {
                                    // existiert noch nicht 
                                    CollectionRoles.InsertOne(roleDB);
                                }
                                else
                                {
                                    // existiert bereits , soll also ersetzt werden , aber mit dem bisherigen TimeStamp 
                                    // und das nur, wenn es nicht identisch ist mit der bereits existierenden 
                                    if (!tmpRole.get_isIdenticalTo(roleDB))
                                    {
                                        roleDB.timestamp = tmpRole.timestamp;

                                        var builder = Builders<clsRollenDefinitionDB>.Filter;
                                        filter = builder.Eq("uid", role.UID) & builder.Eq("timestamp", tmpRole.timestamp);

                                        var rResult = CollectionRoles.ReplaceOne(filter, roleDB);
                                        tmpResult = rResult.IsAcknowledged;

                                    }
                                    else
                                    {
                                        // nichts tun
                                    }

                                }
                            }
                            catch (Exception)
                            {
                                
                                 // es gibt noch überhaupt keine Elemente in der Collection 
                                CollectionRoles.InsertOne(roleDB);
                            }



                        }



                    }
                    catch (Exception)
                    {

                        tmpResult = false;
                    }
                 }       
                                
                                
            }
            catch (Exception)
            {
                tmpResult =  false;
            }

            return tmpResult;
        }

        /// <summary>
        /// speichert eine Kostenart in der Datenbank; 
        /// wenn insertNewDate = true: speichere eine neue Timestamp-Instanz 
        /// andernfalls wird die Kostenart Replaced, sofern sie sich geändert hat  
        /// </summary>
        /// <param name="cost"></param>
        /// <param name="insertNewDate"></param>
        /// <param name="ts"></param>
        /// <returns></returns>
        public bool storeCostDefinitionToDB(clsKostenartDefinition cost, bool insertNewDate, DateTime ts)
        {
            bool tmpResult = true;
            try
            {
                var costDefDB = new clsKostenartDefinitionDB();
                costDefDB.copyFrom(cost);

                if (insertNewDate)
                {
                    costDefDB.timestamp = ts;
                    CollectionCosts.InsertOne(costDefDB);
                }
                else
                {

                    var filter = Builders<clsKostenartDefinitionDB>.Filter.Eq("uid", cost.UID);
                    var sort = Builders<clsKostenartDefinitionDB>.Sort.Ascending("timestamp");

                    try
                    {

                        if (CollectionCosts == null)
                        {
                            // existiert noch nicht 
                            CollectionCosts.InsertOne(costDefDB);
                        }
                        else
                        {

                            try
                            {
                                clsKostenartDefinitionDB tmpCost = CollectionCosts.Find(filter).Sort(sort).ToList().Last();
                                if (tmpCost == null)
                                {
                                    // existiert noch nicht 
                                    CollectionCosts.InsertOne(costDefDB);
                                }
                                else
                                {
                                    // existiert bereits , soll also ersetzt werden , dann mit dem bisherigen TimeStamp 
                                    // aber nur, wenn es nicht identisch ist mit der bereits existierenden 
                                    if (!tmpCost.get_isIdenticalTo(costDefDB))
                                    {
                                        costDefDB.timestamp = tmpCost.timestamp;

                                        var builder = Builders<clsKostenartDefinitionDB>.Filter;
                                        filter = builder.Eq("uid", cost.UID) & builder.Eq("timestamp", tmpCost.timestamp);

                                        var rResult = CollectionCosts.ReplaceOne(filter, costDefDB);
                                        tmpResult = rResult.IsAcknowledged;

                                    }
                                    else
                                    {
                                        // nichts tun
                                    }

                                }
                            }
                            catch (Exception)
                            {
                                // existiert noch nicht 
                                CollectionCosts.InsertOne(costDefDB);
                            }

                        }
                                                
                    }
                    catch (Exception)
                    {

                        tmpResult = false;
                    }
                }


            }
            catch (Exception)
            {
                tmpResult = false;
            }

            return tmpResult;
        }


        /// <summary>
        /// Speichert ein Multiprojekt-Szenario in der Datenbank
        /// </summary>
        /// <param name="c"> - Constellation</param>
        /// <returns></returns>
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

        /// <summary>
        /// Speichert ein Multiprojekt-Portfolio in der Trash-Datenbank
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Löschen des Portfolios  aus der Datenbank
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
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

      
        /// <summary>
        /// benennt alle Projekte mit Namen oldName um
        /// aber nur, wenn der neue Name nicht schon in der Datenbank existiert 
        /// </summary>
        /// <param name="oldName"></param>
        /// <param name="newName"></param>
        /// <param name="userName"></param>
        /// <returns></returns>
        public bool renameProjectsInDB(string oldName, String newName, string userName)
        {
            if (projectNameAlreadyExists(newName, "", DateTime.Now))
            {
                return false;
            }
            else
            {

                try
                {
                    //string oldpvName;
                    //string newpvName;
                    bool chkOk = true;
                    
                    // hier wird überprüft, ob das Projekt selbst
                    // und auch keine der Varianten von einem anderen User schreibgeschützt ist

                    chkOk = checkChgPermission(oldName, "", userName);
                                 
                    Collection listOfVariants = retrieveVariantNamesFromDB(oldName);

                    foreach (string vName in listOfVariants)
                    {
                        if (!chkOk)
                        { break; }
                       
                        chkOk = chkOk && checkChgPermission(oldName, vName, userName);
                       
                    }                  
            

                    // Projekt und seine Varianten können umbenannt werden

                    if (chkOk)
                    {
                        if (protectAllVariants(oldName, userName))
                        {


                        string oldFullName;
                        string newFullName;
                        bool ok = false;

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
                            //Collection listOfVariants = retrieveVariantNamesFromDB(oldName);


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
                               
                                if (a > 0)
                                {
                                    // ur: 28.10.2017: hier wird nun die Routine zum SPeichern der Constellation in der DB verwendet, da Wartungsaufwand geringer
                                    var ergebnis = storeConstellationToDB(c);
                                    ok = ok & ergebnis;
                                 

                                    //////clsConstellationDB chgcDB = new clsConstellationDB();
                                    //////chgcDB.copyfrom(c);
                                    //////// mit Id=null kann kein Replace gemacht werden  
                                    //////chgcDB.Id = cDB.constellationName;

                                    //////var filter = Builders<clsConstellationDB>.Filter.Eq("constellationName", chgcDB.constellationName);
                                    //////var rResult = CollectionConstellations.ReplaceOne(filter, chgcDB);
                                    ////////ok = ok & (rResult.ModifiedCount > 0);
                                    //////ok = ok & rResult.IsAcknowledged;

                                    zaehler = zaehler + 1;
                                    gesamt = gesamt + a;
                                }


                            }
                            // Ende Aktualisierung Constellations ...


                            // dann müssen noch die Dependencies aktualisiert werden ...

                            // CollectionWriteProtection muss aktualisiert werden
                            try
                            {

                                if (ok)
                                {
                                    bool alreadyExisting = CollectionWriteProtections.AsQueryable<clsWriteProtectionItemDB>()
                                        .Any(wp => wp.pName == oldName && wp.vName == "" && wp.type == 0);


                                    if (alreadyExisting)
                                    {
                                        // zuerst dieses Projekt mit Varianten aus CollectionWriteProtections löschen
                                        // neu 3.0 
                                        var wpfilter = Builders<clsWriteProtectionItemDB>.Filter.Eq("pName", oldName) &
                                                                    Builders<clsWriteProtectionItemDB>.Filter.Eq("vName", "") &
                                                                    Builders<clsWriteProtectionItemDB>.Filter.Eq("type", 0);
                                        var wpUpdate = Builders<clsWriteProtectionItemDB>.Update.Set("pName", newName);

                                        var result = CollectionWriteProtections.UpdateOne(wpfilter, wpUpdate);
                                        ok = ok & (result.ModifiedCount > 0);
                                        ////var result = CollectionWriteProtections.DeleteOne(wpfilter);
                                        ////ok = ok & result.IsAcknowledged;

                                        foreach (string vName in listOfVariants)
                                        {
                                            //oldFullName = Projekte.calcProjektKey(oldName, vName);
                                            //newFullName = Projekte.calcProjektKey(newName, vName);

                                            // neu 3.0 
                                            var filter = Builders<clsWriteProtectionItemDB>.Filter.Eq("pName", oldName) &
                                                Builders<clsWriteProtectionItemDB>.Filter.Eq("vName", vName) &
                                                Builders<clsWriteProtectionItemDB>.Filter.Eq("type", 0);
                                            var update = Builders<clsWriteProtectionItemDB>.Update.Set("pName", newName);

                                            var vresult = CollectionWriteProtections.UpdateOne(filter, update);
                                            ok = ok & (vresult.ModifiedCount > 0);
                                            //////var vresult = CollectionWriteProtections.DeleteOne(filter);
                                            //////ok = ok & vresult.IsAcknowledged;

                                        }
                                    }
                                    else
                                    {
                                        clsWriteProtectionItemDB wpItemDB = new clsWriteProtectionItemDB();
                                        clsWriteProtectionItem wpItem = new clsWriteProtectionItem(Projekte.calcProjektKey(newName, ""), 0, userName, false, false);
                                        wpItemDB.copyFrom(wpItem);
                                        CollectionWriteProtections.InsertOne(wpItemDB);

                                        foreach (string vName in listOfVariants)
                                        {
                                            // neu 3.0 

                                            wpItem = new clsWriteProtectionItem(Projekte.calcProjektKey(newName, vName), 0, userName, false, false);
                                            wpItemDB = new clsWriteProtectionItemDB();
                                            wpItemDB.copyFrom(wpItem);
                                            CollectionWriteProtections.InsertOne(wpItemDB);

                                        }
                                    }
                                }
                                
                            }
                                     
                        
                            catch (Exception)
                            {
                                ok = false;
                            }

                            if (ok)
                            { return true; }
                            else
                            { return false; }

                        }
                        else
                        { return false; }

                        }   // hier ist if (protectAllVariants) zu Ende
                        else
                        { return false; }

                    }   // hier ist if(chkOK) zu Ende
                    else
                    { return false; }  
              
                }
                catch (Exception)
                {
                    
                    return false;
                }
                
            }
            // return true;
        }

        /// <summary>
        /// Alle Portfolios (Constellations) aus der Datenbank holen 
        /// Das Ergebnis dieser Funktion ist eine Liste (string, clsConstellation) 
        /// </summary>
        /// <returns></returns>
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


        /// <summary>
        /// speichert Projekt-Dependencies in DB 
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
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

        /// <summary>
        /// Alle Abhängigkeiten aus der Datenbank lesen
        /// und als Ergebnis ein Liste von Abhängigkeiten zurückgeben
        /// </summary>
        /// <returns></returns>
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

       
        /// <summary>
        /// liest einen bestimmten Filter aus der DB  
        /// </summary>
        /// <param name="filtername"></param>
        /// <returns></returns>
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

       /// <summary>
        /// speichert einen Filter mit Namen 'name' in der Datenbank
       /// </summary>
       /// <param name="ptFilter"></param>
       /// <param name="selfilter"></param>
       /// <returns></returns>
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
        
        /// <summary>
        /// löscht einen bestimmten Filter aus der Datenbank
        /// </summary>
        /// <param name="filter"></param>
        /// <returns></returns>
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

        /// <summary>
        /// liest alle Filter aus der Datenbank 
        /// </summary>
        /// <param name="selfilter"></param>
        /// <returns></returns>
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


 
        // TODO: add user permission logic to queries. Use this.User for the currently logged in username
        // TODO: create and use mapper classes for FileInfo and File.


        /// <summary>
        /// Uploads a document from the local file system to the database. See also the asynchronous method see also StoreDocumentToDBAsync(string, string, string[], string
        /// </summary>
        /// <param name="filePath">the path where to find the document on the local file system</param>
        /// <param name="userName">the username of the currently logged in user</param>
        /// <param name="entitled">array of usernames of all users who are eligible to download and upload the document</param>
        /// <param name="description">a description of the document given by the user</param>
        /// <remarks>Warning: permission management should be done by built-in database logic if possible instead of an "entitled" parameter</remarks>
        /// <returns>the string representation of the unique objectId that the database has assigned to the uploaded document</returns>
        public String StoreDocumentToDB(string filePath, string userName, string[] entitled, string description)
        {
            var uploadOptions = new GridFSUploadOptions {
                Metadata = new BsonDocument {
                    { "filetype", Path.GetExtension(filePath) },
                    { "entitled", new BsonArray().AddRange(entitled) },
                    { "description", description }
                }
            };
            Stream stream = new FileStream(filePath, FileMode.Open);

            ObjectId objectId = BucketDocuments.UploadFromStream(Path.GetFileName(filePath), stream, uploadOptions);            
            return objectId.ToString();
        }

        /// <summary>
        /// Asynchronous method to upload a document from the local file system to the database. See description of <see cref="StoreDocumentToDB(string, string, string[], string)"/>
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileType"></param>
        /// <param name="userName"></param>
        /// <param name="entitled"></param>
        /// <param name="description"></param>
        /// <returns>an awaitable <code>Task</code> object that, once the upload is finished, will contain the string representation of the unique objectId that the database has assigned to the document</returns>
        public async Task<String> StoreDocumentToDBAsync(string filePath, int fileType, string userName, string[] entitled, string description)
        {
            var uploadOptions = new GridFSUploadOptions {
                Metadata = new BsonDocument {
                    { "filetype", Path.GetExtension(filePath) },
                    { "entitled", new BsonArray().AddRange(entitled) },
                    { "description", description }
                }
            };
            Stream stream = new FileStream(filePath, FileMode.Open);

            ObjectId objectId = await BucketDocuments.UploadFromStreamAsync(Path.GetFileName(filePath), stream, uploadOptions);
            return objectId.ToString();
        }
        
        /// <summary>
        /// Downloads a document from the database to the local file system by specifying its unique id.
        /// </summary>
        /// <param name="id">string representation of the unique objectId of the document that is retrieved from the database</param>
        /// <param name="userName">the username of the currently logged in user</param>
        /// <param name="filePath">the path where to put the document on the local file system</param>
        /// <returns>the file path where the document was put on the local file system</returns>
        public String retrieveDocumentFromDBById(String id, String userName, String filePath)
        {
            Stream stream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite);
            BucketDocuments.DownloadToStream(new ObjectId(id), stream);
            stream.Close();
            return filePath;
        }

        /// <summary>
        /// Asynchronous method to download a document from the database to the local file system. For description see <see cref="retrieveDocumentFromDBById(string, string, string)"/>
        /// </summary>
        /// <param name="id"></param>
        /// <param name="userName"></param>
        /// <param name="filePath"></param>
        /// <returns>an awaitable <code>Task</code> object that, once the download is finished, will contain the file path where the document was put on the local file system</returns>
        public async Task<String> retrieveDocumentFromDBByIdAsync(String id, String userName, String filePath)
        {
            Stream stream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite);
            await BucketDocuments.DownloadToStreamAsync(new ObjectId(id), stream);
            stream.Close();
            return filePath;
        }
        
        /// <summary>
        /// Downloads a document from the database to the local file system by specifying its filename and revision number.
        /// </summary>
        /// <param name="fileName">the filename of the document that is retrieved from the database</param>
        /// <param name="fileRevision">the revision number of the named document that is retrieved from the database. 
        /// Specify 0 for the original version of the document, 1 for the first revision, ..., -1 for the latest revision, -2 for the second latest, etc.</param>
        /// <param name="userName">the username of the currently logged in user</param>
        /// <param name="filePath">the path where to put the document on the local file system</param>
        /// <returns>the file path where the document was put on the local file system</returns>
        public String retrieveDocumentFromDBByName(String fileName, int fileRevision, String userName, String filePath)
        {
            Stream stream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite);
            BucketDocuments.DownloadToStreamByName(fileName, stream, new GridFSDownloadByNameOptions { Revision = fileRevision });
            stream.Close();
            return filePath;
        }

        /// <summary>
        /// Asynchronous method to download a document from the database to the local file system by specifying its filename and revision number. 
        /// See <see cref="retrieveDocumentFromDBByName(string, int, string, string)"/>
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="fileRevision"></param>
        /// <param name="userName"></param>
        /// <param name="filePath"></param>
        /// <returns>an awaitable <code>Task</code> object that, once the download is finished, contains the file path where the document was put on the local file system</returns>
        public async Task<String> retrieveDocumentFromDBByNameAsync(String fileName, int fileRevision, String userName, String filePath)
        {
            Stream stream = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite);
            await BucketDocuments.DownloadToStreamByNameAsync(fileName, stream, new GridFSDownloadByNameOptions { Revision = fileRevision });
            stream.Close();
            return filePath;
        }



        /// <summary>
        /// ersetzt das angegebene Dokument in der Datenbank
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="timeStamp"></param>
        /// <param name="userName"></param>
        /// <param name="entitled"></param>
        /// <returns></returns>
        public String replaceDocument(string filePath, DateAndTime timeStamp, string userName, string[] entitled)
        {
            //TODO: what does replace mean? replace a single version or replace all versions of the document?
            deleteDocumentFromDB(Path.GetFileName(filePath));
            return StoreDocumentToDB(filePath, userName, entitled, "desc");
        }
        
        /// <summary>
        /// löscht das angegebene Dokument aus der Datenbank 
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public String deleteDocumentFromDB(string fileName)
        {
            //delete all documents with given filename
            return "";
        }
        /// <summary>
        /// gibt eine Liste an Version zu dem angegebenen Dokument zurück
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public List<String> FindAllRevisionsOfDocumentInDB(String fileName)
        {
            List<String> fileNames = new List<String>();
            var filter = Builders<GridFSFileInfo>.Filter.Eq(x => x.Filename, fileName);
            using (var cursor = BucketDocuments.Find(filter/*, new GridFSFindOptions*/))
            {
                var ff = cursor.ToList().DefaultIfEmpty();
                // fileInfo either has the matching file information or is null
                foreach (GridFSFileInfo info in ff)
                {
                    fileNames.Add(info.Filename);
                    Console.WriteLine(info.Filename + ": " + info.UploadDateTime);
                }
            }
            return fileNames;
        }
        /// <summary>
        /// holt die jeweils letzten Versionen der Dokumente aus der Datenbank 
        /// Ergebnis ist eine List von Strings
        /// </summary>
        /// <returns></returns>
        public async Task<List<String>> FindLatestRevisionOfAllDocumentsInDBAsync()
        {
            PipelineDefinition<GridFSFileInfo, BsonDocument> pipeline = new BsonDocument[]{
                new BsonDocument { { "$match", new BsonDocument("metadata.entitled", this.User)} },
                new BsonDocument { { "$sort", new BsonDocument("uploadDate", -1)} },
                new BsonDocument { { "$group", new BsonDocument { {"_id", "$filename"}, {"latest", new BsonDocument("$first", "$$ROOT") } } } }
            };
            
            IMongoCollection<GridFSFileInfo> CollectionDocs = Database.GetCollection<GridFSFileInfo>("documents.files");

            var results = await CollectionDocs.Aggregate<BsonDocument>(pipeline).ToListAsync();

            foreach (BsonDocument elem in results)
            {
                GridFSFileInfo info = new GridFSFileInfo((BsonDocument)elem.GetValue("latest"));
                Console.WriteLine(info.Filename + ": " + info.UploadDateTime);
            }
            return null;
        }
        /// <summary>
        /// löscht die Dokumente 
        /// </summary>
        public void clearDocuments()
        {
            BucketDocuments.Drop();
        }


    }
}
