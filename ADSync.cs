using System;
using System.Collections.Generic;
using System.Text;
using System.DirectoryServices;
using System.Collections;
using System.Configuration;


namespace Technovangelist.Captaris.AlchemyCodeSamples.ADSync
{
    /// <summary>
    /// AD Sync - Command Line Utility to synchronise users from Active Directory    
    ///           to an Alchemy Database. Users in the AlchemyUsers group are       
    ///           identified. A folder is created in their name, and a db group     
    ///           is created based on their login. The folder is then assigned      
    ///           to the group. On the server, a Role is created, the db and group  
    ///           is added, the user is added. All that is left to do is to enable  
    ///           Integrated Security.                              
    /// </summary>
    class ADSync
    {
        static string serverName = ConfigurationManager.AppSettings["ServerName"].ToString();
        static string domainName = ConfigurationManager.AppSettings["DomainName"].ToString();
        static string fQDN = ConfigurationManager.AppSettings["FQDN"].ToString();

        static string adminUser = ConfigurationManager.AppSettings["AdminUserName"].ToString();
        static string adminPassword = ConfigurationManager.AppSettings["AdminUserPassword"].ToString();
        static string defaultPassword = ConfigurationManager.AppSettings["DefaultUserPassword"].ToString();
        static string ADGroupName = ConfigurationManager.AppSettings["ADAlchemyGroup"].ToString();

        static string dbPath = "alchemy://" + fQDN + ":3234/";
        static string dbLocalPath = ConfigurationManager.AppSettings["DatabaseParentDirectory"].ToString();

        static Alchemy.Application app = new Alchemy.Application();
        static AuServerApi.ApplicationClass svrApp = new AuServerApi.ApplicationClass();
        static Alchemy.Database DB;

        static void Main(string[] args)
        {
            int FolderID = 0;

            Console.Clear();
            Console.WriteLine("ADSYNC.EXE: Created for the Captaris Partner Conference 2007 in Kuala Lumpur");
            Console.WriteLine("            Developed by Matt Williams, Captaris International Trainer");
            Console.WriteLine("            Not for Production Use");
            Console.WriteLine();
            Console.WriteLine("            Command Line Arguments are [clean] databasename");
            Console.WriteLine();
            
            UserInfo[] ADUsers = GetADUsers(ADGroupName);

            if (args.Length > 0) 
            {
                switch (args[0])
                {
                    case "clean":
                        if (args.Length > 1)
                        {
                            string dbName = args[1].ToString();
                            if (!dbName.EndsWith(".ald"))
                                dbName = dbName + ".ald";

                            dbPath = dbPath + dbName;
                            dbLocalPath = dbLocalPath + dbName;
                            Console.WriteLine(dbPath);
                            try
                            {
                                ConnectDatabase();

                                foreach (UserInfo username in ADUsers)
                                {
                                    Console.WriteLine("Cleaning {0} ({1}) from the database", username.FullName, username.LoginName);
                                    DeleteRole(username);
                                    DeleteGroup(username);
                                    DeleteFolder(username);
                                }

                                UserInfo AdminUser = new UserInfo(adminUser, adminUser);
                                DeleteRole(AdminUser);
                                DeleteGroup(AdminUser);
                                DeleteFolder(AdminUser);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("There was a problem running the application." + ex.ToString());
                            }
                        }
                        else
                            Console.WriteLine("Enter a db name that is under server control after 'clean'");

                        break;

                    default:

                        string dbname = args[0].ToString();
                        if (!dbname.EndsWith(".ald"))
                            dbname = dbname + ".ald";

                        dbPath = dbPath + dbname;
                        dbLocalPath = dbLocalPath + dbname;

                        Console.WriteLine(dbPath);
                        try
                        {
                            ConnectDatabase();

                            Console.WriteLine(String.Format("Found the following users in {0}:", ADGroupName));
                            foreach (UserInfo username in ADUsers)
                            {
                                Console.WriteLine("\t{0}: {1}", username.LoginName, username.FullName);
                                FolderID = CreateFolder(username);
                                CreateGroup(username, FolderID);
                                CreateRole(username);
                            }

                            Alchemy.Builder build = app.NewBuilder(DB);
                            build.Build(null);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("There was a problem running the application." + ex.ToString());
                        }
                        break;
                }
            }
            else
                Console.WriteLine("Enter either a db name under server control or enter 'clean' followed by dbname");
            
        }

        private static void ConnectDatabase()
        {
            try
            {
                svrApp.Connect(serverName, 0, 0, 0);
                DB = app.Databases.Add(dbPath);

                if (DB.HasSecurityGroups)
                {
                    DB.Login(adminUser, adminPassword, "");
                    Console.WriteLine("Logged in");
                }
                else
                {
                    UserInfo AdminUser = new UserInfo(adminUser, adminUser);
                    DB.SecurityGroups.Add(adminUser, adminPassword);
                    DB.Logout();

                    DB.Login(adminUser, adminPassword, "");
                    DB.Refresh();

                    int AdminFolderID = CreateFolder(AdminUser);
                    Alchemy.Item AdminFolder = DB.GetItemByID(AdminFolderID);
                    AdminFolder.SecurityGroups.Add(AdminUser.LoginName, "");
                    CreateRole(AdminUser);
                    DB.Refresh();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Database not found or other error\n" + ex.ToString());
                throw ex;
            }
        }

        /// <summary>
        /// Delete the role from the Server.
        /// </summary>
        /// <param name="username">Name of the role to delete</param>
        private static void DeleteRole(UserInfo username)
        {
            try
            {
                svrApp.Connect(serverName, 0, 0, 0);
                int dbID = svrApp.get_DatabaseId(dbLocalPath);
                svrApp.DeleteRoleDbAcl(username.LoginName, dbID);
                svrApp.DeleteRoleUser(username.LoginName, domainName + "\\" + username.LoginName);
                svrApp.DeleteRole(username.LoginName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not delete the role: " + username);
                throw ex;
            }
        }

        /// <summary>
        /// Create a new role on the Server.
        /// </summary>
        /// <param name="username">Name of the Role to create. Also acts as the user to add to the role</param>
        private static void CreateRole(UserInfo username)
        {
            try
            {
                svrApp.Connect(serverName, 0, 0, 0);
                svrApp.CreateRole(username.LoginName);
                svrApp.AddRoleUser(username.LoginName, domainName + "\\" + username.LoginName);
                int dbID = svrApp.get_DatabaseId(dbLocalPath);
                svrApp.AddRoleDbAcl(username.LoginName, dbID, "");
                svrApp.AddDbAclGroup(username.LoginName, dbID, adminPassword, username.LoginName);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not create the role: " + username);
                throw ex;
            }
        }

        /// <summary>
        /// Create a new database group on the database.
        /// </summary>
        /// <param name="username"></param>
        /// <param name="FolderID">The folder the group should be assigned to</param>
        private static void CreateGroup(UserInfo username, int FolderID)
        {
            Alchemy.Item Folder;
            try
            {
                DB.SecurityGroups.Add(username.LoginName, defaultPassword);
                Folder = DB.GetItemByID(FolderID);
                Folder.SecurityGroups.Add(username.LoginName, "");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not create the group {0} or assign to folder {1}", username, FolderID.ToString());
                throw ex;
            }
        }

        /// <summary>
        /// Delete a database group from the database.
        /// </summary>
        /// <param name="username"></param>
        private static void DeleteGroup(UserInfo username)
        {
            if (DB.HasSecurityGroups)
            {
                try
                {
                    DB.Login(adminUser, adminPassword, "");
                    DB.Refresh();
                    DB.SecurityGroups.Remove(username.LoginName);
                    DB.Refresh();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Could not delete the group " + username);
                }
            }
        }

        /// <summary>
        /// Delete the folder from the database.
        /// </summary>
        /// <param name="username"></param>
        private static void DeleteFolder(UserInfo username)
        {
            try
            {
                Alchemy.Query query = app.NewQuery();
                query.AddFolderQuery("Folder Title", username.FullName, false);
                query.Search(DB);
                if (query.Results.Count > 0)
                {
                    Alchemy.Items auItems = query.Results[1].Items;
                    if (auItems.Count == 1)
                        auItems[1].Delete();
                }
                else
                    Console.WriteLine("No Folder in Alchemy for that user");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not delete the folder for " + username);
                throw ex;
            }
        }

        /// <summary>
        /// Create a folder for the user in the database.
        /// </summary>
        /// <param name="username"></param>
        /// <returns></returns>
        private static int CreateFolder(UserInfo username)
        {

            int itemID = 0;

            try
            {
                Alchemy.Item newItem;
                newItem = DB.CreateFolder(username.FullName, DB.Root, false);
                itemID = newItem.ID;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not create the folder for " + username);
                throw ex;
            }
            return itemID;
        }

        /// <summary>
        /// Populate an array with all of the users who are members of the AlchemyUsers group.
        /// </summary>
        /// <returns></returns>
        /// <param name="GroupName"></param>
        private static UserInfo[] GetADUsers(string GroupName)
        {
            ArrayList users = new ArrayList();
            DirectoryEntry entry = new DirectoryEntry("LDAP://" + domainName);
            entry.RefreshCache();
            DirectorySearcher search = new DirectorySearcher(entry);
            search.Filter = String.Format("(&(objectClass=group)(cn={0}))", GroupName);
            DirectoryEntry subEntry;
            foreach (object objMember in search.FindAll()[0].GetDirectoryEntry().Properties["member"])
            {
                subEntry = new DirectoryEntry("LDAP://" + domainName + "/" + objMember.ToString());
                //string mstring = objMember.ToString();
                //mstring = mstring.Substring(3, mstring.IndexOf(",") - 3);
                string username = subEntry.Properties["sAMAccountName"][0].ToString();
                string fullname = subEntry.Properties["displayname"][0].ToString();
                users.Add(new UserInfo(username, fullname));
            }

            return (UserInfo[])users.ToArray(typeof(UserInfo));
        }

        public struct UserInfo
        {
            public string LoginName;
            public string FullName;

            public UserInfo(string loginName, string fullName)
            {
                LoginName = loginName;
                FullName = fullName;
            }

        }
    }


}
