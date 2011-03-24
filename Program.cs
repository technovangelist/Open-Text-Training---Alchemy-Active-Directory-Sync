using System;
using System.Collections;
using System.Configuration;
using System.DirectoryServices;
using Alchemy;

namespace ADSync
{

    class Program
    {
        //const string serverName ="capadev";
        string serverName = ConfigurationSettings.AppSettings["ServerName"];
        const string domainName = "captaris";
        const string adminUser = "Administrator";
        const string adminPassword = "password";
        const string defaultPassword = "password";
        static string dbPath = "alchemy://" + serverName + ".captaris.local:3234/";
        static string dbLocalPath = @"C:\databases\";
        static Application app = new Application();
        static AuServerApi.ApplicationClass svrApp = new AuServerApi.ApplicationClass();
        static Database DB;

        static void Main(string[] args)
        {
            Console.Clear();
            Console.WriteLine("ADSYNC.EXE: Created for the Captaris Partner Conference 2007 in Kuala Lumpur");
            Console.WriteLine("            Developed by Matt Williams, Captaris International Trainer");
            Console.WriteLine("            Not for Production Use");
            Console.WriteLine();


            UserInfo[] ADUsers = GetADUsers();

            if (args.Length > 0)
            {
                switch (args[0])
                {
                    case "clean":
                        if (args.Length > 1)
                        {
                            string dbName = args[1].ToString();
                            dbPath = dbPath + dbName;
                            dbLocalPath = dbLocalPath + dbName;
                            Console.WriteLine(dbPath);
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
                        else
                            Console.WriteLine("Enter a db name that is under server control after 'clean'");

                        break;

                    default:

                        string dbname = args[0].ToString();
                        dbPath = dbPath + dbname;
                        dbLocalPath = dbLocalPath + dbname;

                        Console.WriteLine(dbPath);
                        ConnectDatabase();

                        Console.WriteLine("Found the following users in AlchemyUsers:");
                        foreach (UserInfo username in ADUsers)
                        {
                            Console.WriteLine("\t{0}: {1}", username.LoginName, username.FullName);
                            int FolderID;
                            FolderID = CreateFolder(username);
                            CreateGroup(username, FolderID);
                            CreateRole(username);
                        }

                        Builder build = app.NewBuilder(DB);
                        build.Build(null);
                        break;
                }
            }
            else
            {
                Console.WriteLine("Enter either a db name under server control or enter 'clean' followed by dbname");
            }
            
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
                    Item AdminFolder = DB.GetItemByID(AdminFolderID);
                    AdminFolder.SecurityGroups.Add(AdminUser.LoginName, "");
                    CreateRole(AdminUser);
                    DB.Refresh();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Database not found or other error\n" + ex);
            }
        }
        private static void DeleteRole(UserInfo username)
        {
            svrApp.Connect(serverName, 0, 0, 0);
            int dbID = svrApp.get_DatabaseId(dbLocalPath);
            svrApp.DeleteRoleDbAcl(username.LoginName, dbID);
            svrApp.DeleteRoleUser(username.LoginName, domainName + "\\" + username.LoginName);
            svrApp.DeleteRole(username.LoginName);
        }

        private static void CreateRole(UserInfo username)
        {
            svrApp.Connect(serverName, 0, 0, 0);
            svrApp.CreateRole(username.LoginName);
            svrApp.AddRoleUser(username.LoginName, domainName + "\\" + username.LoginName);
            int dbID = svrApp.get_DatabaseId(dbLocalPath);
            svrApp.AddRoleDbAcl(username.LoginName, dbID, "");
            svrApp.AddDbAclGroup(username.LoginName, dbID, adminPassword, username.LoginName);
        }

        private static void CreateGroup(UserInfo username, int FolderID)
        {
            Item Folder;
            DB.SecurityGroups.Add(username.LoginName, defaultPassword);
            Folder = DB.GetItemByID(FolderID);
            Folder.SecurityGroups.Add(username.LoginName, "");
        }

        private static void DeleteGroup(UserInfo username)
        {
            if (DB.HasSecurityGroups)
            {
                DB.Login(adminUser, adminPassword, "");
                DB.Refresh();
                DB.SecurityGroups.Remove(username.LoginName);
                DB.Refresh();
            }
        }

        private static void DeleteFolder(UserInfo username)
        {
            try
            {
                Query query = app.NewQuery();
                query.AddFolderQuery("Folder Title", username.FullName, false);
                query.Search(DB);
                if (query.Results.Count > 0)
                {
                    Items auItems = query.Results[1].Items;
                    if (auItems.Count == 1)
                        auItems[1].Delete();
                }
                else
                    Console.WriteLine("No Folder in Alchemy for that user");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private static int CreateFolder(UserInfo username)
        {

            int itemID = 0;

            try
            {
                Item newItem;
                newItem = DB.CreateFolder(username.FullName, DB.Root, false);
                itemID = newItem.ID;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);

            }
            return itemID;
        }

        private static UserInfo[] GetADUsers()
        {
            ArrayList users = new ArrayList();
            DirectoryEntry entry = new DirectoryEntry("LDAP://captaris");
            entry.RefreshCache();
            DirectorySearcher search = new DirectorySearcher(entry);
            search.Filter = "(&(objectClass=group)(cn=AlchemyUsers))";
            foreach (object objMember in search.FindAll()[0].GetDirectoryEntry().Properties["member"])
            {
                DirectoryEntry subEntry;
                subEntry = new DirectoryEntry("LDAP://captaris/" + objMember);
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
