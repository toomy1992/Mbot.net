using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Discord.Commands;
using Discord;
using System.Linq;
using System.Xml;

namespace Mbot.Modules
{
    public class CommandMethod
    {
        
        internal static IList<IList<Object>> GoogleSheetReader(string spreadsheetid, string tabledata)
        {
            string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
            string ApplicationName = "SwGoh Bot";
            UserCredential credential;

                using (var stream =
                    new FileStream("Modules/RequestedFile/client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(
                        System.Environment.SpecialFolder.Personal);
                    credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });
            SpreadsheetsResource.ValuesResource.GetRequest Req =
                    service.Spreadsheets.Values.Get(spreadsheetid, tabledata);
            ValueRange Data = Req.Execute();
           IList<IList<Object>> DataTable = Data.Values;
           return DataTable;

        }


        internal static string GetWelcomeMsg()
        {
             var stream = new FileStream("Modules/RequestedFile/WelcomeMsg", FileMode.Open, FileAccess.Read);
             return stream.ToString();
        }

        internal static string GetSheetId()
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("Modules/RequestedFile/config.xml");
            XmlNode Token = doc.DocumentElement.SelectSingleNode("/Parameters/GoogleSheet/GoogleSheetID");
            return Token.InnerText;
        }



        internal static string GetToken()
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("Modules/RequestedFile/config.xml");
            XmlNode Token = doc.DocumentElement.SelectSingleNode("/Parameters/Discord/DiscordToken");
            return Token.InnerText;
        }

        internal async static void SendPM(List<User> usersList, SocketCommandContext context)
        {
            foreach (var User in usersList)
           {
               
               
               ulong UserId = ulong.Parse(User.id);
               var GetUser = context.Client.GetUser(UserId);
               bool send = false;
               if(GetUser == null) continue;
               if(User.Pools == null) continue;
               foreach (var Pool  in User.Pools)
               {
                   List<Order> SortedOrder = Pool.Orders.OrderBy(o=>o.PlatoonNum).ToList();
                   String Order ="";
                   foreach (var Orders in SortedOrder)
                    {
                        Order += "Plutoon Number: " + Orders.PlatoonNum + "\t Character: " + Orders.Character +" \n";          
                        send = true;              
                    }

                    await UserExtensions.SendMessageAsync(GetUser, User.name + " Here is your's " +Pool.Platoon+ " asigment ", false, EmBul("Phase: " + Pool.Phase + " Plutoon: " + Pool.Platoon, Order, Pool.Platoon));  
                }
                if(send == true)
                {
                await   context.Channel.SendMessageAsync( User.name  + " recive PM");
                send = false;
                }
           }
        }


         static EmbedBuilder EmBul(string title, string text, string Plutoon)
        {
            EmbedBuilder embedBuilder = new EmbedBuilder();
            embedBuilder.WithTitle(title)
                        .WithDescription(text);
                        
            switch (Plutoon)
            {
                case "TopPlatoon":
                    embedBuilder.WithColor(Discord.Color.Blue);
                    break;
                case "MiddlePlatoon":
                    embedBuilder.WithColor(Discord.Color.Green);
                    break;
                case "BottomPlatoon":
                    embedBuilder.WithColor(Discord.Color.Red);
                    break;
            }

            return embedBuilder;
        }
        internal static List<User> PlatoonOrders(List<User> usersList, string spreadsheetid, string TBType, string TopPlatoon, string MiddlePlatoon, string BottomPlatoon)
        {
            IList<IList<Object>> Data =  GoogleSheetReader(spreadsheetid, TBType);
            
            string Phase = "";
            foreach (var item in Data)
            {
                Phase = (string)item[0];
                
            }
            usersList = AddCharacters(Phase,spreadsheetid,TopPlatoon,usersList,"TopPlatoon");
            usersList = AddCharacters(Phase,spreadsheetid,MiddlePlatoon,usersList,"MiddlePlatoon");
            usersList = AddCharacters(Phase,spreadsheetid,BottomPlatoon,usersList,"BottomPlatoon");
            return usersList;
        }

        private static List<User> AddCharacters(string phase, string spreadsheetid, string Platoon, List<User> usersList, string PlatoonName)
        {
            IList<IList<Object>> Data =  GoogleSheetReader(spreadsheetid, Platoon);
            foreach (User user in usersList)
            {
                List<Pool> Pools;
                if(user.Pools == null)
                {
                    Pools = new List<Pool>();
                }
                else
                {
                    Pools = user.Pools;
                }
                
                List<Order> Characters = new List<Order>(); 
                Pool Pool = new Pool();
                bool PoolAddedCharacter = false;
                bool UserAddedCharacter = false;
                foreach (var Row in Data)
                {
                   
                  
                    int i = 0;
                    foreach (var item in Row)
                    {
                        
                        
                        if (user.name == (string)item)
                        {
                            Order Character = new Order();
                            Pool.Phase = phase;
                            Pool.Platoon = PlatoonName;
                            Character.Character = (string)Row[i-1];
                            Character.PlatoonNum = FindPlatoonNumber(i);
                            Characters.Add(Character);
                            Pool.Orders = Characters;

                            PoolAddedCharacter = true;
                            UserAddedCharacter = true;
                        }
                        i++;
                        
                    }
                    
                }
                if( PoolAddedCharacter == true)
                    {
                    Pools.Add(Pool);
                    PoolAddedCharacter = false;
                    }
                if( UserAddedCharacter == true)
                {
                    user.Pools = Pools;
                    UserAddedCharacter =false;
                }
            }
            

            return usersList;
        }



        private static int FindPlatoonNumber(int i)
        {
            if(i == 2)
            {
                return 1;
            }
            else if(i == 6)
            {
                return 2;
            }
            else if(i == 10)
            {
                return 3;
            }
            else if(i == 14)
            {
                return 4;
            }
            else if(i == 18)
            {
                return 5;
            }
            else if(i == 22)
            {
                return 6;
            }
            else return 0;
        }

        internal static List<User> AddUser(IList<IList<Object>> List)
        {
            List<User> UsersList = new List<User>();
            foreach (var item in List)
            {
                string id  = (string)item[1];
                char[] MyChar = { '<', '@', '>' };
                id = id.TrimStart(MyChar);
                id = id.TrimEnd(MyChar);

                User Subject = new User() as User;
                Subject.name = (string)item[0];
                Subject.id = id;
                UsersList.Add(Subject);
                
            }
            return UsersList;
        }
    }
}