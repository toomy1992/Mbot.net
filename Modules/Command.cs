﻿using Discord.Commands;
using Discord;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System;
using System.Collections.Generic;
using System.Runtime;
using System.Text.RegularExpressions;

using System.Collections.Specialized;

namespace Mbot.Modules
{
    public class Commands :  ModuleBase<SocketCommandContext>
    {
    [Command("SendPlatoonOrders")]
    public async Task Send_Platoon_Orders()
        {
            string SheetId = CommandMethod.GetSheetId();
            IList<IList<Object>> List = CommandMethod.GoogleSheetReader(SheetId,"Discord!A2:B51");
            List<User> UsersList = CommandMethod.AddUser(List);
            UsersList = CommandMethod.PlatoonOrders(UsersList,SheetId,"Platoon!A2:B2","Platoon!C2:Y16","Platoon!C20:Y34","Platoon!C38:Y52");
            CommandMethod.SendPM(UsersList,Context);
            await ReplyAsync ("Request Finished");
        }

    } 

}