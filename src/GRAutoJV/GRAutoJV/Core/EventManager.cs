using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
namespace GRAutoJV.Core
{
    public class EventManager
    {
        public static void RegisterEvents()
        {
            ConnectionManager.oApp.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(Events.ItemEventHandler.ChPoValidate);
            ConnectionManager.oApp.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(Events.FormDataEventHandler.AutoJv);
        }
    }
}
