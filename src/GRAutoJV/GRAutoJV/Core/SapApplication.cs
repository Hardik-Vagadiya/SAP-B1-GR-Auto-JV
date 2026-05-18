using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GRAutoJV.Core
{
    public class SapApplication
    {
        public static void Connect(string[] args)
        {
            try
            {
                //Return if no connection string is found
                if(args.Length == 0)
                {
                    throw new Exception("No SAP Connection String found!");
                }
                string connStr = args[0];
                SAPbouiCOM.SboGuiApi sboGuiApi = new SAPbouiCOM.SboGuiApi();
                sboGuiApi.Connect(connStr);
                ConnectionManager.oApp = sboGuiApi.GetApplication();
            }
            catch(Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message, "UI API Connection Error!");
            }
        }
    }
}
