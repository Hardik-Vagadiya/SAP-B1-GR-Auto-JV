using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GRAutoJV.Core
{
    public static class ConnectionManager
    {
        public static SAPbouiCOM.Application oApp;
        public static SAPbobsCOM.Company oCom;
        public static Models.GlobalVars oGlobalVars;
    }
}
