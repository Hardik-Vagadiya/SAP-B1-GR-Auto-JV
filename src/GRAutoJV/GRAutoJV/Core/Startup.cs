using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GRAutoJV.Core
{
    public class Startup
    {
        public static void Initialize(string[] args)
        {
            SapApplication.Connect(args);
            SapCompany.Connect();
            EventManager.RegisterEvents();
        }
    }
}
