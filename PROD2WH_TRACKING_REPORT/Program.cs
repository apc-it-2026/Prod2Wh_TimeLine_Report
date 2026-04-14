using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PROD2WH_TRACKING_REPORT
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            client = new SJeMES_Framework.Class.ClientClass();
            client.APIURL = "http://localhost:60626//api/CommonCall";
          //   client.APIURL = "http://10.3.0.24:8092/api/CommonCall";
         
            client.UserToken = "080895fb-ebff-423f-945d-a1af07702be2";//
            client.Language = "en";
            Application.Run(new PROD2WH_Tracking_List());
        }

        public static SJeMES_Framework.Class.ClientClass client;
    }
}
