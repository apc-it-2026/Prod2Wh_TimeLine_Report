using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PROD2WH_TRACKING_REPORT
{
    public class Interface
    {

		public static void RunApp(Object obj)
		{
			try
			{
				Program.client = obj as SJeMES_Framework.Class.ClientClass;
				PROD2WH_Tracking_List frm = new PROD2WH_Tracking_List();
				FormCollection collection = Application.OpenForms;
				frm.Owner = collection["frmMain"];
				frm.Show();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}
	}
}
