//#define HANA

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Data.OleDb;
using System.Reflection;

namespace FT_ADDON.Example
{
    [NoForm]
    [FormCode("99999")]
    [ContextMenu("Upload", "upload_id", SAPbouiCOM.BoFormMode.fm_ADD_MODE)]
    class ExampleForm : Form_Base
    {
        public ExampleForm()
        {
            // Add your function here
            /*
            Example:
            AddBeforeItemFunc(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, HelloWorld);
            */
        }

        /* Example
        void HelloWorld()
        {
        }
        */


    }
}
