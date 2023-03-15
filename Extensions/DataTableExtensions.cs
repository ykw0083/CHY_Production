using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FT_ADDON
{
    static class DataTableExtensions
    {
        public static void SecureQuery(this SAPbouiCOM.DataTable dt, string query)
        {
            SQLQuery.SecureQuery(query, () => dt.ExecuteQuery(query));
        }
    }
}
