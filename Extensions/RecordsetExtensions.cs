using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace FT_ADDON
{
    static class RecordsetExtensions
    {
        static readonly Dictionary<Type, Func<SAPbobsCOM.Recordset, object>> basictypes = new Dictionary<Type, Func<SAPbobsCOM.Recordset, object>>
        {
            { typeof(string), GetString },
            { typeof(bool), GetBool },
            { typeof(int), GetInt },
            { typeof(double), GetDouble },
            { typeof(decimal), GetDecimal },
            { typeof(DateTime), GetDateTime },
        };

        static readonly Type maptype = typeof(IDictionary<string, object>);

        public static object GetValue(this SAPbobsCOM.Recordset rc, object field)
        {
            return rc.Fields.Item(field).Value;
        }

        // Will reset record
        public static SAPbobsCOM.Recordset GetRecord(this SAPbobsCOM.Recordset rc, object index, object filter)
        {
            rc.MoveFirst();

            while (!rc.EoF)
            {
                if (rc.GetValue(index) == filter) return rc;

                rc.MoveNext();
            }

            return null;
        }

        public static IEnumerable<T> Query<T>(this SAPbobsCOM.Recordset rc, string query)
        {
            IList<T> list = new List<T>();
            rc.DoQuery(query);
            Type type = typeof(T);

            if (basictypes.TryGetValue(type, out var GetDataFunc))
            {
                while (!rc.EoF)
                {
                    T obj = (T)GetDataFunc(rc);
                    list.Add(obj);
                    rc.MoveNext();
                }

                return list;
            }

            Func<SAPbobsCOM.Recordset, T> func;

            if (maptype.IsAssignableFrom(type))
            {
                func = GetMap<T>;
            }
            else
            {
                func = GetObject<T>;
            }

            while (!rc.EoF)
            {
                T obj = func(rc);
                list.Add(obj);
                rc.MoveNext();
            }

            return list;
        }

        private static T GetObject<T>(SAPbobsCOM.Recordset rc)
        {
            Type type = typeof(T);
            T obj = (T)Activator.CreateInstance(type);

            var map = rc.Fields.OfType<SAPbobsCOM.Field>().ToDictionary(k => k.Name, v => v.Value);

            foreach (var col in map)
            {
                var prop = type.GetProperty(col.Key, BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Public);

                if (prop == null) continue;

                prop.SetValue(obj, col.Value);
            }

            return obj;
        }

        private static T GetMap<T>(SAPbobsCOM.Recordset rc)
        {
            IDictionary<string, object> map = (IDictionary<string, object>)Activator.CreateInstance(typeof(T));

            foreach (var col in rc.Fields.OfType<SAPbobsCOM.Field>())
            {
                map.Add(col.Name, col.Value);
            }

            return (T)map;
        }

        private static object GetString(SAPbobsCOM.Recordset rc)
        {
            return rc.Fields.Item(0).Value.ToString();
        }
        
        private static object GetInt(SAPbobsCOM.Recordset rc)
        {
            return Convert.ToInt32(rc.Fields.Item(0).Value);
        }
        
        private static object GetDouble(SAPbobsCOM.Recordset rc)
        {
            return Convert.ToDouble(rc.Fields.Item(0).Value);
        }
        
        private static object GetDateTime(SAPbobsCOM.Recordset rc)
        {
            return Convert.ToDateTime(rc.Fields.Item(0).Value);
        }
        
        private static object GetBool(SAPbobsCOM.Recordset rc)
        {
            return Convert.ToBoolean(rc.Fields.Item(0).Value);
        }
        
        private static object GetDecimal(SAPbobsCOM.Recordset rc)
        {
            return Convert.ToDecimal(rc.Fields.Item(0).Value);
        }
    }
}
