using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace FT_ADDON
{
    abstract partial class Form_Base
    {
        [AttributeUsage(AttributeTargets.Class)]
        internal protected class Series : Attribute
        {
            string objectname, comboid;

            public Series(string objectname, string comboid)
            {
                this.objectname = objectname;
                this.comboid = comboid;
            }

            public string GetObjectName()
            {
                return objectname;
            }

            public string GetComboId()
            {
                return comboid;
            }
        }
    }

    static class SeriesAttributeExtensions
    {
        public static Form_Base.Series GetSeries(this Type type)
        {
            return type.GetCustomAttribute<Form_Base.Series>();
        }
    }
}
