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
        internal protected class MenuId : Attribute
        {
            private string id { get; set; }

            public MenuId(string _id)
            {
                id = _id;
            }

            public static implicit operator string(MenuId mi) => mi.id;
        }
    }

    static class MenuIdExtensions
    {
        public static Form_Base.MenuId GetMenuId(this Type type)
        {
            var id = type.GetCustomAttribute<Form_Base.MenuId>();

            if (id != null) return id;

            // default no menu id
            return new Form_Base.MenuId("");
        }

        public static bool HasMenuId(this Type type)
        {
            return type.GetCustomAttribute<Form_Base.MenuId>() != null;
        }
    }
}
