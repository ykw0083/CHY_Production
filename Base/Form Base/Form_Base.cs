//#define HANA

using System;
using System.Collections.Generic;
using System.Threading;
using System.Data;
using System.Linq;
using System.Reflection;
using System.IO;
using System.Runtime.CompilerServices;
using System.Security.AccessControl;

namespace FT_ADDON
{
    abstract partial class Form_Base
    {
        #region HIDDEN
        public static Type[] list = (from domainAssembly in AppDomain.CurrentDomain.GetAssemblies()
                                     from assemblyType in domainAssembly.GetTypes()
                                     where typeof(Form_Base).IsAssignableFrom(assemblyType)
                                     where typeof(Form_Base) != assemblyType
                                     select assemblyType).ToArray();

        public static UInt64 runningProcess = 0;

        const string docid_txt = "###docid";
        const string link_btn = "###link";

        protected FormMutex cfl_mtx;
        protected FormMutex cflcond_mtx;
        protected FormMutex doclink_mtx;

        #region FORM PROPERTIES
        protected const string userdt = "CurrentUser";
        protected const string usercol = "UserName";

        const string curdoc = "CurrentDoc";
        const string docobjttype = "ObjType";
        const string doctablename = "TableName";
        const string doclinetablename = "LineTableName";
        const string docstatus = "DocStatus";
        const string docentry = "DocEntry";

        public string queryCode { get => GetType().GetFormCode(); }

        public string formFileName { get => GetType().GetFileName(); }

        public string menuId { get => GetType().GetMenuId(); }

        public string menuName { get => GetType().GetMenuName(); }

        public bool hasDynamicCFL { get => GetType().GetCustomAttribute<NoDynamicCFL>() == null; }
        public bool hasDynamicCFLCondition { get => GetType().GetCustomAttribute<NoDynamicCFLCondition>() == null; }
        public bool hasMenu { get => GetType().HasMenuId(); }
        public bool hasForm { get => GetType().GetCustomAttribute<NoForm>() == null; }
        public bool isManaged { get => GetType().GetCustomAttribute<Unmanaged>() == null; }
        public bool hasContextMenus { get => GetType().GetCustomAttributes<ContextMenu>().Any(); }
        public bool isStaticForm { get => GetType().GetCustomAttribute<StaticForm>() != null; }
        #endregion

        public bool initializing = false;

        public SAPbouiCOM.Form oForm = null;
        private List<FormSession> sessioninfo_list = new List<FormSession>();

        public SAPbouiCOM.ItemEvent itemPVal { get => sessioninfo_list.Last().itemPVal; }
        public SAPbouiCOM.MenuEvent menuPVal { get => sessioninfo_list.Last().menuPVal; }
        public SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo { get => sessioninfo_list.Last().BusinessObjectInfo; }
        public SAPbouiCOM.ContextMenuInfo rcPVal { get => sessioninfo_list.Last().rcPVal; }
        public List<ActionResult> actionResults
        {
            get => sessioninfo_list.Last().actionResults;
            set => sessioninfo_list.Last().actionResults = value;
        }
        public bool BubbleEvent
        {
            get => sessioninfo_list.Last().BubbleEvent;
            set => sessioninfo_list.Last().BubbleEvent = value;
        }
        protected string currentId { get => sessioninfo_list.Last().currentId; }
        protected string colId { get => sessioninfo_list.Last().colId; }
        protected int currentRow { get => sessioninfo_list.Last().currentRow; }
        protected bool beforeAction { get => sessioninfo_list.Last().beforeAction; }
        protected bool actionSuccess { get => sessioninfo_list.Last().actionSuccess; }

        protected SAPbouiCOM.DataTable cflDataTable { get => (itemPVal as SAPbouiCOM.IChooseFromListEvent).SelectedObjects; }

        private Dictionary<SAPbouiCOM.BoEventTypes, Action> _beforeItem = new Dictionary<SAPbouiCOM.BoEventTypes, Action>();
        private Dictionary<SAPbouiCOM.BoEventTypes, Action> _afterItem = new Dictionary<SAPbouiCOM.BoEventTypes, Action>();

        private Dictionary<string, Action> _beforeMenu = new Dictionary<string, Action>();
        private Dictionary<string, Action> _afterMenu = new Dictionary<string, Action>();

        private Dictionary<SAPbouiCOM.BoEventTypes, Action> _beforeData = new Dictionary<SAPbouiCOM.BoEventTypes, Action>();
        private Dictionary<SAPbouiCOM.BoEventTypes, Action> _afterData = new Dictionary<SAPbouiCOM.BoEventTypes, Action>();

        private Dictionary<string, Action> _beforeRightClick = new Dictionary<string, Action>();
        private Dictionary<string, Action> _afterRightClick = new Dictionary<string, Action>();

        protected Dictionary<SAPbouiCOM.BoEventTypes, Action> beforeItem { get => _beforeItem; }
        protected Dictionary<SAPbouiCOM.BoEventTypes, Action> afterItem { get => _afterItem; }

        protected Dictionary<string, Action> beforeMenu { get => _beforeMenu; }
        protected Dictionary<string, Action> afterMenu { get => _afterMenu; }

        protected Dictionary<SAPbouiCOM.BoEventTypes, Action> beforeData { get => _beforeData; }
        protected Dictionary<SAPbouiCOM.BoEventTypes, Action> afterData { get => _afterData; }

        protected Dictionary<string, Action> beforeRightClick { get => _beforeRightClick; }
        protected Dictionary<string, Action> afterRightClick { get => _afterRightClick; }

        #region GET
        protected object GetObjectValue(object obj)
        {
            SAPbouiCOM.Item itm = oForm.Items.Item(obj);

            try
            {
                switch (itm.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                        return (itm.Specific as SAPbouiCOM.EditText).Value;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        return (itm.Specific as SAPbouiCOM.ComboBox).Value;
                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                        return (itm.Specific as SAPbouiCOM.CheckBox).Checked;
                }

                throw new Exception($"GetObjectValue does not support this object - { obj }");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(itm);
                itm = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        
        protected object GetObjectValue(object obj, object col, int row)
        {
            SAPbouiCOM.Item itm = oForm.Items.Item(obj);

            try
            {
                switch (itm.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                        return (itm.Specific as SAPbouiCOM.Grid).DataTable.GetValue(col, row);
                    case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                        return (itm.Specific as SAPbouiCOM.ComboBox).Value;
                }

                throw new Exception($"GetObjectValue does not support this object - { obj }");
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(itm);
                itm = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected void GetCellValue(SAPbouiCOM.Column col, object row, object value)
        {
            SAPbouiCOM.Cell itm = col.Cells.Item(row);

            try
            {
                switch (col.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                        (itm.Specific as SAPbouiCOM.EditText).Value = value.ToString();
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        (itm.Specific as SAPbouiCOM.ComboBox).Select(value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                        (itm.Specific as SAPbouiCOM.CheckBox).Checked = Convert.ToBoolean(value);
                        break;
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(itm);
                itm = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected void GetCellValue(SAPbouiCOM.Column col, SAPbouiCOM.Cell itm, object value)
        {
            switch (col.Type)
            {
                case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                    (itm.Specific as SAPbouiCOM.EditText).Value = value.ToString();
                    break;
                case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                    (itm.Specific as SAPbouiCOM.ComboBox).Select(value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    break;
                case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                    (itm.Specific as SAPbouiCOM.CheckBox).Checked = Convert.ToBoolean(value);
                    break;
            }
        }
        protected SAPbouiCOM.EditText GetText(string itm)
        {
            return oForm.Items.Item(itm).Specific as SAPbouiCOM.EditText;
        }

        protected SAPbouiCOM.Grid GetGrid(string itm)
        {
            return oForm.Items.Item(itm).Specific as SAPbouiCOM.Grid;
        }

        protected SAPbouiCOM.Matrix GetMatrix(string itm)
        {
            return oForm.Items.Item(itm).Specific as SAPbouiCOM.Matrix;
        }

        protected SAPbouiCOM.ComboBox GetCombo(string itm)
        {
            return oForm.Items.Item(itm).Specific as SAPbouiCOM.ComboBox;
        }

        protected SAPbouiCOM.CheckBox GetCheckBox(string itm)
        {
            return oForm.Items.Item(itm).Specific as SAPbouiCOM.CheckBox;
        }

        protected SAPbouiCOM.Button GetButton(string itm)
        {
            return oForm.Items.Item(itm).Specific as SAPbouiCOM.Button;
        }

        protected SAPbouiCOM.ButtonCombo GetButtonCombo(string itm)
        {
            return oForm.Items.Item(itm).Specific as SAPbouiCOM.ButtonCombo;
        }
        #endregion

        #region SET
        protected void SetObjectValue(object obj, object value)
        {
            SAPbouiCOM.Item itm = oForm.Items.Item(obj);

            try
            {
                switch (itm.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                        (itm.Specific as SAPbouiCOM.EditText).Value = value.ToString();
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        (itm.Specific as SAPbouiCOM.ComboBox).Select(value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                        (itm.Specific as SAPbouiCOM.CheckBox).Checked = Convert.ToBoolean(value);
                        break;
                    default:
                        throw new Exception($"SetObjectValue does not support this object - { obj }");
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(itm);
                itm = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        
        protected void SetObjectValue(object obj, object col, int row, object value)
        {
            SAPbouiCOM.Item itm = oForm.Items.Item(obj);

            try
            {
                switch (itm.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                        (itm.Specific as SAPbouiCOM.Grid).DataTable.SetValue(col, row, value);
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                        SetCellValue((itm.Specific as SAPbouiCOM.Matrix).Columns.Item(col), row, value);
                        break;
                    default:
                        throw new Exception($"SetObjectValue does not support this object - { obj }");
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(itm);
                itm = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected void SetObjectValue(SAPbouiCOM.Item itm, object value)
        {
            switch (itm.Type)
            {
                case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                    (itm.Specific as SAPbouiCOM.EditText).Value = value.ToString();
                    break;
                case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                    (itm.Specific as SAPbouiCOM.ComboBox).Select(value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    break;
                case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                    (itm.Specific as SAPbouiCOM.CheckBox).Checked = Convert.ToBoolean(value);
                    break;
            }
        }

        // swapped column parameter with row parameter
        protected void SetCellValue(SAPbouiCOM.Column col, object row, object value)
        {
            SAPbouiCOM.Cell itm = col.Cells.Item(row);

            try
            {
                switch (col.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                        (itm.Specific as SAPbouiCOM.EditText).Value = value.ToString();
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        (itm.Specific as SAPbouiCOM.ComboBox).Select(value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                        (itm.Specific as SAPbouiCOM.CheckBox).Checked = Convert.ToBoolean(value);
                        break;
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(itm);
                itm = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // swapped column parameter with itm parameter
        protected void SetCellValue(SAPbouiCOM.Column col, SAPbouiCOM.Cell itm, object value)
        {
            switch (col.Type)
            {
                case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                    (itm.Specific as SAPbouiCOM.EditText).Value = value.ToString();
                    break;
                case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                    (itm.Specific as SAPbouiCOM.ComboBox).Select(value, SAPbouiCOM.BoSearchKey.psk_ByValue);
                    break;
                case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                    (itm.Specific as SAPbouiCOM.CheckBox).Checked = Convert.ToBoolean(value);
                    break;
            }
        }

        protected void SetGrid(string itm, object col, int row, object value)
        {
            (oForm.Items.Item(itm).Specific as SAPbouiCOM.Grid).DataTable.SetValue(col, row, value);
        }

        protected void SetGrid(SAPbouiCOM.Grid itm, object col, int row, object value)
        {
            itm.DataTable.SetValue(col, row, value);
        }

        protected void SetMatrix(string itm, object col, object row, object value)
        {
            SAPbouiCOM.Column column = (oForm.Items.Item(itm).Specific as SAPbouiCOM.Matrix).Columns.Item(col);

            try
            {
                SetCellValue(column, column.Cells.Item(row), value);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(column);
                column = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected void SetMatrix(SAPbouiCOM.Matrix itm, object col, int row, object value)
        {
            SAPbouiCOM.Column column = itm.Columns.Item(col);

            try
            {
                SetCellValue(column, column.Cells.Item(row), value);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(column);
                column = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void AddBeforeFunc<D, K>(D delegatemap, K key, Action func) where D : IDictionary<K, Action>
        {
            if (!delegatemap.ContainsKey(key))
            {
                delegatemap.Add(key, func);
                delegatemap[key] += CheckBubble;
                return;
            }

            delegatemap[key] += func;
            delegatemap[key] += CheckBubble;
        }
        
        private void AddAfterFunc<D, K>(D delegatemap, K key, Action func) where D : IDictionary<K, Action>
        {
            if (!delegatemap.ContainsKey(key))
            {
                delegatemap.Add(key, func);
                return;
            }

            delegatemap[key] += func;
        }

        protected void AddBeforeItemFunc(SAPbouiCOM.BoEventTypes key, Action func)
        {
            AddBeforeFunc(beforeItem, key, func);
        }

        protected void AddAfterItemFunc(SAPbouiCOM.BoEventTypes key, Action func)
        {
            AddAfterFunc(afterItem, key, func);
        }

        protected void AddBeforeDataFunc(SAPbouiCOM.BoEventTypes key, Action func)
        {
            AddBeforeFunc(beforeData, key, func);
        }

        protected void AddAfterDataFunc(SAPbouiCOM.BoEventTypes key, Action func)
        {
            AddAfterFunc(afterData, key, func);
        }

        protected void AddBeforeMenuFunc(string key, Action func)
        {
            AddBeforeFunc(beforeMenu, key, func);
        }

        protected void AddAfterMenuFunc(string key, Action func)
        {
            AddAfterFunc(afterMenu, key, func);
        }

        protected void AddBeforeRightClickFunc(string key, Action func)
        {
            AddBeforeFunc(beforeRightClick, key, func);
        }

        protected void AddAfterRightClickFunc(string key, Action func)
        {
            AddAfterFunc(afterRightClick, key, func);
        }
        #endregion

        public Form_Base()
        {
            SetFormMutex();
            SetAutoFill();
            SetCFLCondition();
            SetDocLink();
            SetSysDocUpdate();
        }

        protected void CreateReferTableToDoc()
        {
            SAPbouiCOM.DataTable dt = null;

            try
            {
                string tablename = oForm.DataSources.DBDataSources.Item(0).TableName;
                object docobjtypevalue = oForm.DataSources.DBDataSources.Item(0).GetValue(docobjttype, 0);
                object docstatusvalue = oForm.DataSources.DBDataSources.Item(0).GetValue(docstatus, 0);

                dt = oForm.DataSources.DataTables.Add(userdt);
                dt.Columns.Add(usercol, SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                dt.Rows.Add();
                dt.SetValue(usercol, 0, SAP.SBOCompany.UserName);

                dt = oForm.DataSources.DataTables.Add(curdoc);
                dt.Columns.Add(docobjttype, SAPbouiCOM.BoFieldsType.ft_Integer);
                dt.Columns.Add(doctablename, SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                dt.Columns.Add(doclinetablename, SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                dt.Columns.Add(docstatus, SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                dt.Columns.Add(docentry, SAPbouiCOM.BoFieldsType.ft_Integer);
                dt.Rows.Add();
                dt.SetValue(docobjttype, 0, docobjtypevalue);
                dt.SetValue(doctablename, 0, tablename);
                dt.SetValue(doclinetablename, 0, tablename.Substring(1) + "1");
                dt.SetValue(docstatus, 0, docstatusvalue);
                dt.SetValue(docentry, 0, 0);
            }
            catch (Exception)
            {
            }
            finally
            {
                if (dt != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dt);
                    dt = null;
                    GC.Collect();
                }
            }
        }

        protected void DocLinkSetup()
        {
            if (doclink_mtx.IsMutexOwned()) return;

            oForm.DataSources.UserDataSources.Add(docid_txt, SAPbouiCOM.BoDataType.dt_SHORT_TEXT);

            try
            {
                oForm.Freeze(true);
                var item = oForm.Items.Add(docid_txt, SAPbouiCOM.BoFormItemTypes.it_EDIT);

                try
                {
                    item.Width = 1;
                    item.Height = 1;
                    (item.Specific as SAPbouiCOM.EditText).DataBind.SetBound(true, "", docid_txt);

                    item = oForm.Items.Add(link_btn, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                    item.Visible = false;
                    item.Width = 10;
                    item.Height = 10;
                    item.LinkTo = docid_txt;
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(item);
                    item = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        protected void UpdateSystemDocStatus()
        {
            if (!oForm.HasDataTable(curdoc)) return;

            var dt = oForm.DataSources.DataTables.Item(curdoc);

            try
            {
                dt.SetValue(docstatus, 0, oForm.DataSources.DBDataSources.Item(0).GetValue(docstatus, 0));
                dt.SetValue(docentry, 0, oForm.DataSources.DBDataSources.Item(0).GetValue(docentry, 0));
                dt.SetValue(docobjttype, 0, oForm.DataSources.DBDataSources.Item(0).GetValue(docobjttype, 0));
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dt);
                dt = null;
                GC.Collect();
            }
        }

        protected void OpenFormByKey(SAPbouiCOM.BoLinkedObject type, string key)
        {
            oForm.DataSources.UserDataSources.Item(docid_txt).Value = key;
            var lb = oForm.Items.Item(link_btn).Specific as SAPbouiCOM.LinkedButton;

            try
            {
                lb.LinkedObject = type;
                var item = oForm.Items.Item(link_btn);

                try
                {
                    oForm.Freeze(true);
                    item.Visible = true;
                    item.Click();
                    item.Visible = false;
                }
                finally
                {
                    oForm.Freeze(false);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(item);
                    item = null;
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(lb);
                lb = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        
        protected void OpenFormByKey(string udoname, string key)
        {
            oForm.DataSources.UserDataSources.Item(docid_txt).Value = key;
            var lb = oForm.Items.Item(link_btn).Specific as SAPbouiCOM.LinkedButton;

            try
            {
                lb.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_None;
                lb.LinkedObjectType = udoname;
                var item = oForm.Items.Item(link_btn);

                try
                {
                    oForm.Freeze(true);
                    item.Visible = true;
                    item.Click();
                    item.Visible = false;
                }
                finally
                {
                    oForm.Freeze(false);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(item);
                    item = null;
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(lb);
                lb = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        protected void SetCurrentDocEntry(int docentry)
        {
            var dt = oForm.DataSources.DataTables.Item(curdoc);

            dt.SetValue(docstatus, 0, docentry);
        }

        protected void SetCurrentDocStatus(string status)
        {
            var dt = oForm.DataSources.DataTables.Item(curdoc);

            dt.SetValue(docstatus, 0, status);
        }

        protected void SetCurrentObjType(string objtype)
        {
            var dt = oForm.DataSources.DataTables.Item(curdoc);

            dt.SetValue(docobjttype, 0, objtype);
        }

        protected virtual void runtimeTweakBefore()
        {
        }

        protected virtual void runtimeTweakAfter()
        {
        }

        public virtual void initialize()
        {
            initialize(queryCode);
        }

        public virtual void initialize(string menuID)
        {
            if (!hasForm) return;

            bool done = false;

            try
            {
                initializing = true;
                System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                string path = System.Windows.Forms.Application.StartupPath;
                string name = this.GetType().Namespace.Replace(MethodBase.GetCurrentMethod().DeclaringType.Namespace + ".", "")
                                                      .Replace(".", "\\");
                xmlDoc.Load($"{ path }\\{ name }\\{ formFileName }");
                System.Xml.XmlAttributeCollection xmlCol = xmlDoc.LastChild.FirstChild.FirstChild.FirstChild.Attributes;

                foreach (System.Xml.XmlAttribute att in xmlCol)
                {
                    if (att.Value != "FT_Type") continue;

                    att.Value = menuID;
                }

                System.Xml.XmlNode node = xmlDoc.LastChild.FirstChild.FirstChild.FirstChild;
                xmlCol = node.Attributes;
                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = $"FT_{ SAP.getNewformUID() }";
                creationPackage.XmlData = xmlDoc.InnerXml;     // Load form from xml 
                oForm = SAP.SBOApplication.Forms.AddEx(creationPackage);
                oForm.AutoManaged = isManaged;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                // force client height
                //oForm.ClientHeight = int.Parse(xmlCol.GetNamedItem("client_height").Value);

                if (oForm.Items.Count > 0)
                {
                    oForm.Items.OfType<SAPbouiCOM.Item>()
                               .Where(itm => itm.UniqueID.ToLower().Contains("loading"))
                               .ToList()
                               .ForEach(itm =>
                               {
                                   SAPbouiCOM.PictureBox pbox = itm.Specific as SAPbouiCOM.PictureBox;
                                   pbox.Picture = $"{ path }\\Resources\\Loading.jpg";
                               });
                }

                try
                {
                    InitializeFormMutex();
                    runtimeTweakBefore();
                    CFLSetup();
                    CFLConditionSetup();
                    DocLinkSetup();
                    Common.setFormDefaultValue(oForm);
                }
                finally
                {
                    runtimeTweakAfter();
                }

                oForm.Visible = true;
                done = true;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(Common.ReadException(ex), 1, "OK", "", "");
            }
            finally
            {
                initializing = false;

                if (oForm != null && !done) oForm.Close();
            }
        }

        private void SetFormMutex()
        {
            if (hasForm) return;

            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_FORM_LOAD, InitializeFormMutex);
        }

        private void SetAutoFill()
        {
            var autofillstatus = this.GetType().GetAutoFillFromList();

            if (autofillstatus == null) return;

            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, AutoFillChooseFromList);
        }
        
        private void SetCFLCondition()
        {
            if (!hasForm) return;

            if (!hasDynamicCFLCondition) return;

            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_FORM_DRAW, CFLConditionSetup);
        }

        private void SetDocLink()
        {
            if (hasForm) return;

            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_FORM_LOAD, DocLinkSetup);
        }

        private void SetSysDocUpdate() => AddAfterDataFunc(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD, UpdateSystemDocStatus);

        protected void CFLSetup()
        {
            if (!hasDynamicCFL) return;

            if (cfl_mtx.IsMutexOwned()) return;

            if (oForm.Items.Count == 0) return;

            var itemlist = oForm.Items.OfType<SAPbouiCOM.Item>().Where(item => IsItemValidForCFL(item));

            if (!itemlist.Any()) return;

            SAPbouiCOM.ChooseFromList cfl = null;
            SAPbouiCOM.EditText txt = null;

            try
            {
                foreach (var item in itemlist)
                {
                    if (!DynamicChooseFromList.TryGetChooseFromList($"{ oForm.TypeEx }.{ item.UniqueID }", out var dcfl)) continue;

                    try
                    {
                        if (oForm.ChooseFromLists.HasItem(dcfl.parameters)) return;

                        cfl = oForm.ChooseFromLists.Add(dcfl.parameters);
                        txt = item.Specific as SAPbouiCOM.EditText;
                        txt.ChooseFromListUID = cfl.UniqueID;
                        txt.ChooseFromListAlias = dcfl.alias;
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dcfl.parameters);
                        dcfl.parameters = null;
                    }
                }
            }
            finally
            {
                if (cfl != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cfl);
                    cfl = null;
                }

                if (txt != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(txt);
                    txt = null;
                }
            }
        }

        private bool IsItemValidForCFL(SAPbouiCOM.Item item)
        {
            if (item.Type != SAPbouiCOM.BoFormItemTypes.it_EDIT && item.Type != SAPbouiCOM.BoFormItemTypes.it_EXTEDIT) return false;

            SAPbouiCOM.EditText txt = item.Specific as SAPbouiCOM.EditText;
            SAPbouiCOM.DataTable dt = null;
            SAPbouiCOM.DBDataSource db = null;
            SAPbouiCOM.UserDataSource uds = null;
            SAPbouiCOM.DataBind dbind = txt.DataBind;

            try
            {
                if (!dbind.DataBound) return false;

                if (oForm.TryGetDataSource(dbind.TableName, out db)) return db.Fields.Item(dbind.Alias).Type == SAPbouiCOM.BoFieldsType.ft_AlphaNumeric;

                if (oForm.TryGetDataTable(dbind.TableName, out dt)) return dt.Columns.Item(dbind.Alias).Type == SAPbouiCOM.BoFieldsType.ft_AlphaNumeric;

                if (oForm.TryGetUserSource(dbind.TableName, out uds)) return uds.DataType == SAPbouiCOM.BoDataType.dt_SHORT_TEXT || uds.DataType == SAPbouiCOM.BoDataType.dt_LONG_TEXT;

                return false;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dbind);
                dbind = null;
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(txt);
                txt = null;

                if (db != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(db);
                    db = null;
                }

                if (dt != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dt);
                    dt = null;
                }

                if (uds != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(uds);
                    uds = null;
                }
            }
        }

        protected void CFLConditionSetup()
        {
            if (cflcond_mtx.IsMutexOwned()) return;

            if (oForm.ChooseFromLists.Count == 0) return;

            var cfllist = oForm.ChooseFromLists.OfType<SAPbouiCOM.ChooseFromList>();

            foreach (var cfl in cfllist)
            {
                try
                {
                    if (!ChooseFromListCondition.TryGetConditions($"{ queryCode }.{ cfl.UniqueID }", out var conditions)) continue;

                    cfl.SetConditions(conditions);
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cfl);
                }
            }
        }

        private void AutoFillChooseFromList()
        {
            if (cflDataTable == null) return;

            var item = oForm.Items.Item(currentId);
            SAPbouiCOM.DBDataSource db = null;
            SAPbouiCOM.DataTable dt = null;

            try
            {
                string code;
                string tablename;
                string alias;
                int row = 0;

                switch (item.Type)
                {
                    case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                    case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                        GetTextCFLInfo(item, out code, out tablename, out alias);
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        GetComboBoxCFLInfo(item, out code, out tablename, out alias);
                        break;
                    case SAPbouiCOM.BoFormItemTypes.it_GRID:
                        GetGridCFLInfo(item, out code, out tablename, out alias);
                        row = currentRow;
                        break;
                    default:
                        return;
                }

                if (alias == null) return;

                oForm.Freeze(true);

                try
                {
                    if (tablename == null)
                    {
                        if (alias == String.Empty || !oForm.HasUserSource(alias)) return;

                        oForm.SetUserSourceValue(alias, code);
                    }
                    else if (oForm.TryGetDataSource(tablename, out db))
                    {
                        db.SetValue(alias, row, code);
                    }
                    else if (oForm.TryGetDataTable(tablename, out dt))
                    {
                        dt.SetValue(alias, row, code);
                    }
                }
                finally
                {
                    oForm.Freeze(false);
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(item);
                item = null;
                
                if (db != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(db);
                    db = null;
                }
                
                if (dt != null)
                {
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(dt);
                    dt = null;
                }

                GC.Collect();
            }
        }

        private void GetTextCFLInfo(SAPbouiCOM.Item item, out string code, out string tablename, out string alias)
        {
            code = null;
            tablename = null;
            alias = null;
            var txt = item.Specific as SAPbouiCOM.EditText;

            try
            {
                code = cflDataTable.GetValue(txt.ChooseFromListAlias, 0).ToString();
                tablename = txt.DataBind.TableName;
                alias = txt.DataBind.Alias;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(txt);
                txt = null;
                GC.Collect();
            }
        }

        private void GetComboBoxCFLInfo(SAPbouiCOM.Item item, out string code, out string tablename, out string alias)
        {
            code = null;
            tablename = null;
            alias = null;
            var cbox = item.Specific as SAPbouiCOM.ComboBox;

            try
            {
                code = cflDataTable.GetValue(0, 0).ToString();
                tablename = cbox.DataBind.TableName;
                alias = cbox.DataBind.Alias;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(cbox);
                cbox = null;
                GC.Collect();
            }
        }

        private void GetGridCFLInfo(SAPbouiCOM.Item item, out string code, out string tablename, out string alias)
        {
            code = null;
            tablename = null;
            alias = null;

            var grid = item.Specific as SAPbouiCOM.Grid;
            var gdt = grid.DataTable;
            var coltxt = grid.Columns.Item(colId) as SAPbouiCOM.EditTextColumn;

            try
            {
                code = cflDataTable.GetValue(coltxt.ChooseFromListAlias, 0).ToString();
                tablename = gdt.UniqueID;
                alias = coltxt.UniqueID;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(grid);
                grid = null;
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(gdt);
                gdt = null;
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(coltxt);
                coltxt = null;
                GC.Collect();
            }
        }

        private bool blockItem(SAPbouiCOM.Form Form, string item)
        {
            try
            {
                return item.Length == 0 || Form.Items.Item(item).Enabled;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private void CheckBubble()
        {
            if (!BubbleEvent) throw new BubbleCrash();
        }

        private bool PreRunCheck(SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, string uid)
        {
            if (initializing)
            {
                if (!pVal.BeforeAction && (beforeItem.ContainsKey(pVal.EventType) || afterItem.ContainsKey(pVal.EventType)) &&
                    (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD || pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DRAW))
                {
                    throw new Exception(pVal.EventType.ToString() + " is not supported for custom form, use runtimeTweakBefore/runtimeTweakAfter instead");
                }

                return false;
            }

            if (!blockItem(Form, uid)) return false;

            return true;
        }

        protected void Reset()
        {
            _beforeItem = new Dictionary<SAPbouiCOM.BoEventTypes, Action>();
            _afterItem = new Dictionary<SAPbouiCOM.BoEventTypes, Action>();

            _beforeMenu = new Dictionary<string, Action>();
            _afterMenu = new Dictionary<string, Action>();

            _beforeData = new Dictionary<SAPbouiCOM.BoEventTypes, Action>();
            _afterData = new Dictionary<SAPbouiCOM.BoEventTypes, Action>();

            _beforeRightClick = new Dictionary<string, Action>();
            _afterRightClick = new Dictionary<string, Action>();
        }

        private void InsertEvent(SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent evnt)
        {
            oForm = Form;
            sessioninfo_list.Add(new FormSession(evnt));
        }

        private void InsertEvent(SAPbouiCOM.Form Form, SAPbouiCOM.MenuEvent evnt)
        {
            oForm = Form;
            sessioninfo_list.Add(new FormSession(evnt));
        }

        private void InsertEvent(SAPbouiCOM.Form Form, SAPbouiCOM.BusinessObjectInfo evnt)
        {
            oForm = Form;
            sessioninfo_list.Add(new FormSession(evnt));
        }

        private void InsertEvent(SAPbouiCOM.Form Form, SAPbouiCOM.ContextMenuInfo evnt)
        {
            oForm = Form;
            sessioninfo_list.Add(new FormSession(evnt));
        }

        private void RemoveCurrentEvent()
        {
            sessioninfo_list.Remove(sessioninfo_list.Last());
        }

        private static string GetFormCode(Type formtype)
        {
            var formcode = formtype.GetFormCode();
            string code = formcode != null ? formcode : formtype.Namespace.Substring(formtype.BaseType.Namespace.Length + 1);
            return code.Split('.').Last();
        }

        public static bool GetFormTypes(string formTypeEx, out List<Type> formtypes)
        {
            formtypes = list.Where(type => GetFormCode(type) == formTypeEx).ToList();
            return formtypes.Count > 0;
        }

        public static void CreateForm(string formuid, Type formtype)
        {
            if (!AddOn.masterFormList.TryGetValue(formuid, out var list))
            {
                AddOn.masterFormList.Add(formuid, new List<Form_Base>());
                list = AddOn.masterFormList[formuid];
            }

            Form_Base formobj = Activator.CreateInstance(formtype) as Form_Base;
            list.Add(formobj);
        }
        
        public static void CreateForm(string formuid, Form_Base formobj)
        {
            if (!AddOn.masterFormList.TryGetValue(formuid, out var list))
            {
                AddOn.masterFormList.Add(formuid, new List<Form_Base>());
                list = AddOn.masterFormList[formuid];
            }

            list.Add(formobj);
        }

        virtual public void FormRemovalEvent()
        {
        }

        public static void ClearEmptyForms()
        {
            for (int i = 0; i < AddOn.masterFormList.Keys.Count; i++)
            {
                string key = AddOn.masterFormList.Keys.ElementAt(i);

                try
                {
                    SAP.SBOApplication.Forms.Item(key);
                }
                catch (Exception)
                {
                    foreach (var formobj in AddOn.masterFormList[key])
                    {
                        formobj.FormRemovalEvent();
                    }

                    AddOn.masterFormList.Remove(key);
                    --i;
                }
            }
        }

        public static bool GetForms(string formuid, out List<Form_Base> list)
        {
            var form = SAP.SBOApplication.Forms.Item(formuid);

            if (AddOn.masterFormList.TryGetValue(form.UniqueID, out list)) return true;

            if (!GetFormTypes(form.TypeEx, out var formtypes)) return false;

            formtypes = formtypes.Where(formtype => formtype.GetNoForm() != null).ToList();

            if (formtypes.Count == 0) return false;

            list = new List<Form_Base>();

            foreach (var formtype in formtypes)
            {
                CreateForm(form.UniqueID, formtype);
            }

            list = AddOn.masterFormList[form.UniqueID];
            return list.Count > 0;
        }

        public static Form_Type OpenNewForm<Form_Type>() where Form_Type : Form_Base
        {
            Form_Type formobj = NewForm<Form_Type>();

            if (formobj == null) return null;

            formobj.OpenForm();
            return formobj;
        }

        public static Form_Base OpenNewForm(Type formtype)
        {
            Form_Base formobj = NewForm(formtype);

            if (formobj == null) return null;

            formobj.OpenForm();
            return formobj;
        }

        public static Form_Type NewForm<Form_Type>() where Form_Type : Form_Base
        {
            if (!list.ToList().Contains(typeof(Form_Type))) return null;

            return Activator.CreateInstance(typeof(Form_Type)) as Form_Type;
        }

        public static Form_Base NewForm(Type formtype)
        {
            if (!list.ToList().Contains(formtype)) return null;

            return Activator.CreateInstance(formtype) as Form_Base;
        }

        public void OpenForm()
        {
            initialize();

            if (oForm == null) return;

            CreateForm(oForm.UniqueID, this);
        }

        protected bool ExeActionWithEvent(Action action, SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent evnt)
        {
            InsertEvent(Form, evnt);

            try
            {
                return ExeAction(action);
            }
            finally
            {
                RemoveCurrentEvent();
            }
        }
        
        protected bool ExeActionWithEvent(Action action, SAPbouiCOM.Form Form, SAPbouiCOM.MenuEvent evnt)
        {
            InsertEvent(Form, evnt);

            try
            {
                return ExeAction(action);
            }
            finally
            {
                RemoveCurrentEvent();
            }
        }
        
        protected bool ExeActionWithEvent(Action action, SAPbouiCOM.Form Form, SAPbouiCOM.BusinessObjectInfo evnt)
        {
            InsertEvent(Form, evnt);

            try
            {
                return ExeAction(action);
            }
            finally
            {
                RemoveCurrentEvent();
            }
        }
        
        protected bool ExeActionWithEvent(Action action, SAPbouiCOM.Form Form, SAPbouiCOM.ContextMenuInfo evnt)
        {
            InsertEvent(Form, evnt);

            try
            {
                return ExeAction(action);
            }
            finally
            {
                RemoveCurrentEvent();
            }
        }

        protected bool ExeAction(Action action)
        {
            try
            {
                try { action(); }
                catch (BubbleCrash) { }

                if (actionResults.Count > 0) SAP.showActionResult(actionResults);
            }
            catch (MessageException ex)
            {
                BubbleEvent = false;

                if (GetType() == typeof(SystemForm)) return BubbleEvent;

                SAP.stopProgressBar();
                SAP.SBOApplication.MessageBox(ex.Message, 1, "OK", "", "");
            }
            catch (Exception ex)
            {
                BubbleEvent = false;

                if (GetType() == typeof(SystemForm)) return BubbleEvent;

                SAP.stopProgressBar();
                SAP.SBOApplication.MessageBox(Common.ReadException(ex), 1, "OK", "", "");
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            return BubbleEvent;
        }
        
        public virtual void processItemEventbefore(SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal, ref bool _BubbleEvent)
        {
            try
            {
                if (!PreRunCheck(Form, pVal, pVal.ItemUID)) return;

                if (beforeItem.Count == 0) return;

                if (beforeItem.ContainsKey(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS))
                {
                    _BubbleEvent = ExeActionWithEvent(beforeItem[SAPbouiCOM.BoEventTypes.et_ALL_EVENTS], Form, pVal);

                    if (!_BubbleEvent) return;
                }

                if (!beforeItem.ContainsKey(pVal.EventType)) return;

                _BubbleEvent = ExeActionWithEvent(beforeItem[pVal.EventType], Form, pVal);
            }
            finally
            {
                // NECESSARY TO PREVENT CRASH IN SAP
                if (!_BubbleEvent && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) _BubbleEvent = true;
            }
        }

        public virtual void processItemEventafter(SAPbouiCOM.Form Form, SAPbouiCOM.ItemEvent pVal)
        {
            if (!PreRunCheck(Form, pVal, pVal.ItemUID)) return;

            if (afterItem.Count == 0) return;

            if (afterItem.ContainsKey(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS))
            {
                ExeActionWithEvent(afterItem[SAPbouiCOM.BoEventTypes.et_ALL_EVENTS], Form, pVal);
            }

            if (!afterItem.ContainsKey(pVal.EventType)) return;

            ExeActionWithEvent(afterItem[pVal.EventType], Form, pVal);
        }

        public virtual void processMenuEventbefore(SAPbouiCOM.Form Form, SAPbouiCOM.MenuEvent pVal, ref bool _BubbleEvent)
        {
            if (!beforeMenu.ContainsKey(pVal.MenuUID)) return;

            _BubbleEvent = ExeActionWithEvent(beforeMenu[pVal.MenuUID], Form, pVal);
        }

        public virtual void processMenuEventafter(SAPbouiCOM.Form Form, SAPbouiCOM.MenuEvent pVal)
        {
            if (!afterMenu.ContainsKey(pVal.MenuUID)) return;

            ExeActionWithEvent(afterMenu[pVal.MenuUID], Form, pVal);
        }

        public virtual void processDataEventbefore(SAPbouiCOM.Form Form, SAPbouiCOM.BusinessObjectInfo _BusinessObjectInfo, ref bool _BubbleEvent)
        {
            if (beforeData.Count == 0) return;

            if (beforeData.ContainsKey(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS))
            {
                _BubbleEvent = ExeActionWithEvent(beforeData[SAPbouiCOM.BoEventTypes.et_ALL_EVENTS], Form, _BusinessObjectInfo);

                if (!_BubbleEvent) return;
            }

            if (!beforeData.ContainsKey(_BusinessObjectInfo.EventType)) return;

            _BubbleEvent = ExeActionWithEvent(beforeData[_BusinessObjectInfo.EventType], Form, _BusinessObjectInfo);
        }

        public virtual void processDataEventafter(SAPbouiCOM.Form Form, SAPbouiCOM.BusinessObjectInfo _BusinessObjectInfo)
        {
            if (afterData.Count == 0) return;

            if (afterData.ContainsKey(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS))
            {
                ExeActionWithEvent(afterData[SAPbouiCOM.BoEventTypes.et_ALL_EVENTS], Form, _BusinessObjectInfo);
            }

            if (!afterData.ContainsKey(_BusinessObjectInfo.EventType)) return;

            ExeActionWithEvent(afterData[_BusinessObjectInfo.EventType], Form, _BusinessObjectInfo);
        }

        public virtual void processRightClickEventbefore(SAPbouiCOM.Form Form, SAPbouiCOM.ContextMenuInfo pVal, ref bool _BubbleEvent)
        {
            if (!blockItem(Form, pVal.ItemUID)) return;

            if (hasContextMenus)
            {
                ContextMenu.TryAddIn(this);
            }

            if (!beforeRightClick.ContainsKey(pVal.ItemUID)) return;

            _BubbleEvent = ExeActionWithEvent(beforeRightClick[pVal.ItemUID], Form, pVal);
        }

        public virtual void processRightClickEventafter(SAPbouiCOM.Form Form, SAPbouiCOM.ContextMenuInfo pVal)
        {
            if (!blockItem(Form, pVal.ItemUID)) return;

            if (hasContextMenus)
            {
                ContextMenu.TryRemoveFrom(this);
            }

            if (!afterRightClick.ContainsKey(pVal.ItemUID)) return;

            ExeActionWithEvent(afterRightClick[pVal.ItemUID], Form, pVal);
        }
        #endregion
    }

    static class UserFormExtension
    {
        public static SAPbouiCOM.Form GetForm(this SAPbouiCOM.Item item)
        {
            if (SAP.SBOApplication.Forms.Count == 0) return null;

            var list = SAP.SBOApplication.Forms.OfType<SAPbouiCOM.Form>()
                                               .Select(f => f)
                                               .Where(f => f.Items.Count > 0 && 
                                                           f.Items.OfType<SAPbouiCOM.Item>()
                                                                  .Where(i => i == item)
                                                                  .Any())
                                               .ToList();
            return list.Count > 0 ? list.First() : null;
        }

        public static IEnumerable<SAPbouiCOM.Form> GetForms(this SAPbouiCOM.Forms forms)
        {
            if (SAP.SBOApplication.Forms.Count == 0) return new SAPbouiCOM.Form[0];

            return SAP.SBOApplication.Forms.OfType<SAPbouiCOM.Form>();
        }
    }
}
