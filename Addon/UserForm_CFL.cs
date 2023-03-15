using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace FT_ADDON.Addon
{
    //[AutoFillFromList]
    [FormCode("UserForm_CFL")]
    //[MenuId("4352")]
    //[Series("FTS_WO", "Series")]
    class UserForm_CFL : Form_Base
    {
        public Action afterCFL;
        public Action afterExit;
        public Action addRowWithinCFL;
        public string query { get; set; }
        public string FormUID { get; set; }
        public int dsrow { get; set; }
        public string dsname { get; set; }
        public string dsmatrix { get; set; }
        public string dsmatrixcolumn { get; set; }

        private const string btn_choose = "choose";
        private const string grid = "grid";
        public UserForm_CFL()
        {
            AddBeforeItemFunc(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, eventChoose);
            AddBeforeItemFunc(SAPbouiCOM.BoEventTypes.et_CLICK, eventGridClick);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE, eventClose);
        }
        void eventClose()
        {
            if (afterExit != null) afterExit();
        }
        void eventGridClick()
        {
            if (currentId != grid) return;
            SAPbouiCOM.Grid oGrid = GetGrid(grid);

            if (oGrid.SelectionMode == SAPbouiCOM.BoMatrixSelect.ms_Auto)
            {
                if (oGrid.Rows.IsSelected(currentRow))
                    oGrid.Rows.SelectedRows.Remove(currentRow);
                else
                    oGrid.Rows.SelectedRows.Add(currentRow);
                BubbleEvent = false;
            }
            else
                oGrid.Rows.SelectedRows.Add(currentRow);

        }
        void eventChoose()
        {
            if (currentId != btn_choose) return;
            var oSFrom = SAP.SBOApplication.Forms.Item(FormUID);
            try
            {
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(grid).Specific;
                if (oGrid.Rows.SelectedRows.Count <= 0) return;

                string column = "";
                string value = "";
                bool change = false;
                foreach (int x in oGrid.Rows.SelectedRows)
                {
                    var ds = oSFrom.DataSources.DBDataSources.Item(dsname);
                    if (oGrid.SelectionMode == SAPbouiCOM.BoMatrixSelect.ms_Auto)
                        if (dsrow >= ds.Size - 1) 
                            if (addRowWithinCFL != null) addRowWithinCFL();
                    //oSFrom.oForm.DataSources.DBDataSources.Item(dsname).InsertRecord(oSFrom.oForm.DataSources.DBDataSources.Item(dsname).Size - 1);
                    for (int y = 0; y < oForm.DataSources.DataTables.Item("cfl").Columns.Count; y++)
                    {
                        column = oForm.DataSources.DataTables.Item("cfl").Columns.Item(y).Name;


                        if (oForm.DataSources.DataTables.Item("cfl").GetValue(y, x) != null)
                        {
                            value = oForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString();

                            if (oSFrom.DataSources.DBDataSources.Item(dsname).Fields.Cast<SAPbouiCOM.Field>().Where(pp => pp.Name == column).FirstOrDefault() != null)
                            {
                                oSFrom.DataSources.DBDataSources.Item(dsname).SetValue(column, dsrow, value);
                                change = true;
                            }
                        }
                    }
                    dsrow++;
                }

                if (change)
                {
                    if (afterCFL!= null) afterCFL();
                }
            }
            finally
            {
                oForm.Close();
            }
        }
        public override void initialize(string menuID)
        {
            if (!hasForm) return;

            bool done = false;

            try
            {
                initializing = true;
                oForm = SAP.SBOApplication.Forms.Add($"FT_{ SAP.getNewformUID() }", SAPbouiCOM.BoFormTypes.ft_Sizable);

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
        protected override void runtimeTweakAfter()
        {
            oForm.Width = 600;
            oForm.Height = 500;

            //oForm.DataSources.UserDataSources.Add("FormUID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            //oForm.DataSources.UserDataSources.Add("col", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            //oForm.DataSources.UserDataSources.Add("row", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            //oForm.DataSources.UserDataSources.Add("matrixname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);

            SAPbouiCOM.Button oButton = null;
            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.Item oItem = null;

            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 75;
            oItem.Width = 65;
            oItem.Top = oForm.Height - 60;
            oItem.Height = 20;
            oButton = (SAPbouiCOM.Button)oItem.Specific;
            oButton.Caption = "Cancel";

            oItem = oForm.Items.Add(btn_choose, SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = oForm.Height - 60;
            oItem.Height = 20;
            oButton = (SAPbouiCOM.Button)oItem.Specific;
            oButton.Caption = "Choose";

            oItem = oForm.Items.Add(grid, SAPbouiCOM.BoFormItemTypes.it_GRID);
            oItem.Left = 5;
            oItem.Width = oForm.Width - 25;
            oItem.Top = 5;
            oItem.Height = oForm.Height - 90;

            oGrid = (SAPbouiCOM.Grid)oItem.Specific;
            oForm.DataSources.DataTables.Add("cfl");
        }

        public void retrieveGrid(bool multi = false)
        {
            if (afterCFL == null)
            {
                SAP.SBOApplication.StatusBar.SetText("afterCFL function is not assigned.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                oForm.Close();
                return;
            }
            SAP.SBOApplication.StatusBar.SetText("Initialize CFL window...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            oForm.DataSources.DataTables.Item("cfl").ExecuteQuery(query);

            SAPbouiCOM.Item oItem = oForm.Items.Item(grid);
            SAPbouiCOM.Grid oGrid = oItem.Specific as SAPbouiCOM.Grid;
            oGrid.DataTable = oForm.DataSources.DataTables.Item("cfl");
            foreach (SAPbouiCOM.GridColumn column in oGrid.Columns)
            {
                column.Editable = false;
            }
            if (multi)
            {
                //oItem = oForm.Items.Add("ALL", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                //((SAPbouiCOM.Button)oItem.Specific).Caption = "Choose All";
                //oItem.Left = oForm.Width - 95;
                //oItem.Width = 65;
                //oItem.Top = oForm.Height - 60;
                //oItem.Height = 20;

                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            }
            else
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

            SAP.SBOApplication.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
        }
    }
}