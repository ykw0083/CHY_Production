using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace FT_ADDON.Addon
{
    //[AutoFillFromList]
    [FormCode("Work_Order")]
    [MenuId("4352")]
    [Series("FTS_WO", "Series")]

    class Form_FTS_WO : Form_Base
    {
        static int matrixCurrRow = 0;
        static string matrixActiveItem = "";

        const string u_salestype = "U_SalesTyp";
        const string u_project = "U_Project";
        const string u_prodtype = "U_ProdType";
        const string dtu_prodtype = "DTU_ProdType";

        const string dttype = "DT_Type";
        const string u_Type = "U_Type";
        const string statusField = "Status";
        const string canceledField = "Canceled";
        const string series = "Series";
        const string docnum = "DocNum";
        const string ud_Status = "UD_Status";
        const string ud_WhIs = "UD_WhIs";
        const string ud_WhRc = "UD_WhRc";
        const string ud_WhSd = "UD_WhSd";
        const string isFolder = "IsFolder";
        const string fTS_WO = "@FTS_WO";
        const string fTS_WO1 = "@FTS_WO1";
        const string fTS_WO2 = "@FTS_WO2";
        const string fTS_WO3 = "@FTS_WO3";
        const string u_DocDate = "U_DocDate";
        const string u_DueDate = "U_DueDate";
        const string u_Ref2 = "U_Ref2";
        const string whsCode = "WhsCode";
        const string itemCode = "ItemCode";
        const string itemName = "ItemName";
        const string u_itemCode = "U_ItemCode";
        const string u_itemName = "U_ItemName";
        const string matrix1 = "grid1";
        const string matrix2 = "grid2";
        const string matrix3 = "grid3";
        const string u_Quantity = "U_Quantity";
        const string u_Weight = "U_Weight";
        const string u_Amount = "U_Amount";
        const string u_Issued = "U_Issued";
        const string u_Received = "U_Received";
        const string u_DistNumb = "U_DistNumb";
        const string u_MnfSeria = "U_MnfSeria";
        const string u_WhsCode = "U_WhsCode";
        const string u_Area = "U_Area";

        const string visorder = "VisOrder";
        const string btnCopy = "btnCopy";
        const string btnSO = "btnSO";
        const string btnPost = "btnPost";
        const string menuAddRecord = "1282";
        const string menuDeleteDetailRow = "1293";
        const string menuCancelRecord = "1284";
        const string menuRestoreRecord = "1285";
        const string menuCloseRecord = "1286";
        const string menuRefreshRecord = "1304";
        const string menuRepeat = "8801";
        const string statusOpen = "OPEN";
        const string statusClosed = "CLOSED";
        const string statusCanceled = "CANCELED";

        static readonly IDictionary<string, string> matrixDsDict = new Dictionary<string, string>()
        {
            { matrix1, fTS_WO1},
            { matrix2, fTS_WO2},
            { matrix3, fTS_WO3}
        };

        string defaultbtn { get => "1"; }

        public Form_FTS_WO()
        {
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE, activeFormAfter);

            AddBeforeItemFunc(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, defaultBtnClick);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, defaultBtnClickAfter);

            AddAfterDataFunc(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, addNewRecord);

            AddAfterDataFunc(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD, loadRecord);

            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, chooseFromListItem);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, chooseFromListWarehouse);

            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, selectionChange);

            AddBeforeItemFunc(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, postBtnClick);

            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS, setMenuOnFocus);

            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS, calculateTotal);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS, setMenuOnLostFocus);

            AddAfterMenuFunc(menuAddRecord, addNewRecord);
            AddBeforeMenuFunc(menuDeleteDetailRow, deleteRow);
            AddBeforeMenuFunc(menuCancelRecord, checkCancelRecord);
            AddBeforeMenuFunc(menuRestoreRecord, checkRestoreRecord);

            AddBeforeItemFunc(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, eventCFL);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, eventCFLSO);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, eventCFLSelectionChanged);

        }
        void activeFormAfter()
        {
            oForm.EnableMenu(menuCloseRecord, false);
            oForm.EnableMenu(menuRepeat, false);
        }
        private void eventCFLSelectionChanged()
        {
            if (currentId == btnCopy)
            {
                var mymatrix = GetMatrix(matrix2);
                if (mymatrix.RowCount == 0) return;
                mymatrix.FlushToDataSource();
                string query = "";
                var btn = GetButtonCombo(currentId);

                if (btn.Selected?.Value == null) return;

                string code = btn.Selected.Value.ToString();
                btn.Caption = "Copy From";

                if (code == "Sales Order")
                    query = $"{oForm.TypeEx}.{currentId}.getSO";
                else if (code == "Reserve Invoice")
                    query = $"{oForm.TypeEx}.{currentId}.getRIV";

                string query_script = SQLQuery.QueryCode(oForm, query);

                var formobj = NewForm<UserForm_CFL>();
                if (formobj == null)
                {
                    throw new Exception("missing form type");
                }
                formobj.afterCFL = funcAfterCFLCopySO;
                formobj.afterExit = funcAfterCFLExit;
                formobj.addRowWithinCFL = funcAddRowSO;
                formobj.query = query_script;
                formobj.FormUID = oForm.UniqueID;
                formobj.dsrow = oForm.DataSources.DBDataSources.Item(fTS_WO2).Size - 1;
                formobj.dsname = fTS_WO2;
                //formobj.dsmatrix = curmatrix;
                //formobj.dsmatrixcolumn = colId;
                formobj.OpenForm();
                formobj.retrieveGrid(true);
                oForm.Freeze(true);
            }

        }
        private void eventCFLSO()
        {
            if (currentId == btnSO)
            {
                string query = $"{oForm.TypeEx}.{currentId}.getData";
                string query_script = SQLQuery.QueryCode(oForm, query);

                var formobj = NewForm<UserForm_CFL>();
                if (formobj == null)
                {
                    throw new Exception("missing form type");
                }
                formobj.afterCFL = funcAfterCFLSO;
                formobj.afterExit = funcAfterCFLExit;
                //formobj.addRowWithinCFL = funcAddRow;
                formobj.query = query_script;
                formobj.FormUID = oForm.UniqueID;
                formobj.dsrow = 0;
                formobj.dsname = fTS_WO;
                //formobj.dsmatrix = curmatrix;
                //formobj.dsmatrixcolumn = colId;
                formobj.OpenForm();
                formobj.retrieveGrid();
                oForm.Freeze(true);

                BubbleEvent = false;
            }
        }
        private void eventCFL()
        {
            if (currentId == matrix1 || currentId == matrix2 || currentId == matrix3)
            {
                if (colId == u_itemCode || colId == u_WhsCode) return;

                string curmatrix = currentId;
                string curdsname = currentId == matrix1 ? fTS_WO1 : fTS_WO2;
                GetMatrix(currentId).FlushToDataSource();

                SAPbouiCOM.Matrix oMatrix = GetMatrix(curmatrix);
                oMatrix.FlushToDataSource();

                string temp = $"{oForm.TypeEx}.{curmatrix}.{colId}.getData";
                string query = temp;
                string query_script = SQLQuery.QueryCode(oForm, query);

                var formobj = NewForm<UserForm_CFL>();
                if (formobj == null)
                {
                    throw new Exception("missing form type");
                }
                string cur_colid = colId;
                int cur_row = currentRow;
                formobj.afterCFL = () => funcAfterCFL(curmatrix, cur_colid, cur_row);
                formobj.afterExit = funcAfterCFLExit;
                //formobj.addRowWithinCFL = funcAddRow;
                formobj.query = query_script;
                formobj.FormUID = oForm.UniqueID;
                formobj.dsrow = currentRow - 1;
                formobj.dsname = curdsname;
                formobj.dsmatrix = curmatrix;
                formobj.dsmatrixcolumn = colId;
                formobj.OpenForm();
                formobj.retrieveGrid();
                oForm.Freeze(true);

                BubbleEvent = false;
            }
        }
        private void funcAfterCFLSO()
        { }
        private void funcAfterCFL(string matrix, string col, int row)
        {
            GetMatrix(matrix).LoadFromDataSource();
            GetMatrix(matrix).Columns.Item(col).Cells.Item(row).Click();
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        }
        private void funcAfterCFLCopySO()
        {
            SAPbouiCOM.Matrix oMatrix = GetMatrix(matrix2);
            oMatrix.LoadFromDataSource();
            funcArrangeGrids(matrix2, fTS_WO2);
            oForm.Items.Item("RcFolder").Click();
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        }
        private void funcAfterCFLExit()
        {
            oForm.Freeze(false);
        }
        private void funcAddRowSO()
        {

            var ds = oForm.DataSources.DBDataSources.Item(fTS_WO2);
            int size = ds.Size;
            if (!string.IsNullOrWhiteSpace(oForm.DataSources.DBDataSources.Item(fTS_WO2).GetValue(u_itemCode, size - 1)))
                ds.InsertRecord(ds.Size - 1);
            //int maxlineid = 0;
            //string value = "";
            //for (int x = 0; x < ds.Size; x++)
            //{
            //    value = ds.GetValue(LineID, x);
            //    if (string.IsNullOrEmpty(value))
            //        ds.SetValue(LineID, x, "0");
            //    if (Convert.ToInt32(ds.GetValue(LineID, x)) > maxlineid)
            //        maxlineid = Convert.ToInt32(ds.GetValue("LineId", x));
            //}
            //ds.InsertRecord(ds.Size - 1);
            //ds.SetValue(U_ItemCode, ds.Size - 1, "");
            //ds.SetValue(U_ItemName, ds.Size - 1, "");
            //ds.SetValue(U_SerialNo, ds.Size - 1, "");
            //ds.SetValue(U_Package, ds.Size - 1, "");
            //ds.SetValue(U_Location, ds.Size - 1, "");
            //ds.SetValue(U_Month, ds.Size - 1, "0");
            //ds.SetValue(U_Freq, ds.Size - 1, "0");
            //ds.SetValue(U_Confirm, ds.Size - 1, "Y");
            //ds.SetValue(LineID, ds.Size - 1, (maxlineid + 1).ToString());
            //funcArrangeGrids($"@{tbl_detail}", oForm);
            ////if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

        }
        void defaultBtnClickAfter()
        {
            if (currentId != defaultbtn) return;
            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE) return;

            try
            {
                SAP.SBOApplication.ActivateMenuItem(menuRefreshRecord);
            }
            catch
            { }
        }

        void setMenuOnFocus()
        {
            matrixCurrRow = currentRow;
            matrixActiveItem = currentId;

            oForm.EnableMenu(menuDeleteDetailRow, false);
            oForm.EnableMenu(menuCancelRecord, oForm.GetUserSourceValue(ud_Status) == statusOpen);
            if (!matrixDsDict.ContainsKey(itemPVal.ItemUID)) return;

            oForm.EnableMenu(menuDeleteDetailRow, GetMatrix(itemPVal.ItemUID).RowCount > 0);
            oForm.EnableMenu(menuCancelRecord, false);
        }
        void setMenuOnLostFocus()
        {
            if (!matrixDsDict.ContainsKey(itemPVal.ItemUID)) return;

            oForm.EnableMenu(menuDeleteDetailRow, false);
        }
        void checkCancelRecord()
        {
            if (oForm.GetUserSourceValue(ud_Status) != statusOpen)
            {
                SAP.SBOApplication.SetStatusBarMessage($"Document is not in {statusOpen} status", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                BubbleEvent = false;
                return; 
            }
            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                SAP.SBOApplication.SetStatusBarMessage("Document is not in OK mode", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                BubbleEvent = false;
                return;
            }
        }
        void checkRestoreRecord()
        {
            if (oForm.GetUserSourceValue(ud_Status) != statusCanceled)
            {
                SAP.SBOApplication.SetStatusBarMessage("Document is not in CANCEL status", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                BubbleEvent = false;
                return; 
            }
            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                SAP.SBOApplication.SetStatusBarMessage("Document is not in OK mode", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                BubbleEvent = false;
                return;
            }
        }
        void deleteRow()
        {
            if (!matrixDsDict.ContainsKey(matrixActiveItem))
            {
                BubbleEvent = false; 
                return;
            }
            SAPbouiCOM.Matrix oMatrix = GetMatrix(matrixActiveItem);
            
            if (oMatrix.RowCount > 0 && matrixCurrRow > 0)
            {
                oMatrix.DeleteRow(matrixCurrRow);
                funcArrangeGrids(matrixActiveItem, matrixDsDict[matrixActiveItem]);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            BubbleEvent = false;
        }
        void funcManageForm()
        {
            bool isoopen = false;
            if (oForm.GetUserSourceValue(ud_Status) == statusOpen) isoopen = true;

            oForm.Items.Item(u_Type).Enabled = isoopen;
            oForm.Items.Item(u_DocDate).Enabled = isoopen;
            oForm.Items.Item(u_DueDate).Enabled = isoopen;
            oForm.Items.Item(u_Ref2).Enabled = isoopen;

            foreach (KeyValuePair<string, string> entry in matrixDsDict)
            {
                SAPbouiCOM.Matrix oMatrix = GetMatrix(entry.Key);

                for (int x = 1; x <= oMatrix.RowCount; x++)
                {
                    oMatrix.CommonSetting.SetRowEditable(x, isoopen);
                }
            }
            oForm.Update();
        }

        //protected override void runtimeTweakBefore()
        //{
        //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //}
        void loadRecord()
        {
            if (oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue(statusField, 0) == "C")
            {
                if (oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue(canceledField, 0) == "Y")
                {
                    oForm.SetUserSourceValue(ud_Status, statusCanceled);
                }
                else
                    oForm.SetUserSourceValue(ud_Status, statusClosed);

                funcManageForm();
                return;
            }
            oForm.SetUserSourceValue(ud_Status, statusOpen);
            funcManageForm();

            foreach (KeyValuePair<string, string> entry in matrixDsDict)
            {
                funcArrangeGrids(entry.Key, entry.Value);
            }
        }
        void defaultBtnClick()
        {
            if (currentId != defaultbtn) return;

            foreach (KeyValuePair<string, string> entry in matrixDsDict)
            {
                GetMatrix(entry.Key).FlushToDataSource();
            }

        }
        protected override void runtimeTweakAfter()
        {
            base.runtimeTweakAfter();

            GetButton(btnSO).Image = $"{System.Windows.Forms.Application.StartupPath}\\Resources\\CFL.bmp";

            GetButtonCombo(btnCopy).ValidValues.Add("Sales Order", "Sales Order");
            GetButtonCombo(btnCopy).ValidValues.Add("Reserve Invoice", "Reserve Invoice");

            addComboValidValues(dttype, u_Type);

            oForm.Items.Item(docnum).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            oForm.Items.Item(series).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            oForm.Items.Item(u_Issued).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            oForm.Items.Item(u_Received).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

            foreach (KeyValuePair<string, string> entry in matrixDsDict)
            {
                SAPbouiCOM.Matrix oMatrix = GetMatrix(entry.Key);

                try { oMatrix.Columns.Item(u_Amount).Editable = false; } catch { }
                try { oMatrix.Columns.Item(visorder).Editable = false; } catch { }

                if (entry.Key == matrix3)
                {
                    try { oMatrix.Columns.Item(u_Quantity).Editable = false; } catch { }
                    try { oMatrix.Columns.Item(u_Weight).Editable = false; } catch { }
                }
                if (entry.Key == matrix1)
                {
                    var dt = oForm.DataSources.DataTables.Item(dtu_prodtype);
                    var col = oMatrix.Columns.Item(u_prodtype);
                    for (int x = 0; x < dt.Rows.Count; x++)
                    {
                        col.ValidValues.Add(dt.GetValue("Code", x).ToString(), dt.GetValue("Name", x).ToString());
                    }
                }
            }
            oForm.Update();

            addNewRecord();
        }
        void funcCalculateTotal(string gridname, string dbsource, int row)
        {
            var wo = oForm.DataSources.DBDataSources.Item(dbsource);
            var grid = GetMatrix(gridname);
            grid.FlushToDataSource();

            decimal total = 0;

            for (int x = 0; x < wo.Size; x++)
            {
                if (string.IsNullOrWhiteSpace(wo.GetValue(u_itemCode, x))) continue;

                total += decimal.Parse(wo.GetValue(u_Weight, x));
            }

            grid.LoadFromDataSource();

            if (gridname == matrix1)
                oForm.DataSources.DBDataSources.Item(fTS_WO).SetValue(u_Issued, 0, total.ToString());
            if (gridname == matrix2)
                oForm.DataSources.DBDataSources.Item(fTS_WO).SetValue(u_Received, 0, total.ToString());

        }
        void calculateTotal()
        {
            if (!(currentId == matrix1 || currentId == matrix2 || currentId == matrix3)) return;
            if (!(colId == u_Weight || colId == u_Quantity)) return;

            if (currentId == matrix1)
                funcCalculateTotal(matrix1, fTS_WO1, currentRow);
            if (currentId == matrix2)
                funcCalculateTotal(matrix2, fTS_WO2, currentRow);

        }
        void funcArrangeGrids(string gridname, string dbsource)
        {
            var wo = oForm.DataSources.DBDataSources.Item(dbsource);

            var grid = GetMatrix(gridname);

            grid.FlushToDataSource();

            if (wo.Size == 0)
            {
                wo.InsertRecord(0);
                grid.LoadFromDataSource();
            }

            if (grid.RowCount == 0)
            {
                grid.AddRow();
                ((SAPbouiCOM.EditText)grid.Columns.Item(visorder).Cells.Item(1).Specific).Value = "1";
                grid.FlushToDataSource();
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(wo.GetValue(u_itemCode, wo.Size - 1)))
                {
                    wo.InsertRecord(wo.Size);
                }
                for (int x = 0; x < wo.Size; x++)
                {
                    wo.SetValue(visorder, x, (x + 1).ToString());
                }
                grid.LoadFromDataSource();
            }
        }
        void postBtnClick()
        {
            if (itemPVal.ItemUID != btnPost) return;

            if (oForm.GetUserSourceValue(ud_Status) != statusOpen)
            {
                SAP.SBOApplication.SetStatusBarMessage("Document is " + oForm.GetUserSourceValue(ud_Status), SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }
            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                SAP.SBOApplication.SetStatusBarMessage("Please save the document 1st.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }
            int result = SAP.SBOApplication.MessageBox("Post Work Order is a irreversable process, are you sure want to continue?", 2, "Yes", "No");
            if (result == 2) return;

            try
            {
                SAP.SBOCompany.StartTransaction();

                string docEntry = oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue("DocEntry", 0);
                funcCheckWO(docEntry);

                funcCreateGoodIssue(docEntry);
                funcCreateGoodReceive(docEntry);
                funcCreateJournalEntry(docEntry);

                funcCloseWO(docEntry);

                if (SAP.SBOCompany.InTransaction)
                    SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                SAP.SBOApplication.ActivateMenuItem(menuRefreshRecord);
                SAP.SBOApplication.StatusBar.SetSystemMessage("Posting Done", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.StatusBar.SetSystemMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                if (SAP.SBOCompany.InTransaction)
                    SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            }
        }
        void funcCheckWO(string docEntry)
        {
            string spName = "FTS_sp_CheckWO";
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery($"exec {spName} {docEntry}");
            if (rs.RecordCount == 1)
            {
                rs.MoveFirst();
                int errcode = (int)rs.Fields.Item(0).Value;
                if (errcode == -1)
                {
                    string msg = rs.Fields.Item(1).Value.ToString();
                    throw new Exception($"{msg} from {spName}");
                }
            }

        }
        void funcCreateGoodIssue(string docEntry)
        {
            string spName = "FTS_sp_GetWOIssue";
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery($"exec {spName} {docEntry}");
            if (rs.RecordCount == 0)
            {
                throw new Exception($"No record from {spName}");
            }
            rs.MoveFirst();

            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
            oDoc.DocDate = DateTime.Parse(rs.Fields.Item("DocDate").Value.ToString());
            oDoc.Reference2 = rs.Fields.Item("Ref2").Value.ToString();
            oDoc.Comments = rs.Fields.Item("Comments").Value.ToString();
            oDoc.JournalMemo = rs.Fields.Item("JrnlMemo").Value.ToString();
            oDoc.UserFields.Fields.Item("U_WONum").Value = rs.Fields.Item("U_WONum").Value.ToString();

            int cnt = 0;
            while (!rs.EoF)
            {
                if (cnt > 0)
                {
                    oDoc.Lines.Add();
                    oDoc.Lines.SetCurrentLine(oDoc.Lines.Count - 1);
                }

                oDoc.Lines.ItemCode = (string)rs.Fields.Item("ItemCode").Value;
                oDoc.Lines.WarehouseCode = (string)rs.Fields.Item("WhsCode").Value;
                oDoc.Lines.Quantity = (double)rs.Fields.Item("Quantity").Value;
                oDoc.Lines.AccountCode = (string)rs.Fields.Item("Account").Value;
                oDoc.Lines.UserFields.Fields.Item("U_Weight").Value = rs.Fields.Item("U_Weight").Value;
                oDoc.Lines.UserFields.Fields.Item("U_WOIPCost").Value = rs.Fields.Item("U_WOIPCost").Value;

                if (!string.IsNullOrWhiteSpace((string)rs.Fields.Item("DistNumber").Value))
                {
                    oDoc.Lines.BatchNumbers.BatchNumber = rs.Fields.Item("DistNumber").Value.ToString();
                    if (!string.IsNullOrWhiteSpace((string)rs.Fields.Item("MnfSerial").Value))
                        oDoc.Lines.BatchNumbers.ManufacturerSerialNumber = rs.Fields.Item("MnfSerial").Value.ToString();
                    oDoc.Lines.BatchNumbers.Quantity = double.Parse(rs.Fields.Item("Quantity").Value.ToString());
                }
                cnt++;
                rs.MoveNext();
            }

            if (oDoc.Add() != 0)
            {
                throw new Exception(SAP.SBOCompany.GetLastErrorDescription());
            }
        }
        void funcCreateGoodReceive(string docEntry)
        {
            string spName = "FTS_sp_GetWOReceive";
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery($"exec {spName} {docEntry}");
            if (rs.RecordCount == 0)
            {
                throw new Exception($"No record from {spName}");
            }
            rs.MoveFirst();

            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
            oDoc.DocDate = DateTime.Parse(rs.Fields.Item("DocDate").Value.ToString());
            oDoc.Reference2 = rs.Fields.Item("Ref2").Value.ToString();
            oDoc.Comments = rs.Fields.Item("Comments").Value.ToString();
            oDoc.JournalMemo = rs.Fields.Item("JrnlMemo").Value.ToString();
            oDoc.UserFields.Fields.Item("U_WONum").Value = rs.Fields.Item("U_WONum").Value.ToString();

            int cnt = 0;
            while (!rs.EoF)
            {
                if (cnt > 0)
                {
                    oDoc.Lines.Add();
                    oDoc.Lines.SetCurrentLine(oDoc.Lines.Count - 1);
                }

                oDoc.Lines.ItemCode = (string)rs.Fields.Item("ItemCode").Value;
                oDoc.Lines.WarehouseCode = (string)rs.Fields.Item("WhsCode").Value;
                oDoc.Lines.Quantity = (double)rs.Fields.Item("Quantity").Value;
                oDoc.Lines.AccountCode = (string)rs.Fields.Item("Account").Value;
                oDoc.Lines.UserFields.Fields.Item("U_Weight").Value = rs.Fields.Item("U_Weight").Value;
                oDoc.Lines.UserFields.Fields.Item("U_WOIPCost").Value = rs.Fields.Item("U_WOIPCost").Value;
                oDoc.Lines.LineTotal = (double)rs.Fields.Item("Amount").Value;
                oDoc.Lines.UserFields.Fields.Item("U_WOIMCost").Value = rs.Fields.Item("U_WOIMCost").Value;
                oDoc.Lines.UserFields.Fields.Item("U_WOSPValue").Value = rs.Fields.Item("U_WOSPCost").Value;
                oDoc.Lines.UserFields.Fields.Item("U_WOOPCost").Value = rs.Fields.Item("U_WOOPCost").Value;

                if (!string.IsNullOrWhiteSpace((string)rs.Fields.Item("DistNumber").Value))
                {
                    oDoc.Lines.BatchNumbers.BatchNumber = rs.Fields.Item("DistNumber").Value.ToString();
                    if (!string.IsNullOrWhiteSpace((string)rs.Fields.Item("MnfSerial").Value))
                        oDoc.Lines.BatchNumbers.ManufacturerSerialNumber = rs.Fields.Item("MnfSerial").Value.ToString();
                    oDoc.Lines.BatchNumbers.Quantity = double.Parse(rs.Fields.Item("Quantity").Value.ToString());
                }
                cnt++;
                rs.MoveNext();
            }

            if (oDoc.Add() != 0)
            {
                throw new Exception(SAP.SBOCompany.GetLastErrorDescription());
            }
        }
        void funcCreateJournalEntry(string docEntry)
        {
            string spName = "FTS_sp_GetWOJournal";
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery($"exec {spName} {docEntry}");
            if (rs.RecordCount == 0)
            {
                return;
                //throw new Exception($"No record from {spName}");
            }
            rs.MoveFirst();

            SAPbobsCOM.JournalEntries oDoc = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            oDoc.ReferenceDate = DateTime.Parse(rs.Fields.Item("RefDate").Value.ToString());
            oDoc.Memo = rs.Fields.Item("Memo").Value.ToString();
            oDoc.Reference = rs.Fields.Item("Ref1").Value.ToString();
            oDoc.Reference2 = rs.Fields.Item("Ref2").Value.ToString();
            oDoc.UserFields.Fields.Item("U_WONum").Value = rs.Fields.Item("U_WONum").Value.ToString();

            int cnt = 0;
            while (!rs.EoF)
            {
                if (cnt > 0)
                {
                    oDoc.Lines.Add();
                    oDoc.Lines.SetCurrentLine(oDoc.Lines.Count - 1);
                }

                oDoc.Lines.AccountCode = (string)rs.Fields.Item("Account").Value;
                if ((double)rs.Fields.Item("Amount").Value > 0)
                    oDoc.Lines.Debit = (double)rs.Fields.Item("Amount").Value;
                else
                    oDoc.Lines.Credit = (double)rs.Fields.Item("Amount").Value * -1;
                oDoc.Lines.LineMemo =(string) rs.Fields.Item("LineMemo").Value;
                oDoc.Lines.CostingCode = (string)rs.Fields.Item("ProfitCode").Value;

                cnt++;
                rs.MoveNext();
            }

            if (oDoc.Add() != 0)
            {
                throw new Exception(SAP.SBOCompany.GetLastErrorDescription());
            }
        }
        void funcCloseWO(string docEntry)
        {
            string spName = "FTS_sp_CloseWO";
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery($"exec {spName} {docEntry}");
            if (rs.RecordCount == 1)
            {
                rs.MoveFirst();
                int errcode = (int)rs.Fields.Item(0).Value;
                if (errcode == -1)
                {
                    string msg = rs.Fields.Item(1).Value.ToString();
                    throw new Exception($"{msg} from {spName}");
                }
            }

        }

        void selectionChange()
        {
            //currentId;
            //currentRow;
            //colId;
            

            if (currentId == "U_Type")
            {
                foreach (KeyValuePair<string, string> entry in matrixDsDict)
                {
                    if (GetMatrix(entry.Key).RowCount == 0) funcArrangeGrids(entry.Key, entry.Value);
                }
                funcManageForm();
            }
        }

        private void chooseFromListWarehouse()
        {
            if (itemPVal.ColUID != u_WhsCode) return;
            var dt = ((SAPbouiCOM.IChooseFromListEvent)itemPVal).SelectedObjects;

            if (dt == null) return;

            if (!(currentId == matrix1 || currentId == matrix2 || currentId == matrix3)) return;

            var item = oForm.Items.Item(currentId);
            string code = "";
            string tablename = null;
            string alias = null;
            SAPbouiCOM.Matrix grid = null;
            switch (item.Type)
            {
                case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                    grid = ((SAPbouiCOM.Matrix)item.Specific);
                    var txt = grid.Columns.Item(itemPVal.ColUID).Cells.Item(itemPVal.Row).Specific as SAPbouiCOM.EditText;
                    code = dt.GetValue(whsCode, 0).ToString();
                    tablename = txt.DataBind.TableName;
                    alias = u_itemCode;
                    break;
                default:
                    return;
            }

            if (alias == null) return;

            if (grid == null) return;

            if (tablename == null)
            {
                if (alias == String.Empty || !oForm.HasUserSource(alias)) return;

                oForm.SetUserSourceValue(alias, code);
            }
            else if (oForm.HasDataSource(tablename))
            {
                grid.FlushToDataSource();
                oForm.DataSources.DBDataSources.Item(tablename).SetValue(u_WhsCode, itemPVal.Row - 1, code);
                grid.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

            }

        }

        private void chooseFromListItem()
        {
            if (itemPVal.ColUID != u_itemCode) return;
            var dt = ((SAPbouiCOM.IChooseFromListEvent)itemPVal).SelectedObjects;

            if (dt == null) return;

            if (!(currentId == matrix1 || currentId == matrix2 || currentId == matrix3)) return;
            
            var item = oForm.Items.Item(currentId);
            string code = "";
            string name = "";
            string tablename = null;
            string alias = null;
            SAPbouiCOM.Matrix grid = null;
            switch (item.Type)
            {
                case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                    grid = ((SAPbouiCOM.Matrix)item.Specific);
                    var txt = grid.Columns.Item(itemPVal.ColUID).Cells.Item(itemPVal.Row).Specific as SAPbouiCOM.EditText;
                    code = dt.GetValue(itemCode, 0).ToString();
                    name = dt.GetValue(itemName, 0).ToString();
                    tablename = txt.DataBind.TableName;
                    alias = u_itemCode;
                    break;
                default:
                    return;
            }

            if (alias == null) return;

            if (grid == null) return;

            if (tablename == null)
            {
                if (alias == String.Empty || !oForm.HasUserSource(alias)) return;

                oForm.SetUserSourceValue(alias, code);
            }
            else if (oForm.HasDataSource(tablename))
            {
                grid.FlushToDataSource();
                oForm.DataSources.DBDataSources.Item(tablename).SetValue(u_itemCode, itemPVal.Row - 1, code);
                oForm.DataSources.DBDataSources.Item(tablename).SetValue(u_itemName, itemPVal.Row - 1, name);
                string wh = "";
                if (currentId == matrix1)
                    wh = oForm.GetUserSourceValue(ud_WhIs);
                if (currentId == matrix2)
                    wh = oForm.GetUserSourceValue(ud_WhRc);
                if (currentId == matrix3)
                    wh = oForm.GetUserSourceValue(ud_WhSd);

                oForm.DataSources.DBDataSources.Item(tablename).SetValue(u_WhsCode, itemPVal.Row - 1, wh);

                grid.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                funcArrangeGrids(currentId, tablename);
                funcManageForm();
            }

        }

        void addComboValidValues(string datatablename, string comboname)
        {
            var oCombo = this.GetCombo(comboname);
            SAPbouiCOM.DataTable udt = oForm.DataSources.DataTables.Item(datatablename);
            for (int y = 0; y < udt.Rows.Count; y++)
            {
                oCombo.ValidValues.Add(udt.GetValue(0, y).ToString(), udt.GetValue(1, y).ToString());
            }
        }

        void addNewRecord()
        {
            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            Common.setFormDefaultValue(oForm);
            oForm.SetUserSourceValue(ud_Status, statusOpen);

            funcManageForm();

            if (!string.IsNullOrWhiteSpace(oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue(u_Type, 0)))
            {
                foreach (KeyValuePair<string, string> entry in matrixDsDict)
                {
                    funcArrangeGrids(entry.Key, entry.Value);
                }
            }

            //selectDefaultSeries();
            string temp = oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue(u_DocDate, 0);

            if (string.IsNullOrWhiteSpace(temp))
                oForm.DataSources.DBDataSources.Item(fTS_WO).SetValue(u_DocDate, 0, DateTime.Today.ToString("yyyyMMdd"));
            temp = oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue(u_DueDate, 0);
            if (string.IsNullOrWhiteSpace(temp))
                oForm.DataSources.DBDataSources.Item(fTS_WO).SetValue(u_DueDate, 0, DateTime.Today.ToString("yyyyMMdd"));

            oForm.Items.Item(isFolder).Click();

        }



    }
}
