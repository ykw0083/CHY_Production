using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace FT_ADDON.Addon
{
    //[AutoFillFromList]
    //[FormCode("UDO_FT_FTS_WO")]
    [FormCode("999999")]
    //[MenuId("4352")]
    [Series("FTS_WO", "Series")]
    class UDO_FTS_WO : Form_Base
    {
        const string dttype = "DT_Type";
        const string u_Type = "U_Type";
        const string status = "Status";
        const string canceled = "Canceled";
        const string series = "Series";
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
        string defaultbtn { get => "1"; }
        string visorder { get => "VisOrder"; }
        
        static bool formstartup = false;

        public UDO_FTS_WO()
        {
            AddBeforeItemFunc(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, defaultBtnClick);

            AddAfterDataFunc(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, addNewRecord);
            AddAfterDataFunc(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD, loadRecord);

            //addAfterItemFunc(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, chooseFromListItem);
            //addAfterItemFunc(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST, chooseFromListWarehouse);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT, selectionChange);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, postBtnClick);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS, calculateTotal);

            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_FORM_LOAD, formLoad);
            AddAfterItemFunc(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED, udoFormLoadDone);

            AddAfterMenuFunc("1282", addNewRecord);
            AddAfterMenuFunc("1293", deleteRecord);
        }
        void deleteRecord()
        {

        }
        void funcManageForm()
        {
            bool isoopen = false;
            if (oForm.DataSources.UserDataSources.Item(ud_Status).Value == "OPEN") isoopen = true;

            oForm.Items.Item(u_Type).Enabled = isoopen;
            oForm.Items.Item(u_DocDate).Enabled = isoopen;
            oForm.Items.Item(u_DueDate).Enabled = isoopen;
            oForm.Items.Item(u_Ref2).Enabled = isoopen;

            GetMatrix(matrix1).Columns.Item(u_itemCode).Editable = isoopen;
            GetMatrix(matrix1).Columns.Item(u_itemName).Editable = isoopen;
            GetMatrix(matrix1).Columns.Item(u_Quantity).Editable = isoopen;
            GetMatrix(matrix1).Columns.Item(u_Weight).Editable = isoopen;
            GetMatrix(matrix1).Columns.Item(u_Amount).Editable = false;
            GetMatrix(matrix1).Columns.Item(u_DistNumb).Editable = isoopen;
            GetMatrix(matrix1).Columns.Item(u_MnfSeria).Editable = isoopen;
            GetMatrix(matrix1).Columns.Item(u_WhsCode).Editable = isoopen;

            GetMatrix(matrix2).Columns.Item(u_itemCode).Editable = isoopen;
            GetMatrix(matrix2).Columns.Item(u_itemName).Editable = isoopen;
            GetMatrix(matrix2).Columns.Item(u_Quantity).Editable = isoopen;
            GetMatrix(matrix2).Columns.Item(u_Weight).Editable = isoopen;
            GetMatrix(matrix2).Columns.Item(u_Amount).Editable = false;
            GetMatrix(matrix2).Columns.Item(u_DistNumb).Editable = isoopen;
            GetMatrix(matrix2).Columns.Item(u_MnfSeria).Editable = isoopen;
            GetMatrix(matrix2).Columns.Item(u_WhsCode).Editable = isoopen;
            GetMatrix(matrix2).Columns.Item(u_Area).Editable = isoopen;

            GetMatrix(matrix3).Columns.Item(u_itemCode).Editable = isoopen;
            GetMatrix(matrix3).Columns.Item(u_itemName).Editable = isoopen;
            GetMatrix(matrix3).Columns.Item(u_Quantity).Editable = isoopen;
            GetMatrix(matrix3).Columns.Item(u_Weight).Editable = isoopen;
            GetMatrix(matrix3).Columns.Item(u_Amount).Editable = false;
            GetMatrix(matrix3).Columns.Item(u_DistNumb).Editable = isoopen;
            GetMatrix(matrix3).Columns.Item(u_MnfSeria).Editable = isoopen;
            GetMatrix(matrix3).Columns.Item(u_WhsCode).Editable = isoopen;

        }

        //protected override void runtimeTweakBefore()
        //{
        //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
        //}
        void loadRecord()
        {
            if (oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue(status, 0) == "C")
            {
                if (oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue(canceled, 0) == "Y")
                {
                    oForm.DataSources.UserDataSources.Item(ud_Status).Value = "CANCELED";
                }
                else
                    oForm.DataSources.UserDataSources.Item(ud_Status).Value = "CLOSED";

                return;
            }
            oForm.DataSources.UserDataSources.Item(ud_Status).Value = "OPEN";

            funcArrangeGrids(matrix1, fTS_WO1);
            funcArrangeGrids(matrix2, fTS_WO2);
            funcArrangeGrids(matrix3, fTS_WO3);

            funcManageForm();
        }
        void defaultBtnClick()
        {
            if (currentId != defaultbtn) return;
            GetMatrix(matrix1).FlushToDataSource();
            GetMatrix(matrix2).FlushToDataSource();
            GetMatrix(matrix3).FlushToDataSource();

        }
        void udoFormLoadDone()
        {
            if (!formstartup) return;

            if (currentId != "0_U_FD") return;

            formstartup = false;
            //createSeries();
            addComboValidValues(dttype, u_Type);

            oForm.Items.Item(series).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            //oForm.Items.Item(u_Issued).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            //oForm.Items.Item(u_Received).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            addNewRecord();
        }
        void formLoad()
        {
            formstartup = true;
        }
        protected override void runtimeTweakAfter()
        {
            base.runtimeTweakAfter();

            formLoad();
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

                //if (row - 1 == x)
                //{
                //    decimal amount = decimal.Parse(wo.GetValue(u_Quantity, x)) * decimal.Parse(wo.GetValue(u_Weight, x));
                //    wo.SetValue(u_Amount, x, amount.ToString());
                //}

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
            if (currentId == matrix3)
                funcCalculateTotal(matrix3, fTS_WO3, currentRow);

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
                grid.AddRow();
            else
            {
                //if (!string.IsNullOrWhiteSpace(((SAPbouiCOM.EditText)grid.Columns.Item(u_itemCode).Cells.Item(grid.RowCount).Specific).Value))
                if (!string.IsNullOrWhiteSpace(wo.GetValue(u_itemCode, wo.Size - 1)))
                {
                    wo.InsertRecord(wo.Size);
                    grid.LoadFromDataSource();
                    //grid.AddRow(1, grid.RowCount + 1);
                }
            }
            for (int x = 1; x <= grid.RowCount; x++)
            {
                ((SAPbouiCOM.EditText)grid.Columns.Item(visorder).Cells.Item(x).Specific).Value = x.ToString();
            }
            grid.FlushToDataSource();
        }

        void postBtnClick()
        {
            if (itemPVal.ItemUID != "btnPost") return;

            if (oForm.DataSources.UserDataSources.Item(ud_Status).Value != "OPEN")
            {
                SAP.SBOApplication.SetStatusBarMessage("Document is " + oForm.DataSources.UserDataSources.Item(ud_Status).Value, SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }
            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                SAP.SBOApplication.SetStatusBarMessage("Please save the document 1st.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                return;
            }

            try
            {
                SAP.SBOCompany.StartTransaction();

                string docEntry = oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue("DocEntry", 0);
                funcCheckWO(docEntry);
                funcCreateGoodIssue(docEntry);
                funcCreateGoodReceive(docEntry);

                funcCloseWO(docEntry);

                if (SAP.SBOCompany.InTransaction)
                    SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

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
            oDoc.Reference2 = rs.Fields.Item("DocDate").Value.ToString();
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
            oDoc.Reference2 = rs.Fields.Item("DocDate").Value.ToString();
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
                oDoc.Lines.UserFields.Fields.Item("U_WOSPCost").Value = rs.Fields.Item("U_WOSPCost").Value;
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
                throw new Exception($"No record from {spName}");
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
                if (decimal.Parse((string)rs.Fields.Item("Amount").Value) > 0)
                    oDoc.Lines.Debit = (double)rs.Fields.Item("Amount").Value;
                else
                    oDoc.Lines.Credit = (double)rs.Fields.Item("Amount").Value * -1;
                oDoc.Lines.LineMemo = (string)rs.Fields.Item("LineMemo").Value;
                oDoc.Lines.CostingCode = (string)rs.Fields.Item("ProfitCode").Value;

                cnt++;
                rs.MoveNext();
            }

            if (oDoc.Add() != 0)
            {
                throw new Exception(SAP.SBOCompany.GetLastErrorDescription());
            }
        }

        void selectionChange()
        {
            //currentId;
            //currentRow;
            //colId;
            

            if (currentId == "U_Type")
            {
                funcArrangeGrids(matrix1, fTS_WO1);
                funcArrangeGrids(matrix2, fTS_WO2);
                funcArrangeGrids(matrix3, fTS_WO3);
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
                    wh = oForm.DataSources.UserDataSources.Item(ud_WhIs).Value;
                if (currentId == matrix2)
                    wh = oForm.DataSources.UserDataSources.Item(ud_WhRc).Value;
                if (currentId == matrix3)
                    wh = oForm.DataSources.UserDataSources.Item(ud_WhSd).Value;

                oForm.DataSources.DBDataSources.Item(tablename).SetValue(u_WhsCode, itemPVal.Row - 1, wh);

                grid.LoadFromDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                funcArrangeGrids(currentId, tablename);
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
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
            Common.setFormDefaultValue(oForm);
            oForm.SetUserSourceValue(ud_Status, "OPEN");
            //selectDefaultSeries();
            string temp = oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue(u_DocDate, 0);

            if (string.IsNullOrWhiteSpace(temp))
                oForm.DataSources.DBDataSources.Item(fTS_WO).SetValue(u_DocDate, 0, DateTime.Today.ToString("yyyyMMdd"));
            temp = oForm.DataSources.DBDataSources.Item(fTS_WO).GetValue(u_DueDate, 0);
            if (string.IsNullOrWhiteSpace(temp))
                oForm.DataSources.DBDataSources.Item(fTS_WO).SetValue(u_DueDate, 0, DateTime.Today.ToString("yyyyMMdd"));

            //oForm.Items.Item(isFolder).Click();
            funcManageForm();

        }



    }
}
