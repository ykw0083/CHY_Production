using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FT_ADDON.Addon
{
    class AddOnSettingsJS : AddOnSettings
    {
        public override bool Setup()
        {
            //UserTable udt = new UserTable("SQLQuery", "Query Table");

            //if (!udt.createField("Query", "Query", SAPbobsCOM.BoFieldTypes.db_Memo, 254, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!UserTable.createField("OIGE", "WONum", "WO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, "")) return false;
            if (!UserTable.createField("OJDT", "WONum", "WO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, "")) return false;
            if (!UserTable.createField("OITM", "Spec", "Spec", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) return false;

            if (!UserTable.createField("IGE1", "Weight", "Weight(kg/pc)", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!UserTable.createField("IGE1", "WOIMCost", "WO Input Material Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!UserTable.createField("IGE1", "WOIPCost", "WO Input Process Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!UserTable.createField("IGE1", "WOSPValue", "WO Side Product Value", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!UserTable.createField("IGE1", "WOOPCost", "WO Output Process Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;

            UserTable WOINPUTCOST = new UserTable("WOINPUTCOST", "WO Order Input Cost", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            if (!WOINPUTCOST.createField("WOType", "WO Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!WOINPUTCOST.createField("DateFr", "Date From", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!WOINPUTCOST.createField("DateTo", "Date To", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;

            UserTable WOINPUTCOSTDT = new UserTable("WOINPUTCOSTDT", "WO Order Input Cost DT", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            if (!WOINPUTCOSTDT.createField("Spec", "Spec", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!WOINPUTCOSTDT.createField("ThicknessFrom", "Thickness From", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOINPUTCOSTDT.createField("ThicknessTo", "Thickness To", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOINPUTCOSTDT.createField("LengthFrom", "Length From", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOINPUTCOSTDT.createField("LengthTo", "Length To", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOINPUTCOSTDT.createField("CostPerPC", "Cost Per PC", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOINPUTCOSTDT.createField("CostPerMT", "Cost Per MT", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOINPUTCOSTDT.createField("ExpAct", "Exp Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;

            WOINPUTCOST.Children.Add(WOINPUTCOSTDT);
            if (!WOINPUTCOST.createUDO("", false, false, false, true, false, "")) return false;


            UserTable WOOUTPUTCOST = new UserTable("WOOUTPUTCOST", "WO Order Output Cost", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            if (!WOOUTPUTCOST.createField("WOType", "WO Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!WOOUTPUTCOST.createField("DateFr", "Date From", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!WOOUTPUTCOST.createField("DateTo", "Date To", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;

            UserTable WOOUTPUTCOSTDT = new UserTable("WOOUTPUTCOSTDT", "WO Order Output Cost DT", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            if (!WOOUTPUTCOSTDT.createField("Spec", "Spec", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!WOOUTPUTCOSTDT.createField("ThicknessFrom", "Thickness From", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOOUTPUTCOSTDT.createField("ThicknessTo", "Thickness To", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOOUTPUTCOSTDT.createField("LengthFrom", "Length From", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOOUTPUTCOSTDT.createField("LengthTo", "Length To", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOOUTPUTCOSTDT.createField("CostPerPC", "Cost Per PC", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOOUTPUTCOSTDT.createField("CostPerMT", "Cost Per MT", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOOUTPUTCOSTDT.createField("ExpAct", "Exp Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;

            WOOUTPUTCOST.Children.Add(WOOUTPUTCOSTDT);
            if (!WOOUTPUTCOST.createUDO("", false, false, false, true, false, "")) return false;


            UserTable WOSIDEOUTPUT = new UserTable("WOSIDEOUTPUT", "WO Order Side Cost", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            if (!WOSIDEOUTPUT.createField("WOType", "WO Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!WOSIDEOUTPUT.createField("DateFr", "Date From", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!WOSIDEOUTPUT.createField("DateTo", "Date To", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;

            UserTable WOSIDEOUTPUTDT = new UserTable("WOSIDEOUTPUTDT", "WO Order Side Cost DT", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            if (!WOSIDEOUTPUTDT.createField("Spec", "Spec", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!WOSIDEOUTPUTDT.createField("ThicknessFrom", "Thickness From", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOSIDEOUTPUTDT.createField("ThicknessTo", "Thickness To", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOSIDEOUTPUTDT.createField("LengthFrom", "Length From", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOSIDEOUTPUTDT.createField("LengthTo", "Length To", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOSIDEOUTPUTDT.createField("CostPerPC", "Cost Per PC", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOSIDEOUTPUTDT.createField("CostPerMT", "Cost Per MT", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Rate)) return false;
            if (!WOSIDEOUTPUTDT.createField("ExpAct", "Exp Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;

            WOSIDEOUTPUT.Children.Add(WOSIDEOUTPUTDT);
            if (!WOSIDEOUTPUT.createUDO("", false, false, false, true, false, "")) return false;


            UserTable udt = new UserTable("WOType", "WO Type");
            if (!udt.createField("WIPAct", "WIP Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;

            UserTable udtwo = new UserTable("FTS_WO", "WO Order", SAPbobsCOM.BoUTBTableType.bott_Document);
            if (!udtwo.createField("Ref2", "Ref 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("DocDate", "Posting Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", true, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("DueDate", "Expected Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "", true, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("Issued", "Issued Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", true, SAPbobsCOM.BoFldSubTypes.st_Measurement)) return false;
            if (!udtwo.createField("Received", "Received Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", true, SAPbobsCOM.BoFldSubTypes.st_Measurement)) return false;
            if (!udtwo.createField("ItemCode", "Side Product Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("ItemName", "Side Product Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("Quantity", "Side Product Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) return false;
            if (!udtwo.createField("Weight", "Side Product Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Measurement)) return false;
            if (!udtwo.createField("DistNumber", "Side Product Batch Num.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("MnfSerial", "Side Product Batch Ref.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("WhsCode", "Side Product Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("Amount", "Side Product Amount", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) return false;
            if (!udtwo.createField("SONo", "SO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) return false;
            if (!udtwo.createField("Project", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("SalesType", "Sales Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo.createField("Machine", "Machine Master", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;


            UserTable udtwo1 = new UserTable("FTS_WO1", "Work Order Issued", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            if (!udtwo1.createField("ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) return false;
            if (!udtwo1.createField("Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Measurement)) return false;
            if (!udtwo1.createField("DistNumber", "Batch Num.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("MnfSerial", "Batch Ref.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("WhsCode", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("Amount", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) return false;
            if (!udtwo1.createField("Length", "Length(mm)", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) return false;
            if (!udtwo1.createField("Project", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("SalesType", "Sales Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("ProdType", "Production Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("ProductionCost", "Production Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!udtwo1.createField("ProductionCut", "Production Cut", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!udtwo1.createField("ProductionDrill", "Production Drill", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!udtwo1.createField("ProductionPaint", "Production Paint", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!udtwo1.createField("ProductionBlast", "Production Blast", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!udtwo1.createField("ProductionOther", "Production Others", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!udtwo1.createField("AverageCost", "Average Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Price)) return false;
            if (!udtwo1.createField("TotalCost", "Total Cost", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) return false;
            if (!udtwo1.createField("UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("Machine", "Machine Master", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo1.createField("Qty_PCS", "Qty PCS", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) return false;

            UserTable udtwo2 = new UserTable("FTS_WO2", "Work Order Received", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            if (!udtwo2.createField("ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) return false;
            if (!udtwo2.createField("Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Measurement)) return false;
            if (!udtwo2.createField("DistNumber", "Batch Num.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("MnfSerial", "Batch Ref.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("WhsCode", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("Amount", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) return false;
            if (!udtwo2.createField("Area", "Area", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("Length", "Length(mm)", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) return false;
            if (!udtwo2.createField("Project", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("SalesType", "Sales Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("Batch01", "Batch Attribute 01", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("Batch02", "Batch Attribute 02", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("Machine", "Machine Master", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo2.createField("SONo", "SO Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) return false;
            if (!udtwo2.createField("Qty_PCS", "Qty PCS", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) return false;

            UserTable udtwo3 = new UserTable("FTS_WO3", "Work Order Side Item", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
            if (!udtwo3.createField("ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo3.createField("ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo3.createField("Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) return false;
            if (!udtwo3.createField("Weight", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Measurement)) return false;
            if (!udtwo3.createField("DistNumber", "Batch Num.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo3.createField("MnfSerial", "Batch Ref.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo3.createField("WhsCode", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo3.createField("Amount", "Amount", SAPbobsCOM.BoFieldTypes.db_Float, 0, "", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) return false;
            if (!udtwo3.createField("UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo3.createField("Machine", "Machine Master", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo3.createField("Batch01", "Batch Attribute 01", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;
            if (!udtwo3.createField("Batch02", "Batch Attribute 02", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false, SAPbobsCOM.BoFldSubTypes.st_None)) return false;

            udtwo.Children.Add(udtwo1);
            udtwo.Children.Add(udtwo2);
            udtwo.Children.Add(udtwo3);
            if (!udtwo.createUDO("", true, true, true, false, true, "AFTS_WO")) return false;

            #region Custom Form Setting table
            //UserTable FT_CFS = new UserTable("FT_CFS", "Custom Form Setting", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            //if (!FT_CFS.createField("FNAME", "Form Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", false)) return false;
            //if (!FT_CFS.createField("USRID", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", false)) return false;
            //if (!FT_CFS.createField("MATRIX", "Matrix Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", false)) return false;
            //if (!FT_CFS.createField("DSNAME", "Table Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", false)) return false;
            //if (!FT_CFS.createField("MATRIX", "Matrix Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", false)) return false;
            //if (!FT_CFS.createField("MATRIX", "Matrix Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", false)) return false;
            //if (!FT_CFS.createField("MATRIX", "Matrix Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", false)) return false;
            //if (!FT_CFS.createField("MATRIX", "Matrix Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", false)) return false;

            //UserTable FT_CFSDL = new UserTable("FT_CFSDL", "Custom Form Setting Detail", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            //if (!FT_CFSDL.createField("SEQ", "Sequence", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "1", false)) return false;
            //if (!FT_CFSDL.createField("CNAME", "Column Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "", false)) return false;
            //if (!FT_CFSDL.createField("NONVIEW", "Cannot View", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "1", false)) return false;
            //if (!FT_CFSDL.createField("NONEDIT", "Cannot Edit", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "1", false)) return false;

            //FT_CFS.children.Add(FT_CFSDL);
            //if (!FT_CFS.createUDO("", false, false, false, true, false, "")) return false;

            //UserTable FT_SPCFSQL = new UserTable("FT_SPCFSQL", "Copy From SQL", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            //if (!FT_SPCFSQL.createField("UDO", "UDO", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) return false;
            //if (!FT_SPCFSQL.createField("Header", "Header", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, false, false, "Y:Yes|N:No")) return false;
            //if (!FT_SPCFSQL.createField("HColumn", "Column Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) return false;
            //if (!FT_SPCFSQL.createField("Btn", "Button Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) return false;
            //if (!FT_SPCFSQL.createField("BtnName", "Button Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) return false;
            //if (!FT_SPCFSQL.createField("BtnSQL", "Button Copy from SQL", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "", false)) return false;
            #endregion


            GC.Collect();

            //if (Common.createQuery("Get Discountable Item Code", "SELECT \"ItemCode\" FROM \"OITM\" WHERE \"U_IncDiscount\"='Y'", out int qr))
            //{
            //    Common.createFormattedSearch("UDO_FT_DISTCUST", qr, "0_U_G", "C_0_2");
            //    Common.createFormattedSearch("UDO_FT_DISTGRP", qr, "0_U_G", "C_0_2");
            //}

            return true;
        }
    }
}
