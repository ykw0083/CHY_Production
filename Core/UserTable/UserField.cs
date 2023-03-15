using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FT_ADDON
{
    public class UserField
    {
        public string tablename
        {
            get => oUserFieldsMD.TableName;
            set => oUserFieldsMD.TableName = value;
        }
        private int fieldid
        {
            get
            {
                using (RecordSet rc = new RecordSet())
                {
                    rc.DoQuery($"SELECT \"FieldID\" from \"CUFD\" where \"TableID\"='{ tablename }' AND \"AliasID\"='{ fieldname }'");

                    if (rc.RecordCount == 0) return -1;

                    return Convert.ToInt32(rc.GetValue("FieldID").ToString());
                }
            }
        }
        public string fieldname
        {
            get => oUserFieldsMD.Name;
            set => oUserFieldsMD.Name = value;
        }
        public string fieldinfo
        {
            get => oUserFieldsMD.Description;
            set => UpdateFieldInfo(value);
        }
        public SAPbobsCOM.BoFieldTypes fieldtype
        {
            get => oUserFieldsMD.Type;
            set => UpdateFieldType(value);
        }
        public int fieldsize
        {
            get => oUserFieldsMD.Size;
            set => UpdateSize(value);
        }
        public string defaultvalue
        {
            get => oUserFieldsMD.DefaultValue;
            set => UpdateDefaultValue(value);
        }
        public bool mandatory
        {
            get => oUserFieldsMD.Mandatory == SAPbobsCOM.BoYesNoEnum.tYES;
            set => UpdateMandatory(value);
        }
        public SAPbobsCOM.BoFldSubTypes subtype
        {
            get => oUserFieldsMD.SubType;
            set => UpdateSubType(value);
        }
        public string validvalues
        {
            set => UpdateValidValues(value);
        }
        public string linkedtable
        {
            get => oUserFieldsMD.LinkedTable;
            set => UpdateLinkedTable(value);
        }
        public SAPbobsCOM.UDFLinkedSystemObjectTypesEnum systable
        {
            get => oUserFieldsMD.LinkedSystemObject;
            set => UpdateSystemTable(value);
        }
        public bool canfind { get; set; }
        public bool canmodify { get; set; }

        SAPbobsCOM.UserFieldsMD oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
        bool change = false;
        bool? exists = null;

        public UserField(string tablename, string fieldname)
        {
            this.tablename = tablename;
            this.fieldname = fieldname;

            if (!Exists()) return;

            if (oUserFieldsMD.GetByKey(tablename, fieldid)) return;

            throw new Exception(SAP.SBOCompany.GetLastErrorDescription());
        }

        ~UserField()
        {
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oUserFieldsMD);
            oUserFieldsMD = null;
            GC.Collect();
        }

        public bool Create()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (oUserFieldsMD.Add() == 0) return true;

            SAP.SBOApplication.MessageBox($"Error : { SAP.SBOCompany.GetLastErrorDescription() }", 1, "&Ok", "", "");
            return false;
        }

        public bool Update()
        {
            if (!change) return true;

            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (oUserFieldsMD.Update() == 0) return true;

            SAP.SBOApplication.MessageBox($"Error : { SAP.SBOCompany.GetLastErrorDescription() }", 1, "Ok", "", "");
            return false;
        }

        public bool Exists()
        {
            if (exists.HasValue) return exists.Value;

            using (RecordSet rc = new RecordSet())
            {
                rc.DoQuery($"SELECT \"AliasID\" FROM \"CUFD\" WHERE \"TableID\"='{ tablename }' AND \"AliasID\" = '{ fieldname }'");
                exists = rc.RecordCount > 0;
                return exists.Value;
            }
        }

        private void UpdateFieldInfo(string value)
        {
            if (oUserFieldsMD.Description == value) return;

            oUserFieldsMD.Description = value;
            change = true;
        }

        private void UpdateFieldType(SAPbobsCOM.BoFieldTypes value)
        {
            if (oUserFieldsMD.Type == value) return;

            oUserFieldsMD.Type = value;
            change = true;
        }

        private void UpdateSubType(SAPbobsCOM.BoFldSubTypes value)
        {
            if (oUserFieldsMD.SubType == value) return;

            oUserFieldsMD.SubType = value;
            change = true;
        }

        private void UpdateSize(int value)
        {
            int limit = GetLimit();
            int size = Math.Min(limit, value);

            if (oUserFieldsMD.Size < size)
            {
                oUserFieldsMD.Size = size;
                change = true;
            }

            if (oUserFieldsMD.EditSize < size)
            {
                oUserFieldsMD.EditSize = size;
                change = true;
            }
        }

        private void UpdateDefaultValue(string value)
        {
            if (oUserFieldsMD.DefaultValue == value) return;

            oUserFieldsMD.DefaultValue = value;
            change = true;
        }

        private void UpdateValidValues(string value)
        {
            if (String.IsNullOrEmpty(value)) return;

            int IvalidValues = 0;
            int initialCount = oUserFieldsMD.ValidValues.Count;

            foreach (string vv in value.Split('|'))
            {
                IvalidValues++;
                string[] parm = vv.Split(':');
                bool isNew = false;

                if (IvalidValues > initialCount)
                {
                    isNew = true;
                    oUserFieldsMD.ValidValues.Add();
                    change = true;
                }

                oUserFieldsMD.ValidValues.SetCurrentLine(IvalidValues - 1);

                if (isNew)
                {
                    oUserFieldsMD.ValidValues.Value = parm[0];
                    oUserFieldsMD.ValidValues.Description = parm[1];
                    change = true;
                    continue;
                }

                if (oUserFieldsMD.ValidValues.Value != parm[0])
                {
                    oUserFieldsMD.ValidValues.Value = parm[0];
                    change = true;
                }

                if (oUserFieldsMD.ValidValues.Description != parm[1])
                {
                    oUserFieldsMD.ValidValues.Description = parm[1];
                    change = true;
                }
            }
        }

        private void UpdateMandatory(bool value)
        {
            var sapmondary = value ? SAPbobsCOM.BoYesNoEnum.tYES : SAPbobsCOM.BoYesNoEnum.tNO;

            if (oUserFieldsMD.Mandatory == sapmondary) return;

            oUserFieldsMD.Mandatory = sapmondary;
            change = true;
        }

        private void UpdateLinkedTable(string value)
        {
            bool localchange = change;
            try
            {
                using (RecordSet rc = new RecordSet())
                {
                    rc.DoQuery($"SELECT * FROM \"OUDO\" WHERE \"Code\"='{ value }'");

                    if (rc.RecordCount == 0)
                    {
                        if (oUserFieldsMD.LinkedTable == value) return;

                        oUserFieldsMD.LinkedSystemObject = 0;
                        oUserFieldsMD.LinkedTable = value;
                        oUserFieldsMD.LinkedUDO = "";
                        change = true;
                        return;
                    }

                    if (oUserFieldsMD.LinkedUDO == value) return;

                    oUserFieldsMD.LinkedSystemObject = 0;
                    oUserFieldsMD.LinkedTable = "";
                    oUserFieldsMD.LinkedUDO = value;
                    change = true;
                }
            }
            catch { change = localchange; }
            
        }

        private void UpdateSystemTable(SAPbobsCOM.UDFLinkedSystemObjectTypesEnum value)
        {
            if (oUserFieldsMD.LinkedSystemObject == value) return;

            oUserFieldsMD.LinkedSystemObject = systable;
            oUserFieldsMD.LinkedTable = "";
            oUserFieldsMD.LinkedUDO = "";
            change = true;
        }

        private int GetLimit()
        {
            switch (fieldtype)
            {
                case SAPbobsCOM.BoFieldTypes.db_Numeric:
                    return 11;
                case SAPbobsCOM.BoFieldTypes.db_Memo:
                case SAPbobsCOM.BoFieldTypes.db_Date:
                    return 0;
                case SAPbobsCOM.BoFieldTypes.db_Float:
                    return 16;
                default:
                    return 254;
            }
        }
    }
}
