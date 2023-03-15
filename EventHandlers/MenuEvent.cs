using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;

namespace FT_ADDON
{
    class MenuEvent
    {
        public static void processMenuEvent(ref SAPbouiCOM.MenuEvent pVal)
        {
            try
            {
                if (PurchaseOrder_Base.isCustomPurchaseOrder(pVal.MenuUID))
                {
                    InitPOForm.VInventory();
                }
                else if (Form_Base.GetFormTypes(pVal.MenuUID, out var list))
                {
                    foreach (var formtype in list)
                    {
                        Form_Base.OpenNewForm(formtype);
                    }
                }
            }
            catch (Exception ex)
            {
                SAP.stopProgressBar();
                SAP.SBOApplication.MessageBox(Common.ReadException(ex), 1, "OK", "", "");
            }
        }

        public static void processMenuEvent2(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                SAPbouiCOM.Form oForm = null;

                try
                {
                    oForm = SAP.SBOApplication.Forms.ActiveForm;
                }
                catch
                {
                    return;
                }

                if (oForm == null) return;

                if (!Form_Base.GetForms(oForm.UniqueID, out var list)) return;

                if (pVal.BeforeAction)
                {
                    foreach (var formobj in list)
                    {
                        formobj.processMenuEventbefore(oForm, pVal, ref BubbleEvent);
                    }

                    return;
                }

                foreach (var formobj in list)
                {
                    formobj.processMenuEventafter(oForm, pVal);
                }
            }
            catch (Exception ex)
            {
                SAP.stopProgressBar();
                SAP.SBOApplication.MessageBox(Common.ReadException(ex), 1, "OK", "", "");
            }
        }
    }
}
