using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace QualityControl.Resources
{
    class Form_SYS_ReceiptFromProduction
    {
        public Form_SYS_ReceiptFromProduction()
        {
            Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //GRPO FormTypeEx 143
                if (pVal.FormTypeEx == "65214" && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    SAPbouiCOM.Form oform = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.ActionSuccess)
                    {
                        // Assuming 'oform' is your SAPbouiCOM.Form object
                        //SAPbouiCOM.Item GlnStatic = oform.Items.Item("1470002039");
                        //SAPbouiCOM.Item GlnEditText = oform.Items.Item("1470002038");



                    }
                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
                        try
                        {
                            oForm.Freeze(true);

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oform.Items.Item("13").Specific;

                            for (int i = 1; i < oMatrix.VisualRowCount; i++)
                            {
                                // string itemcode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value;
                                string WORDocEntry = "";

                                SAPbouiCOM.EditText WDocEntry = (SAPbouiCOM.EditText)oMatrix.Columns.Item("61").Cells.Item(i).Specific;
                                WORDocEntry = WDocEntry.Value.ToString();
                                if(WORDocEntry != "")
                                {
                                    SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    SAPbobsCOM.Recordset rSet1 = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    string qStr = "SELECT T0.\"ItemCode\", T2.\"WhsCode\" " + Environment.NewLine +
                                                 "FROM \"OITM\" T0 " + Environment.NewLine +
                                                 "INNER JOIN \"@FIL_MR_INPLNIM\" T1 ON T1.\"U_ITEMCODE\" = T0.\"ItemCode\" " + Environment.NewLine + "INNER JOIN \"@FIL_MH_INPLAN\" T4 ON T4.\"Code\" = T1.\"Code\" " + Environment.NewLine +
                                                 "CROSS JOIN \"OWHS\" T2 " + Environment.NewLine +
                                                 "WHERE T2.\"WhsCode\" = (select \"WhsCode\" from \"OWHS\" T3 where T3.\"U_WHSTYPE\"='Q') AND T1.\"U_ACTIVE\" = 'Y' AND T4.\"U_ACTIVE\" = 'Y' AND T0.\"ItemCode\" = '" +
                                                 ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";

                                    rSet.DoQuery(qStr);

                                    string qStr1 = $"select \"Warehouse\" from \"OWOR\" where \"DocNum\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("61").Cells.Item(i).Specific).Value + "' AND " +
                                                  $" \"ItemCode\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";


                                    rSet1.DoQuery(qStr1);

                                    if (rSet.RecordCount > 0 && rSet.Fields.Item("WhsCode").Value.ToString() != "")
                                    {
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(i).Specific).Value = rSet1.Fields.Item("Warehouse").Value.ToString();
                                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(i).Specific).Value = rSet.Fields.Item("WhsCode").Value.ToString();
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            oForm.Freeze(false);
                        }
                        oForm.Freeze(false);

                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage("Error in Itemevnt for SAP Screen - " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
        }
    }
}
