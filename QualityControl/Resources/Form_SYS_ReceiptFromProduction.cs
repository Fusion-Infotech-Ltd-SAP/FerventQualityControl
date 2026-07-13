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

        private bool _receiptProcessing = false;
        private HashSet<string> _receiptActivatedDone = new HashSet<string>();

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                // Receipt from Production FormTypeEx = 65214
                if (pVal.FormTypeEx != "65214" || pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                    return;

                if (_receiptProcessing)
                    return;

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("13").Specific;

                if (pVal.ItemUID == "13" && pVal.ColUID == "15" && pVal.Row > 0 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.BeforeAction == false)
                {
                    try
                    {
                        oForm.Freeze(true);

                        SAPbouiCOM.IChooseFromListEvent oCFLEvent = (SAPbouiCOM.IChooseFromListEvent)pVal;

                        SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;

                        if (oDataTable == null || oDataTable.Rows.Count == 0)
                            return;

                        string selectedValue = oDataTable.GetValue("WhsCode", 0).ToString();

                        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(pVal.Row).Specific).Value = selectedValue;

                    }
                   
                    finally
                    {
                        oForm.Freeze(false);
                    }
                    return;
                }


                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
                {
                    if (_receiptActivatedDone.Contains(FormUID))
                        return;

                    try
                    {
                        _receiptProcessing = true;
                        oForm.Freeze(true);

                        for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                        {
                            ProcessReceiptFromProductionRow(oMatrix, i);
                        }
                        _receiptActivatedDone.Add(FormUID);
                    }
                    finally
                    {
                        oForm.Freeze(false);
                        _receiptProcessing = false;
                    }

                    return;
                }
            }
            catch (Exception ex)
            {
                _receiptProcessing = false;
               // Global.GFunc.ShowError( "Error in ItemEvent for Receipt from Production - " + ex.ToString());
            }
        }

        private void ProcessReceiptFromProductionRow(SAPbouiCOM.Matrix oMatrix, int row)
        {
            string itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(row).Specific).Value.Trim();

            string worDocNum = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("61").Cells.Item(row).Specific).Value.Trim();

            if (string.IsNullOrEmpty(itemCode) || string.IsNullOrEmpty(worDocNum))
                return;

            SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            SAPbobsCOM.Recordset rSet1 =(SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string qStr =
                "SELECT T0.\"ItemCode\", T2.\"WhsCode\" " + Environment.NewLine +
                "FROM \"OITM\" T0 " + Environment.NewLine +
                "INNER JOIN \"@FIL_MR_INPLNIM\" T1 ON T1.\"U_ITEMCODE\" = T0.\"ItemCode\" " + Environment.NewLine +
                "INNER JOIN \"@FIL_MH_INPLAN\" T4 ON T4.\"Code\" = T1.\"Code\" " + Environment.NewLine +
                "CROSS JOIN \"OWHS\" T2 " + Environment.NewLine +
                "WHERE T2.\"WhsCode\" = (SELECT \"WhsCode\" FROM \"OWHS\" T3 WHERE T3.\"U_WHSTYPE\" = 'Q') " +
                "AND T1.\"U_ACTIVE\" = 'Y' " +
                "AND T4.\"U_ACTIVE\" = 'Y' " +
                "AND T0.\"ItemCode\" = '" + itemCode.Replace("'", "''") + "';";

            rSet.DoQuery(qStr);

            string qStr1 =
                "SELECT \"Warehouse\" FROM \"OWOR\" " +
                "WHERE \"DocNum\" = '" + worDocNum.Replace("'", "''") + "' " +
                "AND \"ItemCode\" = '" + itemCode.Replace("'", "''") + "';";

            rSet1.DoQuery(qStr1);

            if (rSet1.RecordCount > 0)
            {
                string productionWhs = rSet1.Fields.Item("Warehouse").Value.ToString();

                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(row).Specific).Value = productionWhs;
            }

            if (rSet.RecordCount > 0 && rSet.Fields.Item("WhsCode").Value.ToString() != "")
            {
                string qcWhs = rSet.Fields.Item("WhsCode").Value.ToString();

                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(row).Specific).Value = qcWhs;
            }
        }
        //private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        //{
        //    BubbleEvent = true;
        //    try
        //    {
        //        //GRPO FormTypeEx 143
        //        if (pVal.FormTypeEx == "65214" && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
        //        {
        //            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);


        //            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
        //            {
        //                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);
        //                try
        //                {
        //                    oForm.Freeze(true);

        //                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oform.Items.Item("13").Specific;

        //                    for (int i = 1; i < oMatrix.VisualRowCount; i++)
        //                    {
        //                        // string itemcode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value;
        //                        string WORDocEntry = "";

        //                        SAPbouiCOM.EditText WDocEntry = (SAPbouiCOM.EditText)oMatrix.Columns.Item("61").Cells.Item(i).Specific;
        //                        WORDocEntry = WDocEntry.Value.ToString();
        //                        if(WORDocEntry != "")
        //                        {
        //                            SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //                            SAPbobsCOM.Recordset rSet1 = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //                            string qStr = "SELECT T0.\"ItemCode\", T2.\"WhsCode\" " + Environment.NewLine +
        //                                         "FROM \"OITM\" T0 " + Environment.NewLine +
        //                                         "INNER JOIN \"@FIL_MR_INPLNIM\" T1 ON T1.\"U_ITEMCODE\" = T0.\"ItemCode\" " + Environment.NewLine + "INNER JOIN \"@FIL_MH_INPLAN\" T4 ON T4.\"Code\" = T1.\"Code\" " + Environment.NewLine +
        //                                         "CROSS JOIN \"OWHS\" T2 " + Environment.NewLine +
        //                                         "WHERE T2.\"WhsCode\" = (select \"WhsCode\" from \"OWHS\" T3 where T3.\"U_WHSTYPE\"='Q') AND T1.\"U_ACTIVE\" = 'Y' AND T4.\"U_ACTIVE\" = 'Y' AND T0.\"ItemCode\" = '" +
        //                                         ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";

        //                            rSet.DoQuery(qStr);

        //                            string qStr1 = $"select \"Warehouse\" from \"OWOR\" where \"DocNum\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("61").Cells.Item(i).Specific).Value + "' AND " +
        //                                          $" \"ItemCode\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";


        //                            rSet1.DoQuery(qStr1);

        //                            if (rSet.RecordCount > 0 && rSet.Fields.Item("WhsCode").Value.ToString() != "")
        //                            {
        //                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(i).Specific).Value = rSet1.Fields.Item("Warehouse").Value.ToString();
        //                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("15").Cells.Item(i).Specific).Value = rSet.Fields.Item("WhsCode").Value.ToString();
        //                            }
        //                        }
        //                    }
        //                }
        //                catch (Exception ex)
        //                {
        //                    oForm.Freeze(false);
        //                }

        //                finally
        //                {
        //                    oform.Freeze(false);

        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Application.SBO_Application.SetStatusBarMessage("Error in Itemevnt for SAP Screen - " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
        //    }
        //}
    }
}
