using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace QualityControl.Resources
{
    class Form_SYS_GRPO
    {
        public Form_SYS_GRPO()
        {
         Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //GRPO FormTypeEx 143
                if (pVal.FormTypeEx == "143" && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    SAPbouiCOM.Form oform = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                  
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        // only PaneLevel 1
                        if (oForm.PaneLevel != 1)
                            return;

                        try
                        {
                            oForm.Freeze(true);

                            SAPbouiCOM.Matrix oMatrix =(SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                            for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                            {
                                ProcessChooseFormListAndTab(oMatrix, i);
                            }
                        }
                        finally
                        {
                            oForm.Freeze(false);
                        }
                    }

                    else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && pVal.BeforeAction == false && pVal.CharPressed == 9)
                    {
                        if (pVal.ItemUID == "38" && pVal.Row > 0)
                        {
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                            if (oForm.PaneLevel != 1)
                                return;

                            try
                            {
                                oForm.Freeze(true);

                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                                ProcessChooseFormListAndTab(oMatrix, pVal.Row);
                            }
                            finally
                            {
                                oForm.Freeze(false);
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage("Error in Itemevnt for SAP Screen - " + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
            }
        }

        private void ProcessChooseFormListAndTab(SAPbouiCOM.Matrix oMatrix, int i)
        {
            string PORDocEntry = "";

            SAPbouiCOM.EditText PDocEntry = (SAPbouiCOM.EditText)oMatrix.Columns.Item("45").Cells.Item(i).Specific;
            PORDocEntry = PDocEntry.Value.ToString();

            if (PORDocEntry != "")
            {
                SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rSet1 = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string qStr = "SELECT T0.\"ItemCode\", T2.\"WhsCode\" " + Environment.NewLine +
                             "FROM \"OITM\" T0 " + Environment.NewLine +
                             "INNER JOIN \"@FIL_MR_INPLNIM\" T1 ON T1.\"U_ITEMCODE\" = T0.\"ItemCode\" " + Environment.NewLine +
                             "INNER JOIN \"@FIL_MH_INPLAN\" T4 ON T4.\"Code\" = T1.\"Code\" " + Environment.NewLine +
                             "CROSS JOIN \"OWHS\" T2 " + Environment.NewLine +
                             "WHERE T2.\"WhsCode\" = (select \"WhsCode\" from \"OWHS\" T3 where T3.\"U_WHSTYPE\"='Q') AND T1.\"U_ACTIVE\" = 'Y' AND T4.\"U_ACTIVE\" = 'Y' AND T0.\"ItemCode\" = '" +
                             ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";

                rSet.DoQuery(qStr);

                string qStr1 = $"select \"WhsCode\" from \"POR1\" where \"DocEntry\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("45").Cells.Item(i).Specific).Value + "' AND " +
                               $" \"ItemCode\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "' AND " +
                               $" \"LineNum\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("46").Cells.Item(i).Specific).Value + "';";

                rSet1.DoQuery(qStr1);

                if (rSet1.RecordCount > 0)
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(i).Specific).Value = rSet1.Fields.Item("WhsCode").Value.ToString();

                if (rSet.RecordCount > 0 && rSet.Fields.Item("WhsCode").Value.ToString() != "")
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("24").Cells.Item(i).Specific).Value = rSet.Fields.Item("WhsCode").Value.ToString();
            }
            else
            {
                SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rSet1 = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string qStr = "SELECT T0.\"ItemCode\", T2.\"WhsCode\" " + Environment.NewLine +
                             "FROM \"OITM\" T0 " + Environment.NewLine +
                             "INNER JOIN \"@FIL_MR_INPLNIM\" T1 ON T1.\"U_ITEMCODE\" = T0.\"ItemCode\" " + Environment.NewLine +
                             "INNER JOIN \"@FIL_MH_INPLAN\" T4 ON T4.\"Code\" = T1.\"Code\" " + Environment.NewLine +
                             "CROSS JOIN \"OWHS\" T2 " + Environment.NewLine +
                             "WHERE T2.\"WhsCode\" = (select \"WhsCode\" from \"OWHS\" T3 where T3.\"U_WHSTYPE\"='Q') AND T1.\"U_ACTIVE\" = 'Y' AND T4.\"U_ACTIVE\" = 'Y' AND T0.\"ItemCode\" = '" +
                             ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";

                rSet.DoQuery(qStr);

                if (rSet.RecordCount > 0 && rSet.Fields.Item("WhsCode").Value.ToString() != "")
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("24").Cells.Item(i).Specific).Value = rSet.Fields.Item("WhsCode").Value.ToString();

                string qStr1 = $"select \"DfltWH\" from \"OITM\" where  \"ItemCode\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";

                rSet1.DoQuery(qStr1);

                if (rSet1.RecordCount > 0)
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(i).Specific).Value = rSet1.Fields.Item("DfltWH").Value.ToString();
            }
        }

        //private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        //{
        //    BubbleEvent = true;
        //    try
        //    {
        //        //GRPO FormTypeEx 143
        //        if (pVal.FormTypeEx == "143" && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
        //        {
        //            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

        //            if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && pVal.ActionSuccess)
        //            {
        //                // Assuming 'oform' is your SAPbouiCOM.Form object
        //                //SAPbouiCOM.Item GlnStatic = oform.Items.Item("1470002039");
        //                //SAPbouiCOM.Item GlnEditText = oform.Items.Item("1470002038");



        //            }
        //            else if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
        //            {
        //                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

        //                oForm.Freeze(true);

        //                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oform.Items.Item("38").Specific;

        //                for (int i = 1; i < oMatrix.VisualRowCount; i++)
        //                {
        //                    //string itemcode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("43").Cells.Item(i).Specific).Value;
        //                    string PORDocEntry = "";

        //                    SAPbouiCOM.EditText PDocEntry = (SAPbouiCOM.EditText)oMatrix.Columns.Item("45").Cells.Item(i).Specific;
        //                    PORDocEntry = PDocEntry.Value.ToString();

        //                    if (PORDocEntry != "")
        //                    {
        //                        SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //                        SAPbobsCOM.Recordset rSet1 = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //                        string qStr = "SELECT T0.\"ItemCode\", T2.\"WhsCode\" " + Environment.NewLine +
        //                                     "FROM \"OITM\" T0 " + Environment.NewLine +
        //                                     "INNER JOIN \"@FIL_MR_INPLNIM\" T1 ON T1.\"U_ITEMCODE\" = T0.\"ItemCode\" " + Environment.NewLine + "INNER JOIN \"@FIL_MH_INPLAN\" T4 ON T4.\"Code\" = T1.\"Code\" " + Environment.NewLine + "CROSS JOIN \"OWHS\" T2 " + Environment.NewLine + "WHERE T2.\"WhsCode\" = (select \"WhsCode\" from \"OWHS\" T3 where T3.\"U_WHSTYPE\"='Q') AND T1.\"U_ACTIVE\" = 'Y' AND T4.\"U_ACTIVE\" = 'Y' AND T0.\"ItemCode\" = '" +
        //                                     ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";

        //                        rSet.DoQuery(qStr);

        //                        string qStr1 = $"select \"WhsCode\" from \"POR1\" where \"DocEntry\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("45").Cells.Item(i).Specific).Value + "' AND " +
        //                                       $" \"ItemCode\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "' AND " +
        //                                       $" \"LineNum\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("46").Cells.Item(i).Specific).Value + "';";

        //                        rSet1.DoQuery(qStr1);

        //                        if (rSet1.RecordCount > 0)
        //                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(i).Specific).Value = rSet1.Fields.Item("WhsCode").Value.ToString();
        //                        if (rSet.RecordCount > 0)
        //                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("24").Cells.Item(i).Specific).Value = rSet.Fields.Item("WhsCode").Value.ToString();
        //                    }
        //                    else
        //                    {
        //                        SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //                        SAPbobsCOM.Recordset rSet1 = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //                        string qStr = "SELECT T0.\"ItemCode\", T2.\"WhsCode\" " + Environment.NewLine +
        //                                     "FROM \"OITM\" T0 " + Environment.NewLine +
        //                                     "INNER JOIN \"@FIL_MR_INPLNIM\" T1 ON T1.\"U_ITEMCODE\" = T0.\"ItemCode\" " + Environment.NewLine + "INNER JOIN \"@FIL_MH_INPLAN\" T4 ON T4.\"Code\" = T1.\"Code\" " + Environment.NewLine + "CROSS JOIN \"OWHS\" T2 " + Environment.NewLine + "WHERE T2.\"WhsCode\" = (select \"WhsCode\" from \"OWHS\" T3 where T3.\"U_WHSTYPE\"='Q') AND T1.\"U_ACTIVE\" = 'Y' AND T4.\"U_ACTIVE\" = 'Y' AND T0.\"ItemCode\" = '" +
        //                                     ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";

        //                        rSet.DoQuery(qStr);
        //                        if (rSet.RecordCount > 0)
        //                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("24").Cells.Item(i).Specific).Value = rSet.Fields.Item("WhsCode").Value.ToString();

        //                        string qStr1 = $"select \"DfltWH\" from \"OITM\" where  \"ItemCode\" = '" + ((SAPbouiCOM.EditText)oMatrix.Columns.Item("1").Cells.Item(i).Specific).Value + "';";

        //                        rSet1.DoQuery(qStr1);

        //                        if (rSet1.RecordCount > 0)
        //                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(i).Specific).Value = rSet1.Fields.Item("DfltWH").Value.ToString();

        //                    }
        //                }
        //                oForm.Freeze(false);

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
