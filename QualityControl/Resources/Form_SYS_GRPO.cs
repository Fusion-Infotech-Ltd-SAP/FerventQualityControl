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
        bool ValidatioStart = false;
        bool CFLStart = false;
        public Form_SYS_GRPO()
        {
            Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }
        private HashSet<string> _activatedDone = new HashSet<string>();

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.FormTypeEx == "143" && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    if (ValidatioStart == true || CFLStart == true)
                        return;


                    if (pVal.FormTypeEx == "143" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && pVal.BeforeAction == false && pVal.ActionSuccess == true && pVal.ItemUID == "38" && pVal.ColUID == "24" && pVal.Row > 0)
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (oForm.PaneLevel != 1)
                            return;

                        try
                        {
                            oForm.Freeze(true);

                            SAPbouiCOM.IChooseFromListEvent oCFLEvent = (SAPbouiCOM.IChooseFromListEvent)pVal;

                            SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;

                            if (oDataTable == null || oDataTable.Rows.Count == 0)
                                return;

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                            string selectedValue = oDataTable.GetValue("WhsCode", 0).ToString();

                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(pVal.Row).Specific).Value = selectedValue;

                        }
                        finally
                        {
                            oForm.Freeze(false);
                        }
                    }

                    if (pVal.ItemUID == "38" && pVal.ColUID == "1" && pVal.Row > 0 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE && pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (oForm.PaneLevel != 1)
                            return;

                        try
                        {
                            oForm.Freeze(true);
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;
                            ValidatioStart = true;
                            ProcessChooseFormListAndTab(oMatrix, pVal.Row);
                        }
                        finally
                        {
                            ValidatioStart = false;
                            oForm.Freeze(false);
                        }

                        return;
                    }


                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && pVal.BeforeAction == false)
                    {
                        string formKey = pVal.FormUID;

                        if (_activatedDone.Contains(formKey))
                            return;

                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                        if (oForm.PaneLevel != 1)
                            return;

                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("38").Specific;

                        if (!IsCopiedFromPO(oMatrix))
                            return;

                        try
                        {
                            oForm.Freeze(true);
                            ValidatioStart = true;

                            for (int row = 1; row <= oMatrix.RowCount; row++)
                            {
                                if (row > oMatrix.VisualRowCount)
                                    continue;

                                ProcessChooseFormListAndTab(oMatrix, row);
                            }

                            _activatedDone.Add(formKey);
                        }
                        finally
                        {
                            ValidatioStart = false;
                            oForm.Freeze(false);
                        }

                        return;
                    }

                }
            }
            catch (Exception ex)
            {
                ValidatioStart = false;
                Global.GFunc.ShowError("Error in Itemevent for SAP Screen - " + ex.ToString());
            }
        }

        private bool IsCopiedFromPO(SAPbouiCOM.Matrix oMatrix)
        {
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                string baseType = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("43")
                    .Cells.Item(i).Specific).Value.Trim();

                if (baseType == "22")
                    return true;
            }

            return false;
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

    }
}
