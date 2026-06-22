using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QualityControl.Resources
{
    [FormAttribute("QualityControl.Resources.Form_NO_IPndList", "Resources/Form_NO_IPndList.b1f")]
    class Form_NO_IPndList : UserFormBase
    {
        public Form_NO_IPndList()
        {
        }

        private SAPbouiCOM.Button BTSELECT, Cancel;
        private SAPbouiCOM.Grid GRID01;
        public override void OnInitializeComponent()
        {
            this.GRID01 = ((SAPbouiCOM.Grid)(this.GetItem("GRID01").Specific));
            this.BTSELECT = ((SAPbouiCOM.Button)(this.GetItem("BTSELECT").Specific));
            this.BTSELECT.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTSELECT_PressedAfter);
            this.Cancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {

        }

        private void BTSELECT_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
            try
            {
                CheckBoxItem(ref oform);
            }
            catch (Exception EX)
            {

            }

        }

        private static void CheckBoxItem(ref SAPbouiCOM.Form pForm)
        {
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)pForm.Items.Item("GRID01").Specific;
            string DocEntry = "";
            string ObjType = "";
            string BaseDocNum = "";
            string ItemCode = "";
            string ItemName = "";
            decimal Quantity = 0;
            string Status = "";
            string ScrapWhsCode = "";
            string NormalWhsCode = "";
            string Bacth = "";
            string QcWhs = "";
            string rejectWhs = "";

            if (oGrid.DataTable.Rows.Count > 0)
            {
                for (int i = 0; i < oGrid.DataTable.Rows.Count; i++)
                {
                    if (oGrid.DataTable.GetValue("Selected", i).ToString() == "Y")
                    {
                        BaseDocNum = oGrid.DataTable.GetValue("DocNum", i).ToString();
                        DocEntry = oGrid.DataTable.GetValue("DocEntry", i).ToString();
                        ObjType = oGrid.DataTable.GetValue("ObjType", i).ToString();
                        ItemCode = oGrid.DataTable.GetValue("ItemCode", i).ToString();
                        ItemName = oGrid.DataTable.GetValue("ItemName", i).ToString();
                        Quantity = Convert.ToDecimal(oGrid.DataTable.GetValue("Quantity", i).ToString());
                        Status = oGrid.DataTable.GetValue("Status", i).ToString();
                        ScrapWhsCode = oGrid.DataTable.GetValue("ScrapWhsCode", i).ToString();
                        rejectWhs = oGrid.DataTable.GetValue("RejectWhs", i).ToString();
                        NormalWhsCode = oGrid.DataTable.GetValue("NormalWhsCode", i).ToString();
                        //Bacth = oGrid.DataTable.GetValue("BatchNum", i).ToString();
                        QcWhs = oGrid.DataTable.GetValue("QcWarehouse", i).ToString();
                    }
                }
            }
            else
            {
                Global.GFunc.ShowError("No rows in the grid.");
                return;
            }

            pForm.Close();
            Form_UDO_InspectionDecision InspectionDecision = new Form_UDO_InspectionDecision();
            InspectionDecision.Show();
            SAPbouiCOM.Form cForm = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            if (cForm == null)
            {
                Global.GFunc.ShowError("Failed to load the form.");
                return;
            }

            try
            {
                //cForm.SupportedModes = 15;
                if (!string.IsNullOrEmpty(BaseDocNum))
                {
                    Form_UDO_InspectionDecision.LoadQCValue(ref cForm, BaseDocNum, DocEntry, ObjType, ItemCode, ItemName, Quantity, ScrapWhsCode, rejectWhs, NormalWhsCode, Bacth, QcWhs, Status);
                }
                else
                {
                    Global.GFunc.ShowError("Status is empty.");
                }
            }
            catch (Exception ex)
            {
                cForm.Items.Item("2").Click();
                cForm.Freeze(false);

            }
            finally
            {

                cForm.Freeze(false);
            }
        }

        internal static void LoadPendingQC(ref SAPbouiCOM.Form pForm)
        {
            pForm.Freeze(true);
            string FormDate = pForm.DataSources.UserDataSources.Item("UDFD").ValueEx;
            string ToDate = pForm.DataSources.UserDataSources.Item("UDTD").ValueEx;
            string DocType = pForm.DataSources.UserDataSources.Item("UDDOCTYPE").ValueEx;
            string ItemGroup = pForm.DataSources.UserDataSources.Item("UDITMGRPCD").ValueEx;
            string ItemCode = pForm.DataSources.UserDataSources.Item("UDITEMCODE").ValueEx;
            string CardCode = pForm.DataSources.UserDataSources.Item("UDCARDCODE").ValueEx;
            string Status = pForm.DataSources.UserDataSources.Item("UDSTATUS").ValueEx;

            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)pForm.Items.Item("GRID01").Specific;
            oGrid.DataTable.Clear();
            string qStr = "";
            //qStr = "call FIL_SP_QC_CheckList ( '" + FormDate + "','" + ToDate + "', '" + DocType + "','" + ItemGroup + "', '" + ItemCode + "','" + CardCode + "', '" + Status + "')";
            qStr = "call FIL_SP_QC_CheckList ( '" + FormDate + "','" + ToDate + "', '" + DocType + "','" + ItemGroup + "', '" + ItemCode + "','" + CardCode + "')";
            oGrid.DataTable.ExecuteQuery(qStr);
            if (oGrid.DataTable.IsEmpty)
            {
                oGrid.DataTable.Clear();
                Global.GFunc.ShowWarning("No Pending QC");
                pForm.PaneLevel = 1;
                pForm.Freeze(false);
                return;
            }

            SAPbouiCOM.EditTextColumn DocEntry = (SAPbouiCOM.EditTextColumn ) oGrid.Columns.Item("DocEntry");
            DocEntry.LinkedObjectType = DocType;
            oGrid.Columns.Item("Selected").Editable = true;
            oGrid.Columns.Item("DocEntry").Editable = false;
            oGrid.Columns.Item("ObjType").Editable = false;
            oGrid.Columns.Item("DocNum").Editable = false;
            oGrid.Columns.Item("DocDate").Editable = false;
            oGrid.Columns.Item("ItemCode").Editable = false;
            oGrid.Columns.Item("ItemName").Editable = false;
            oGrid.Columns.Item("Quantity").Editable = false;
            oGrid.Columns.Item("QcWarehouse").Editable = false;
            oGrid.Columns.Item("ScrapWhsCode").Editable = false;
            oGrid.Columns.Item("RejectWhs").Editable = false;
            oGrid.Columns.Item("NormalWhsCode").Editable = false;
            oGrid.Columns.Item("Accepted Quantity").Editable = false;
            oGrid.Columns.Item("Rejected Quantity").Editable = false;
            oGrid.Columns.Item("Balance Quantity").Editable = false;
        //    oGrid.Columns.Item("BatchNum").Editable = false;
            oGrid.Columns.Item("Status").Editable = false;




              oGrid.Columns.Item("Selected").TitleObject.Caption = "Check";
            //oGrid.Columns.Item("Status").TitleObject.Caption = "Status";
            //oGrid.Columns.Item("SourceDocument").TitleObject.Caption = "Source Document";
            //oGrid.Columns.Item("DocDate").TitleObject.Caption = "Doc Date";
            //oGrid.Columns.Item("DocNumber").TitleObject.Caption = "Doc Number";
            //oGrid.Columns.Item("QcNumber").TitleObject.Caption = "QC Number";
            //oGrid.Columns.Item("QCDate").TitleObject.Caption = "QC Date";
            //oGrid.Columns.Item("TransferDate").TitleObject.Caption = "Transfer Date";
            //oGrid.Columns.Item("TransferNumber").TitleObject.Caption = "Transfer Number";
            //oGrid.Columns.Item("ItemCode").TitleObject.Caption = "Item Code";
            //oGrid.Columns.Item("ItemDescription").TitleObject.Caption = "Item Description";
            //oGrid.Columns.Item("Batch").TitleObject.Caption = "Batch";
            //oGrid.Columns.Item("ActualQuantity").TitleObject.Caption = "Actual Quantity";
            //oGrid.Columns.Item("PendingQuantity").TitleObject.Caption = "Pending Quantity";
            //oGrid.Columns.Item("QCQuantity").TitleObject.Caption = "QC Quantity";
            //oGrid.Columns.Item("AcceptedQuatity").TitleObject.Caption = "Accepted Quantity";
            //oGrid.Columns.Item("RejectedQuantity").TitleObject.Caption = "Rejected Quantity";

            ////oGrid.Columns.Item("DocDate").TitleObject.Caption = "Doc Date";
            ////oGrid.Columns.Item("DistNumber").TitleObject.Caption = "Batch / PC Num";
            ////oGrid.Columns.Item("OpenQty").TitleObject.Caption = "Open Qty";
            //oGrid.Columns.Item("DocNumber").TitleObject.Sortable = true;
            //oGrid.Columns.Item("DocDate").TitleObject.Sortable = true;
            //oGrid.Columns.Item("Batch").TitleObject.Sortable = true;
            ////oGrid.Columns.Item("DistNumber").TitleObject.Sortable = true;
            ////oGrid.Columns.Item("OpenQty").RightJustified = true;
            oGrid.Columns.Item("Selected").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
            oGrid.AutoResizeColumns();

            pForm.Freeze(false);
        }


    }
}
