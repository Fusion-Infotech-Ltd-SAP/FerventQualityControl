using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QualityControl.Resources
{
    [FormAttribute("QualityControl.Resources.TransferPendingList", "Resources/TransferPendingList.b1f")]
    class TransferPendingList : UserFormBase
    {
        public TransferPendingList()
        {
        }

        private SAPbouiCOM.Grid GRID01;
        public override void OnInitializeComponent()
        {
            this.GRID01 = ((SAPbouiCOM.Grid)(this.GetItem("GRID01").Specific));
            this.GRID01.LinkPressedBefore += new SAPbouiCOM._IGridEvents_LinkPressedBeforeEventHandler(this.GRID01_LinkPressedBefore);
            this.OnCustomInitialize();

        }


        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {

        }


        public void loadPendingList()
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_MH_PNDLIST");
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GRID01").Specific;
            oGrid.DataTable.Clear();

            string sql = @"
                        SELECT 
                            ""DocEntry"" AS ""Qc Entry"",
                            ""DocNum"" AS ""Qc Number"",
                            ""U_DOCDATE"" AS ""Qc Date"",
                            ""U_TRNSFRID"" AS ""Inventory Entry"",
                            ""U_TRNSFRNM"" AS ""Transfer No"",
                            ""U_ITEMCODE"" AS ""Item Code"",
                            ""U_ITEMNAME"" AS ""Item Name"",
                            ""U_ACPTQTY"" AS ""Accepted Qty"",
                            ""U_SCRPQTY"" AS ""Rejected Quantity""
                        FROM ""@FIL_DH_INSPDECN""
                        WHERE ""U_TRNSFRID"" IS NULL
                        ORDER BY ""DocEntry""";


            oGrid.DataTable.ExecuteQuery(sql);

            if (oGrid.DataTable.IsEmpty)
            {
                oGrid.DataTable.Clear();
                Global.GFunc.ShowWarning("No Pending Invertory List");
                oForm.PaneLevel = 1;
                oForm.Freeze(false);
                return;
            }

            SAPbouiCOM.EditTextColumn oDocEntryCol = (SAPbouiCOM.EditTextColumn)oGrid.Columns.Item("Qc Entry");
            oDocEntryCol.LinkedObjectType = "FIL_D_INSPDECN";
            //oGrid.Columns.Item("Qc Entry").Editable = false;
            //oGrid.Columns.Item("Qc Number").Editable = false;
            //oGrid.Columns.Item("Qc Date").Editable = false;
            //oGrid.Columns.Item("Inventory Entry").Editable = false;
            //oGrid.Columns.Item("Transfer No").Editable = false;
            //oGrid.Columns.Item("Item Code").Editable = false;
            //oGrid.Columns.Item("Item Name").Editable = false;
            //oGrid.Columns.Item("Accepted Qty").Editable = false;
            //oGrid.Columns.Item("Rejected Quantity").Editable = false;

            oGrid.AutoResizeColumns();
        }

        private void GRID01_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.ColUID != "Qc Entry")
                    return;

                BubbleEvent = false; 

                SAPbouiCOM.Form oListForm =  Application.SBO_Application.Forms.Item(pVal.FormUID);

                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oListForm.Items.Item("GRID01").Specific;

                string docEntry = oGrid.DataTable.GetValue("Qc Entry", pVal.Row).ToString().Trim();

                if (string.IsNullOrEmpty(docEntry))
                    return;

                OpenInspectionDecisionForm(docEntry);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText( ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void OpenInspectionDecisionForm(string docEntry)
        {
            SAPbouiCOM.Form oForm = null;

            try
            {
                
                 Application.SBO_Application.ActivateMenuItem("FIL_INSPDECN");
                  oForm = Application.SBO_Application.Forms.ActiveForm;
 
                oForm.Freeze(true);

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                ((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCENTRY").Specific).Value = docEntry;

                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
              
            }
        }
    }
}
