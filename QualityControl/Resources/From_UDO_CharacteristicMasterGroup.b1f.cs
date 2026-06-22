using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QualityControl.Resources
{
    [FormAttribute("QualityControl.Resources.From_UDO_CharacteristicMasterGroup", "Resources/From_UDO_CharacteristicMasterGroup.b1f")]
    class From_UDO_CharacteristicMasterGroup : UserFormBase
    {

        private SAPbouiCOM.Button Btn, Cancel;
        private SAPbouiCOM.StaticText STCODE, STNAME;
        private SAPbouiCOM.EditText ETCODE, ETNAME, ETDOCENTRY;
        private SAPbouiCOM.Matrix MTX_01;
        

        public From_UDO_CharacteristicMasterGroup()
        {
        }

        public override void OnInitializeComponent()
        {
            this.ETCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCODE").Specific));
            this.ETNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETNAME").Specific));
            this.STCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STCODE").Specific));
            this.STNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STNAME").Specific));
            this.MTX_01 = ((SAPbouiCOM.Matrix)(this.GetItem("MTX_01").Specific));
            this.MTX_01.ComboSelectAfter += new SAPbouiCOM._IMatrixEvents_ComboSelectAfterEventHandler(this.MTX_01_ComboSelectAfter);
            this.Btn = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Btn.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Btn_PressedAfter);
            this.Btn.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Btn_PressedBefore);
            this.Cancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.ETDOCENTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCENTRY").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_0").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_1").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_2").Specific));
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.RightClickBefore += new SAPbouiCOM.Framework.FormBase.RightClickBeforeHandler(this.Form_RightClickBefore);

        }

        private void OnCustomInitialize()
        {

        }

        private void MTX_01_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_MH_CHARMGRPS");
            SAPbouiCOM.DBDataSource oDBL = (SAPbouiCOM.DBDataSource)oform.DataSources.DBDataSources.Item("@FIL_MR_CHARMGRP");

            try
            {
                if (pVal.InnerEvent == false && pVal.Row > 0 && pVal.ColUID == "CLCHARCODE" && pVal.ItemUID == "MTX_01")
                {
                    oform.Freeze(true);
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oform.Items.Item("MTX_01").Specific;
                    int rowIndex = pVal.Row;

                    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("CLCHARCODE").Cells.Item(rowIndex).Specific;
                    string selectedValue = "";

                    if (oCombo.Selected != null)
                        selectedValue = oCombo.Selected.Value;

                    string qStr = $"SELECT \"Code\", CASE WHEN TRIM(IFNULL(\"Name\", '')) = '' THEN \"Code\" ELSE \"Name\" END AS \"Name\" " +
                                  $"FROM \"@FIL_MH_QINSPMAS\" WHERE \"Code\" = '{selectedValue}'";

                    SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rSet.DoQuery(qStr);


                    SAPbouiCOM.EditText Name = (SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCHARNAME").Cells.Item(pVal.Row).Specific;
                    if (rSet.RecordCount > 0)
                    {
                        //Code.Value = rSet.Fields.Item("Code").Value.ToString();
                        Name.Value = rSet.Fields.Item("Name").Value.ToString();
                    }


                    oform.Freeze(false);

                    for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                    {
                        if (((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("CLCHARCODE").Cells.Item(pVal.Row).Specific).Value != "" && ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("CLCHARCODE").Cells.Item(pVal.Row).Specific).Value == ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("CLCHARCODE").Cells.Item(i).Specific).Value && i != pVal.Row)
                        {
                            //BubbleEvent = false;
                            Global.GFunc.ShowError("Parameter already selected in row " + i.ToString());
                            ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("CLCHARCODE").Cells.Item(pVal.Row).Specific).Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCHARNAME").Cells.Item(pVal.Row).Specific).Value = "";

                            oMatrix.SetCellFocus(pVal.Row, 1);
                        }

                    }

                    if (oCombo.Selected != null)
                    {

                        Global.GFunc.SetNewLine(oMatrix, oDBL, pVal.Row, "CLCHARCODE");

                    }
                }

            }
            catch (Exception ex)
            {
                oform.Freeze(false);
            }

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);

            SAPbouiCOM.DBDataSource DBDataSourceLine = (SAPbouiCOM.DBDataSource)oform.DataSources.DBDataSources.Item("@FIL_MR_CHARMGRP");
            oform.Items.Item("ETCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            oform.Freeze(true);
            SAPbouiCOM.Matrix MTX_01 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX_01").Specific;
            Global.GFunc.SetNewLine(MTX_01, DBDataSourceLine, 1, "");
            oform.Freeze(false);

        }

        private void Btn_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oform.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ValidateForm(ref oform, ref BubbleEvent);
                ClearBlankRow(ref oform);
            }

        }

      

        private void Btn_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.DBDataSource DBDataSourceLine = (SAPbouiCOM.DBDataSource)oform.DataSources.DBDataSources.Item("@FIL_MR_CHARMGRP");
            if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                oform.Freeze(true);
                SAPbouiCOM.Matrix MTX_01 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX_01").Specific;
                Menu objMenu = new Menu();
                objMenu.SetNextCode(oform, "ETCODE", "FIL_MH_CHARMGRP");
                Global.GFunc.SetNewLine(MTX_01, DBDataSourceLine, 1, "");
                oform.Freeze(false);
  
            }

        }


        private bool ValidateForm(ref SAPbouiCOM.Form pForm, ref bool BubbleEvent)
        {
            string Code = pForm.DataSources.DBDataSources.Item("@FIL_MH_CHARMGRP").GetValue("Code", 0);
            string Name = pForm.DataSources.DBDataSources.Item("@FIL_MH_CHARMGRP").GetValue("Name", 0);
           
            if (Name == "")
            {
                Global.GFunc.ShowError("Select Name");
                pForm.ActiveItem = "ETNAME";
                return BubbleEvent = false;
            }
            return BubbleEvent;
        }


        private const string MNU_DUPLICATE = "FIL_DUPL";
       
        private void Form_RightClickBefore(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(eventInfo.FormUID);
            bool shouldShow = (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE);
            AddOrRemoveRightClickMenu(shouldShow);

        }
        private void AddOrRemoveRightClickMenu(bool add)
        {
            // Most add-ons attach to Data menu ("1280") so it appears in context menu too
            SAPbouiCOM.MenuItem dataMenu = Application.SBO_Application.Menus.Item("1280");
            SAPbouiCOM.Menus subMenus = dataMenu.SubMenus;

            if (add)
            {
                if (!subMenus.Exists(MNU_DUPLICATE))
                {
                    SAPbouiCOM.MenuCreationParams cp = (SAPbouiCOM.MenuCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                    cp.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    cp.UniqueID = MNU_DUPLICATE;
                    cp.String = "Duplicate";
                    cp.Enabled = true;
                    cp.Position = subMenus.Count + 1;

                    subMenus.AddEx(cp);
                }
            }
            else
            {
                if (subMenus.Exists(MNU_DUPLICATE))
                    subMenus.RemoveEx(MNU_DUPLICATE);
            }
        }

        private static void ClearBlankRow(ref SAPbouiCOM.Form pForm)
        {

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)pForm.Items.Item("MTX_01").Specific;
            for (int i = oMatrix.RowCount; i >= 1; i--)
            {
                string cCombo = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("CLCHARCODE").Cells.Item(i).Specific).Value;
                if (((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("CLCHARCODE").Cells.Item(i).Specific).Value == "")
                    oMatrix.DeleteRow(i);
                else
                    break;
            }

        }

        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.LinkedButton LinkedButton1;
        private SAPbouiCOM.ComboBox ComboBox0;
    }
}
