using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QualityControl.Resources
{
    [FormAttribute("QualityControl.Resources.Form_UDO_InspectionParam", "Resources/Form_UDO_InspectionParam.b1f")]
    class Form_UDO_InspectionParam : UserFormBase
    {
        public Form_UDO_InspectionParam()
        {
        }

        private SAPbouiCOM.Button Add,Cancel;
        private SAPbouiCOM.StaticText STCODE , STNAME ;
        private SAPbouiCOM.EditText ETCODE, ETNAME, ETDOCENTRY;
      

      
        public override void OnInitializeComponent()
        {
            this.Add = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Add.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Add_PressedAfter);
            this.Add.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Add_PressedBefore);
            this.Cancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.STCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STCODE").Specific));
            this.ETCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCODE").Specific));
            this.STNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STNAME").Specific));
            this.ETNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETNAME").Specific));
            this.ETDOCENTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCENTRY").Specific));
            this.OnCustomInitialize();

        }

      
        public override void OnInitializeFormEvents()
        {
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private void OnCustomInitialize()
        {
        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
            oform.Items.Item("ETCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
        }


        private void Add_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oform.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ValidateForm(ref oform, ref BubbleEvent);
            }

        }

        private void Add_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);

            if(oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                Menu objMenu = new Menu();
                objMenu.SetNextCode(oform, "ETCODE", "FIL_MH_QINSPMAS");
            }
            

        }
        private bool ValidateForm(ref SAPbouiCOM.Form pForm, ref bool BubbleEvent)
        {
            string Code = pForm.DataSources.DBDataSources.Item("@FIL_MH_QINSPMAS").GetValue("Code", 0);
            string Name = pForm.DataSources.DBDataSources.Item("@FIL_MH_QINSPMAS").GetValue("Name", 0);
            
            if (Name == "")
            {
                Global.GFunc.ShowError("Select Characteristic Name");
                pForm.ActiveItem = "ETNAME";
                return BubbleEvent = false;
            }
            return BubbleEvent;
        }

    }
}
