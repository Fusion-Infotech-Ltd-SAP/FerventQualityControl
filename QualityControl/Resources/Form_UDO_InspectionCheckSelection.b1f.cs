using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QualityControl.Resources
{
    [FormAttribute("QualityControl.Resources.Form_UDO_InspectionCheckSelection", "Resources/Form_UDO_InspectionCheckSelection.b1f")]
    class Form_UDO_InspectionCheckSelection : UserFormBase
    {


        private SAPbouiCOM.StaticText STFROMDATE, STTODATE, STDOCTYPE, STITMGRPCD, STITEMCODE, STITEMNAME, STBPCODE, STBPNAME, STSTATUS;
        private SAPbouiCOM.EditText ETFROMDATE,ETTODATE, ETITEMCODE, ETITEMNAME, ETBPCODE, ETBPNAME;



        private SAPbouiCOM.ComboBox CBDOCTYPE, CBITMGRPCD, CBSTATUS;
        private SAPbouiCOM.Button Generate, Cancel;

        public Form_UDO_InspectionCheckSelection()
        {
        }

        public override void OnInitializeComponent()
        {
            this.STFROMDATE = ((SAPbouiCOM.StaticText)(this.GetItem("STFROMDATE").Specific));
            this.STTODATE = ((SAPbouiCOM.StaticText)(this.GetItem("STTODATE").Specific));
            this.ETFROMDATE = ((SAPbouiCOM.EditText)(this.GetItem("ETFROMDATE").Specific));
            this.ETTODATE = ((SAPbouiCOM.EditText)(this.GetItem("ETTODATE").Specific));
            this.STDOCTYPE = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCTYPE").Specific));
            this.CBDOCTYPE = ((SAPbouiCOM.ComboBox)(this.GetItem("CBDOCTYPE").Specific));
            this.STITMGRPCD = ((SAPbouiCOM.StaticText)(this.GetItem("STITMGRPCD").Specific));
            this.CBITMGRPCD = ((SAPbouiCOM.ComboBox)(this.GetItem("CBITMGRPCD").Specific));
            this.STITEMCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STITEMCODE").Specific));
            this.STITEMNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STITEMNAME").Specific));
            this.ETITEMCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETITEMCODE").Specific));
            this.ETITEMCODE.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETITEMCODE_ChooseFromListBefore);
            this.ETITEMCODE.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETITEMCODE_ChooseFromListAfter);
            this.ETITEMNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETITEMNAME").Specific));
            this.STBPCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STBPCODE").Specific));
            this.STBPNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STBPNAME").Specific));
            this.ETBPCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETBPCODE").Specific));
            this.ETBPCODE.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETBPCODE_ChooseFromListAfter);
            this.ETBPNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETBPNAME").Specific));
            this.STSTATUS = ((SAPbouiCOM.StaticText)(this.GetItem("STSTATUS").Specific));
            this.CBSTATUS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSTATUS").Specific));
            this.Generate = ((SAPbouiCOM.Button)(this.GetItem("AddBtn").Specific));
            this.Generate.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Generate_PressedBefore);
            this.Generate.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Generate_PressedAfter);
            this.Cancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.OnCustomInitialize();

        }


        public override void OnInitializeFormEvents()
        {
            //  this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }



        private void OnCustomInitialize()
        {

        }

       

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
           // throw new System.NotImplementedException();

        }
        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            SAPbouiCOM.ComboBox CBSTATUS = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBSTATUS").Specific;
            CBSTATUS.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue);

            //var oDB = oForm.DataSources.DBDataSources.Item("@MY_UDT");

            //if (string.IsNullOrWhiteSpace(oDB.GetValue("U_Status", 0)))
            //{
            //    oForm.Freeze(true);

            //    oDB.SetValue("U_Status", 0, "P");

            //    oForm.Freeze(false);
            //}
        }
        private void ETITEMCODE_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_MH_CHKSLC");
                SAPbouiCOM.ComboBox ITMGRPCD = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBITMGRPCD").Specific;
                string SelctedCBITMGRPCD = ITMGRPCD.Selected?.Value ?? string.Empty;
               
                SAPbouiCOM.ChooseFromList ocfl = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item("CFL_OITM");
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                // 🔹 First Condition
                SAPbouiCOM.Condition oCon1 = oCons.Add();
                oCon1.Alias = "frozenFor";
                oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon1.CondVal = "N";

                if(SelctedCBITMGRPCD != string.Empty)
                {
                    // 2nd condition
                    oCon1.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    SAPbouiCOM.Condition oCon2 = oCons.Add();
                    oCon2.Alias = "ItmsGrpCod";
                    oCon2.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon2.CondVal = SelctedCBITMGRPCD;
                }


                // Apply conditions
                ocfl.SetConditions(oCons);

            }
            catch (Exception ex)
            {

            }
        }
        private void ETITEMCODE_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_MH_CHKSLC"); 
                SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string Uid = cflEvent.ChooseFromListUID;  // cfl unique id
           
                SAPbouiCOM.DataTable dtbCFL = cflEvent.SelectedObjects;

                if (!dtbCFL.IsEmpty)
                {
                 
                    ((SAPbouiCOM.EditText)oform.Items.Item("ETITEMCODE").Specific).Value = dtbCFL.GetValue("ItemCode", 0).ToString();
                    ((SAPbouiCOM.EditText)oform.Items.Item("ETITEMNAME").Specific).Value = dtbCFL.GetValue("ItemName", 0).ToString();
    
                }


            }
            catch(Exception ex)
            {

            }

        }

        private void ETBPCODE_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_MH_CHKSLC");
                SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string Uid = cflEvent.ChooseFromListUID;  // cfl unique id

                SAPbouiCOM.DataTable dtbCFL = cflEvent.SelectedObjects;

                if (!dtbCFL.IsEmpty)
                {

                    ((SAPbouiCOM.EditText)oform.Items.Item("ETBPCODE").Specific).Value = dtbCFL.GetValue("CardCode", 0).ToString();
                    ((SAPbouiCOM.EditText)oform.Items.Item("ETBPNAME").Specific).Value = dtbCFL.GetValue("CardName", 0).ToString();

                }
            }
            catch(Exception ex)
            {

            }

        }

        private void Generate_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
            string FromDate = ((SAPbouiCOM.EditText)oForm.Items.Item("ETFROMDATE").Specific).Value;
            string ToDate = ((SAPbouiCOM.EditText)oForm.Items.Item("ETTODATE").Specific).Value;
            SAPbouiCOM.ComboBox CBDOCTYPE = (SAPbouiCOM.ComboBox)oForm.Items.Item("CBDOCTYPE").Specific;
            string SelctedDOCTYPE = CBDOCTYPE.Selected?.Description ?? string.Empty;
            //string status = ((SAPbouiCOM.ComboBox)oForm.Items.Item("CBSTATUS").Specific).Value;

            if (FromDate == "")
            {
                BubbleEvent = false;
                Global.GFunc.ShowError("Select FromDate");
                oForm.ActiveItem = "ETFROMDATE";
                return;
            }
            if (ToDate == "")
            {
                BubbleEvent = false;
                Global.GFunc.ShowError("Select ToDate");
                oForm.ActiveItem = "ETTODATE";
                return;
            }

            if (SelctedDOCTYPE == "")
            {
                BubbleEvent = false;
                Global.GFunc.ShowError("Select Document");
                oForm.ActiveItem = "CBDOCTYPE";
                return;
            }

         
            //if (status == "")
            //{
            //    Global.GFunc.ShowError("Select Status");
            //    oForm.ActiveItem = "CBSTATUS";
            //    return;
            //}

        }

        private void Generate_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
            try
            {
                OpenQCInspectionSelection(ref oform);
            }
            catch (Exception EX)
            {

            }

        }

        private static void OpenQCInspectionSelection(ref SAPbouiCOM.Form pForm)
        {

            string FromDate = ((SAPbouiCOM.EditText)pForm.Items.Item("ETFROMDATE").Specific).Value;
            string ToDate = ((SAPbouiCOM.EditText)pForm.Items.Item("ETTODATE").Specific).Value;
            SAPbouiCOM.ComboBox CBDOCTYPE = (SAPbouiCOM.ComboBox)pForm.Items.Item("CBDOCTYPE").Specific;
            string SelctedDOCTYPE = CBDOCTYPE.Selected?.Value ?? string.Empty;
            SAPbouiCOM.ComboBox ITMGRPCD = (SAPbouiCOM.ComboBox)pForm.Items.Item("CBITMGRPCD").Specific;
            string SelctedCBITMGRPCD = ITMGRPCD.Selected?.Value ?? string.Empty;
            string ItemCode = ((SAPbouiCOM.EditText)pForm.Items.Item("ETITEMCODE").Specific).Value;
            string CardCode = ((SAPbouiCOM.EditText)pForm.Items.Item("ETBPCODE").Specific).Value;
            string status = ((SAPbouiCOM.ComboBox)pForm.Items.Item("CBSTATUS").Specific).Value;
            if(status == "")
            {
                status = "P";
            }
         
            //Form cForm = LoadForm("FFS_FRM_NO_QCDOCS", false, true);
            Form_NO_IPndList IpdnList = new Form_NO_IPndList();
            IpdnList.Show();
            SAPbouiCOM.Form cForm = Application.SBO_Application.Forms.Item("FIL_FRM_NO_IPNDLIST");
            try
            {
                cForm.SupportedModes = 0;
                cForm.Freeze(true);
                cForm.DataSources.UserDataSources.Item("UDPARENTID").ValueEx = pForm.UniqueID;
                cForm.DataSources.UserDataSources.Item("UDFD").ValueEx = FromDate;
                cForm.DataSources.UserDataSources.Item("UDTD").ValueEx = ToDate;
                cForm.DataSources.UserDataSources.Item("UDDOCTYPE").ValueEx = SelctedDOCTYPE;
                cForm.DataSources.UserDataSources.Item("UDITMGRPCD").ValueEx = SelctedCBITMGRPCD;
                cForm.DataSources.UserDataSources.Item("UDITEMCODE").ValueEx = ItemCode;
                cForm.DataSources.UserDataSources.Item("UDCARDCODE").ValueEx = CardCode;
                cForm.DataSources.UserDataSources.Item("UDSTATUS").ValueEx = status;


                Form_NO_IPndList.LoadPendingQC(ref cForm);
                cForm.Freeze(false);
            }
            catch (Exception ex)
            {
                cForm.Items.Item("2").Click();
                cForm.Freeze(false);
               // ShowError(ex.Message);
            }
        }




    }
}
