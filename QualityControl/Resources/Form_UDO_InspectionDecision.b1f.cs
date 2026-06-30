using QualityControl.Helper;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace QualityControl.Resources
{
    [FormAttribute("QualityControl.Resources.Form_UDO_InspectionDecision", "Resources/Form_UDO_InspectionDecision.b1f")]
    class Form_UDO_InspectionDecision : UserFormBase
    {
        public Form_UDO_InspectionDecision()
        {
        }

        private SAPbouiCOM.StaticText STBSEDOCNO,STDOCNUM,STITEMCODE,STITEMNAME,STACTQTY, STQCQTY,STDOCDATE, STTRNSFRNM, STREMARKS,
                                      STACPTQTY,STGRDAQTY,STGRDBQTY,STGRDCQTY,STGRDDQTY,STGRDEQTY, STACPTITM, STGRDAITM, STGRDBITM, STGRDCITM, STGRDDITM,
                                      STGRDEITM, STGRDCWHS, STGRDAWHS, STGRDBWHS, STACPTWHS, STGRDDWHS, STGRDEWHS, STSCRPQTY, STSCRPITM, STSCRPWHS,STBATCH,STQCWHS;

       

        private SAPbouiCOM.EditText ETBSEDOCNO,ETDOCNUM,ETITEMCODE,ETITEMNAME,ETACTQTY,ETQCQTY,ETDOCDATE,ETTRNSFRID,ETTRNSFRNM,ETREMARKS,
                                    ETACPTQTY,ETGRDAQTY,ETGRDBQTY,ETGRDCQTY,ETGRDDQTY,ETGRDEQTY,ETACPTITM,ETGRDAITM,ETGRDBITM,ETGRDCITM,
                                    ETGRDDITM,ETGRDEITM,ETSCRPQTY,ETSCRPITM,ETDOCENTRY, ETGRDAWHS, ETGRDCWHS, ETGRDDWHS, ETGRDEWHS, ETSCRPWHS, ETACPTWHS,ETBATCH,ETQCWHS;



        private SAPbouiCOM.ComboBox CBSERIES;
        private SAPbouiCOM.Button Add,Cancel, BRWSBTN, DISPBTN, DELBTN;



        private SAPbouiCOM.Matrix MTX01;
        private SAPbouiCOM.LinkedButton LinkedButton;


        public override void OnInitializeComponent()
        {
            this.STBSEDOCNO = ((SAPbouiCOM.StaticText)(this.GetItem("STBSEDOCNO").Specific));
            this.ETBSEDOCNO = ((SAPbouiCOM.EditText)(this.GetItem("ETBSEDOCNO").Specific));
            this.STDOCNUM = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCNUM").Specific));
            this.CBSERIES = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSERIES").Specific));
            this.ETDOCNUM = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCNUM").Specific));
            this.STITEMCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STITEMCODE").Specific));
            this.ETITEMCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETITEMCODE").Specific));
            this.STITEMNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STITEMNAME").Specific));
            this.ETITEMNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETITEMNAME").Specific));
            this.STACTQTY = ((SAPbouiCOM.StaticText)(this.GetItem("STACTQTY").Specific));
            this.ETACTQTY = ((SAPbouiCOM.EditText)(this.GetItem("ETACTQTY").Specific));
            this.STQCQTY = ((SAPbouiCOM.StaticText)(this.GetItem("STQCQTY").Specific));
            this.ETQCQTY = ((SAPbouiCOM.EditText)(this.GetItem("ETQCQTY").Specific));
            this.ETQCQTY.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETQCQTY_LostFocusAfter);
            this.STDOCDATE = ((SAPbouiCOM.StaticText)(this.GetItem("STDOCDATE").Specific));
            this.ETDOCDATE = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCDATE").Specific));
            this.ETDOCDATE.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETDOCDATE_LostFocusAfter);
            this.STTRNSFRNM = ((SAPbouiCOM.StaticText)(this.GetItem("STTRNSFRNM").Specific));
            this.ETTRNSFRID = ((SAPbouiCOM.EditText)(this.GetItem("ETTRNSFRID").Specific));
            this.ETTRNSFRNM = ((SAPbouiCOM.EditText)(this.GetItem("ETTRNSFRNM").Specific));
            this.MTX01 = ((SAPbouiCOM.Matrix)(this.GetItem("MTX01").Specific));
            this.MTX01.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.MTX01_LostFocusAfter);
            this.STREMARKS = ((SAPbouiCOM.StaticText)(this.GetItem("STREMARKS").Specific));
            this.ETREMARKS = ((SAPbouiCOM.EditText)(this.GetItem("ETREMARKS").Specific));
            this.Add = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Add.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Add_PressedAfter);
            this.Add.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Add_PressedBefore);
            this.Cancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.STACPTQTY = ((SAPbouiCOM.StaticText)(this.GetItem("STACPTQTY").Specific));
            this.STGRDAQTY = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDAQTY").Specific));
            this.STGRDBQTY = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDBQTY").Specific));
            this.STGRDCQTY = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDCQTY").Specific));
            this.STGRDDQTY = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDDQTY").Specific));
            this.STGRDEQTY = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDEQTY").Specific));
            this.ETACPTQTY = ((SAPbouiCOM.EditText)(this.GetItem("ETACPTQTY").Specific));
            this.ETGRDAQTY = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDAQTY").Specific));
            this.ETGRDAQTY.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETGRDAQTY_LostFocusAfter);
            this.ETGRDBQTY = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDBQTY").Specific));
            this.ETGRDBQTY.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETGRDBQTY_LostFocusAfter);
            this.ETGRDCQTY = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDCQTY").Specific));
            this.ETGRDCQTY.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETGRDCQTY_LostFocusAfter);
            this.ETGRDDQTY = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDDQTY").Specific));
            this.ETGRDDQTY.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETGRDDQTY_LostFocusAfter);
            this.ETGRDEQTY = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDEQTY").Specific));
            this.ETGRDEQTY.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETGRDEQTY_LostFocusAfter);
            this.STACPTITM = ((SAPbouiCOM.StaticText)(this.GetItem("STACPTITM").Specific));
            this.STGRDAITM = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDAITM").Specific));
            this.STGRDBITM = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDBITM").Specific));
            this.STGRDCITM = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDCITM").Specific));
            this.STGRDDITM = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDDITM").Specific));
            this.STGRDEITM = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDEITM").Specific));
            this.ETACPTITM = ((SAPbouiCOM.EditText)(this.GetItem("ETACPTITM").Specific));
            this.ETGRDAITM = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDAITM").Specific));
            this.ETGRDAITM.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.ETGRDAITM_KeyDownAfter);
            this.ETGRDAITM.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETGRDAITM_ChooseFromListBefore);
            this.ETGRDAITM.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETGRDAITM_ChooseFromListAfter);
            this.ETGRDBITM = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDBITM").Specific));
            this.ETGRDBITM.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.ETGRDBITM_KeyDownAfter);
            this.ETGRDBITM.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETGRDBITM_ChooseFromListBefore);
            this.ETGRDBITM.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETGRDBITM_ChooseFromListAfter);
            this.ETGRDCITM = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDCITM").Specific));
            this.ETGRDCITM.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.ETGRDCITM_KeyDownAfter);
            this.ETGRDCITM.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.ETGRDCITM_ChooseFromListBefore);
            this.ETGRDCITM.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETGRDCITM_ChooseFromListAfter);
            this.ETGRDDITM = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDDITM").Specific));
            this.ETGRDEITM = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDEITM").Specific));
            this.STGRDCWHS = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDCWHS").Specific));
            this.STGRDAWHS = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDAWHS").Specific));
            this.STGRDBWHS = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDBWHS").Specific));
            this.STACPTWHS = ((SAPbouiCOM.StaticText)(this.GetItem("STACPTWHS").Specific));
            this.STGRDDWHS = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDDWHS").Specific));
            this.STGRDEWHS = ((SAPbouiCOM.StaticText)(this.GetItem("STGRDEWHS").Specific));
            //                            this.ETGRDBWHS.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.ETGRDBWHS_KeyDownAfter);
            this.STSCRPQTY = ((SAPbouiCOM.StaticText)(this.GetItem("STSCRPQTY").Specific));
            this.ETSCRPQTY = ((SAPbouiCOM.EditText)(this.GetItem("ETSCRPQTY").Specific));
            this.ETSCRPQTY.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.ETSCRPQTY_LostFocusAfter);
            this.STSCRPITM = ((SAPbouiCOM.StaticText)(this.GetItem("STSCRPITM").Specific));
            this.ETSCRPITM = ((SAPbouiCOM.EditText)(this.GetItem("ETSCRPITM").Specific));
            this.STSCRPWHS = ((SAPbouiCOM.StaticText)(this.GetItem("STSCRPWHS").Specific));
            this.ETDOCENTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCENTRY").Specific));
            this.ETACPTWHS = ((SAPbouiCOM.EditText)(this.GetItem("ETACPTWHS").Specific));
            this.ETGRDAWHS = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDAWHS").Specific));
            this.ETGRDAWHS = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDBWHS").Specific));
            this.ETGRDCWHS = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDCWHS").Specific));
            this.ETGRDDWHS = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDDWHS").Specific));
            this.ETGRDEWHS = ((SAPbouiCOM.EditText)(this.GetItem("ETGRDEWHS").Specific));
            this.ETSCRPWHS = ((SAPbouiCOM.EditText)(this.GetItem("ETSCRPWHS").Specific));
            this.LinkedButton = ((SAPbouiCOM.LinkedButton)(this.GetItem("LNTRNSFRNM").Specific));
            this.STBATCH = ((SAPbouiCOM.StaticText)(this.GetItem("STBATCH").Specific));
            this.ETBATCH = ((SAPbouiCOM.EditText)(this.GetItem("ETBATCH").Specific));
            this.STQCWHS = ((SAPbouiCOM.StaticText)(this.GetItem("STQCWHS").Specific));
            this.ETQCWHS = ((SAPbouiCOM.EditText)(this.GetItem("ETQCWHS").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("ETPRDNTRYA").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("ETPRDNTRYB").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("ETPRDNTRYC").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("ETPRDNTRYD").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("ETPRDNTRYE").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LNPRDNTRYA").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LNPRDNTRYB").Specific));
            this.LinkedButton2 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LNPRDNTRYC").Specific));
            this.LinkedButton3 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LNPRDNTRYD").Specific));
            this.LinkedButton4 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LNPRDNTRYE").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("STTACPTQTY").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("ETTACPTQTY").Specific));
            this.LinkedButton5 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBITEMCODE").Specific));
            this.LinkedButton6 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBSCRPITM").Specific));
            this.LinkedButton7 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDEITM").Specific));
            this.LinkedButton8 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDDITM").Specific));
            this.LinkedButton9 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDCITM").Specific));
            this.LinkedButton10 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDBITM").Specific));
            this.LinkedButton11 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDAITM").Specific));
            this.LinkedButton12 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBACPTITM").Specific));
            this.LinkedButton13 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBACPTWHS").Specific));
            this.LinkedButton14 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDAWHS").Specific));
            this.LinkedButton15 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDBWHS").Specific));
            this.LinkedButton16 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDCWHS").Specific));
            this.LinkedButton17 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDDWHS").Specific));
            this.LinkedButton18 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBGRDEWHS").Specific));
            this.LinkedButton19 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LBSCRPWHS").Specific));
            this.LinkedButton20 = ((SAPbouiCOM.LinkedButton)(this.GetItem("LNBASENTRY").Specific));
            this.LinkedButton20.PressedBefore += new SAPbouiCOM._ILinkedButtonEvents_PressedBeforeEventHandler(this.LinkedButton20_PressedBefore);
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("ETBASENTRY").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("ETBASETYPE").Specific));
            this.EditText8 = ((SAPbouiCOM.EditText)(this.GetItem("ETACPITEN").Specific));
            this.EditText9 = ((SAPbouiCOM.EditText)(this.GetItem("ETGRAITMN").Specific));
            this.EditText10 = ((SAPbouiCOM.EditText)(this.GetItem("ETGRBITMN").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("ETGRCITMN").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("ETSCPITMN").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("STITEM").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("STITEMN").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("STWHS").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("TAB01").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("TAB02").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("MTXATTCH").Specific));
            this.BRWSBTN = ((SAPbouiCOM.Button)(this.GetItem("BRWSBTN").Specific));
            this.BRWSBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.BRWSBTN_ClickAfter);
            this.DISPBTN = ((SAPbouiCOM.Button)(this.GetItem("DISPBTN").Specific));
            this.DISPBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.DISPBTN_ClickAfter);
            this.DELBTN = ((SAPbouiCOM.Button)(this.GetItem("DELBTN").Specific));
            this.DELBTN.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.DELBTN_ClickAfter);
            this.BTSTKTRNS = ((SAPbouiCOM.Button)(this.GetItem("BTSTKTRNS").Specific));
            this.BTSTKTRNS.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.BTSTKTRNS_PressedAfter);
            this.OnCustomInitialize();

        }


        public override void OnInitializeFormEvents()
        {
            this.DataAddAfter += new SAPbouiCOM.Framework.FormBase.DataAddAfterHandler(this.Form_DataAddAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }


        private void OnCustomInitialize()
        {

        }

        private void ETDOCDATE_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                DateTime postDate = DateTime.ParseExact(((SAPbouiCOM.EditText)oForm.Items.Item("ETDOCDATE").Specific).Value, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);

                if (!Global.GFunc.LoadSeriesAndSetDocNum(oForm, "CBSERIES", "FIL_D_INSPDECN", "@FIL_DH_INSPDECN", postDate))
                {
                    return;
                }
            }
            catch (Exception ex)
            {
            }


        }


        private void LinkedButton20_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            SAPbouiCOM.LinkedButton oLinkButton;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            oLinkButton = (SAPbouiCOM.LinkedButton)oForm.Items.Item("LNBASENTRY").Specific;

            string type = ((SAPbouiCOM.EditText)oForm.Items.Item("ETBASETYPE").Specific).Value;

            if (type == "20")
            {
                oLinkButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GoodsReceiptPO;
            }
            else if (type == "59")
            {
                oLinkButton.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_GoodsReceipt;
            }

        }
        internal static void LoadQCValue(ref SAPbouiCOM.Form pForm, string BaseDocNum,string DocEntry,string ObjType, string ItemCode, string ItemName, decimal Quantity, string ScrapWhsCode,string rejectWhs,string NormalWhsCode,string Bacth,string QcWhs, string Status)
        {
            pForm.Freeze(true);
            bool BubbleEvent = true;
            InitFormAddMode(ref pForm, ref BubbleEvent);

            if (Status == "Complete")
            {
               // ExistingRecord(ref pForm, QCNum);
            }
            else
            {
                SAPbouiCOM.Matrix oMatrix;
                string qStr;
                string qStr1;
                SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rSet1 = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbouiCOM.Matrix MTX01 = (SAPbouiCOM.Matrix)pForm.Items.Item("MTX01").Specific;
                MTX01.AutoResizeColumns();//define matrix
                SAPbouiCOM.DBDataSource oDataSource = (SAPbouiCOM.DBDataSource)pForm.DataSources.DBDataSources.Item("@FIL_DR_INSPDECN");

                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETBSEDOCNO").Specific).Value = BaseDocNum;
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETITEMCODE").Specific).Value = ItemCode;
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETITEMNAME").Specific).Value = ItemName;
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETACPTITM").Specific).Value = ItemCode;
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETACPITEN").Specific).Value = ItemName;
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETSCRPITM").Specific).Value = ItemCode;
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETSCPITMN").Specific).Value = ItemName;
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETACPTWHS").Specific).Value = NormalWhsCode;
                if(ObjType == "20")
                {
                    ((SAPbouiCOM.IEditText)pForm.Items.Item("ETSCRPWHS").Specific).Value = rejectWhs;
                }
                else
                {
                    ((SAPbouiCOM.IEditText)pForm.Items.Item("ETSCRPWHS").Specific).Value = ScrapWhsCode;
                }
     
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETACTQTY").Specific).Value = Quantity.ToString();
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETQCQTY").Specific).Value = Quantity.ToString();
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETACPTQTY").Specific).Value = Quantity.ToString();
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETTACPTQTY").Specific).Value = Quantity.ToString();
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETBATCH").Specific).Value = Bacth;
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETQCWHS").Specific).Value = QcWhs;
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETBASENTRY").Specific).Value = DocEntry.ToString();
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETBASETYPE").Specific).Value = ObjType.ToString();
                //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_BASETYPE", 0, ObjType.ToString());
                //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_BASENTRY", 0, DocEntry.ToString());
                if (ObjType == "20")
                {
                    pForm.Items.Item("BTSTKTRNS").Visible = false;
                   ((SAPbouiCOM.StaticText)pForm.Items.Item("STSCRPQTY").Specific).Caption = "Reject Qty";
                    pForm.Items.Item("ETGRDAQTY").Enabled = false;
                    pForm.Items.Item("ETGRDBQTY").Enabled = false;
                    pForm.Items.Item("ETGRDCQTY").Enabled = false;
                    //pForm.Items.Item("ETGRDDQTY").Enabled = false;
                    //pForm.Items.Item("ETGRDEQTY").Enabled = false;

                    pForm.Items.Item("ETGRDAITM").Enabled = false;
                    pForm.Items.Item("ETGRDBITM").Enabled = false;
                    pForm.Items.Item("ETGRDCITM").Enabled = false;
                    //pForm.Items.Item("ETGRDDITM").Enabled = false;
                    //pForm.Items.Item("ETGRDEITM").Enabled = false;

                    pForm.Items.Item("ETGRDAWHS").Enabled = false;
                    pForm.Items.Item("ETGRDBWHS").Enabled = false;
                    pForm.Items.Item("ETGRDCWHS").Enabled = false;
                    //pForm.Items.Item("ETGRDDWHS").Enabled = false;
                    //pForm.Items.Item("ETGRDEWHS").Enabled = false;

                    pForm.Items.Item("ETPRDNTRYA").Enabled = false;
                    pForm.Items.Item("ETPRDNTRYB").Enabled = false;
                    pForm.Items.Item("ETPRDNTRYC").Enabled = false;
                    //pForm.Items.Item("ETPRDNTRYD").Enabled = false;
                    //pForm.Items.Item("ETPRDNTRYE").Enabled = false; 
                }
                else if (ObjType == "59")
                {
                    pForm.Items.Item("BTSTKTRNS").Visible = false;
                    pForm.Items.Item("ETGRDAQTY").Enabled = true;
                    pForm.Items.Item("ETGRDBQTY").Enabled = true;
                    pForm.Items.Item("ETGRDCQTY").Enabled = true;
                    //pForm.Items.Item("ETGRDDQTY").Enabled = true;
                    //pForm.Items.Item("ETGRDEQTY").Enabled = true;

                    pForm.Items.Item("ETGRDAITM").Enabled = true;
                    pForm.Items.Item("ETGRDBITM").Enabled = true;
                    pForm.Items.Item("ETGRDCITM").Enabled = true;
                    //pForm.Items.Item("ETGRDDITM").Enabled = true;
                    //pForm.Items.Item("ETGRDEITM").Enabled = true;

                    pForm.Items.Item("ETGRDAWHS").Enabled = true;
                    pForm.Items.Item("ETGRDBWHS").Enabled = true;
                    pForm.Items.Item("ETGRDCWHS").Enabled = true;
                    //pForm.Items.Item("ETGRDDWHS").Enabled = true;
                    //pForm.Items.Item("ETGRDEWHS").Enabled = true;

                    pForm.Items.Item("ETPRDNTRYA").Enabled = true;
                    pForm.Items.Item("ETPRDNTRYB").Enabled = true;
                    pForm.Items.Item("ETPRDNTRYC").Enabled = true;
                    //pForm.Items.Item("ETPRDNTRYD").Enabled = true;
                    //pForm.Items.Item("ETPRDNTRYE").Enabled = true;

                    //qStr1 = $"SELECT" +
                    //    $" T0.\"U_GRDAITM\" AS \"GRDAITM\",(SELECT \"DfltWH\" FROM \"OITM\" WHERE \"ItemCode\" = T0.\"U_GRDAITM\") AS \"GRDAWHS\"," +
                    //    $" T0.\"U_GRDBITM\" AS \"GRDBITM\",(SELECT \"DfltWH\" FROM \"OITM\" WHERE \"ItemCode\" = T0.\"U_GRDBITM\") AS \"GRDBWHS\"," +
                    //    $" T0.\"U_GRDCITM\" AS \"GRDCITM\", (SELECT \"DfltWH\" FROM \"OITM\" WHERE \"ItemCode\" = T0.\"U_GRDCITM\") AS \"GRDCWHS\", " +
                    //    $" T0.\"U_GRDDITM\" AS \"GRDDITM\", (SELECT \"DfltWH\" FROM \"OITM\" WHERE \"ItemCode\" = T0.\"U_GRDDITM\") AS \"GRDDWHS\", " +
                    //    $" T0.\"U_GRDEITM\" AS \"GRDEITM\", (SELECT \"DfltWH\" FROM \"OITM\" WHERE \"ItemCode\" = T0.\"U_GRDEITM\") AS \"GRDEWHS\" " +
                    //    $" FROM \"OITM\" T0 " +
                    //    $" WHERE T0.\"ItemCode\" = '" + ItemCode + "'";

                    //rSet1.DoQuery(qStr1);

                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDAITM", 0, rSet1.Fields.Item("GRDAITM").Value.ToString());
                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDAWHS", 0, rSet1.Fields.Item("GRDAWHS").Value.ToString());

                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDBITM", 0, rSet1.Fields.Item("GRDBITM").Value.ToString());
                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDBWHS", 0, rSet1.Fields.Item("GRDBWHS").Value.ToString());

                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDCITM", 0, rSet1.Fields.Item("GRDCITM").Value.ToString());
                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDCWHS", 0, rSet1.Fields.Item("GRDCWHS").Value.ToString());

                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDDITM", 0, rSet1.Fields.Item("GRDDITM").Value.ToString());
                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDDWHS", 0, rSet1.Fields.Item("GRDDWHS").Value.ToString());

                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDEITM", 0, rSet1.Fields.Item("GRDEITM").Value.ToString());
                    //pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("U_GRDEWHS", 0, rSet1.Fields.Item("GRDEWHS").Value.ToString());

                }

                qStr = $"SELECT" +
                           $" T0.\"U_CHRGCODE\" AS \"ChrgCode\", T0.\"U_CHRGNAME\" AS \"ChrgName\",T1.\"U_CHARCODE\" AS \"CharCode\", " +
                           $" T1.\"U_CHARNAME\" AS \"CharName\", T1.\"U_TRGVALUE\" AS \"Trgvalue\",T1.\"U_LOWVALUE\" AS \"LowValue\", " +
                           $" T1.\"U_UPRVALUE\" AS \"UprValue\",T2.\"U_ITEMCODE\" AS \"ItemCode\" " +
                        $" FROM \"@FIL_MH_INPLAN\" T0 " +
                        $" INNER JOIN \"@FIL_MR_INPLNPM\" T1 ON T1.\"Code\" = T0.\"Code\" " +
                        $" INNER JOIN \"@FIL_MR_INPLNIM\" T2 ON T2.\"Code\" = T0.\"Code\" " +
                       $" WHERE T0.\"U_ACTIVE\" = 'Y' AND T2.\"U_ACTIVE\" = 'Y'  AND T2.\"U_ITEMCODE\" = '" + ItemCode + "'";
            
                rSet.DoQuery(qStr);
                pForm.Freeze(true);
                oDataSource.Clear();
                while (!rSet.EoF)
                {
                    oDataSource.InsertRecord(oDataSource.Size);
                    oDataSource.SetValue("LineId", oDataSource.Size - 1, oDataSource.Size.ToString());
                    oDataSource.SetValue("U_CHRGNAME", oDataSource.Size - 1, rSet.Fields.Item("ChrgName").Value.ToString());
                    oDataSource.SetValue("U_CHARNAME", oDataSource.Size - 1, rSet.Fields.Item("CharName").Value.ToString());
                    oDataSource.SetValue("U_TRGVALUE", oDataSource.Size - 1, rSet.Fields.Item("Trgvalue").Value.ToString());
                    oDataSource.SetValue("U_LOWVALUE", oDataSource.Size - 1, rSet.Fields.Item("LowValue").Value.ToString());
                    oDataSource.SetValue("U_UPRVALUE", oDataSource.Size - 1, rSet.Fields.Item("UprValue").Value.ToString());

                    rSet.MoveNext();
                }
                oDataSource.InsertRecord(oDataSource.Size);
                oDataSource.SetValue("LineId", oDataSource.Size - 1, oDataSource.Size.ToString());
                MTX01.Clear();
                MTX01.LoadFromDataSource();

            }

            pForm.Freeze(false);
        }

        private static void InitFormAddMode(ref SAPbouiCOM.Form pForm, ref bool BubbleEvent)
        {
            try
            {
                if (pForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    pForm.Items.Item("TAB01").Click();
                    Global.GFunc.LoadSeriesAndSetDocNum(pForm, "CBSERIES", "FIL_D_INSPDECN", "@FIL_DH_INSPDECN");
                }

                //Current Date
                SAPbouiCOM.EditText oedt = (SAPbouiCOM.EditText)pForm.Items.Item("ETDOCDATE").Specific; //assign / define a edittext
                oedt.Active = true;
                oedt.String = "W";
            }

            catch (Exception EX)
            {

            }
        }

        private void ETSCRPQTY_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            try
            {
                oform.Freeze(true);
                QuantityCalculation(ref oform);
                oform.Freeze(false);
            }
            catch (Exception ex)
            {
                oform.Freeze(false);
            }
        }

        private void ETQCQTY_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            try
            {
                oform.Freeze(true);
                QuantityCalculation(ref oform);
                oform.Freeze(false);
            }
            catch (Exception ex)
            {
                oform.Freeze(false);
            }

        }

        private void DELBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

            SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_INSPATCH");

            SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

            MTXATTCH.FlushToDataSource();

            for (int i = 1; i <= MTXATTCH.RowCount; i++)
            {
                if (MTXATTCH.IsRowSelected(i))
                {
                    int rowIndex = i - 1;

                    if (rowIndex >= 0 && rowIndex < DBDataSourceLine.Size)
                    {
                        string filePath =
                            DBDataSourceLine.GetValue("U_ATCHMENT", rowIndex).Trim();

                        // Physical file delete
                        if (!string.IsNullOrWhiteSpace(filePath) &&
                            File.Exists(filePath))
                        {
                            try
                            {
                                File.Delete(filePath);
                            }
                            catch (Exception ex)
                            {
                                Application.SBO_Application.MessageBox("File delete error: " + ex.Message);
                            }
                        }

                        DBDataSourceLine.RemoveRecord(rowIndex);

                        // LineId resequence
                        for (int j = 0; j < DBDataSourceLine.Size; j++)
                        {
                            DBDataSourceLine.Offset = j;
                            DBDataSourceLine.SetValue("LineId",j,(j + 1).ToString());
                        }

                        MTXATTCH.LoadFromDataSource();

                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }

                        Application.SBO_Application.MessageBox("File deleted successfully");
                    }

                    break;
                }
            }
        }

        private void DISPBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

                for (int i = 1; i <= MTXATTCH.VisualRowCount; i++)
                {
                    if (MTXATTCH.IsRowSelected(i))
                    {
                        string filePath = ((SAPbouiCOM.EditText)MTXATTCH.Columns.Item("CLATTACH").Cells.Item(i).Specific).Value.Trim();

                        if (string.IsNullOrWhiteSpace(filePath))
                        {
                            Application.SBO_Application.MessageBox("File path is empty.");
                            return;
                        }

                        if (!System.IO.File.Exists(filePath))
                        {
                            Application.SBO_Application.MessageBox("File not found:\n" + filePath);
                            return;
                        }

                        System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo();

                        psi.FileName = filePath;
                        psi.UseShellExecute = true;

                        System.Diagnostics.Process.Start(psi);

                        return;
                    }
                }

                Application.SBO_Application.MessageBox("Please select a row.");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox("Error: " + ex.Message);
            }
        }
        
        private void BRWSBTN_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);

                SAPbouiCOM.DBDataSource DBDataSourceLine = oForm.DataSources.DBDataSources.Item("@FIL_DR_INSPATCH");

                SAPbouiCOM.Matrix MTXATTCH = (SAPbouiCOM.Matrix)oForm.Items.Item("MTXATTCH").Specific;

                //Browse file
                string sourceFile = FileDialogHelper.ShowFileDialog();

                if (string.IsNullOrWhiteSpace(sourceFile))
                    return;

                //Get Document Number
                string docNo = oForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("DocNum", 0).Trim();

                if (string.IsNullOrEmpty(docNo))
                    docNo = "Draft";

                //Get SAP attachment path
                SAPbobsCOM.Recordset rs =  (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                rs.DoQuery($"SELECT \"AttachPath\" FROM OADP");

                string attachPath = rs.Fields.Item("AttachPath").Value.ToString();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(rs);

                if (string.IsNullOrWhiteSpace(attachPath))
                    throw new Exception("SAP Attachment Path not found");

                //Create new file name
                string fileName =  System.IO.Path.GetFileNameWithoutExtension(sourceFile);

                string extension = System.IO.Path.GetExtension(sourceFile);

                string newFileName =  docNo + "_" + fileName + extension;

                string destinationFile = System.IO.Path.Combine(attachPath, newFileName);

                //avoid duplicate file
                int i = 1;

                while (File.Exists(destinationFile))
                {
                    destinationFile = Path.Combine( attachPath, docNo + "_" + fileName + "_" + i + extension
                    );

                    i++;
                }

                //Copy file
                File.Copy(sourceFile, destinationFile);

                //Matrix Row
                int lastRow = MTXATTCH.VisualRowCount;

                bool needNewRow =
                    lastRow == 0 ||
                    !string.IsNullOrEmpty(
                    ((SAPbouiCOM.EditText)
                    MTXATTCH.Columns.Item("CLATTACH")
                    .Cells.Item(lastRow).Specific).Value);

                if (needNewRow)
                {
                    Global.GFunc.SetNewLine(MTXATTCH,DBDataSourceLine,1,"");

                    lastRow = MTXATTCH.VisualRowCount;
                }

                ((SAPbouiCOM.EditText) MTXATTCH.Columns.Item("CLATTACH").Cells.Item(lastRow).Specific).Value = destinationFile;

                MTXATTCH.FlushToDataSource();

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                Application.SBO_Application.StatusBar.SetText("File uploaded successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        

        private void ETGRDEQTY_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            try
            {
                oform.Freeze(true);
                QuantityCalculation(ref oform);
                oform.Freeze(false);
            }
            catch (Exception ex)
            {
                oform.Freeze(false);
            }
        }

        private void ETGRDDQTY_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            try
            {
                oform.Freeze(true);
                QuantityCalculation(ref oform);
                oform.Freeze(false);
            }
            catch (Exception ex)
            {
                oform.Freeze(false);
            }

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");

            string trnsfrId = oform.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_TRNSFRID", 0).Trim();
            if (!string.IsNullOrEmpty(trnsfrId))
            {
                oform.Items.Item("BTSTKTRNS").Visible = false;
            }
            else
            {
                oform.Items.Item("BTSTKTRNS").Visible = true;
            }

        }

        private void BTSTKTRNS_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbobsCOM.Company oCompany = Global.oComp;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");

            bool success = false;

            try
            {
                Global.G_UI_Application.StatusBar.SetText(
                    "Processing... Please wait",
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Warning
                );

                oForm.Freeze(true);

                if (!oCompany.InTransaction)
                    oCompany.StartTransaction();

                AddInventoryTransfer(ref oForm);
                AddProductionOrder(ref oForm);

                if (oCompany.InTransaction)
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                if(oCompany.InTransaction == true)
                {
                   success = true;
                }
                else
                {
                    success = false;
                }


            }
            catch (Exception ex)
            {
                if (oCompany.InTransaction)
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                Global.G_UI_Application.StatusBar.SetText(
                    "Error. Nothing saved: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Long,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                oForm.Freeze(false);

                if (success)
                {
                    Global.G_UI_Application.StatusBar.SetText("Both documents saved successfully", SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    Application.SBO_Application.ActivateMenuItem("1304");
                }
            }

        }

       

        private void ETGRDCQTY_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            try
            {
                oform.Freeze(true);
                QuantityCalculation(ref oform);
                oform.Freeze(false);
            }
            catch (Exception ex)
            {
                oform.Freeze(false);
            }

        }

       

        private void ETGRDBQTY_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            try
            {
                oform.Freeze(true);
                QuantityCalculation(ref oform);
                oform.Freeze(false);
            }
            catch (Exception ex)
            {
                oform.Freeze(false);
            }

        }

        private void ETGRDAQTY_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            try
            {
                oform.Freeze(true);
                QuantityCalculation(ref oform);
                oform.Freeze(false);
            }
            catch(Exception ex)
            {
                oform.Freeze(false);
            }

        }
        private void MTX01_LostFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oform.Items.Item("MTX01").Specific;
                string clResultStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLRESULT").Cells.Item(pVal.Row).Specific).Value;
                if (!decimal.TryParse(clResultStr, out decimal clResult)) return;
                decimal ClLowValue = Convert.ToDecimal(((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLLOWVALUE").Cells.Item(pVal.Row).Specific).Value);
                decimal clUprValue = Convert.ToDecimal(((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLUPRVALUE").Cells.Item(pVal.Row).Specific).Value);

                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item("CLSTATUS").Cells.Item(pVal.Row).Specific;

                if (clResult >= ClLowValue && clResult <= clUprValue)
                {
                    oCombo.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else
                {
                    oCombo.Select("R", SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
            }
            catch(Exception ex)
            {

            }

        }


        private void Form_DataAddAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oForm = Global.G_UI_Application.Forms.Item(pVal.FormUID);
            AddInventoryTransfer(ref oForm);
            AddProductionOrder(ref oForm);
        }

        private void Add_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Global.G_UI_Application.Forms.Item(pVal.FormUID);
            if (pVal.ActionSuccess)
            {
             //   ClearFormFields(oForm);
            }

        }

        private void Add_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
            if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oform.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || pVal.CharPressed == 13)
            {
                ClearBlankRow(ref oform);
                ValidateForm(ref oform, ref BubbleEvent);
               
            }

        }

        private static void ClearBlankRow(ref SAPbouiCOM.Form pForm)
        {

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)pForm.Items.Item("MTX01").Specific;
            for (int i = oMatrix.RowCount; i >= 1; i--)
            {
                string cCombo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCHRGNAME").Cells.Item(i).Specific).Value;
                if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLCHRGNAME").Cells.Item(i).Specific).Value == "")
                    oMatrix.DeleteRow(i);
                else
                    break;
            }

        }

        private static void AddInventoryTransfer(ref SAPbouiCOM.Form pForm)
        {
            pForm.Freeze(true);
            SAPbobsCOM.StockTransfer oInventoryTransfer = (SAPbobsCOM.StockTransfer)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer);
            try
            {
                string fromWarehouse = ((SAPbouiCOM.IEditText)pForm.Items.Item("ETQCWHS").Specific).Value;
                oInventoryTransfer.FromWarehouse = fromWarehouse;
                string itemCode = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_ITEMCODE", 0);
                //string batchNumber = ((SAPbouiCOM.IEditText)pForm.Items.Item("ETBATCH").Specific).Value;

                double qcQty =Convert.ToDouble(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_ACPTQTY", 0));

                double scrapQty =Convert.ToDouble(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_SCRPQTY", 0));

                string acceptWarehouse = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_ACPTWHS", 0);
                oInventoryTransfer.ToWarehouse = acceptWarehouse;
                string scrapWarehouse = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_SCRPWHS", 0);


                List<(string ToWarehouse, double Quantity)> DetailsValue = new List<(string, double)>
                {
                    (acceptWarehouse, qcQty), (scrapWarehouse, scrapQty)
                };


                foreach (var target in DetailsValue)
                {
                    if (target.Quantity == 0)
                        continue;
                    oInventoryTransfer.Lines.ItemCode = itemCode;
                    oInventoryTransfer.Lines.FromWarehouseCode = fromWarehouse;
                    oInventoryTransfer.Lines.WarehouseCode = target.ToWarehouse;
                    oInventoryTransfer.Lines.Quantity = target.Quantity;

                    //SAPbobsCOM.BatchNumbers oBatch = oInventoryTransfer.Lines.BatchNumbers;
                    //oBatch.BatchNumber = batchNumber;
                    //oBatch.Quantity = target.Quantity;
                    //oBatch.Add();

                    oInventoryTransfer.Lines.Add();
                }

                if (oInventoryTransfer.Add() != 0)
                {
                    throw new Exception(Global.oComp.GetLastErrorDescription());
                    
                }
                else
                {
                    SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    int TransferEntry = int.Parse(Global.oComp.GetNewObjectKey());
                    Global.GFunc.ShowSuccess("Inventory Transfer Posted Successfully. DocEntry: " + TransferEntry);
                    string DocEntry = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("DocEntry", 0);
                    string qStr = "CALL SP_UpdateQCInventoryTransfer(" + DocEntry + ", " + TransferEntry + ")";
                    rSet.DoQuery(qStr);
                    pForm.Freeze(false);

                    //if (pForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    //{
                    //    Global.G_UI_Application.ActivateMenuItem("1304");
                    //}
                }
            }
            catch (Exception ex)
            {
                pForm.Freeze(false);
                Global.GFunc.ShowError("Error: " + ex.Message);
            }
            finally
            {
                pForm.Freeze(false);
            }

        }

        private static void AddProductionOrder(ref SAPbouiCOM.Form pForm)
        {

            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
            string batch = ((SAPbouiCOM.IEditText)oform.Items.Item("ETBATCH").Specific).Value;
            string ItemA = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDAITM").Specific).Value;
            string ItemB = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDBITM").Specific).Value;
            string ItemC = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDCITM").Specific).Value;
            string ItemD = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDDITM").Specific).Value;
            string ItemE = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDEITM").Specific).Value;

            string WhsA = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDAWHS").Specific).Value;
            string WhsB = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDBWHS").Specific).Value;
            string WhsC = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDCWHS").Specific).Value;
            string WhsD = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDDWHS").Specific).Value;
            string WhsE = ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDEWHS").Specific).Value;

            double QtyA = double.TryParse(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDAQTY").Specific).Value.Trim(), out double gradeAQty) ? gradeAQty : 0;
            double QtyB = double.TryParse(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDBQTY").Specific).Value.Trim(), out double gradeBQty) ? gradeBQty : 0;
            double QtyC = double.TryParse(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDCQTY").Specific).Value.Trim(), out double gradeCQty) ? gradeCQty : 0;
            double QtyD = double.TryParse(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDDQTY").Specific).Value.Trim(), out double gradeDQty) ? gradeDQty : 0;
            double QtyE = double.TryParse(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDEQTY").Specific).Value.Trim(), out double gradeEQty) ? gradeEQty : 0;


            List<(string fgItemCode, string warehouse, double qty,int flag)> fgItems = new List<(string, string, double,int)>
                {
                    (ItemA, WhsA, QtyA,1),
                    (ItemB, WhsB, QtyB,2),
                    (ItemC, WhsC, QtyC,3),
                    (ItemD, WhsD, QtyD,4),
                    (ItemE, WhsE, QtyE,5)
                };

            foreach (var item in fgItems)
            {
                if (item.qty > 0)
                {
                    try
                    {
                        if (!Global.oComp.InTransaction)
                            Global.oComp.StartTransaction();

                        int prodOrderDocEntry = -1;
                        string fgItemCode = item.fgItemCode;
                        string warehouse = item.warehouse;
                        double plannedQty = item.qty;

                        //Production Order
                        SAPbobsCOM.ProductionOrders oProdOrder = (SAPbobsCOM.ProductionOrders)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oProductionOrders);
                        oProdOrder.ItemNo = fgItemCode;
                        oProdOrder.PlannedQuantity = plannedQty;
                        oProdOrder.ProductionOrderType = SAPbobsCOM.BoProductionOrderTypeEnum.bopotSpecial;
                        oProdOrder.PostingDate = DateTime.Today;
                        oProdOrder.DueDate = DateTime.Today.AddDays(3);
                        oProdOrder.Warehouse = warehouse;
                        oProdOrder.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposPlanned;


                        oProdOrder.Lines.ItemNo = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_ITEMCODE", 0);
                        oProdOrder.Lines.Warehouse = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_QCWHS", 0);
                        oProdOrder.Lines.ProductionOrderIssueType = SAPbobsCOM.BoIssueMethod.im_Manual;
                        oProdOrder.Lines.PlannedQuantity = plannedQty;

                        if (oProdOrder.Add() != 0)
                        {
                            Global.oComp.GetLastError(out int errCode, out string errMsg);
                            Global.GFunc.ShowError($"Production Order Error for {fgItemCode}: {errMsg}");
                        }
                        else
                        {

                            prodOrderDocEntry = int.Parse(Global.oComp.GetNewObjectKey());
                            oProdOrder.GetByKey(prodOrderDocEntry);
                            oProdOrder.ProductionOrderStatus = SAPbobsCOM.BoProductionOrderStatusEnum.boposReleased;
                            oProdOrder.Update();
                        }




                        // 2. ISSUE FOR PRODUCTION

                        SAPbobsCOM.Documents oIssue = (SAPbobsCOM.Documents)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                        oIssue.DocDate = DateTime.Today;
                        oIssue.BPL_IDAssignedToInvoice = 1;

                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string query = $"SELECT \"LineNum\", \"ItemCode\",\"PlannedQty\", \"wareHouse\" FROM \"WOR1\" WHERE \"DocEntry\" = {prodOrderDocEntry}";
                        oRS.DoQuery(query);

                        int line = 0;
                        while (!oRS.EoF)
                        {
                            if (line > 0) oIssue.Lines.Add();
                            oIssue.Lines.Quantity = Convert.ToDouble(oRS.Fields.Item("PlannedQty").Value);
                            oIssue.Lines.WarehouseCode = oRS.Fields.Item("Warehouse").Value.ToString();
                            oIssue.Lines.BaseEntry = prodOrderDocEntry;
                            oIssue.Lines.BaseLine = Convert.ToInt32(oRS.Fields.Item("LineNum").Value);
                            oIssue.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oProductionOrders;


                            //SAPbobsCOM.BatchNumbers oBatch1 = oIssue.Lines.BatchNumbers;
                            //oBatch1.BatchNumber = batch;
                            //oBatch1.Quantity = Convert.ToDouble(oRS.Fields.Item("PlannedQty").Value);
                            //oBatch1.Add();

                            line++;
                            oRS.MoveNext();
                        }

                        if (oIssue.Add() != 0)
                        {
                            Global.oComp.GetLastError(out int errCode, out string errMsg);
                            Global.GFunc.ShowError($"ISSUE FOR PRODUCTION Error for {fgItemCode}: {errMsg}");
                        }


                        // 2. RECEIPT FROM PRODUCTION
                        SAPbobsCOM.Documents oReceipt = (SAPbobsCOM.Documents)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                        oReceipt.DocDate = DateTime.Today;
                        oReceipt.BPL_IDAssignedToInvoice = 1;


                        oReceipt.Lines.Quantity = plannedQty;
                        oReceipt.Lines.WarehouseCode = warehouse;
                        oReceipt.Lines.BaseEntry = prodOrderDocEntry;
                        //oReceipt.Lines.BaseLine = 0; Do not use this for Receive From Production
                        oReceipt.Lines.BaseType = 202;

                        //SAPbobsCOM.BatchNumbers oBatch = oReceipt.Lines.BatchNumbers;
                        //oBatch.BatchNumber = oBatch.BatchNumber = DateTime.Now.ToString("ddMMyyyyHHmmss");
                        //oBatch.Quantity = plannedQty;
                        //oBatch.Add();

                        if (oReceipt.Add() != 0)
                        {
                            Global.oComp.GetLastError(out int errCode, out string errMsg);
                            Global.GFunc.ShowError($"RECEIPT FROM PRODUCTION Error for {fgItemCode}: {errMsg}");
                        }

                        if (Global.oComp.InTransaction)
                            Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                        if (item.flag == 1)
                        {
                            SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string DocEntry = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("DocEntry", 0);
                            string qStr = $"UPDATE \"@FIL_DH_INSPDECN\" AS H " +
                                          $"SET " +
                                          $"H.\"U_PRDNTRYA\" = (SELECT W.\"DocEntry\" FROM \"OWOR\" W WHERE W.\"DocEntry\" = " + prodOrderDocEntry + ") " +
                                          $"WHERE H.\"DocEntry\" = " + DocEntry + "";
                            rSet.DoQuery(qStr);
                        }

                        else if (item.flag == 2)
                        {
                            SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string DocEntry = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("DocEntry", 0);
                            string qStr = $"UPDATE \"@FIL_DH_INSPDECN\" AS H " +
                                          $"SET " +
                                          $"H.\"U_PRDNTRYB\" = (SELECT W.\"DocEntry\" FROM \"OWOR\" W WHERE W.\"DocEntry\" = " + prodOrderDocEntry + ") " +
                                          $"WHERE H.\"DocEntry\" = " + DocEntry + "";
                            rSet.DoQuery(qStr);
                        }

                        else if (item.flag == 3)
                        {
                            SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string DocEntry = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("DocEntry", 0);
                            string qStr = $"UPDATE \"@FIL_DH_INSPDECN\" AS H " +
                                          $"SET " +
                                          $"H.\"U_PRDNTRYC\" = (SELECT W.\"DocEntry\" FROM \"OWOR\" W WHERE W.\"DocEntry\" = " + prodOrderDocEntry + ") " +
                                          $"WHERE H.\"DocEntry\" = " + DocEntry + "";
                            rSet.DoQuery(qStr);
                        }

                        else if (item.flag == 4)
                        {
                            SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string DocEntry = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("DocEntry", 0);
                            string qStr = $"UPDATE \"@FIL_DH_INSPDECN\" AS H " +
                                          $"SET " +
                                          $"H.\"U_PRDNTRYD\" = (SELECT W.\"DocEntry\" FROM \"OWOR\" W WHERE W.\"DocEntry\" = " + prodOrderDocEntry + ") " +
                                          $"WHERE H.\"DocEntry\" = " + DocEntry + "";
                            rSet.DoQuery(qStr);
                        }

                        else if (item.flag == 5)
                        {
                            SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            string DocEntry = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("DocEntry", 0);
                            string qStr = $"UPDATE \"@FIL_DH_INSPDECN\" AS H " +
                                          $"SET " +
                                          $"H.\"U_PRDNTRYE\" = (SELECT W.\"DocEntry\" FROM \"OWOR\" W WHERE W.\"DocEntry\" = " + prodOrderDocEntry + ") " +
                                          $"WHERE H.\"DocEntry\" = " + DocEntry + "";
                            rSet.DoQuery(qStr);
                        }
                    }
                    catch (Exception ex)
                    {
                        if (Global.oComp.InTransaction)
                            Global.oComp.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        Global.GFunc.ShowError("Process failed: " + ex.Message);
                    }

                }

            }

        }
        private bool ValidateForm(ref SAPbouiCOM.Form pForm, ref bool BubbleEvent)
        {
            double ActQty = Convert.ToDouble(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_ACTQTY", 0));
            double qcQty = Convert.ToDouble(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_QCQTY", 0));
            //string Result = pForm.DataSources.DBDataSources.Item("@FIL_DR_INSPDECN").GetValue("U_RESULT", 0).Trim();
            double ACPTQTY = Convert.ToDouble(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_ACPTQTY", 0));

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)pForm.Items.Item("MTX01").Specific;

            if (qcQty > ActQty)
            {
                Global.GFunc.ShowError("QC Quanity can't be gretter than Actual Quantity");
                pForm.ActiveItem = "ETQCQTY";
                return BubbleEvent = false;
            }
            else if (qcQty == 0)
            {
                Global.GFunc.ShowError("QC Quanity can't be 0");
                pForm.ActiveItem = "ETQCQTY";
                return BubbleEvent = false;    
            }
            else if (ACPTQTY < 0)
            {
                Global.GFunc.ShowError("Accepted Quanity can't be Negative");
                pForm.ActiveItem = "ETACPTQTY";
                return BubbleEvent = false;
            }

            double premium = string.IsNullOrWhiteSpace( pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDAQTY", 0))? 0 : Convert.ToDouble(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDAQTY", 0).Trim());
            string GRDAITM = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDAITM", 0).ToString();
            string GRDAWHS = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDAWHS", 0).ToString();
            if(GRDAITM != "" || GRDAWHS != "")
            {
                if(premium == 0)
                {
                    Global.GFunc.ShowError("Please entry premium Qty");
                    pForm.ActiveItem = "ETGRDAQTY";
                    return BubbleEvent = false;
                }
            }

            if (premium > 0)
            {
                if(GRDAITM == "")
                {
                    Global.GFunc.ShowError("Please select premium Item");
                    pForm.ActiveItem = "ETGRDAITM";
                    return BubbleEvent = false;
                }

                if (GRDAWHS == "")
                {
                    Global.GFunc.ShowError("Please select premium Warehouse");
                    pForm.ActiveItem = "ETGRDAWHS";
                    return BubbleEvent = false;
                }
            }else if (premium < 0)
            {
                Global.GFunc.ShowError("Premium Quanity can't be negative");
                pForm.ActiveItem = "ETGRDAQTY";
                return BubbleEvent = false;
            }

            double nonpremium = string.IsNullOrWhiteSpace(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDBQTY", 0)) ? 0 : Convert.ToDouble(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDBQTY", 0).Trim());
            string GRDBITM = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDBITM", 0).ToString();
            string GRDBWHS = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDBWHS", 0).ToString();

            if (GRDBITM != "" || GRDBWHS != "")
            {
                if (nonpremium == 0)
                {
                    Global.GFunc.ShowError("Please entry nonpremium Qty");
                    pForm.ActiveItem = "ETGRDBQTY";
                    return BubbleEvent = false;
                }
            }

            if (nonpremium > 0)
            {
                if (GRDBITM == "")
                {
                    Global.GFunc.ShowError("Please select non premium Item");
                    pForm.ActiveItem = "ETGRDBITM";
                    return BubbleEvent = false;
                }

                if (GRDBWHS == "")
                {
                    Global.GFunc.ShowError("Please select non premium Warehouse");
                    pForm.ActiveItem = "ETGRDBWHS";
                    return BubbleEvent = false;
                }
            }
            else if (nonpremium < 0)
            {
                Global.GFunc.ShowError("Non Premium Quanity can't be negative");
                pForm.ActiveItem = "ETGRDBQTY";
                return BubbleEvent = false;
            }

            double lot = string.IsNullOrWhiteSpace(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDCQTY", 0)) ? 0 : Convert.ToDouble(pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDCQTY", 0).Trim());
            string GRDCITM = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDCITM", 0).ToString();
            string GRDCWHS = pForm.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").GetValue("U_GRDCWHS", 0).ToString();
            if (GRDCITM != "" || GRDCWHS != "")
            {
                if (lot == 0)
                {
                    Global.GFunc.ShowError("Please entry lot Qty");
                    pForm.ActiveItem = "ETGRDCQTY";
                    return BubbleEvent = false;
                }
            }

            if (lot > 0)
            {
                if (GRDCITM == "")
                {
                    Global.GFunc.ShowError("Please select lot Item");
                    pForm.ActiveItem = "ETGRDCITM";
                    return BubbleEvent = false;
                }
                if (GRDCWHS == "")
                {
                    Global.GFunc.ShowError("Please select lot Warehouse");
                    pForm.ActiveItem = "ETGRDCWHS";
                    return BubbleEvent = false;
                }
            }
            else if (lot < 0)
            {
                Global.GFunc.ShowError("Lot Quanity can't be negative");
                pForm.ActiveItem = "ETGRDCQTY";
                return BubbleEvent = false;
            }


            for (int i = 0; i <= oMatrix.RowCount - 1; i++)
            {
                string resultVal = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLRESULT").Cells.Item(i+1).Specific).Value;

                if (string.IsNullOrEmpty(resultVal))
                {
                    Global.GFunc.ShowError($"Result is empty in row {i+1}. Please select a result.");
                    oMatrix.SetCellFocus(i+1, 6);
                    return BubbleEvent = false;
                }
            }


            //Qc Warehouse Location Check
            SAPbobsCOM.Recordset oRs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string qcWhs = ((SAPbouiCOM.IEditText)pForm.Items.Item("ETQCWHS").Specific).Value;
            string acceptWhs = ((SAPbouiCOM.IEditText)pForm.Items.Item("ETACPTWHS").Specific).Value;
            string WhsA = ((SAPbouiCOM.IEditText)pForm.Items.Item("ETGRDAWHS").Specific).Value;
            string WhsB = ((SAPbouiCOM.IEditText)pForm.Items.Item("ETGRDBWHS").Specific).Value;
            string WhsC = ((SAPbouiCOM.IEditText)pForm.Items.Item("ETGRDCWHS").Specific).Value;
            string scrapWhs = ((SAPbouiCOM.IEditText)pForm.Items.Item("ETSCRPWHS").Specific).Value;

            string qstr = $@"
                            SELECT COUNT(*) AS ""mismatchCount""
                            FROM OWHS
                            WHERE ""WhsCode"" IN ('{WhsA}','{WhsB}','{WhsC}','{acceptWhs}','{scrapWhs}')
                            AND ""Location"" <> (
                                    SELECT ""Location""
                                    FROM OWHS
                                    WHERE ""WhsCode"" = '{qcWhs}'
                                )
                            ";

            oRs.DoQuery(qstr);
            int mismatchCount = Convert.ToInt32(oRs.Fields.Item("mismatchCount").Value);

            if(mismatchCount > 0)
            {
                Global.GFunc.ShowError("QC Warehouse location must match the Transfer warehouse Location.");
                return BubbleEvent = false;
            }

            if (!ValidateWarehouseDefaultBin(pForm))
            {
                BubbleEvent = false;
                return BubbleEvent = false; ;
            }



            return BubbleEvent;
        }

        private bool ValidateWarehouseDefaultBin(SAPbouiCOM.Form pForm)
        {
            string[] whsCodes =
            {
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETQCWHS").Specific).Value.Trim(),
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETACPTWHS").Specific).Value.Trim(),
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETGRDAWHS").Specific).Value.Trim(),
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETGRDBWHS").Specific).Value.Trim(),
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETGRDCWHS").Specific).Value.Trim(),
                ((SAPbouiCOM.IEditText)pForm.Items.Item("ETSCRPWHS").Specific).Value.Trim()
             };

            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            foreach (string whs in whsCodes.Distinct())
            {
                if (string.IsNullOrEmpty(whs))
                    continue;

                string sql = $@"
                        SELECT ""WhsCode""
                        FROM ""OWHS""
                        WHERE ""WhsCode"" = '{whs.Replace("'", "''")}'
                          AND ""BinActivat"" = 'Y'
                          AND ""DftBinAbs"" IS NULL";

                rs.DoQuery(sql);

                if (!rs.EoF)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText($"Warehouse [{whs}] bin activated, but default bin is not set.",SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error );
                    return false;
                }
            }

            return true;
        }

        private void QuantityCalculation(ref SAPbouiCOM.Form oform)
        {
            double GradeAQty = Convert.ToDouble(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDAQTY").Specific).Value);
            double GradeBQty = Convert.ToDouble(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDBQTY").Specific).Value);
            double GradeCQty = Convert.ToDouble(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDCQTY").Specific).Value);
            double GradeDQty = Convert.ToDouble(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDDQTY").Specific).Value);
            double GradeEQty = Convert.ToDouble(((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDEQTY").Specific).Value);

            double QcQuantity = Convert.ToDouble(((SAPbouiCOM.IEditText)oform.Items.Item("ETQCQTY").Specific).Value);
            double ScrapQuanity = Convert.ToDouble(((SAPbouiCOM.IEditText)oform.Items.Item("ETSCRPQTY").Specific).Value);

            double Grade = GradeAQty + GradeBQty + GradeCQty + GradeDQty + GradeEQty;
            double Total = QcQuantity - ScrapQuanity;
            double AcceptQty = Total - Grade;
            ((SAPbouiCOM.IEditText)oform.Items.Item("ETACPTQTY").Specific).Value = AcceptQty.ToString();
            ((SAPbouiCOM.IEditText)oform.Items.Item("ETTACPTQTY").Specific).Value = Total.ToString();
           // ((SAPbouiCOM.IEditText)oform.Items.Item("ETGRDAQTY").Specific).Value = GradeAQty.ToString();
        }


        private void ETGRDAITM_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException();
            try
            {

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
                SAPbouiCOM.ChooseFromList ocfl = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item("CFL_OITM2");

                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                // 🔹 First Condition
                SAPbouiCOM.Condition oCon1 = oCons.Add();
                oCon1.Alias = "frozenFor";
                oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon1.CondVal = "N";
                // Apply conditions
                ocfl.SetConditions(oCons);

            }
            catch (Exception ex)
            {

            }
        }

        private void ETGRDBITM_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException();
            try
            {

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
                SAPbouiCOM.ChooseFromList ocfl = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item("CFL_OITM3");

                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                // 🔹 First Condition
                SAPbouiCOM.Condition oCon1 = oCons.Add();
                oCon1.Alias = "frozenFor";
                oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon1.CondVal = "N";
                // Apply conditions
                ocfl.SetConditions(oCons);

            }
            catch (Exception ex)
            {

            }
        }

        private void ETGRDCITM_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            // throw new System.NotImplementedException();
            try
            {

                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
                SAPbouiCOM.ChooseFromList ocfl = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item("CFL_OITM4");

                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                // 🔹 First Condition
                SAPbouiCOM.Condition oCon1 = oCons.Add();
                oCon1.Alias = "frozenFor";
                oCon1.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon1.CondVal = "N";
                // Apply conditions
                ocfl.SetConditions(oCons);

            }
            catch (Exception ex)
            {

            }
        }
        private void ETGRDAITM_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.DBDataSource dbs = oform.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN");
                SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string Uid = cflEvent.ChooseFromListUID;  // cfl unique id 

                SAPbouiCOM.DataTable dtbCFL = cflEvent.SelectedObjects;
                SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (!dtbCFL.IsEmpty)
                {
                    string ItemCode = dtbCFL.GetValue("ItemCode", 0).ToString();
                    string qStr = $"SELECT \"ItemCode\",\"ItemName\",\"DfltWH\" FROM OITM WHERE \"ItemCode\" ='" + ItemCode + "' ";
                    rSet.DoQuery(qStr);

                    if (!rSet.EoF)
                    {
                        string itemName = rSet.Fields.Item("ItemName").Value.ToString().Trim();
                        string dfltWhs = rSet.Fields.Item("DfltWH").Value.ToString().Trim();


                        dbs.SetValue("U_GRDAITM", 0, ItemCode);
                        dbs.SetValue("U_GRDAITMN", 0, itemName);
                        dbs.SetValue("U_GRDAWHS", 0, dfltWhs);

                       // oform.Items.Item("ETGRDAITM").Click();
                    }

                }
            }
            catch (Exception ex)
            {
                Global.GFunc.ShowError(ex.Message);
            }

        }
        private void ETGRDAITM_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.DBDataSource dbs = oform.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN");
                // Check source EditText (where key is pressed)
                 //string ETGRDAITM= dbs.GetValue("U_GRDAITM", 0).ToString();
                SAPbouiCOM.EditText ETGRDAITM =
                            (SAPbouiCOM.EditText)oform.Items.Item("ETGRDAITM").Specific;
                if (pVal.ItemUID == "ETGRDAITM" && ETGRDAITM.Value == "")
                {
                    // Example: Enter key
                    //if (pVal.CharPressed == 13)
                    //{
                        dbs.SetValue("U_GRDAITM", 0, "");
                        dbs.SetValue("U_GRDAITMN", 0, "");
                        dbs.SetValue("U_GRDAWHS", 0, "");

                        Global.GFunc.ShowSuccess("Target field cleared");
                    //}
                }
            }
            catch(Exception ex)
            {
                Global.GFunc.ShowError(ex.Message);
            }

        }
        private void ETGRDBITM_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.DBDataSource dbs = oform.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN");
                SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string Uid = cflEvent.ChooseFromListUID;  // cfl unique id

                SAPbouiCOM.DataTable dtbCFL = cflEvent.SelectedObjects;

                SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (!dtbCFL.IsEmpty)
                {
                    string ItemCode = dtbCFL.GetValue("ItemCode", 0).ToString();
                    string qStr = $"SELECT \"ItemCode\",\"ItemName\",\"DfltWH\" FROM OITM WHERE \"ItemCode\" ='" + ItemCode + "' ";
                    rSet.DoQuery(qStr);

                    if (!rSet.EoF)
                    {
                        string itemName = rSet.Fields.Item("ItemName").Value.ToString().Trim();
                        string dfltWhs = rSet.Fields.Item("DfltWH").Value.ToString().Trim();


                        dbs.SetValue("U_GRDBITM", 0, ItemCode);
                        dbs.SetValue("U_GRDBITMN", 0, itemName);
                        dbs.SetValue("U_GRDBWHS", 0, dfltWhs);

                        // oform.Items.Item("ETGRDAITM").Click();
                    }

                }
            }
            catch (Exception)
            {

            }

        }
        private void ETGRDBITM_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.DBDataSource dbs = oform.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN");
                // Check source EditText (where key is pressed)
                //string ETGRDAITM= dbs.GetValue("U_GRDAITM", 0).ToString();
                SAPbouiCOM.EditText ETGRDBITM =
                            (SAPbouiCOM.EditText)oform.Items.Item("ETGRDBITM").Specific;
                if (pVal.ItemUID == "ETGRDBITM" && ETGRDBITM.Value == "")
                {
                    // Example: Enter key
                    //if (pVal.CharPressed == 13)
                    //{
                    dbs.SetValue("U_GRDBITM", 0, "");
                    dbs.SetValue("U_GRDBITMN", 0, "");
                    dbs.SetValue("U_GRDBWHS", 0, "");

                    Global.GFunc.ShowSuccess("Target field cleared");
                    //}
                }
            }
            catch (Exception ex)
            {
                Global.GFunc.ShowError(ex.Message);
            }


        }


        private void ETGRDCITM_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.DBDataSource dbs = oform.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN");
                SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                string Uid = cflEvent.ChooseFromListUID;  // cfl unique id

                SAPbouiCOM.DataTable dtbCFL = cflEvent.SelectedObjects;

                SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (!dtbCFL.IsEmpty)
                {
                    string ItemCode = dtbCFL.GetValue("ItemCode", 0).ToString();
                    string qStr = $"SELECT \"ItemCode\",\"ItemName\",\"DfltWH\" FROM OITM WHERE \"ItemCode\" ='" + ItemCode + "' ";
                    rSet.DoQuery(qStr);

                    if (!rSet.EoF)
                    {
                        string itemName = rSet.Fields.Item("ItemName").Value.ToString().Trim();
                        string dfltWhs = rSet.Fields.Item("DfltWH").Value.ToString().Trim();


                        dbs.SetValue("U_GRDCITM", 0, ItemCode);
                        dbs.SetValue("U_GRDCITMN", 0, itemName);
                        dbs.SetValue("U_GRDCWHS", 0, dfltWhs);

                        // oform.Items.Item("ETGRDAITM").Click();
                    }

                }
                //if (!dtbCFL.IsEmpty)
                //{

                //    ((SAPbouiCOM.EditText)oform.Items.Item("ETGRDCITM").Specific).Value = dtbCFL.GetValue("ItemCode", 0).ToString();
                //    ((SAPbouiCOM.EditText)oform.Items.Item("ETGRCITMN").Specific).Value = dtbCFL.GetValue("ItemName", 0).ToString();
                //    oform.Items.Item("ETGRDCITM").Click();
                //    ((SAPbouiCOM.EditText)oform.Items.Item("ETGRDCWHS").Specific).Value = dtbCFL.GetValue("DfltWH", 0).ToString();
                //}
            }
            catch (Exception)
            {

            }

        }

        private void ETGRDCITM_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
                SAPbouiCOM.DBDataSource dbs = oform.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN");
                SAPbouiCOM.EditText ETGRDCITM =
                            (SAPbouiCOM.EditText)oform.Items.Item("ETGRDCITM").Specific;
                if (pVal.ItemUID == "ETGRDCITM" && ETGRDCITM.Value == "")
                {
                    // Example: Enter key
                    //if (pVal.CharPressed == 13)
                    //{
                    dbs.SetValue("U_GRDCITM", 0, "");
                    dbs.SetValue("U_GRDCITMN", 0, "");
                    dbs.SetValue("U_GRDCWHS", 0, "");

                    Global.GFunc.ShowSuccess("Target field cleared");
                    //}
                }
            }
            catch (Exception ex)
            {
                Global.GFunc.ShowError(ex.Message);
            }

        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.LinkedButton LinkedButton1;
        private SAPbouiCOM.LinkedButton LinkedButton2;
        private SAPbouiCOM.LinkedButton LinkedButton3;
        private SAPbouiCOM.LinkedButton LinkedButton4;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.LinkedButton LinkedButton5;
        private SAPbouiCOM.LinkedButton LinkedButton6;
        private SAPbouiCOM.LinkedButton LinkedButton7;
        private SAPbouiCOM.LinkedButton LinkedButton8;
        private SAPbouiCOM.LinkedButton LinkedButton9;
        private SAPbouiCOM.LinkedButton LinkedButton10;
        private SAPbouiCOM.LinkedButton LinkedButton11;
        private SAPbouiCOM.LinkedButton LinkedButton12;
        private SAPbouiCOM.LinkedButton LinkedButton13;
        private SAPbouiCOM.LinkedButton LinkedButton14;
        private SAPbouiCOM.LinkedButton LinkedButton15;
        private SAPbouiCOM.LinkedButton LinkedButton16;
        private SAPbouiCOM.LinkedButton LinkedButton17;
        private SAPbouiCOM.LinkedButton LinkedButton18;
        private SAPbouiCOM.LinkedButton LinkedButton19;
        private SAPbouiCOM.LinkedButton LinkedButton20;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.EditText EditText8;
        private SAPbouiCOM.EditText EditText9;
        private SAPbouiCOM.EditText EditText10;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button BTSTKTRNS;
    }
}
