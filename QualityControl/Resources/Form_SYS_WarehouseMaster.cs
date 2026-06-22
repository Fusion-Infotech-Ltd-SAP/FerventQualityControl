using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;

namespace QualityControl.Resources
{
    class Form_SYS_WarehouseMaster
    {
        bool ValidatioStart = false;
        bool CFLStart = false;
        public Form_SYS_WarehouseMaster()
        {
            Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
        }
        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //Warehouse FormTypeEx 62
                if (pVal.FormTypeEx == "62" && pVal.EventType != SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    SAPbouiCOM.Form oform = Application.SBO_Application.Forms.GetFormByTypeAndCount(pVal.FormType, pVal.FormTypeCount);

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD && pVal.BeforeAction)
                    {
                        // Assuming 'oform' is your SAPbouiCOM.Form object
                        SAPbouiCOM.Item GlnStatic = oform.Items.Item("1470002039");
                        SAPbouiCOM.Item GlnEditText = oform.Items.Item("1470002038");

                        // For Wharehouse Type
                        SAPbouiCOM.Item StWhsType = oform.Items.Add("STWHSTYPE", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        StWhsType.Top = GlnStatic.Top - 150;
                        StWhsType.Height = GlnStatic.Height;
                        StWhsType.Width = GlnStatic.Width;
                        StWhsType.Left = GlnStatic.Left;
                        StWhsType.FromPane = 1;
                        StWhsType.ToPane = 1;
                        StWhsType.LinkTo = "CBWHSTYPE";
                        ((SAPbouiCOM.StaticText)StWhsType.Specific).Caption = "Warehouse Type";

                        SAPbouiCOM.Item CBWhsType = oform.Items.Add("CBWHSTYPE", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                        ((SAPbouiCOM.ComboBox)CBWhsType.Specific).DataBind.SetBound(true, "OWHS", "U_WHSTYPE");
                        CBWhsType.Top = GlnEditText.Top - 150;
                        CBWhsType.Height = GlnEditText.Height;
                        CBWhsType.Width = GlnEditText.Width;
                        CBWhsType.Left = GlnEditText.Left;
                        CBWhsType.DisplayDesc = true;
                        CBWhsType.FromPane = 1;
                        CBWhsType.ToPane = 1;
                        CBWhsType.LinkTo = "STWHSTYPE";


                        // For Quality
                        SAPbouiCOM.Item StQuality = oform.Items.Add("STQUALITY", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        StQuality.Top = GlnStatic.Top - 130;
                        StQuality.Height = GlnStatic.Height;
                        StQuality.Width = GlnStatic.Width;
                        StQuality.Left = GlnStatic.Left;
                        StQuality.FromPane = 1;
                        StQuality.ToPane = 1;
                        StQuality.LinkTo = "ETQUALITY";
                        ((SAPbouiCOM.StaticText)StQuality.Specific).Caption = "Quality";

                        SAPbouiCOM.Item EtQuality = oform.Items.Add("ETQUALITY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        ((SAPbouiCOM.EditText)EtQuality.Specific).DataBind.SetBound(true, "OWHS", "U_QUALITY");
                        EtQuality.Top = GlnEditText.Top - 130;
                        EtQuality.Height = GlnEditText.Height;
                        EtQuality.Width = GlnEditText.Width;
                        EtQuality.Left = GlnEditText.Left;
                        EtQuality.FromPane = 1;
                        EtQuality.ToPane = 1;
                        EtQuality.LinkTo = "STQUALITY";

                        //// For LOT
                        //SAPbouiCOM.Item STLOT = oform.Items.Add("STLOT", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        //STLOT.Top = GlnStatic.Top - 110;
                        //STLOT.Height = GlnStatic.Height;
                        //STLOT.Width = GlnStatic.Width;
                        //STLOT.Left = GlnStatic.Left;
                        //STLOT.FromPane = 1;
                        //STLOT.ToPane = 1;
                        //STLOT.LinkTo = "ETLOT";
                        //((SAPbouiCOM.StaticText)STLOT.Specific).Caption = "Lot";

                        //SAPbouiCOM.Item ETLOT = oform.Items.Add("ETLOT", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        //((SAPbouiCOM.EditText)ETLOT.Specific).DataBind.SetBound(true, "OWHS", "U_LOT");
                        //ETLOT.Top = GlnEditText.Top - 110;
                        //ETLOT.Height = GlnEditText.Height;
                        //ETLOT.Width = GlnEditText.Width;
                        //ETLOT.Left = GlnEditText.Left;
                        //ETLOT.FromPane = 1;
                        //ETLOT.ToPane = 1;
                        //ETLOT.LinkTo = "STLOT";


                        //// For LACQUER
                        //SAPbouiCOM.Item STLACQUER = oform.Items.Add("STLACQUER", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        //STLACQUER.Top = GlnStatic.Top - 90;
                        //STLACQUER.Height = GlnStatic.Height;
                        //STLACQUER.Width = GlnStatic.Width;
                        //STLACQUER.Left = GlnStatic.Left;
                        //STLACQUER.FromPane = 1;
                        //STLACQUER.ToPane = 1;
                        //STLACQUER.LinkTo = "ETLACQUER";
                        //((SAPbouiCOM.StaticText)STLACQUER.Specific).Caption = "Lacquer";

                        //SAPbouiCOM.Item ETLACQUER = oform.Items.Add("ETLACQUER", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        //((SAPbouiCOM.EditText)ETLACQUER.Specific).DataBind.SetBound(true, "OWHS", "U_LACQUER");
                        //ETLACQUER.Top = GlnEditText.Top - 90;
                        //ETLACQUER.Height = GlnEditText.Height;
                        //ETLACQUER.Width = GlnEditText.Width;
                        //ETLACQUER.Left = GlnEditText.Left;
                        //ETLACQUER.FromPane = 1;
                        //ETLACQUER.ToPane = 1;
                        //ETLACQUER.LinkTo = "STLACQUER";


                        //// For ECONOMIC
                        //SAPbouiCOM.Item STECONOMIC = oform.Items.Add("STECONOMIC", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        //STECONOMIC.Top = GlnStatic.Top - 70;
                        //STECONOMIC.Height = GlnStatic.Height;
                        //STECONOMIC.Width = GlnStatic.Width;
                        //STECONOMIC.Left = GlnStatic.Left;
                        //STECONOMIC.FromPane = 1;
                        //STECONOMIC.ToPane = 1;
                        //STECONOMIC.LinkTo = "ETECONOMIC";
                        //((SAPbouiCOM.StaticText)STECONOMIC.Specific).Caption = "Economic";

                        //SAPbouiCOM.Item ETECONOMIC = oform.Items.Add("ETECONOMIC", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        //((SAPbouiCOM.EditText)ETECONOMIC.Specific).DataBind.SetBound(true, "OWHS", "U_ECONOMIC");
                        //ETECONOMIC.Top = GlnEditText.Top - 70;
                        //ETECONOMIC.Height = GlnEditText.Height;
                        //ETECONOMIC.Width = GlnEditText.Width;
                        //ETECONOMIC.Left = GlnEditText.Left;
                        //ETECONOMIC.FromPane = 1;
                        //ETECONOMIC.ToPane = 1;
                        //ETECONOMIC.LinkTo = "STECONOMIC";


                        //// For BROKEN
                        //SAPbouiCOM.Item STBROKEN = oform.Items.Add("STBROKEN", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        //STBROKEN.Top = GlnStatic.Top - 50;
                        //STBROKEN.Height = GlnStatic.Height;
                        //STBROKEN.Width = GlnStatic.Width;
                        //STBROKEN.Left = GlnStatic.Left;
                        //STBROKEN.FromPane = 1;
                        //STBROKEN.ToPane = 1;
                        //STBROKEN.LinkTo = "ETBROKEN";
                        //((SAPbouiCOM.StaticText)STBROKEN.Specific).Caption = "Broken";

                        //SAPbouiCOM.Item ETBROKEN = oform.Items.Add("ETBROKEN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        //((SAPbouiCOM.EditText)ETBROKEN.Specific).DataBind.SetBound(true, "OWHS", "U_BROKEN");
                        //ETBROKEN.Top = GlnEditText.Top - 50;
                        //ETBROKEN.Height = GlnEditText.Height;
                        //ETBROKEN.Width = GlnEditText.Width;
                        //ETBROKEN.Left = GlnEditText.Left;
                        //ETBROKEN.FromPane = 1;
                        //ETBROKEN.ToPane = 1;
                        //ETBROKEN.LinkTo = "STBROKEN";

                        //// For BURST
                        //SAPbouiCOM.Item STBURST = oform.Items.Add("STBURST", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        //STBURST.Top = GlnStatic.Top - 30;
                        //STBURST.Height = GlnStatic.Height;
                        //STBURST.Width = GlnStatic.Width;
                        //STBURST.Left = GlnStatic.Left;
                        //STBURST.FromPane = 1;
                        //STBURST.ToPane = 1;
                        //STBURST.LinkTo = "ETBURST";
                        //((SAPbouiCOM.StaticText)STBURST.Specific).Caption = "Burst";

                        //SAPbouiCOM.Item ETBURST = oform.Items.Add("ETBURST", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        //((SAPbouiCOM.EditText)ETBURST.Specific).DataBind.SetBound(true, "OWHS", "U_BURST");
                        //ETBURST.Top = GlnEditText.Top - 30;
                        //ETBURST.Height = GlnEditText.Height;
                        //ETBURST.Width = GlnEditText.Width;
                        //ETBURST.Left = GlnEditText.Left;
                        //ETBURST.FromPane = 1;
                        //ETBURST.ToPane = 1;
                        //ETBURST.LinkTo = "STBURST";

                        // For SCRAP
                        SAPbouiCOM.Item STSCRAP = oform.Items.Add("STSCRAP", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        STSCRAP.Top = GlnStatic.Top - 10;
                        STSCRAP.Height = GlnStatic.Height;
                        STSCRAP.Width = GlnStatic.Width;
                        STSCRAP.Left = GlnStatic.Left;
                        STSCRAP.FromPane = 1;
                        STSCRAP.ToPane = 1;
                        STSCRAP.LinkTo = "ETSCRAP";
                        ((SAPbouiCOM.StaticText)STSCRAP.Specific).Caption = "Scrap";

                        SAPbouiCOM.Item ETSCRAP = oform.Items.Add("ETSCRAP", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        ((SAPbouiCOM.EditText)ETSCRAP.Specific).DataBind.SetBound(true, "OWHS", "U_SCRAP");
                        ETSCRAP.Top = GlnEditText.Top - 10;
                        ETSCRAP.Height = GlnEditText.Height;
                        ETSCRAP.Width = GlnEditText.Width;
                        ETSCRAP.Left = GlnEditText.Left;
                        ETSCRAP.FromPane = 1;
                        ETSCRAP.ToPane = 1;
                        ETSCRAP.LinkTo = "STSCRAP";

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
