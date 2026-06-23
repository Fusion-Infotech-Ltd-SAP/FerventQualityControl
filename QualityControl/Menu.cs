using QualityControl.Resources;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace QualityControl
{
    class Menu
    {
        public void BasicStart()
        {
            CompanyConnection(); //1)Company connection 
            CreateMainMenu("43520", "FIL_MN_QC", "Quality Control", 15, 2, false);//parent 2 step
            //CreateMainMenu("FIL_MN_TARGET", "FIL_MASTER", "Master", 0, 2, false);
            //String Menu
            CreateMainMenu("FIL_MN_QC", "FIL_QINSPMAS", "Inspection Parameter Master", 0, 1, false);
            CreateMainMenu("FIL_MN_QC", "FIL_CHARMGRPS", "Characteristic Master Group", 1, 1, false);
            CreateMainMenu("FIL_MN_QC", "FIL_INSPLAN", "Inspection Plan", 2, 1, false);
            CreateMainMenu("FIL_MN_QC", "FIL_CHKSLCTN", "Inspection Check Selection", 3, 1, false);
            CreateMainMenu("FIL_MN_QC", "FIL_PNDLIST", "QC Checked Stock Not Transfer", 4, 1, false);
            CreateMainMenu("FIL_MN_QC", "FIL_INSPDECN", "Inspection Decision", 5, 1, false);
        
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "FIL_QINSPMAS")
                {
                    string formUID = "FIL_FRM_MH_QINSPMAS"; // Unique ID for the form
                                                             // Check if the form is already open
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form is already open.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }
                    Form_UDO_InspectionParam activeForm = new Form_UDO_InspectionParam();
                    activeForm.Show();
                    SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_MH_QINSPMAS");
                    oform.Items.Item("ETCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                    oform.Items.Item("ETCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                    SetNextCode(oform, "ETCODE", "FIL_MH_QINSPMAS");
                }
               
                else if (pVal.BeforeAction && pVal.MenuUID == "FIL_CHARMGRPS")
                {
                    string formUID = "FIL_FRM_MH_CHARMGRPS"; // Unique ID for the form
                                                            // Check if the form is already open
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form is already open.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }

                    From_UDO_CharacteristicMasterGroup activeForm = new From_UDO_CharacteristicMasterGroup();
                    activeForm.Show();
                    SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_MH_CHARMGRPS");
                    oform.Freeze(true);
                    oform.Items.Item("ETCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                    oform.Items.Item("ETCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                    SetNextCode(oform, "ETCODE", "FIL_MH_CHARMGRP");
                    SAPbouiCOM.DBDataSource DBDataSourceLine = (SAPbouiCOM.DBDataSource)oform.DataSources.DBDataSources.Item("@FIL_MR_CHARMGRP");
                    SAPbouiCOM.Matrix MTX_01 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX_01").Specific;
                    Global.GFunc.SetNewLine(MTX_01, DBDataSourceLine, 1, "");

                    string sqlQuerybpl = string.Format("SELECT {0}Code{0},{0}Name{0}  FROM {0}@FIL_MH_QINSPMAS{0}", '"');


                    SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rSet.DoQuery(sqlQuerybpl);
                    MTX_01.Columns.Item("CLCHARCODE").ValidValues.Add("", "");
                    while (!rSet.EoF)
                    {
                        MTX_01.Columns.Item("CLCHARCODE").ValidValues.Add(rSet.Fields.Item("Code").Value.ToString(), rSet.Fields.Item("Name").Value.ToString());
                        rSet.MoveNext();
                    }

                    oform.Freeze(false);

                }


                else if (pVal.BeforeAction && pVal.MenuUID == "FIL_INSPLAN")
                {
                    string formUID = "FIL_FRM_MH_INSPLAN"; // Unique ID for the form
                                                             // Check if the form is already open
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form is already open.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }
                    Form_UDO_QualityInspection activeForm = new Form_UDO_QualityInspection();
                    activeForm.Show();

                    string qStr = "";
                    SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_MH_INSPLAN");
                    oform.Items.Item("CBITMGRPCD").Enabled = true;

                    SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    qStr = string.Format("SELECT {0}ItmsGrpCod{0}, {0}ItmsGrpNam{0} FROM {0}OITB{0} ORDER BY {0}ItmsGrpCod{0}",'"');
                    rSet.DoQuery(qStr);
                    while (!rSet.EoF)
                    {
                        ((SAPbouiCOM.ComboBox)oform.Items.Item("CBITMGRPCD").Specific).ValidValues.Add(rSet.Fields.Item("ItmsGrpCod").Value.ToString(), rSet.Fields.Item("ItmsGrpNam").Value.ToString());
                        rSet.MoveNext();
                    }
                    oform.Items.Item("CBSUBGRP").Enabled = true;
                    string qStr1 = "";
                    qStr1 = string.Format("SELECT distinct {0}U_SUBGRP{0}  FROM {0}OITM{0}", '"');
                    rSet.DoQuery(qStr1);
                    while (!rSet.EoF)
                    {
                        ((SAPbouiCOM.ComboBox)oform.Items.Item("CBSUBGRP").Specific).ValidValues.Add(rSet.Fields.Item("U_SUBGRP").Value.ToString(), rSet.Fields.Item("U_SUBGRP").Value.ToString());
                        rSet.MoveNext();
                    }
                    oform.Items.Item("CBHICKNSS").Enabled = true;
                    string qStr2 = "";
                  //  qStr2 = string.Format("SELECT distinct {0}U_Thickness{0}  FROM {0}OITM{0}", '"');
                    //Fervent
                    qStr2 = string.Format("SELECT distinct {0}U_THICKNESS{0}  FROM {0}OITM{0}", '"');
                    
                    rSet.DoQuery(qStr2);
                    while (!rSet.EoF)
                    {
                     // ((SAPbouiCOM.ComboBox)oform.Items.Item("CBHICKNSS").Specific).ValidValues.Add(rSet.Fields.Item("U_Thickness").Value.ToString(), rSet.Fields.Item("U_Thickness").Value.ToString());
                        //Fervent
                     ((SAPbouiCOM.ComboBox)oform.Items.Item("CBHICKNSS").Specific).ValidValues.Add(rSet.Fields.Item("U_THICKNESS").Value.ToString(), rSet.Fields.Item("U_THICKNESS").Value.ToString());
                        rSet.MoveNext();
                    }


                    oform.Items.Item("TAB3").Click();
                    SAPbouiCOM.Matrix MTX03 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX03").Specific;
                    //SAPbouiCOM.DBDataSource DBDataSourceLine = (SAPbouiCOM.DBDataSource)oform.DataSources.DBDataSources.Item("@FIL_MR_INPLNIM");
                    //Global.GFunc.SetNewLineForEditText(MTX03, DBDataSourceLine, 1, "");
                    oform.Items.Item("TAB1").Click();


                }

                else if (pVal.BeforeAction && pVal.MenuUID == "FIL_CHKSLCTN")
                {
                    string formUID = "FIL_FRM_MH_CHKSLC"; // Unique ID for the form
                                                           // Check if the form is already open
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form is already open.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }
                    Form_UDO_InspectionCheckSelection activeForm = new Form_UDO_InspectionCheckSelection();
                    activeForm.Show();

                    string qStr = "";
                    SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_MH_CHKSLC");
                    SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    qStr = string.Format("SELECT {0}ItmsGrpCod{0},{0}ItmsGrpNam{0}  FROM {0}OITB{0}", '"');
                    rSet.DoQuery(qStr);
                    while (!rSet.EoF)
                    {
                        ((SAPbouiCOM.ComboBox)oform.Items.Item("CBITMGRPCD").Specific).ValidValues.Add(rSet.Fields.Item("ItmsGrpCod").Value.ToString(), rSet.Fields.Item("ItmsGrpNam").Value.ToString());
                        rSet.MoveNext();
                    }

                    qStr = string.Format("SELECT A.{0}DocumentType{0},A.{0}Description{0},A.{0}ObjType{0},A.{0}BaseType{0}  from {0}VW_QCDocument{0} A", '"');
                    rSet.DoQuery(qStr);
                    while (!rSet.EoF)
                    {
                        ((SAPbouiCOM.ComboBox)oform.Items.Item("CBDOCTYPE").Specific).ValidValues.Add(rSet.Fields.Item("ObjType").Value.ToString(), rSet.Fields.Item("Description").Value.ToString());
                        rSet.MoveNext();
                    }
                }


                else if (pVal.BeforeAction && pVal.MenuUID == "FIL_PNDLIST")
                {
                    string formUID = "FIL_FRM_MH_PNDLIST"; // Unique ID for the form
                                                          // Check if the form is already open
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form is already open.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }
                    TransferPendingList activeForm = new TransferPendingList();
                    activeForm.Show();

                    activeForm.loadPendingList();

                    //string qStr = "";
                    //SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_MH_CHKSLC");
                    //SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //qStr = string.Format("SELECT {0}ItmsGrpCod{0},{0}ItmsGrpNam{0}  FROM {0}OITB{0}", '"');
                    //rSet.DoQuery(qStr);
                    //while (!rSet.EoF)
                    //{
                    //    ((SAPbouiCOM.ComboBox)oform.Items.Item("CBITMGRPCD").Specific).ValidValues.Add(rSet.Fields.Item("ItmsGrpCod").Value.ToString(), rSet.Fields.Item("ItmsGrpNam").Value.ToString());
                    //    rSet.MoveNext();
                    //}

                    //qStr = string.Format("SELECT A.{0}DocumentType{0},A.{0}Description{0},A.{0}ObjType{0},A.{0}BaseType{0}  from {0}VW_QCDocument{0} A", '"');
                    //rSet.DoQuery(qStr);
                    //while (!rSet.EoF)
                    //{
                    //    ((SAPbouiCOM.ComboBox)oform.Items.Item("CBDOCTYPE").Specific).ValidValues.Add(rSet.Fields.Item("ObjType").Value.ToString(), rSet.Fields.Item("Description").Value.ToString());
                    //    rSet.MoveNext();
                    //}
                }

                else if (pVal.BeforeAction && pVal.MenuUID == "FIL_INSPDECN")
                {
                    string formUID = "FIL_FRM_DH_INSPDECN"; // Unique ID for the form
                                                          // Check if the form is already open
                    if (IsFormOpen(formUID))
                    {
                        Global.G_UI_Application.Forms.Item(formUID).Select();
                        Global.G_UI_Application.StatusBar.SetText("Form is already open.",
                            SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        return;
                    }

                    Form_UDO_InspectionDecision activeForm = new Form_UDO_InspectionDecision();
                    activeForm.Show();
                    SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
      
                    try
                    {
                        if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oform.Items.Item("CBSERIES").Specific;                                                                         //1 step do generate series:
                            Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_INSPDECN");
                            string oComValue = ocmb.Selected.Value;
                            long DocNo = oform.BusinessObject.GetNextSerialNumber(oComValue, "FIL_D_INSPDECN");

                            oform.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("DocNum", 0, DocNo.ToString());
                        }
                        oform.Items.Item("BTSTKTRNS").Visible = false;
                        //Current Date
                        SAPbouiCOM.EditText oedt = (SAPbouiCOM.EditText)oform.Items.Item("ETDOCDATE").Specific; //assign / define a edittext
                        oedt.Active = true;
                        oedt.String = "W";

                        SAPbouiCOM.Matrix MTX01 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX01").Specific;
                        MTX01.AutoResizeColumns();//define matrix
                    }

                    catch (Exception EX)
                    {

                    }


                }
                //Delete Row
                else if (!pVal.BeforeAction && pVal.MenuUID == "1293")
                {
                    SAPbouiCOM.Form pForm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = pForm.UniqueID.ToString();
                    
                    switch (formtype)
                    {
                        case "FIL_FRM_MH_INSPLAN":
                            {
                                try
                                {
                                    pForm.Freeze(true);
                                    SAPbouiCOM.DBDataSource oDataSource = pForm.DataSources.DBDataSources.Item("@FIL_MR_INPLNIM");
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)pForm.Items.Item("MTX03").Specific;
                                    Global.GFunc.DeleteRow(oMatrix, oDataSource);
                                    pForm.Freeze(false);
                                    
                                }
                                catch(Exception ex)
                                {
                                    pForm.Freeze(false);
                                }
                                break;
                            }
                            //case "FIL_FRM_MD_LOCATION":
                            //    {
                            //        break;
                            //    }
                    }
                }
                //Find
                else if (pVal.BeforeAction && pVal.MenuUID == "1281")
                {
                    SAPbouiCOM.Form ofrm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = ofrm.UniqueID.ToString();
                    switch (formtype)
                    {
                        case "FIL_FRM_MH_INSPLAN":
                            {
                                SAPbouiCOM.Form pForm = Application.SBO_Application.Forms.ActiveForm;

                                pForm.Items.Item("CBITMGRPCD").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                pForm.Items.Item("ETCHRGCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                pForm.Items.Item("ETCHRGNAME").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                pForm.Items.Item("CBACTIVE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                pForm.Items.Item("CBSUBGRP").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                pForm.Items.Item("CBHICKNSS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                                break;
                            }


                    }
                }
                //Add Form Mode Menu
                else if (!pVal.BeforeAction && pVal.MenuUID == "1282")
                {
                    SAPbouiCOM.Form ofrm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = ofrm.UniqueID.ToString();
                    switch (formtype)
                    {
                        case "FIL_FRM_MH_QINSPMAS":
                            {
                                SetNextCode(ofrm, "ETCODE", "FIL_MH_QINSPMAS");
                                break;
                            }
                              
                        case "FIL_FRM_MH_CHARMGRPS":
                            {
                                SetNextCode(ofrm, "ETCODE", "FIL_MH_CHARMGRP");
                                break;
                            }
                        case "FIL_FRM_MH_INSPLAN":
                            {
                                string formUID = "FIL_FRM_MH_INSPLAN"; // Unique ID for the form
                                string qStr = "";
                                SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_MH_INSPLAN");
                                oform.Items.Item("CBITMGRPCD").Enabled = true;
                                oform.Items.Item("CBSUBGRP").Enabled = true;
                                oform.Items.Item("CBHICKNSS").Enabled = true;
                                //SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                //qStr = string.Format("SELECT {0}ItmsGrpCod{0},{0}ItmsGrpNam{0}  FROM {0}OITB{0}", '"');
                                //rSet.DoQuery(qStr);
                                //while (!rSet.EoF)
                                //{
                                //    ((SAPbouiCOM.ComboBox)oform.Items.Item("CBITMGRPCD").Specific).ValidValues.Add(rSet.Fields.Item("ItmsGrpCod").Value.ToString(), rSet.Fields.Item("ItmsGrpNam").Value.ToString());
                                //    rSet.MoveNext();
                                //}
                                oform.Items.Item("TAB3").Click();
                                SAPbouiCOM.Matrix MTX03 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX03").Specific;
                                //SAPbouiCOM.DBDataSource DBDataSourceLine = (SAPbouiCOM.DBDataSource)oform.DataSources.DBDataSources.Item("@FIL_MR_INPLNIM");
                                //Global.GFunc.SetNewLineForEditText(MTX03, DBDataSourceLine, 1, "");
                                oform.Items.Item("TAB1").Click();
                                break;
                            }
                        //case "FIL_FRM_MH_CHKSLC":
                        //    {
                        //        string formUID = "FIL_FRM_MH_CHKSLC"; // Unique ID for the form
                        //        string qStr = "";
                        //        SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_MH_CHKSLC");
                        //        SAPbobsCOM.Recordset rSet = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //        qStr = string.Format("SELECT {0}ItmsGrpCod{0},{0}ItmsGrpNam{0}  FROM {0}OITB{0}", '"');
                        //        rSet.DoQuery(qStr);
                        //        while (!rSet.EoF)
                        //        {
                        //            ((SAPbouiCOM.ComboBox)oform.Items.Item("CBITMGRPCD").Specific).ValidValues.Add(rSet.Fields.Item("ItmsGrpCod").Value.ToString(), rSet.Fields.Item("ItmsGrpNam").Value.ToString());
                        //            rSet.MoveNext();
                        //        }

                        //        qStr = string.Format("SELECT A.{0}DocumentType{0},A.{0}Description{0},A.{0}ObjType{0},A.{0}BaseType{0}  from {0}VW_QCDocument{0} A", '"');
                        //        rSet.DoQuery(qStr);
                        //        while (!rSet.EoF)
                        //        {
                        //            ((SAPbouiCOM.ComboBox)oform.Items.Item("CBDOCTYPE").Specific).ValidValues.Add(rSet.Fields.Item("ObjType").Value.ToString(), rSet.Fields.Item("Description").Value.ToString());
                        //            rSet.MoveNext();
                        //        }
                        //        break;
                        //    }
                        //case "FIL_FRM_DH_INSPDECN":
                        //    {
                        //        string formUID = "FIL_FRM_DH_INSPDECN"; // Unique ID for the form
                        //        SAPbouiCOM.Form oform = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item("FIL_FRM_DH_INSPDECN");
                        //        try
                        //        {
                        //            if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        //            {
                        //                SAPbouiCOM.ComboBox ocmb = (SAPbouiCOM.ComboBox)oform.Items.Item("CBSERIES").Specific;                                                                         //1 step do generate series:
                        //                Global.GFunc.LoadComboBoxSeries(ocmb, "FIL_D_INSPDECN");
                        //                string oComValue = ocmb.Selected.Value;
                        //                long DocNo = oform.BusinessObject.GetNextSerialNumber(oComValue, "FIL_D_INSPDECN");

                        //                oform.DataSources.DBDataSources.Item("@FIL_DH_INSPDECN").SetValue("DocNum", 0, DocNo.ToString());
                        //            }

                        //            //Current Date
                        //            SAPbouiCOM.EditText oedt = (SAPbouiCOM.EditText)oform.Items.Item("ETDOCDATE").Specific; //assign / define a edittext
                        //            oedt.Active = true;
                        //            oedt.String = "W";

                        //            SAPbouiCOM.Matrix MTX01 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX01").Specific;
                        //            MTX01.AutoResizeColumns();//define matrix
                        //        }

                        //        catch (Exception EX)
                        //        {

                        //        }

                        //        break;
                        //    }

                    }
                }
                //First
                else if (!pVal.BeforeAction && pVal.MenuUID == "1288")
                {
                    SAPbouiCOM.Form ofrm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = ofrm.UniqueID.ToString();
                    switch (formtype)
                    {

                        case "FIL_FRM_DH_INSPDECN":
                            {
                                SAPbouiCOM.Form pForm = Application.SBO_Application.Forms.ActiveForm;
                                FormMode(pForm);
                                break;
                            }

                    }
                }
                //Previous
                else if (!pVal.BeforeAction && pVal.MenuUID == "1289")
                {
                    SAPbouiCOM.Form ofrm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = ofrm.UniqueID.ToString();
                    switch (formtype)
                    {

                        case "FIL_FRM_DH_INSPDECN":
                            {
                                SAPbouiCOM.Form pForm = Application.SBO_Application.Forms.ActiveForm;
                                FormMode(pForm);
                                break;
                            }
                    }
                }
                //Next
                else if (!pVal.BeforeAction && pVal.MenuUID == "1290")
                {
                    SAPbouiCOM.Form ofrm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = ofrm.UniqueID.ToString();
                    switch (formtype)
                    {

                        case "FIL_FRM_DH_INSPDECN":
                            {
                                SAPbouiCOM.Form pForm = Application.SBO_Application.Forms.ActiveForm;
                                FormMode(pForm);
                                break;
                            }
                    }
                }
                //Last
                else if (pVal.BeforeAction && pVal.MenuUID == "1291")
                {
                    SAPbouiCOM.Form ofrm = (SAPbouiCOM.Form)Application.SBO_Application.Forms.ActiveForm;
                    string formtype = ofrm.UniqueID.ToString();
                    switch (formtype)
                    {

                        case "FIL_FRM_DH_INSPDECN":
                            {
                                SAPbouiCOM.Form pForm = Application.SBO_Application.Forms.ActiveForm;
                                FormMode(pForm);
                                break;
                            }
                            

                    }
                }
                //Habib(2026_05_14)

                else if (!pVal.BeforeAction && pVal.MenuUID == "FIL_DUPL")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;

                    if (oForm.UniqueID == "FIL_FRM_MH_CHARMGRPS")
                    {
                        try
                        {
                            oForm.Freeze(true);

                            SAPbouiCOM.DBDataSource oHeader = oForm.DataSources.DBDataSources.Item("@FIL_MH_CHARMGRP");

                            SAPbouiCOM.DBDataSource oLine =  oForm.DataSources.DBDataSources.Item("@FIL_MR_CHARMGRP");

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MTX_01").Specific;

                            string oldCode = oHeader.GetValue("Code", 0).Trim();

                            if (string.IsNullOrEmpty(oldCode))
                            {
                                Application.SBO_Application.StatusBar.SetText("Please select an existing record first.",SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                return;
                            }

                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                         //   oForm.Items.Item("ETCODE").Enabled = false;

                            // Clear header values
                            oHeader.SetValue("Code", 0, "");
                            oHeader.SetValue("Name", 0, "");

                            SetNextCode(oForm, "ETCODE", "FIL_MH_CHARMGRP");

                            SAPbobsCOM.Recordset rs =(SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                            string safeOldCode = oldCode.Replace("'", "''");

                            string query = $@"SELECT ""U_CHARCODE"", ""U_CHARNAME"" FROM ""@FIL_MR_CHARMGRP"" WHERE ""Code"" = '{safeOldCode}' ORDER BY ""LineId"" ";

                            rs.DoQuery(query);
                            oLine.Clear();

                            int row = 0;

                            while (!rs.EoF)
                            {
                                oLine.InsertRecord(row);
                                oLine.SetValue("LineId", row, (row + 1).ToString());

                                oLine.SetValue("U_CHARCODE",row,rs.Fields.Item("U_CHARCODE").Value.ToString());
                                oLine.SetValue("U_CHARNAME",row,rs.Fields.Item("U_CHARNAME").Value.ToString());
                                row++;
                                rs.MoveNext();
                            }

                            oMatrix.LoadFromDataSource();

                            Global.GFunc.SetNewLine(oMatrix, oLine, 1, "");

                          //  oForm.Items.Item("ETCODE").Enabled = false;

                            Application.SBO_Application.StatusBar.SetText("Data duplicated successfully.",SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        }
                        catch (Exception ex)
                        {
                            Application.SBO_Application.StatusBar.SetText(ex.Message,SAPbouiCOM.BoMessageTime.bmt_Short,SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        }
                        finally
                        {
                            oForm.Freeze(false);
                        }
                    }

                    
                }


            }

            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }
        //private void GenerateCostVersion(SAPbouiCOM.Form oForm)
        //{
        //    try
        //    {
        //        string styleCode = ((SAPbouiCOM.EditText)oForm.Items.Item("ETSTYLCD").Specific).Value;
        //        string draftCode = ((SAPbouiCOM.EditText)oForm.Items.Item("ETDRFTSO").Specific).Value.Trim();

        //        string nextValue = string.Empty;
        //        SAPbobsCOM.Recordset oRS2 = (SAPbobsCOM.Recordset)
        //            Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //        string sql = $@"
        //            SELECT
        //               LPAD(TO_NVARCHAR(COUNT(*) + 1), 2, '0') AS ""RESULT""
        //            FROM ""@FIL_MH_OCSM""
        //            WHERE ""U_STYLENO"" = '{styleCode}'
        //            AND ""U_DRFTNTRY"" = '{draftCode}'; ";

        //        oRS2.DoQuery(sql);

        //        if (!oRS2.EoF)
        //            nextValue = oRS2.Fields.Item("RESULT").Value.ToString();

        //        // --- Generate the Style Code ---
        //        string verCode = $"{styleCode}-{nextValue}";

        //        // --- Set Value to ETSTYLID ---
        //        if (oForm.Items.Item("ETSTYLCD") != null)
        //        {
        //            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("ETVERSON").Specific;
        //            oEdit.Value = verCode;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Application.SBO_Application.StatusBar.SetText(
        //            "Error generating Route Code: " + ex.Message,
        //            SAPbouiCOM.BoMessageTime.bmt_Short,
        //            SAPbouiCOM.BoStatusBarMessageType.smt_Error
        //        );
        //    }
        //}
        private void FormMode(SAPbouiCOM.Form pForm)
        {
            try
            {
                pForm.Items.Item("ETQCQTY").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETDOCNUM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                pForm.Items.Item("ETBSEDOCNO").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                pForm.Items.Item("ETITEMCODE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                pForm.Items.Item("ETDOCDATE").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETACPTQTY").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETSCRPQTY").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETREMARKS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETACPTITM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);


                pForm.Items.Item("ETGRDAITM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETGRDBITM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETGRDCITM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                pForm.Items.Item("ETACPTWHS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                pForm.Items.Item("ETGRDAWHS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETGRDBWHS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETGRDCWHS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                pForm.Items.Item("ETSCRPITM").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("ETSCRPWHS").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("CBSERIES").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                pForm.Items.Item("MTX01").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }
            catch (Exception exception1)
            {

                Application.SBO_Application.StatusBar.SetText("DeleteRow  Method Failed:" + exception1.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            }
        }

        public bool IsFormOpen(string formUID)
        {
            try
            {
                foreach (SAPbouiCOM.Form form in Application.SBO_Application.Forms)
                {
                    if (form.UniqueID == formUID)
                    {
                        return true; // Form is already open (SAPbouiCOM.Form)Application.SBO_Application.Forms
                    }
                }
            }
            catch (Exception ex)
            {
                Global.G_UI_Application.StatusBar.SetText("Error checking form: " + ex.Message,
                   SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            return false; // Form is not open
        }

        public void SetNextCode( SAPbouiCOM.Form oForm,string editTextItemId,string tableName)
        {
            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(editTextItemId).Specific;
            long code;
            if (tableName == "FIL_MH_QINSPMAS")
            {
                code = 506000001;
            }
            else
            {
                code = 507000001;
            }

            SAPbobsCOM.Recordset rs =(SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string sql = $@"SELECT MAX(""Code"") AS ""Code"" FROM ""@{tableName}""";

            rs.DoQuery(sql);

            if (!rs.EoF)
            {
                object valueObj = rs.Fields.Item("Code").Value;

                if (valueObj != null && valueObj != DBNull.Value)
                {
                    string value = valueObj.ToString().Trim();

                    if (long.TryParse(value, out long maxCode))
                    {
                        code = maxCode + 1;
                    }
                }
            }

            oEdit.Value = code.ToString();
        }
        //public void SetNextCode( SAPbouiCOM.Form oForm, string editTextItemId, string tableName,long startCode = 10000001)
        //{
        //    SAPbouiCOM.EditText oEdit =
        //        (SAPbouiCOM.EditText)oForm.Items.Item(editTextItemId).Specific;

        //    long code = startCode;

        //    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //    string sql = $@"SELECT MAX(""Code"") AS ""Code"" FROM ""@{tableName}""";

        //    rs.DoQuery(sql);

        //    if (!rs.EoF && rs.Fields.Item("Code").Value.ToString() != "")
        //    {
        //        code = Convert.ToInt64(rs.Fields.Item("Code").Value) + 1;
        //    }

        //    oEdit.Value = code.ToString();
        //}

        private void CompanyConnection()
        {

            try
            {
                string sErrorMsg;
                string cookie;
                string connStr;
                // Global.ocomp.
                Global.oComp = new SAPbobsCOM.Company();
                cookie = Global.oComp.GetContextCookie();
                //    Global.oCompany = new SAPbobsCOM.Company();
                //   cookie =Global.oCompany.GetContextCookie();
                connStr = Application.SBO_Application.Company.GetConnectionContext(cookie);
                Global.oComp.SetSboLoginContext(connStr);
                ////   if (Global.CF.IsSAPHANA())
                ////  {
                ////   Global.oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
                //// }
                //// else
                //// {
                //Global.ocomp.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2019;
                // }
                // Global.oCompany.Connect();
                Global.G_UI_Application = Application.SBO_Application;
                Global.oComp = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany(); // Reassign the ocomp with the session we conencted with sap b1
                                                                                                       // sErrorMsg = Global.oCompany.GetLastErrorDescription();
                Application.SBO_Application.StatusBar.SetText("Quality Control Addon Connected Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch
            {
                Application.SBO_Application.MessageBox(Global.oComp.GetLastErrorDescription().ToString(), 1, "OK", "", "");
            }
        }
        public void CreateMainMenu(string ParentMenuID, string MenuID, string MenuName, int Position, int imenutype, bool flgimg) // POP UP- PARENT
        {
            try
            {
                SAPbouiCOM.Menus oMenus = null; // Define a variable to "menus"
                SAPbouiCOM.MenuItem oMenuItem = null; // Define a variable to MenuItem

                oMenus = Application.SBO_Application.Menus;  // Assign a SAP menu

                SAPbouiCOM.MenuCreationParams oCreationPackage = null;   //Define a variable to menu creating parameter
                oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
                oMenuItem = Application.SBO_Application.Menus.Item(ParentMenuID); // "43520" moudles'  //assign a Parent menu




                switch (imenutype)
                {
                    case 2:
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                            break;
                        }
                    case 1:
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            break;
                        }
                    case 3:
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_SEPERATOR;
                            break;
                        }
                }

                oCreationPackage.UniqueID = MenuID;
                oCreationPackage.String = MenuName;
                oCreationPackage.Enabled = true;
                oCreationPackage.Position = Position;  //postion is integer and it start from 0 value

                //string path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath).ToString();
                string path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath).ToString();
                //string Img = string.Concat(path, @"\BANKREC1.png");
                //oCreationPackage.Image = Img;
                if (flgimg == true)
                {
                    if (MenuID == "FIL_MN_QC")
                    {
                        string Bank = string.Concat(path, @"\OJ91D303.JPG");
                        oCreationPackage.Image = Bank;
                    }

                }
                oMenus = oMenuItem.SubMenus;

                try
                {
                    //  If the menu already exists this code will fail
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception ex)
                {

                }
            }
            catch
            {

            }
        }


    }
}
