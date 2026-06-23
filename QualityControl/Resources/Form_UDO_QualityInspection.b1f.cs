using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QualityControl.Resources
{
    [FormAttribute("QualityControl.Resources.Form_UDO_QualityInspection", "Resources/Form_UDO_QualityInspection.b1f")]
    class Form_UDO_QualityInspection : UserFormBase
    {
       
        public Form_UDO_QualityInspection()
        {
        }

        private SAPbouiCOM.Matrix MTX01, MTX02, MTX03;
        private SAPbouiCOM.Folder TAB1, TAB2, TAB3;
        private SAPbouiCOM.Button Add, Cancel;
        private SAPbouiCOM.StaticText STITMGRPCD, STCHRGCODE, STCHRGNAME, STSUBGRP, STHICKNSS;
        private SAPbouiCOM.EditText ETCHRGCODE, ETCHRGNAME, ETDOCENTRY;
        private SAPbouiCOM.ComboBox CBITMGRPCD, CBSUBGRP, CBHICKNSS;
        private SAPbouiCOM.CheckBox CBACTIVE;



        public override void OnInitializeComponent()
        {
            this.TAB1 = ((SAPbouiCOM.Folder)(this.GetItem("TAB1").Specific));
            this.TAB2 = ((SAPbouiCOM.Folder)(this.GetItem("TAB2").Specific));
            this.TAB2.ClickAfter += new SAPbouiCOM._IFolderEvents_ClickAfterEventHandler(this.TAB2_ClickAfter);
            this.TAB3 = ((SAPbouiCOM.Folder)(this.GetItem("TAB3").Specific));
            this.MTX01 = ((SAPbouiCOM.Matrix)(this.GetItem("MTX01").Specific));
            this.MTX02 = ((SAPbouiCOM.Matrix)(this.GetItem("MTX02").Specific));
            this.MTX03 = ((SAPbouiCOM.Matrix)(this.GetItem("MTX03").Specific));
            this.MTX03.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.MTX03_ChooseFromListBefore);
            this.MTX03.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.MTX03_ChooseFromListAfter);
            this.Add = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Add.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Add_PressedAfter);
            this.Add.PressedBefore += new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Add_PressedBefore);
            this.Cancel = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.STITMGRPCD = ((SAPbouiCOM.StaticText)(this.GetItem("STITMGRPCD").Specific));
            this.STCHRGCODE = ((SAPbouiCOM.StaticText)(this.GetItem("STCHRGCODE").Specific));
            this.STCHRGNAME = ((SAPbouiCOM.StaticText)(this.GetItem("STCHRGNAME").Specific));
            this.ETCHRGCODE = ((SAPbouiCOM.EditText)(this.GetItem("ETCHRGCODE").Specific));
            this.ETCHRGCODE.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.ETCHRGCODE_ChooseFromListAfter);
            this.ETCHRGNAME = ((SAPbouiCOM.EditText)(this.GetItem("ETCHRGNAME").Specific));
            this.ETDOCENTRY = ((SAPbouiCOM.EditText)(this.GetItem("ETDOCENTRY").Specific));
            this.CBITMGRPCD = ((SAPbouiCOM.ComboBox)(this.GetItem("CBITMGRPCD").Specific));
            this.CBITMGRPCD.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.CBITMGRPCD_ComboSelectAfter);
            this.CBACTIVE = ((SAPbouiCOM.CheckBox)(this.GetItem("CBACTIVE").Specific));
            this.CBSUBGRP = ((SAPbouiCOM.ComboBox)(this.GetItem("CBSUBGRP").Specific));
            this.CBHICKNSS = ((SAPbouiCOM.ComboBox)(this.GetItem("CBHICKNSS").Specific));
            this.STSUBGRP = ((SAPbouiCOM.StaticText)(this.GetItem("STSUBGRP").Specific));
            this.STHICKNSS = ((SAPbouiCOM.StaticText)(this.GetItem("STHICKNSS").Specific));
            this.OnCustomInitialize();

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_MH_INSPLAN");
            if(oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
            {
                AddRow_Item(oForm);
            }

            
        }

        private void CBITMGRPCD_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_MH_INSPLAN");
                SAPbouiCOM.Matrix MTX03 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX03").Specific;
                SAPbouiCOM.DBDataSource DBDataSourceLine =
                    oform.DataSources.DBDataSources.Item("@FIL_MR_INPLNIM");

                if (MTX03.RowCount == 0)
                {
                    Global.GFunc.SetNewLineForEditText(MTX03, DBDataSourceLine, 1, "");
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        public override void OnInitializeFormEvents()
        {
            this.RightClickBefore += new SAPbouiCOM.Framework.FormBase.RightClickBeforeHandler(this.Form_RightClickBefore);
            this.DataAddBefore += new SAPbouiCOM.Framework.FormBase.DataAddBeforeHandler(this.Form_DataAddBefore);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private void MTX03_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_MH_INSPLAN");
                SAPbouiCOM.ChooseFromList ocfl = (SAPbouiCOM.ChooseFromList)oForm.ChooseFromLists.Item("CFL_OITM");

                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                string itmGrp = oForm.DataSources.DBDataSources.Item("@FIL_MH_INPLAN").GetValue("U_ITMGRPCD", 0).Trim();

                string subGrp = oForm.DataSources.DBDataSources.Item("@FIL_MH_INPLAN").GetValue("U_SUBGRP", 0).Trim();

                string thickness = oForm.DataSources.DBDataSources.Item("@FIL_MH_INPLAN").GetValue("U_THICKNSS", 0).Trim();


                // 1st condition (always)
                SAPbouiCOM.Condition oCon = oCons.Add();
                oCon.Alias = "ItmsGrpCod";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = itmGrp;

                // 2nd condition: Inventory Item = Y
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.Alias = "InvntItem";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "Y";

                // 2nd condition (if exists)
                if (subGrp != "")
                {
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    oCon = oCons.Add();
                    oCon.Alias = "U_SUBGRP";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = subGrp;
                }

                // 3rd condition (if exists)
                if (thickness != "")
                {
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    oCon = oCons.Add();
                    oCon.Alias = "U_THICKNESS";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = thickness;
                }

                ocfl.SetConditions(oCons);


            }
            catch (Exception ex)
            {

            }

        }

        private void Form_RightClickBefore(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form ofrom = (SAPbouiCOM.Form)Application.SBO_Application.Forms.Item(eventInfo.FormUID);

            if (eventInfo.ItemUID == "MTX01")
            {
                ofrom.EnableMenu("1293", true);
            }
            else
            {
                ofrom.EnableMenu("1293", false);
            }


        }

        private void OnCustomInitialize()
        {

        }
        private bool ValidateForm(ref SAPbouiCOM.Form pForm, ref bool BubbleEvent)
        {
            SAPbobsCOM.Recordset itmGrpCheck = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string DocEntry =  pForm.DataSources.DBDataSources.Item("@FIL_MH_INPLAN").GetValue("DocEntry", 0);
           string ITMGRPCD = pForm.DataSources.DBDataSources.Item("@FIL_MH_INPLAN").GetValue("U_ITMGRPCD", 0);
           string SUBGRP = pForm.DataSources.DBDataSources.Item("@FIL_MH_INPLAN").GetValue("U_SUBGRP", 0);
           string THICKNSS = pForm.DataSources.DBDataSources.Item("@FIL_MH_INPLAN").GetValue("U_THICKNSS", 0);

          
            if (ITMGRPCD == "")
            {
                Global.GFunc.ShowError("Please Select Item Group");
                pForm.ActiveItem = "CBITMGRPCD";
                return BubbleEvent = false;
            }

            SAPbobsCOM.Recordset itemg = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
           string itemgs = string.Format("SELECT A.{0}U_ISFG{0}  from {0}OITB{0} A Where  A.{0}ItmsGrpCod{0}= '" + ITMGRPCD + "'", '"');
            itemg.DoQuery(itemgs);

            if (itemg.RecordCount > 0)
            {
              string ISFG = itemg.Fields.Item("U_ISFG").Value.ToString();
                if (SUBGRP == "" && ISFG == "Y")
                {
                    Global.GFunc.ShowError("Please Select Sub Group");
                    pForm.ActiveItem = "CBSUBGRP";
                    return BubbleEvent = false;
                }
                else if (THICKNSS == "" && ISFG == "Y")
                {
                    Global.GFunc.ShowError("Please Select Thickness");
                    pForm.ActiveItem = "CBHICKNSS";
                    return BubbleEvent = false;
                }
            }

            //Item Group Duplicate Check
           if(pForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                string qstr = $@"
                        SELECT COUNT(*) AS ""Cnt""
                        FROM ""@FIL_MH_INPLAN""
                        WHERE ""U_ITMGRPCD"" = '{ITMGRPCD}'
                        AND ""U_SUBGRP"" = '{SUBGRP}'
                        AND ""U_THICKNSS"" = '{THICKNSS}'
                        AND ""U_ACTIVE"" = 'Y'
                        ";

                itmGrpCheck.DoQuery(qstr);
                int dataCount = Convert.ToInt32(itmGrpCheck.Fields.Item("Cnt").Value);
                if (dataCount > 0)
                {
                    Global.GFunc.ShowError("Quality Inspection already created with this Item Group,Sub Group with THickness ");
                    return BubbleEvent = false;
                }

                //Item Duplicate
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)pForm.Items.Item("MTX03").Specific;

                for (int i = 1; i <= oMatrix.VisualRowCount; i++)
                {
                    string item = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLITEMCODE")
                                    .Cells.Item(i).Specific).Value;

                    SAPbobsCOM.Recordset oRs =
                    (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(
                    SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string sql = $@"
                                SELECT COUNT(*) AS ""Count""
                                FROM ""@FIL_MR_INPLNIM"" T0
                                INNER JOIN ""@FIL_MH_INPLAN"" T1 
                                    ON T0.""Code"" = T1.""Code""
                                WHERE T0.""U_ITEMCODE"" = '{item}'
                                  AND T1.""U_ACTIVE"" = 'Y'";

                    oRs.DoQuery(sql);

                    int count = Convert.ToInt32(oRs.Fields.Item("Count").Value);

                    if (count > 0)
                    {
                        Global.GFunc.ShowError($"Item {item} already exists. Duplicate not allowed!");
                        return BubbleEvent = false;

                    }
                }
            }
           

           
            return BubbleEvent;
        }

        private static void ClearBlankRow(ref SAPbouiCOM.Form pForm)
        {

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)pForm.Items.Item("MTX03").Specific;
            for (int i = oMatrix.RowCount; i >= 1; i--)
            {
                string cCombo = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLITEMCODE").Cells.Item(i).Specific).Value;
                if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("CLITEMCODE").Cells.Item(i).Specific).Value == "")
                    oMatrix.DeleteRow(i);
                else
                    break;
            }

        }
        private void ETCHRGCODE_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_MH_INSPLAN"); // DEFINE A FORM
                SAPbouiCOM.DBDataSource DBDataSourceHeader = (SAPbouiCOM.DBDataSource)oform.DataSources.DBDataSources.Item("@FIL_MH_INPLAN");
                SAPbouiCOM.DBDataSource DBDataSourceLine = (SAPbouiCOM.DBDataSource)oform.DataSources.DBDataSources.Item("@FIL_MR_INPLNPM");

                if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oform.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal; //Assign a cfl and call event
                    string Uid = cflEvent.ChooseFromListUID;  // cfl unique id
                    string strCode = "";

                    SAPbouiCOM.DataTable dtbCFL = cflEvent.SelectedObjects;
                    if (!dtbCFL.IsEmpty)
                    {
                        strCode = dtbCFL.GetValue("Code", 0).ToString();

                        ((SAPbouiCOM.EditText)oform.Items.Item("ETCHRGCODE").Specific).Value = dtbCFL.GetValue("Code", 0).ToString();
                        ((SAPbouiCOM.EditText)oform.Items.Item("ETCHRGNAME").Specific).Value = dtbCFL.GetValue("Name", 0).ToString();

                        SAPbouiCOM.Matrix MTX01 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX01").Specific;

                        string strqry = string.Format("SELECT A.{0}U_CHARCODE{0},A.{0}U_CHARNAME{0}  from {0}@FIL_MR_CHARMGRP{0} A Where A.{0}Code{0}= '" + strCode + "' ", '"');


                        SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        ors.DoQuery(strqry);

                        MTX01.FlushToDataSource(); // move the data from matrix to dbdatasource
                        MTX01.Clear();// clear or remove from matrix
                        DBDataSourceLine.Clear(); // clear or remove from dbdatasource(in backend table)
                        for (int irow = 1; irow <= ors.RecordCount; irow++) //ToString load a data from record set to dbdatasource
                        {
                            DBDataSourceLine.InsertRecord(irow - 1);
                            DBDataSourceLine.SetValue("LineId", irow - 1, irow.ToString());
                            DBDataSourceLine.SetValue("U_CHARCODE", irow - 1, ors.Fields.Item("U_CHARCODE").Value.ToString());
                            DBDataSourceLine.SetValue("U_CHARNAME", irow - 1, ors.Fields.Item("U_CHARNAME").Value.ToString());
                            //to move recordset
                            ors.MoveNext();
                        }
                        MTX01.LoadFromDataSource();
                        MTX01.AutoResizeColumns();

                    }
                }
            }
            catch (Exception ex)
            {

            }


        }



        private void Add_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
            //if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            //{
            //    SAPbouiCOM.DBDataSource ods = (SAPbouiCOM.DBDataSource)oform.DataSources.DBDataSources.Item("@FIL_MH_INPLAN");
            //    ods.SetValue("Code", 0, Global.GFunc.GetCodeGeneration("@FIL_MH_INPLAN").ToString());
            //}
            if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oform.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
            {
                ClearBlankRow(ref oform);
                ValidateForm(ref oform, ref BubbleEvent);
                AddRow_Item(oform);
            }

        }

        private void Form_DataAddBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            // throw new System.NotImplementedException();
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item(pVal.FormUID);
            int Code = Global.GFunc.GetCodeGeneration("@FIL_MH_INPLAN");
            oform.DataSources.DBDataSources.Item("@FIL_MH_INPLAN").SetValue("Code", 0, Code.ToString());
        }
        private void Add_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item("FIL_FRM_MH_INSPLAN");
            oForm.Freeze(true);

            oForm.Freeze(false);

        }

        private void MTX03_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_MH_INSPLAN");

            SAPbouiCOM.Matrix MTX03 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX03").Specific;

            SAPbouiCOM.DBDataSource DBDataSourceLine = oform.DataSources.DBDataSources.Item("@FIL_MR_INPLNIM");

            SAPbouiCOM.ISBOChooseFromListEventArg cflEvent = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;

            SAPbouiCOM.DataTable dtbCFL = cflEvent.SelectedObjects;

            if (dtbCFL == null || pVal.ColUID != "CLITEMCODE")
                return;

            try
            {
                oform.Freeze(true);

                MTX03.FlushToDataSource();

                int dsRow = pVal.Row - 1;

                for (int r = 0; r < dtbCFL.Rows.Count; r++)
                {
                    string itemCode = dtbCFL.GetValue("ItemCode", r).ToString();

                    string sqlQuery =
                        "SELECT \"ItemCode\", \"ItemName\", \"InvntryUom\" " +
                        "FROM \"OITM\" " +
                        "WHERE \"ItemCode\" = '" + itemCode.Replace("'", "''") + "'";

                    SAPbobsCOM.Recordset ors =
                        (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(
                            SAPbobsCOM.BoObjectTypes.BoRecordset);

                    ors.DoQuery(sqlQuery);

                    if (!ors.EoF)
                    {
                        if (dsRow >= DBDataSourceLine.Size)
                        {
                            DBDataSourceLine.InsertRecord(DBDataSourceLine.Size);
                        }

                        DBDataSourceLine.SetValue(
                            "U_ITEMCODE",
                            dsRow,
                            ors.Fields.Item("ItemCode").Value.ToString());

                        DBDataSourceLine.SetValue(
                            "U_ITEMNAME",
                            dsRow,
                            ors.Fields.Item("ItemName").Value.ToString());

                       
                        dsRow++;
                    }
                }

                for (int i = DBDataSourceLine.Size - 1; i >= 0; i--)
                {
                    string code = DBDataSourceLine.GetValue("U_ITEMCODE", i).Trim();

                    if (string.IsNullOrEmpty(code))
                    {
                        DBDataSourceLine.RemoveRecord(i);
                    }
                }

                for (int i = 0; i < DBDataSourceLine.Size; i++)
                {
                    DBDataSourceLine.SetValue("LineId", i, (i + 1).ToString());
                }

                DBDataSourceLine.InsertRecord(DBDataSourceLine.Size);

                DBDataSourceLine.SetValue( "LineId", DBDataSourceLine.Size - 1,DBDataSourceLine.Size.ToString());

                MTX03.Clear();
                MTX03.LoadFromDataSource();

                if (oform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    oform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }

                MTX03.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.Message);
            }
            finally
            {
                oform.Freeze(false);
            }
        }


        private static void AddRow_Item(SAPbouiCOM.Form pForm)
        {
            //oForm.Freeze(true);
            pForm.Freeze(true);
            SAPbouiCOM.DBDataSource oDataSource = pForm.DataSources.DBDataSources.Item("@FIL_MR_INPLNIM");
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)pForm.Items.Item("MTX03").Specific;
            oMatrix.FlushToDataSource();
            oDataSource.InsertRecord(oDataSource.Size);
            if (oDataSource.Size > 1)
                for (int i = oDataSource.Size - 2; i >= 0; i--)
                {
                    if (oDataSource.GetValue("U_ITEMCODE", i) == "")
                        oDataSource.RemoveRecord(i);
                    else
                        break;
                }
            for (int i = 0; i < oDataSource.Size; i++)
            {
                oDataSource.SetValue("LineId", i, (i + 1).ToString());
            }
            oMatrix.Clear();
            oMatrix.LoadFromDataSource();
            //oForm.Freeze(false);
            pForm.Freeze(false);
        }


        private void TAB2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            SAPbouiCOM.Form oform = Application.SBO_Application.Forms.Item("FIL_FRM_MH_INSPLAN");
            SAPbouiCOM.DBDataSource DBDataSourceLine =  oform.DataSources.DBDataSources.Item("@FIL_MR_INPLNDC");

            SAPbouiCOM.Matrix MTX02 = (SAPbouiCOM.Matrix)oform.Items.Item("MTX02").Specific;

            try
            {
                oform.Freeze(true);

                if (oform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    if (MTX02.RowCount > 0)
                    {
                        oform.Freeze(false);
                        return;
                    }
                    string strqry = string.Format("SELECT A.{0}DocumentType{0}, A.{0}Description{0}, A.{0}ObjType{0}, A.{0}BaseType{0} FROM {0}VW_QCDocument{0} A",'"');
                    SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    ors.DoQuery(strqry);

                    DBDataSourceLine.Clear();

                    int row = 0;

                    while (!ors.EoF)
                    {
                        DBDataSourceLine.InsertRecord(row);

                        DBDataSourceLine.SetValue("LineId", row, (row + 1).ToString());
                        DBDataSourceLine.SetValue("U_DSCRPTN", row, ors.Fields.Item("Description").Value.ToString());
                        DBDataSourceLine.SetValue("U_DOCTYPE", row, ors.Fields.Item("DocumentType").Value.ToString());
                        DBDataSourceLine.SetValue("U_OBJTYPE", row, ors.Fields.Item("ObjType").Value.ToString());
                        DBDataSourceLine.SetValue("U_BASETYPE", row, ors.Fields.Item("BaseType").Value.ToString());
                       // DBDataSourceLine.SetValue("U_ACTIVE", row, "Y");

                        row++;
                        ors.MoveNext();
                    }

                    MTX02.LoadFromDataSource();
                    MTX02.AutoResizeColumns();
                }
                else
                {
                    //SAPbouiCOM.DBDataSource oDBHeader = oform.DataSources.DBDataSources.Item("@FIL_MH_INPLAN");

                    //string DocEntry = oDBHeader.GetValue("DocEntry", 0).Trim();

                    //DBDataSourceLine.Clear();

                    //string strqry = @"
                    //        SELECT 
                    //            ""LineId"",
                    //            ""U_DSCRPTN"",
                    //            ""U_DOCTYPE"",
                    //            ""U_OBJTYPE"",
                    //            ""U_BASETYPE"",
                    //            ""U_ACTIVE""
                    //        FROM ""@FIL_MR_INPLNDC""
                    //        WHERE ""Code"" = '" + DocEntry + @"'
                    //        ORDER BY ""LineId""";

                    //SAPbobsCOM.Recordset ors =
                    //    (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    //ors.DoQuery(strqry);

                    //int row = 0;

                    //while (!ors.EoF)
                    //{
                    //    DBDataSourceLine.InsertRecord(row);

                    //    DBDataSourceLine.SetValue("LineId", row, ors.Fields.Item("LineId").Value.ToString());
                    //    DBDataSourceLine.SetValue("U_DSCRPTN", row, ors.Fields.Item("U_DSCRPTN").Value.ToString());
                    //    DBDataSourceLine.SetValue("U_DOCTYPE", row, ors.Fields.Item("U_DOCTYPE").Value.ToString());
                    //    DBDataSourceLine.SetValue("U_OBJTYPE", row, ors.Fields.Item("U_OBJTYPE").Value.ToString());
                    //    DBDataSourceLine.SetValue("U_BASETYPE", row, ors.Fields.Item("U_BASETYPE").Value.ToString());
                    //    DBDataSourceLine.SetValue("U_ACTIVE", row, ors.Fields.Item("U_ACTIVE").Value.ToString());

                    //    row++;
                    //    ors.MoveNext();
                    //}

                    //MTX02.LoadFromDataSource();
                    //MTX02.AutoResizeColumns();
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                oform.Freeze(false);
            }
        }

    }
}
