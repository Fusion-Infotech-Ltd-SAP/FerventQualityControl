using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QualityControl
{
    class GlobalFunction
    {
        public bool setComboBoxValue(SAPbouiCOM.ComboBox oComboBox, string strQry)
        {
            bool flag;
            try
            {

                int count = oComboBox.ValidValues.Count;//0
                if (count > 0)
                {
                    while (true)
                    {
                        if (count <= 0)
                        {
                            break;
                        }
                        oComboBox.ValidValues.Remove(count - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                        count--;
                    }
                }
                //IN VS-CORE DOTNET WE USE DATASET- AS LIKE SAME FUNCTIONALITY, IT WILL USING RECORDSET IN SAP
                SAPbobsCOM.Recordset businessObject = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string str = Convert.ToString(oComboBox.ValidValues.Count);
                if (oComboBox.ValidValues.Count == 0)
                {

                    businessObject.DoQuery(strQry); //doquery means- executinga query
                    businessObject.MoveFirst();
                    int num2 = businessObject.RecordCount - 1; //linelevel count 
                    int num = 0;
                    while (true)
                    {
                        int num3 = num2;
                        if (num > num3)
                        {
                            break;
                        }
                        oComboBox.ValidValues.Add(Convert.ToString(businessObject.Fields.Item(0).Value), Convert.ToString(businessObject.Fields.Item(1).Value));

                        businessObject.MoveNext(); // it will move to next cursor to recordset
                        num++;
                    }
                }
                oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_ValueDescription;
                oComboBox.Item.DisplayDesc = true;
                flag = true;
            }
            catch (Exception exception1)
            {

                Application.SBO_Application.StatusBar.SetText("setComboBoxValue Function Failed:" + exception1.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                flag = true;

            }
            return flag;
        }
        SAPbouiCOM.EditText omatcol;
        SAPbouiCOM.ComboBox omatcolb;
        public void SetNewLine(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSDetail, int RowID = 1, string ColumnUID = "")
        {
            try
            {

                if (ColumnUID != "")
                {
                    omatcolb = (SAPbouiCOM.ComboBox)oMatrix.Columns.Item(ColumnUID).Cells.Item(RowID).Specific;
                }

                if (ColumnUID.Equals(""))  //no column assign ; eventhough no values exist in previous column then also can add new lines.
                {
                    oMatrix.FlushToDataSource();
                    oMatrix.AddRow(1, -1);
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1;
                    oDBDSDetail.SetValue("LineId", oDBDSDetail.Offset, Convert.ToString(oMatrix.VisualRowCount));
                    oMatrix.SetLineData(oMatrix.VisualRowCount);
                    oMatrix.FlushToDataSource();
                }
                else if (oMatrix.VisualRowCount <= 0)  //1st time row creation
                {
                    oMatrix.FlushToDataSource();
                    oMatrix.AddRow(1, -1);//1-1=4
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1; //starting index from 0
                    oDBDSDetail.SetValue("LineId", oDBDSDetail.Offset, Convert.ToString(oMatrix.VisualRowCount));
                    oMatrix.SetLineData(oMatrix.VisualRowCount);
                    oMatrix.FlushToDataSource();
                }
                else if (!(omatcolb.Value).Equals("") && (RowID == oMatrix.VisualRowCount))  // column assigned ; only add a row when present column value is not null.
                {
                    oMatrix.FlushToDataSource();
                    oMatrix.AddRow(1, -1);
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1;
                    oDBDSDetail.SetValue("LineId", oDBDSDetail.Offset, Convert.ToString(oMatrix.VisualRowCount));
                    oMatrix.SetLineData(oMatrix.VisualRowCount);
                    //oMatrix.LoadFromDataSource();
                    oMatrix.FlushToDataSource();
                }
            }
            catch (Exception exception1)
            {

            }
        }

        public void SetNewLineForEditText(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSDetail, int RowID = 1, string ColumnUID = "")
        {
            try
            {

                if (ColumnUID != "")
                {
                    omatcol = (SAPbouiCOM.EditText)oMatrix.Columns.Item(ColumnUID).Cells.Item(RowID).Specific;
                }

                if (ColumnUID.Equals(""))  //no column assign ; eventhough no values exist in previous column then also can add new lines.
                {
                    oMatrix.FlushToDataSource();
                    oMatrix.AddRow(1, -1);
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1;
                    oDBDSDetail.SetValue("LineId", oDBDSDetail.Offset, Convert.ToString(oMatrix.VisualRowCount));
                    oMatrix.SetLineData(oMatrix.VisualRowCount);
                    oMatrix.FlushToDataSource();
                }
                else if (oMatrix.VisualRowCount <= 0)  //1st time row creation
                {
                    oMatrix.FlushToDataSource();
                    oMatrix.AddRow(1, -1);//1-1=4
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1; //starting index from 0
                    oDBDSDetail.SetValue("LineId", oDBDSDetail.Offset, Convert.ToString(oMatrix.VisualRowCount));
                    oMatrix.SetLineData(oMatrix.VisualRowCount);
                    oMatrix.FlushToDataSource();
                }
                else if (!(omatcol.Value).Equals("") && (RowID == oMatrix.VisualRowCount))  // column assigned ; only add a row when present column value is not null.
                {
                    oMatrix.FlushToDataSource();
                    oMatrix.AddRow(1, -1);
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size);
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1;
                    oDBDSDetail.SetValue("LineId", oDBDSDetail.Offset, Convert.ToString(oMatrix.VisualRowCount));
                    oMatrix.SetLineData(oMatrix.VisualRowCount);
                    //oMatrix.LoadFromDataSource();
                    oMatrix.FlushToDataSource();
                }
            }
            catch (Exception exception1)
            {

            }
        }
        public bool LoadComboBoxSeries(SAPbouiCOM.ComboBox oComboBox, string UDOID)
        {
            bool flag;
            try
            {
                oComboBox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                oComboBox.ValidValues.LoadSeries(UDOID, SAPbouiCOM.BoSeriesMode.sf_Add);  // ONLY TO LOAD A COMBOBOX
                oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                oComboBox.Item.DisplayDesc = true;
                flag = true;
            }
            catch (Exception exception1)
            {
                Application.SBO_Application.SetStatusBarMessage("error");
                flag = false;

            }
            return flag;
        }

        public void DeleteRow(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource oDBDSDetail)
        {
            try
            {
                
                oMatrix.FlushToDataSource();
                int visualRowCount = oMatrix.VisualRowCount+1; //5
                int rowNum = 1;
                while (true)
                {
                    int num3 = visualRowCount; //5
                    if (rowNum >= num3) //5=>5
                    {
                        oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1);
                        oMatrix.LoadFromDataSource();
                        break;
                    }

                    oMatrix.GetLineData(rowNum);
                    oDBDSDetail.Offset = rowNum - 1;
                    oDBDSDetail.SetValue("LineId", oDBDSDetail.Offset, Convert.ToString(rowNum));
                    oMatrix.SetLineData(rowNum);
                    oMatrix.FlushToDataSource();
                    rowNum++;
                }
                //1Item1
                //2Item2
                //3Item3
                //4Iten5
            }
            catch (Exception exception1)
            {

                Application.SBO_Application.StatusBar.SetText("DeleteRow  Method Failed:" + exception1.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            }
        }

        public int GetCodeGeneration(string TableName)
        {
            int num;
            try
            {
                SAPbobsCOM.Recordset businessObject = (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sqlQuery = string.Format("SELECT ifnull(Max({0}DocEntry{0}),0) + 1 as  {0}Code{0} from {0}" + TableName.Trim().ToString() + "{0}", '"');
                businessObject.DoQuery(sqlQuery);
                num = Convert.ToInt32(businessObject.Fields.Item("Code").Value.ToString());

            }
            catch (Exception exception1)
            {

                Application.SBO_Application.StatusBar.SetText("GetCodeGeneration Function Failed:" + exception1.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                num = -1;

            }
            return num;
        }
        public string GetCodeGeneration1000(string tableName)
        {
            string nextCode = "1000";

            try
            {
                SAPbobsCOM.Recordset rs =
                    (SAPbobsCOM.Recordset)Global.oComp.GetBusinessObject(
                        SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = $@"
            SELECT 
                TO_NVARCHAR(IFNULL(MAX(CAST(""Code"" AS INTEGER)), 999) + 1) 
                AS ""NextCode""
            FROM ""{tableName}""
        ";

                rs.DoQuery(query);

                nextCode = rs.Fields.Item("NextCode").Value.ToString();
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            return nextCode;
        }
        public void ShowError(string ErrorMessage)
        {
            Application.SBO_Application.StatusBar.SetText(ErrorMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        }
        public void ShowSuccess(string ErrorMessage)
        {
            Application.SBO_Application.StatusBar.SetText(ErrorMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        public void ShowWarning(string ErrorMessage)
        {
            Application.SBO_Application.StatusBar.SetText(ErrorMessage, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }
    }
}
