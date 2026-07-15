using System;
using System.Runtime.InteropServices;

namespace QualityControl.Helper
{
    internal static class PermissionInstaller
    {
        public static void Install()
        {
            EnsurePermission(PermissionConstants.Root, "Quality Control", "", SAPbobsCOM.BoUPTOptions.bou_FullReadNone, null);
            EnsurePermission(PermissionConstants.InspectionParameter, "Inspection Parameter Master", PermissionConstants.Root, SAPbobsCOM.BoUPTOptions.bou_FullReadNone, "FIL_FRM_MH_QINSPMAS");
            EnsurePermission(PermissionConstants.CharacteristicGroup, "Characteristic Master Group", PermissionConstants.Root, SAPbobsCOM.BoUPTOptions.bou_FullReadNone, "FIL_FRM_MH_CHARMGRPS");
            EnsurePermission(PermissionConstants.InspectionPlan, "Inspection Plan", PermissionConstants.Root, SAPbobsCOM.BoUPTOptions.bou_FullReadNone, "FIL_FRM_MH_INSPLAN");
            EnsurePermission(PermissionConstants.CheckSelection, "Inspection Check Selection", PermissionConstants.Root, SAPbobsCOM.BoUPTOptions.bou_FullReadNone, "FIL_FRM_MH_CHKSLC,FIL_FRM_NO_IPNDLIST");
            EnsurePermission(PermissionConstants.PendingList, "QC Checked Stock Not Transfer", PermissionConstants.Root, SAPbobsCOM.BoUPTOptions.bou_FullReadNone, "FIL_FRM_MH_PNDLIST");
            EnsurePermission(PermissionConstants.InspectionDecision, "Inspection Decision", PermissionConstants.Root, SAPbobsCOM.BoUPTOptions.bou_FullReadNone, "FIL_FRM_DH_INSPDECN");
        }

        private static void EnsurePermission(string permissionId, string name, string parentId, SAPbobsCOM.BoUPTOptions options, string formTypes)
        {
            SAPbobsCOM.UserPermissionTree permission = null;

            try
            {
                permission = (SAPbobsCOM.UserPermissionTree)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree);
                if (permission.GetByKey(permissionId))
                {
                    if (permission.Options != options)
                    {
                        permission.Options = options;
                        if (permission.Update() != 0)
                        {
                            int errCode;
                            string errMsg;
                            Global.oComp.GetLastError(out errCode, out errMsg);
                            Global.GFunc.ShowError("QC authorization upgrade failed [" + permissionId + "]: " + errMsg);
                        }
                    }

                    return;
                }

                permission.PermissionID = permissionId;
                permission.Name = name;
                permission.Options = options;

                if (!string.IsNullOrEmpty(parentId))
                    permission.ParentID = parentId;

                if (!string.IsNullOrEmpty(formTypes))
                {
                    string[] forms = formTypes.Split(',');
                    for (int i = 0; i < forms.Length; i++)
                    {
                        if (i > 0)
                            permission.UserPermissionForms.Add();

                        permission.UserPermissionForms.FormType = forms[i];
                    }
                }

                if (permission.Add() != 0)
                {
                    int errCode;
                    string errMsg;
                    Global.oComp.GetLastError(out errCode, out errMsg);
                    Global.GFunc.ShowError("QC authorization setup failed [" + permissionId + "]: " + errMsg);
                }
            }
            catch (Exception ex)
            {
                Global.GFunc.ShowError("QC authorization setup failed [" + permissionId + "]: " + ex.Message);
            }
            finally
            {
                if (permission != null)
                    Marshal.ReleaseComObject(permission);
            }
        }
    }
}
