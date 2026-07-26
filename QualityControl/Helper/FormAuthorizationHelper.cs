using System;

namespace QualityControl.Helper
{
    internal static class FormAuthorizationHelper
    {
        public static bool CanOpenMenu(string menuUid)
        {
            switch (menuUid)
            {
                case "FIL_QINSPMAS":
                    return AuthorizationService.EnsureView(PermissionConstants.InspectionParameter, "Inspection Parameter Master");
                case "FIL_CHARMGRPS":
                    return AuthorizationService.EnsureView(PermissionConstants.CharacteristicGroup, "Characteristic Master Group");
                case "FIL_INSPLAN":
                    return AuthorizationService.EnsureView(PermissionConstants.InspectionPlan, "Inspection Plan");
                case "FIL_CHKSLCTN":
                    return AuthorizationService.EnsureView(PermissionConstants.CheckSelection, "Inspection Check Selection");
                case "FIL_PNDLIST":
                    return AuthorizationService.EnsureView(PermissionConstants.PendingList, "QC Checked Stock Not Transfer");
                case "FIL_INSPDECN":
                    return AuthorizationService.EnsureView(PermissionConstants.InspectionDecision, "Inspection Decision");
            }

            return true;
        }

        public static bool CanSave(SAPbouiCOM.Form form)
        {
            if (form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                return AuthorizationService.EnsureFull(GetPermissionByForm(form.UniqueID), form.Title);

            if (form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                return AuthorizationService.EnsureFull(GetPermissionByForm(form.UniqueID), form.Title);

            return true;
        }

        public static void ApplyReadOnly(SAPbouiCOM.Form form, string permissionId)
        {
            if (AuthorizationService.GetAuthorization(permissionId) != AuthorizationLevel.ReadOnly)
                return;

            try
            {
                form.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                form.EnableMenu("1282", false);
                form.EnableMenu("1293", false);
                LockEditableItems(form);
            }
            catch
            {
            }
        }

        private static void LockEditableItems(SAPbouiCOM.Form form)
        {
            for (int i = 0; i < form.Items.Count; i++)
            {
                SAPbouiCOM.Item item = null;

                try
                {
                    item = form.Items.Item(i);

                    if (item.UniqueID == "1" || item.UniqueID == "2")
                        continue;

                    switch (item.Type)
                    {
                        case SAPbouiCOM.BoFormItemTypes.it_EDIT:
                        case SAPbouiCOM.BoFormItemTypes.it_EXTEDIT:
                        case SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX:
                        case SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX:
                            SetItemEditable(item, false);
                            break;

                        case SAPbouiCOM.BoFormItemTypes.it_BUTTON:
                            item.Enabled = false;
                            break;

                        case SAPbouiCOM.BoFormItemTypes.it_MATRIX:
                            LockMatrix((SAPbouiCOM.Matrix)item.Specific);
                            break;

                        case SAPbouiCOM.BoFormItemTypes.it_GRID:
                            LockGrid((SAPbouiCOM.Grid)item.Specific);
                            break;
                    }
                }
                catch
                {
                }
            }
        }

        private static void SetItemEditable(SAPbouiCOM.Item item, bool editable)
        {
            try
            {
                if (editable)
                {
                    item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_All, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                }
                else
                {
                    item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                    item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                    item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_View, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                    item.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, (int)SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                }
            }
            catch
            {
                try
                {
                    item.Enabled = editable;
                }
                catch
                {
                }
            }
        }

        private static void LockMatrix(SAPbouiCOM.Matrix matrix)
        {
            for (int i = 0; i < matrix.Columns.Count; i++)
            {
                try
                {
                    matrix.Columns.Item(i).Editable = false;
                }
                catch
                {
                }
            }
        }

        private static void LockGrid(SAPbouiCOM.Grid grid)
        {
            for (int i = 0; i < grid.Columns.Count; i++)
            {
                try
                {
                    grid.Columns.Item(i).Editable = false;
                }
                catch
                {
                }
            }
        }

        public static bool IsConfigurationForm(string formUid)
        {
            return formUid == "FIL_FRM_MH_QINSPMAS" ||
                   formUid == "FIL_FRM_MH_CHARMGRPS" ||
                   formUid == "FIL_FRM_MH_INSPLAN" ||
                   formUid == "FIL_FRM_MH_CHKSLC";
        }

        public static bool IsReportForm(string formUid)
        {
            return formUid == "FIL_FRM_MH_PNDLIST" ||
                   formUid == "FIL_FRM_NO_IPNDLIST";
        }

        public static bool IsInspectionForm(string formUid)
        {
            return formUid == "FIL_FRM_DH_INSPDECN";
        }

        public static string GetPermissionByForm(string formUid)
        {
            switch (formUid)
            {
                case "FIL_FRM_MH_QINSPMAS":
                    return PermissionConstants.InspectionParameter;
                case "FIL_FRM_MH_CHARMGRPS":
                    return PermissionConstants.CharacteristicGroup;
                case "FIL_FRM_MH_INSPLAN":
                    return PermissionConstants.InspectionPlan;
                case "FIL_FRM_MH_CHKSLC":
                case "FIL_FRM_NO_IPNDLIST":
                    return PermissionConstants.CheckSelection;
                case "FIL_FRM_MH_PNDLIST":
                    return PermissionConstants.PendingList;
                case "FIL_FRM_DH_INSPDECN":
                    return PermissionConstants.InspectionDecision;
                default:
                    return PermissionConstants.Root;
            }
        }
    }
}
