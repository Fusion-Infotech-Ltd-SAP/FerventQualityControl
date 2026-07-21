using System;
using System.Runtime.InteropServices;

namespace QualityControl.Helper
{
    internal enum AuthorizationLevel
    {
        None = 0,
        ReadOnly = 1,
        Full = 2
    }

    internal static class AuthorizationService
    {
        //public static AuthorizationLevel GetAuthorization(string permissionId)
        //{
        //    AuthorizationLevel diPermission = GetAuthorizationFromDiUser(permissionId);
        //    if (diPermission != AuthorizationLevel.None)
        //        return diPermission;

        //    try
        //    {
        //        dynamic company = Global.G_UI_Application.Company;
        //        object permission = company.GetPermission(permissionId);
        //        return ConvertPermission(permission);
        //    }
        //    catch(Exception ex)
        //    {
        //        return AuthorizationLevel.None;
        //    }
        //}

        public static AuthorizationLevel GetAuthorization(string permissionId)
        {
            try
            {
                return GetAuthorizationFromDiUser(permissionId);
            }
            catch (Exception ex)
            {
                Global.GFunc.ShowError(
                    "Permission check error: " + ex.Message);

                return AuthorizationLevel.None;
            }
        }
        private static AuthorizationLevel GetAuthorizationFromDiUser(string permissionId)
        {
            SAPbobsCOM.Users user = null;

            try
            {
                user = (SAPbobsCOM.Users)Global.oComp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);
                if (!user.GetByKey(Global.oComp.UserSignature))
                    return AuthorizationLevel.None;

                if (user.Superuser == SAPbobsCOM.BoYesNoEnum.tYES)
                    return AuthorizationLevel.Full;

                SAPbobsCOM.UserPermission permissions = user.UserPermission;
                for (int i = 0; i < permissions.Count; i++)
                {
                    permissions.SetCurrentLine(i);
                    if (!string.Equals(permissions.PermissionID, permissionId, StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (permissions.Permission == SAPbobsCOM.BoPermission.boper_Full)
                        return AuthorizationLevel.Full;

                    if (permissions.Permission == SAPbobsCOM.BoPermission.boper_ReadOnly)
                        return AuthorizationLevel.ReadOnly;

                    return AuthorizationLevel.None;
                }
            }
            catch
            {
            }
            finally
            {
                if (user != null)
                    Marshal.ReleaseComObject(user);
            }

            return AuthorizationLevel.None;
        }

        private static AuthorizationLevel ConvertPermission(object permission)
        {
            if (permission == null)
                return AuthorizationLevel.None;

            string value = permission.ToString();

            if (value.IndexOf("Full", StringComparison.OrdinalIgnoreCase) >= 0)
                return AuthorizationLevel.Full;

            if (value.IndexOf("Read", StringComparison.OrdinalIgnoreCase) >= 0)
                return AuthorizationLevel.ReadOnly;

            int code;
            if (int.TryParse(value, out code))
            {
                if (code == 2)
                    return AuthorizationLevel.Full;

                if (code == 1)
                    return AuthorizationLevel.ReadOnly;
            }

            return AuthorizationLevel.None;
        }

        public static bool CanView(string permissionId)
        {
            AuthorizationLevel level = GetAuthorization(permissionId);
            return level == AuthorizationLevel.ReadOnly || level == AuthorizationLevel.Full;
        }

        public static bool CanFull(string permissionId)
        {
            return GetAuthorization(permissionId) == AuthorizationLevel.Full;
        }

        public static bool EnsureView(string permissionId, string actionName)
        {
            if (CanView(permissionId))
                return true;

            ShowNoAuthorization(permissionId, actionName);
            return false;
        }

        public static bool EnsureFull(string permissionId, string actionName)
        {
            if (CanFull(permissionId))
                return true;

            ShowNoAuthorization(permissionId, actionName);
            return false;
        }

        //private static void ShowNoAuthorization(string actionName)
        //{
        //    Global.GFunc.ShowError("No authorization for " + actionName + ".");
        //}

        private static void ShowNoAuthorization(string permissionId, string actionName)
        {
            AuthorizationLevel level = GetAuthorization(permissionId);

            string mode = "No Authorization";

            if (level == AuthorizationLevel.Full)
                mode = "Full";
            else if (level == AuthorizationLevel.ReadOnly)
                mode = "Read Only";

            Global.GFunc.ShowError(
                "No authorization for " + actionName +
                ". Current Permission Mode: " + mode);
        }
    }
}
