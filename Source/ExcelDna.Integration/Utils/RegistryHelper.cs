using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;

namespace ExcelDna.Integration.Utils
{
    public static class RegistryHelper
    {
        public static RegistryKey CreateVolatileSubKey(RegistryKey baseKey, string subKeyName, RegistryKeyPermissionCheck permissionCheck)
        {
            var keyType = baseKey.GetType();

            var handleField = keyType.GetField("hkey", BindingFlags.NonPublic | BindingFlags.Instance);
            var baseKeyHandle = (SafeHandleZeroOrMinusOneIsInvalid)(handleField.GetValue(baseKey));

            var accessRights = (int)(RegSAM.QueryValue);
            if (permissionCheck == RegistryKeyPermissionCheck.ReadWriteSubTree)
                accessRights = (int)(RegSAM.QueryValue | RegSAM.CreateSubKey | RegSAM.SetValue);

            int lpdwDisposition;
            var hkResult = IntPtr.Zero;

            try
            {
                var errorCode = RegCreateKeyEx(baseKeyHandle, subKeyName, 0, null, (int)RegOption.Volatile,
                    accessRights, IntPtr.Zero, out hkResult, out lpdwDisposition);

                if (errorCode == 5)
                    throw new UnauthorizedAccessException();

                if (errorCode != 0 || hkResult.ToInt32() <= 0)
                    throw new Win32Exception();
            }
            finally
            {
                if (hkResult.ToInt32() > 0)
                    Marshal.Release(hkResult);
            }

            return baseKey.OpenSubKey(subKeyName, true);
        }

        [Flags]
        public enum RegOption
        {
            NonVolatile = 0x0,
            Volatile = 0x1,
            CreateLink = 0x2,
            BackupRestore = 0x4,
            OpenLink = 0x8
        }

        [Flags]
        [SuppressMessage("ReSharper", "InconsistentNaming")]
        public enum RegSAM
        {
            QueryValue = 0x0001,
            SetValue = 0x0002,
            CreateSubKey = 0x0004,
            EnumerateSubKeys = 0x0008,
            Notify = 0x0010,
            CreateLink = 0x0020,
            WOW64_32Key = 0x0200,
            WOW64_64Key = 0x0100,
            WOW64_Res = 0x0300,
            Read = 0x00020019,
            Write = 0x00020006,
            Execute = 0x00020019,
            AllAccess = 0x000f003f
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        private static extern int RegCreateKeyEx(SafeHandleZeroOrMinusOneIsInvalid hKey, string lpSubKey, int reserved, string lpClass, int dwOptions, int samDesigner, IntPtr lpSecurityAttributes, out IntPtr hkResult, out int lpdwDisposition);
    }
}
