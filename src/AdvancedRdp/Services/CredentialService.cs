using System.Runtime.InteropServices;
using System.Text;

namespace AdvancedRdp.Services;

public class CredentialService
{
    private readonly string _applicationName;

    public CredentialService(string applicationName)
    {
        _applicationName = applicationName;
    }

    public void SavePassword(string key, string password)
    {
        var target = BuildTarget(key);
        var passwordBytes = Encoding.Unicode.GetBytes(password);

        var credential = new NativeMethods.CREDENTIAL
        {
            TargetName = target,
            UserName = Environment.UserName,
            Type = NativeMethods.CRED_TYPE_GENERIC,
            Persist = NativeMethods.CRED_PERSIST_LOCAL_MACHINE,
            CredentialBlobSize = (uint)passwordBytes.Length,
            CredentialBlob = Marshal.StringToCoTaskMemUni(password)
        };

        try
        {
            if (!NativeMethods.CredWrite(ref credential, 0))
            {
                throw new InvalidOperationException($"CredWrite failed: {Marshal.GetLastWin32Error()}");
            }
        }
        finally
        {
            if (credential.CredentialBlob != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(credential.CredentialBlob);
            }
        }
    }

    public string? GetPassword(string key)
    {
        var target = BuildTarget(key);
        if (!NativeMethods.CredRead(target, NativeMethods.CRED_TYPE_GENERIC, 0, out var credPtr))
        {
            return null;
        }

        try
        {
            var credential = Marshal.PtrToStructure<NativeMethods.CREDENTIAL>(credPtr);
            if (credential.CredentialBlobSize == 0 || credential.CredentialBlob == IntPtr.Zero)
            {
                return null;
            }

            return Marshal.PtrToStringUni(credential.CredentialBlob, (int)credential.CredentialBlobSize / 2);
        }
        finally
        {
            NativeMethods.CredFree(credPtr);
        }
    }

    public void DeletePassword(string key)
    {
        var target = BuildTarget(key);
        NativeMethods.CredDelete(target, NativeMethods.CRED_TYPE_GENERIC, 0);
    }

    private string BuildTarget(string key) => $"{_applicationName}/{key}";

    private static class NativeMethods
    {
        public const int CRED_TYPE_GENERIC = 1;
        public const int CRED_PERSIST_LOCAL_MACHINE = 2;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct CREDENTIAL
        {
            public uint Flags;
            public uint Type;
            public string TargetName;
            public string? Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public uint CredentialBlobSize;
            public IntPtr CredentialBlob;
            public uint Persist;
            public uint AttributeCount;
            public IntPtr Attributes;
            public string? TargetAlias;
            public string UserName;
        }

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool CredWrite([In] ref CREDENTIAL userCredential, [In] uint flags);

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credentialPtr);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool CredDelete(string target, int type, int flags);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern void CredFree([In] IntPtr cred);
    }
}
