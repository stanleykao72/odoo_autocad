// OdooAutoCAD.Core/AutoCAD/OleMessageFilter.cs
// COM Message Filter for handling rejected COM calls from AutoCAD.
//
// AutoCAD (2010+) rejects incoming COM calls when its WPF components are
// performing layout processing. Without this filter, rejected calls cause
// "Unhandled Access Violation" crashes inside AutoCAD.
//
// References:
//   https://keanw.com/2010/02/handling-com-calls-rejected-by-autocad-from-an-external-net-application.html
//   https://infosys.beckhoff.com/content/1033/tc3_automationinterface/242727947.html

using System.Runtime.InteropServices;

namespace OdooAutoCAD.Core.AutoCAD;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int dwCallType, IntPtr hTaskCaller,
        int dwTickCount, IntPtr lpInterfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount,
        int dwRejectType);

    [PreserveSig]
    int MessagePending(IntPtr hTaskCallee, int dwTickCount,
        int dwPendingType);
}

/// <summary>
/// Implements IOleMessageFilter to automatically retry COM calls that AutoCAD
/// rejects during WPF layout processing (RPC_E_CALL_REJECTED / 0x80010001).
/// Must be registered on an STA thread via <see cref="Register"/> before
/// making any COM calls to AutoCAD.
/// </summary>
public class OleMessageFilter : IOleMessageFilter
{
    private const int SERVERCALL_ISHANDLED = 0;
    private const int SERVERCALL_RETRYLATER = 2;
    private const int PENDINGMSG_WAITDEFPROCESS = 2;
    private const int RetryCallAfterMs = 1000;

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(
        IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    /// <summary>
    /// Registers the message filter on the current STA thread.
    /// Safe to call multiple times — each call replaces the previous filter.
    /// </summary>
    public static void Register()
    {
        IOleMessageFilter newFilter = new OleMessageFilter();
        CoRegisterMessageFilter(newFilter, out _);
    }

    /// <summary>
    /// Revokes the message filter on the current STA thread.
    /// </summary>
    public static void Revoke()
    {
        CoRegisterMessageFilter(null, out _);
    }

    int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller,
        int dwTickCount, IntPtr lpInterfaceInfo)
    {
        return SERVERCALL_ISHANDLED;
    }

    int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount,
        int dwRejectType)
    {
        if (dwRejectType == SERVERCALL_RETRYLATER)
        {
            // Retry after 1 second
            return RetryCallAfterMs;
        }

        // Cancel the call
        return -1;
    }

    int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount,
        int dwPendingType)
    {
        return PENDINGMSG_WAITDEFPROCESS;
    }
}
