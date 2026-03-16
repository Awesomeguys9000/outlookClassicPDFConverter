using System;
using System.Runtime.InteropServices;

namespace AttachmentPdfConverter
{
    // ── IDTExtensibility2 (COM Shared Add-in entry point) ──────────────────

    [ComImport, Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        [DispId(1)] void OnConnection(
            [MarshalAs(UnmanagedType.IDispatch)] object Application,
            ext_ConnectMode ConnectMode,
            [MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)] void OnDisconnection(
            ext_DisconnectMode RemoveMode,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)] void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)] void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)] void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    // ── Office Ribbon Interfaces ───────────────────────────────────────────

    [ComImport, Guid("000C0396-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        string GetCustomUI([MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    [ComImport, Guid("000C0395-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonControl
    {
        [DispId(1)] string Id { [return: MarshalAs(UnmanagedType.BStr)] get; }
        [DispId(2)] object Context { [return: MarshalAs(UnmanagedType.IDispatch)] get; }
        [DispId(3)] string Tag { [return: MarshalAs(UnmanagedType.BStr)] get; }
    }

    [ComImport, Guid("000C03A7-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IRibbonUI
    {
        [DispId(1)] void Invalidate();
        [DispId(2)] void InvalidateControl([MarshalAs(UnmanagedType.BStr)] string ControlID);
    }
}
