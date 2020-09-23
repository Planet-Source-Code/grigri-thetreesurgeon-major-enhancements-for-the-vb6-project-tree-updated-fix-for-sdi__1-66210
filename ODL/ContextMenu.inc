  typedef struct {
    long cbSize;
    long fMask;
    long hwnd;
    long lpVerb;
    long lpParameters;
    long lpDirectory;
    long nShow;
    long dwHotKey;
    long hIcon;
  } CMINVOKECOMMANDINFO;
  
  typedef enum {
    GCS_VERBA     = 0x00000000,
    GCS_HELPTEXTA = 0x00000001,
    GCS_VALIDATEA = 0x00000002,
    GCS_VERBW     = 0x00000004,
    GCS_HELPTEXTW = 0x00000005,
    GCS_VALIDATEW = 0x00000006,
    GCS_UNICODE   = 0x00000004,
  } GETCOMMANDSTRINGFLAGS;

  typedef enum {
    CMF_NORMAL        = 0x00000000,
    CMF_DEFAULTONLY   = 0x00000001,
    CMF_VERBSONLY     = 0x00000002,
    CMF_EXPLORE       = 0x00000004,
    CMF_NOVERBS       = 0x00000008,
    CMF_CANRENAME     = 0x00000010,
    CMF_NODEFAULT     = 0x00000020,
    CMF_INCLUDESTATIC = 0x00000040,
    CMF_RESERVED      = 0xffff0000
  } QUERYCONTEXTMENUFLAGS;
  
  [
      uuid(000214e4-0000-0000-c000-000000000046),
      helpstring("IContextMenu Interface"),
      odl
  ]
  interface IContextMenu : IUnknown
  {
      HRESULT QueryContextMenu(
                  [in] long hMenu,
                  [in] long indexMenu,
                  [in] long idCmdFirst,
                  [in] long idCmdLast,
                  [in] QUERYCONTEXTMENUFLAGS uFlags);

      HRESULT InvokeCommand([in] CMINVOKECOMMANDINFO * lpcmi);

      HRESULT GetCommandString([in] long    idCmd,
                               [in] GETCOMMANDSTRINGFLAGS uType,
                               [in, out] int *pwReserved,
                               [in] long    pszName,
                               [in] long    cchMax);
  }
  
[
    uuid(000214F4-0000-0000-C000-000000000046),
    helpstring("IContextMenu2 interface"),
    odl
]
interface IContextMenu2 : stdole.IUnknown {

    HRESULT QueryContextMenu(
        [in] long hMenu,
        [in] long indexMenu,
        [in] long idCmdFirst,
        [in] long idCmdLast,
        [in] QUERYCONTEXTMENUFLAGS uFlags);

    HRESULT InvokeCommand(
        [in, out] CMINVOKECOMMANDINFO* lpici);
    
    HRESULT GetCommandString(
        [in] long idCmd,
        [in] GETCOMMANDSTRINGFLAGS uType,
        [in, out] int *pwReserved,
        [in] long pszName,
        [in] int cchMax);

    HRESULT HandleMenuMsg(
        [in] long uMsg,
        [in] long wParam,
        [in] long lParam);
};

[
    uuid(bcfce0a0-ec17-11d0-8d10-00a0c90f2719),
    helpstring("IContextMenu3 interface"),
    odl
]
interface IContextMenu3 : stdole.IUnknown {

    HRESULT QueryContextMenu(
        [in] long hMenu,
        [in] long indexMenu,
        [in] long idCmdFirst,
        [in] long idCmdLast,
        [in] QUERYCONTEXTMENUFLAGS uFlags);

    HRESULT InvokeCommand(
        [in, out] CMINVOKECOMMANDINFO* lpici);
    
    HRESULT GetCommandString(
        [in] long idCmd,
        [in] GETCOMMANDSTRINGFLAGS uType,
        [in, out] int *pwReserved,
        [in] long pszName,
        [in] int cchMax);

    HRESULT HandleMenuMsg(
        [in] long uMsg,
        [in] long wParam,
        [in] long lParam);
        
    HRESULT HandleMenuMsg2(
        [in] long uMsg,
        [in] long wParam,
        [in] long lParam,
        [out, retval] long *pResult);
};