    // IShellFolder::GetDisplayNameOf/SetNameOf uFlags
    typedef enum {
        SHGDN_NORMAL             = 0,
        SHGDN_INFOLDER           = 1,
        SHGDN_INCLUDE_NONFILESYS = 0x2000,
        SHGDN_FORADDRESSBAR      = 0x4000,
        SHGDN_FORPARSING         = 0x8000,
    } SHGNO;

    // IShellFolder::EnumObjects
    typedef enum {
        SHCONTF_FOLDERS         = 32,
        SHCONTF_NONFOLDERS      = 64,
        SHCONTF_INCLUDEHIDDEN   = 128,
    } SHCONTF;

    // IShellFolder::GetAttributesOf flags
    typedef enum {
        SFGAO_CANCOPY           = 0x00000001, // DROPEFFECT_COPY
        SFGAO_CANMOVE           = 0x00000002, // DROPEFFECT_MOVE
        SFGAO_CANLINK           = 0x00000004, // DROPEFFECT_LINK
        SFGAO_CANRENAME         = 0x00000010,
        SFGAO_CANDELETE         = 0x00000020,
        SFGAO_HASPROPSHEET      = 0x00000040,
        SFGAO_DROPTARGET        = 0x00000100,
        SFGAO_CAPABILITYMASK    = 0x00000177,
        SFGAO_LINK              = 0x00010000,
        SFGAO_SHARE             = 0x00020000,
        SFGAO_READONLY          = 0x00040000,
        SFGAO_GHOSTED           = 0x00080000,
        SFGAO_HIDDEN            = 0x00080000,
        SFGAO_DISPLAYATTRMASK   = 0x000F0000,
        SFGAO_FILESYSANCESTOR   = 0x10000000,
        SFGAO_FOLDER            = 0x20000000,
        SFGAO_FILESYSTEM        = 0x40000000,
        SFGAO_HASSUBFOLDER      = 0x80000000,
        SFGAO_CONTENTSMASK      = 0x80000000,
        SFGAO_VALIDATE          = 0x01000000,
        SFGAO_REMOVABLE         = 0x02000000,
        SFGAO_COMPRESSED        = 0x04000000,
        SFGAO_BROWSABLE         = 0x08000000,
        SFGAO_NONENUMERATED     = 0x00100000,
        SFGAO_NEWCONTENT        = 0x00200000,
    }SFGAO;

    typedef struct _SHITEMID {
        short cb;
        unsigned char  abID[1];
    } SHITEMID;

    typedef struct _ITEMIDLIST {
        SHITEMID mkid;
    } ITEMIDLIST;

    [
        uuid(000214e6-0000-0000-c000-000000000046),
        helpstring("IShellFolder Interface"),
        odl
    ]
    interface IShellFolder : IUnknown
    {
        HRESULT ParseDisplayName([in] long hwndOwner,
                                 [in] long pbcReserved,
                                 [in] long lpszDisplayName,
                                 [in] long * pchEaten,
                                 [in, out] long * ppidl,
                                 [in, out] long *pdwAttributes);

        HRESULT EnumObjects(
                  [in] long hwndOwner,
                  [in] long grfFlags,
                  [out, retval] long * ppenumIDList);

        HRESULT BindToObject([in] long pidl,
                             [in] long pbcReserved,
                             [in] long riid,
                             [out, retval] IUnknown **ppvOut);

        HRESULT BindToStorage([in] long pidl,
                              [in] long pbcReserved,
                              [in] long riid,
                              [in, out] long * ppvObj);

        HRESULT CompareIDs([in] long lParam,
                           [in] long pidl1,
                           [in] long pidl2);

        HRESULT CreateViewObject([in] long hwndOwner,
                                 [in] long riid,
                                 [in, out] long * ppvOut);

        HRESULT GetAttributesOf([in] long cidl,
                                [in, out] long * apidl,
                                [in, out] long * rgfInOut);

        HRESULT GetUIObjectOf([in] long hwndOwner,
                              [in] long cidl,
                              [in, out] long * apidl,
                              [in] long riid,
                              [in, out] long * prgfInOut,
                              [out, retval] IUnknown **ppvOut);

        HRESULT GetDisplayNameOf([in] long pidl,
                                 [in] long uFlags,
                                 [in] long lpName);

        HRESULT SetNameOf([in] long hwndOwner,
                          [in] long pidl,
                          [in] long lpszName,
                          [in] long uFlags,
                          [in, out] long * ppidlOut);
    };
    
[
    dllname("SHELL32.DLL")
]
module shell32 {
    [entry("SHGetDesktopFolder")]
    HRESULT SHGetDesktopFolder(
        [out, retval] IShellFolder **ppshf);
        
    [entry("SHGetMalloc")]
    HRESULT SHGetMalloc(
        [out, retval] IMalloc ** ppMalloc);
}