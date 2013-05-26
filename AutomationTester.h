// AutomationTester.h : Declaration of the AutomationTester
#pragma once
#include "resource.h"       // main symbols
#include <atlctl.h>
#include <atlsafe.h>
#include "COMAutomationTester_i.h"
#include "_IAutomationTesterEvents_CP.h"

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


class AutomationTesterLic
{
protected:
   static BOOL VerifyLicenseKey(BSTR bstr)
   {
      return !lstrcmpW(bstr, L"AutomationTester license");
   }

   static BOOL GetLicenseKey(DWORD dwReserved, BSTR* pBstr)
   {
      if( pBstr == NULL )
        return FALSE;
      *pBstr = SysAllocString(L"AutomationTester license");
      return TRUE;
   }

   static BOOL IsLicenseValid()
   {
       return TRUE;
   }
};


// AutomationTester
class ATL_NO_VTABLE AutomationTester :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CStockPropImpl<AutomationTester, IAutomationTester>,
    public IPersistStreamInitImpl<AutomationTester>,
    public IOleControlImpl<AutomationTester>,
    public IOleObjectImpl<AutomationTester>,
    public IOleInPlaceActiveObjectImpl<AutomationTester>,
    public IViewObjectExImpl<AutomationTester>,
    public IOleInPlaceObjectWindowlessImpl<AutomationTester>,
    public IConnectionPointContainerImpl<AutomationTester>,
    public CProxy_IAutomationTesterEvents<AutomationTester>,
    public IPersistStorageImpl<AutomationTester>,
    public IQuickActivateImpl<AutomationTester>,
    public IPropertyNotifySinkCP<AutomationTester>,
    public CComCoClass<AutomationTester, &CLSID_AutomationTester>,
    public CComControl<AutomationTester>
{
public:
    CContainedWindow m_ctlEdit;

             AutomationTester();
    virtual ~AutomationTester();

DECLARE_CLASSFACTORY2(AutomationTesterLic)

DECLARE_OLEMISC_STATUS(OLEMISC_RECOMPOSEONRESIZE |
    OLEMISC_CANTLINKINSIDE |
    OLEMISC_INSIDEOUT |
    OLEMISC_ACTIVATEWHENVISIBLE |
    OLEMISC_SETCLIENTSITEFIRST
)

DECLARE_REGISTRY_RESOURCEID(IDR_AUTOMATIONTESTER)


BEGIN_COM_MAP(AutomationTester)
    COM_INTERFACE_ENTRY(IAutomationTester)
    COM_INTERFACE_ENTRY(IDispatch)
    COM_INTERFACE_ENTRY(IViewObjectEx)
    COM_INTERFACE_ENTRY(IViewObject2)
    COM_INTERFACE_ENTRY(IViewObject)
    COM_INTERFACE_ENTRY(IOleInPlaceObjectWindowless)
    COM_INTERFACE_ENTRY(IOleInPlaceObject)
    COM_INTERFACE_ENTRY2(IOleWindow, IOleInPlaceObjectWindowless)
    COM_INTERFACE_ENTRY(IOleInPlaceActiveObject)
    COM_INTERFACE_ENTRY(IOleControl)
    COM_INTERFACE_ENTRY(IOleObject)
    COM_INTERFACE_ENTRY(IPersistStreamInit)
    COM_INTERFACE_ENTRY2(IPersist, IPersistStreamInit)
    COM_INTERFACE_ENTRY(IConnectionPointContainer)
    COM_INTERFACE_ENTRY(IQuickActivate)
    COM_INTERFACE_ENTRY(IPersistStorage)
END_COM_MAP()

BEGIN_PROP_MAP(AutomationTester)
    PROP_DATA_ENTRY("_cx", m_sizeExtent.cx, VT_UI4)
    PROP_DATA_ENTRY("_cy", m_sizeExtent.cy, VT_UI4)
#ifndef _WIN32_WCE
    PROP_ENTRY_TYPE("BackColor", DISPID_BACKCOLOR, CLSID_StockColorPage, VT_UI4)
#endif
#ifndef _WIN32_WCE
    PROP_ENTRY_TYPE("Font", DISPID_FONT, CLSID_StockFontPage, VT_DISPATCH)
#endif
    PROP_ENTRY_TYPE("Text", DISPID_TEXT, CLSID_NULL, VT_BSTR)
    // Example entries
    // PROP_ENTRY_TYPE("Property Name", dispid, clsid, vtType)
    // PROP_PAGE(CLSID_StockColorPage)
END_PROP_MAP()

BEGIN_CONNECTION_POINT_MAP(AutomationTester)
    CONNECTION_POINT_ENTRY(IID_IPropertyNotifySink)
    CONNECTION_POINT_ENTRY(__uuidof(_IAutomationTesterEvents))
END_CONNECTION_POINT_MAP()

BEGIN_MSG_MAP(AutomationTester)
    MESSAGE_HANDLER(WM_CREATE, OnCreate)
    MESSAGE_HANDLER(WM_SETFOCUS, OnSetFocus)
    CHAIN_MSG_MAP(CComControl<AutomationTester>)
ALT_MSG_MAP(1)
    // Replace this with message map entries for superclassed Edit
END_MSG_MAP()
// Handler prototypes:
//  LRESULT MessageHandler(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled);
//  LRESULT CommandHandler(WORD wNotifyCode, WORD wID, HWND hWndCtl, BOOL& bHandled);
//  LRESULT NotifyHandler(int idCtrl, LPNMHDR pnmh, BOOL& bHandled);

    BOOL PreTranslateAccelerator(LPMSG pMsg, HRESULT& hRet)
    {
        if(pMsg->message == WM_KEYDOWN)
        {
            switch(pMsg->wParam)
            {
            case VK_LEFT:
            case VK_RIGHT:
            case VK_UP:
            case VK_DOWN:
            case VK_HOME:
            case VK_END:
            case VK_NEXT:
            case VK_PRIOR:
                hRet = S_FALSE;
                return TRUE;
            }
        }
        //TODO: Add your additional accelerator handling code here
        return FALSE;
    }

    LRESULT OnSetFocus(UINT uMsg, WPARAM wParam, LPARAM lParam, BOOL& bHandled)
    {
        LRESULT lRes = CComControl<AutomationTester>::OnSetFocus(uMsg, wParam, lParam, bHandled);
        if (m_bInPlaceActive)
        {
            if(!IsChild(::GetFocus()))
                m_ctlEdit.SetFocus();
        }
        return lRes;
    }

    LRESULT OnCreate(UINT /*uMsg*/, WPARAM /*wParam*/, LPARAM /*lParam*/, BOOL& /*bHandled*/)
    {
        RECT rc;
        GetWindowRect(&rc);
        rc.right -= rc.left;
        rc.bottom -= rc.top;
        rc.top = rc.left = 0;
        m_ctlEdit.Create(m_hWnd, rc, nullptr, WS_CHILD | WS_VISIBLE   | WS_VSCROLL |
                                              ES_LEFT  | ES_MULTILINE | ES_AUTOVSCROLL);
        m_ctlEdit.SetWindowText(m_bstrText);
        return 0;
    }

    STDMETHOD(SetObjectRects)(LPCRECT prcPos,LPCRECT prcClip)
    {
        IOleInPlaceObjectWindowlessImpl<AutomationTester>::SetObjectRects(prcPos, prcClip);
        int cx, cy;
        cx = prcPos->right - prcPos->left;
        cy = prcPos->bottom - prcPos->top;
        ::SetWindowPos(m_ctlEdit.m_hWnd, NULL, 0,
            0, cx, cy, SWP_NOZORDER | SWP_NOACTIVATE);
        return S_OK;
    }

// IViewObjectEx
    DECLARE_VIEW_STATUS(VIEWSTATUS_SOLIDBKGND | VIEWSTATUS_OPAQUE)

// IAutomationTester
    OLE_COLOR m_clrBackColor;
    void OnBackColorChanged()
    {
        ATLTRACE(_T("OnBackColorChanged\n"));
    }
    CComPtr<IFontDisp> m_pFont;
    void OnFontChanged()
    {
        ATLTRACE(_T("OnFontChanged\n"));
    }
    CComBSTR m_bstrText;
    void OnTextChanged()
    {
        ATLTRACE(_T("OnTextChanged\n"));
        if (!::IsWindow(m_ctlEdit.m_hWnd)) return;
        
        m_ctlEdit.SetWindowText(m_bstrText);
    }

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    HRESULT FinalConstruct()
    {
        return S_OK;
    }

    void FinalRelease()
    {
    }
    
    STDMETHOD(get_PropBSTR       )(BSTR        * pVal  );
    STDMETHOD(put_PropBSTR       )(BSTR          newVal);
    STDMETHOD(get_PropBYTE       )(BYTE        * pVal  );
    STDMETHOD(put_PropBYTE       )(BYTE          newVal);
    STDMETHOD(get_PropCHAR       )(CHAR        * pVal  );
    STDMETHOD(put_PropCHAR       )(CHAR          newVal);
    STDMETHOD(get_PropCY         )(CY          * pVal  );
    STDMETHOD(put_PropCY         )(CY            newVal);
    STDMETHOD(get_PropDATE       )(DATE        * pVal  );
    STDMETHOD(put_PropDATE       )(DATE          newVal);
    STDMETHOD(get_PropDECIMAL    )(DECIMAL     * pVal  );
    STDMETHOD(put_PropDECIMAL    )(DECIMAL       newVal);
    STDMETHOD(get_PropDOUBLE     )(DOUBLE      * pVal  );
    STDMETHOD(put_PropDOUBLE     )(DOUBLE        newVal);
    STDMETHOD(get_PropFLOAT      )(FLOAT       * pVal  );
    STDMETHOD(put_PropFLOAT      )(FLOAT         newVal);
    STDMETHOD(get_PropLONG       )(LONG        * pVal  );
    STDMETHOD(put_PropLONG       )(LONG          newVal);
    STDMETHOD(get_PropLONGLONG   )(LONGLONG    * pVal  );
    STDMETHOD(put_PropLONGLONG   )(LONGLONG      newVal);
    STDMETHOD(get_PropSCODE      )(SCODE       * pVal  );
    STDMETHOD(put_PropSCODE      )(SCODE         newVal);
    STDMETHOD(get_PropSHORT      )(SHORT       * pVal  );
    STDMETHOD(put_PropSHORT      )(SHORT         newVal);
    STDMETHOD(get_PropULONG      )(ULONG       * pVal  );
    STDMETHOD(put_PropULONG      )(ULONG         newVal);
    STDMETHOD(get_PropULONGLONG  )(ULONGLONG   * pVal  );
    STDMETHOD(put_PropULONGLONG  )(ULONGLONG     newVal);
    STDMETHOD(get_PropUSHORT     )(USHORT      * pVal  );
    STDMETHOD(put_PropUSHORT     )(USHORT        newVal);
    STDMETHOD(get_PropVARIANTBOOL)(VARIANT_BOOL* pVal  );
    STDMETHOD(put_PropVARIANTBOOL)(VARIANT_BOOL  newVal);
    STDMETHOD(get_PropDISPATCH   )(IDispatch*  * pVal  );
    STDMETHOD(put_PropDISPATCH   )(IDispatch*    newVal);
    STDMETHOD(get_PropUNKNOWN    )(IUnknown*   * pVal  );
    STDMETHOD(put_PropUNKNOWN    )(IUnknown*     newVal);
    STDMETHOD(get_PropAryBSTR       )(SAFEARRAY/*(BSTR        )*/** pVal  );
    STDMETHOD(put_PropAryBSTR       )(SAFEARRAY/*(BSTR        )*/*  newVal);
    STDMETHOD(get_PropAryBYTE       )(SAFEARRAY/*(BYTE        )*/** pVal  );
    STDMETHOD(put_PropAryBYTE       )(SAFEARRAY/*(BYTE        )*/*  newVal);
    STDMETHOD(get_PropAryCHAR       )(SAFEARRAY/*(CHAR        )*/** pVal  );
    STDMETHOD(put_PropAryCHAR       )(SAFEARRAY/*(CHAR        )*/*  newVal);
    STDMETHOD(get_PropAryCY         )(SAFEARRAY/*(CY          )*/** pVal  );
    STDMETHOD(put_PropAryCY         )(SAFEARRAY/*(CY          )*/*  newVal);
    STDMETHOD(get_PropAryDATE       )(SAFEARRAY/*(DATE        )*/** pVal  );
    STDMETHOD(put_PropAryDATE       )(SAFEARRAY/*(DATE        )*/*  newVal);
    STDMETHOD(get_PropAryDECIMAL    )(SAFEARRAY/*(DECIMAL     )*/** pVal  );
    STDMETHOD(put_PropAryDECIMAL    )(SAFEARRAY/*(DECIMAL     )*/*  newVal);
    STDMETHOD(get_PropAryDOUBLE     )(SAFEARRAY/*(DOUBLE      )*/** pVal  );
    STDMETHOD(put_PropAryDOUBLE     )(SAFEARRAY/*(DOUBLE      )*/*  newVal);
    STDMETHOD(get_PropAryFLOAT      )(SAFEARRAY/*(FLOAT       )*/** pVal  );
    STDMETHOD(put_PropAryFLOAT      )(SAFEARRAY/*(FLOAT       )*/*  newVal);
    STDMETHOD(get_PropAryLONG       )(SAFEARRAY/*(LONG        )*/** pVal  );
    STDMETHOD(put_PropAryLONG       )(SAFEARRAY/*(LONG        )*/*  newVal);
    STDMETHOD(get_PropAryLONGLONG   )(SAFEARRAY/*(LONGLONG    )*/** pVal  );
    STDMETHOD(put_PropAryLONGLONG   )(SAFEARRAY/*(LONGLONG    )*/*  newVal);
    STDMETHOD(get_PropArySCODE      )(SAFEARRAY/*(SCODE       )*/** pVal  );
    STDMETHOD(put_PropArySCODE      )(SAFEARRAY/*(SCODE       )*/*  newVal);
    STDMETHOD(get_PropArySHORT      )(SAFEARRAY/*(SHORT       )*/** pVal  );
    STDMETHOD(put_PropArySHORT      )(SAFEARRAY/*(SHORT       )*/*  newVal);
    STDMETHOD(get_PropAryULONG      )(SAFEARRAY/*(ULONG       )*/** pVal  );
    STDMETHOD(put_PropAryULONG      )(SAFEARRAY/*(ULONG       )*/*  newVal);
    STDMETHOD(get_PropAryULONGLONG  )(SAFEARRAY/*(ULONGLONG   )*/** pVal  );
    STDMETHOD(put_PropAryULONGLONG  )(SAFEARRAY/*(ULONGLONG   )*/*  newVal);
    STDMETHOD(get_PropAryUSHORT     )(SAFEARRAY/*(USHORT      )*/** pVal  );
    STDMETHOD(put_PropAryUSHORT     )(SAFEARRAY/*(USHORT      )*/*  newVal);
    STDMETHOD(get_PropAryVARIANTBOOL)(SAFEARRAY/*(VARIANT_BOOL)*/** pVal  );
    STDMETHOD(put_PropAryVARIANTBOOL)(SAFEARRAY/*(VARIANT_BOOL)*/*  newVal);
    STDMETHOD(get_PropAryDISPATCH   )(SAFEARRAY/*(IDispatch*  )*/** pVal  );
    STDMETHOD(put_PropAryDISPATCH   )(SAFEARRAY/*(IDispatch*  )*/*  newVal);
    STDMETHOD(get_PropAryUNKNOWN    )(SAFEARRAY/*(IUnknown*   )*/** pVal  );
    STDMETHOD(put_PropAryUNKNOWN    )(SAFEARRAY/*(IUnknown*   )*/*  newVal);
    STDMETHOD(testBSTR       )(BSTR         valIn, BSTR        * refIn, BSTR        * refInOut, BSTR        * refRet);
    STDMETHOD(testBYTE       )(BYTE         valIn, BYTE        * refIn, BYTE        * refInOut, BYTE        * refRet);
    STDMETHOD(testCHAR       )(CHAR         valIn, CHAR        * refIn, CHAR        * refInOut, CHAR        * refRet);
    STDMETHOD(testCY         )(CY           valIn, CY          * refIn, CY          * refInOut, CY          * refRet);
    STDMETHOD(testDATE       )(DATE         valIn, DATE        * refIn, DATE        * refInOut, DATE        * refRet);
    STDMETHOD(testDECIMAL    )(DECIMAL      valIn, DECIMAL     * refIn, DECIMAL     * refInOut, DECIMAL     * refRet);
    STDMETHOD(testDOUBLE     )(DOUBLE       valIn, DOUBLE      * refIn, DOUBLE      * refInOut, DOUBLE      * refRet);
    STDMETHOD(testFLOAT      )(FLOAT        valIn, FLOAT       * refIn, FLOAT       * refInOut, FLOAT       * refRet);
    STDMETHOD(testLONG       )(LONG         valIn, LONG        * refIn, LONG        * refInOut, LONG        * refRet);
    STDMETHOD(testLONGLONG   )(LONGLONG     valIn, LONGLONG    * refIn, LONGLONG    * refInOut, LONGLONG    * refRet);
    STDMETHOD(testSCODE      )(SCODE        valIn, SCODE       * refIn, SCODE       * refInOut, SCODE       * refRet);
    STDMETHOD(testSHORT      )(SHORT        valIn, SHORT       * refIn, SHORT       * refInOut, SHORT       * refRet);
    STDMETHOD(testULONG      )(ULONG        valIn, ULONG       * refIn, ULONG       * refInOut, ULONG       * refRet);
    STDMETHOD(testULONGLONG  )(ULONGLONG    valIn, ULONGLONG   * refIn, ULONGLONG   * refInOut, ULONGLONG   * refRet);
    STDMETHOD(testUSHORT     )(USHORT       valIn, USHORT      * refIn, USHORT      * refInOut, USHORT      * refRet);
    STDMETHOD(testVARIANTBOOL)(VARIANT_BOOL valIn, VARIANT_BOOL* refIn, VARIANT_BOOL* refInOut, VARIANT_BOOL* refRet);
    STDMETHOD(testDISPATCH   )(IDispatch*   valIn, IDispatch*  * refIn, IDispatch*  * refInOut, IDispatch*  * refRet);
    STDMETHOD(testUNKNOWN    )(IUnknown*    valIn, IUnknown*   * refIn, IUnknown*   * refInOut, IUnknown*   * refRet);
    STDMETHOD(testAryBSTR       )(SAFEARRAY/*(BSTR        )*/* valIn, SAFEARRAY/*(BSTR        )*/** refIn, SAFEARRAY/*(BSTR        )*/** refInOut, SAFEARRAY/*(BSTR        )*/** refRet);
    STDMETHOD(testAryBYTE       )(SAFEARRAY/*(BYTE        )*/* valIn, SAFEARRAY/*(BYTE        )*/** refIn, SAFEARRAY/*(BYTE        )*/** refInOut, SAFEARRAY/*(BYTE        )*/** refRet);
    STDMETHOD(testAryCHAR       )(SAFEARRAY/*(CHAR        )*/* valIn, SAFEARRAY/*(CHAR        )*/** refIn, SAFEARRAY/*(CHAR        )*/** refInOut, SAFEARRAY/*(CHAR        )*/** refRet);
    STDMETHOD(testAryCY         )(SAFEARRAY/*(CY          )*/* valIn, SAFEARRAY/*(CY          )*/** refIn, SAFEARRAY/*(CY          )*/** refInOut, SAFEARRAY/*(CY          )*/** refRet);
    STDMETHOD(testAryDATE       )(SAFEARRAY/*(DATE        )*/* valIn, SAFEARRAY/*(DATE        )*/** refIn, SAFEARRAY/*(DATE        )*/** refInOut, SAFEARRAY/*(DATE        )*/** refRet);
    STDMETHOD(testAryDECIMAL    )(SAFEARRAY/*(DECIMAL     )*/* valIn, SAFEARRAY/*(DECIMAL     )*/** refIn, SAFEARRAY/*(DECIMAL     )*/** refInOut, SAFEARRAY/*(DECIMAL     )*/** refRet);
    STDMETHOD(testAryDOUBLE     )(SAFEARRAY/*(DOUBLE      )*/* valIn, SAFEARRAY/*(DOUBLE      )*/** refIn, SAFEARRAY/*(DOUBLE      )*/** refInOut, SAFEARRAY/*(DOUBLE      )*/** refRet);
    STDMETHOD(testAryFLOAT      )(SAFEARRAY/*(FLOAT       )*/* valIn, SAFEARRAY/*(FLOAT       )*/** refIn, SAFEARRAY/*(FLOAT       )*/** refInOut, SAFEARRAY/*(FLOAT       )*/** refRet);
    STDMETHOD(testAryLONG       )(SAFEARRAY/*(LONG        )*/* valIn, SAFEARRAY/*(LONG        )*/** refIn, SAFEARRAY/*(LONG        )*/** refInOut, SAFEARRAY/*(LONG        )*/** refRet);
    STDMETHOD(testAryLONGLONG   )(SAFEARRAY/*(LONGLONG    )*/* valIn, SAFEARRAY/*(LONGLONG    )*/** refIn, SAFEARRAY/*(LONGLONG    )*/** refInOut, SAFEARRAY/*(LONGLONG    )*/** refRet);
    STDMETHOD(testArySCODE      )(SAFEARRAY/*(SCODE       )*/* valIn, SAFEARRAY/*(SCODE       )*/** refIn, SAFEARRAY/*(SCODE       )*/** refInOut, SAFEARRAY/*(SCODE       )*/** refRet);
    STDMETHOD(testArySHORT      )(SAFEARRAY/*(SHORT       )*/* valIn, SAFEARRAY/*(SHORT       )*/** refIn, SAFEARRAY/*(SHORT       )*/** refInOut, SAFEARRAY/*(SHORT       )*/** refRet);
    STDMETHOD(testAryULONG      )(SAFEARRAY/*(ULONG       )*/* valIn, SAFEARRAY/*(ULONG       )*/** refIn, SAFEARRAY/*(ULONG       )*/** refInOut, SAFEARRAY/*(ULONG       )*/** refRet);
    STDMETHOD(testAryULONGLONG  )(SAFEARRAY/*(ULONGLONG   )*/* valIn, SAFEARRAY/*(ULONGLONG   )*/** refIn, SAFEARRAY/*(ULONGLONG   )*/** refInOut, SAFEARRAY/*(ULONGLONG   )*/** refRet);
    STDMETHOD(testAryUSHORT     )(SAFEARRAY/*(USHORT      )*/* valIn, SAFEARRAY/*(USHORT      )*/** refIn, SAFEARRAY/*(USHORT      )*/** refInOut, SAFEARRAY/*(USHORT      )*/** refRet);
    STDMETHOD(testAryVARIANTBOOL)(SAFEARRAY/*(VARIANT_BOOL)*/* valIn, SAFEARRAY/*(VARIANT_BOOL)*/** refIn, SAFEARRAY/*(VARIANT_BOOL)*/** refInOut, SAFEARRAY/*(VARIANT_BOOL)*/** refRet);
    STDMETHOD(testAryDISPATCH   )(SAFEARRAY/*(IDispatch*  )*/* valIn, SAFEARRAY/*(IDispatch*  )*/** refIn, SAFEARRAY/*(IDispatch*  )*/** refInOut, SAFEARRAY/*(IDispatch*  )*/** refRet);
    STDMETHOD(testAryUNKNOWN    )(SAFEARRAY/*(IUnknown*   )*/* valIn, SAFEARRAY/*(IUnknown*   )*/** refIn, SAFEARRAY/*(IUnknown*   )*/** refInOut, SAFEARRAY/*(IUnknown*   )*/** refRet);
    
private:
    HRESULT TextAppend(wchar_t const* const szText);
    
    BSTR         _propBSTR       ;
    BYTE         _propBYTE       ;
    CHAR         _propCHAR       ;
    CY           _propCY         ;
    DATE         _propDATE       ;
    DECIMAL      _propDECIMAL    ;
    DOUBLE       _propDOUBLE     ;
    FLOAT        _propFLOAT      ;
    LONG         _propLONG       ;
    LONGLONG     _propLONGLONG   ;
    SCODE        _propSCODE      ;
    SHORT        _propSHORT      ;
    ULONG        _propULONG      ;
    ULONGLONG    _propULONGLONG  ;
    USHORT       _propUSHORT     ;
    VARIANT_BOOL _propVARIANTBOOL;
    IDispatch*   _propDISPATCH   ;
    IUnknown*    _propUNKNOWN    ;
    CComSafeArray<BSTR        > _propAryBSTR       ;
    CComSafeArray<BYTE        > _propAryBYTE       ;
    CComSafeArray<CHAR        > _propAryCHAR       ;
    CComSafeArray<CY          > _propAryCY         ;
    CComSafeArray<DATE        > _propAryDATE       ;
    CComSafeArray<DECIMAL     > _propAryDECIMAL    ;
    CComSafeArray<DOUBLE      > _propAryDOUBLE     ;
    CComSafeArray<FLOAT       > _propAryFLOAT      ;
    CComSafeArray<LONG        > _propAryLONG       ;
    CComSafeArray<LONGLONG    > _propAryLONGLONG   ;
    CComSafeArray<SCODE       > _propArySCODE      ;
    CComSafeArray<SHORT       > _propArySHORT      ;
    CComSafeArray<ULONG       > _propAryULONG      ;
    CComSafeArray<ULONGLONG   > _propAryULONGLONG  ;
    CComSafeArray<USHORT      > _propAryUSHORT     ;
    CComSafeArray<VARIANT_BOOL> _propAryVARIANTBOOL;
    CComSafeArray<IDispatch*  > _propAryDISPATCH   ;
    CComSafeArray<IUnknown*   > _propAryUNKNOWN    ;
};

OBJECT_ENTRY_AUTO(__uuidof(AutomationTester), AutomationTester)
