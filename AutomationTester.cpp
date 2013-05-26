
#include "stdafx.h"
#include "AutomationTester.h"

#include <comutil.h>

#include <cassert>

#undef min
#undef max
#include <algorithm>
#include <type_traits>


#if _MSC_BUILD
    #define PRINTF_I64 L"%I64i"
    #define PRINTF_UI64 L"%I64u"
#else
    #error Add correct printf int64/uint64 for your compiler
#endif

namespace {
    bool Ensure(HRESULT const result)
    {
        auto const okay = SUCCEEDED(result);
        assert(okay);
        return okay;
    }
    
    template<size_t sBufferLen>
    void safe_swprintf(wchar_t (&buffer)[sBufferLen], wchar_t const* const szFormat, ...)
    {
        va_list va; va_start(va, szFormat);
            _vsnwprintf_s(buffer, sBufferLen, szFormat, va);
        va_end(va);
        
        buffer[sBufferLen - 1] = L'\0';
    }

    template<typename T>
    HRESULT PropArrayGet(CComSafeArray<T>& ary, SAFEARRAY** out_ary)
    {
        if (!out_ary) return E_INVALIDARG;
        if ( ary    ) return ary.CopyTo(out_ary);
        
        // Play around with the lower bound. According to specs it shouldn't matter, but it might
        // trip faulty client code.
        CComSafeArrayBound bounds(0u, rand());
        CComSafeArray<T> aryNew;
        auto const hr = aryNew.Create(&bounds);
        if (!Ensure(hr)) return hr;
        
        *out_ary = aryNew.Detach();
        return S_OK;
    }
    
    template<typename T>
    HRESULT PropArrayPut(CComSafeArray<T>& ary, SAFEARRAY*& in_ary)
    { return ary.CopyFrom(in_ary); }
    
    template<typename T>
    void PropArrayAddMinMax(CComSafeArray<T>& ary)
    { ary.Add(std::numeric_limits<T>::min());
      ary.Add(std::numeric_limits<T>::max());}
}


#pragma warning(push)
#pragma warning(disable: 4355) // 'this' : used in base member initializer list
AutomationTester::AutomationTester() : m_ctlEdit(_T("Edit"), this, 1)
                                     , _propBSTR       (nullptr)
                                     , _propBYTE       (0)
                                     , _propCHAR       (0)
                                     , _propDATE       (0)
                                     , _propDOUBLE     (0)
                                     , _propFLOAT      (0)
                                     , _propLONG       (0)
                                     , _propLONGLONG   (0)
                                     , _propSCODE      (0)
                                     , _propSHORT      (0)
                                     , _propULONG      (0)
                                     , _propULONGLONG  (0)
                                     , _propUSHORT     (0)
                                     , _propVARIANTBOOL(0)
                                     , _propDISPATCH   (nullptr)
                                     , _propUNKNOWN    (nullptr)
{
    memset(&_propCY     , 0, sizeof(_propCY     ));
    memset(&_propDECIMAL, 0, sizeof(_propDECIMAL));
    
    CY      cyMin ; cyMin .int64 = std::numeric_limits<LONGLONG>::min();
    CY      cyMax ; cyMax .int64 = std::numeric_limits<LONGLONG>::max();
    DECIMAL decMin; decMin.scale = 0; decMin.sign = 1; decMin.Hi32 = 0xFFFFFFFF; decMin.Lo64 = 0xFFFFFFFFFFFFFFFF;
    DECIMAL decMax = decMin; decMax.sign = 0;
    DECIMAL decNearZero; decNearZero.scale = 28; decNearZero.sign = 0; decNearZero.Hi32 = 0; decNearZero.Lo64 = 1;
    
    _bstr_t bstr = L"Many threats to security can only be defeated from inside. Your mind shall be carefully blanked, and conditioned with the nature and past of a criminal. Join with the criminal and rebellious, endure their squalor and chaos, and then, when it is time, liquidate them from within.";
    _propAryBSTR.Add(bstr);
    
    PropArrayAddMinMax(_propAryBYTE);
    PropArrayAddMinMax(_propAryCHAR);
    _propAryCY.Add(cyMin); _propAryCY.Add(cyMax);
    //PropArrayAddMinMax(_propAryDATE);
    _propAryDECIMAL.Add(decMin); _propAryDECIMAL.Add(decMax); _propAryDECIMAL.Add(decNearZero);
    //PropArrayAddMinMax(_propAryDISPATCH);
    PropArrayAddMinMax(_propAryDOUBLE);
    PropArrayAddMinMax(_propAryFLOAT);
    PropArrayAddMinMax(_propAryLONG);
    PropArrayAddMinMax(_propAryLONGLONG);
    PropArrayAddMinMax(_propArySCODE);
    PropArrayAddMinMax(_propArySHORT);
    PropArrayAddMinMax(_propAryULONG);
    PropArrayAddMinMax(_propAryULONGLONG);
    //PropArrayAddMinMax(_propAryUNKNOWN);
    PropArrayAddMinMax(_propAryUSHORT);
    _propAryVARIANTBOOL.Add((SHORT)0); _propAryVARIANTBOOL.Add(-1);
    
    m_bWindowOnly = TRUE;
}
#pragma warning(pop)

AutomationTester::~AutomationTester()
{
    if (_propDISPATCH   ) _propDISPATCH   ->Release(); _propDISPATCH    = nullptr;
    if (_propUNKNOWN    ) _propUNKNOWN    ->Release(); _propUNKNOWN     = nullptr;
}

STDMETHODIMP AutomationTester::get_PropBSTR       (BSTR        * pVal) { *pVal = _propBSTR       ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropBYTE       (BYTE        * pVal) { *pVal = _propBYTE       ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropCHAR       (CHAR        * pVal) { *pVal = _propCHAR       ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropCY         (CY          * pVal) { *pVal = _propCY         ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropDATE       (DATE        * pVal) { *pVal = _propDATE       ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropDECIMAL    (DECIMAL     * pVal) { *pVal = _propDECIMAL    ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropDOUBLE     (DOUBLE      * pVal) { *pVal = _propDOUBLE     ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropFLOAT      (FLOAT       * pVal) { *pVal = _propFLOAT      ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropLONG       (LONG        * pVal) { *pVal = _propLONG       ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropLONGLONG   (LONGLONG    * pVal) { *pVal = _propLONGLONG   ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropSCODE      (SCODE       * pVal) { *pVal = _propSCODE      ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropSHORT      (SHORT       * pVal) { *pVal = _propSHORT      ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropULONG      (ULONG       * pVal) { *pVal = _propULONG      ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropULONGLONG  (ULONGLONG   * pVal) { *pVal = _propULONGLONG  ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropUSHORT     (USHORT      * pVal) { *pVal = _propUSHORT     ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropVARIANTBOOL(VARIANT_BOOL* pVal) { *pVal = _propVARIANTBOOL; return S_OK; }
STDMETHODIMP AutomationTester::get_PropDISPATCH   (IDispatch*  * pVal) { if (_propDISPATCH   ) _propDISPATCH   ->AddRef(); *pVal = _propDISPATCH   ; return S_OK; }
STDMETHODIMP AutomationTester::get_PropUNKNOWN    (IUnknown*   * pVal) { if (_propUNKNOWN    ) _propUNKNOWN    ->AddRef(); *pVal = _propUNKNOWN    ; return S_OK; }

STDMETHODIMP AutomationTester::get_PropAryBSTR       (SAFEARRAY/*(BSTR        )*/** pVal  ) { return PropArrayGet(_propAryBSTR       , pVal); }
STDMETHODIMP AutomationTester::get_PropAryBYTE       (SAFEARRAY/*(BYTE        )*/** pVal  ) { return PropArrayGet(_propAryBYTE       , pVal); }
STDMETHODIMP AutomationTester::get_PropAryCHAR       (SAFEARRAY/*(CHAR        )*/** pVal  ) { return PropArrayGet(_propAryCHAR       , pVal); }
STDMETHODIMP AutomationTester::get_PropAryCY         (SAFEARRAY/*(CY          )*/** pVal  ) { return PropArrayGet(_propAryCY         , pVal); }
STDMETHODIMP AutomationTester::get_PropAryDATE       (SAFEARRAY/*(DATE        )*/** pVal  ) { return PropArrayGet(_propAryDATE       , pVal); }
STDMETHODIMP AutomationTester::get_PropAryDECIMAL    (SAFEARRAY/*(DECIMAL     )*/** pVal  ) { return PropArrayGet(_propAryDECIMAL    , pVal); }
STDMETHODIMP AutomationTester::get_PropAryDOUBLE     (SAFEARRAY/*(DOUBLE      )*/** pVal  ) { return PropArrayGet(_propAryDOUBLE     , pVal); }
STDMETHODIMP AutomationTester::get_PropAryFLOAT      (SAFEARRAY/*(FLOAT       )*/** pVal  ) { return PropArrayGet(_propAryFLOAT      , pVal); }
STDMETHODIMP AutomationTester::get_PropAryLONG       (SAFEARRAY/*(LONG        )*/** pVal  ) { return PropArrayGet(_propAryLONG       , pVal); }
STDMETHODIMP AutomationTester::get_PropAryLONGLONG   (SAFEARRAY/*(LONGLONG    )*/** pVal  ) { return PropArrayGet(_propAryLONGLONG   , pVal); }
STDMETHODIMP AutomationTester::get_PropArySCODE      (SAFEARRAY/*(SCODE       )*/** pVal  ) { return PropArrayGet(_propArySCODE      , pVal); }
STDMETHODIMP AutomationTester::get_PropArySHORT      (SAFEARRAY/*(SHORT       )*/** pVal  ) { return PropArrayGet(_propArySHORT      , pVal); }
STDMETHODIMP AutomationTester::get_PropAryULONG      (SAFEARRAY/*(ULONG       )*/** pVal  ) { return PropArrayGet(_propAryULONG      , pVal); }
STDMETHODIMP AutomationTester::get_PropAryULONGLONG  (SAFEARRAY/*(ULONGLONG   )*/** pVal  ) { return PropArrayGet(_propAryULONGLONG  , pVal); }
STDMETHODIMP AutomationTester::get_PropAryUSHORT     (SAFEARRAY/*(USHORT      )*/** pVal  ) { return PropArrayGet(_propAryUSHORT     , pVal); }
STDMETHODIMP AutomationTester::get_PropAryVARIANTBOOL(SAFEARRAY/*(VARIANT_BOOL)*/** pVal  ) { return PropArrayGet(_propAryVARIANTBOOL, pVal); }
STDMETHODIMP AutomationTester::get_PropAryDISPATCH   (SAFEARRAY/*(IDispatch*  )*/** pVal  ) { return PropArrayGet(_propAryDISPATCH   , pVal); }
STDMETHODIMP AutomationTester::get_PropAryUNKNOWN    (SAFEARRAY/*(IUnknown*   )*/** pVal  ) { return PropArrayGet(_propAryUNKNOWN    , pVal); }

STDMETHODIMP AutomationTester::put_PropBSTR       (BSTR         newVal) { _propBSTR        = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropBYTE       (BYTE         newVal) { _propBYTE        = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropCHAR       (CHAR         newVal) { _propCHAR        = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropCY         (CY           newVal) { _propCY          = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropDATE       (DATE         newVal) { _propDATE        = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropDECIMAL    (DECIMAL      newVal) { _propDECIMAL     = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropDOUBLE     (DOUBLE       newVal) { _propDOUBLE      = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropFLOAT      (FLOAT        newVal) { _propFLOAT       = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropLONG       (LONG         newVal) { _propLONG        = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropLONGLONG   (LONGLONG     newVal) { _propLONGLONG    = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropSCODE      (SCODE        newVal) { _propSCODE       = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropSHORT      (SHORT        newVal) { _propSHORT       = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropULONG      (ULONG        newVal) { _propULONG       = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropULONGLONG  (ULONGLONG    newVal) { _propULONGLONG   = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropUSHORT     (USHORT       newVal) { _propUSHORT      = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropVARIANTBOOL(VARIANT_BOOL newVal) { _propVARIANTBOOL = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropDISPATCH   (IDispatch*   newVal)
{ if (_propDISPATCH   ) _propDISPATCH   ->Release();
  if (newVal) newVal->AddRef();
  _propDISPATCH    = newVal; return S_OK; }
STDMETHODIMP AutomationTester::put_PropUNKNOWN    (IUnknown*    newVal)
{ if (_propUNKNOWN    ) _propUNKNOWN    ->Release();
  if (newVal) newVal->AddRef();
  _propUNKNOWN     = newVal; return S_OK; }

STDMETHODIMP AutomationTester::put_PropAryBSTR       (SAFEARRAY/*(BSTR        )*/*  newVal) { return PropArrayPut(_propAryBSTR       , newVal); }
STDMETHODIMP AutomationTester::put_PropAryBYTE       (SAFEARRAY/*(BYTE        )*/*  newVal) { return PropArrayPut(_propAryBYTE       , newVal); }
STDMETHODIMP AutomationTester::put_PropAryCHAR       (SAFEARRAY/*(CHAR        )*/*  newVal) { return PropArrayPut(_propAryCHAR       , newVal); }
STDMETHODIMP AutomationTester::put_PropAryCY         (SAFEARRAY/*(CY          )*/*  newVal) { return PropArrayPut(_propAryCY         , newVal); }
STDMETHODIMP AutomationTester::put_PropAryDATE       (SAFEARRAY/*(DATE        )*/*  newVal) { return PropArrayPut(_propAryDATE       , newVal); }
STDMETHODIMP AutomationTester::put_PropAryDECIMAL    (SAFEARRAY/*(DECIMAL     )*/*  newVal) { return PropArrayPut(_propAryDECIMAL    , newVal); }
STDMETHODIMP AutomationTester::put_PropAryDOUBLE     (SAFEARRAY/*(DOUBLE      )*/*  newVal) { return PropArrayPut(_propAryDOUBLE     , newVal); }
STDMETHODIMP AutomationTester::put_PropAryFLOAT      (SAFEARRAY/*(FLOAT       )*/*  newVal) { return PropArrayPut(_propAryFLOAT      , newVal); }
STDMETHODIMP AutomationTester::put_PropAryLONG       (SAFEARRAY/*(LONG        )*/*  newVal) { return PropArrayPut(_propAryLONG       , newVal); }
STDMETHODIMP AutomationTester::put_PropAryLONGLONG   (SAFEARRAY/*(LONGLONG    )*/*  newVal) { return PropArrayPut(_propAryLONGLONG   , newVal); }
STDMETHODIMP AutomationTester::put_PropArySCODE      (SAFEARRAY/*(SCODE       )*/*  newVal) { return PropArrayPut(_propArySCODE      , newVal); }
STDMETHODIMP AutomationTester::put_PropArySHORT      (SAFEARRAY/*(SHORT       )*/*  newVal) { return PropArrayPut(_propArySHORT      , newVal); }
STDMETHODIMP AutomationTester::put_PropAryULONG      (SAFEARRAY/*(ULONG       )*/*  newVal) { return PropArrayPut(_propAryULONG      , newVal); }
STDMETHODIMP AutomationTester::put_PropAryULONGLONG  (SAFEARRAY/*(ULONGLONG   )*/*  newVal) { return PropArrayPut(_propAryULONGLONG  , newVal); }
STDMETHODIMP AutomationTester::put_PropAryUSHORT     (SAFEARRAY/*(USHORT      )*/*  newVal) { return PropArrayPut(_propAryUSHORT     , newVal); }
STDMETHODIMP AutomationTester::put_PropAryVARIANTBOOL(SAFEARRAY/*(VARIANT_BOOL)*/*  newVal) { return PropArrayPut(_propAryVARIANTBOOL, newVal); }
STDMETHODIMP AutomationTester::put_PropAryDISPATCH   (SAFEARRAY/*(IDispatch*  )*/*  newVal) { return PropArrayPut(_propAryDISPATCH   , newVal); }
STDMETHODIMP AutomationTester::put_PropAryUNKNOWN    (SAFEARRAY/*(IUnknown*   )*/*  newVal) { return PropArrayPut(_propAryUNKNOWN    , newVal); }

// METHODS

STDMETHODIMP AutomationTester::testBSTR(BSTR valIn, BSTR* refIn, BSTR* refInOut, BSTR* refRet)
{
    TextAppend(L"");
    TextAppend(L"testBSTR:");
    
    TextAppend(L"\tvalIn:\t[["  ); TextAppend( valIn   ); TextAppend(L"]]");
    TextAppend(L"\trefIn:\t[["  ); TextAppend(*refIn   ); TextAppend(L"]]");
    TextAppend(L"\trefInOut:\t[["); TextAppend(*refInOut); TextAppend(L"]]");
    if (wcscmp(valIn, *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    _bstr_t bstrOut, bstrRet;
    bstrOut += "Input text: "; bstrOut += *refInOut;
    bstrRet += "Return: "    ; bstrRet += bstrOut;
    
    *refInOut = bstrOut.Detach();
    *refRet   = bstrRet.Detach();
    return S_OK;
}

STDMETHODIMP AutomationTester::testBYTE(BYTE valIn, BYTE* refIn, BYTE* refInOut, BYTE* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testBYTE:");
    
    safe_swprintf(buffer, L"\tvalIn:\t%d"   , (int) valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%d"   , (int)*refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%d", (int)*refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet    = *refInOut / 2;
    *refInOut *=  2;
    
    return S_OK;
}

STDMETHODIMP AutomationTester::testCHAR(CHAR valIn, CHAR* refIn, CHAR* refInOut, CHAR* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testCHAR:");
    
    safe_swprintf(buffer, L"\tvalIn:\t%c"   ,  valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%c"   , *refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%c", *refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet   = *refInOut;
    *refInOut = (CHAR)(isupper(*refInOut) ? tolower(*refInOut) : toupper(*refInOut));
    
    return S_OK;
}

STDMETHODIMP AutomationTester::testCY(CY valIn, CY* refIn, CY* refInOut, CY* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testCY:");
    
    safe_swprintf(buffer, L"\tvalIn:\t" PRINTF_I64 L"." PRINTF_I64 L" (" PRINTF_I64 L")"
                        , valIn   . int64 / 10000, std::abs(valIn   . int64 - ((valIn   . int64 / 10000) * 10000))
                        , valIn   . int64);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t" PRINTF_I64 L"." PRINTF_I64 L" (" PRINTF_I64 L")"
                        , refIn   ->int64 / 10000, std::abs(refIn   ->int64 - ((refIn   ->int64 / 10000) * 10000))
                        , refIn   ->int64);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t" PRINTF_I64 L"." PRINTF_I64 L" (" PRINTF_I64 L")"
                        , refInOut->int64 / 10000, std::abs(refInOut->int64 - ((refInOut->int64 / 10000) * 10000))
                        , refInOut->int64);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn.int64 != refIn->int64) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    refRet  ->int64  = refInOut->int64 / 2;
    refInOut->int64 *=  2;
    
    return S_OK;
}

STDMETHODIMP AutomationTester::testDATE(DATE valIn, DATE* refIn, DATE* refInOut, DATE* refRet)
{
    TextAppend(L"");
    TextAppend(L"testDATE:");
    TextAppend(L"\tNOT IMPLEMENTED");
    return E_NOTIMPL;
}

STDMETHODIMP AutomationTester::testDECIMAL(DECIMAL valIn, DECIMAL* refIn, DECIMAL* refInOut, DECIMAL* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testDECIMAL:");
    
    CComVariant valInStr   ; valInStr   .decVal =  valIn   ; valInStr   .vt = VT_DECIMAL;
    CComVariant refInStr   ; refInStr   .decVal = *refIn   ; refInStr   .vt = VT_DECIMAL;
    CComVariant refInOutStr; refInOutStr.decVal = *refInOut; refInOutStr.vt = VT_DECIMAL;
    if (!Ensure(valInStr   .ChangeType(VT_BSTR))) return E_FAIL;
    if (!Ensure(refInStr   .ChangeType(VT_BSTR))) return E_FAIL;
    if (!Ensure(refInOutStr.ChangeType(VT_BSTR))) return E_FAIL;
    
    safe_swprintf(buffer, L"\tvalIn:\t%s"   , valInStr   .bstrVal);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%s"   , refInStr   .bstrVal);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%s", refInOutStr.bstrVal);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    auto const hrCompare = VarDecCmp(&valIn, refIn);
    if (!Ensure(hrCompare)) return E_FAIL;
    if ((hrCompare != VARCMP_EQ) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    DECIMAL decTwo = {0, {0, 0}, 0, {2, 0}}, decOut, decRet;
    VarDecDiv(refInOut, &decTwo, &decRet);
    VarDecMul(refInOut, &decTwo, &decOut);
    
    *refRet   = decRet;
    *refInOut = decOut;
    
    return S_OK;
}

STDMETHODIMP AutomationTester::testDOUBLE(DOUBLE valIn, DOUBLE* refIn, DOUBLE* refInOut, DOUBLE* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testDOUBLE:");
    
    safe_swprintf(buffer, L"\tvalIn:\t%E"   ,  valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%E"   , *refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%E", *refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet    = *refInOut / 2;
    *refInOut *=  2;
    return S_OK;
}

STDMETHODIMP AutomationTester::testFLOAT(FLOAT valIn, FLOAT* refIn, FLOAT* refInOut, FLOAT* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testFLOAT:");
    
    safe_swprintf(buffer, L"\tvalIn:\t%E"   , (double) valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%E"   , (double)*refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%E", (double)*refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet    = *refInOut / 2;
    *refInOut *=  2;
    return S_OK;
}

STDMETHODIMP AutomationTester::testLONG(LONG valIn, LONG* refIn, LONG* refInOut, LONG* refRet)
{
    static_assert(sizeof(LONG) == sizeof(int), "Fixup the format specifiers.");
    
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testLONG:");
    
    safe_swprintf(buffer, L"\tvalIn:\t%d"   ,  valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%d"   , *refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%d", *refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet    = *refInOut / 2;
    *refInOut *=  2;
    return S_OK;
}

STDMETHODIMP AutomationTester::testLONGLONG(LONGLONG valIn, LONGLONG* refIn, LONGLONG* refInOut, LONGLONG* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testLONGLONG:");
    
    safe_swprintf(buffer, L"\tvalIn:\t"    PRINTF_I64 L"",  valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t"    PRINTF_I64 L"", *refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t" PRINTF_I64 L"", *refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet    = *refInOut / 2;
    *refInOut *=  2;
    
    return S_OK;
}

STDMETHODIMP AutomationTester::testSCODE(SCODE valIn, SCODE* refIn, SCODE* refInOut, SCODE* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testSCODE:");
    
    safe_swprintf(buffer, L"\tvalIn:\t%08x"   ,  valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%08x"   , *refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%08x", *refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet   = S_FALSE;
    *refInOut = E_ABORT;
    
    return S_OK;
}

STDMETHODIMP AutomationTester::testSHORT(SHORT valIn, SHORT* refIn, SHORT* refInOut, SHORT* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testSHORT:");
    
    safe_swprintf(buffer, L"\tvalIn:\t%d"   , (int) valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%d"   , (int)*refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%d", (int)*refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet    = *refInOut / 2;
    *refInOut *=  2;
    return S_OK;
}

STDMETHODIMP AutomationTester::testULONG(ULONG valIn, ULONG* refIn, ULONG* refInOut, ULONG* refRet)
{
    static_assert(sizeof(ULONG) == sizeof(int), "Fixup the format specifiers.");
    
    wchar_t buffer[0x1ff];
    TextAppend(L"testULONG:");
    safe_swprintf(buffer, L"\tvalIn:\t%u"   ,  valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%u"   , *refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%u", *refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet    = *refInOut / 2;
    *refInOut *=  2;
    return S_OK;
}

STDMETHODIMP AutomationTester::testULONGLONG(ULONGLONG valIn, ULONGLONG* refIn, ULONGLONG* refInOut, ULONGLONG* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testULONGLONG:");
    
    safe_swprintf(buffer, L"\tvalIn:\t"    PRINTF_UI64 L"",  valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t"    PRINTF_UI64 L"", *refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t" PRINTF_UI64 L"", *refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet    = *refInOut / 2;
    *refInOut *=  2;
    
    return S_OK;
}

STDMETHODIMP AutomationTester::testUSHORT(USHORT valIn, USHORT* refIn, USHORT* refInOut, USHORT* refRet)
{
    static_assert(sizeof(LONG) == sizeof(int), "Fixup the format specifiers.");
    
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testLONG:");
    
    safe_swprintf(buffer, L"\tvalIn:\t%d"   , (int) valIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%d"   , (int)*refIn   );
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%d", (int)*refInOut);
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet    = *refInOut / 2;
    *refInOut *=  2;
    return S_OK;
}

STDMETHODIMP AutomationTester::testVARIANTBOOL(VARIANT_BOOL valIn, VARIANT_BOOL* refIn, VARIANT_BOOL* refInOut, VARIANT_BOOL* refRet)
{
    wchar_t buffer[0x1ff];
    TextAppend(L"");
    TextAppend(L"testVARIANTBOOL:");
    
    safe_swprintf(buffer, L"\tvalIn:\t%s"   ,  valIn    ? L"TRUE" : L"FALSE");
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefIn:\t%s"   , *refIn    ? L"TRUE" : L"FALSE");
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    safe_swprintf(buffer, L"\trefInOut:\t%s", *refInOut ? L"TRUE" : L"FALSE");
    if (!Ensure(TextAppend(buffer))) return E_FAIL;
    
    if ((valIn != *refIn) &&
        (!Ensure(TextAppend(L"\tWARN: valIn != refIn")))) return E_FAIL;
    
    *refRet   = -1;
    *refInOut = *refInOut ? 0 : -1;
    return S_OK;
}

STDMETHODIMP AutomationTester::testDISPATCH(IDispatch* valIn, IDispatch** refIn, IDispatch** refInOut, IDispatch** refRet)
{
    *refRet = nullptr;
    TextAppend(L"");
    TextAppend(L"testOLECOLOR:");
    TextAppend(L"\tNOT IMPLEMENTED");
    return E_NOTIMPL;
}

STDMETHODIMP AutomationTester::testUNKNOWN(IUnknown* valIn, IUnknown** refIn, IUnknown** refInOut, IUnknown** refRet)
{
    *refRet = nullptr;
    TextAppend(L"");
    TextAppend(L"testOLECOLOR:");
    TextAppend(L"\tNOT IMPLEMENTED");
    return E_NOTIMPL;
}

STDMETHODIMP AutomationTester::testAryBSTR       (SAFEARRAY/*(BSTR        )*/* valIn, SAFEARRAY/*(BSTR        )*/** refIn, SAFEARRAY/*(BSTR        )*/** refInOut, SAFEARRAY/*(BSTR        )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryBYTE       (SAFEARRAY/*(BYTE        )*/* valIn, SAFEARRAY/*(BYTE        )*/** refIn, SAFEARRAY/*(BYTE        )*/** refInOut, SAFEARRAY/*(BYTE        )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryCHAR       (SAFEARRAY/*(CHAR        )*/* valIn, SAFEARRAY/*(CHAR        )*/** refIn, SAFEARRAY/*(CHAR        )*/** refInOut, SAFEARRAY/*(CHAR        )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryCY         (SAFEARRAY/*(CY          )*/* valIn, SAFEARRAY/*(CY          )*/** refIn, SAFEARRAY/*(CY          )*/** refInOut, SAFEARRAY/*(CY          )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryDATE       (SAFEARRAY/*(DATE        )*/* valIn, SAFEARRAY/*(DATE        )*/** refIn, SAFEARRAY/*(DATE        )*/** refInOut, SAFEARRAY/*(DATE        )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryDECIMAL    (SAFEARRAY/*(DECIMAL     )*/* valIn, SAFEARRAY/*(DECIMAL     )*/** refIn, SAFEARRAY/*(DECIMAL     )*/** refInOut, SAFEARRAY/*(DECIMAL     )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryDOUBLE     (SAFEARRAY/*(DOUBLE      )*/* valIn, SAFEARRAY/*(DOUBLE      )*/** refIn, SAFEARRAY/*(DOUBLE      )*/** refInOut, SAFEARRAY/*(DOUBLE      )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryFLOAT      (SAFEARRAY/*(FLOAT       )*/* valIn, SAFEARRAY/*(FLOAT       )*/** refIn, SAFEARRAY/*(FLOAT       )*/** refInOut, SAFEARRAY/*(FLOAT       )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryLONG       (SAFEARRAY/*(LONG        )*/* valIn, SAFEARRAY/*(LONG        )*/** refIn, SAFEARRAY/*(LONG        )*/** refInOut, SAFEARRAY/*(LONG        )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryLONGLONG   (SAFEARRAY/*(LONGLONG    )*/* valIn, SAFEARRAY/*(LONGLONG    )*/** refIn, SAFEARRAY/*(LONGLONG    )*/** refInOut, SAFEARRAY/*(LONGLONG    )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testArySCODE      (SAFEARRAY/*(SCODE       )*/* valIn, SAFEARRAY/*(SCODE       )*/** refIn, SAFEARRAY/*(SCODE       )*/** refInOut, SAFEARRAY/*(SCODE       )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testArySHORT      (SAFEARRAY/*(SHORT       )*/* valIn, SAFEARRAY/*(SHORT       )*/** refIn, SAFEARRAY/*(SHORT       )*/** refInOut, SAFEARRAY/*(SHORT       )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryULONG      (SAFEARRAY/*(ULONG       )*/* valIn, SAFEARRAY/*(ULONG       )*/** refIn, SAFEARRAY/*(ULONG       )*/** refInOut, SAFEARRAY/*(ULONG       )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryULONGLONG  (SAFEARRAY/*(ULONGLONG   )*/* valIn, SAFEARRAY/*(ULONGLONG   )*/** refIn, SAFEARRAY/*(ULONGLONG   )*/** refInOut, SAFEARRAY/*(ULONGLONG   )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryUSHORT     (SAFEARRAY/*(USHORT      )*/* valIn, SAFEARRAY/*(USHORT      )*/** refIn, SAFEARRAY/*(USHORT      )*/** refInOut, SAFEARRAY/*(USHORT      )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryVARIANTBOOL(SAFEARRAY/*(VARIANT_BOOL)*/* valIn, SAFEARRAY/*(VARIANT_BOOL)*/** refIn, SAFEARRAY/*(VARIANT_BOOL)*/** refInOut, SAFEARRAY/*(VARIANT_BOOL)*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryDISPATCH   (SAFEARRAY/*(IDispatch*  )*/* valIn, SAFEARRAY/*(IDispatch*  )*/** refIn, SAFEARRAY/*(IDispatch*  )*/** refInOut, SAFEARRAY/*(IDispatch*  )*/** refRet) { return E_NOTIMPL; }
STDMETHODIMP AutomationTester::testAryUNKNOWN    (SAFEARRAY/*(IUnknown*   )*/* valIn, SAFEARRAY/*(IUnknown*   )*/** refIn, SAFEARRAY/*(IUnknown*   )*/** refInOut, SAFEARRAY/*(IUnknown*   )*/** refRet) { return E_NOTIMPL; }


// Util method for outputting debug/notification text
HRESULT AutomationTester::TextAppend(wchar_t const* const szText)
{
    BSTR bstrText = nullptr;
    if (!Ensure(get_Text(&bstrText))) return E_FAIL;
    
    _bstr_t bstrTextWrap(bstrText, false); //< Attach, else we leak!
    bstrTextWrap += szText;
    bstrTextWrap += L"\r\n";
    if (!Ensure(put_Text(bstrTextWrap))) return E_FAIL;
    
    return S_FALSE;
}

