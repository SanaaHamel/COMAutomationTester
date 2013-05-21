
local iBaseID = 1

local idlProp       = "[propget, id(%d)] HRESULT Prop|([out, retval] ^* pVal  );\n" ..
                      "[propput, id(%d)] HRESULT Prop|([in]          ^  newVal);"
local idlPropAry    = "[propget, id(%d)] HRESULT PropAry|([out, retval] SAFEARRAY(^)* pVal  );\n" ..
                      "[propput, id(%d)] HRESULT PropAry|([in]          SAFEARRAY(^)  newVal);"
local idlFunc       = "[id(%d)] HRESULT test|(^ valIn, ^* refIn, [in,out] ^* refInOut, [out,retval] ^* refRet);"
local idlFuncAry    = "[id(%d)] HRESULT testAry|(SAFEARRAY(^) valIn, SAFEARRAY(^)* refIn, [in,out] SAFEARRAY(^)* refInOut, [out,retval] SAFEARRAY(^)* refRet);"
local propDecl      = "STDMETHOD(get_Prop|)(^* pVal  );\n" ..
                      "STDMETHOD(put_Prop|)(^  newVal);"
local propAryDecl   = "STDMETHOD(get_PropAry|)(SAFEARRAY/*(^)*/** pVal  );\n" ..
                      "STDMETHOD(put_PropAry|)(SAFEARRAY/*(^)*/*  newVal);"
local funcDecl      = "STDMETHOD(test|)(^ valIn, ^* refIn, ^* refInOut, ^* refRet);"
local funcAryDecl   = "STDMETHOD(testAry|)(SAFEARRAY/*(^)*/* valIn, SAFEARRAY/*(^)*/** refIn, SAFEARRAY/*(^)*/** refInOut, SAFEARRAY/*(^)*/** refRet);"
local propMember    = "^ _prop|;"
local propAryMember = "SAFEARRAY/*(^)*/* _propAry|;"
local propImplGetSimple = "STDMETHODIMP AutomationTester::get_Prop|(^* pVal) { *pVal = _prop|; return S_OK; }"
local propImplGetObject = "STDMETHODIMP AutomationTester::get_Prop|(^* pVal) { if (_prop|) _prop|->AddRef(); *pVal = _prop|; return S_OK; }"
local propImplGetArray  = "STDMETHODIMP AutomationTester::get_PropAry|(SAFEARRAY/*(^)*/** pVal  ) { return PropArrayGet(_propAry|, $, pVal); }"
local propImplSetSimple = "STDMETHODIMP AutomationTester::put_Prop|(^ newVal) { _prop| = newVal; return S_OK; }"
local propImplSetObject = "STDMETHODIMP AutomationTester::put_Prop|(^ newVal)\n" ..
                          "{ if (_prop|) _prop|->Release();\n" ..
                          "  if (newVal) newVal->AddRef();\n"  ..
                          "  _prop| = newVal; return S_OK; }"
local propImplSetArray  = "STDMETHODIMP AutomationTester::put_PropAry|(SAFEARRAY/*(^)*/*  newVal) { return PropArrayPut(_propAry|, $, newVal); }"
local funcImpl       = "STDMETHODIMP AutomationTester::test|(^ valIn, ^* refIn, ^* refInOut, ^* refRet) { return E_NOTIMPL; }"
local funcAryImpl    = "STDMETHODIMP AutomationTester::testAry|(SAFEARRAY/*(^)*/* valIn, SAFEARRAY/*(^)*/** refIn, SAFEARRAY/*(^)*/** refInOut, SAFEARRAY/*(^)*/** refRet) { return E_NOTIMPL; }"
local propMemInitNum = "                                     , _prop|(0)"
local propMemInitPtr = "                                     , _prop|(nullptr)"
local propMemInitAry = "                                     , _propAry|(nullptr)"
local propMemDtorPtr = "if (_prop|) _prop|->Release(); _prop| = nullptr;"
local propMemDtorAry = "if (_propAry|) Ensure(SafeArrayDestroy(_propAry|)); _propAry| = nullptr;"


local testCaseDBase = [[
   function test|
      varInOriginal    = 25
      varInOutOriginal = 32
      valInOut  = varInOutOriginal
      valReturn = -1
      ? "| PreInOut:"
      ? valInOut
      ? "| PreReturn:"
      ? valReturn
      valReturn = this.ACTIVEX1.nativeObject.test|(varInOriginal, varInOriginal, valInOut)
      ? "| PostInOut:"
      ? valInOut
      ? "| PostReturn:"
      ? valReturn
      ? "| InOutDiff:"
      ? valInOut <> varInOutOriginal
      return]]
local testCaseDBaseInvoke = "test|()"


local typesDispatch = {"IDispatch*", "IUnknown*"}
local types = {"BSTR", "BYTE", "CHAR", "CY", "DATE", "DECIMAL", "DOUBLE", "FLOAT", "LONG", "LONGLONG"
              ,"SCODE", "SHORT", "ULONG", "ULONGLONG", "USHORT", "VARIANT_BOOL"}
local typesVT = {"VT_BSTR", "VT_UI1", "VT_I1", "VT_CY", "VT_DATE", "VT_DECIMAL", "VT_R8", "VT_R4"
                ,"VT_I4", "VT_I8", "VT_ERROR", "VT_I2", "VT_UI4", "VT_UI8", "VT_UI2", "VT_BOOL"
                ,"VT_DISPATCH", "VT_UNKNOWN"}
for _, t in ipairs(typesDispatch) do types[#types + 1] = t end
local typeprefix = {}
for k, v in pairs(types) do typeprefix[k] = v:gsub("[_%*]" , "")
                                             :gsub("^I", ""):upper() end
local function ApplyPadding(t)
    local n = 0
    for _, s in ipairs(t) do n    = math.max(n, #s)  end
    for _, s in ipairs(t) do t[_] = s .. (" "):rep(n - #s) end
end
--ApplyPadding(types)
--ApplyPadding(typeprefix)
--ApplyPadding(typesVT)

for k, v in ipairs(typesDispatch) do typesDispatch[v] = k    end -- Convert dispatch array to map
for k, v in ipairs(types        ) do typesVT[v] = typesVT[k] end -- Convert typesVT array to map


local function FmtDefault(iID, strType, strTypePrefix, strFormat)
    local s = strFormat:format(iID, iID):gsub("|" , strTypePrefix)
                                        :gsub("%^", strType)
    return s
end

local function FmtDispatch(a, b, id, t, p)
    return FmtDefault(id, t, p, typesDispatch[t:gsub(" ", "")] and a or b)
end

local function FmtPointer(a, b, id, t, p)
    return FmtDefault(id, t, p, (typesDispatch[t:gsub(" ", "")] or (t:gsub(" ", "") == "BSTR")) and a or b)
end

local function FmtSafeArray(a, id, t, p)
    local vt = typesVT[t]
    a = a:gsub("%$", vt)
    return FmtDefault(id, t, p, a)
end

local tblFormats  = {idlProp --< IDL
                    ,function(...) return FmtSafeArray(idlPropAry, ...) end
                    ,idlFunc, idlFuncAry
                     --< Header/decl
                    ,propDecl, propAryDecl, funcDecl, funcAryDecl, propMember, propAryMember
                     --< Impl
                    ,function(...) return FmtDispatch(propImplGetObject, propImplGetSimple, ...) end
                    ,function(...) return FmtSafeArray(propImplGetArray, ...) end
                    ,function(...) return FmtDispatch(propImplSetObject, propImplSetSimple, ...) end
                    ,function(...) return FmtSafeArray(propImplSetArray, ...) end
                    ,funcImpl
                    ,funcAryImpl
                    ,function(...) return FmtPointer (propMemInitPtr   , propMemInitNum   , ...) end
                    ,propMemInitAry
                    ,function(...) return FmtPointer  (propMemDtorPtr   , "", ...) end
                    ,function(...) return FmtSafeArray(propMemDtorAry   ,     ...) end}
tblFormats = {testCaseDBase, testCaseDBaseInvoke}
local iID = iBaseID
for _, fmt in pairs(tblFormats) do
    for k, t in pairs(types) do
        local fnFormat = (type(fmt) == "function") and fmt or FmtDefault
        print(fnFormat(iID, t, typeprefix[k], fmt))
        iID = iID + 1
    end
end


