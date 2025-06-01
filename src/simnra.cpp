#include "simnra.h"


VARIANT MakeIntArg(int val) {
    VARIANT v; VariantInit(&v);
    v.vt = VT_I4; v.intVal = val;
    return v;
}

VARIANT MakeDoubleArg(double val) {
    VARIANT v; VariantInit(&v);
    v.vt = VT_R8; v.dblVal = val;
    return v;
}

VARIANT MakeStringArg(const std::wstring& val) {
    VARIANT v; VariantInit(&v);
    v.vt = VT_BSTR; v.bstrVal = SysAllocString(val.c_str());
    return v;
}

void check_hresult(HRESULT hr, const std::string& msg) {
    if (FAILED(hr)) {
        _com_error err(hr);
        std::wcerr << L"Error: " << std::wstring(msg.begin(), msg.end()) << L" - " << err.ErrorMessage() << std::endl;
        exit(1);
    }
}

void InvokeVoidMethod(IDispatch* disp, const wchar_t* method, const std::vector<VARIANT>& args ) {
    DISPID dispid;
    LPOLESTR nameCopy = const_cast<LPOLESTR>(method);
    HRESULT hr = disp->GetIDsOfNames(IID_NULL, &nameCopy, 1, LOCALE_USER_DEFAULT, &dispid);
    check_hresult(hr, "GetIDsOfNames failed");
    
    std::vector<VARIANT> args_reverse(args.rbegin(), args.rend());
    DISPPARAMS params = {
        args_reverse.data(),
        nullptr,
        static_cast<UINT>(args.size()),
        0
    };

    hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, nullptr, nullptr, nullptr);
    check_hresult(hr, "InvokeVoidMethod failed");
}


VARIANT GetPropertyValue(IDispatch* disp, const wchar_t* propertyName, const std::vector<VARIANT>& args) {
    if (!disp) throw std::runtime_error("IDispatch is null");

    DISPID dispid;
    LPOLESTR name = const_cast<LPOLESTR>(propertyName);
    HRESULT hr = disp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    check_hresult(hr, "GetIDsOfNames failed");
    
    std::vector<VARIANT> args_reverse (args.rbegin(), args.rend());
    // Prepare DISPPARAMS
    DISPPARAMS dp;
    dp.cArgs = static_cast<UINT>(args_reverse.size());
    dp.rgvarg = const_cast<VARIANT*>(args_reverse.data());  // right-to-left
    dp.cNamedArgs = 0;
    dp.rgdispidNamedArgs = nullptr;

    VARIANT result;
    VariantInit(&result);

    hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
    check_hresult(hr, "Invoke PROPERTYGET failed");

    return result;  // caller owns and must VariantClear
}

void SetPropertyValue(IDispatch* disp, const wchar_t* propertyName, const std::vector<VARIANT>& args) {
    if (!disp) throw std::runtime_error("IDispatch is null");
    if (args.empty()) throw std::runtime_error("Setter requires at least one argument (value to set)");

    DISPID dispid;
    LPOLESTR name = const_cast<LPOLESTR>(propertyName);
    HRESULT hr = disp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    check_hresult(hr, "GetIDsOfNames failed");

    DISPID namedArg = DISPID_PROPERTYPUT;

    std::vector<VARIANT> args_reverse (args.rbegin(), args.rend());

    DISPPARAMS dp;
    dp.cArgs = static_cast<UINT>(args_reverse.size());
    dp.rgvarg = const_cast<VARIANT*>(args_reverse.data());
    dp.cNamedArgs = 1;
    dp.rgdispidNamedArgs = &namedArg;

    hr = disp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, nullptr, nullptr, nullptr);
    check_hresult(hr, "Invoke PROPERTYPUT failed");
}


IDispatch* CreateDispatch(const wchar_t* progID) {
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(progID, &clsid);
    if (FAILED(hr)) throw std::runtime_error("CLSIDFromProgID failed for " + std::string(progID, progID + wcslen(progID)));

    IDispatch* disp = nullptr;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&disp);
    if (FAILED(hr) || !disp) throw std::runtime_error("CoCreateInstance failed for " + std::string(progID, progID + wcslen(progID)));
    return disp;
}

