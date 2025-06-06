#ifndef EXCEL_HPP
#define EXCEL_HPP
#include "Outlook.hpp"

IDispatch *get_excel_app()
{
    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(constr_to_wc("Excel.Application"), &clsid);
    if (FAILED(hr))
    {
        wcout << constr_to_wc("Cannot find Excel.Application") << endl;
        return nullptr;
    }

    IDispatch *pExcel = nullptr;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pExcel);
    if (FAILED(hr))
    {
        wcout << constr_to_wc("Cannot create Excel instance") << endl;
        return nullptr;
    }

    DISPID dispID;
    OLECHAR *name = constr_to_wc("Visible");
    pExcel->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispID);
    VARIANT x;
    x.vt = VT_BOOL;
    x.boolVal = VARIANT_FALSE;
    DISPID dispidNamed = DISPID_PROPERTYPUT;
    DISPPARAMS dp = {&x, &dispidNamed, 1, 1};
    pExcel->Invoke(dispID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);

    return pExcel;
}

IDispatch *get_workbooks(IDispatch *pExcel)
{
    DISPID dispidWorkbooks;
    OLECHAR *name = constr_to_wc("Workbooks");
    HRESULT hr = pExcel->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispidWorkbooks);
    if (FAILED(hr))
        return nullptr;

    DISPPARAMS noArgs = {nullptr, nullptr, 0, 0};
    VARIANT result;
    VariantInit(&result);
    hr = pExcel->Invoke(dispidWorkbooks, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, nullptr, nullptr);
    if (FAILED(hr))
        return nullptr;

    return result.pdispVal;
}

IDispatch *get_workbook(IDispatch *pWorkbooks, string filepath)
{
    if (!pWorkbooks)
        return nullptr;

    DISPID dispidOpen;
    OLECHAR *openName = constr_to_wc("Open");
    HRESULT hr = pWorkbooks->GetIDsOfNames(IID_NULL, &openName, 1, LOCALE_USER_DEFAULT, &dispidOpen);
    if (FAILED(hr))
        return nullptr;

    VARIANT vtFilePath;
    vtFilePath.vt = VT_BSTR;
    vtFilePath.bstrVal = SysAllocString(string_to_widechar(filepath).c_str());

    DISPPARAMS params = {&vtFilePath, nullptr, 1, 0};
    VARIANT result;
    VariantInit(&result);

    hr = pWorkbooks->Invoke(dispidOpen, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, nullptr, nullptr);

    SysFreeString(vtFilePath.bstrVal);

    if (FAILED(hr))
        return nullptr;

    return result.pdispVal;
}

IDispatch *get_sheet(IDispatch *pSheets, int index)
{
    if (!pSheets)
        return nullptr;

    DISPID dispidItem;
    OLECHAR *itemName = constr_to_wc("Item");
    HRESULT hr = pSheets->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT, &dispidItem);
    if (FAILED(hr))
        return nullptr;

    VARIANT varIndex;
    varIndex.vt = VT_I4;
    varIndex.lVal = index;

    DISPPARAMS params = {&varIndex, nullptr, 1, 0};
    VARIANT result;
    VariantInit(&result);

    hr = pSheets->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, nullptr, nullptr);
    if (FAILED(hr))
        return nullptr;

    return result.pdispVal;
}
#endif
