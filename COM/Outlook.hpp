#ifndef OUTLOOK_HPP
#define OUTLOOK_HPP
#include <initguid.h>
DEFINE_GUID(GUID_NULL, 0x00000000, 0x0000, 0x0000,
            0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00);
#include <windows.h>
#include <oleauto.h>
#include <iostream>
#include <locale>
#include <codecvt>
using namespace std;

static wstring string_to_widechar(string str)
{
    wstring_convert<codecvt_utf8_utf16<wchar_t>> converter;
    return converter.from_bytes(str);
}

#define constr_to_wc(x) ((wchar_t *)string_to_widechar(x).c_str())

int outlook_send(string target_address, string cc_address, string email_subject, string email_body, string attachment)
{

    CLSID clsid;

    HRESULT hr = CLSIDFromProgID(L"Outlook.Application", &clsid);
    if (FAILED(hr))
    {
        cout << "Failed to find outlook" << endl;
        return -1;
    }

    IDispatch *pOutlook = nullptr;
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pOutlook);
    if (FAILED(hr))
    {
        std::cout << "Failed to create outlook instance" << std::endl;
        return 1;
    }

    IDispatch *pSession = nullptr;
    {
        OLECHAR *name = constr_to_wc("Session");
        DISPID dispid;
        pOutlook->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
        DISPPARAMS noArgs = {NULL, NULL, 0, 0};
        VARIANT result;
        VariantInit(&result);
        pOutlook->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
        pSession = result.pdispVal;
    }

    {
        DISPID createID;
        OLECHAR *name = constr_to_wc("CreateItem");
        pOutlook->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &createID);

        VARIANT arg;
        arg.vt = VT_I4;
        arg.lVal = 0;

        DISPPARAMS params = {&arg, NULL, 1, 0};
        VARIANT result;
        VariantInit(&result);
        pOutlook->Invoke(createID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, NULL, NULL);
        IDispatch *pMail = result.pdispVal;

        DISPID dispidNamed = DISPID_PROPERTYPUT;

        // To
        DISPID toID;
        OLECHAR *toName = constr_to_wc("To");
        pMail->GetIDsOfNames(IID_NULL, &toName, 1, LOCALE_USER_DEFAULT, &toID);
        VARIANT to;
        to.vt = VT_BSTR;
        to.bstrVal = SysAllocString(string_to_widechar(target_address).c_str());
        DISPPARAMS dpTo = {&to, &dispidNamed, 1, 1};
        pMail->Invoke(toID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpTo, NULL, NULL, NULL);

        // Cc
        VARIANT cc;
        if (!cc_address.empty())
        {
            DISPID ccID;
            OLECHAR *ccName = constr_to_wc("Cc");
            pMail->GetIDsOfNames(IID_NULL, &ccName, 1, LOCALE_USER_DEFAULT, &ccID);
            cc.vt = VT_BSTR;
            cc.bstrVal = SysAllocString(string_to_widechar(cc_address).c_str());
            DISPPARAMS dpCc = {&cc, &dispidNamed, 1, 1};
            pMail->Invoke(ccID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpCc, NULL, NULL, NULL);
        }

        // Subject
        DISPID subjID;
        OLECHAR *subjName = constr_to_wc("Subject");
        pMail->GetIDsOfNames(IID_NULL, &subjName, 1, LOCALE_USER_DEFAULT, &subjID);
        VARIANT subj;
        subj.vt = VT_BSTR;
        subj.bstrVal = SysAllocString(string_to_widechar(email_subject).c_str());
        DISPPARAMS dpSubj = {&subj, &dispidNamed, 1, 1};
        pMail->Invoke(subjID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpSubj, NULL, NULL, NULL);

        // HTMLBody
        DISPID bodyID;
        OLECHAR *bodyName = constr_to_wc("HTMLBody");
        pMail->GetIDsOfNames(IID_NULL, &bodyName, 1, LOCALE_USER_DEFAULT, &bodyID);
        std::wstring wbody = string_to_widechar("<pre>" + email_body + "</pre>");
        VARIANT body;
        body.vt = VT_BSTR;
        body.bstrVal = SysAllocString(wbody.c_str());
        DISPPARAMS dpBody = {&body, &dispidNamed, 1, 1};
        hr = pMail->Invoke(bodyID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dpBody, NULL, NULL, NULL);
        if (FAILED(hr))
        {
            std::cout << "Failed to set HTMLBody, HRESULT=0x" << std::hex << hr << std::endl;
        }

        if (attachment.compare("") != 0)
        {
            DISPID attachID;
            OLECHAR *attachName = constr_to_wc("Attachments");
            pMail->GetIDsOfNames(IID_NULL, &attachName, 1, LOCALE_USER_DEFAULT, &attachID);

            DISPPARAMS noArgs = {NULL, NULL, 0, 0};
            VARIANT result;
            VariantInit(&result);
            pMail->Invoke(attachID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
            IDispatch *pAttachments = result.pdispVal;

            DISPID addID;
            OLECHAR *addName = constr_to_wc("Add");
            pAttachments->GetIDsOfNames(IID_NULL, &addName, 1, LOCALE_USER_DEFAULT, &addID);

            VARIANT arg;
            arg.vt = VT_BSTR;
            arg.bstrVal = SysAllocString(string_to_widechar(attachment).c_str());
            DISPPARAMS dpAdd = {&arg, NULL, 1, 0};
            VARIANT varResult;
            VariantInit(&varResult);
            HRESULT hr = pAttachments->Invoke(addID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpAdd, &varResult, NULL, NULL);
            if (FAILED(hr))
            {
                std::wcout << "Failed to attach file" << std::hex << hr << std::endl;
            }
        }

        // Send
        DISPID sendID;
        OLECHAR *sendName = constr_to_wc("Send");
        pMail->GetIDsOfNames(IID_NULL, &sendName, 1, LOCALE_USER_DEFAULT, &sendID);
        DISPPARAMS noArgs = {NULL, NULL, 0, 0};
        pMail->Invoke(sendID, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &noArgs, NULL, NULL, NULL);

        // Clean up
        SysFreeString(to.bstrVal);
        if (!cc_address.empty())
            SysFreeString(cc.bstrVal);
        SysFreeString(subj.bstrVal);
        SysFreeString(body.bstrVal);
        if (!attachment.empty())
            SysFreeString(arg.bstrVal);
        if (pMail)
            pMail->Release();
        if (pSession)
            pSession->Release();
    }

    if (pOutlook)
        pOutlook->Release();

    return 0;
}

#endif
