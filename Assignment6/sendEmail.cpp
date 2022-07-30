#include "stdafx.h"
#include <tchar.h>
#include <Windows.h>

#include "EASendMailObj.tlh"
using namespace EASendMailObjLib;

#include <atlbase.h>
#include <atlcom.h>

const int ConnectNormal = 0;
const int ConnectSSLAuto = 1;
const int ConnectSTARTTLS = 2;
const int ConnectDirectSSL = 3;
const int ConnectTryTLS = 4;

#define IDC_SRCFASTSENDER 1
static _ATL_FUNC_INFO OnSent = {CC_STDCALL, VT_EMPTY, 6,
    {VT_I4, VT_BSTR, VT_I4, VT_BSTR, VT_BSTR, VT_BSTR }};
class CFastSenderEvents:public IDispEventSimpleImpl<IDC_SRCFASTSENDER,
                                            CFastSenderEvents,
                                            &__uuidof(_IFastSenderEvents)>
{
public:
    CFastSenderEvents(){};
BEGIN_SINK_MAP(CFastSenderEvents)
    SINK_ENTRY_INFO(IDC_SRCFASTSENDER, __uuidof(_IFastSenderEvents), 1,
                &CFastSenderEvents::OnSentHandler, &OnSent)
END_SINK_MAP()
public:
    INT             m_nSent;
protected:
    HRESULT __stdcall OnSentHandler(long lRet,
        BSTR ErrorDesc,
        long nKey,
        BSTR tParam,
        BSTR Sender,
        BSTR Recipients)
    {
        _bstr_t rcpt = Recipients;
        if(lRet == 0)
        {
            _tprintf(_T("email was sent to %s successfully\r\n"), (const TCHAR*)rcpt);
        }
        else
        {
            _bstr_t error = ErrorDesc;
            _tprintf(_T("failed to sent email to %s with error %s \r\n"),
                (const TCHAR*)rcpt, (const TCHAR*)error);

        }
        m_nSent++;
        return S_OK;
    }
};

int _tmain(int argc, _TCHAR* argv[])
{
    ::CoInitialize(NULL);

    const INT nRcpt = 3;
    const TCHAR* arRcpt[nRcpt] = {
        _T("test@adminsystem.com"),
        _T("test1@adminsystem.com"),
        _T("test2@adminsystem.com")
    };
    IFastSenderPtr oFastSender = NULL;
    oFastSender.CreateInstance(__uuidof(EASendMailObjLib::FastSender));

    CFastSenderEvents oEvents;
    oEvents.DispEventAdvise(oFastSender.GetInterfacePtr());
    oEvents.m_nSent = 0;

    IMailPtr oSmtp = NULL;
    oSmtp.CreateInstance(__uuidof(EASendMailObjLib::Mail));
    oSmtp->LicenseCode = _T("TryIt");

    // Set your sender email address
    oSmtp->FromAddr = _T("test@emailarchitect.net");

    // Your SMTP server address
    oSmtp->ServerAddr = _T("smtp.emailarchitect.net");

    // User and password for ESMTP authentication, if your server doesn't
    // require User authentication, please remove the following codes.
    oSmtp->UserName = _T("test@emailarchitect.net");
    oSmtp->Password = _T("testpassword");

    // Most mordern SMTP servers require SSL/TLS connection now.
    // ConnectTryTLS means if server supports SSL/TLS, SSL/TLS will be used automatically.
    oSmtp->ConnectType = ConnectTryTLS;

    // If your SMTP server uses 587 port
    // oSmtp->ServerPort = 587;

    // If your SMTP server requires SSL/TLS connection on 25/587/465 port
    // oSmtp->ServerPort = 25; // 25 or 587 or 465
    // oSmtp->ConnectType = ConnectSSLAuto;

    for(int i = 0; i < nRcpt; i++)
    {
        oSmtp->ClearRecipient();
        oSmtp->AddRecipientEx(arRcpt[i], 0);
        oSmtp->Subject = _T("test mass email from visual c++");
        oSmtp->BodyText = _T("mass email test body sent from Visual C++");
        //submit email to inner thread pool
        oFastSender->Send(oSmtp, i, _T("anything"));
    }

    // Waiting for all email finished.
    while(oEvents.m_nSent < nRcpt)
    {
        MSG msg;
        while(PeekMessage(&msg, NULL, 0, 0, PM_REMOVE))
        {
            if(msg.message == WM_QUIT)
                return 0;
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
        ::Sleep(10);
    }
    oEvents.DispEventUnadvise(oFastSender.GetInterfacePtr());

    return 0;
}