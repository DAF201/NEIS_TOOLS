#ifndef HTTP_HPP
#define HTTP_HPP
#include <windows.h>
#include <wininet.h>
#include <iostream>
#include <cstring>
#include <string>
#pragma comment(lib, "wininet.lib")
using namespace std;
namespace http
{

    string http_get(string url)
    {
        string result = "";

        HINTERNET hInternet = InternetOpenA("WinINetExample", INTERNET_OPEN_TYPE_PRECONFIG, NULL, NULL, 0);
        if (!hInternet)
        {
            std::cerr << "InternetOpen failed\n";
            return "";
        }

        HINTERNET hUrl = InternetOpenUrlA(hInternet, url.c_str(), NULL, 0, INTERNET_FLAG_RELOAD, 0);
        if (!hUrl)
        {
            std::cerr << "InternetOpenUrl failed\n";
            InternetCloseHandle(hInternet);
            return "";
        }

        char buffer[4096];
        DWORD bytesRead;
        while (InternetReadFile(hUrl, buffer, sizeof(buffer) - 1, &bytesRead) && bytesRead > 0)
        {
            buffer[bytesRead] = '\0';
            result += buffer;
        }

        InternetCloseHandle(hUrl);
        InternetCloseHandle(hInternet);
        return result;
    }

    string http_post(string url, string &formData)
    {
        string result;
        HINTERNET hInternet = InternetOpenA("", INTERNET_OPEN_TYPE_PRECONFIG, NULL, NULL, 0);
        if (!hInternet)
            return "";

        URL_COMPONENTSA uc = {sizeof(uc)};
        char host[256], path[1024];
        uc.lpszHostName = host;
        uc.dwHostNameLength = sizeof(host);
        uc.lpszUrlPath = path;
        uc.dwUrlPathLength = sizeof(path);

        if (!InternetCrackUrlA(url.c_str(), 0, 0, &uc))
        {
            InternetCloseHandle(hInternet);
            return "";
        }

        HINTERNET hConnect = InternetConnectA(hInternet, uc.lpszHostName, uc.nPort, NULL, NULL, INTERNET_SERVICE_HTTP, 0, 0);
        if (!hConnect)
        {
            InternetCloseHandle(hInternet);
            return "";
        }

        const char *acceptTypes[] = {"*/*", NULL};
        HINTERNET hRequest = HttpOpenRequestA(hConnect, "POST", uc.lpszUrlPath, NULL,
                                              NULL, acceptTypes, INTERNET_FLAG_RELOAD | INTERNET_FLAG_NO_CACHE_WRITE, 0);
        if (!hRequest)
        {
            InternetCloseHandle(hConnect);
            InternetCloseHandle(hInternet);
            return "";
        }

        string headers = "Content-Type: application/x-www-form-urlencoded\r\n";

        if (!HttpSendRequestA(hRequest, headers.c_str(), headers.length(),
                              (LPVOID)formData.c_str(), formData.length()))
        {
            InternetCloseHandle(hRequest);
            InternetCloseHandle(hConnect);
            InternetCloseHandle(hInternet);
            return "";
        }

        char buffer[4096];
        DWORD bytesRead;
        while (InternetReadFile(hRequest, buffer, sizeof(buffer) - 1, &bytesRead) && bytesRead > 0)
        {
            buffer[bytesRead] = '\0';
            result += buffer;
        }

        InternetCloseHandle(hRequest);
        InternetCloseHandle(hConnect);
        InternetCloseHandle(hInternet);

        return result;
    }
}
#endif
/*
#include "http.hpp"
#include <iostream>
using namespace std;
int main()
{
    cout << http::http_get("http://example.com") << endl;
    std::string data = "username=user&password=password";
    std::string response = http::http_post("http://httpbin.org/post", data);
    std::cout << "Server Response:\n"
              << response << std::endl;
}
*/