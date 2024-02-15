#pragma once

#ifndef UNICODE
#define UNICODE
#endif

#include <combaseapi.h>
#include <UIAutomation.h>
#include <Windows.h>

#include <string>
#include <iostream>
#include <mutex>
#include <condition_variable>
#include <thread>

namespace utils
{
    static const std::wstring httpsPart = L"https:";
    static const std::wstring qEqualPart = L"q=";
    static const std::wstring textToAppend = L"test:";
    static const std::wstring textToAppendMod = L"test%3";

    static const std::wstring browserWindowNameEdgeAndChrome = L"Chrome_WidgetWin_1";
    static const std::wstring browserWindowNameFirefox = L"MozillaWindowClass";

    struct SyncBlock
    {
        std::mutex mx;
        std::condition_variable cv;
        bool processed{ false };
    };

    std::wstring BstrToWstring(BSTR bstr)
    {
        if (!bstr)
            return {};

        auto len = SysStringLen(bstr);
        return std::wstring(bstr, len);
    }

    void PrintCurrentName(IUIAutomationElement* elem)
    {
        if (!elem)
            return;

        BSTR nameW;
        if (auto h = elem->get_CurrentName(&nameW); FAILED(h))
            std::wcerr << L"Failed to get CurrentName property" << std::endl;

        std::wcout << L"CurrentName property: " << utils::BstrToWstring(nameW) << std::endl;
        SysFreeString(nameW);
    }

    template <typename T>
    class UiaPtrWrapper
    {
    public:
        UiaPtrWrapper() : ptr{ nullptr } {}

        UiaPtrWrapper(T* p) : ptr{ p } {}

        ~UiaPtrWrapper()
        {
            if (ptr)
            {
                ptr->Release();
                ptr = nullptr;
            }
        }

        UiaPtrWrapper(UiaPtrWrapper&& other) noexcept
            : ptr{ other.ptr }
        {
            other.ptr = nullptr;
        }

        UiaPtrWrapper& operator=(UiaPtrWrapper&& other) noexcept
        {
            if (this != &other)
            {
                if (ptr)
                    ptr->Release();

                ptr = other.ptr;
                other.ptr = nullptr;
            }
            return *this;
        }

        T*& get()
        {
            return ptr;
        }

        T* operator->() const
        {
            return ptr;
        }

        explicit operator bool() const
        {
            return ptr != nullptr;
        }

    private:
        UiaPtrWrapper(const UiaPtrWrapper&) = delete;
        UiaPtrWrapper& operator=(const UiaPtrWrapper&) = delete;

        T* ptr;
    };

    class VariantWrapper
    {
    public:
        VariantWrapper() = default;

        ~VariantWrapper()
        {
            VariantClear(&var);
        }

        VARIANT& get() { return var; }

    private:
        VARIANT var;
    };
}
