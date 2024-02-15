#pragma once
#include "Utils.h"

namespace uia
{
    using UIElemPtr = utils::UiaPtrWrapper<IUIAutomationElement>;
    using UIAutoPtr = utils::UiaPtrWrapper<IUIAutomation>;
    using UICondPtr = utils::UiaPtrWrapper<IUIAutomationCondition>;
    using UIValPattPtr = utils::UiaPtrWrapper<IUIAutomationValuePattern>;
    using UIElemArrayPtr = utils::UiaPtrWrapper<IUIAutomationElementArray>;

    // Searches first of or all Edit Controls, where user puts URL
    class BrowserUrlFinder
    {
    public:
        BrowserUrlFinder() {};

        UIElemPtr findUrl(UIAutoPtr& uiAuto, UIElemPtr& rootElem)
        {
            if (!uiAuto || !rootElem)
                return nullptr;

            if (!conditionForUrl && !prepareConditionForUrl(uiAuto))
                return nullptr;

            UIElemPtr url;
            if (auto h = rootElem->FindFirst(TreeScope_Descendants, conditionForUrl.get(), &(url.get())); FAILED(h) || !url)
                return nullptr;

            //utils::PrintCurrentName(url.get());

            return url;
        }

        UIElemArrayPtr findAllBrowserWindowsOpened(UIAutoPtr& uiAuto, UIElemPtr& rootElem)
        {
            if (!uiAuto || !rootElem)
                return nullptr;

            if (!conditionBrowserCombined && !prepareBrowserConditions(uiAuto))
                return nullptr;

            UIElemArrayPtr elemArr;

            if (auto h = rootElem->FindAll(TreeScope_Children/*TreeScope_Descendants*/, conditionBrowserCombined.get(), &(elemArr.get())); FAILED(h) || !elemArr)
                return nullptr;

            return elemArr;
        }

    private:

        bool prepareConditionForUrl(UIAutoPtr& uiAuto)
        {
            utils::VariantWrapper varWrap;
            auto& var = varWrap.get();
            var.vt = VT_I4;
            var.lVal = UIA_EditControlTypeId;

            if (auto h = uiAuto->CreatePropertyCondition(UIA_ControlTypePropertyId, var, &(conditionForUrl.get())); FAILED(h))
                return false;

            return true;
        }

        bool prepareBrowserConditions(UIAutoPtr& uiAuto)
        {
            UICondPtr conditionForEditEdgeAndChrome, conditionForEditFirefox;

            if (conditionForEditEdgeAndChrome = prepareBrowserCondition(uiAuto, utils::browserWindowNameEdgeAndChrome); !conditionForEditEdgeAndChrome)
                return false;

            if (conditionForEditFirefox = prepareBrowserCondition(uiAuto, utils::browserWindowNameFirefox); !conditionForEditFirefox)
                return false;

            if (auto h = uiAuto->CreateOrCondition(
                conditionForEditEdgeAndChrome.get(),
                conditionForEditFirefox.get(),
                &(conditionBrowserCombined.get()))
                ; FAILED(h) || !conditionBrowserCombined)
            {
                return false;
            }

            return true;
        }

        static UICondPtr prepareBrowserCondition(UIAutoPtr& uiAuto, const std::wstring& value)
        {
            utils::VariantWrapper varWrap;
            auto& var = varWrap.get();
            var.vt = VT_BSTR;
            if (var.bstrVal = SysAllocString(value.data()); !var.bstrVal)
                return nullptr;

            UICondPtr condition;
            if (auto h = uiAuto->CreatePropertyCondition(UIA_ClassNamePropertyId, var, &(condition.get())); FAILED(h))
                return nullptr;

            return condition;
        }

    private:
        BrowserUrlFinder(const BrowserUrlFinder&) = delete;
        BrowserUrlFinder& operator=(const BrowserUrlFinder&) = delete;
        BrowserUrlFinder(BrowserUrlFinder&&) = delete;
        BrowserUrlFinder& operator=(BrowserUrlFinder&&) = delete;

        UICondPtr conditionBrowserCombined;
        UICondPtr conditionForUrl;
    };

    // Event handler for URL manipulation
    // 
    // Algorithm is the following:
    // 1. detect when search is performed (https + q= should be presented in a text)
    // 2. obtain value pattern to change text
    // 3. manipulate url
    // 4. simulate Enter key pressed
    // 
    // Important to note that If we see "test:" part presented in URL we'll do nothing
    class UrlEventHandler : public IUIAutomationEventHandler
    {
    public:
        UrlEventHandler() : refCount{ 1 }
        {
            prepareKbdInput();
        }

        // AddRef, Release and QueryInterface just put here as is from MS doc

        ULONG STDMETHODCALLTYPE AddRef()
        {
            ULONG ret = InterlockedIncrement(&refCount);
            return ret;
        }

        ULONG STDMETHODCALLTYPE Release()
        {
            ULONG ret = InterlockedDecrement(&refCount);
            if (ret == 0)
            {
                delete this;
                return 0;
            }
            return ret;
        }

        HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface)
        {
            if (riid == __uuidof(IUnknown))
                *ppInterface = static_cast<IUIAutomationEventHandler*>(this);
            else if (riid == __uuidof(IUIAutomationEventHandler))
                *ppInterface = static_cast<IUIAutomationEventHandler*>(this);
            else
            {
                *ppInterface = NULL;
                return E_NOINTERFACE;
            }
            this->AddRef();
            return S_OK;
        }

        // Main handler
        HRESULT STDMETHODCALLTYPE HandleAutomationEvent(IUIAutomationElement* pSender, EVENTID eventID)
        {
            switch (eventID)
            {
            case UIA_Text_TextChangedEventId:
            {
                HandleUrlNew(pSender);
                break;
            }
            default:
                //std::wcout << "Some Event Received: " << eventID << std::endl;
                break;
            }

            return S_OK;
        }

    private:
        void prepareKbdInput()
        {
            ZeroMemory(kbdInputs, sizeof(kbdInputs));

            kbdInputs[0].type = INPUT_KEYBOARD;
            kbdInputs[0].ki.wVk = VK_RETURN;

            kbdInputs[1].type = INPUT_KEYBOARD;
            kbdInputs[1].ki.wVk = VK_RETURN;
            kbdInputs[1].ki.dwFlags = KEYEVENTF_KEYUP;
        }

        bool HandleUrlNew(IUIAutomationElement* pSender)
        {
            if (!pSender)
                return false;

            utils::VariantWrapper var;
            pSender->GetCurrentPropertyValue(UIA_ValueValuePropertyId, &(var.get()));
            auto currUrl = utils::BstrToWstring(var.get().bstrVal);

            //std::wcout << "> Name: " << currUrl << std::endl;

            if (currUrl.find_first_of(utils::httpsPart) == currUrl.npos)
                return false;

            const auto queryStartPos = currUrl.find(utils::qEqualPart, utils::httpsPart.size());
            if (queryStartPos == currUrl.npos)
                return false;

            if (currUrl.find(utils::textToAppend, queryStartPos + 2) != currUrl.npos)
                return false;

            if (currUrl.find(utils::textToAppendMod, queryStartPos + 2) != currUrl.npos)
                return false;

            currUrl.insert(queryStartPos + 2, utils::textToAppend);
            //std::wcout << "Value To Set : " << currUrl << std::endl;

            IUnknown* patternObj;
            if (auto h = pSender->GetCurrentPattern(UIA_ValuePatternId, &patternObj); FAILED(h) || !patternObj)
            {
                //std::wcerr << "Failed to obtain Pattern Object" << std::endl;
                return false;
            }

            UIValPattPtr valuePattern(reinterpret_cast<IUIAutomationValuePattern*>(patternObj));
            if (!valuePattern)
            {
                //std::wcerr << "Failed to obtain Value Pattern" << std::endl;
                return false;
            }

            BSTR updatedUrlValue = SysAllocString(currUrl.data());
            if (auto h = valuePattern->SetValue(updatedUrlValue); FAILED(h))
            {
                //std::wcerr << "Failed to Set Value: " << currUrl << std::endl;
                SysFreeString(updatedUrlValue);
                return false;
            }

            SysFreeString(updatedUrlValue);

            // need to simulate Enter key pressed on a keyboard to perform a search
            // with modified URL string
            SendInput(ARRAYSIZE(kbdInputs), kbdInputs, sizeof(INPUT));

            return true;
        }

        LONG refCount;
        INPUT kbdInputs[2]; // KEYDOWN + KEYUP
    };

    using UrlEventHPtr = utils::UiaPtrWrapper<UrlEventHandler>;

    // Event handler for detecting new browser windows opened
    //
    // Algorithm:
    // 1. detect only specific class names corresponding to Edge, Firefox and Chrome
    // 2. signal to the main thread about successfull detection
    class BrowserWindowEventHandler : public IUIAutomationEventHandler
    {
    public:
        BrowserWindowEventHandler(utils::SyncBlock& sBlock)
            : refCount{ 1 }, synBlock{sBlock}
        {}

        ULONG STDMETHODCALLTYPE AddRef()
        {
            ULONG ret = InterlockedIncrement(&refCount);
            return ret;
        }

        ULONG STDMETHODCALLTYPE Release()
        {
            ULONG ret = InterlockedDecrement(&refCount);
            if (ret == 0)
            {
                delete this;
                return 0;
            }
            return ret;
        }

        HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppInterface)
        {
            if (riid == __uuidof(IUnknown))
                *ppInterface = static_cast<IUIAutomationEventHandler*>(this);
            else if (riid == __uuidof(IUIAutomationEventHandler))
                *ppInterface = static_cast<IUIAutomationEventHandler*>(this);
            else
            {
                *ppInterface = NULL;
                return E_NOINTERFACE;
            }
            this->AddRef();
            return S_OK;
        }

        HRESULT STDMETHODCALLTYPE HandleAutomationEvent(IUIAutomationElement* pSender, EVENTID eventID)
        {
            switch (eventID)
            {
            case UIA_Window_WindowOpenedEventId:
            {
                utils::VariantWrapper varStr;
                pSender->get_CurrentClassName(&(varStr.get().bstrVal));

                auto windowClass = utils::BstrToWstring(varStr.get().bstrVal);

                if (windowClass.find(utils::browserWindowNameEdgeAndChrome) != windowClass.npos ||
                    windowClass.find(utils::browserWindowNameFirefox) != windowClass.npos)
                {
                    //std::wcout << "> New Browser Window Opened" << std::endl;
                    
                    // need to signal to the main thread about new window opened
                    {
                        std::lock_guard lk(synBlock.mx);
                        synBlock.processed = true;
                    }
                    synBlock.cv.notify_one();
                }
                break;
            }
            default:
                //std::wcout << "Some Event Received: " << eventID << std::endl;
                break;
            }

            return S_OK;
        }

    private:
        LONG refCount;

        utils::SyncBlock& synBlock;
    };

    using BrowserEventHPtr = utils::UiaPtrWrapper<BrowserWindowEventHandler>;

    bool AddUrlHandler(UIAutoPtr& ui, UIElemPtr& urlElem, UrlEventHPtr& urlHandler)
    {
        auto h = ui->AddAutomationEventHandler(
            UIA_Text_TextChangedEventId,
            urlElem.get(),
            TreeScope_Element,
            nullptr,
            reinterpret_cast<IUIAutomationEventHandler*>(urlHandler.get()));
            
        return FAILED(h) ? false : true;
    }

    bool AddBrowserWindowHandler(UIAutoPtr& ui, UIElemPtr& winElem, BrowserEventHPtr& winHandler)
    {
        auto h = ui->AddAutomationEventHandler(
            UIA_Window_WindowOpenedEventId,
            winElem.get(),
            TreeScope_Children,
            nullptr,
            reinterpret_cast<IUIAutomationEventHandler*>(winHandler.get()));

        return FAILED(h) ? false : true;
    }

    // Manages initialization specific to the UI Automation and provides ability
    // to add new URL manipulation handlers
    class UIManager
    {
    public:
        UIManager() = default;

        ~UIManager()
        {
            ui->RemoveAllEventHandlers();
            CoUninitialize();
        }

        bool init(utils::SyncBlock& sBlock)
        {
            if (auto h = CoInitializeEx(nullptr, COINIT_MULTITHREADED); h != S_OK && h != S_FALSE)
            {
                std::wcerr << "Failed to initialize COM" << std::endl;
                return false;
            }

            if (auto h = CoCreateInstance(CLSID_CUIAutomation, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&(ui.get())))
                ; FAILED(h))
            {
                std::wcerr << "Failed to initialized UIAutomation COM object" << std::endl;
                return false;
            }

            if (auto h = ui->GetRootElement(&(rootElem.get())); FAILED(h) || !rootElem)
            {
                std::wcout << "Failed to get Root Element" << std::endl;
                return false;
            }

            //utils::PrintCurrentName(rootElem.get());

            // allocate event handlers
            urlHandler = UrlEventHPtr(new UrlEventHandler());
            browserHandler = BrowserEventHPtr(new BrowserWindowEventHandler(sBlock));

            // detect all currently opened browser windows and add URL manipulators to them
            auto urlArray = urlReader.findAllBrowserWindowsOpened(ui, rootElem);
            if (urlArray)
            {
                int numOfBrowserWindows{ 0 };
                if (auto h = urlArray->get_Length(&numOfBrowserWindows); SUCCEEDED(h))
                {
                    std::cout << "Number of opened browser windows found: " << numOfBrowserWindows << std::endl;
                    for (int id{ 0 }; id < numOfBrowserWindows; id++)
                    {
                        UIElemPtr newBrowserWindow;
                        urlArray->GetElement(id, &(newBrowserWindow.get()));
                        if (newBrowserWindow)
                            tryAddNewHandler(newBrowserWindow);
                        //uia::AddUrlHandler(ui, newUrlElem, urlHandler);
                    }
                }
            }

            if (!browserHandler || !uia::AddBrowserWindowHandler(ui, rootElem, browserHandler))
            {
                std::wcout << "Failed to add window handler" << std::endl;
                return false;
            }

            return true;
        }

        // tries to find new Edit Control and to add corresponding event handler
        bool tryAddNewHandler(UIElemPtr& root)
        {
            auto newUrlElem = urlReader.findUrl(ui, root);

            if (!newUrlElem)
                return false;

            if (!uia::AddUrlHandler(ui, newUrlElem, urlHandler))
                return false;

            return true;
        }
        
        bool tryAddNewHandler()
        {
            return tryAddNewHandler(rootElem);
        }

    private:
        UIElemPtr rootElem;
        BrowserUrlFinder urlReader;
        UrlEventHPtr urlHandler;
        BrowserEventHPtr browserHandler;
        UIAutoPtr ui;
    };
}
