// MIT License
//
// Copyright (c) 2025 WinApp - xmc0211 <xmc0211@qq.com>
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

#include "WinApp.h"
#include <unordered_map>
#include <array>
#include <strsafe.h>
#include <assert.h>
#include <functional>
#include <Psapi.h>
#include <functiondiscoverykeys.h>
#include <NotificationActivationCallback.h>
#include <propvarutil.h>
#include <MsXml6.h>
#include <comdef.h>
#include "Convert.h"


#pragma comment(lib, "msxml6.lib")

// Provide support for private API functions
namespace DllImporter {

	// Load a private API function from a library
	template <typename FTp>
	HRESULT LoadFunctionFromLibrary(std::_tstring LibraryName, std::_tstring FunctionName, FTp* Function) {
		HMODULE hModule = GetModuleHandle(LibraryName.c_str());
        SetLastError(ERROR_SUCCESS);
		BOOL bShouldLoad = FALSE;

		if (Function == NULL) return E_HANDLE;

		if (hModule == NULL) {
			bShouldLoad = TRUE;
			hModule = LoadLibrary(LibraryName.c_str());
			if (hModule == NULL) return E_INVALIDARG; // Library not exist
		}

		FTp pFunction = (FTp)GetProcAddress(hModule, LPT2LPC(FunctionName).c_str());
		if (pFunction == NULL) return E_INVALIDARG; // Function not exist

		*Function = pFunction;

		if (bShouldLoad) FreeLibrary(hModule);
		return S_OK;
	}
    
	// Private API function definations
	typedef HRESULT(FAR STDAPICALLTYPE* T_SetCurrentProcessExplicitAppUserModelID)(
		PCWSTR AppID
		);
	typedef HRESULT(FAR STDAPICALLTYPE* T_PropVariantToString)(
		REFPROPVARIANT propvar,
		PWSTR psz,
		UINT cch
		);
	typedef HRESULT(FAR STDAPICALLTYPE* T_RoGetActivationFactory)(
		HSTRING activatableClassId,
		REFIID iid,
		void** factory
		);
	typedef HRESULT(FAR STDAPICALLTYPE* T_WindowsCreateStringReference)(
		PCWSTR sourceString,
		UINT32 Length,
		HSTRING_HEADER* hstringHeader,
		HSTRING* string
		);
	typedef PCWSTR(FAR STDAPICALLTYPE* T_WindowsGetStringRawBuffer)(
		HSTRING string, 
		UINT32* Length
		);
	typedef HRESULT(FAR STDAPICALLTYPE* T_WindowsDeleteString)(
		HSTRING string
		);
    typedef NTSTATUS(WINAPI* T_RtlGetVersion)(
        PRTL_OSVERSIONINFOW pVersionInfo
        );

	static T_SetCurrentProcessExplicitAppUserModelID pSetCurrentProcessExplicitAppUserModelID = NULL;
	static T_PropVariantToString pPropVariantToString = NULL;
	static T_RoGetActivationFactory pRoGetActivationFactory = NULL;
	static T_WindowsCreateStringReference pWindowsCreateStringReference = NULL;
	static T_WindowsGetStringRawBuffer pWindowsGetStringRawBuffer = NULL;
	static T_WindowsDeleteString pWindowsDeleteString = NULL;
    static T_RtlGetVersion pRtlGetVersion = NULL;

	BOOL IsCompatible() {
		return !(pSetCurrentProcessExplicitAppUserModelID == NULL ||
			pPropVariantToString == NULL ||
			pRoGetActivationFactory == NULL ||
			pWindowsCreateStringReference == NULL ||
			pWindowsGetStringRawBuffer == NULL ||
			pWindowsDeleteString == NULL ||
            pRtlGetVersion == NULL);
	}

	HRESULT SetCurrentProcessExplicitAppUserModelID(
		PCWSTR AppID
	) {
		assert(pSetCurrentProcessExplicitAppUserModelID != 0);
		return pSetCurrentProcessExplicitAppUserModelID(AppID);
	}

	HRESULT PropVariantToString(
		REFPROPVARIANT propvar,
		PWSTR psz,
		UINT cch
	) {
		assert(pPropVariantToString != 0);
		return pPropVariantToString(propvar, psz, cch);
	}

	template <typename Tp>
	HRESULT RoGetActivationFactory(
		HSTRING activatableClassId,
		Details::ComPtrRef<Tp> factory
	) noexcept {
		assert(pRoGetActivationFactory != 0);
		return pRoGetActivationFactory(activatableClassId, IID_INS_ARGS(factory.ReleaseAndGetAddressOf()));
	}

	HRESULT WindowsCreateStringReference(
		PCWSTR sourceString,
		UINT32 Length,
		HSTRING_HEADER* hstringHeader,
		HSTRING* string
	) {
		assert(pWindowsCreateStringReference != 0);
		return pWindowsCreateStringReference(sourceString, Length, hstringHeader, string);
	}

	PCWSTR WindowsGetStringRawBuffer(
		HSTRING string,
		UINT32* Length
	) {
		assert(pWindowsGetStringRawBuffer != 0);
		return pWindowsGetStringRawBuffer(string, Length);
	}

	HRESULT WindowsDeleteString(
		HSTRING string
	) {
		assert(pWindowsDeleteString != 0);
		return pWindowsDeleteString(string);
	}

    NTSTATUS RtlGetVersion(
        PRTL_OSVERSIONINFOW pVersionInfo
    ) {
        assert(pRtlGetVersion != 0);
        return pRtlGetVersion(pVersionInfo);
    }

	HRESULT Initialize() {
		HRESULT hRes = S_OK;
		hRes = LoadFunctionFromLibrary(TEXT("SHELL32.DLL"), TEXT("SetCurrentProcessExplicitAppUserModelID"), &pSetCurrentProcessExplicitAppUserModelID);
		if (FAILED(hRes)) return hRes;
		hRes = LoadFunctionFromLibrary(TEXT("PROPSYS.DLL"), TEXT("PropVariantToString"), &pPropVariantToString);
		if (FAILED(hRes)) return hRes;
		hRes = LoadFunctionFromLibrary(TEXT("COMBASE.DLL"), TEXT("RoGetActivationFactory"), &pRoGetActivationFactory);
		if (FAILED(hRes)) return hRes;
		hRes = LoadFunctionFromLibrary(TEXT("COMBASE.DLL"), TEXT("WindowsCreateStringReference"), &pWindowsCreateStringReference);
		if (FAILED(hRes)) return hRes;
		hRes = LoadFunctionFromLibrary(TEXT("COMBASE.DLL"), TEXT("WindowsGetStringRawBuffer"), &pWindowsGetStringRawBuffer);
		if (FAILED(hRes)) return hRes;
		hRes = LoadFunctionFromLibrary(TEXT("COMBASE.DLL"), TEXT("WindowsDeleteString"), &pWindowsDeleteString);
		if (FAILED(hRes)) return hRes;
        hRes = LoadFunctionFromLibrary(TEXT("NTDLL.DLL"), TEXT("RtlGetVersion"), &pRtlGetVersion);
        if (FAILED(hRes)) return hRes;
		return S_OK;
	}
};

// String Wrapper
class StringWrapper {
public:
    StringWrapper(_In_reads_(Length) PCWSTR stringRef, UINT32 Length) noexcept {
        HRESULT hRes = DllImporter::WindowsCreateStringReference(stringRef, Length, &header, &hstring);
        if (!SUCCEEDED(hRes)) {
            RaiseException(static_cast<DWORD>(STATUS_INVALID_PARAMETER), EXCEPTION_NONCONTINUABLE, 0, nullptr);
        }
    }

    StringWrapper(std::wstring const& stringRef) noexcept {
        HRESULT hRes =
            DllImporter::WindowsCreateStringReference(stringRef.c_str(), static_cast<UINT32>(stringRef.length()), &header, &hstring);
        if (FAILED(hRes)) {
            RaiseException(static_cast<DWORD>(STATUS_INVALID_PARAMETER), EXCEPTION_NONCONTINUABLE, 0, nullptr);
        }
    }

    ~StringWrapper() {
        DllImporter::WindowsDeleteString(hstring);
    }

    inline HSTRING Get() const noexcept {
        return hstring;
    }

private:
    HSTRING hstring;
    HSTRING_HEADER header;
};



/* ==================================== WinApp Begin ==================================== */

// Default shelllink format and path ( %APPDATA%\\DEFAULT_SHELL_LINKS_PATH\\*DEFAULT_LINK_FORMAT )
#define DEFAULT_SHELL_LINKS_PATH TEXT("\\Microsoft\\Windows\\Start Menu\\Programs\\")
#define DEFAULT_LINK_FORMAT TEXT(".lnk")

WinApp::WinApp() 
	: _AppName(TEXT(""))
	, _AppID(TEXT(""))
	, _bRegistered(FALSE)
    , _bIsReady(FALSE)
	, _ShortcutPolicy(ShortcutPolicy_RequireNoCreate)
{
    HRESULT hRes = CoInitialize(NULL);
    assert(SUCCEEDED(hRes) || hRes == RPC_E_CHANGED_MODE);
    DllImporter::Initialize();
}
WinApp::~WinApp() {
    CoUninitialize();
}
std::_tstring WinApp::GetAppName() {
	return _AppName;
}
std::_tstring WinApp::GetAppID() {
	return _AppID;
}
WinAppRegisterShortcutPolicy WinApp::GetShortcutPolicy() {
	return _ShortcutPolicy;
}
BOOL WinApp::IsRegistered() {
    return _bRegistered;
}
BOOL WinApp::IsReady() {
    return _bIsReady;
}

// File Path Helper
namespace FilePathHelper {
    HRESULT DefaultExecutablePath(TCHAR* path, DWORD nSize = MAX_PATH) {
        if (path == NULL) return E_INVALIDARG;
        DWORD written = GetModuleFileNameEx(GetCurrentProcess(), nullptr, path, nSize);
        return (written > 0) ? S_OK : HRESULT_FROM_WIN32(written);
    }
    HRESULT DefaultShellLinksDirectory(TCHAR* path, DWORD nSize = MAX_PATH) {
        if (path == NULL) return E_INVALIDARG;
        DWORD written = GetEnvironmentVariable(TEXT("APPDATA"), path, nSize);
        HRESULT hRes = written > 0 ? S_OK : E_INVALIDARG;
        if (SUCCEEDED(hRes)) {
            errno_t result = _tcscat_s(path, nSize, DEFAULT_SHELL_LINKS_PATH);
            hRes = (result == 0) ? S_OK : E_INVALIDARG;
        }
        return hRes;
    }
    HRESULT DefaultShellLinkPath(std::_tstring const& appname, TCHAR* path, DWORD nSize = MAX_PATH) {
        if (path == NULL) return E_INVALIDARG;
        HRESULT hRes = DefaultShellLinksDirectory(path, nSize);
        if (SUCCEEDED(hRes)) {
            const std::_tstring appLink(appname + DEFAULT_LINK_FORMAT);
            errno_t result = _tcscat_s(path, nSize, appLink.c_str());
            hRes = (result == 0) ? S_OK : E_INVALIDARG;
        }
        return hRes;
    }
    std::_tstring ParentDirectory(TCHAR* path, DWORD size) {
        if (path == NULL) return TEXT("");
        size_t lastSeparator = 0;
        for (size_t i = 0; i < size; i++) {
            if (path[i] == TEXT('\\') || path[i] == TEXT('/')) {
                lastSeparator = i;
            }
        }
        return { path, lastSeparator };
    }

    PCWSTR AsString(ComPtr<IXmlDocument>& XmlDocument) {
        HSTRING xml;
        ComPtr<IXmlNodeSerializer> ser;
        HRESULT hRes = XmlDocument.As<IXmlNodeSerializer>(&ser);
        hRes = ser->GetXml(&xml);
        if (SUCCEEDED(hRes)) {
            return DllImporter::WindowsGetStringRawBuffer(xml, nullptr);
        }
        return nullptr;
    }
    PCWSTR AsString(HSTRING hstring) {
        return DllImporter::WindowsGetStringRawBuffer(hstring, nullptr);
    }
};

HRESULT WinApp::ValidateShellLink(BOOL& bWasChanged) {

    std::_tstring RegAppID = GetAppID(), RegAppName = GetAppName();
    WinAppRegisterShortcutPolicy RegAppPolicy = GetShortcutPolicy();
    if (RegAppID.empty() || RegAppName.empty()) return E_INVALIDARG;

    TCHAR path[MAX_PATH] = { TEXT('\0') };
    BOOL WasChanged = FALSE;
    FilePathHelper::DefaultShellLinkPath(RegAppName, path);
    // Check if the file exist
    DWORD attr = GetFileAttributes(path);
    if (attr == INVALID_FILE_ATTRIBUTES) return E_NOT_SET;
    
    // Let's load the file as shell link to validate.
    // - Create a shell link
    // - Create a persistant file
    // - Load the path as data for the persistant file
    // - Read the property AUMI and validate with the current
    // - Review if AUMI is equal.
    ComPtr<IShellLink> ShellLink;
    HRESULT hRes = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&ShellLink));
    if (SUCCEEDED(hRes)) {
        ComPtr<IPersistFile> PersistFile;
        hRes = ShellLink.As(&PersistFile);
        if (SUCCEEDED(hRes)) {
            hRes = PersistFile->Load(LPT2LPW(std::_tstring(path)).c_str(), STGM_READWRITE);
            if (SUCCEEDED(hRes)) {
                return S_OK;
                ComPtr<IPropertyStore> PropertyStore;
                hRes = ShellLink.As(&PropertyStore);
                if (SUCCEEDED(hRes)) {
                    PROPVARIANT AppIdPropVar = { 0 };
                    hRes = PropertyStore->GetValue(PKEY_AppUserModel_ID, &AppIdPropVar);
                    if (SUCCEEDED(hRes)) {
                        WCHAR AUMI[MAX_PATH];
                        hRes = DllImporter::PropVariantToString(AppIdPropVar, AUMI, MAX_PATH);
                        bWasChanged = FALSE;
                        if (FAILED(hRes) || LPT2LPW(RegAppID) != AUMI) {
                            if (RegAppPolicy == ShortcutPolicy_RequireCreate) {
                                // AUMI Changed for the same app, let's update the current value! =)
                                bWasChanged = TRUE;
                                PropVariantClear(&AppIdPropVar);
                                hRes = InitPropVariantFromString(LPT2LPW(RegAppID).c_str(), &AppIdPropVar);
                                if (SUCCEEDED(hRes)) {
                                    hRes = PropertyStore->SetValue(PKEY_AppUserModel_ID, AppIdPropVar);
                                    if (SUCCEEDED(hRes)) {
                                        hRes = PropertyStore->Commit();
                                        if (SUCCEEDED(hRes) && SUCCEEDED(PersistFile->IsDirty())) {
                                            hRes = PersistFile->Save(LPT2LPW(std::_tstring(path)).c_str(), TRUE);
                                        }
                                    }
                                }
                            }
                            else {
                                // Not allowed to touch the shortcut to fix the AUMI
                                hRes = E_FAIL;
                            }
                        }
                        PropVariantClear(&AppIdPropVar);
                    }
                }
            }
        }
    }
    return hRes;
}
HRESULT WinApp::CreateShellLink() {

    std::_tstring RegAppID = GetAppID(), RegAppName = GetAppName();
    WinAppRegisterShortcutPolicy RegAppPolicy = GetShortcutPolicy();
    if (RegAppID.empty() || RegAppName.empty()) return E_INVALIDARG;
    
    TCHAR ExePath[MAX_PATH]{ TEXT('\0') };
    TCHAR SlPath[MAX_PATH]{ TEXT('\0') };
    FilePathHelper::DefaultShellLinkPath(RegAppName, SlPath);
    FilePathHelper::DefaultExecutablePath(ExePath);
    std::_tstring ExeDir = FilePathHelper::ParentDirectory(ExePath, sizeof(ExePath) / sizeof(ExePath[0]));
    ComPtr<IShellLink> ShellLink;
    HRESULT hRes = CoCreateInstance(CLSID_ShellLink, nullptr, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&ShellLink));
    if (SUCCEEDED(hRes)) {
        hRes = ShellLink->SetPath(ExePath);
        if (SUCCEEDED(hRes)) {
            hRes = ShellLink->SetArguments(TEXT(""));
            if (SUCCEEDED(hRes)) {
                hRes = ShellLink->SetWorkingDirectory(ExeDir.c_str());
                if (SUCCEEDED(hRes)) {
                    ComPtr<IPropertyStore> PropertyStore;
                    hRes = ShellLink.As(&PropertyStore);
                    if (SUCCEEDED(hRes)) {
                        PROPVARIANT AppIdPropVar;
                        hRes = InitPropVariantFromString(LPT2LPW(RegAppID).c_str(), &AppIdPropVar);
                        if (SUCCEEDED(hRes)) {
                            hRes = PropertyStore->SetValue(PKEY_AppUserModel_ID, AppIdPropVar);
                            if (SUCCEEDED(hRes)) {
                                hRes = PropertyStore->Commit();
                                if (SUCCEEDED(hRes)) {
                                    ComPtr<IPersistFile> PersistFile;
                                    hRes = ShellLink.As(&PersistFile);
                                    if (SUCCEEDED(hRes)) {
                                        hRes = PersistFile->Save(LPT2LPW(std::_tstring(SlPath)).c_str(), TRUE);
                                    }
                                }
                            }
                            PropVariantClear(&AppIdPropVar);
                        }
                    }
                }
            }
        }
    }
    return hRes;
}

HRESULT WinApp::SetAppName(std::_tstring AppName) {
	if (AppName.empty()) return E_INVALIDARG;
	_AppName = AppName;
	return S_OK;
}
HRESULT WinApp::SetAppID(std::_tstring AppID) {
	if (AppID.empty()) return E_INVALIDARG;
	_AppID = AppID;
	return S_OK;
}
HRESULT WinApp::SetAppID(std::_tstring CompanyName, std::_tstring ProductName, std::_tstring SubProductName, std::_tstring VersionInfo) {
	std::_tstring AppID = CompanyName + TEXT(".") + ProductName + TEXT(".") + SubProductName + TEXT(".") + VersionInfo;
	return SetAppID(AppID);
}
HRESULT WinApp::SetShorcutPolicy(WinAppRegisterShortcutPolicy Policy) {
	_ShortcutPolicy = Policy;
    return S_OK;
}

HRESULT WinApp::RegisterApp(BOOL bWaitForSystemIfAppNotExist) {
	if (_bRegistered == TRUE) return S_OK;
	if (!DllImporter::IsCompatible()) return E_NOT_SUPPORT;

	std::_tstring RegAppID = GetAppID(), RegAppName = GetAppName();
    WinAppRegisterShortcutPolicy RegAppPolicy = GetShortcutPolicy();
	if (RegAppID.empty() || RegAppName.empty()) return E_INVALIDARG;
    HRESULT hRes;

    BOOL bWasChanged = FALSE, bIsExist = TRUE;
    if (RegAppPolicy != ShortcutPolicy_Ignore) {
        if (FAILED(ValidateShellLink(bWasChanged))) {
            hRes = CreateShellLink();
            bIsExist = FALSE;
            bWasChanged = TRUE;
            if (FAILED(hRes)) return hRes;
        }
    }
    
    hRes = DllImporter::SetCurrentProcessExplicitAppUserModelID(LPT2LPW(RegAppID).c_str());
    if (FAILED(hRes)) return hRes;
    if (bWasChanged && bWaitForSystemIfAppNotExist) {
        CreateThread(NULL, 0, [](LPVOID lpParam)->DWORD {
            assert(lpParam != 0);
            WinApp* Object = reinterpret_cast <WinApp*> (lpParam);
            assert(Object != 0);
            Sleep(WINAPP_SYSTEM_TIMEOUT);
            Object->_bIsReady = TRUE;
            return 0;
            }, reinterpret_cast <LPVOID> (this), 0, 0);
    }
    else _bIsReady = TRUE;
    _bRegistered = TRUE;
    if (bIsExist && !bWasChanged) return S_SKIPPED;
    return S_OK;
}

/* ===================================== WinApp End ===================================== */





/* =================================== WinToast Begin =================================== */

// WinToast Handler
namespace WinToastHandler {
    struct ToastHandler {
        WinToast_Clicked _WinToastActivated_Clicked;
        WinToast_ActionSelected _WinToastActivated_ActionSelected;
        WinToast_Replied _WinToastActivated_Replied;
        WinToast_Dismissed _WinToastDismissed;
        WinToast_Failed _WinToastFailed;

        ToastHandler();
        void ToastActivated() const;
        void ToastActivated(INT ActionIndex) const;
        void ToastActivated(LPCSTR Response) const;
        void ToastDismissed(WinToastDismissalReason State) const;
        void ToastFailed() const;
    };

    ToastHandler::ToastHandler() {
        _WinToastActivated_Clicked = 0;
        _WinToastActivated_ActionSelected = 0;
        _WinToastActivated_Replied = 0;
        _WinToastDismissed = 0;
        _WinToastFailed = 0;
    }
    void ToastHandler::ToastActivated() const {
        if (_WinToastActivated_Clicked) _WinToastActivated_Clicked();
    }
    void ToastHandler::ToastActivated(INT ActionIndex) const {
        if (_WinToastActivated_ActionSelected) _WinToastActivated_ActionSelected(ActionIndex);
    }
    void ToastHandler::ToastActivated(LPCSTR Response) const {
        if (_WinToastActivated_Replied) _WinToastActivated_Replied(LPC2LPT(Response));
    }
    void ToastHandler::ToastDismissed(WinToastDismissalReason State) const {
        if (_WinToastDismissed) _WinToastDismissed(State);
    }
    void ToastHandler::ToastFailed() const {
        if (_WinToastFailed) _WinToastFailed();
    }

    static ToastHandler Handler;
};

// Internal Date Time Class
class InternalDateTime : public IReference<DateTime> {
public:
    static INT64 Now() {
        FILETIME now;
        GetSystemTimeAsFileTime(&now);
        return ((((INT64)now.dwHighDateTime) << 32) | now.dwLowDateTime);
    }
    InternalDateTime(DateTime DateTime) : _DateTime(DateTime) {}
    InternalDateTime(INT64 MillisecondsFromNow) {
        _DateTime.UniversalTime = Now() + MillisecondsFromNow * 10000;
    }
    virtual ~InternalDateTime() = default;
    operator INT64() {
        return _DateTime.UniversalTime;
    }
    HRESULT STDMETHODCALLTYPE get_Value(DateTime* DateTime) {
        *DateTime = _DateTime;
        return S_OK;
    }
    HRESULT STDMETHODCALLTYPE QueryInterface(const IID& riid, void** ppvObject) {
        if (!ppvObject) {
            return E_HANDLE;
        }
        if (riid == __uuidof(IUnknown) || riid == __uuidof(IReference<DateTime>)) {
            *ppvObject = static_cast<IUnknown*>(static_cast<IReference<DateTime>*>(this));
            return S_OK;
        }
        return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE Release() {
        return 1;
    }
    ULONG STDMETHODCALLTYPE AddRef() {
        return 2;
    }
    HRESULT STDMETHODCALLTYPE GetIids(ULONG*, IID**) {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetRuntimeClassName(HSTRING*) {
        return E_NOTIMPL;
    }
    HRESULT STDMETHODCALLTYPE GetTrustLevel(TrustLevel*) {
        return E_NOTIMPL;
    }

protected:
    DateTime _DateTime;
};

// Xml Document Helper
namespace XmlHelper {
    HRESULT SetNodeStringValue(std::wstring const& string, IXmlNode* node, IXmlDocument* xml) {
        ComPtr<IXmlText> TextNode;
        HRESULT hRes = xml->CreateTextNode(StringWrapper(string).Get(), &TextNode);
        if (SUCCEEDED(hRes)) {
            ComPtr<IXmlNode> stringNode;
            hRes = TextNode.As(&stringNode);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlNode> appendedChild;
                hRes = node->AppendChild(stringNode.Get(), &appendedChild);
            }
        }
        return hRes;
    }

    template <typename FunctorT>
    HRESULT SetEventHandlers(
        IToastNotification* Notification,
        std::shared_ptr<WinToastHandler::ToastHandler> EventHandler,
        INT64 ExpirationTime,
        EventRegistrationToken& ActivatedToken,
        EventRegistrationToken& DismissedToken,
        EventRegistrationToken& FailedToken,
        FunctorT&& MarkAsReadyForDeletionFunc
    ) {
        HRESULT hRes = Notification->add_Activated(
            Callback<Implements<RuntimeClassFlags<ClassicCom>, ITypedEventHandler<ToastNotification*, IInspectable*>>>(
                [EventHandler, MarkAsReadyForDeletionFunc](IToastNotification* notify, IInspectable* inspectable)
                {
                    ComPtr<IToastActivatedEventArgs> activatedEventArgs;
                    HRESULT hRes = inspectable->QueryInterface(activatedEventArgs.GetAddressOf());
                    if (SUCCEEDED(hRes)) {
                        HSTRING argumentsHandle;
                        hRes = activatedEventArgs->get_Arguments(&argumentsHandle);
                        if (SUCCEEDED(hRes)) {
                            PCWSTR arguments = FilePathHelper::AsString(argumentsHandle);

                            if (wcscmp(arguments, L"action=reply") == 0)
                            {
                                ComPtr<IToastActivatedEventArgs2> inputBoxActivatedEventArgs;
                                HRESULT hRes2 = inspectable->QueryInterface(inputBoxActivatedEventArgs.GetAddressOf());

                                if (SUCCEEDED(hRes2))
                                {
                                    ComPtr<Collections::IPropertySet> replyHandle;
                                    inputBoxActivatedEventArgs->get_UserInput(&replyHandle);

                                    ComPtr<__FIMap_2_HSTRING_IInspectable> replyMap;
                                    hRes = replyHandle.As(&replyMap);

                                    if (SUCCEEDED(hRes))
                                    {
                                        IInspectable* propertySet;
                                        hRes = replyMap.Get()->Lookup(StringWrapper(L"textBox").Get(), &propertySet);
                                        if (SUCCEEDED(hRes))
                                        {
                                            ComPtr<IPropertyValue> propertyValue;
                                            hRes = propertySet->QueryInterface(IID_PPV_ARGS(&propertyValue));

                                            if (SUCCEEDED(hRes))
                                            {
                                                // Successfully queried IPropertyValue, now extract the value
                                                HSTRING userInput;
                                                hRes = propertyValue->GetString(&userInput);

                                                // Convert the HSTRING to a wide string
                                                PCWSTR strValueW = FilePathHelper::AsString(userInput);

                                                // Convert the wide string to a STL std::string to pass it as parameter
                                                // into the event.
                                                // std::wstring ogWstr(strValueW);
                                                // std::string str(ogWstr.length(), ' ');
                                                // std::copy(ogWstr.begin(), ogWstr.end(), str.begin());

                                                if (SUCCEEDED(hRes))
                                                {
                                                    EventHandler->ToastActivated(LPW2LPC(strValueW).c_str());
                                                    return S_OK;
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (arguments && *arguments) {
                                EventHandler->ToastActivated(static_cast<INT>(wcstol(arguments, nullptr, 10)));
                                DllImporter::WindowsDeleteString(argumentsHandle);
                                MarkAsReadyForDeletionFunc();
                                return S_OK;
                            }
                            DllImporter::WindowsDeleteString(argumentsHandle);
                        }
                    }
                    EventHandler->ToastActivated();
                    MarkAsReadyForDeletionFunc();
                    return S_OK;
                })
            .Get(),
            &ActivatedToken);

        if (SUCCEEDED(hRes)) {
            hRes = Notification->add_Dismissed(
                Callback<Implements<RuntimeClassFlags<ClassicCom>, ITypedEventHandler<ToastNotification*, ToastDismissedEventArgs*>>>(
                    [EventHandler, ExpirationTime, MarkAsReadyForDeletionFunc](IToastNotification* notify, IToastDismissedEventArgs* e) {
                        ToastDismissalReason reason;
                        if (SUCCEEDED(e->get_Reason(&reason))) {
                            if (reason == ToastDismissalReason_UserCanceled && ExpirationTime &&
                                InternalDateTime::Now() >= ExpirationTime) {
                                reason = ToastDismissalReason_TimedOut;
                            }
                            EventHandler->ToastDismissed(static_cast<WinToastDismissalReason>(reason));
                        }
                        MarkAsReadyForDeletionFunc();
                        return S_OK;
                    })
                .Get(),
                &DismissedToken);
            if (SUCCEEDED(hRes)) {
                hRes = Notification->add_Failed(
                    Callback<Implements<RuntimeClassFlags<ClassicCom>, ITypedEventHandler<ToastNotification*, ToastFailedEventArgs*>>>(
                        [EventHandler, MarkAsReadyForDeletionFunc](IToastNotification* notify, IToastFailedEventArgs* e) {
                            EventHandler->ToastFailed();
                            MarkAsReadyForDeletionFunc();
                            return S_OK;
                        })
                    .Get(),
                    &FailedToken);
            }
        }
        return hRes;
    }

    HRESULT AddAttribute(IXmlDocument* xml, std::wstring const& name, IXmlNamedNodeMap* attributeMap) {
        ComPtr<IXmlAttribute> srcAttribute;
        HRESULT hRes = xml->CreateAttribute(StringWrapper(name).Get(), &srcAttribute);
        if (SUCCEEDED(hRes)) {
            ComPtr<IXmlNode> node;
            hRes = srcAttribute.As(&node);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlNode> pNode;
                hRes = attributeMap->SetNamedItem(node.Get(), &pNode);
            }
        }
        return hRes;
    }

    HRESULT CreateElement(IXmlDocument* xml, std::wstring const& RootNode, std::wstring const& ElementName,
        std::vector<std::wstring> const& AttributeNames) {
        ComPtr<IXmlNodeList> rootList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(RootNode).Get(), &rootList);
        if (SUCCEEDED(hRes)) {
            ComPtr<IXmlNode> root;
            hRes = rootList->Item(0, &root);
            if (SUCCEEDED(hRes)) {
                ComPtr<ABI::Windows::Data::Xml::Dom::IXmlElement> audioElement;
                hRes = xml->CreateElement(StringWrapper(ElementName).Get(), &audioElement);
                if (SUCCEEDED(hRes)) {
                    ComPtr<IXmlNode> audioNodeTmp;
                    hRes = audioElement.As(&audioNodeTmp);
                    if (SUCCEEDED(hRes)) {
                        ComPtr<IXmlNode> audioNode;
                        hRes = root->AppendChild(audioNodeTmp.Get(), &audioNode);
                        if (SUCCEEDED(hRes)) {
                            ComPtr<IXmlNamedNodeMap> Attributes;
                            hRes = audioNode->get_Attributes(&Attributes);
                            if (SUCCEEDED(hRes)) {
                                for (auto const& it : AttributeNames) {
                                    hRes = AddAttribute(xml, it, Attributes.Get());
                                }
                            }
                        }
                    }
                }
            }
        }
        return hRes;
    }

//
// Available as of Windows 10 Anniversary Update
// Ref: https://docs.microsoft.com/en-us/windows/uwp/design/shell/tiles-and-Notifications/adaptive-INTeractive-toasts
//
// NOTE: This will add a new text field, so be aware when iterating over
//       the toast's text fields or getting a count of them.
//
    
    HRESULT AddDuration_Xml(IXmlDocument* xml, std::wstring const& duration) {
        ComPtr<IXmlNodeList> NodeList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"toast").Get(), &NodeList);
        if (SUCCEEDED(hRes)) {
            UINT32 Length;
            hRes = NodeList->get_Length(&Length);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlNode> ToastNode;
                hRes = NodeList->Item(0, &ToastNode);
                if (SUCCEEDED(hRes)) {
                    ComPtr<IXmlElement> ToastElement;
                    hRes = ToastNode.As(&ToastElement);
                    if (SUCCEEDED(hRes)) {
                        hRes = ToastElement->SetAttribute(StringWrapper(L"duration").Get(), StringWrapper(duration).Get());
                    }
                }
            }
        }
        return hRes;
    }
    HRESULT AddScenario_Xml(IXmlDocument* xml, std::wstring const& scenario) {
        ComPtr<IXmlNodeList> NodeList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"toast").Get(), &NodeList);
        if (SUCCEEDED(hRes)) {
            UINT32 Length;
            hRes = NodeList->get_Length(&Length);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlNode> ToastNode;
                hRes = NodeList->Item(0, &ToastNode);
                if (SUCCEEDED(hRes)) {
                    ComPtr<IXmlElement> ToastElement;
                    hRes = ToastNode.As(&ToastElement);
                    if (SUCCEEDED(hRes)) {
                        hRes = ToastElement->SetAttribute(StringWrapper(L"scenario").Get(), StringWrapper(scenario).Get());
                    }
                }
            }
        }
        return hRes;
    }
    HRESULT AddInput_Xml(IXmlDocument* xml, std::wstring const& backgroundText, std::wstring const& replyText)
    {
        std::vector<std::wstring> attrbs;
        attrbs.push_back(L"id");
        attrbs.push_back(L"type");
        attrbs.push_back(L"placeHolderContent");

        std::vector<std::wstring> attrbs2;
        attrbs2.push_back(L"content");
        attrbs2.push_back(L"arguments");

        CreateElement(xml, L"toast", L"actions", {});

        CreateElement(xml, L"actions", L"input", attrbs);
        CreateElement(xml, L"actions", L"action", attrbs2);

        ComPtr<IXmlNodeList> NodeList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"input").Get(), &NodeList);
        if (SUCCEEDED(hRes))
        {
            ComPtr<IXmlNode> inputNode;
            hRes = NodeList->Item(0, &inputNode);
            if (SUCCEEDED(hRes))
            {
                ComPtr<IXmlElement> ToastElement;
                hRes = inputNode.As(&ToastElement);
                if (SUCCEEDED(hRes)) {
                    ToastElement->SetAttribute(StringWrapper(L"id").Get(), StringWrapper(L"textBox").Get());
                    ToastElement->SetAttribute(StringWrapper(L"type").Get(), StringWrapper(L"text").Get());
                    hRes = ToastElement->SetAttribute(StringWrapper(L"placeHolderContent").Get(), StringWrapper(backgroundText).Get());
                }
            }
        }

        ComPtr<IXmlNodeList> NodeList2;
        hRes = xml->GetElementsByTagName(StringWrapper(L"action").Get(), &NodeList2);
        if (SUCCEEDED(hRes))
        {
            ComPtr<IXmlNode> actionNode;
            hRes = NodeList2->Item(0, &actionNode);
            if (SUCCEEDED(hRes))
            {
                ComPtr<IXmlElement> actionElement;
                hRes = actionNode.As(&actionElement);
                if (SUCCEEDED(hRes)) {
                    actionElement->SetAttribute(StringWrapper(L"content").Get(), StringWrapper(replyText).Get());
                    actionElement->SetAttribute(StringWrapper(L"arguments").Get(), StringWrapper(L"action=reply").Get());
                    actionElement->SetAttribute(StringWrapper(L"hint-inputId").Get(), StringWrapper(L"textBox").Get());
                }
            }
        }
        return hRes;
    }
    HRESULT SetTextField_Xml(IXmlDocument* xml, std::wstring const& text, UINT32 pos) {
        ComPtr<IXmlNodeList> NodeList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"text").Get(), &NodeList);
        if (SUCCEEDED(hRes)) {
            ComPtr<IXmlNode> node;
            hRes = NodeList->Item(pos, &node);
            if (SUCCEEDED(hRes)) {
                hRes = SetNodeStringValue(text, node.Get(), xml);
            }
        }
        return hRes;
    }
    HRESULT SetBindToastGeneric_Xml(IXmlDocument* xml) {
        ComPtr<IXmlNodeList> NodeList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"binding").Get(), &NodeList);
        if (SUCCEEDED(hRes)) {
            UINT32 Length;
            hRes = NodeList->get_Length(&Length);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlNode> ToastNode;
                hRes = NodeList->Item(0, &ToastNode);
                if (SUCCEEDED(hRes)) {
                    ComPtr<IXmlElement> ToastElement;
                    hRes = ToastNode.As(&ToastElement);
                    if (SUCCEEDED(hRes)) {
                        hRes = ToastElement->SetAttribute(StringWrapper(L"template").Get(), StringWrapper(L"ToastGeneric").Get());
                    }
                }
            }
        }
        return hRes;
    }
    HRESULT SetImageField_Xml(IXmlDocument* xml, std::wstring const& path, BOOL isToastGeneric,
        BOOL isCropHintCircle) {
        assert(path.size() < MAX_PATH);

        wchar_t imagePath[MAX_PATH] = L"file:///";
        HRESULT hRes = StringCchCatW(imagePath, MAX_PATH, path.c_str());
        if (SUCCEEDED(hRes)) {
            ComPtr<IXmlNodeList> NodeList;
            HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"image").Get(), &NodeList);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlNode> node;
                hRes = NodeList->Item(0, &node);

                ComPtr<IXmlElement> imageElement;
                HRESULT hResImage = node.As(&imageElement);
                if (SUCCEEDED(hRes) && SUCCEEDED(hResImage) && isToastGeneric) {
                    hRes = imageElement->SetAttribute(StringWrapper(L"placement").Get(), StringWrapper(L"appLogoOverride").Get());
                    if (SUCCEEDED(hRes) && isCropHintCircle) {
                        hRes = imageElement->SetAttribute(StringWrapper(L"hint-crop").Get(), StringWrapper(L"circle").Get());
                    }
                }
                if (SUCCEEDED(hRes)) {
                    ComPtr<IXmlNamedNodeMap> Attributes;
                    hRes = node->get_Attributes(&Attributes);
                    if (SUCCEEDED(hRes)) {
                        ComPtr<IXmlNode> EditedNode;
                        hRes = Attributes->GetNamedItem(StringWrapper(L"src").Get(), &EditedNode);
                        if (SUCCEEDED(hRes)) {
                            SetNodeStringValue(imagePath, EditedNode.Get(), xml);
                        }
                    }
                }
            }
        }
        return hRes;
    }
    HRESULT SetAudioField_Xml(IXmlDocument* xml, std::wstring const& path,
        WinToastAudioOption option) {
        std::vector<std::wstring> attrs;
        if (!path.empty()) {
            attrs.push_back(L"src");
        }
        if (option == WinToastAudioOption::AudioOption_Loop) {
            attrs.push_back(L"loop");
        }
        if (option == WinToastAudioOption::AudioOption_Silent) {
            attrs.push_back(L"silent");
        }
        CreateElement(xml, L"toast", L"audio", attrs);

        ComPtr<IXmlNodeList> NodeList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"audio").Get(), &NodeList);
        if (SUCCEEDED(hRes)) {
            ComPtr<IXmlNode> node;
            hRes = NodeList->Item(0, &node);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlNamedNodeMap> Attributes;
                hRes = node->get_Attributes(&Attributes);
                if (SUCCEEDED(hRes)) {
                    ComPtr<IXmlNode> EditedNode;
                    if (!path.empty()) {
                        if (SUCCEEDED(hRes)) {
                            hRes = Attributes->GetNamedItem(StringWrapper(L"src").Get(), &EditedNode);
                            if (SUCCEEDED(hRes)) {
                                hRes = SetNodeStringValue(path, EditedNode.Get(), xml);
                            }
                        }
                    }

                    if (SUCCEEDED(hRes)) {
                        switch (option) {
                        case WinToastAudioOption::AudioOption_Loop:
                            hRes = Attributes->GetNamedItem(StringWrapper(L"loop").Get(), &EditedNode);
                            if (SUCCEEDED(hRes)) {
                                hRes = SetNodeStringValue(L"TRUE", EditedNode.Get(), xml);
                            }
                            break;
                        case WinToastAudioOption::AudioOption_Silent:
                            hRes = Attributes->GetNamedItem(StringWrapper(L"silent").Get(), &EditedNode);
                            if (SUCCEEDED(hRes)) {
                                hRes = SetNodeStringValue(L"TRUE", EditedNode.Get(), xml);
                            }
                        default:
                            break;
                        }
                    }
                }
            }
        }
        return hRes;
    }
    HRESULT AddAction_Xml(IXmlDocument* xml, std::wstring const& content, std::wstring const& arguments) {
        ComPtr<IXmlNodeList> NodeList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"actions").Get(), &NodeList);
        if (SUCCEEDED(hRes)) {
            UINT32 Length;
            hRes = NodeList->get_Length(&Length);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlNode> actionsNode;
                if (Length > 0) {
                    hRes = NodeList->Item(0, &actionsNode);
                }
                else {
                    hRes = xml->GetElementsByTagName(StringWrapper(L"toast").Get(), &NodeList);
                    if (SUCCEEDED(hRes)) {
                        hRes = NodeList->get_Length(&Length);
                        if (SUCCEEDED(hRes)) {
                            ComPtr<IXmlNode> ToastNode;
                            hRes = NodeList->Item(0, &ToastNode);
                            if (SUCCEEDED(hRes)) {
                                ComPtr<IXmlElement> ToastElement;
                                hRes = ToastNode.As(&ToastElement);
                                if (SUCCEEDED(hRes)) {
                                    hRes = ToastElement->SetAttribute(StringWrapper(L"template").Get(),
                                        StringWrapper(L"ToastGeneric").Get());
                                }
                                if (SUCCEEDED(hRes)) {
                                    hRes = ToastElement->SetAttribute(StringWrapper(L"duration").Get(),
                                        StringWrapper(L"long").Get());
                                }
                                if (SUCCEEDED(hRes)) {
                                    ComPtr<IXmlElement> actionsElement;
                                    hRes = xml->CreateElement(StringWrapper(L"actions").Get(), &actionsElement);
                                    if (SUCCEEDED(hRes)) {
                                        hRes = actionsElement.As(&actionsNode);
                                        if (SUCCEEDED(hRes)) {
                                            ComPtr<IXmlNode> appendedChild;
                                            hRes = ToastNode->AppendChild(actionsNode.Get(), &appendedChild);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                if (SUCCEEDED(hRes)) {
                    ComPtr<IXmlElement> actionElement;
                    hRes = xml->CreateElement(StringWrapper(L"action").Get(), &actionElement);
                    if (SUCCEEDED(hRes)) {
                        hRes = actionElement->SetAttribute(StringWrapper(L"content").Get(), StringWrapper(content).Get());
                    }
                    if (SUCCEEDED(hRes)) {
                        hRes = actionElement->SetAttribute(StringWrapper(L"arguments").Get(), StringWrapper(arguments).Get());
                    }
                    if (SUCCEEDED(hRes)) {
                        ComPtr<IXmlNode> actionNode;
                        hRes = actionElement.As(&actionNode);
                        if (SUCCEEDED(hRes)) {
                            ComPtr<IXmlNode> appendedChild;
                            hRes = actionsNode->AppendChild(actionNode.Get(), &appendedChild);
                        }
                    }
                }
            }
        }
        return hRes;
    }
    HRESULT SetHeroImage_Xml(IXmlDocument* xml, std::wstring const& path, BOOL isInlineImage) {
        ComPtr<IXmlNodeList> NodeList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"binding").Get(), &NodeList);
        if (SUCCEEDED(hRes)) {
            UINT32 Length;
            hRes = NodeList->get_Length(&Length);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlNode> bindingNode;
                if (Length > 0) {
                    hRes = NodeList->Item(0, &bindingNode);
                }
                if (SUCCEEDED(hRes)) {
                    ComPtr<IXmlElement> imageElement;
                    hRes = xml->CreateElement(StringWrapper(L"image").Get(), &imageElement);
                    if (SUCCEEDED(hRes) && isInlineImage == FALSE) {
                        hRes = imageElement->SetAttribute(StringWrapper(L"placement").Get(), StringWrapper(L"hero").Get());
                    }
                    if (SUCCEEDED(hRes)) {
                        hRes = imageElement->SetAttribute(StringWrapper(L"src").Get(), StringWrapper(path).Get());
                    }
                    if (SUCCEEDED(hRes)) {
                        ComPtr<IXmlNode> actionNode;
                        hRes = imageElement.As(&actionNode);
                        if (SUCCEEDED(hRes)) {
                            ComPtr<IXmlNode> appendedChild;
                            hRes = bindingNode->AppendChild(actionNode.Get(), &appendedChild);
                        }
                    }
                }
            }
        }
        return hRes;
    }
    HRESULT SetAttributionTextField_Xml(IXmlDocument* xml, std::wstring const& text) {
        CreateElement(xml, L"binding", L"text", { L"placement" });
        ComPtr<IXmlNodeList> NodeList;
        HRESULT hRes = xml->GetElementsByTagName(StringWrapper(L"text").Get(), &NodeList);
        if (SUCCEEDED(hRes)) {
            UINT32 NodeListLength;
            hRes = NodeList->get_Length(&NodeListLength);
            if (SUCCEEDED(hRes)) {
                for (UINT32 i = 0; i < NodeListLength; i++) {
                    ComPtr<IXmlNode> TextNode;
                    hRes = NodeList->Item(i, &TextNode);
                    if (SUCCEEDED(hRes)) {
                        ComPtr<IXmlNamedNodeMap> Attributes;
                        hRes = TextNode->get_Attributes(&Attributes);
                        if (SUCCEEDED(hRes)) {
                            ComPtr<IXmlNode> EditedNode;
                            if (SUCCEEDED(hRes)) {
                                hRes = Attributes->GetNamedItem(StringWrapper(L"placement").Get(), &EditedNode);
                                if (FAILED(hRes) || !EditedNode) {
                                    continue;
                                }
                                hRes = SetNodeStringValue(L"attribution", EditedNode.Get(), xml);
                                if (SUCCEEDED(hRes)) {
                                    return SetTextField_Xml(xml, text, i);
                                }
                            }
                        }
                    }
                }
            }
        }
        return hRes;
    }
}

// Audio Helper
namespace AudioHelper {
    const std::unordered_map<WinToastAudioSystemFile, std::_tstring> Files = {
        {WinToastAudioSystemFile::DefaultSound, TEXT("ms-winsoundevent:Notification.Default")        },
        {WinToastAudioSystemFile::IM,           TEXT("ms-winsoundevent:Notification.IM")             },
        {WinToastAudioSystemFile::Mail,         TEXT("ms-winsoundevent:Notification.Mail")           },
        {WinToastAudioSystemFile::Reminder,     TEXT("ms-winsoundevent:Notification.Reminder")       },
        {WinToastAudioSystemFile::SMS,          TEXT("ms-winsoundevent:Notification.SMS")            },
        {WinToastAudioSystemFile::Alarm,        TEXT("ms-winsoundevent:Notification.Looping.Alarm")  },
        {WinToastAudioSystemFile::Alarm2,       TEXT("ms-winsoundevent:Notification.Looping.Alarm2") },
        {WinToastAudioSystemFile::Alarm3,       TEXT("ms-winsoundevent:Notification.Looping.Alarm3") },
        {WinToastAudioSystemFile::Alarm4,       TEXT("ms-winsoundevent:Notification.Looping.Alarm4") },
        {WinToastAudioSystemFile::Alarm5,       TEXT("ms-winsoundevent:Notification.Looping.Alarm5") },
        {WinToastAudioSystemFile::Alarm6,       TEXT("ms-winsoundevent:Notification.Looping.Alarm6") },
        {WinToastAudioSystemFile::Alarm7,       TEXT("ms-winsoundevent:Notification.Looping.Alarm7") },
        {WinToastAudioSystemFile::Alarm8,       TEXT("ms-winsoundevent:Notification.Looping.Alarm8") },
        {WinToastAudioSystemFile::Alarm9,       TEXT("ms-winsoundevent:Notification.Looping.Alarm9") },
        {WinToastAudioSystemFile::Alarm10,      TEXT("ms-winsoundevent:Notification.Looping.Alarm10")},
        {WinToastAudioSystemFile::Call,         TEXT("ms-winsoundevent:Notification.Looping.Call")   },
        {WinToastAudioSystemFile::Call1,        TEXT("ms-winsoundevent:Notification.Looping.Call1")  },
        {WinToastAudioSystemFile::Call2,        TEXT("ms-winsoundevent:Notification.Looping.Call2")  },
        {WinToastAudioSystemFile::Call3,        TEXT("ms-winsoundevent:Notification.Looping.Call3")  },
        {WinToastAudioSystemFile::Call4,        TEXT("ms-winsoundevent:Notification.Looping.Call4")  },
        {WinToastAudioSystemFile::Call5,        TEXT("ms-winsoundevent:Notification.Looping.Call5")  },
        {WinToastAudioSystemFile::Call6,        TEXT("ms-winsoundevent:Notification.Looping.Call6")  },
        {WinToastAudioSystemFile::Call7,        TEXT("ms-winsoundevent:Notification.Looping.Call7")  },
        {WinToastAudioSystemFile::Call8,        TEXT("ms-winsoundevent:Notification.Looping.Call8")  },
        {WinToastAudioSystemFile::Call9,        TEXT("ms-winsoundevent:Notification.Looping.Call9")  },
        {WinToastAudioSystemFile::Call10,       TEXT("ms-winsoundevent:Notification.Looping.Call10") },
    };

    std::_tstring SystemFileToLocation(WinToastAudioSystemFile SystemFile) {
        auto const iter = Files.find(SystemFile);
        if (iter == Files.end()) return TEXT("");
        return iter->second;
    }
}

WinToast::WinToast(WinApp* pApp, WinToastType Type)
    : _App(pApp)
    , _bHasInput(FALSE)
    , _Actions({})
    , _ImagePath({})
    , _HeroImagePath({})
    , _InlineHeroImage(FALSE)
    , _AudioPath({})
    , _AttributionText({})
    , _InputReplyText({})
    , _InputBackgroundText({})
    , _Scenario({ TEXT("Default") })
    , _Expiration(0)
    , _AudioOption(WinToastAudioOption::AudioOption_Default)
    , _Type(Type)
    , _Duration(WinToastDuration::Duration_System)
    , _CropHint(WinToastCropHint::CropHint_Square)
{
    assert(pApp != 0);
    while (!pApp->IsReady()) Sleep(WINAPP_READY_WAIT_TIMEOUT);
    constexpr static std::size_t TextFieldsCount[] = { 1, 2, 2, 3, 1, 2, 2, 3 };
    _TextFields = std::vector<std::_tstring>(TextFieldsCount[Type], TEXT(""));
    HRESULT hRes = CoInitialize(NULL);
    assert(SUCCEEDED(hRes) || hRes == RPC_E_CHANGED_MODE);
}
WinToast::~WinToast() {
    _TextFields.clear();
    _Actions.clear();
    CoUninitialize();
}

// NotifyData
WinToast::NotifyData::NotifyData() 
    : Notify(NULL)
    , ReadyForDeletion(FALSE)
    , PreviouslyTokenRemoved(FALSE)
    , ActivatedToken({0})
    , DismissedToken({0})
    , FailedToken({0})
{
    
}
WinToast::NotifyData::NotifyData(
    ComPtr<IToastNotification> Notify, 
    EventRegistrationToken ActivatedToken, 
    EventRegistrationToken DismissedToken, 
    EventRegistrationToken FailedToken
)
    : Notify(Notify)
    , ActivatedToken(ActivatedToken)
    , DismissedToken(DismissedToken)
    , FailedToken(FailedToken)
    , ReadyForDeletion(FALSE)
    , PreviouslyTokenRemoved(FALSE)
{
    assert(ActivatedToken.value != 0);
    assert(DismissedToken.value != 0);
    assert(FailedToken.value != 0);
}
WinToast::NotifyData::~NotifyData() {
    RemoveTokens();
}
void WinToast::NotifyData::RemoveTokens() {
    if (!ReadyForDeletion || PreviouslyTokenRemoved) return;

    ComPtr<IToastNotification> temp;
    if (Notify.As(&temp) != S_OK) return;

    if (ActivatedToken.value != 0) {
        Notify->remove_Activated(ActivatedToken);
    }
    if (DismissedToken.value != 0) {
        Notify->remove_Dismissed(DismissedToken);
    }
    if (FailedToken.value != 0) {
        Notify->remove_Failed(FailedToken);
    }

    PreviouslyTokenRemoved = TRUE;
}
void WinToast::NotifyData::MarkAsReadyForDeletion() {
    ReadyForDeletion = TRUE;
}
BOOL WinToast::NotifyData::IsReadyForDeletion() const {
    return ReadyForDeletion;
}
IToastNotification* WinToast::NotifyData::Notification() {
    return Notify.Get();
}

std::size_t WinToast::TextFieldsCount() const {
    return _TextFields.size();
}
std::size_t WinToast::ActionsCount() const {
    return _Actions.size();
}
BOOL WinToast::HasImage() const {
    return _Type < WinToastType::ToastType_Text01;
}
BOOL WinToast::HasHeroImage() const {
    return HasImage() && !_HeroImagePath.empty();
}
std::vector<std::_tstring> const& WinToast::GetTextFields() const {
    return _TextFields;
}
std::_tstring const& WinToast::GetTextField(WinToastTextField Pos) const {
    auto const position = static_cast<std::size_t>(Pos);
    if (position >= _TextFields.size()) return TEXT("");
    return GetTextFields()[position];
}
std::_tstring const& WinToast::GetActionLabel(std::size_t Pos) const {
    if (Pos >= _Actions.size()) return TEXT("");
    return _Actions[Pos];
}
std::_tstring const& WinToast::GetImagePath() const {
    return _ImagePath;
}
std::_tstring const& WinToast::GetHeroImagePath() const {
    return _HeroImagePath;
}
std::_tstring const& WinToast::GetAudioPath() const {
    return _AudioPath;
}
std::_tstring const& WinToast::GetAttributionText() const {
    return _AttributionText;
}
std::_tstring const& WinToast::GetScenario() const {
    return _Scenario;
}
std::_tstring const& WinToast::GetInputBackgroundText() const {
    return _InputBackgroundText;
}
std::_tstring const& WinToast::GetInputReplyText() const {
    return _InputReplyText;
}
INT64 WinToast::GetExpiration() const {
    return _Expiration;
}
WinToastType WinToast::GetType() const {
    return _Type;
}
WinToastAudioOption WinToast::GetAudioOption() const {
    return _AudioOption;
}
WinToastDuration WinToast::GetDuration() const {
    return _Duration;
}
BOOL WinToast::IsToastGeneric() const {
    return HasHeroImage() || _CropHint == WinToastCropHint::CropHint_Circle;
}
BOOL WinToast::IsInlineHeroImage() const {
    return _InlineHeroImage;
}
BOOL WinToast::IsCropHintCircle() const {
    return _CropHint == WinToastCropHint::CropHint_Circle;
}
BOOL WinToast::IsInput() const {
    return _bHasInput;
}

static BOOL IsSupportingModernFeatures() {
    constexpr auto MinimumSupportedVersion = 6;
    RTL_OSVERSIONINFOW Ver;
    if (DllImporter::RtlGetVersion(&Ver) != S_OK) return FALSE;
    return Ver.dwMajorVersion > MinimumSupportedVersion;
}
static BOOL IsWin10AnniversaryOrHigher() {
    RTL_OSVERSIONINFOW Ver;
    if (DllImporter::RtlGetVersion(&Ver) != S_OK) return FALSE;
    return Ver.dwBuildNumber >= 14393;
}
void WinToast::MarkAsReadyForDeletion(INT64 id) {
    // Flush the buffer by removing all the toasts that are ready for deletion
    for (auto it = _Buffer.begin(); it != _Buffer.end();) {
        if (it->second.IsReadyForDeletion()) {
            it->second.RemoveTokens();
            it = _Buffer.erase(it);
        }
        else {
            ++it;
        }
    }

    // Mark the toast as ready for deletion (if it exists) so that it will be removed from the buffer in the next iteration
    auto const iter = _Buffer.find(id);
    if (iter != _Buffer.end()) {
        _Buffer[id].MarkAsReadyForDeletion();
    }
}

void WinToast::SetFirstLine(std::_tstring const& Text) {
    SetTextField(Text, TextField_FirstLine);
}
void WinToast::SetSecondLine(std::_tstring const& Text) {
    SetTextField(Text, TextField_SecondLine);
}
void WinToast::SetThirdLine(std::_tstring const& Text) {
    SetTextField(Text, TextField_ThirdLine);
}
void WinToast::SetTextField(std::_tstring const& Text, WinToastTextField Pos) {
    auto const position = static_cast<std::size_t>(Pos);
    if (position >= _TextFields.size()) return;
    _TextFields[position] = Text;
}
void WinToast::SetAttributionText(std::_tstring const& AttributionText) {
    _AttributionText = AttributionText;
}
void WinToast::SetImagePath(std::_tstring const& ImagePath, WinToastCropHint CropHint) {
    _ImagePath = ImagePath;
    _CropHint = CropHint;
}
void WinToast::SetHeroImagePath(std::_tstring const& ImagePath, BOOL bIsInlineImage) {
    _HeroImagePath = ImagePath;
    _InlineHeroImage = bIsInlineImage;
}
void WinToast::SetAudioPath(WinToastAudioSystemFile SystemAudio) {
    SetAudioPath(AudioHelper::SystemFileToLocation(SystemAudio));
}
void WinToast::SetAudioPath(std::_tstring const& AudioPath) {
    _AudioPath = AudioPath;
}
void WinToast::SetAudioOption(WinToastAudioOption AudioOption) {
    _AudioOption = AudioOption;
}
void WinToast::SetDuration(WinToastDuration Duration) {
    _Duration = Duration;
}
void WinToast::SetExpiration(INT64 MillisecondsFromNow) {
    _Expiration = MillisecondsFromNow;
}
void WinToast::SetScenario(WinToastScenario Scenario) {
    switch (Scenario) {
    case WinToastScenario::Scenario_Default:
        _Scenario = TEXT("Default");
        break;
    case WinToastScenario::Scenario_Alarm:
        _Scenario = TEXT("Alarm");
        break;
    case WinToastScenario::Scenario_IncomingCall:
        _Scenario = TEXT("IncomingCall");
        break;
    case WinToastScenario::Scenario_Reminder:
        _Scenario = TEXT("Reminder");
        break;
    }
}
void WinToast::AddAction(std::_tstring const& Label) {
    _Actions.push_back(Label);
}
void WinToast::AddInput(std::_tstring const& BackgroundText, std::_tstring const& ReplyText) {
    _bHasInput = TRUE;
    _InputBackgroundText = BackgroundText;
    _InputReplyText = ReplyText;
}

void WinToast::SetCallbackClicked(WinToast_Clicked Func) {
    WinToastHandler::Handler._WinToastActivated_Clicked = Func;
}
void WinToast::SetCallbackActionSelected(WinToast_ActionSelected Func) {
    WinToastHandler::Handler._WinToastActivated_ActionSelected = Func;
}
void WinToast::SetCallbackReplied(WinToast_Replied Func) {
    WinToastHandler::Handler._WinToastActivated_Replied = Func;
}
void WinToast::SetCallbackDismissed(WinToast_Dismissed Func) {
    WinToastHandler::Handler._WinToastDismissed = Func;
}
void WinToast::SetCallbackFailed(WinToast_Failed Func) {
    WinToastHandler::Handler._WinToastFailed = Func;
}

INT64 WinToast::Show(INT64* pToastID) {
    std::shared_ptr<WinToastHandler::ToastHandler> handler = std::make_shared<WinToastHandler::ToastHandler>(WinToastHandler::Handler);
    INT64 id = -1;
    if (_App == NULL) return E_HANDLE;
    if (!_App->IsRegistered()) return E_HANDLE;
    if (!handler) return E_FAIL;

    std::wstring AppID = LPT2LPW(_App->GetAppID());

    ComPtr<IToastNotificationManagerStatics> NotificationManager;
    HRESULT hRes = DllImporter::RoGetActivationFactory(
        StringWrapper(RuntimeClass_Windows_UI_Notifications_ToastNotificationManager).Get(), &NotificationManager);
    if (SUCCEEDED(hRes)) {
        ComPtr<IToastNotifier> Notifier;
        hRes = NotificationManager->CreateToastNotifierWithId(StringWrapper(AppID).Get(), &Notifier);
        if (SUCCEEDED(hRes)) {
            ComPtr<IToastNotificationFactory> NotificationFactory;
            hRes = DllImporter::RoGetActivationFactory(
                StringWrapper(RuntimeClass_Windows_UI_Notifications_ToastNotification).Get(), &NotificationFactory);
            if (SUCCEEDED(hRes)) {
                ComPtr<IXmlDocument> XmlDocument;
                hRes = NotificationManager->GetTemplateContent(ToastTemplateType(GetType()), &XmlDocument);
                if (SUCCEEDED(hRes) && IsToastGeneric()) {
                    hRes = XmlHelper::SetBindToastGeneric_Xml(XmlDocument.Get());
                }
                if (SUCCEEDED(hRes)) {
                    for (UINT32 i = 0, FieldsCount = static_cast<UINT32>(TextFieldsCount()); i < FieldsCount && SUCCEEDED(hRes); i++) {
                        hRes = XmlHelper::SetTextField_Xml(XmlDocument.Get(), LPT2LPW(GetTextField(WinToastTextField(i))), i);
                    }

                    // Modern feature are supported Windows > Windows 10
                    if (SUCCEEDED(hRes) && IsSupportingModernFeatures()) {
                        // Note that we do this *after* using GettextFieldsCount() to
                        // iterate/fill the template's text fields, since we're adding yet another text field.
                        if (SUCCEEDED(hRes) && !GetAttributionText().empty()) {
                            hRes = XmlHelper::SetAttributionTextField_Xml(XmlDocument.Get(), LPT2LPW(GetAttributionText()));
                        }

                        std::array<WCHAR, 12> buf;
                        for (std::size_t i = 0, sActionsCount = ActionsCount(); i < sActionsCount && SUCCEEDED(hRes); i++) {
                            _snwprintf_s(buf.data(), buf.size(), _TRUNCATE, L"%zd", i);
                            hRes = XmlHelper::AddAction_Xml(XmlDocument.Get(), LPT2LPW(GetActionLabel(i)), buf.data());
                        }

                        if (SUCCEEDED(hRes)) {
                            hRes = (GetAudioPath().empty() && GetAudioOption() == WinToastAudioOption::AudioOption_Default)
                                ? hRes
                                : XmlHelper::SetAudioField_Xml(XmlDocument.Get(), LPT2LPW(GetAudioPath()), GetAudioOption());
                        }

                        if (SUCCEEDED(hRes) && GetDuration() != WinToastDuration::Duration_System) {
                            hRes = XmlHelper::AddDuration_Xml(XmlDocument.Get(),
                                (GetDuration() == WinToastDuration::Duration_Short) ? L"short" : L"long");
                        }

                        if (SUCCEEDED(hRes) && IsInput()) {
                            hRes = XmlHelper::AddInput_Xml(XmlDocument.Get(), LPT2LPW(GetInputBackgroundText()), LPT2LPW(GetInputReplyText()));
                        }

                        if (SUCCEEDED(hRes)) {
                            hRes = XmlHelper::AddScenario_Xml(XmlDocument.Get(), LPT2LPW(GetScenario()));
                        }
                    }

                    if (SUCCEEDED(hRes)) {
                        BOOL IsWin10AnniversaryOrAbove = IsWin10AnniversaryOrHigher();
                        BOOL IsCircleCropHint = IsWin10AnniversaryOrAbove ? IsCropHintCircle() : FALSE;
                        hRes = HasImage()
                            ? XmlHelper::SetImageField_Xml(XmlDocument.Get(), LPT2LPW(GetImagePath()), IsToastGeneric(), IsCircleCropHint)
                            : hRes;
                        if (SUCCEEDED(hRes) && IsWin10AnniversaryOrAbove && HasHeroImage()) {
                            hRes = XmlHelper::SetHeroImage_Xml(XmlDocument.Get(), LPT2LPW(GetHeroImagePath()), IsInlineHeroImage());
                        }
                        if (SUCCEEDED(hRes)) {
                            ComPtr<IToastNotification> Notification;
                            hRes = NotificationFactory->CreateToastNotification(XmlDocument.Get(), &Notification);
                            if (SUCCEEDED(hRes)) {
                                INT64 Expiration = 0, relativeExpiration = GetExpiration();
                                if (relativeExpiration > 0) {
                                    InternalDateTime ExpirationDateTime(relativeExpiration);
                                    Expiration = ExpirationDateTime;
                                    hRes = Notification->put_ExpirationTime(&ExpirationDateTime);
                                }

                                EventRegistrationToken ActivatedToken = { 0 }, DismissedToken = { 0 }, FailedToken = { 0 };

                                GUID guid;
                                HRESULT hResGuid = CoCreateGuid(&guid);
                                id = guid.Data1;
                                if (SUCCEEDED(hRes) && SUCCEEDED(hResGuid)) {
                                    hRes = XmlHelper::SetEventHandlers(Notification.Get(), handler, Expiration, ActivatedToken, DismissedToken,
                                        FailedToken, [this, id]() { MarkAsReadyForDeletion(id); });
                                }

                                if (SUCCEEDED(hRes)) {
                                    _Buffer.emplace(id, NotifyData(Notification, ActivatedToken, DismissedToken, FailedToken));
                                    // XML Document: FilePathHelper::AsString(XmlDocument);
                                    hRes = Notifier->Show(Notification.Get());
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    if (FAILED(hRes)) return hRes;
    if (pToastID != 0) *pToastID = id;
    return S_OK;
}


/* ==================================== WinToast End ==================================== */





/* ================================== WinTaskbarIcon Begin ================================== */

WinTaskbarIcon::WinTaskbarIcon(WinApp* pApp, HWND hWnd)
    : _pTaskbar(NULL)
    , _pJumpList(NULL)
    , _hWnd(hWnd)
    , _nStateProgressValue(0)
    , _nStateProgressTotal(100)
    , _nFlashCount(TASKBAR_FLASH_INFINITY)
    , _dwFlashFlags(TaskbarFlash_All)
    , _dwFlashTimeout(500)
    , _StateNow(TaskbarState_Normal)
    , _nJumpListMaxSlots(0)
    , _bJumpListCanAdd(0)
{
    HRESULT hRes = CoInitialize(NULL);
    assert(SUCCEEDED(hRes) || hRes == RPC_E_CHANGED_MODE);
    hRes = CoCreateInstance(CLSID_TaskbarList, NULL, CLSCTX_ALL, IID_PPV_ARGS(&_pTaskbar));
    assert(SUCCEEDED(hRes) && _pTaskbar != NULL);
    hRes = _pTaskbar->HrInit();
    assert(SUCCEEDED(hRes));
    hRes = CoCreateInstance(CLSID_DestinationList, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_pJumpList));
    assert(SUCCEEDED(hRes) && _pJumpList != NULL);
    hRes = _pTaskbar->HrInit();
    assert(SUCCEEDED(hRes));
    if (pApp) {
        while (!pApp->IsReady()) Sleep(WINAPP_READY_WAIT_TIMEOUT);
        hRes = _pJumpList->SetAppID(LPT2LPW(pApp->GetAppID()).c_str());
        assert(SUCCEEDED(hRes));
    }
}
WinTaskbarIcon::~WinTaskbarIcon() {
    if (_pTaskbar) {
        _pTaskbar->Release();
    }
    if (_pJumpList) {
        _pJumpList->Release();
    }
    CoUninitialize();
}

/* ----------- Taskbar Icon State ----------- */

BOOL WinTaskbarIcon::IsProgressState() {
    if (_StateNow == TaskbarState_Progress) return TRUE;
    if (_StateNow == TaskbarState_Error) return TRUE;
    if (_StateNow == TaskbarState_Paused) return TRUE;
    if (_StateNow == TaskbarState_Indeterminate) return TRUE;
    return FALSE;
}

HRESULT WinTaskbarIcon::SetProgressValue(ULONGLONG nProgressValue, ULONGLONG nProgressTotal) {
    assert(_pTaskbar != 0);
    _nStateProgressValue = nProgressValue;
    if (nProgressTotal != -1) _nStateProgressTotal = nProgressTotal;
    if (IsProgressState()) _pTaskbar->SetProgressValue(_hWnd, nProgressValue, nProgressTotal);
    return S_OK;
}
HRESULT WinTaskbarIcon::SetFlashValue(DWORD dwTimeout, UINT nCount) {
    _nFlashCount = nCount;
    _dwFlashTimeout = dwTimeout;
    return S_OK;
}
HRESULT WinTaskbarIcon::SetFlashFlag(DWORD dwFlags) {
    _dwFlashFlags = dwFlags;
    return S_OK;
}
HRESULT WinTaskbarIcon::SetState(WinTaskbarIcon_State State) {
    assert(_pTaskbar != 0);
    TBPFLAG Flag;
    if (State == TaskbarState_Flash) {
        FLASHWINFO Info;
        Info.cbSize = sizeof(FLASHWINFO);
        Info.dwTimeout = _dwFlashTimeout;
        Info.hwnd = _hWnd;
        Info.dwFlags = 0;
        if (_nFlashCount == TASKBAR_FLASH_INFINITY) {
            Info.dwFlags |= FLASHW_TIMERNOFG;
        }
        else {
            Info.dwFlags |= FLASHW_TIMER;
            Info.uCount = _nFlashCount;
        }
        if (_dwFlashFlags & TaskbarFlash_Caption) Info.dwFlags |= FLASHW_CAPTION;
        if (_dwFlashFlags & TaskbarFlash_Tray) Info.dwFlags |= FLASHW_TRAY;
        FlashWindowEx(&Info);
        _StateNow = State;
        return S_OK;
    }
    switch (State) {
    case TaskbarState_Error: { Flag = TBPF_ERROR; break; }
    case TaskbarState_Paused: { Flag = TBPF_PAUSED; break; }
    case TaskbarState_Normal: { Flag = TBPF_NOPROGRESS; break; }
    case TaskbarState_Progress: { Flag = TBPF_NORMAL; break; }
    case TaskbarState_Indeterminate: { Flag = TBPF_INDETERMINATE; break; }
    default: { Flag = TBPF_NOPROGRESS; break; }
    }
    _pTaskbar->SetProgressState(_hWnd, Flag);
    _StateNow = State;
    if (IsProgressState()) _pTaskbar->SetProgressValue(_hWnd, _nStateProgressValue, _nStateProgressTotal);
    return S_OK;
}

/* ------------- Taskbar Button ------------- */

WinTaskbarIcon::TaskbarButton::TaskbarButton()
    : _Button()
{
    ZeroMemory(&_Button, sizeof(_Button));
    _Button.dwFlags = THBF_ENABLED;
    _Button.dwMask = THB_FLAGS;
}
WinTaskbarIcon::TaskbarButton::TaskbarButton(UINT nButtonID)
    : _Button()
{
    ZeroMemory(&_Button, sizeof(_Button));
    _Button.dwFlags = THBF_ENABLED;
    _Button.dwMask = THB_FLAGS;
    SetButtonID(nButtonID);
}
WinTaskbarIcon::TaskbarButton::TaskbarButton(const TaskbarButton& other)
    : _Button(other._Button)
{
    _Button.hIcon = CopyIcon(other._Button.hIcon);
}
WinTaskbarIcon::TaskbarButton::~TaskbarButton() {
    if (_Button.hIcon) {
        DestroyIcon(_Button.hIcon);
        _Button.hIcon = NULL;
    }
}

HRESULT WinTaskbarIcon::TaskbarButton::AddFlag(WinTaskbarButton_Flag Flag) {
    _Button.dwFlags |= (THUMBBUTTONFLAGS)Flag;
#pragma warning(push)
#pragma warning(disable: 26813)
    if (Flag == TaskbarButton_Disabled) RemoveFlag(TaskbarButton_Enabled);
    if (Flag == TaskbarButton_Enabled) RemoveFlag(TaskbarButton_Disabled);
#pragma warning(pop)
    return S_OK;
}
HRESULT WinTaskbarIcon::TaskbarButton::RemoveFlag(WinTaskbarButton_Flag Flag) {
    _Button.dwFlags &= (THUMBBUTTONFLAGS)(~Flag);
    return S_OK;
}
HRESULT WinTaskbarIcon::TaskbarButton::SetImage(HICON hIcon) {
    if (hIcon == NULL) {
        _Button.dwMask &= (~THB_ICON);
        _Button.hIcon = NULL;
    }
    else {
        _Button.dwMask |= THB_ICON;
        _Button.hIcon = CopyIcon(hIcon);
    }
    return S_OK;
}
HRESULT WinTaskbarIcon::TaskbarButton::SetTip(std::_tstring Tip) {
    if (Tip.empty()) {
        _Button.dwMask &= (~THB_TOOLTIP);
        wcscpy_s(_Button.szTip, _countof(_Button.szTip), L"");
    }
    else {
        _Button.dwMask |= THB_TOOLTIP;
        wcscpy_s(_Button.szTip, _countof(_Button.szTip), LPT2LPW(Tip).c_str());
    }
    return S_OK;
}
HRESULT WinTaskbarIcon::TaskbarButton::SetButtonID(UINT nID) {
    _Button.iId = nID;
    return S_OK;
}
THUMBBUTTON* WinTaskbarIcon::TaskbarButton::GetButtonObject() {
    return &_Button;
}

HRESULT WinTaskbarIcon::AddButton(TaskbarButton* pButton, UINT nCount) {
    assert(_pTaskbar != 0);
    if (pButton == NULL) return E_POINTER;
    if (nCount > MAX_TOOLBARBUTTON) return E_INVALIDARG;
    HRESULT hRes = S_OK;
    THUMBBUTTON ButtonsAdd[MAX_TOOLBARBUTTON] = {THUMBBUTTON()};
    for (UINT idx = 0; idx < nCount; idx++) {
        THUMBBUTTON* pNowButton = pButton[idx].GetButtonObject();
        assert(pNowButton != 0);
        ButtonsAdd[idx] = *pNowButton;
    }
    hRes = _pTaskbar->ThumbBarAddButtons(_hWnd, nCount, ButtonsAdd);
    return hRes;
}
HRESULT WinTaskbarIcon::UpdateButton(TaskbarButton* pButton, UINT nCount) {
    assert(_pTaskbar != 0);
    if (pButton == NULL) return E_POINTER;
    if (nCount > MAX_TOOLBARBUTTON) return E_INVALIDARG;
    HRESULT hRes = S_OK;
    THUMBBUTTON ButtonsUpdate[MAX_TOOLBARBUTTON] = { THUMBBUTTON() };
    for (UINT idx = 0; idx < nCount; idx++) {
        THUMBBUTTON* pNowButton = pButton[idx].GetButtonObject();
        assert(pNowButton != 0);
        ButtonsUpdate[idx] = *pNowButton;
    }
    hRes = _pTaskbar->ThumbBarUpdateButtons(_hWnd, nCount, ButtonsUpdate);
    return hRes;
}

/* ------------ Taskbar JumpList ------------ */

WinTaskbarIcon::JumpListShelllink::JumpListShelllink()
    : _pShelllink(NULL)
{
    HRESULT hRes = CoInitialize(NULL);
    assert(SUCCEEDED(hRes) || hRes == RPC_E_CHANGED_MODE);
    hRes = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC, IID_PPV_ARGS(&_pShelllink));
    assert(SUCCEEDED(hRes) && _pShelllink != NULL);
    SetName(TEXT("Untitled"));
}
WinTaskbarIcon::JumpListShelllink::~JumpListShelllink() {
    if (_pShelllink != NULL) {
        _pShelllink->Release();
    }
    CoUninitialize();
}
HRESULT WinTaskbarIcon::JumpListShelllink::SetName(std::_tstring Name) {
    assert(_pShelllink != 0);
    HRESULT hRes;
    IPropertyStore* pPropertyStore = NULL;
    hRes = _pShelllink->QueryInterface(IID_PPV_ARGS(&pPropertyStore));
    if (SUCCEEDED(hRes))
    {
        PROPVARIANT pvTitle;
        hRes = InitPropVariantFromString(LPT2LPW(Name).c_str(), &pvTitle);
        if (SUCCEEDED(hRes))
        {
            hRes = pPropertyStore->SetValue(PKEY_Title, pvTitle);
            if (SUCCEEDED(hRes))
            {
                hRes = pPropertyStore->Commit();
            }
            PropVariantClear(&pvTitle);
        }
        pPropertyStore->Release();
    }
    return hRes;
}
HRESULT WinTaskbarIcon::JumpListShelllink::SetFilePath(std::_tstring FilePath) {
    assert(_pShelllink != 0);
    return _pShelllink->SetPath(FilePath.c_str());
}
HRESULT WinTaskbarIcon::JumpListShelllink::SetArguments(std::_tstring Arguments) {
    assert(_pShelllink != 0);
    return _pShelllink->SetArguments(Arguments.c_str());
}
HRESULT WinTaskbarIcon::JumpListShelllink::SetWorkingDirectory(std::_tstring WorkingDirectory) {
    assert(_pShelllink != 0);
    return _pShelllink->SetWorkingDirectory(WorkingDirectory.c_str());
}
HRESULT WinTaskbarIcon::JumpListShelllink::SetIconPath(std::_tstring IconPath, INT nIndex) {
    assert(_pShelllink != 0);
    return _pShelllink->SetIconLocation(IconPath.c_str(), nIndex);
}
HRESULT WinTaskbarIcon::JumpListShelllink::SetDescription(std::_tstring Description) {
    assert(_pShelllink != 0);
    return _pShelllink->SetDescription(Description.c_str());
}
HRESULT WinTaskbarIcon::JumpListShelllink::SetShowCmd(UINT nShowCmd) {
    assert(_pShelllink != 0);
    return _pShelllink->SetShowCmd(nShowCmd);
}
ComPtr<IShellLink> WinTaskbarIcon::JumpListShelllink::GetShelllinkObject() {
    assert(_pShelllink != 0);
    return _pShelllink;
}

WinTaskbarIcon::JumpListCollection::JumpListCollection() 
    : _pCollection(NULL)
{
    HRESULT hRes = CoInitialize(NULL);
    assert(SUCCEEDED(hRes) || hRes == RPC_E_CHANGED_MODE);
    hRes = CoCreateInstance(CLSID_EnumerableObjectCollection, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&_pCollection));
    assert(SUCCEEDED(hRes) && _pCollection != NULL);
}
WinTaskbarIcon::JumpListCollection::~JumpListCollection() {
    if (_pCollection != NULL) {
        _pCollection->Release();
    }
    CoUninitialize();
}
HRESULT WinTaskbarIcon::JumpListCollection::SetName(std::_tstring Name) {
    assert(_pCollection != 0);
    _Name = Name;
    return S_OK;
}
HRESULT WinTaskbarIcon::JumpListCollection::AddShelllink(JumpListShelllink* pShelllink) {
    assert(_pCollection != 0);
    if (pShelllink == NULL) return E_POINTER;
    return _pCollection->AddObject(pShelllink->GetShelllinkObject().Get());
}
std::_tstring WinTaskbarIcon::JumpListCollection::GetName() {
    return _Name;
}
ComPtr<IObjectCollection> WinTaskbarIcon::JumpListCollection::GetCollectionObject() {
    return _pCollection;
}

UINT WinTaskbarIcon::GetMaxSlots() {
    return _nJumpListMaxSlots;
}
HRESULT WinTaskbarIcon::ResetJumpList() {
    assert(_pJumpList != 0);
    IObjectArray* pRemovedItems = NULL;
    HRESULT hRes = _pJumpList->BeginList(&_nJumpListMaxSlots, IID_PPV_ARGS(&pRemovedItems));
    if (FAILED(hRes) || pRemovedItems == NULL) return hRes;
    pRemovedItems->Release();
    _bJumpListCanAdd = TRUE;
    return S_OK;
}
HRESULT WinTaskbarIcon::AddJumpList(JumpListCollection* pCollection) {
    assert(_pJumpList != 0);
    if (pCollection == NULL) return E_POINTER;
    if (!_bJumpListCanAdd) return E_ACCESSDENIED;
    ComPtr<IObjectCollection> pBaseCollection = pCollection->GetCollectionObject();
    assert(pBaseCollection != 0);
    HRESULT hRes = _pJumpList->AppendCategory(LPT2LPW(pCollection->GetName()).c_str(), pBaseCollection.Get());
    return hRes;
}
HRESULT WinTaskbarIcon::SetUserTaskJumpList(JumpListCollection* pCollection) {
    assert(_pJumpList != 0);
    if (pCollection == NULL) return E_POINTER;
    if (!_bJumpListCanAdd) return E_ACCESSDENIED;
    ComPtr<IObjectCollection> pBaseCollection = pCollection->GetCollectionObject();
    assert(pBaseCollection != 0);
    ComPtr<IObjectArray> pArray;
    pBaseCollection->QueryInterface(IID_PPV_ARGS(&pArray));
    assert(pArray != 0);
    HRESULT hRes = _pJumpList->AddUserTasks(pArray.Get());
    return hRes;
}
HRESULT WinTaskbarIcon::CommitJumpList() {
    assert(_pJumpList != 0);
    HRESULT hRes = _pJumpList->CommitList();
    if (FAILED(hRes)) return hRes;
    _bJumpListCanAdd = FALSE;
    return S_OK;
}

/* ----------------- Other ------------------ */

HRESULT WinTaskbarIcon::ShowTaskbarIcon(BOOL bShow) {
    assert(_pTaskbar != 0);
    if (bShow) return _pTaskbar->AddTab(_hWnd);
    return _pTaskbar->DeleteTab(_hWnd);
}
HRESULT WinTaskbarIcon::SetOverlayIcon(HICON hIcon, std::_tstring Hint) {
    assert(_pTaskbar != 0);
    return _pTaskbar->SetOverlayIcon(_hWnd, hIcon, LPT2LPW(Hint).c_str());
}

/* =================================== WinTaskbarIcon End =================================== */
