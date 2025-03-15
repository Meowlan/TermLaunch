#include <windows.h>
#include <shlobj.h>
#include <exdisp.h>
#include <shlguid.h>
#include <iostream>
#include <string>
#include <memory>
#include <UIAutomation.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "UIAutomationCore.lib")

#define ID_TRAY_APP_ICON 4368
#define WM_TRAYICON (WM_USER + 1)

NOTIFYICONDATA nid;

// define our hotkey id
#define HOTKEY_ID 1

// raii wrapper for com
template<typename T>
class ComPtr {
private:
    T* ptr;
public:
    ComPtr() : ptr(nullptr) {}
    ~ComPtr() { if (ptr) ptr->Release(); }
    T** operator&() { return &ptr; }
    T* operator->() { return ptr; }
    operator bool() const { return ptr != nullptr; }
    T* Get() const { return ptr; }
};

enum class WindowType {
    Unknown,
    FileExplorer,
    Desktop,
    Other
};

struct WindowInfo {
    HWND hwnd = nullptr;
    WindowType type = WindowType::Unknown;
    std::wstring className;
    std::wstring title;
};

// helper function to get user profile path
std::wstring GetUserProfilePath() {
    char* userProfile = nullptr;
    size_t size = 0;

    if (_dupenv_s(&userProfile, &size, "USERPROFILE") == 0 && userProfile) {
        int bufferSize = MultiByteToWideChar(CP_ACP, 0, userProfile, -1, nullptr, 0);
        std::wstring userProfileW(bufferSize, 0);
        MultiByteToWideChar(CP_ACP, 0, userProfile, -1, &userProfileW[0], bufferSize);

        free(userProfile);

        if (!userProfileW.empty() && userProfileW.back() == L'\0') {
            userProfileW.pop_back();
        }

        return userProfileW;
    }

    return L"";
}

// get the active path from the focused file explorer window
std::wstring GetActiveExplorerPath(HWND explorerHwnd) {
    HRESULT hr = CoInitialize(nullptr);
    if (FAILED(hr)) {
        return L"";
    }

    std::wstring path;

    std::cout << "Falling back to shell approach" << std::endl;

    if (path.empty()) {
        HWND foregroundWindow = GetForegroundWindow();

        ComPtr<IShellWindows> shellWindows;
        hr = CoCreateInstance(CLSID_ShellWindows, nullptr, CLSCTX_ALL, IID_IShellWindows, (void**)&shellWindows);
        if (FAILED(hr) || !shellWindows) return L"";

        long count = 0;
        shellWindows->get_Count(&count);

        for (long i = 0; i < count; i++) {
            VARIANT variant;
            VariantInit(&variant);
            variant.vt = VT_I4;
            variant.lVal = i;

            ComPtr<IDispatch> dispatch;
            hr = shellWindows->Item(variant, &dispatch);
            if (FAILED(hr) || !dispatch) continue;

            ComPtr<IWebBrowserApp> webBrowserApp;
            hr = dispatch->QueryInterface(IID_IWebBrowserApp, (void**)&webBrowserApp);
            if (FAILED(hr) || !webBrowserApp) continue;

            HWND hwnd;
            webBrowserApp->get_HWND((LONG_PTR*)&hwnd);
            if (hwnd != explorerHwnd && hwnd != foregroundWindow) continue;

            ComPtr<IServiceProvider> serviceProvider;
            hr = dispatch->QueryInterface(IID_IServiceProvider, (void**)&serviceProvider);
            if (FAILED(hr) || !serviceProvider) continue;

            ComPtr<IShellBrowser> shellBrowser;
            hr = serviceProvider->QueryService(SID_STopLevelBrowser, IID_IShellBrowser, (void**)&shellBrowser);
            if (FAILED(hr) || !shellBrowser) continue;

            ComPtr<IShellView> shellView;
            hr = shellBrowser->QueryActiveShellView(&shellView);
            if (FAILED(hr) || !shellView) continue;

            ComPtr<IFolderView> folderView;
            hr = shellView->QueryInterface(IID_IFolderView, (void**)&folderView);
            if (FAILED(hr) || !folderView) continue;

            ComPtr<IPersistFolder2> persistFolder;
            hr = folderView->GetFolder(IID_IPersistFolder2, (void**)&persistFolder);
            if (FAILED(hr) || !persistFolder) continue;

            LPITEMIDLIST pidl = nullptr;
            hr = persistFolder->GetCurFolder(&pidl);
            if (FAILED(hr) || !pidl) continue;

            wchar_t folderPath[MAX_PATH];
            if (SHGetPathFromIDListW(pidl, folderPath)) {
                path = folderPath;
            }
            CoTaskMemFree(pidl);

            if (!path.empty()) break;
        }
    }

    CoUninitialize();
    return path;
}

// get the desktop path
std::wstring GetDesktopPath() {
    wchar_t path[MAX_PATH];
    if (SUCCEEDED(SHGetFolderPathW(NULL, CSIDL_DESKTOP, NULL, 0, path))) {
        return path;
    }
    return GetUserProfilePath() + L"\\Desktop";
}

// get information about the foreground window
WindowInfo GetForegroundWindowInfo() {
    WindowInfo info;
    info.hwnd = GetForegroundWindow();

    if (info.hwnd) {
        // get window class name
        wchar_t className[256] = { 0 };
        GetClassNameW(info.hwnd, className, sizeof(className) / sizeof(wchar_t));
        info.className = className;

        // get window title
        wchar_t title[256] = { 0 };
        GetWindowTextW(info.hwnd, title, sizeof(title) / sizeof(wchar_t));
        info.title = title;

        // determine window type
        if (info.className == L"CabinetWClass" ||
            info.className == L"Microsoft.UI.Content.DesktopChildSiteBridge") {
            info.type = WindowType::FileExplorer;
        }
        else if (info.className == L"WorkerW" || info.className == L"Progman") {
            info.type = WindowType::Desktop;
        }
        else {
            info.type = WindowType::Other;
        }
    }

    return info;
}

// get the path for the current focused window
std::wstring GetFocusedWindowPath() {
    WindowInfo info = GetForegroundWindowInfo();
    std::wstring path;

    switch (info.type) {
    case WindowType::FileExplorer:
        path = GetActiveExplorerPath(info.hwnd);
        break;

    case WindowType::Desktop:
        path = GetDesktopPath();
        break;

    case WindowType::Other:
    case WindowType::Unknown:
    default:
        path = L"";
        break;
    }

    if (path.empty()) {
        path = GetUserProfilePath();
    }

    return path;
}

// create the tray icon
void CreateTrayIcon(HWND hwnd) {
    memset(&nid, 0, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = hwnd;
    nid.uID = ID_TRAY_APP_ICON;
    nid.uFlags = NIF_ICON | NIF_MESSAGE | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon = LoadIcon(NULL, IDI_APPLICATION);
    wcscpy_s(nid.szTip, L"Term Launch");

    Shell_NotifyIcon(NIM_ADD, &nid);
}

void RemoveTrayIcon() {
    Shell_NotifyIcon(NIM_DELETE, &nid);
}

void ShowContextMenu(HWND hwnd) {
    HMENU hMenu = CreatePopupMenu();
    AppendMenu(hMenu, MF_STRING || MF_DISABLED, 1, L"TermLaunch v2");
    AppendMenu(hMenu, MF_SEPARATOR, 2, L"seperator");
    AppendMenu(hMenu, MF_STRING, 3, L"Exit");

    POINT pt;
    GetCursorPos(&pt);
    SetForegroundWindow(hwnd);
    TrackPopupMenu(hMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, pt.x, pt.y, 0, hwnd, NULL);
    DestroyMenu(hMenu);
}

// message handling function
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_HOTKEY:
        if (wParam == HOTKEY_ID) {
            WindowInfo info = GetForegroundWindowInfo();
            std::wstring path = GetFocusedWindowPath();

            std::wcout << L"\n---------------------------------------" << std::endl;
            std::wcout << L"Hotkey pressed! Window info:" << std::endl;
            std::wcout << L"  Class: " << info.className << std::endl;
            std::wcout << L"  Title: " << info.title << std::endl;

            if (info.type == WindowType::FileExplorer) {
                std::wcout << L"  Type: File Explorer" << std::endl;
            }
            else if (info.type == WindowType::Desktop) {
                std::wcout << L"  Type: Desktop" << std::endl;
            }
            else {
                std::wcout << L"  Type: Other" << std::endl;
            }

            if (!path.empty()) {
                std::wcout << L"Path: " << path << std::endl;
            }
            else {
                std::wcout << L"No path available for this window." << std::endl;
            }
            std::wcout << L"---------------------------------------" << std::endl;

            STARTUPINFO si;
            PROCESS_INFORMATION pi;
            ZeroMemory(&si, sizeof(si));
            si.cb = sizeof(si);
            ZeroMemory(&pi, sizeof(pi));

            // open the terminal here
            if (!CreateProcess(L"C:\\Windows\\System32\\cmd.exe", NULL, NULL, NULL, FALSE, CREATE_NEW_CONSOLE, NULL, path.c_str(), &si, &pi)) {
                std::cerr << "Failed to create process: " << GetLastError() << std::endl;
                return 1;
            }

            CloseHandle(pi.hProcess);
            CloseHandle(pi.hThread);

            return 0;
        }
        break;

    case WM_TRAYICON:
        switch (lParam) {
        case WM_LBUTTONUP:
        case WM_RBUTTONUP:
            ShowContextMenu(hwnd);
            break;
        }

        break;

    case WM_COMMAND:
        if (LOWORD(wParam) == 3) {
            PostQuitMessage(0);
        }
        break;

    case WM_DESTROY:
        RemoveTrayIcon();
        PostQuitMessage(0);
        return 0;
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // register the window class
    const wchar_t CLASS_NAME[] = L"TermLaunchWindow";

    WNDCLASS wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = GetModuleHandle(NULL);
    wc.lpszClassName = CLASS_NAME;

    RegisterClass(&wc);

    // create a message-only window (invisible)
    HWND hwnd = CreateWindowEx(
        0,                          // optional window styles
        CLASS_NAME,                 // window class
        L"Term Launch",            // window title
        WS_OVERLAPPEDWINDOW,        // window style

        // size and position
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,

        NULL,       // parent window    
        NULL,       // menu
        GetModuleHandle(NULL),  // instance handle
        NULL        // additional application data
    );

    if (hwnd == NULL) {
        std::cerr << "Failed to create window" << std::endl;
        return 1;
    }

    // hide the window - we just need it for hotkey processing
    ShowWindow(hwnd, SW_HIDE);
    CreateTrayIcon(hwnd);

    // register the hotkey (ctrl+alt+t)
    if (!RegisterHotKey(hwnd, HOTKEY_ID, MOD_CONTROL | MOD_ALT, 'T')) {
        std::cerr << "Failed to register hotkey" << std::endl;
        return 1;
    }

    std::cout << "Term Launch is running!" << std::endl;
    std::cout << "Press Ctrl+Alt+T to get the path of the focused window." << std::endl;
    std::cout << "Press Ctrl+C in this console window to exit." << std::endl;

    // message loop
    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    // unregister the hotkey when done
    UnregisterHotKey(hwnd, HOTKEY_ID);
    RemoveTrayIcon();

    return 0;
}