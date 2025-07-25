#include <Windows.h>
#include <iostream>
#include <vector>
#include <set>
#include <string>
#include <tuple>
#include <filesystem>
#include <thread>
#include <atomic>
#include <sstream>
#include <CommCtrl.h>

#pragma comment(lib, "comctl32.lib")

namespace fs = std::filesystem;

// Window controls IDs
#define ID_LISTBOX_ACTIONS 1001
#define ID_BUTTON_PREVIOUS 1002
#define ID_BUTTON_NEXT 1003
#define ID_BUTTON_EXECUTE_CURRENT 1004
#define ID_BUTTON_EXECUTE_ALL 1005
#define ID_STATIC_STATUS 1006
#define ID_BUTTON_SCAN_DIRECTORY 1007
#define ID_BUTTON_STOP 1008

// Global variables
HWND g_hWnd;
HWND g_hListBox;
HWND g_hStatusText;
std::atomic<bool> g_monitorRunning(false);
std::atomic<bool> g_executionRunning(false);
std::thread* g_executionThread = nullptr;

// Action structure
struct Action {
    enum Type {
        LAUNCH_APP,
        WAIT_CTRL,
        PRESS_KEY,
        PRESS_TAB,
        SET_CLIPBOARD,
        PASTE_CTRL_V,
        PASTE_CTRL_SHIFT_V,
        PRESS_ENTER,
        PRESS_F11,
        PRESS_F2,
        MOVE_MOUSE,
        LEFT_CLICK,
        RIGHT_CLICK,
        SLEEP
    };

    Type type;
    std::string description;
    std::string data;
    int param1;
    int param2;
};

// Global action list and current position
std::vector<Action> g_actions;
int g_currentActionIndex = 0;

// Function declarations from original code
void MoveMouse(int x, int y);
void LeftClick();
void RightClick();
void MultPressKey(std::string key);
void PressKey(int k);
void SimulateCtrlV();
void PressCtrlShiftV();
void SetClipboardText(const std::string& text);
void ReplaceBackslashesWithForwardSlashes(std::string& path);
void ctrl_wait();
std::wstring string_to_wstring(const std::string& str);
bool window_foreground(std::string window);
void launch_an_app();
std::vector<std::string> personel_number_extractor(std::vector<std::string> FileSets);

// Implementation of original functions
void MoveMouse(int x, int y) {
    SetCursorPos(x, y);
}

void LeftClick() {
    INPUT input = { 0 };
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
    SendInput(1, &input, sizeof(INPUT));

    ZeroMemory(&input, sizeof(INPUT));
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = MOUSEEVENTF_LEFTUP;
    SendInput(1, &input, sizeof(INPUT));
}

void RightClick() {
    INPUT input = { 0 };
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = MOUSEEVENTF_RIGHTDOWN;
    SendInput(1, &input, sizeof(INPUT));

    ZeroMemory(&input, sizeof(INPUT));
    input.type = INPUT_MOUSE;
    input.mi.dwFlags = MOUSEEVENTF_RIGHTUP;
    SendInput(1, &input, sizeof(INPUT));
}

void MultPressKey(std::string key) {
    for (char k : key) {
        INPUT input = { 0 };
        input.type = INPUT_KEYBOARD;
        input.ki.wVk = k;
        input.ki.dwFlags = 0;
        SendInput(1, &input, sizeof(INPUT));

        input.ki.dwFlags = KEYEVENTF_KEYUP;
        SendInput(1, &input, sizeof(INPUT));
    }
}

void PressKey(int k) {
    INPUT input = { 0 };
    input.type = INPUT_KEYBOARD;
    input.ki.wVk = k;
    input.ki.dwFlags = 0;
    SendInput(1, &input, sizeof(INPUT));

    input.ki.dwFlags = KEYEVENTF_KEYUP;
    SendInput(1, &input, sizeof(INPUT));
}

void SimulateCtrlV() {
    keybd_event(VK_CONTROL, 0, 0, 0);
    keybd_event('V', 0, 0, 0);
    keybd_event('V', 0, KEYEVENTF_KEYUP, 0);
    keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
}

void PressCtrlShiftV() {
    INPUT inputs[6] = {};

    inputs[0].type = INPUT_KEYBOARD;
    inputs[0].ki.wVk = VK_CONTROL;

    inputs[1].type = INPUT_KEYBOARD;
    inputs[1].ki.wVk = VK_SHIFT;

    inputs[2].type = INPUT_KEYBOARD;
    inputs[2].ki.wVk = 'V';

    inputs[3].type = INPUT_KEYBOARD;
    inputs[3].ki.wVk = 'V';
    inputs[3].ki.dwFlags = KEYEVENTF_KEYUP;

    inputs[4].type = INPUT_KEYBOARD;
    inputs[4].ki.wVk = VK_SHIFT;
    inputs[4].ki.dwFlags = KEYEVENTF_KEYUP;

    inputs[5].type = INPUT_KEYBOARD;
    inputs[5].ki.wVk = VK_CONTROL;
    inputs[5].ki.dwFlags = KEYEVENTF_KEYUP;

    SendInput(6, inputs, sizeof(INPUT));
}

void SetClipboardText(const std::string& text) {
    if (!OpenClipboard(NULL)) {
        return;
    }

    EmptyClipboard();

    HGLOBAL hGlobal = GlobalAlloc(GMEM_MOVEABLE, (text.size() + 1) * sizeof(char));
    if (!hGlobal) {
        CloseClipboard();
        return;
    }

    char* pGlobal = static_cast<char*>(GlobalLock(hGlobal));
    if (pGlobal) {
        memcpy(pGlobal, text.c_str(), text.size() + 1);
        GlobalUnlock(hGlobal);
        SetClipboardData(CF_TEXT, hGlobal);
    }
    else {
        GlobalFree(hGlobal);
    }

    CloseClipboard();
}

void ReplaceBackslashesWithForwardSlashes(std::string& path) {
    size_t pos = 0;
    while ((pos = path.find('/', pos)) != std::string::npos) {
        path.replace(pos, 1, "\\");
        pos += 1;
    }
}

void ctrl_wait() {
    while (true) {
        if (GetAsyncKeyState(VK_CONTROL) & 0x8000) {
            break;
        }
        Sleep(50);
    }
}

std::wstring string_to_wstring(const std::string& str) {
    int size_needed = MultiByteToWideChar(CP_UTF8, 0, str.c_str(), (int)str.size(), NULL, 0);
    std::wstring wstr(size_needed, 0);
    MultiByteToWideChar(CP_UTF8, 0, str.c_str(), (int)str.size(), &wstr[0], size_needed);
    return wstr;
}

LPCWSTR ConvertToLPCWSTR(const char* charString) {
    int len = MultiByteToWideChar(CP_ACP, 0, charString, -1, NULL, 0);
    wchar_t* wString = new wchar_t[len];
    MultiByteToWideChar(CP_ACP, 0, charString, -1, wString, len);
    return wString;
}

void launch_an_app() {
    const char* exePath = "";

    LPCWSTR wideExePath = ConvertToLPCWSTR(exePath);

    HINSTANCE result = ShellExecute(NULL, L"open", wideExePath, NULL, NULL, SW_SHOW);

    if ((int)result <= 32) {
        SetWindowTextW(g_hStatusText, L"Failed to launch application");
    }
    else {
        SetWindowTextW(g_hStatusText, L"Application launched successfully");
    }

    delete[] wideExePath;
    Sleep(5000);

    HWND hwnd = FindWindow(NULL, L"Enterprise Scan");
    if (hwnd != NULL) {
        SetForegroundWindow(hwnd);
        ShowWindow(hwnd, SW_MAXIMIZE);
    }
}

std::vector<std::string> personel_number_extractor(std::vector<std::string> FileSets) {
    std::vector<std::string> returnvec{};
    for (auto v : FileSets) {
        std::string str{};
        str = str + v[33 + 0] + v[33 + 1] + v[33 + 2] + v[33 + 3] + v[33 + 4];
        returnvec.push_back(str);
    }
    return returnvec;
}

// Execute a single action
void ExecuteAction(const Action& action) {
    std::wstring statusMsg = string_to_wstring("Executing: " + action.description);
    SetWindowTextW(g_hStatusText, statusMsg.c_str());

    switch (action.type) {
    case Action::LAUNCH_APP:
        launch_an_app();
        break;
    case Action::WAIT_CTRL:
        SetWindowTextW(g_hStatusText, L"Waiting for CTRL key press...");
        ctrl_wait();
        break;
    case Action::PRESS_KEY:
        PressKey(action.param1);
        break;
    case Action::PRESS_TAB:
        PressKey(VK_TAB);
        Sleep(100);
        break;
    case Action::SET_CLIPBOARD:
        SetClipboardText(action.data);
        break;
    case Action::PASTE_CTRL_V:
        SimulateCtrlV();
        break;
    case Action::PASTE_CTRL_SHIFT_V:
        PressCtrlShiftV();
        break;
    case Action::PRESS_ENTER:
        PressKey(VK_RETURN);
        break;
    case Action::PRESS_F11:
        PressKey(VK_F11);
        break;
    case Action::PRESS_F2:
        PressKey(VK_F2);
        break;
    case Action::MOVE_MOUSE:
        MoveMouse(action.param1, action.param2);
        break;
    case Action::LEFT_CLICK:
        LeftClick();
        break;
    case Action::RIGHT_CLICK:
        RightClick();
        break;
    case Action::SLEEP:
        Sleep(action.param1);
        break;
    }
}

// Update the listbox to highlight current action
void UpdateListBox() {
    SendMessage(g_hListBox, LB_SETCURSEL, g_currentActionIndex, 0);

    // Ensure the current item is visible
    SendMessage(g_hListBox, LB_SETTOPINDEX, max(0, g_currentActionIndex - 5), 0);
}

// Build action list from directory scan
void BuildActionList(const std::string& baseDir) {
    g_actions.clear();
    g_currentActionIndex = 0;

    std::vector<std::set<std::string>> fileSets;
    std::vector<std::string> personelnumbers;

    // Add initial actions
    g_actions.push_back({ Action::LAUNCH_APP, "Launch Enterprise Scan Application", "", 0, 0 });
    g_actions.push_back({ Action::WAIT_CTRL, "Wait for CTRL key press", "", 0, 0 });
    g_actions.push_back({ Action::SLEEP, "Sleep 1000ms", "", 1000, 0 });
    g_actions.push_back({ Action::PRESS_TAB, "Press TAB", "", 0, 0 });
    g_actions.push_back({ Action::PRESS_TAB, "Press TAB", "", 0, 0 });
    g_actions.push_back({ Action::PRESS_TAB, "Press TAB", "", 0, 0 });

    // Scan directory
    for (const auto& entry : fs::directory_iterator(baseDir)) {
        if (fs::is_directory(entry.path())) {
            std::set<std::string> filePaths;

            for (const auto& fileEntry : fs::directory_iterator(entry.path())) {
                if (fs::is_regular_file(fileEntry.path())) {
                    std::string filename = fileEntry.path().string();

                    // Check file extensions using character comparisons
                    bool skipFile = false;
                    size_t len = filename.length();

                    if (len >= 8 && filename.substr(len - 8) == "COMMANDS") skipFile = true;
                    else if (len >= 4 && filename.substr(len - 4) == ".BIN") skipFile = true;
                    else if (len >= 3 && filename.substr(len - 3) == "log") skipFile = true;
                    else if (len >= 4 && filename.substr(len - 4) == "xlsx") skipFile = true;

                    if (!skipFile) {
                        filePaths.insert(filename);
                    }
                }
            }

            if (!filePaths.empty()) {
                fileSets.push_back(filePaths);
                personelnumbers.push_back(entry.path().string());
            }
        }
    }

    personelnumbers = personel_number_extractor(personelnumbers);

    // Build action list for each personnel number
    int j = 0;
    for (auto entry : personelnumbers) {
        g_actions.push_back({ Action::SET_CLIPBOARD, "Set clipboard: Personnel " + entry, entry, 0, 0 });
        g_actions.push_back({ Action::WAIT_CTRL, "Wait for CTRL key press", "", 0, 0 });
        g_actions.push_back({ Action::SLEEP, "Sleep 1000ms", "", 1000, 0 });
        g_actions.push_back({ Action::PASTE_CTRL_SHIFT_V, "Paste personnel number (Ctrl+Shift+V)", "", 0, 0 });
        g_actions.push_back({ Action::SLEEP, "Sleep 1000ms", "", 1000, 0 });
        g_actions.push_back({ Action::PRESS_ENTER, "Press ENTER", "", 0, 0 });

        for (auto path : fileSets[j]) {
            g_actions.push_back({ Action::WAIT_CTRL, "Wait for CTRL key press", "", 0, 0 });
            g_actions.push_back({ Action::SLEEP, "Sleep 1000ms", "", 1000, 0 });
            g_actions.push_back({ Action::PRESS_F11, "Press F11", "", 0, 0 });

            std::string modPath = path;
            ReplaceBackslashesWithForwardSlashes(modPath);
            g_actions.push_back({ Action::SET_CLIPBOARD, "Set clipboard: " + modPath, modPath, 0, 0 });
            g_actions.push_back({ Action::SLEEP, "Sleep 100ms", "", 100, 0 });
            g_actions.push_back({ Action::WAIT_CTRL, "Wait for CTRL key press", "", 0, 0 });
            g_actions.push_back({ Action::PASTE_CTRL_V, "Paste file path (Ctrl+V)", "", 0, 0 });
            g_actions.push_back({ Action::SLEEP, "Sleep 1000ms", "", 1000, 0 });
            g_actions.push_back({ Action::PRESS_ENTER, "Press ENTER", "", 0, 0 });
        }

        g_actions.push_back({ Action::WAIT_CTRL, "Wait for CTRL key press", "", 0, 0 });
        g_actions.push_back({ Action::SLEEP, "Sleep 1000ms", "", 1000, 0 });
        g_actions.push_back({ Action::PRESS_F2, "Press F2", "", 0, 0 });
        g_actions.push_back({ Action::WAIT_CTRL, "Wait for CTRL key press", "", 0, 0 });
        g_actions.push_back({ Action::SLEEP, "Sleep 1000ms", "", 1000, 0 });
        g_actions.push_back({ Action::PRESS_ENTER, "Press ENTER", "", 0, 0 });
        j++;
    }

    // Populate listbox
    SendMessage(g_hListBox, LB_RESETCONTENT, 0, 0);
    for (size_t i = 0; i < g_actions.size(); i++) {
        std::string item = std::to_string(i + 1) + ". " + g_actions[i].description;
        std::wstring witem = string_to_wstring(item);
        SendMessage(g_hListBox, LB_ADDSTRING, 0, (LPARAM)witem.c_str());
    }

    UpdateListBox();
}

// Execute all actions from current position
void ExecuteAllFromCurrent() {
    g_executionRunning = true;

    for (int i = g_currentActionIndex; i < g_actions.size() && g_executionRunning; i++) {
        g_currentActionIndex = i;

        // Update UI in main thread
        PostMessage(g_hWnd, WM_USER + 1, 0, 0);

        ExecuteAction(g_actions[i]);

        if (!g_executionRunning) break;
    }

    g_executionRunning = false;
    SetWindowTextW(g_hStatusText, L"Execution completed");
}

// Window procedure
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam) {
    switch (uMsg) {
    case WM_CREATE:
        // Create controls
        g_hListBox = CreateWindowW(L"LISTBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | LBS_NOTIFY,
            10, 10, 600, 400, hwnd, (HMENU)ID_LISTBOX_ACTIONS, NULL, NULL);

        CreateWindowW(L"BUTTON", L"Scan Directory",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            620, 10, 150, 30, hwnd, (HMENU)ID_BUTTON_SCAN_DIRECTORY, NULL, NULL);

        CreateWindowW(L"BUTTON", L"Previous",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            620, 50, 150, 30, hwnd, (HMENU)ID_BUTTON_PREVIOUS, NULL, NULL);

        CreateWindowW(L"BUTTON", L"Next",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            620, 90, 150, 30, hwnd, (HMENU)ID_BUTTON_NEXT, NULL, NULL);

        CreateWindowW(L"BUTTON", L"Execute Current",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            620, 130, 150, 30, hwnd, (HMENU)ID_BUTTON_EXECUTE_CURRENT, NULL, NULL);

        CreateWindowW(L"BUTTON", L"Execute All",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            620, 170, 150, 30, hwnd, (HMENU)ID_BUTTON_EXECUTE_ALL, NULL, NULL);

        CreateWindowW(L"BUTTON", L"Stop Execution",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            620, 210, 150, 30, hwnd, (HMENU)ID_BUTTON_STOP, NULL, NULL);

        g_hStatusText = CreateWindowW(L"STATIC", L"Ready",
            WS_CHILD | WS_VISIBLE | WS_BORDER,
            10, 420, 760, 40, hwnd, (HMENU)ID_STATIC_STATUS, NULL, NULL);

        // Initialize with default directory
        BuildActionList("");
        break;

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_BUTTON_SCAN_DIRECTORY: {
            // Simple directory input dialog
            char dirPath[MAX_PATH] = "";
            if (MessageBoxW(hwnd, L"Using default directory. Click OK to continue.", L"Directory", MB_OKCANCEL) == IDOK) {
                BuildActionList(dirPath);
                SetWindowTextW(g_hStatusText, L"Directory scanned and actions loaded");
            }
            break;
        }

        case ID_BUTTON_PREVIOUS:
            if (g_currentActionIndex > 0) {
                g_currentActionIndex--;
                UpdateListBox();
            }
            break;

        case ID_BUTTON_NEXT:
            if (g_currentActionIndex < g_actions.size() - 1) {
                g_currentActionIndex++;
                UpdateListBox();
            }
            break;

        case ID_BUTTON_EXECUTE_CURRENT:
            if (g_currentActionIndex < g_actions.size()) {
                ExecuteAction(g_actions[g_currentActionIndex]);
            }
            break;

        case ID_BUTTON_EXECUTE_ALL:
            if (!g_executionRunning && g_executionThread == nullptr) {
                g_executionThread = new std::thread(ExecuteAllFromCurrent);
                g_executionThread->detach();
                delete g_executionThread;
                g_executionThread = nullptr;
            }
            break;

        case ID_BUTTON_STOP:
            g_executionRunning = false;
            SetWindowTextW(g_hStatusText, L"Execution stopped");
            break;

        case ID_LISTBOX_ACTIONS:
            if (HIWORD(wParam) == LBN_SELCHANGE) {
                int sel = SendMessage(g_hListBox, LB_GETCURSEL, 0, 0);
                if (sel != LB_ERR) {
                    g_currentActionIndex = sel;
                }
            }
            break;
        }
        break;

    case WM_USER + 1:
        // Update UI from execution thread
        UpdateListBox();
        break;

    case WM_DESTROY:
        g_executionRunning = false;
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProc(hwnd, uMsg, wParam, lParam);
    }
    return 0;
}

// Main function
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    // Initialize common controls
    INITCOMMONCONTROLSEX icex;
    icex.dwSize = sizeof(INITCOMMONCONTROLSEX);
    icex.dwICC = ICC_LISTVIEW_CLASSES;
    InitCommonControlsEx(&icex);

    // Register window class
    const wchar_t* CLASS_NAME = L"GuiAutomationController";

    WNDCLASSW wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.hCursor = LoadCursor(NULL, IDC_ARROW);

    RegisterClassW(&wc);

    // Create window
    g_hWnd = CreateWindowExW(
        0,
        CLASS_NAME,
        L"GUI Automation Controller - OpenText Enterprise Scan",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 500,
        NULL, NULL, hInstance, NULL
    );

    if (g_hWnd == NULL) {
        return 0;
    }

    ShowWindow(g_hWnd, nCmdShow);
    UpdateWindow(g_hWnd);

    // Message loop
    MSG msg = {};
    while (GetMessage(&msg, NULL, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    return 0;
}