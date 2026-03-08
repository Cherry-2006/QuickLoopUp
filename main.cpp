#include <windows.h>
#include <uiautomation.h>
#include <winhttp.h>
#include <string>
#include <vector>
#include <cmath>
#include <cwctype>
#include <thread>
#include <WebView2.h>
#include <ShellScalingApi.h>
#include <dwmapi.h>
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "user32.lib")
#pragma comment(lib, "gdi32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shcore.lib")
#pragma comment(lib, "dwmapi.lib")

#define WM_APP_TRIGGER_EXTRACT (WM_APP + 1)
#define WM_APP_WEBVIEW_READY   (WM_APP + 2)
#define WM_APP_SHOW_CONTENT    (WM_APP + 3)
#define WM_APP_HIDE_POPUP      (WM_APP + 4)

HHOOK hMouseHook = NULL, hKeyboardHook = NULL;
DWORD mainThreadId = 0;
ULONGLONG downTime = 0;
POINT downPos = {0, 0};
HWND hPopupWindow = NULL;
bool popupVisible = false, isDragging = false;
ICoreWebView2Controller* webViewController = nullptr;
ICoreWebView2* webView = nullptr;
bool isWebViewReady = false;
HINTERNET hPersistentSession = NULL;

struct PopupData { int x, y; std::wstring htmlContent; };

LRESULT CALLBACK LowLevelMouseProc(int, WPARAM, LPARAM);
LRESULT CALLBACK LowLevelKeyboardProc(int, WPARAM, LPARAM);
LRESULT CALLBACK PopupWindowProc(HWND, UINT, WPARAM, LPARAM);
void ExtractTextAndSearch(int x, int y, bool wasDragging);
void FetchDefinitionAsync(std::wstring word, int x, int y);
void InitializeWebView();
void CreatePopupWindow(HINSTANCE);
void ShowPopup(int x, int y, const std::wstring& html);
std::wstring BackupClipboardExtraction(int x, int y, bool wasDragging);

// ── CSS ───────────────────────────────────────────────────
static const char* CSS =
    "*{margin:0;padding:0;box-sizing:border-box}"
    "body{background:#111;color:#e0e0e0;font-family:'Segoe UI',-apple-system,sans-serif;margin:0;padding:0;"
    "user-select:text;overflow-y:auto;overflow-x:hidden;cursor:default;border-radius:12px;animation:fi .2s ease both}"
    "@keyframes fi{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:translateY(0)}}"
    "::-webkit-scrollbar{width:5px}::-webkit-scrollbar-track{background:0 0}"
    "::-webkit-scrollbar-thumb{background:rgba(255,255,255,.15);border-radius:10px}"
    ".c{padding:18px 20px}.wh{display:flex;align-items:baseline;gap:10px;margin-bottom:4px}"
    ".wt{font-size:24px;font-weight:700;color:#fff}.ph{color:#888;font-size:15px;font-style:italic}"
    ".pb{display:inline-block;font-size:11px;font-weight:600;color:#aaa;background:rgba(255,255,255,.06);"
    "border:1px solid rgba(255,255,255,.12);padding:3px 10px;border-radius:20px;text-transform:uppercase;"
    "letter-spacing:.8px;margin-bottom:12px}"
    ".dc{background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);border-radius:10px;"
    "padding:12px 14px;margin-bottom:8px;transition:background .15s}.dc:hover{background:rgba(255,255,255,.07)}"
    ".dn{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;"
    "background:#333;color:#fff;font-size:11px;font-weight:700;margin-right:10px;flex-shrink:0;vertical-align:top;margin-top:2px}"
    ".dt{color:#ccc;font-size:15px;line-height:1.55}.de{color:#777;font-size:13px;font-style:italic;margin-top:6px;padding-left:30px}"
    ".sd{height:1px;margin:14px 0;background:rgba(255,255,255,.1)}.ss{margin-top:4px}"
    ".sl{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px;color:#888;margin-bottom:8px}"
    ".sc{display:flex;flex-wrap:wrap;gap:6px}"
    ".sp{font-size:12px;color:#bbb;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.12);"
    "border-radius:16px;padding:4px 12px;transition:all .15s;cursor:default}.sp:hover{background:rgba(255,255,255,.1);color:#fff}"
    ".fb{margin-top:14px;padding-top:12px;border-top:1px solid rgba(255,255,255,.08);display:flex;justify-content:center;gap:8px}"
    ".ft{display:flex;align-items:center;gap:6px;padding:8px 16px;font-size:12px;font-weight:600;color:#999;"
    "border-radius:8px;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.08);"
    "text-decoration:none;transition:all .15s;cursor:pointer}.ft:hover{background:rgba(255,255,255,.1);color:#fff}";

// ── HTML helpers ──────────────────────────────────────────
std::wstring MakeSimplePage(const std::string& body) {
    std::string h = "<html><head><style>" + std::string(CSS) + "</style></head><body>" + body + "</body></html>";
    int len = MultiByteToWideChar(CP_UTF8, 0, h.c_str(), -1, NULL, 0);
    std::wstring r(len, 0); MultiByteToWideChar(CP_UTF8, 0, h.c_str(), -1, &r[0], len); return r;
}

std::wstring GetLoadingHtml(const std::wstring& word = L"") {
    std::string w(word.begin(), word.end());
    std::string wd = w.empty() ? "" : "<div style='margin-top:14px;font-size:14px;color:#777'>Looking up <span style='color:#fff;font-weight:600'>"+ w +"</span></div>";
    return MakeSimplePage(
        "<div style='display:flex;flex-direction:column;justify-content:center;align-items:center;height:100vh'>"
        "<div style='width:32px;height:32px;border-radius:50%;border:3px solid rgba(255,255,255,.08);"
        "border-top:3px solid #fff;animation:s .8s linear infinite'></div>" + wd +
        "</div><style>@keyframes s{to{transform:rotate(360deg)}}</style>");
}

std::string HtmlEsc(const std::string& s) {
    std::string r; r.reserve(s.size());
    for (char c : s) { if(c=='&')r+="&amp;";else if(c=='<')r+="&lt;";else if(c=='>')r+="&gt;";else if(c=='"')r+="&quot;";else r+=c; }
    return r;
}

// ── JSON helpers ──────────────────────────────────────────
std::string ExtractJsonStr(const std::string& j, size_t p) {
    if (p >= j.size() || j[p] != '"') return "";
    std::string r; size_t i = p + 1;
    while (i < j.size()) {
        if (j[i]=='\\' && i+1<j.size()) { r += (j[i+1]=='"'?'"':j[i+1]); i+=2; }
        else if (j[i]=='"') break;
        else { r += j[i]; i++; }
    } return r;
}

std::string FindJsonVal(const std::string& j, const std::string& key, size_t from=0) {
    std::string n = "\""+key+"\":\""; size_t p = j.find(n, from);
    return p==std::string::npos ? "" : ExtractJsonStr(j, p+n.size()-1);
}

std::vector<std::string> FindAllJsonVals(const std::string& j, const std::string& key, int max=10) {
    std::vector<std::string> r; std::string n="\""+key+"\":\""; size_t p=0;
    while ((int)r.size()<max) { p=j.find(n,p); if(p==std::string::npos)break; auto v=ExtractJsonStr(j,p+n.size()-1); if(!v.empty())r.push_back(v); p+=n.size(); }
    return r;
}

std::vector<std::string> FindJsonArr(const std::string& j, const std::string& key) {
    std::vector<std::string> r; std::string n="\""+key+"\":["; size_t p=j.find(n);
    if(p==std::string::npos) return r;
    size_t s=p+n.size(), e=j.find("]",s); if(e==std::string::npos)return r;
    std::string a=j.substr(s,e-s); size_t i=0;
    while(i<a.size()&&r.size()<8) { size_t q=a.find('"',i); if(q==std::string::npos)break;
        auto v=ExtractJsonStr(a,q); if(!v.empty())r.push_back(v);
        size_t qe=a.find('"',q+1); while(qe!=std::string::npos&&qe>0&&a[qe-1]=='\\')qe=a.find('"',qe+1);
        i=(qe==std::string::npos)?a.size():qe+1; }
    return r;
}

// ── WinMain ───────────────────────────────────────────────
int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE, LPSTR, int) {
    SetProcessDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);
    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    mainThreadId = GetCurrentThreadId();

    hPersistentSession = WinHttpOpen(L"QuickLoopUp/2.0", WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, NULL, NULL, 0);
    if (hPersistentSession) WinHttpSetTimeouts(hPersistentSession, 3000, 3000, 5000, 5000);

    CreatePopupWindow(hInstance);
    InitializeWebView();
    hMouseHook = SetWindowsHookExW(WH_MOUSE_LL, LowLevelMouseProc, hInstance, 0);
    hKeyboardHook = SetWindowsHookExW(WH_KEYBOARD_LL, LowLevelKeyboardProc, hInstance, 0);
    if (!hMouseHook || !hKeyboardHook) { MessageBoxW(NULL, L"Failed to install hooks!", L"Error", MB_ICONERROR); return 1; }
    SetPriorityClass(GetCurrentProcess(), REALTIME_PRIORITY_CLASS);

    MSG msg;
    while (GetMessageW(&msg, NULL, 0, 0)) {
        if (msg.message == WM_APP_TRIGGER_EXTRACT) {
            bool drg = isDragging;
            std::thread(ExtractTextAndSearch, (int)msg.wParam, (int)msg.lParam, drg).detach();
        } else if (msg.message == WM_APP_SHOW_CONTENT) {
            PopupData* p = (PopupData*)msg.lParam;
            if (p) { ShowPopup(p->x, p->y, p->htmlContent); delete p; }
        } else if (msg.message == WM_APP_HIDE_POPUP) {
            ShowWindow(hPopupWindow, SW_HIDE);
            if (webViewController) { webViewController->put_IsVisible(FALSE); if (webView) webView->Navigate(L"about:blank"); }
            popupVisible = false;
        } else if (msg.message == WM_APP_WEBVIEW_READY) {
            isWebViewReady = true;
            ShowPopup(-1, -1, MakeSimplePage(
                "<div style='display:flex;flex-direction:column;justify-content:center;align-items:center;height:100vh'>"
                "<div style='font-size:22px;font-weight:700;color:#fff'>QuickLoopUp</div>"
                "<div style='font-size:13px;color:#777;margin-top:8px'>Ready</div></div>"));
            SetTimer(hPopupWindow, 1, 1500, NULL);
        } else { TranslateMessage(&msg); DispatchMessageW(&msg); }
    }
    UnhookWindowsHookEx(hMouseHook); UnhookWindowsHookEx(hKeyboardHook);
    if (hPersistentSession) WinHttpCloseHandle(hPersistentSession);
    CoUninitialize(); return 0;
}

// ── WebView2 COM ──────────────────────────────────────────
class ControllerHandler : public ICoreWebView2CreateCoreWebView2ControllerCompletedHandler {
    ULONG m_ref = 1;
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override {
        if (riid==IID_IUnknown||riid==IID_ICoreWebView2CreateCoreWebView2ControllerCompletedHandler)
        { *ppv=static_cast<ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*>(this); AddRef(); return S_OK; }
        *ppv=nullptr; return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&m_ref); }
    ULONG STDMETHODCALLTYPE Release() override { ULONG r=InterlockedDecrement(&m_ref); if(!r)delete this; return r; }
    HRESULT STDMETHODCALLTYPE Invoke(HRESULT hr, ICoreWebView2Controller* ctrl) override {
        if (FAILED(hr)||!ctrl) return S_OK;
        webViewController = ctrl; webViewController->AddRef();
        webViewController->get_CoreWebView2(&webView);
        static const GUID IID_C2 = {0xC979903E,0xD4CA,0x4228,{0x92,0xEB,0xCE,0x99,0x9B,0x31,0xAB,0x71}};
        ICoreWebView2Controller2* c2 = nullptr;
        if (SUCCEEDED(webViewController->QueryInterface(IID_C2,(void**)&c2))) {
            COREWEBVIEW2_COLOR clr={0,0,0,0}; c2->put_DefaultBackgroundColor(clr); c2->Release();
        }
        ICoreWebView2Settings* s; if (SUCCEEDED(webView->get_Settings(&s))) {
            s->put_IsScriptEnabled(TRUE); s->put_AreDefaultContextMenusEnabled(FALSE); s->put_IsStatusBarEnabled(FALSE); s->Release();
        }
        webViewController->put_IsVisible(TRUE);
        RECT b; GetClientRect(hPopupWindow, &b); webViewController->put_Bounds(b);
        webView->NavigateToString(L"<html><body></body></html>");
        PostThreadMessageW(mainThreadId, WM_APP_WEBVIEW_READY, 0, 0);
        return S_OK;
    }
};

class EnvHandler : public ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler {
    ULONG m_ref = 1;
public:
    HRESULT STDMETHODCALLTYPE QueryInterface(REFIID riid, void** ppv) override {
        if (riid==IID_IUnknown||riid==IID_ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler)
        { *ppv=static_cast<ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*>(this); AddRef(); return S_OK; }
        *ppv=nullptr; return E_NOINTERFACE;
    }
    ULONG STDMETHODCALLTYPE AddRef() override { return InterlockedIncrement(&m_ref); }
    ULONG STDMETHODCALLTYPE Release() override { ULONG r=InterlockedDecrement(&m_ref); if(!r)delete this; return r; }
    HRESULT STDMETHODCALLTYPE Invoke(HRESULT hr, ICoreWebView2Environment* env) override {
        if (FAILED(hr)||!env) return S_OK;
        env->CreateCoreWebView2Controller(hPopupWindow, new ControllerHandler()); return S_OK;
    }
};

typedef HRESULT(*CreateEnvFn)(PCWSTR, PCWSTR, ICoreWebView2EnvironmentOptions*, ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler*);

void InitializeWebView() {
    // Try to extract embedded DLL from resources first
    HMODULE h = NULL;
    HRSRC res = FindResourceW(NULL, MAKEINTRESOURCEW(101), (LPCWSTR)RT_RCDATA);
    if (res) {
        HGLOBAL hg = LoadResource(NULL, res); DWORD sz = SizeofResource(NULL, res);
        if (hg && sz) {
            void* data = LockResource(hg);
            wchar_t tmp[MAX_PATH]; GetTempPathW(MAX_PATH, tmp);
            std::wstring dllPath = std::wstring(tmp) + L"WebView2Loader.dll";
            HANDLE hf = CreateFileW(dllPath.c_str(), GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
            if (hf != INVALID_HANDLE_VALUE) { DWORD w; WriteFile(hf, data, sz, &w, NULL); CloseHandle(hf); h = LoadLibraryW(dllPath.c_str()); }
        }
    }
    if (!h) h = LoadLibraryW(L"WebView2Loader.dll"); // fallback for dev builds
    if (!h) { MessageBoxW(NULL, L"WebView2Loader.dll not found!", L"Error", MB_ICONERROR); return; }
    auto fn = (CreateEnvFn)GetProcAddress(h, "CreateCoreWebView2EnvironmentWithOptions");
    if (!fn) { MessageBoxW(NULL, L"Failed to find entry point in DLL.", L"Error", MB_ICONERROR); return; }
    wchar_t ad[MAX_PATH]; GetEnvironmentVariableW(L"LOCALAPPDATA", ad, MAX_PATH);
    std::wstring udf = std::wstring(ad) + L"\\QuickLoopUp";
    SetEnvironmentVariableW(L"WEBVIEW2_ADDITIONAL_BROWSER_ARGUMENTS",
        L"--single-process --in-process-gpu --disable-gpu --disable-software-rasterizer --disable-crash-reporter "
        L"--disable-extensions --disable-background-networking --disable-sync --no-first-run "
        L"--disable-features=AudioServiceOutOfProcess,IsolateOrigins,site-per-process");
    SetEnvironmentVariableW(L"WEBVIEW2_CRASHPAD_DISABLE", L"1");
    fn(nullptr, udf.c_str(), nullptr, new EnvHandler());
}

// ── Hooks ─────────────────────────────────────────────────
LRESULT CALLBACK LowLevelKeyboardProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0 && popupVisible && (wParam == WM_KEYDOWN || wParam == WM_SYSKEYDOWN))
        PostThreadMessageW(mainThreadId, WM_APP_HIDE_POPUP, 0, 0);
    return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
}

LRESULT CALLBACK LowLevelMouseProc(int nCode, WPARAM wParam, LPARAM lParam) {
    if (nCode >= 0) {
        auto* ms = (MSLLHOOKSTRUCT*)lParam;
        if (ms->flags & LLMHF_INJECTED) return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
        if (popupVisible && wParam == WM_LBUTTONDOWN) {
            RECT r; GetWindowRect(hPopupWindow, &r);
            if (ms->pt.x<r.left||ms->pt.x>r.right||ms->pt.y<r.top||ms->pt.y>r.bottom)
                PostThreadMessageW(mainThreadId, WM_APP_HIDE_POPUP, 0, 0);
        }
        if (wParam == WM_LBUTTONDOWN) { downTime=GetTickCount64(); downPos=ms->pt; isDragging=false; SetTimer(hPopupWindow,2,800,NULL); }
        else if (wParam == WM_MOUSEMOVE && downTime) {
            double d = std::sqrt(std::pow(ms->pt.x-downPos.x,2)+std::pow(ms->pt.y-downPos.y,2));
            if (d > 10.0) { isDragging=true; downPos=ms->pt; SetTimer(hPopupWindow,2,800,NULL); }
        }
        else if (wParam == WM_LBUTTONUP) { downTime=0; isDragging=false; KillTimer(hPopupWindow,2); }
    }
    return CallNextHookEx(hMouseHook, nCode, wParam, lParam);
}

// ── Clipboard fallback ────────────────────────────────────
std::wstring BackupClipboardExtraction(int x, int y, bool wasDragging) {
    HWND hw = GetDesktopWindow();
    std::wstring oldClip;
    if (OpenClipboard(hw)) { HANDLE h=GetClipboardData(CF_UNICODETEXT); if(h){auto*p=(wchar_t*)GlobalLock(h);if(p)oldClip=p;GlobalUnlock(h);} CloseClipboard(); }
    if (OpenClipboard(hw)) { EmptyClipboard(); CloseClipboard(); }

    INPUT inp[4]={};
    inp[0].type=INPUT_KEYBOARD; inp[0].ki.wVk=VK_CONTROL;
    inp[1].type=INPUT_KEYBOARD; inp[1].ki.wVk='C';
    inp[2].type=INPUT_KEYBOARD; inp[2].ki.wVk='C'; inp[2].ki.dwFlags=KEYEVENTF_KEYUP;
    inp[3].type=INPUT_KEYBOARD; inp[3].ki.wVk=VK_CONTROL; inp[3].ki.dwFlags=KEYEVENTF_KEYUP;
    SendInput(4,inp,sizeof(INPUT)); Sleep(40);

    std::wstring txt;
    if (OpenClipboard(hw)) { HANDLE h=GetClipboardData(CF_UNICODETEXT); if(h){auto*p=(wchar_t*)GlobalLock(h);if(p)txt=p;GlobalUnlock(h);} CloseClipboard(); }

    if (txt.empty() && !wasDragging) {
        bool isText = false;
        IUIAutomation* a=NULL; POINT pt={x,y};
        if (SUCCEEDED(CoCreateInstance(__uuidof(CUIAutomation),NULL,CLSCTX_INPROC_SERVER,__uuidof(IUIAutomation),(void**)&a)) && a) {
            IUIAutomationElement* el=NULL;
            if (SUCCEEDED(a->ElementFromPoint(pt,&el)) && el) {
                CONTROLTYPEID ct=0; el->get_CurrentControlType(&ct);
                isText = (ct==UIA_EditControlTypeId||ct==UIA_TextControlTypeId||ct==UIA_DocumentControlTypeId||ct==UIA_HyperlinkControlTypeId);
                el->Release();
            } a->Release();
        }
        if (isText) {
            INPUT ms[4]={}; ms[0].type=INPUT_MOUSE;ms[0].mi.dwFlags=MOUSEEVENTF_LEFTDOWN;
            ms[1].type=INPUT_MOUSE;ms[1].mi.dwFlags=MOUSEEVENTF_LEFTUP;
            ms[2].type=INPUT_MOUSE;ms[2].mi.dwFlags=MOUSEEVENTF_LEFTDOWN;
            ms[3].type=INPUT_MOUSE;ms[3].mi.dwFlags=MOUSEEVENTF_LEFTUP;
            SendInput(4,ms,sizeof(INPUT)); Sleep(30);
            SendInput(4,inp,sizeof(INPUT)); Sleep(40);
            if (OpenClipboard(hw)) { HANDLE h=GetClipboardData(CF_UNICODETEXT); if(h){auto*p=(wchar_t*)GlobalLock(h);if(p)txt=p;GlobalUnlock(h);} CloseClipboard(); }
        }
    }

    if (!oldClip.empty()) {
        if (OpenClipboard(hw)) { EmptyClipboard(); size_t sz=(oldClip.length()+1)*sizeof(wchar_t);
            HGLOBAL hm=GlobalAlloc(GMEM_MOVEABLE,sz); if(hm){memcpy(GlobalLock(hm),oldClip.c_str(),sz);GlobalUnlock(hm);SetClipboardData(CF_UNICODETEXT,hm);} CloseClipboard(); }
    } else { if (OpenClipboard(hw)) { EmptyClipboard(); CloseClipboard(); } }
    return txt;
}

// ── Text extraction ───────────────────────────────────────
void ExtractTextAndSearch(int x, int y, bool wasDragging) {
    if (!isWebViewReady) return;
    bool ok = false; std::wstring word;
    HRESULT hr = CoInitialize(NULL);
    IUIAutomation* au=NULL;
    hr = CoCreateInstance(__uuidof(CUIAutomation),NULL,CLSCTX_INPROC_SERVER,__uuidof(IUIAutomation),(void**)&au);
    if (SUCCEEDED(hr) && au) {
        IUIAutomationElement* el=NULL; POINT pt={x,y};
        if (SUCCEEDED(au->ElementFromPoint(pt,&el)) && el) {
            IUIAutomationTextPattern* tp=NULL;
            if (SUCCEEDED(el->GetCurrentPatternAs(UIA_TextPatternId,__uuidof(IUIAutomationTextPattern),(void**)&tp)) && tp) {
                IUIAutomationTextRangeArray* sel=NULL;
                if (SUCCEEDED(tp->GetSelection(&sel)) && sel) {
                    int len=0; sel->get_Length(&len);
                    if (len>0) { IUIAutomationTextRange* r=NULL; sel->GetElement(0,&r);
                        if(r){BSTR t;if(SUCCEEDED(r->GetText(-1,&t))&&t){word=t;SysFreeString(t);if(!word.empty())ok=true;}r->Release();} }
                    sel->Release();
                }
                if (!ok) { IUIAutomationTextRange* r=NULL;
                    if(SUCCEEDED(tp->RangeFromPoint(pt,&r))&&r){r->ExpandToEnclosingUnit(TextUnit_Word);
                    BSTR t;if(SUCCEEDED(r->GetText(-1,&t))&&t){word=t;SysFreeString(t);if(!word.empty())ok=true;}r->Release();} }
                tp->Release();
            } el->Release();
        } au->Release();
    } CoUninitialize();

    if (!ok || word.empty()) word = BackupClipboardExtraction(x, y, wasDragging);
    while (!word.empty() && iswspace(word.back())) word.pop_back();
    while (!word.empty() && iswspace(word.front())) word.erase(word.begin());
    if (!word.empty()) {
        PostThreadMessageW(mainThreadId, WM_APP_SHOW_CONTENT, 0, (LPARAM)new PopupData{x,y,GetLoadingHtml(word)});
        FetchDefinitionAsync(word, x, y);
    } else { PostThreadMessageW(mainThreadId, WM_APP_HIDE_POPUP, 0, 0); }
}

// ── Google redirect ───────────────────────────────────────
std::wstring GoogleHtml(const std::wstring& q) {
    std::wstring sw; for(auto c:q){if(c==L'\\')sw+=L"\\\\";else if(c==L'`')sw+=L"\\`";else sw+=c;}
    return L"<html><head><style>body{background:#111;display:flex;flex-direction:column;justify-content:center;"
           L"align-items:center;height:100vh;margin:0;border-radius:12px;font-family:'Segoe UI',sans-serif;"
           L"animation:f .2s ease both}@keyframes f{from{opacity:0}to{opacity:1}}"
           L".t{color:#777;font-size:14px;margin-top:12px}"
           L".s{width:28px;height:28px;border-radius:50%;border:3px solid rgba(255,255,255,.08);"
           L"border-top:3px solid #fff;animation:r .8s linear infinite}@keyframes r{to{transform:rotate(360deg)}}"
           L"</style></head><body><div class='s'></div><div class='t'>Searching the web...</div>"
           L"<script>setTimeout(()=>window.location.replace('https://www.google.com/search?q='+encodeURIComponent(`"
           + sw + L"`)),300)</script></body></html>";
}

// ── Dictionary fetch ──────────────────────────────────────
void FetchDefinitionAsync(std::wstring word, int x, int y) {
    for (auto c : word) if (iswspace(c)) {
        PostThreadMessageW(mainThreadId, WM_APP_SHOW_CONTENT, 0, (LPARAM)new PopupData{x,y,GoogleHtml(word)}); return;
    }
    bool found = false; std::wstring finalHtml;
    if (hPersistentSession) {
        HINTERNET hc = WinHttpConnect(hPersistentSession, L"api.dictionaryapi.dev", INTERNET_DEFAULT_HTTPS_PORT, 0);
        if (hc) {
            std::wstring path = L"/api/v2/entries/en/" + word;
            HINTERNET hr = WinHttpOpenRequest(hc, L"GET", path.c_str(), NULL, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, WINHTTP_FLAG_SECURE);
            if (hr) {
                if (WinHttpSendRequest(hr, NULL, 0, NULL, 0, 0, 0) && WinHttpReceiveResponse(hr, NULL)) {
                    DWORD sc=0, ds=sizeof(sc);
                    WinHttpQueryHeaders(hr, WINHTTP_QUERY_STATUS_CODE|WINHTTP_QUERY_FLAG_NUMBER, NULL, &sc, &ds, NULL);
                    if (sc == 200) {
                        DWORD sz=0, dl=0; std::string resp;
                        do { if(WinHttpQueryDataAvailable(hr,&sz)&&sz>0){std::vector<char>b(sz+1,0);if(WinHttpReadData(hr,&b[0],sz,&dl))resp.append(&b[0],dl);} } while(sz>0);
                        if (!resp.empty()) {
                            std::string wn(word.begin(), word.end());
                            auto ph = FindJsonVal(resp,"phonetic"), pos = FindJsonVal(resp,"partOfSpeech");
                            auto defs = FindAllJsonVals(resp,"definition",4), exs = FindAllJsonVals(resp,"example",4);
                            auto syns = FindJsonArr(resp,"synonyms");
                            std::string h = "<div class='c'><div class='wh'><span class='wt'>"+HtmlEsc(wn)+"</span>";
                            if(!ph.empty()) h += "<span class='ph'>"+HtmlEsc(ph)+"</span>";
                            h += "</div>"; if(!pos.empty()) h += "<div class='pb'>"+HtmlEsc(pos)+"</div>";
                            for(size_t i=0;i<defs.size();i++) {
                                h += "<div class='dc'><span class='dn'>"+std::to_string(i+1)+"</span><span class='dt'>"+HtmlEsc(defs[i])+"</span></div>";
                                if(i<exs.size()&&!exs[i].empty()) h += "<div class='de'>\""+HtmlEsc(exs[i])+"\"</div>";
                            }
                            if(!syns.empty()) { h += "<div class='sd'></div><div class='ss'><div class='sl'>Synonyms</div><div class='sc'>";
                                for(auto&s:syns) h += "<span class='sp'>"+HtmlEsc(s)+"</span>"; h += "</div></div>"; }
                            h += "<div class='fb'>"
                                 "<a class='ft' onclick=\"window.location='https://www.google.com/search?q=define+"+wn+"'\">Full Definition</a>"
                                 "<a class='ft' onclick=\"window.location='https://www.google.com/search?q="+wn+"+synonyms'\">More Synonyms</a></div></div>";
                            finalHtml = MakeSimplePage(h); found = true;
                        }
                    } else if (sc == 404) { finalHtml = GoogleHtml(word); found = true; }
                }
                WinHttpCloseHandle(hr);
            } WinHttpCloseHandle(hc);
        }
    }
    if (!found) finalHtml = GoogleHtml(word);
    PostThreadMessageW(mainThreadId, WM_APP_SHOW_CONTENT, 0, (LPARAM)new PopupData{x,y,finalHtml});
}

// ── Window ────────────────────────────────────────────────
void CreatePopupWindow(HINSTANCE hInstance) {
    WNDCLASSW wc={};
    wc.lpfnWndProc=PopupWindowProc; wc.hInstance=hInstance; wc.lpszClassName=L"QuickLoopUpPopup";
    wc.hbrBackground=CreateSolidBrush(RGB(17,17,17)); wc.hCursor=LoadCursor(NULL,IDC_ARROW);
    RegisterClassW(&wc);
    hPopupWindow = CreateWindowExW(WS_EX_TOPMOST|WS_EX_TOOLWINDOW|WS_EX_NOACTIVATE,
        L"QuickLoopUpPopup",L"",WS_POPUP|WS_CLIPCHILDREN,0,0,440,380,NULL,NULL,hInstance,NULL);
    int cp=2; DwmSetWindowAttribute(hPopupWindow,33,&cp,sizeof(cp));
    MARGINS m={1,1,1,1}; DwmExtendFrameIntoClientArea(hPopupWindow,&m);
}

void ShowPopup(int x, int y, const std::wstring& html) {
    HMONITOR hm = MonitorFromPoint({x,y},MONITOR_DEFAULTTONEAREST);
    UINT dx=96,dy=96; GetDpiForMonitor(hm,MDT_EFFECTIVE_DPI,&dx,&dy);
    double s = (double)dx/96.0; int w=(int)(440*s), h=(int)(380*s);
    if (webView) webView->NavigateToString(html.c_str());
    MONITORINFO mi={sizeof(mi)}; GetMonitorInfo(hm,&mi);
    if (x==-1&&y==-1) { x=mi.rcWork.left+((mi.rcWork.right-mi.rcWork.left)-w)/2; y=mi.rcWork.top+((mi.rcWork.bottom-mi.rcWork.top)-h)/2; }
    else { int off=(int)(25*s); x-=w/2; y+=off;
        if(x<mi.rcWork.left+5)x=mi.rcWork.left+5; if(x+w>mi.rcWork.right-5)x=mi.rcWork.right-w-5;
        if(y+h>mi.rcWork.bottom-5)y=y-h-off*2; }
    SetWindowPos(hPopupWindow,HWND_TOPMOST,x,y,w,h,SWP_SHOWWINDOW|SWP_NOACTIVATE);
    if(webViewController){webViewController->put_IsVisible(TRUE);RECT b;GetClientRect(hPopupWindow,&b);webViewController->put_Bounds(b);}
    popupVisible = true;
}

LRESULT CALLBACK PopupWindowProc(HWND hw, UINT msg, WPARAM wp, LPARAM lp) {
    switch (msg) {
    case WM_SIZE: if(webViewController){RECT b;GetClientRect(hw,&b);webViewController->put_Bounds(b);} return 0;
    case WM_TIMER:
        if(wp==1){ShowWindow(hw,SW_HIDE);popupVisible=false;KillTimer(hw,1);}
        else if(wp==2){KillTimer(hw,2);if(!popupVisible)PostThreadMessageW(mainThreadId,WM_APP_TRIGGER_EXTRACT,downPos.x,downPos.y);}
        return 0;
    case WM_NCHITTEST: return HTCLIENT;
    } return DefWindowProc(hw,msg,wp,lp);
}
