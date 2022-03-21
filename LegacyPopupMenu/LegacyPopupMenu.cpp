// LegacyPopupMenu.cpp : アプリケーションのエントリ ポイントを定義します。
//

#include "framework.h"
#include "LegacyPopupMenu.h"
#include <commCtrl.h>
#include <string>
#include <windows.h>
#include <windowsx.h>

#pragma comment(lib, "comctl32.lib")

#define MENU_WIDTH_PLUS 5
#define MAX_LOADSTRING 100
#define MN_SETHMENU 0x01E0
//  0x01E1  MN_GETHMENU これは公開されている
#define MN_SIZEWINDOW	0x01E2
#define MN_OPENHIERARCHY	0x01E3
#define MN_CLOSEHIERARCHY	0x01E4
#define MN_SELECTITEM	0x01E5
#define MN_CANCELMENUS	0x01E6
#define MN_SELECTFIRSTVALIDITEM	0x01E7
#define MN_GETPPOPUPMENU	0x01EA
#define MN_FINDMENUWINDOWFROMPOINT	0x01EB
#define MN_SHOWPOPUPWINDOW	0x01EC
#define MN_BUTTONDOWN	0x01ED
#define MN_MOUSEMOVE	0x01EE
#define MN_BUTTONUP	0x01EF
#define MN_SETTIMERTOOPENHIERARCHY	0x01F0
#define MN_DBLCLK	0x01F1
#define WM_UAHDESTROYWINDOW 0x0090
#define WM_UAHDRAWMENU 0x0091
#define WM_UAHDRAWMENUITEM 0x0092
#define WM_UAHINITMENU 0x0093
#define WM_UAHMEASUREMENUITEM 0x0094
#define WM_UAHNCPAINTMENUPOPUP 0x0095

HHOOK hook;
bool isPopUp = false;
int cnt = 0;
HMENU hMenu;                              // メニューのハンドル
HMENU hSubMenu1;                          // メニューのハンドル
HMENU hSubMenu2;                          // メニューのハンドル
HMENU hSubMenu3;                          // メニューのハンドル
LPMEASUREITEMSTRUCT lpMI;                 // LPMEASUREITEMSTRUCT構造体
LPDRAWITEMSTRUCT menuItem;                    // DRAWITEMSTRUCT構造体

#define MAX_LOADSTRING 100

// グローバル変数:
HINSTANCE hInst;                                // 現在のインターフェイス
WCHAR szTitle[MAX_LOADSTRING];                  // タイトル バーのテキスト
WCHAR szWindowClass[MAX_LOADSTRING];            // メイン ウィンドウ クラス名

// このコード モジュールに含まれる関数の宣言を転送します:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
LRESULT CALLBACK	CBTProc(int nCode, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK	DialogProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData);
std::wstring MsgToString(int msg);

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
  _In_opt_ HINSTANCE hPrevInstance,
  _In_ LPWSTR    lpCmdLine,
  _In_ int       nCmdShow) {
  UNREFERENCED_PARAMETER(hPrevInstance);
  UNREFERENCED_PARAMETER(lpCmdLine);

  // TODO: ここにコードを挿入してください。

  // グローバル文字列を初期化する
  LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
  LoadStringW(hInstance, IDC_LEGACYPOPUPMENU, szWindowClass, MAX_LOADSTRING);
  MyRegisterClass(hInstance);

  // アプリケーション初期化の実行:
  if (!InitInstance(hInstance, nCmdShow)) {
    return FALSE;
  }

  HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_LEGACYPOPUPMENU));

  MSG msg;

  // メイン メッセージ ループ:
  while (GetMessage(&msg, nullptr, 0, 0)) {
    if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg)) {
      TranslateMessage(&msg);
      DispatchMessage(&msg);
    }
  }

  return (int)msg.wParam;
}

//
//  関数: MyRegisterClass()
//
//  目的: ウィンドウ クラスを登録します。
//
ATOM MyRegisterClass(HINSTANCE hInstance) {
  WNDCLASSEXW wcex;

  wcex.cbSize = sizeof(WNDCLASSEX);

  wcex.style = CS_HREDRAW | CS_VREDRAW;
  wcex.lpfnWndProc = WndProc;
  wcex.cbClsExtra = 0;
  wcex.cbWndExtra = 0;
  wcex.hInstance = hInstance;
  wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_LEGACYPOPUPMENU));
  wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
  wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
  wcex.lpszMenuName = nullptr; //MAKEINTRESOURCEW(IDC_LEGACYPOPUPMENU);
  wcex.lpszClassName = szWindowClass;
  wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

  return RegisterClassExW(&wcex);
}

//
//   関数: InitInstance(HINSTANCE, int)
//
//   目的: インスタンス ハンドルを保存して、メイン ウィンドウを作成します
//
//   コメント:
//
//        この関数で、グローバル変数でインスタンス ハンドルを保存し、
//        メイン プログラム ウィンドウを作成および表示します。
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow) {
  hInst = hInstance; // グローバル変数にインスタンス ハンドルを格納する

  HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
    CW_USEDEFAULT, 0, 640, 480, nullptr, nullptr, hInstance, nullptr);

  if (!hWnd) {
    return FALSE;
  }

  ShowWindow(hWnd, nCmdShow);
  UpdateWindow(hWnd);

  hook = SetWindowsHookEx(WH_CBT, (HOOKPROC)CBTProc, NULL, GetCurrentThreadId());

  return TRUE;
}

LRESULT CALLBACK CBTProc(int nCode, WPARAM wParam, LPARAM lParam) {
  TCHAR buff[MAX_PATH];
  CHAR pszClassA[MAX_CLASS_NAME];
  if (nCode == HCBT_CREATEWND) {
    HWND hWnd = (HWND)wParam;
    ::GetWindowText(hWnd, buff, MAX_PATH);
    ::GetClassNameA(hWnd, pszClassA, MAX_CLASS_NAME);
    if (!lstrcmpA(pszClassA, "#32768") && isPopUp) {
      SetWindowSubclass((HWND)wParam, DialogProc, (UINT_PTR)DialogProc, 0);
      LPCBT_CREATEWND window = ((LPCBT_CREATEWND)lParam);
      LPCREATESTRUCT lpcsWnd = window->lpcs;
      DWORD dwStyle = lpcsWnd->style;
      DWORD dwExStyle = lpcsWnd->dwExStyle;
    }
  } else if (nCode == HCBT_DESTROYWND) {
    RemoveWindowSubclass((HWND)wParam, DialogProc, (UINT_PTR)DialogProc);
  } else if (nCode == HCBT_ACTIVATE) {
  }
  return CallNextHookEx(0, nCode, wParam, lParam);
}

LRESULT CALLBACK DialogProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR uIdSubclass, DWORD_PTR dwRefData) {
  OutputDebugString((std::to_wstring(++cnt) + L":子:" + MsgToString(message) + L"\n").c_str());
  switch (message) {
    case MN_SELECTITEM:
      {
        OutputDebugString((std::to_wstring(++cnt) + L":選択インデックス:" + std::to_wstring((int)wParam) + L"\n").c_str());
        //if ((int)wParam==-3|| (int)wParam == -4) {
        //return 0;
        //}
        return DefSubclassProc(hWnd, message, wParam, lParam);
      }

      // case WM_PAINT:
        // case WM_PRINTCLIENT:
    case WM_PRINT:
      {
        RECT windowRect{};
        GetWindowRect(hWnd, &windowRect);
        RECT clientRect{};
        GetClientRect(hWnd, &clientRect);

        HDC hdc;
        DWORD flags;
        PAINTSTRUCT ps;
        if (message == WM_PAINT) {
          hdc = BeginPaint(hWnd, &ps);
        } else {
          hdc = (HDC)wParam;
          flags = (DWORD)lParam;
        }

        int windowWidth = windowRect.right - windowRect.left;
        int windowHeight = windowRect.bottom - windowRect.top;
        int clientWidth = clientRect.right - clientRect.left;
        int clientHeight = clientRect.bottom - clientRect.top;

        RECT drawRect{};
        // 全体
        drawRect = { 0, 0, windowWidth, windowHeight };
        FillRect(hdc, &drawRect, CreateSolidBrush(RGB(33, 33, 33)));
        // 枠線
        drawRect = { 0, 0, windowWidth, windowHeight };
        FrameRect(hdc, &drawRect, CreateSolidBrush(RGB(255, 255, 0)));
        // 上
        ExcludeClipRect(hdc, 0, 0, windowWidth, 3);
        // 下
        ExcludeClipRect(hdc, 0, windowHeight - 3, windowWidth, windowHeight);
        // 左
        ExcludeClipRect(hdc, 0, 0, 3, windowHeight);
        // 右
        ExcludeClipRect(hdc, windowWidth - 3, 0, windowWidth, windowHeight);
        // ▲の部分
        // ▼の部分
        if (6 < windowHeight - clientHeight) {
          ExcludeClipRect(hdc, 3, 3, 20, 23);
          ExcludeClipRect(hdc, 40, 3, windowWidth, 23);
          ExcludeClipRect(hdc, 3, windowHeight - 23, 20, windowHeight - 3);
          ExcludeClipRect(hdc, 40, windowHeight - 23, windowWidth, windowHeight - 3);
        }

        return DefSubclassProc(hWnd, message, wParam, lParam);
      }
    case WM_NCPAINT:
      {
        auto result = DefSubclassProc(hWnd, message, wParam, lParam);

        RECT windowRect{};
        GetWindowRect(hWnd, &windowRect);
        RECT clientRect{};
        GetClientRect(hWnd, &clientRect);

        int windowWidth = windowRect.right - windowRect.left;
        int windowHeight = windowRect.bottom - windowRect.top;
        int clientWidth = clientRect.right - clientRect.left;
        int clientHeight = clientRect.bottom - clientRect.top;

        HDC hdc = ::GetWindowDC(hWnd);

        // ▲を除いた部分
        // ▼を除いた部分
        if (6 < windowHeight - clientHeight) {
          ExcludeClipRect(hdc, 20, 3, 40, 23);
          ExcludeClipRect(hdc, 20, windowHeight - 23, 40, windowHeight - 3);
        }

        RECT drawRect{};
        // 全体
        drawRect = { 0, 0, windowWidth, windowHeight };
        FillRect(hdc, &drawRect, CreateSolidBrush(RGB(33, 33, 33)));
        // 枠線
        drawRect = { 0, 0, windowWidth, windowHeight };
        FrameRect(hdc, &drawRect, CreateSolidBrush(RGB(255, 255, 0)));

        ::ReleaseDC(hWnd, hdc);

        return 0;
      }
  }
  return DefSubclassProc(hWnd, message, wParam, lParam);
}

void appendMenuItem(HMENU hmenu, int nId, const wchar_t* lpszItemName, HMENU hmenuSub) {
  MENUITEMINFO mii = { sizeof(MENUITEMINFO), MIIM_ID | MIIM_TYPE | MIIM_SUBMENU,
             lpszItemName != NULL ? MFT_STRING : MFT_SEPARATOR, 0, nId, hmenuSub };
  mii.dwTypeData = (LPWSTR)lpszItemName;
  InsertMenuItem(hmenu, nId, FALSE, &mii);
}

void InitializeMenuItem(HMENU hMenu, HMENU hSubMenu, int itemId, int type) {
  WCHAR text[UCHAR_MAX]{};
  memset(&text[0], 0, sizeof(text));
  LoadString(hInst, itemId, (LPTSTR)text, UCHAR_MAX);

  MENUITEMINFO mii{};

  mii.cbSize = sizeof(MENUITEMINFO);
  mii.fMask = MIIM_TYPE | MIIM_CHECKMARKS | MIIM_DATA | MIIM_ID | MIIM_STATE | MIIM_SUBMENU;
  mii.fType = type;
  mii.fState = MFS_ENABLED;
  mii.wID = itemId;
  mii.hSubMenu = hSubMenu;
  mii.dwItemData = (ULONG_PTR)text;
  mii.dwTypeData = (LPWSTR)text;

  InsertMenuItem(hMenu, 0xFFFFFFFF, TRUE, &mii);
}

//
//  関数: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  目的: メイン ウィンドウのメッセージを処理します。
//
//  WM_COMMAND  - アプリケーション メニューの処理
//  WM_PAINT    - メイン ウィンドウを描画する
//  WM_DESTROY  - 中止メッセージを表示して戻る
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam) {
  OutputDebugString((std::to_wstring(++cnt) + L":親:" + MsgToString(message) + L"\n").c_str());
  POINT po;
  static HMENU tmp, hmenuR, hmenuL;
  int wmId, wmEvent;
  SIZE sz;
  HFONT hFont, hFontOld;
  static LOGFONT g_menufont;
  static int g_icon_vx, g_icon_vy;
  static DWORD g_menu_backcolor;
  static DWORD g_menu_textcolor;
  static HICON hIcon;
  HBRUSH hBrush;
  MENUITEMINFO mi;

  switch (message) {
    case WM_RBUTTONUP:
      isPopUp = true;
      po.x = LOWORD(lParam);
      po.y = HIWORD(lParam);
      ClientToScreen(hWnd, &po);
      TrackPopupMenu(hSubMenu1, TPM_LEFTALIGN | TPM_BOTTOMALIGN, po.x, po.y, 0, hWnd, NULL);
      return 0;

    case WM_CREATE:
      {
        hMenu = CreateMenu();

        hSubMenu1 = CreatePopupMenu();
        hSubMenu2 = CreatePopupMenu();
        hSubMenu3 = CreatePopupMenu();

        InitializeMenuItem(hMenu, hSubMenu1, IDS_MENU_FILE, MFT_STRING);
        InitializeMenuItem(hMenu, hSubMenu2, IDS_MENU_HELP, MFT_STRING);

        InitializeMenuItem(hSubMenu1, hSubMenu3, IDS_MENU_FILE_1, MFT_STRING | MFT_OWNERDRAW);
        InitializeMenuItem(hSubMenu1, NULL, NULL, MFT_SEPARATOR | MFT_OWNERDRAW);
        InitializeMenuItem(hSubMenu1, NULL, IDS_MENU_FILE_3, MFT_STRING | MFT_OWNERDRAW);
        InitializeMenuItem(hSubMenu1, NULL, IDS_MENU_FILE_4, MFT_STRING | MFT_OWNERDRAW);
        InitializeMenuItem(hSubMenu1, NULL, IDS_MENU_FILE_5, MFT_STRING | MFT_OWNERDRAW);
        
        InitializeMenuItem(hSubMenu2, NULL, IDS_MENU_HELP_VERSION, MFT_STRING | MFT_OWNERDRAW);
        SetMenuDefaultItem(hSubMenu2, IDS_MENU_HELP_VERSION, 0);
        InitializeMenuItem(hSubMenu3, NULL, IDS_MENU_FILE_2, MFT_STRING | MFT_OWNERDRAW);

        ::SetMenu(hWnd, hMenu);

        NONCLIENTMETRICS metrics{};
        metrics.cbSize = sizeof(NONCLIENTMETRICS);
        SystemParametersInfo(SPI_GETNONCLIENTMETRICS, sizeof(metrics),
          &metrics, 0);
        g_menufont = metrics.lfMenuFont;
        g_icon_vx = (int)GetSystemMetrics(SM_CXSMICON);
        g_icon_vy = (int)GetSystemMetrics(SM_CYSMICON);
        //hIcon = LoadIcon(hInst, MAKEINTRESOURCE(IDI_TEST));
        g_menu_backcolor = RGB(55, 55, 55);
        g_menu_textcolor = RGB(255, 255, 255);

        const int cxBorder = ::GetSystemMetrics(SM_CXBORDER);
        const int cyBorder = ::GetSystemMetrics(SM_CYBORDER);
        const int cxEdge = ::GetSystemMetrics(SM_CXEDGE);
        const int cyEdge = ::GetSystemMetrics(SM_CYEDGE);
        const int cxFrame = ::GetSystemMetrics(SM_CXFRAME);
        const int cyFrame = ::GetSystemMetrics(SM_CYFRAME);
        const int cxSmIcon = ::GetSystemMetrics(SM_CXSMICON);
        const int cySmIcon = ::GetSystemMetrics(SM_CYSMICON);
      }
      break;

    case WM_INITMENUPOPUP:
      {
        MENUITEMINFO menuInfo{};

        HMENU hmenuPopup = (HMENU)wParam;
        int index = LOWORD(lParam);

        isPopUp = false;
        if (hmenuPopup != hMenu) {
          isPopUp = true;
        }
      }
      return 0;

    case WM_MEASUREITEM:
      {
        static int height = 0;
        lpMI = (LPMEASUREITEMSTRUCT)lParam;
        WCHAR text[MAX_LOADSTRING];
        LoadString(hInst, lpMI->itemID, text, MAX_LOADSTRING);

        HDC hdc = ::GetDC(hWnd);
        hFont = CreateFontIndirect(&g_menufont);
        hFontOld = (HFONT)SelectObject(hdc, hFont);
        ::GetTextExtentPoint32(hdc, text, lstrlen(text), &sz);
        lpMI->itemWidth = sz.cx + g_icon_vx + MENU_WIDTH_PLUS + 50;
        lpMI->itemHeight = sz.cy + 5;
        height += 5;
        ::ReleaseDC(hWnd, hdc);
        break;
      }

    case WM_DRAWITEM:
      {
        menuItem = (LPDRAWITEMSTRUCT)lParam;

        if (menuItem->CtlType == ODT_MENU) {
          MENUITEMINFO menuInfo{};
          menuInfo.cbSize = sizeof(MENUITEMINFO);
          menuInfo.fMask = MIIM_TYPE | MIIM_CHECKMARKS | MIIM_DATA | MIIM_ID | MIIM_STATE | MIIM_SUBMENU;
          GetMenuItemInfo((HMENU)menuItem->hwndItem, menuItem->itemID, FALSE, &menuInfo);

          RECT itemRect = menuItem->rcItem;
          HDC hdc = menuItem->hDC;

          RECT windowRect;
          GetWindowRect(hWnd, &windowRect);
          RECT clientRect;
          GetClientRect(hWnd, &clientRect);

          hBrush = CreateSolidBrush(RGB(33, 33, 33));
          if (menuItem->itemID == 0) {
            FillRect(hdc, &itemRect, CreateSolidBrush(RGB(33, 33, 33)));
            itemRect.bottom = itemRect.top + 1;
            FillRect(hdc, &itemRect, CreateSolidBrush(RGB(255, 0, 0)));
            return 0;
          }

          hFont = CreateFontIndirect(&g_menufont);
          hFontOld = (HFONT)SelectObject(hdc, hFont);

          FillRect(hdc, &itemRect, CreateSolidBrush(RGB(33, 33, 33)));
          if (menuItem->itemState & ODS_SELECTED) {
            hBrush = CreateSolidBrush(g_menu_backcolor);
            SetTextColor(hdc, g_menu_textcolor);
            SetBkColor(hdc, g_menu_backcolor);
          } else {
            SetTextColor(hdc, g_menu_textcolor);
            SetBkColor(hdc, RGB(33, 33, 33));
            hBrush = CreateSolidBrush(RGB(33, 33, 33));
          }
          ::FillRect(hdc, &itemRect, hBrush);
          ::DeleteObject(hBrush);
          //DrawIconEx(hdc, rc.left, rc.top + 1, hIcon, g_icon_vx, g_icon_vy, 0,	NULL, DI_IMAGE | DI_MASK);

          WCHAR text[MAX_LOADSTRING];
          ::LoadString(hInst, menuItem->itemID, text, MAX_LOADSTRING);
          // ::SetTextAlign(hdc, TA_RIGHT);
          ::ExtTextOut(hdc, itemRect.left + g_icon_vx + MENU_WIDTH_PLUS, itemRect.top + 2, ETO_OPAQUE, &itemRect, text, lstrlen(text), NULL);
          ::SelectObject(hdc, hFontOld);
          ::DeleteObject(hFont);
        }
        break;
      }

    case WM_COMMAND:
      {
        int wmId = LOWORD(wParam);
        // 選択されたメニューの解析:
        switch (wmId) {
          case IDS_MENU_HELP_VERSION:
          case IDM_ABOUT:
            DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
            break;
          case IDM_EXIT:
            DestroyWindow(hWnd);
            break;
          default:
            return DefWindowProc(hWnd, message, wParam, lParam);
        }
      }
      break;
    case WM_PAINT:
      {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hWnd, &ps);
        // TODO: HDC を使用する描画コードをここに追加してください...
        EndPaint(hWnd, &ps);
      }
      break;
    case WM_DESTROY:
      PostQuitMessage(0);
      break;
    default:
      return DefWindowProc(hWnd, message, wParam, lParam);
  }
  return 0;
}

// バージョン情報ボックスのメッセージ ハンドラーです。
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam) {
  UNREFERENCED_PARAMETER(lParam);
  switch (message) {
    case WM_INITDIALOG:
      return (INT_PTR)TRUE;

    case WM_COMMAND:
      if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL) {
        EndDialog(hDlg, LOWORD(wParam));
        return (INT_PTR)TRUE;
      }
      break;
  }
  return (INT_PTR)FALSE;
}

/// <summary>
/// Windowメッセージコードを文字列に変換します。
/// </summary>
/// <param name="msg"></param>
/// <returns></returns>
#define MSG_TO_STRING(val) case val: text = L#val; break
std::wstring MsgToString(int msg) {

  std::wstring text;
  switch (msg) {
    MSG_TO_STRING(MN_SIZEWINDOW);
    MSG_TO_STRING(MN_OPENHIERARCHY);
    MSG_TO_STRING(MN_CLOSEHIERARCHY);
    MSG_TO_STRING(MN_SELECTITEM);
    MSG_TO_STRING(MN_CANCELMENUS);
    MSG_TO_STRING(MN_SELECTFIRSTVALIDITEM);
    MSG_TO_STRING(MN_GETPPOPUPMENU);
    MSG_TO_STRING(MN_FINDMENUWINDOWFROMPOINT);
    MSG_TO_STRING(MN_SHOWPOPUPWINDOW);
    MSG_TO_STRING(MN_BUTTONDOWN);
    MSG_TO_STRING(MN_MOUSEMOVE);
    MSG_TO_STRING(MN_BUTTONUP);
    MSG_TO_STRING(MN_SETTIMERTOOPENHIERARCHY);
    MSG_TO_STRING(MN_DBLCLK);

    MSG_TO_STRING(WM_NULL);
    MSG_TO_STRING(WM_CREATE);
    MSG_TO_STRING(WM_DESTROY);
    MSG_TO_STRING(WM_MOVE);
    MSG_TO_STRING(WM_SIZE);
    MSG_TO_STRING(WM_ACTIVATE);
    //MSG_TO_STRING(WA_INACTIVE);
    //MSG_TO_STRING(WA_ACTIVE);
    //MSG_TO_STRING(WA_CLICKACTIVE);
    MSG_TO_STRING(WM_SETFOCUS);
    MSG_TO_STRING(WM_KILLFOCUS);
    MSG_TO_STRING(WM_ENABLE);
    MSG_TO_STRING(WM_SETREDRAW);
    MSG_TO_STRING(WM_SETTEXT);
    MSG_TO_STRING(WM_GETTEXT);
    MSG_TO_STRING(WM_GETTEXTLENGTH);
    MSG_TO_STRING(WM_PAINT);
    MSG_TO_STRING(WM_CLOSE);
    MSG_TO_STRING(WM_QUERYENDSESSION);
    MSG_TO_STRING(WM_QUIT);
    MSG_TO_STRING(WM_QUERYOPEN);
    MSG_TO_STRING(WM_ERASEBKGND);
    MSG_TO_STRING(WM_SYSCOLORCHANGE);
    MSG_TO_STRING(WM_ENDSESSION);
    MSG_TO_STRING(WM_SHOWWINDOW);
    MSG_TO_STRING(WM_WININICHANGE);
    //MSG_TO_STRING(WM_SETTINGCHANGE);
    MSG_TO_STRING(WM_DEVMODECHANGE);
    MSG_TO_STRING(WM_ACTIVATEAPP);
    MSG_TO_STRING(WM_FONTCHANGE);
    MSG_TO_STRING(WM_TIMECHANGE);
    MSG_TO_STRING(WM_CANCELMODE);
    MSG_TO_STRING(WM_SETCURSOR);
    MSG_TO_STRING(WM_MOUSEACTIVATE);
    MSG_TO_STRING(WM_CHILDACTIVATE);
    MSG_TO_STRING(WM_QUEUESYNC);
    MSG_TO_STRING(WM_GETMINMAXINFO);
    MSG_TO_STRING(WM_PAINTICON);
    MSG_TO_STRING(WM_ICONERASEBKGND);
    MSG_TO_STRING(WM_NEXTDLGCTL);
    MSG_TO_STRING(WM_SPOOLERSTATUS);
    MSG_TO_STRING(WM_DRAWITEM);
    MSG_TO_STRING(WM_MEASUREITEM);
    MSG_TO_STRING(WM_DELETEITEM);
    MSG_TO_STRING(WM_VKEYTOITEM);
    MSG_TO_STRING(WM_CHARTOITEM);
    MSG_TO_STRING(WM_SETFONT);
    MSG_TO_STRING(WM_GETFONT);
    MSG_TO_STRING(WM_SETHOTKEY);
    MSG_TO_STRING(WM_GETHOTKEY);
    MSG_TO_STRING(WM_QUERYDRAGICON);
    MSG_TO_STRING(WM_COMPAREITEM);
    MSG_TO_STRING(WM_GETOBJECT);
    MSG_TO_STRING(WM_COMPACTING);
    MSG_TO_STRING(WM_COMMNOTIFY);
    MSG_TO_STRING(WM_WINDOWPOSCHANGING);
    MSG_TO_STRING(WM_WINDOWPOSCHANGED);
    MSG_TO_STRING(WM_POWER);
    MSG_TO_STRING(WM_COPYDATA);
    MSG_TO_STRING(WM_CANCELJOURNAL);
    MSG_TO_STRING(WM_NOTIFY);
    MSG_TO_STRING(WM_INPUTLANGCHANGEREQUEST);
    MSG_TO_STRING(WM_INPUTLANGCHANGE);
    MSG_TO_STRING(WM_TCARD);
    MSG_TO_STRING(WM_HELP);
    MSG_TO_STRING(WM_USERCHANGED);
    MSG_TO_STRING(WM_NOTIFYFORMAT);
    MSG_TO_STRING(WM_CONTEXTMENU);
    MSG_TO_STRING(WM_STYLECHANGING);
    MSG_TO_STRING(WM_STYLECHANGED);
    MSG_TO_STRING(WM_DISPLAYCHANGE);
    MSG_TO_STRING(WM_GETICON);
    MSG_TO_STRING(WM_SETICON);
    MSG_TO_STRING(WM_NCCREATE);
    MSG_TO_STRING(WM_NCDESTROY);
    MSG_TO_STRING(WM_NCCALCSIZE);
    MSG_TO_STRING(WM_NCHITTEST);
    MSG_TO_STRING(WM_NCPAINT);
    MSG_TO_STRING(WM_NCACTIVATE);
    MSG_TO_STRING(WM_GETDLGCODE);
    MSG_TO_STRING(WM_NCMOUSEMOVE);
    MSG_TO_STRING(WM_NCLBUTTONDOWN);
    MSG_TO_STRING(WM_NCLBUTTONUP);
    MSG_TO_STRING(WM_NCLBUTTONDBLCLK);
    MSG_TO_STRING(WM_NCRBUTTONDOWN);
    MSG_TO_STRING(WM_NCRBUTTONUP);
    MSG_TO_STRING(WM_NCRBUTTONDBLCLK);
    MSG_TO_STRING(WM_NCMBUTTONDOWN);
    MSG_TO_STRING(WM_NCMBUTTONUP);
    MSG_TO_STRING(WM_NCMBUTTONDBLCLK);
    //MSG_TO_STRING(WM_KEYFIRST);
    MSG_TO_STRING(WM_KEYDOWN);
    MSG_TO_STRING(WM_KEYUP);
    MSG_TO_STRING(WM_CHAR);
    MSG_TO_STRING(WM_DEADCHAR);
    MSG_TO_STRING(WM_SYSKEYDOWN);
    MSG_TO_STRING(WM_SYSKEYUP);
    MSG_TO_STRING(WM_SYSCHAR);
    MSG_TO_STRING(WM_SYSDEADCHAR);
    MSG_TO_STRING(WM_KEYLAST);
    MSG_TO_STRING(WM_IME_STARTCOMPOSITION);
    MSG_TO_STRING(WM_IME_ENDCOMPOSITION);
    MSG_TO_STRING(WM_IME_COMPOSITION);
    //MSG_TO_STRING(WM_IME_KEYLAST);
    MSG_TO_STRING(WM_INITDIALOG);
    MSG_TO_STRING(WM_COMMAND);
    MSG_TO_STRING(WM_SYSCOMMAND);
    MSG_TO_STRING(WM_TIMER);
    MSG_TO_STRING(WM_HSCROLL);
    MSG_TO_STRING(WM_VSCROLL);
    MSG_TO_STRING(WM_INITMENU);
    MSG_TO_STRING(WM_INITMENUPOPUP);
    MSG_TO_STRING(WM_MENUSELECT);
    MSG_TO_STRING(WM_MENUCHAR);
    MSG_TO_STRING(WM_ENTERIDLE);
    MSG_TO_STRING(WM_CTLCOLORMSGBOX);
    MSG_TO_STRING(WM_CTLCOLOREDIT);
    MSG_TO_STRING(WM_CTLCOLORLISTBOX);
    MSG_TO_STRING(WM_CTLCOLORBTN);
    MSG_TO_STRING(WM_CTLCOLORDLG);
    MSG_TO_STRING(WM_CTLCOLORSCROLLBAR);
    MSG_TO_STRING(WM_CTLCOLORSTATIC);
    //MSG_TO_STRING(WM_MOUSEFIRST);
    MSG_TO_STRING(WM_MOUSEMOVE);
    MSG_TO_STRING(WM_LBUTTONDOWN);
    MSG_TO_STRING(WM_LBUTTONUP);
    MSG_TO_STRING(WM_LBUTTONDBLCLK);
    MSG_TO_STRING(WM_RBUTTONDOWN);
    MSG_TO_STRING(WM_RBUTTONUP);
    MSG_TO_STRING(WM_RBUTTONDBLCLK);
    MSG_TO_STRING(WM_MBUTTONDOWN);
    MSG_TO_STRING(WM_MBUTTONUP);
    MSG_TO_STRING(WM_MBUTTONDBLCLK);
    MSG_TO_STRING(WM_MOUSEWHEEL);
    //MSG_TO_STRING(WM_MOUSELAST);
    MSG_TO_STRING(WM_PARENTNOTIFY);
    MSG_TO_STRING(WM_ENTERMENULOOP);
    MSG_TO_STRING(WM_EXITMENULOOP);
    MSG_TO_STRING(WM_NEXTMENU);
    MSG_TO_STRING(WM_SIZING);
    MSG_TO_STRING(WM_CAPTURECHANGED);
    MSG_TO_STRING(WM_MOVING);
    MSG_TO_STRING(WM_POWERBROADCAST);
    MSG_TO_STRING(WM_DEVICECHANGE);
    MSG_TO_STRING(WM_IME_SETCONTEXT);
    MSG_TO_STRING(WM_IME_NOTIFY);
    MSG_TO_STRING(WM_IME_CONTROL);
    MSG_TO_STRING(WM_IME_COMPOSITIONFULL);
    MSG_TO_STRING(WM_IME_SELECT);
    MSG_TO_STRING(WM_IME_CHAR);
    MSG_TO_STRING(WM_IME_KEYDOWN);
    MSG_TO_STRING(WM_IME_KEYUP);
    MSG_TO_STRING(WM_MDICREATE);
    MSG_TO_STRING(WM_MDIDESTROY);
    MSG_TO_STRING(WM_MDIACTIVATE);
    MSG_TO_STRING(WM_MDIRESTORE);
    MSG_TO_STRING(WM_MDINEXT);
    MSG_TO_STRING(WM_MDIMAXIMIZE);
    MSG_TO_STRING(WM_MDITILE);
    MSG_TO_STRING(WM_MDICASCADE);
    MSG_TO_STRING(WM_MDIICONARRANGE);
    MSG_TO_STRING(WM_MDIGETACTIVE);
    MSG_TO_STRING(WM_MDISETMENU);
    MSG_TO_STRING(WM_ENTERSIZEMOVE);
    MSG_TO_STRING(WM_EXITSIZEMOVE);
    MSG_TO_STRING(WM_DROPFILES);
    MSG_TO_STRING(WM_MDIREFRESHMENU);
    MSG_TO_STRING(WM_MOUSEHOVER);
    MSG_TO_STRING(WM_MOUSELEAVE);
    MSG_TO_STRING(WM_CUT);
    MSG_TO_STRING(WM_COPY);
    MSG_TO_STRING(WM_PASTE);
    MSG_TO_STRING(WM_CLEAR);
    MSG_TO_STRING(WM_UNDO);
    MSG_TO_STRING(WM_RENDERFORMAT);
    MSG_TO_STRING(WM_RENDERALLFORMATS);
    MSG_TO_STRING(WM_DESTROYCLIPBOARD);
    MSG_TO_STRING(WM_DRAWCLIPBOARD);
    MSG_TO_STRING(WM_PAINTCLIPBOARD);
    MSG_TO_STRING(WM_VSCROLLCLIPBOARD);
    MSG_TO_STRING(WM_SIZECLIPBOARD);
    MSG_TO_STRING(WM_ASKCBFORMATNAME);
    MSG_TO_STRING(WM_CHANGECBCHAIN);
    MSG_TO_STRING(WM_HSCROLLCLIPBOARD);
    MSG_TO_STRING(WM_QUERYNEWPALETTE);
    MSG_TO_STRING(WM_PALETTEISCHANGING);
    MSG_TO_STRING(WM_PALETTECHANGED);
    MSG_TO_STRING(WM_HOTKEY);
    MSG_TO_STRING(WM_PRINT);
    MSG_TO_STRING(WM_PRINTCLIENT);
    MSG_TO_STRING(WM_HANDHELDFIRST);
    MSG_TO_STRING(WM_HANDHELDLAST);
    MSG_TO_STRING(WM_AFXFIRST);
    MSG_TO_STRING(WM_AFXLAST);
    MSG_TO_STRING(WM_PENWINFIRST);
    MSG_TO_STRING(WM_PENWINLAST);
    MSG_TO_STRING(WM_APP);
    MSG_TO_STRING(WM_USER);
    MSG_TO_STRING(WM_CTLCOLOR);
    // RichEdit messages
    //MSG_TO_STRING(WM_CONTEXTMENU);
    //MSG_TO_STRING(WM_PRINTCLIENT);
    MSG_TO_STRING(EM_GETLIMITTEXT);
    MSG_TO_STRING(EM_POSFROMCHAR);
    MSG_TO_STRING(EM_CHARFROMPOS);
    MSG_TO_STRING(EM_SCROLLCARET);
    //MSG_TO_STRING(EM_CANPASTE);
    //MSG_TO_STRING(EM_DISPLAYBAND);
    //MSG_TO_STRING(EM_EXGETSEL);
    //MSG_TO_STRING(EM_EXLIMITTEXT);
    //MSG_TO_STRING(EM_EXLINEFROMCHAR);
    //MSG_TO_STRING(EM_EXSETSEL);
    //MSG_TO_STRING(EM_FINDTEXT);
    //MSG_TO_STRING(EM_FORMATRANGE);
    //MSG_TO_STRING(EM_GETCHARFORMAT);
    //MSG_TO_STRING(EM_GETEVENTMASK);
    //MSG_TO_STRING(EM_GETOLEINTERFACE);
    //MSG_TO_STRING(EM_GETPARAFORMAT);
    //MSG_TO_STRING(EM_GETSELTEXT);
    //MSG_TO_STRING(EM_HIDESELECTION);
    //MSG_TO_STRING(EM_PASTESPECIAL);
    //MSG_TO_STRING(EM_REQUESTRESIZE);
    //MSG_TO_STRING(EM_SELECTIONTYPE);
    //MSG_TO_STRING(EM_SETBKGNDCOLOR);
    //MSG_TO_STRING(EM_SETCHARFORMAT);
    //MSG_TO_STRING(EM_SETEVENTMASK);
    //MSG_TO_STRING(EM_SETOLECALLBACK);
    //MSG_TO_STRING(EM_SETPARAFORMAT);
    //MSG_TO_STRING(EM_SETTARGETDEVICE);
    //MSG_TO_STRING(EM_STREAMIN);
    //MSG_TO_STRING(EM_STREAMOUT);
    //MSG_TO_STRING(EM_GETTEXTRANGE);
    //MSG_TO_STRING(EM_FINDWORDBREAK);
    //MSG_TO_STRING(EM_SETOPTIONS);
    //MSG_TO_STRING(EM_GETOPTIONS);
    //MSG_TO_STRING(EM_FINDTEXTEX);
    //MSG_TO_STRING(EM_GETWORDBREAKPROCEX);
    //MSG_TO_STRING(EM_SETWORDBREAKPROCEX);
    //// Richedit v2.0 messages
    //MSG_TO_STRING(EM_SETUNDOLIMIT);
    //MSG_TO_STRING(EM_REDO);
    //MSG_TO_STRING(EM_CANREDO);
    //MSG_TO_STRING(EM_GETUNDONAME);
    //MSG_TO_STRING(EM_GETREDONAME);
    //MSG_TO_STRING(EM_STOPGROUPTYPING);
    //MSG_TO_STRING(EM_SETTEXTMODE);
    //MSG_TO_STRING(EM_GETTEXTMODE);
    //MSG_TO_STRING(EM_AUTOURLDETECT);
    //MSG_TO_STRING(EM_GETAUTOURLDETECT);
    //MSG_TO_STRING(EM_SETPALETTE);
    //MSG_TO_STRING(EM_GETTEXTEX);
    //MSG_TO_STRING(EM_GETTEXTLENGTHEX);
    //// Asia specific messages
    //MSG_TO_STRING(EM_SETPUNCTUATION);
    //MSG_TO_STRING(EM_GETPUNCTUATION);
    //MSG_TO_STRING(EM_SETWORDWRAPMODE);
    //MSG_TO_STRING(EM_GETWORDWRAPMODE);
    //MSG_TO_STRING(EM_SETIMECOLOR);
    //MSG_TO_STRING(EM_GETIMECOLOR);
    //MSG_TO_STRING(EM_SETIMEOPTIONS);
    //MSG_TO_STRING(EM_GETIMEOPTIONS);
    //MSG_TO_STRING(EM_CONVPOSITION);
    //MSG_TO_STRING(EM_SETLANGOPTIONS);
    //MSG_TO_STRING(EM_GETLANGOPTIONS);
    //MSG_TO_STRING(EM_GETIMECOMPMODE);
    //MSG_TO_STRING(EM_FINDTEXTW);
    //MSG_TO_STRING(EM_FINDTEXTEXW);
    ////Rich Edit 3.0 Asia msgs
    //MSG_TO_STRING(EM_RECONVERSION);
    //MSG_TO_STRING(EM_SETIMEMODEBIAS);
    //MSG_TO_STRING(EM_GETIMEMODEBIAS);
    //// BiDi Specific messages
    //MSG_TO_STRING(EM_SETBIDIOPTIONS);
    //MSG_TO_STRING(EM_GETBIDIOPTIONS);
    //MSG_TO_STRING(EM_SETTYPOGRAPHYOPTIONS);
    //MSG_TO_STRING(EM_GETTYPOGRAPHYOPTIONS);
    //// Extended Edit style specific messages
    //MSG_TO_STRING(EM_SETEDITSTYLE);
    //MSG_TO_STRING(EM_GETEDITSTYLE);
    default: text = L""; break;
  }

  if (text == L"" && ((msg & /*WM_REFLECT*/WM_USER + 0x1C00) == /*WM_REFLECT*/WM_USER + 0x1C00)) {
    std::wstring subtext = MsgToString(msg - /*WM_REFLECT*/WM_USER + 0x1C00);
    if (subtext == L"") subtext = L"???";
    text = L"WM_REFLECT + " + subtext;
  }
  if (text == L"") {
    text = std::to_wstring(msg);
  }

  return text;
}