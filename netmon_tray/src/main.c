#define WIN32_LEAN_AND_MEAN
#define NOMINMAX
#include <windows.h>
#include <shellapi.h>
#include <windowsx.h>
#include <winsock2.h>
#include <ws2tcpip.h>
#include <iphlpapi.h>
#include <icmpapi.h>
#include <winnetwk.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <stdarg.h>

#ifndef ARRAYSIZE
#define ARRAYSIZE(a) (sizeof(a) / sizeof((a)[0]))
#endif

#define APP_NAME               "NetMon Tray"
#define APP_CLASS              "NetMonTrayMainWnd"
#define POPUP_CLASS            "NetMonTrayPopupWnd"
#define ONLINE_CLASS           "NetMonTrayOnlineWnd"
#define DEFAULT_INI_NAME       "netmon.ini"
#define DEFAULT_CACHE_NAME     "netmon.cache.ini"

#define MAX_COMPUTERS          64
#define MAX_SECTION_TEXT       8192
#define MAX_TEXT               256
#define MAX_NAME               64
#define MAX_ADDR               64
#define MAX_USER               128
#define MAX_PASS               128
#define MAX_MAC                32

#define STATUS_UNKNOWN         (-1)
#define STATUS_OFFLINE         0
#define STATUS_ONLINE          1

#define ID_TRAY_ICON           1
#define WM_TRAYICON            (WM_APP + 1)
#define WM_STATUS_CHANGED      (WM_APP + 2)
#define WM_RELOAD_CONFIG       (WM_APP + 3)

#define ID_MENU_SHOW_ONLINE    5001
#define ID_MENU_RELOAD         5002
#define ID_MENU_EXIT           5003

#define ID_WAKE_BASE           1000
#define ID_SHUTDOWN_BASE       2000
#define ID_RDP_BASE            3000
#define ID_REBOOT_BASE         4000
#define ID_TIMER_POPUP         1

#define ACTION_WAKE            1
#define ACTION_SHUTDOWN        2
#define ACTION_RDP             3
#define ACTION_REBOOT          4

typedef struct Computer {
    char section[MAX_NAME];
    char name[MAX_NAME];
    char address[MAX_ADDR];
    char username[MAX_USER];
    char password[MAX_PASS];
    char mac[MAX_MAC];
    DWORD intervalSec;
    DWORD lastProbeTick;
    int status;
    int allowWake;
    int allowShutdown;
    int allowRdp;
    int allowReboot;
} Computer;

typedef struct ActionRequest {
    int action;
    int index;
} ActionRequest;

static HINSTANCE g_hInst = NULL;
static HWND g_mainWnd = NULL;
static HWND g_popupWnd = NULL;
static HWND g_onlineWnd = NULL;
static HFONT g_uiFont = NULL;
static HANDLE g_monitorThread = NULL;
static volatile LONG g_stopMonitor = 0;
static CRITICAL_SECTION g_lock;
static Computer g_computers[MAX_COMPUTERS];
static int g_computerCount = 0;
static BOOL g_popupOnHover = TRUE;
static BOOL g_alwaysShowOnline = FALSE;
static UINT g_icmpTimeoutMs = 700;
static UINT g_tcpTimeoutMs = 350;
static char g_broadcastAddress[MAX_ADDR] = "255.255.255.255";
static char g_iniPath[MAX_PATH] = {0};
static char g_cachePath[MAX_PATH] = {0};
static DWORD g_lastPopupHoverTick = 0;
static UINT g_taskbarCreatedMsg = 0;
static BOOL g_trayIconAdded = FALSE;
static HBRUSH g_onlineBrush = NULL;
static POINT g_onlineSavedPos = {0, 0};
static BOOL g_onlinePosValid = FALSE;
static DWORD g_lastTrayToggleTick = 0;

static void ShowBalloon(const char* title, const char* text, DWORD infoFlags);
static void UpdateTrayTooltip(void);
static void UpdatePopupControls(void);
static void UpdateOnlineWindowText(void);
static void ToggleOnlineWindow(void);
static void ToggleOnlineWindowFromTray(void);
static void ResizeOnlineWindowToFit(void);
static BOOL LoadConfiguration(void);
static DWORD WINAPI MonitorThreadProc(LPVOID lpParameter);
static void QueueImmediateProbe(int index);
static BOOL AddTrayIcon(void);
static void RemoveTrayIcon(void);
static void ShowContextMenuAt(int x, int y);
static BOOL AskActionConfirmation(int action, int index);
static void FormatWin32MessageText(DWORD code, char* out, size_t outSize);
static BOOL LaunchRdpSession(const Computer* pc, char* errorText, size_t errorSize);
static void ShowActionResult(int action, BOOL ok, const char* text);
static void LoadOnlineWindowPlacement(void);
static void SaveOnlineWindowPlacement(void);
static BOOL RequestRemotePowerAction(const Computer* pc, BOOL reboot, char* errorText, size_t errorSize);
static int ParseIpOctets(const char* address, int octets[4]);
static int CompareComputersByIp(const void* a, const void* b);
static void BuildSortedComputerSnapshot(Computer* out, int* outCount);
static int CountNetworkSeparators(const Computer* list, int count);

static int StrIEquals(const char* a, const char* b) {
    if (!a || !b) return 0;
    return _stricmp(a, b) == 0;
}

static void SafeCopy(char* dst, size_t dstSize, const char* src) {
    if (!dst || dstSize == 0) return;
    if (!src) src = "";
    strncpy(dst, src, dstSize - 1);
    dst[dstSize - 1] = '\0';
}

static void AppendText(char* dst, size_t dstSize, size_t* used, const char* fmt, ...) {
    va_list args;
    int written;
    size_t pos;

    if (!dst || !used || dstSize == 0) return;
    pos = *used;
    if (pos >= dstSize) {
        dst[dstSize - 1] = '\0';
        *used = dstSize - 1;
        return;
    }

    va_start(args, fmt);
    written = vsnprintf(dst + pos, dstSize - pos, fmt, args);
    va_end(args);

    if (written < 0) {
        dst[dstSize - 1] = '\0';
        *used = strlen(dst);
        return;
    }

    if ((size_t)written >= dstSize - pos) {
        *used = dstSize - 1;
    } else {
        *used = pos + (size_t)written;
    }
}


static void TrimTrailingWhitespace(char* text) {
    size_t len;
    if (!text) return;
    len = strlen(text);
    while (len > 0) {
        char ch = text[len - 1];
        if (ch == '\r' || ch == '\n' || ch == ' ' || ch == '\t' || ch == '.') {
            text[len - 1] = '\0';
            --len;
        } else {
            break;
        }
    }
}

static void FormatWin32MessageText(DWORD code, char* out, size_t outSize) {
    DWORD flags;
    DWORD rc;

    if (!out || outSize == 0) return;
    out[0] = '\0';

    flags = FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;
    rc = FormatMessageA(flags,
                        NULL,
                        code,
                        MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT),
                        out,
                        (DWORD)outSize,
                        NULL);
    if (rc == 0) {
        snprintf(out, outSize, "Unknown Windows error %lu", (unsigned long)code);
        return;
    }

    TrimTrailingWhitespace(out);
}


static int ParseIpOctets(const char* address, int octets[4]) {
    unsigned int a, b, c, d;
    if (!address || !octets) return 0;
    if (sscanf(address, "%u.%u.%u.%u", &a, &b, &c, &d) != 4) return 0;
    if (a > 255 || b > 255 || c > 255 || d > 255) return 0;
    octets[0] = (int)a;
    octets[1] = (int)b;
    octets[2] = (int)c;
    octets[3] = (int)d;
    return 1;
}

static int CompareComputersByIp(const void* a, const void* b) {
    const Computer* ca = (const Computer*)a;
    const Computer* cb = (const Computer*)b;
    int ao[4] = {999, 999, 999, 999};
    int bo[4] = {999, 999, 999, 999};

    ParseIpOctets(ca->address, ao);
    ParseIpOctets(cb->address, bo);

    for (int i = 0; i < 4; ++i) {
        if (ao[i] != bo[i]) return ao[i] - bo[i];
    }
    return _stricmp(ca->name[0] ? ca->name : ca->section, cb->name[0] ? cb->name : cb->section);
}

static void BuildSortedComputerSnapshot(Computer* out, int* outCount) {
    int count = 0;

    if (!out || !outCount) return;

    EnterCriticalSection(&g_lock);
    count = g_computerCount;
    if (count > MAX_COMPUTERS) count = MAX_COMPUTERS;
    if (count > 0) {
        memcpy(out, g_computers, sizeof(Computer) * count);
    }
    LeaveCriticalSection(&g_lock);

    if (count > 1) {
        qsort(out, (size_t)count, sizeof(Computer), CompareComputersByIp);
    }

    *outCount = count;
}

static int CountNetworkSeparators(const Computer* list, int count) {
    int prev[4] = {0, 0, -1, 0};
    int cur[4] = {0, 0, -1, 0};
    int separators = 0;

    if (!list || count <= 1) return 0;

    if (!ParseIpOctets(list[0].address, prev)) {
        prev[2] = -1;
    }

    for (int i = 1; i < count; ++i) {
        if (!ParseIpOctets(list[i].address, cur)) {
            cur[2] = -1;
        }
        if (cur[2] != prev[2]) {
            ++separators;
        }
        prev[2] = cur[2];
    }

    return separators;
}

static void BuildAppPaths(void) {
    char exePath[MAX_PATH];
    char* slash;

    GetModuleFileNameA(NULL, exePath, ARRAYSIZE(exePath));
    slash = strrchr(exePath, '\\');
    if (slash) {
        *(slash + 1) = '\0';
    } else {
        exePath[0] = '.';
        exePath[1] = '\0';
    }

    snprintf(g_iniPath, ARRAYSIZE(g_iniPath), "%s%s", exePath, DEFAULT_INI_NAME);
    snprintf(g_cachePath, ARRAYSIZE(g_cachePath), "%s%s", exePath, DEFAULT_CACHE_NAME);
}

static void ReadIniString(const char* section, const char* key, const char* defValue,
                          char* out, DWORD outSize, const char* path) {
    GetPrivateProfileStringA(section, key, defValue, out, outSize, path);
}

static DWORD ReadIniDword(const char* section, const char* key, DWORD defValue, const char* path) {
    return GetPrivateProfileIntA(section, key, (INT)defValue, path);
}

static BOOL WriteIniString(const char* section, const char* key, const char* value, const char* path) {
    return WritePrivateProfileStringA(section, key, value, path);
}

static void LoadOnlineWindowPlacement(void) {
    char bufX[32];
    char bufY[32];

    bufX[0] = '\0';
    bufY[0] = '\0';
    ReadIniString("window", "online_x", "", bufX, ARRAYSIZE(bufX), g_cachePath);
    ReadIniString("window", "online_y", "", bufY, ARRAYSIZE(bufY), g_cachePath);
    if (bufX[0] && bufY[0]) {
        g_onlineSavedPos.x = atoi(bufX);
        g_onlineSavedPos.y = atoi(bufY);
        g_onlinePosValid = TRUE;
    } else {
        g_onlinePosValid = FALSE;
    }
}

static void SaveOnlineWindowPlacement(void) {
    RECT rc;
    char buf[32];

    if (!g_onlineWnd) return;
    if (!GetWindowRect(g_onlineWnd, &rc)) return;

    snprintf(buf, sizeof(buf), "%d", rc.left);
    WriteIniString("window", "online_x", buf, g_cachePath);
    snprintf(buf, sizeof(buf), "%d", rc.top);
    WriteIniString("window", "online_y", buf, g_cachePath);

    g_onlineSavedPos.x = rc.left;
    g_onlineSavedPos.y = rc.top;
    g_onlinePosValid = TRUE;
}

static int ParseMacAddress(const char* text, BYTE mac[6]) {
    unsigned int values[6];
    if (!text || !*text) return 0;
    if (sscanf(text, "%2x:%2x:%2x:%2x:%2x:%2x",
               &values[0], &values[1], &values[2], &values[3], &values[4], &values[5]) == 6 ||
        sscanf(text, "%2x-%2x-%2x-%2x-%2x-%2x",
               &values[0], &values[1], &values[2], &values[3], &values[4], &values[5]) == 6) {
        int i;
        for (i = 0; i < 6; ++i) {
            mac[i] = (BYTE)values[i];
        }
        return 1;
    }
    return 0;
}

static void FormatMacAddress(const BYTE mac[6], char* out, size_t outSize) {
    snprintf(out, outSize, "%02X-%02X-%02X-%02X-%02X-%02X",
             mac[0], mac[1], mac[2], mac[3], mac[4], mac[5]);
}

static BOOL ResolveIPv4(const char* address, DWORD* outAddrNetworkOrder) {
    struct in_addr addr4;
    if (!address || !outAddrNetworkOrder) return FALSE;
    if (InetPtonA(AF_INET, address, &addr4) == 1) {
        *outAddrNetworkOrder = addr4.s_addr;
        return TRUE;
    }
    return FALSE;
}

static BOOL SendArpForMac(const char* address, char* outMac, size_t outMacSize) {
    DWORD ip;
    BYTE macBytes[6];
    ULONG macLen = 6;
    ULONG macBuf[2] = {0, 0};

    if (!ResolveIPv4(address, &ip)) return FALSE;
    if (SendARP(ip, 0, macBuf, &macLen) != NO_ERROR || macLen < 6) return FALSE;

    memcpy(macBytes, macBuf, 6);
    FormatMacAddress(macBytes, outMac, outMacSize);
    return TRUE;
}

static BOOL SaveCachedMac(const char* section, const char* mac) {
    if (!section || !*section || !mac || !*mac) return FALSE;
    return WritePrivateProfileStringA(section, "mac", mac, g_cachePath);
}

static BOOL ProbeHostIcmp(const char* address, UINT timeoutMs) {
    HANDLE hIcmp;
    DWORD ip;
    char sendData[4] = "NM";
    BYTE replyBuffer[128];
    DWORD result;

    if (!ResolveIPv4(address, &ip)) return FALSE;
    hIcmp = IcmpCreateFile();
    if (hIcmp == INVALID_HANDLE_VALUE) return FALSE;

    result = IcmpSendEcho(hIcmp, ip, sendData, (WORD)sizeof(sendData), NULL,
                          replyBuffer, sizeof(replyBuffer), timeoutMs);
    IcmpCloseHandle(hIcmp);
    return result > 0;
}

static BOOL ProbeHostTcp(const char* address, USHORT port, UINT timeoutMs) {
    SOCKET s = INVALID_SOCKET;
    struct sockaddr_in sa;
    u_long nonBlocking = 1;
    fd_set writeSet;
    fd_set errSet;
    TIMEVAL tv;
    int rc;
    int soErr = 0;
    int soErrLen = (int)sizeof(soErr);

    ZeroMemory(&sa, sizeof(sa));
    sa.sin_family = AF_INET;
    sa.sin_port = htons(port);
    if (InetPtonA(AF_INET, address, &sa.sin_addr) != 1) return FALSE;

    s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if (s == INVALID_SOCKET) return FALSE;

    ioctlsocket(s, FIONBIO, &nonBlocking);
    connect(s, (struct sockaddr*)&sa, sizeof(sa));

    FD_ZERO(&writeSet);
    FD_ZERO(&errSet);
    FD_SET(s, &writeSet);
    FD_SET(s, &errSet);

    tv.tv_sec = timeoutMs / 1000;
    tv.tv_usec = (timeoutMs % 1000) * 1000;

    rc = select(0, NULL, &writeSet, &errSet, &tv);
    if (rc > 0 && FD_ISSET(s, &writeSet)) {
        getsockopt(s, SOL_SOCKET, SO_ERROR, (char*)&soErr, &soErrLen);
        closesocket(s);
        return soErr == 0;
    }

    closesocket(s);
    return FALSE;
}

static int ProbeComputerStatus(const char* address) {
    if (ProbeHostIcmp(address, g_icmpTimeoutMs)) return STATUS_ONLINE;
    if (ProbeHostTcp(address, 445, g_tcpTimeoutMs)) return STATUS_ONLINE;
    if (ProbeHostTcp(address, 135, g_tcpTimeoutMs)) return STATUS_ONLINE;
    return STATUS_OFFLINE;
}

static void DeriveBroadcastFromIp(const char* ip, char* out, size_t outSize) {
    unsigned int a, b, c, d;
    if (sscanf(ip, "%u.%u.%u.%u", &a, &b, &c, &d) == 4) {
        snprintf(out, outSize, "%u.%u.%u.255", a, b, c);
    } else {
        SafeCopy(out, outSize, "255.255.255.255");
    }
}

static BOOL SendWakePacket(const char* macText, const char* preferredBroadcast) {
    BYTE mac[6];
    BYTE packet[102];
    SOCKET s;
    BOOL ok = FALSE;
    BOOL broadcastOpt = TRUE;
    struct sockaddr_in sa;
    char broadcastIp[MAX_ADDR];
    int i;

    if (!ParseMacAddress(macText, mac)) return FALSE;

    memset(packet, 0xFF, 6);
    for (i = 0; i < 16; ++i) {
        memcpy(packet + 6 + (i * 6), mac, 6);
    }

    SafeCopy(broadcastIp, ARRAYSIZE(broadcastIp), preferredBroadcast);
    if (broadcastIp[0] == '\0') {
        SafeCopy(broadcastIp, ARRAYSIZE(broadcastIp), "255.255.255.255");
    }

    s = socket(AF_INET, SOCK_DGRAM, IPPROTO_UDP);
    if (s == INVALID_SOCKET) return FALSE;

    setsockopt(s, SOL_SOCKET, SO_BROADCAST, (const char*)&broadcastOpt, sizeof(broadcastOpt));
    ZeroMemory(&sa, sizeof(sa));
    sa.sin_family = AF_INET;
    sa.sin_port = htons(9);
    if (InetPtonA(AF_INET, broadcastIp, &sa.sin_addr) != 1) {
        closesocket(s);
        return FALSE;
    }

    if (sendto(s, (const char*)packet, sizeof(packet), 0, (const struct sockaddr*)&sa, sizeof(sa)) == (int)sizeof(packet)) {
        ok = TRUE;
    }

    sa.sin_port = htons(7);
    sendto(s, (const char*)packet, sizeof(packet), 0, (const struct sockaddr*)&sa, sizeof(sa));
    closesocket(s);
    return ok;
}

static BOOL RunHiddenProcessAndWait(const char* commandLine, DWORD timeoutMs, DWORD* outExitCode) {
    STARTUPINFOA si;
    PROCESS_INFORMATION pi;
    char cmd[512];
    DWORD exitCode = (DWORD)-1;
    DWORD waitRc;
    BOOL ok;

    ZeroMemory(&si, sizeof(si));
    ZeroMemory(&pi, sizeof(pi));
    si.cb = sizeof(si);
    si.dwFlags = STARTF_USESHOWWINDOW;
    si.wShowWindow = SW_HIDE;

    SafeCopy(cmd, sizeof(cmd), commandLine);
    ok = CreateProcessA(NULL, cmd, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi);
    if (!ok) return FALSE;

    waitRc = WaitForSingleObject(pi.hProcess, timeoutMs);
    if (waitRc == WAIT_TIMEOUT) {
        TerminateProcess(pi.hProcess, 0x102);
        exitCode = WAIT_TIMEOUT;
    } else if (!GetExitCodeProcess(pi.hProcess, &exitCode)) {
        exitCode = (DWORD)-1;
    }

    if (outExitCode) *outExitCode = exitCode;
    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return TRUE;
}

static BOOL RequestRemotePowerAction(const Computer* pc, BOOL reboot, char* errorText, size_t errorSize) {
    NETRESOURCEA nr;
    char remotePath[MAX_ADDR + 16];
    char command[512];
    char sysMsg[256];
    DWORD addRc;
    DWORD exitCode = (DWORD)-1;
    DWORD launchErr;
    const char* actionWord = reboot ? "restart" : "shutdown";

    if (!pc || !pc->address[0]) {
        SafeCopy(errorText, errorSize, "Missing host address.");
        return FALSE;
    }

    snprintf(remotePath, sizeof(remotePath), "\\\\%s\\IPC$", pc->address);
    ZeroMemory(&nr, sizeof(nr));
    nr.dwType = RESOURCETYPE_ANY;
    nr.lpRemoteName = remotePath;

    WNetCancelConnection2A(remotePath, 0, TRUE);
    addRc = WNetAddConnection2A(&nr,
                                pc->password[0] ? pc->password : NULL,
                                pc->username[0] ? pc->username : NULL,
                                CONNECT_TEMPORARY);

    if (addRc != NO_ERROR && addRc != ERROR_ALREADY_ASSIGNED && addRc != ERROR_DEVICE_ALREADY_REMEMBERED) {
        FormatWin32MessageText(addRc, sysMsg, sizeof(sysMsg));
        snprintf(errorText, errorSize,
                 "IPC$ authentication failed for %s (code %lu: %s). Check host, login, password, SMB access and firewall.",
                 pc->address,
                 (unsigned long)addRc,
                 sysMsg);
        return FALSE;
    }

    snprintf(command, sizeof(command), "shutdown.exe /m \\\\%s /%c /t 0 /f", pc->address, reboot ? 'r' : 's');
    if (!RunHiddenProcessAndWait(command, 15000, &exitCode)) {
        launchErr = GetLastError();
        FormatWin32MessageText(launchErr, sysMsg, sizeof(sysMsg));
        snprintf(errorText, errorSize,
                 "Unable to launch shutdown.exe for %s (error %lu: %s).",
                 pc->address,
                 (unsigned long)launchErr,
                 sysMsg);
        WNetCancelConnection2A(remotePath, 0, TRUE);
        return FALSE;
    }

    WNetCancelConnection2A(remotePath, 0, TRUE);
    if (exitCode == WAIT_TIMEOUT) {
        snprintf(errorText, errorSize,
                 "shutdown.exe timed out while contacting %s. The host may be unreachable or blocked by firewall/RPC.",
                 pc->address);
        return FALSE;
    }
    if (exitCode != 0) {
        FormatWin32MessageText(exitCode, sysMsg, sizeof(sysMsg));
        snprintf(errorText, errorSize,
                 "Remote %s failed for %s. shutdown.exe returned code %lu: %s. Common causes: access denied, Remote Shutdown privilege missing, firewall, RPC, or UAC remote restrictions.",
                 actionWord,
                 pc->address,
                 (unsigned long)exitCode,
                 sysMsg);
        return FALSE;
    }

    return TRUE;
}

static void ComputerToStatusText(const Computer* pc, char* out, size_t outSize) {
    const char* state = "UNKNOWN";
    if (pc->status == STATUS_ONLINE) state = "ONLINE";
    else if (pc->status == STATUS_OFFLINE) state = "OFFLINE";

    snprintf(out, outSize, "%s [%s] : %s", pc->name[0] ? pc->name : pc->section,
             pc->address, state);
}

static void UpdateTrayTooltip(void) {
    int online = 0;
    int total;
    char tip[128];
    NOTIFYICONDATAA nid;

    if (!g_trayIconAdded || !g_mainWnd) return;

    EnterCriticalSection(&g_lock);
    total = g_computerCount;
    for (int i = 0; i < total; ++i) {
        if (g_computers[i].status == STATUS_ONLINE) ++online;
    }
    LeaveCriticalSection(&g_lock);

    snprintf(tip, sizeof(tip), APP_NAME " - online %d/%d", online, total);
    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = g_mainWnd;
    nid.uID = ID_TRAY_ICON;
    nid.uFlags = NIF_TIP;
    SafeCopy(nid.szTip, ARRAYSIZE(nid.szTip), tip);
    Shell_NotifyIconA(NIM_MODIFY, &nid);
}

static void ShowBalloon(const char* title, const char* text, DWORD infoFlags) {
    NOTIFYICONDATAA nid;

    if (!g_trayIconAdded || !g_mainWnd) return;

    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = g_mainWnd;
    nid.uID = ID_TRAY_ICON;
    nid.uFlags = NIF_INFO;
    nid.dwInfoFlags = infoFlags;
    SafeCopy(nid.szInfoTitle, ARRAYSIZE(nid.szInfoTitle), title ? title : APP_NAME);
    SafeCopy(nid.szInfo, ARRAYSIZE(nid.szInfo), text ? text : "");
    Shell_NotifyIconA(NIM_MODIFY, &nid);
}

static void UpdatePopupControls(void) {
    (void)0;
}

static void UpdateOnlineWindowText(void) {
    if (!g_onlineWnd) return;
    ResizeOnlineWindowToFit();
    InvalidateRect(g_onlineWnd, NULL, TRUE);
}

static void ResizeOnlineWindowToFit(void) {
    Computer list[MAX_COMPUTERS];
    int count = 0;
    int separators = 0;
    HDC hdc;
    HFONT oldFont = NULL;
    SIZE sz;
    int width = 160;
    int clientWidth;
    int clientHeight;
    UINT swpFlags;
    int maxWidth = 480;
    const int border = 1;
    const int paddingX = 10;
    const int paddingY = 8;
    const int rowHeight = 18;
    const int separatorGap = 8;
    const char* emptyText = "No computers configured.";

    if (!g_onlineWnd) return;

    BuildSortedComputerSnapshot(list, &count);
    separators = CountNetworkSeparators(list, count);

    hdc = GetDC(g_onlineWnd);
    if (!hdc) return;
    if (g_uiFont) oldFont = (HFONT)SelectObject(hdc, g_uiFont);

    if (count <= 0) {
        GetTextExtentPoint32A(hdc, emptyText, (int)strlen(emptyText), &sz);
        width = sz.cx;
    } else {
        for (int i = 0; i < count; ++i) {
            const char* name = list[i].name[0] ? list[i].name : list[i].section;
            if (GetTextExtentPoint32A(hdc, name, (int)strlen(name), &sz) && sz.cx > width) {
                width = sz.cx;
            }
        }
    }

    if (oldFont) SelectObject(hdc, oldFont);
    ReleaseDC(g_onlineWnd, hdc);

    if (width < 160) width = 160;
    if (width > maxWidth) width = maxWidth;

    clientWidth = border * 2 + paddingX * 2 + width;
    if (count <= 0) {
        clientHeight = border * 2 + paddingY * 2 + rowHeight;
    } else {
        clientHeight = border * 2 + paddingY * 2 + count * rowHeight + separators * separatorGap;
    }

    swpFlags = SWP_NOMOVE | SWP_NOACTIVATE;
    if (IsWindowVisible(g_onlineWnd) || g_alwaysShowOnline) swpFlags |= SWP_SHOWWINDOW;
    SetWindowPos(g_onlineWnd, HWND_TOPMOST, 0, 0, clientWidth, clientHeight, swpFlags);
}

static void ToggleOnlineWindow(void) {
    if (!g_onlineWnd) return;
    g_alwaysShowOnline = !IsWindowVisible(g_onlineWnd);
    if (g_alwaysShowOnline) {
        UpdateOnlineWindowText();
        if (g_onlinePosValid) {
            SetWindowPos(g_onlineWnd, HWND_TOPMOST, g_onlineSavedPos.x, g_onlineSavedPos.y, 0, 0,
                         SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        } else {
            ShowWindow(g_onlineWnd, SW_SHOWNOACTIVATE);
            SetWindowPos(g_onlineWnd, HWND_TOPMOST, 0, 0, 0, 0,
                         SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        }
    } else {
        SaveOnlineWindowPlacement();
        ShowWindow(g_onlineWnd, SW_HIDE);
    }
}

static void ToggleOnlineWindowFromTray(void) {
    DWORD now = GetTickCount();
    if (g_lastTrayToggleTick != 0 && (now - g_lastTrayToggleTick) < 250) {
        return;
    }
    g_lastTrayToggleTick = now;
    ToggleOnlineWindow();
}

static void QueueImmediateProbe(int index) {

    EnterCriticalSection(&g_lock);
    if (index >= 0 && index < g_computerCount) {
        g_computers[index].lastProbeTick = 0;
    }
    LeaveCriticalSection(&g_lock);
}

static BOOL AskActionConfirmation(int action, int index) {
    Computer pc;
    char text[512];
    const char* target;
    HWND owner;

    ZeroMemory(&pc, sizeof(pc));
    EnterCriticalSection(&g_lock);
    if (index >= 0 && index < g_computerCount) {
        pc = g_computers[index];
    }
    LeaveCriticalSection(&g_lock);

    target = pc.name[0] ? pc.name : pc.address;
    owner = g_mainWnd;

    if (action == ACTION_SHUTDOWN) {
        snprintf(text, sizeof(text),
                 "Power off '%s' (%s)?\r\n\r\nThis will send a remote shutdown command immediately.",
                 target,
                 pc.address[0] ? pc.address : "unknown host");
    } else if (action == ACTION_REBOOT) {
        snprintf(text, sizeof(text),
                 "Reboot '%s' (%s)?\r\n\r\nThis will send a remote restart command immediately.",
                 target,
                 pc.address[0] ? pc.address : "unknown host");
    } else {
        snprintf(text, sizeof(text),
                 "Power on '%s' (%s)?\r\n\r\nThis will send a Wake-on-LAN packet immediately.",
                 target,
                 pc.address[0] ? pc.address : "unknown host");
    }

    return MessageBoxA(owner,
                       text,
                       action == ACTION_SHUTDOWN ? "Confirm power off" :
                       (action == ACTION_REBOOT ? "Confirm reboot" : "Confirm power on"),
                       MB_YESNO | MB_ICONQUESTION | MB_DEFBUTTON2 | MB_TOPMOST) == IDYES;
}

static void ShowContextMenuAt(int x, int y) {
    HMENU menu = CreatePopupMenu();
    POINT pt;
    int count;
    Computer snapshot[MAX_COMPUTERS];
    UINT cmd;

    if (!menu) return;
    if ((x <= 0 && y <= 0) || x < 0 || y < 0) {
        GetCursorPos(&pt);
        x = pt.x;
        y = pt.y;
    }

    AppendMenuA(menu, MF_STRING | (IsWindowVisible(g_onlineWnd) ? MF_CHECKED : 0), ID_MENU_SHOW_ONLINE,
                IsWindowVisible(g_onlineWnd) ? "Hide online list" : "Show online list");
    AppendMenuA(menu, MF_STRING, ID_MENU_RELOAD, "Reload config");
    AppendMenuA(menu, MF_SEPARATOR, 0, NULL);

    EnterCriticalSection(&g_lock);
    count = g_computerCount;
    if (count > MAX_COMPUTERS) count = MAX_COMPUTERS;
    memcpy(snapshot, g_computers, sizeof(Computer) * count);
    LeaveCriticalSection(&g_lock);

    if (count == 0) {
        AppendMenuA(menu, MF_GRAYED, 0, "No computers configured");
    } else {
        for (int i = 0; i < count; ++i) {
            HMENU sub = CreatePopupMenu();
            char title[128];
            char ipLine[128];
            UINT flags;

            if (!sub) continue;
            snprintf(title, sizeof(title), "%s%s", snapshot[i].name[0] ? snapshot[i].name : snapshot[i].section,
                     snapshot[i].status == STATUS_ONLINE ? " [ON]" : (snapshot[i].status == STATUS_OFFLINE ? " [OFF]" : " [??]"));
            snprintf(ipLine, sizeof(ipLine), "IP: %s", snapshot[i].address[0] ? snapshot[i].address : "n/a");
            AppendMenuA(sub, MF_GRAYED, 0, ipLine);
            AppendMenuA(sub, MF_SEPARATOR, 0, NULL);

            if (snapshot[i].allowWake) {
                flags = MF_STRING;
                if (!snapshot[i].mac[0]) flags |= MF_GRAYED;
                AppendMenuA(sub, flags, ID_WAKE_BASE + i, "On");
            }
            if (snapshot[i].allowShutdown) {
                flags = MF_STRING;
                if (snapshot[i].status != STATUS_ONLINE) flags |= MF_GRAYED;
                AppendMenuA(sub, flags, ID_SHUTDOWN_BASE + i, "Off");
            }
            if (snapshot[i].allowReboot) {
                flags = MF_STRING;
                if (snapshot[i].status != STATUS_ONLINE) flags |= MF_GRAYED;
                AppendMenuA(sub, flags, ID_REBOOT_BASE + i, "Reboot");
            }
            if (snapshot[i].allowRdp) {
                flags = MF_STRING;
                if (snapshot[i].status != STATUS_ONLINE) flags |= MF_GRAYED;
                AppendMenuA(sub, flags, ID_RDP_BASE + i, "RDP");
            }
            if (!snapshot[i].allowWake && !snapshot[i].allowShutdown && !snapshot[i].allowReboot && !snapshot[i].allowRdp) {
                AppendMenuA(sub, MF_GRAYED, 0, "No actions enabled");
            }

            AppendMenuA(menu, MF_POPUP, (UINT_PTR)sub, title);
        }
    }

    AppendMenuA(menu, MF_SEPARATOR, 0, NULL);
    AppendMenuA(menu, MF_STRING, ID_MENU_EXIT, "Exit");

    SetForegroundWindow(g_mainWnd);
    cmd = TrackPopupMenu(menu,
                         TPM_RIGHTBUTTON | TPM_BOTTOMALIGN | TPM_LEFTALIGN | TPM_RETURNCMD,
                         x, y, 0, g_mainWnd, NULL);
    if (cmd != 0) {
        PostMessageA(g_mainWnd, WM_COMMAND, MAKEWPARAM(cmd, 0), 0);
    }
    PostMessageA(g_mainWnd, WM_NULL, 0, 0);
    DestroyMenu(menu);
}

static BOOL AddTrayIcon(void) {
    NOTIFYICONDATAA nid;

    if (!g_mainWnd) return FALSE;

    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = g_mainWnd;
    nid.uID = ID_TRAY_ICON;
    nid.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
    nid.uCallbackMessage = WM_TRAYICON;
    nid.hIcon = LoadIconA(NULL, IDI_APPLICATION);
    SafeCopy(nid.szTip, ARRAYSIZE(nid.szTip), APP_NAME);

    if (!Shell_NotifyIconA(NIM_ADD, &nid)) {
        return FALSE;
    }

    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = g_mainWnd;
    nid.uID = ID_TRAY_ICON;
    nid.uVersion = NOTIFYICON_VERSION_4;
    Shell_NotifyIconA(NIM_SETVERSION, &nid);

    g_trayIconAdded = TRUE;
    return TRUE;
}

static void RemoveTrayIcon(void) {
    NOTIFYICONDATAA nid;

    if (!g_trayIconAdded || !g_mainWnd) return;

    ZeroMemory(&nid, sizeof(nid));
    nid.cbSize = sizeof(nid);
    nid.hWnd = g_mainWnd;
    nid.uID = ID_TRAY_ICON;
    Shell_NotifyIconA(NIM_DELETE, &nid);
    g_trayIconAdded = FALSE;
}


static BOOL LaunchRdpSession(const Computer* pc, char* errorText, size_t errorSize) {
    char command[512];
    STARTUPINFOA si;
    PROCESS_INFORMATION pi;
    BOOL ok;
    DWORD err;
    char sysMsg[256];

    if (!pc || !pc->address[0]) {
        SafeCopy(errorText, errorSize, "RDP launch failed: missing target address.");
        return FALSE;
    }

    snprintf(command, sizeof(command), "mstsc.exe /v:%s", pc->address);
    ZeroMemory(&si, sizeof(si));
    ZeroMemory(&pi, sizeof(pi));
    si.cb = sizeof(si);

    ok = CreateProcessA(NULL, command, NULL, NULL, FALSE, 0, NULL, NULL, &si, &pi);
    if (!ok) {
        err = GetLastError();
        FormatWin32MessageText(err, sysMsg, sizeof(sysMsg));
        snprintf(errorText, errorSize,
                 "Unable to start mstsc.exe for %s (error %lu: %s).",
                 pc->address,
                 (unsigned long)err,
                 sysMsg);
        return FALSE;
    }

    CloseHandle(pi.hThread);
    CloseHandle(pi.hProcess);
    return TRUE;
}

static void ShowActionResult(int action, BOOL ok, const char* text) {
    const char* title = APP_NAME;
    UINT iconFlags = NIIF_INFO;
    UINT boxFlags = MB_OK | MB_TOPMOST | MB_ICONINFORMATION;

    if (action == ACTION_WAKE) {
        title = ok ? "Power on result" : "Power on error";
    } else if (action == ACTION_SHUTDOWN) {
        title = ok ? "Power off result" : "Power off error";
    } else if (action == ACTION_REBOOT) {
        title = ok ? "Reboot result" : "Reboot error";
    } else if (action == ACTION_RDP) {
        title = ok ? "RDP result" : "RDP error";
    }

    if (!ok) {
        iconFlags = NIIF_ERROR;
        boxFlags = MB_OK | MB_TOPMOST | MB_ICONERROR;
    }

    ShowBalloon(title, text, iconFlags);
    MessageBoxA(g_mainWnd ? g_mainWnd : NULL, text, title, boxFlags);
}

static DWORD WINAPI ActionThreadProc(LPVOID lpParameter) {
    ActionRequest* req = (ActionRequest*)lpParameter;
    Computer pc;
    char statusText[256];
    char errText[512];
    BOOL ok = FALSE;
    char broadcast[MAX_ADDR];

    ZeroMemory(&pc, sizeof(pc));
    ZeroMemory(statusText, sizeof(statusText));
    ZeroMemory(errText, sizeof(errText));
    ZeroMemory(broadcast, sizeof(broadcast));

    if (!req) return 0;

    EnterCriticalSection(&g_lock);
    if (req->index >= 0 && req->index < g_computerCount) {
        pc = g_computers[req->index];
        SafeCopy(broadcast, sizeof(broadcast), g_broadcastAddress);
    }
    LeaveCriticalSection(&g_lock);

    if (req->action == ACTION_WAKE) {
        if (!pc.mac[0]) {
            SafeCopy(errText, sizeof(errText), "Wake-on-LAN requires a MAC address. Add mac=... to the INI or let the app learn it while the PC is online.");
        } else {
            if (!broadcast[0]) DeriveBroadcastFromIp(pc.address, broadcast, sizeof(broadcast));
            ok = SendWakePacket(pc.mac, broadcast);
            if (!ok && StrIEquals(broadcast, "255.255.255.255")) {
                DeriveBroadcastFromIp(pc.address, broadcast, sizeof(broadcast));
                ok = SendWakePacket(pc.mac, broadcast);
            }
        }
        if (ok) {
            snprintf(statusText, sizeof(statusText), "Wake-on-LAN packet sent successfully to %s (%s).", pc.name[0] ? pc.name : pc.address, pc.address);
            QueueImmediateProbe(req->index);
            ShowActionResult(req->action, TRUE, statusText);
        } else {
            if (!errText[0]) SafeCopy(errText, sizeof(errText), "Unable to send wake packet.");
            ShowActionResult(req->action, FALSE, errText);
        }
    } else if (req->action == ACTION_SHUTDOWN) {
        ok = RequestRemotePowerAction(&pc, FALSE, errText, sizeof(errText));
        if (ok) {
            snprintf(statusText, sizeof(statusText), "Remote shutdown command sent successfully to %s (%s).", pc.name[0] ? pc.name : pc.address, pc.address);
            QueueImmediateProbe(req->index);
            ShowActionResult(req->action, TRUE, statusText);
        } else {
            if (!errText[0]) SafeCopy(errText, sizeof(errText), "Unable to request remote shutdown.");
            ShowActionResult(req->action, FALSE, errText);
        }
    } else if (req->action == ACTION_REBOOT) {
        ok = RequestRemotePowerAction(&pc, TRUE, errText, sizeof(errText));
        if (ok) {
            snprintf(statusText, sizeof(statusText), "Remote reboot command sent successfully to %s (%s).", pc.name[0] ? pc.name : pc.address, pc.address);
            QueueImmediateProbe(req->index);
            ShowActionResult(req->action, TRUE, statusText);
        } else {
            if (!errText[0]) SafeCopy(errText, sizeof(errText), "Unable to request remote reboot.");
            ShowActionResult(req->action, FALSE, errText);
        }
    } else if (req->action == ACTION_RDP) {
        ok = LaunchRdpSession(&pc, errText, sizeof(errText));
        if (ok) {
            snprintf(statusText, sizeof(statusText), "RDP client started for %s (%s).", pc.name[0] ? pc.name : pc.address, pc.address);
            ShowActionResult(req->action, TRUE, statusText);
        } else {
            if (!errText[0]) SafeCopy(errText, sizeof(errText), "Unable to start RDP client.");
            ShowActionResult(req->action, FALSE, errText);
        }
    }

    HeapFree(GetProcessHeap(), 0, req);
    return 0;
}

static void StartActionAsync(int action, int index) {
    ActionRequest* req = (ActionRequest*)HeapAlloc(GetProcessHeap(), HEAP_ZERO_MEMORY, sizeof(ActionRequest));
    HANDLE hThread;
    if (!req) return;
    req->action = action;
    req->index = index;
    hThread = CreateThread(NULL, 0, ActionThreadProc, req, 0, NULL);
    if (hThread) {
        CloseHandle(hThread);
    } else {
        HeapFree(GetProcessHeap(), 0, req);
        ShowActionResult(action, FALSE, "Unable to start background action thread.");
    }
}

static BOOL LoadConfiguration(void) {
    char sectionNames[MAX_SECTION_TEXT];
    Computer oldComputers[MAX_COMPUTERS];
    int oldCount = 0;
    DWORD len;
    char* current;

    if (GetFileAttributesA(g_iniPath) == INVALID_FILE_ATTRIBUTES) {
        char msg[512];
        snprintf(msg, sizeof(msg), "Configuration file not found:\n%s", g_iniPath);
        MessageBoxA(NULL, msg, APP_NAME, MB_ICONERROR | MB_OK);
        return FALSE;
    }

    EnterCriticalSection(&g_lock);
    oldCount = g_computerCount;
    if (oldCount > MAX_COMPUTERS) oldCount = MAX_COMPUTERS;
    memcpy(oldComputers, g_computers, sizeof(Computer) * oldCount);
    ZeroMemory(g_computers, sizeof(g_computers));
    g_computerCount = 0;
    g_popupOnHover = ReadIniDword("general", "popup_on_hover", 1, g_iniPath) ? TRUE : FALSE;
    g_alwaysShowOnline = ReadIniDword("general", "always_show_online", 0, g_iniPath) ? TRUE : FALSE;
    g_icmpTimeoutMs = ReadIniDword("general", "icmp_timeout_ms", 700, g_iniPath);
    g_tcpTimeoutMs = ReadIniDword("general", "tcp_timeout_ms", 350, g_iniPath);
    ReadIniString("general", "broadcast", "255.255.255.255", g_broadcastAddress, ARRAYSIZE(g_broadcastAddress), g_iniPath);
    LeaveCriticalSection(&g_lock);

    ZeroMemory(sectionNames, sizeof(sectionNames));
    len = GetPrivateProfileSectionNamesA(sectionNames, sizeof(sectionNames), g_iniPath);
    if (len == 0) {
        ShowBalloon(APP_NAME, "No INI sections found.", NIIF_WARNING);
        return FALSE;
    }

    current = sectionNames;
    while (*current) {
        if (!StrIEquals(current, "general") && g_computerCount < MAX_COMPUTERS) {
            Computer pc;
            ZeroMemory(&pc, sizeof(pc));
            SafeCopy(pc.section, ARRAYSIZE(pc.section), current);
            ReadIniString(current, "name", current, pc.name, ARRAYSIZE(pc.name), g_iniPath);
            ReadIniString(current, "address", "", pc.address, ARRAYSIZE(pc.address), g_iniPath);
            ReadIniString(current, "username", "", pc.username, ARRAYSIZE(pc.username), g_iniPath);
            ReadIniString(current, "password", "", pc.password, ARRAYSIZE(pc.password), g_iniPath);
            ReadIniString(current, "mac", "", pc.mac, ARRAYSIZE(pc.mac), g_iniPath);
            if (!pc.mac[0]) {
                ReadIniString(current, "mac", "", pc.mac, ARRAYSIZE(pc.mac), g_cachePath);
            }
            pc.intervalSec = ReadIniDword(current, "interval_sec", 5, g_iniPath);
            if (pc.intervalSec == 0) pc.intervalSec = 5;
            pc.allowWake = ReadIniDword(current, "allow_wake", 1, g_iniPath) ? 1 : 0;
            pc.allowShutdown = ReadIniDword(current, "allow_shutdown", 1, g_iniPath) ? 1 : 0;
            pc.allowRdp = ReadIniDword(current, "allow_rdp", 1, g_iniPath) ? 1 : 0;
            pc.allowReboot = ReadIniDword(current, "allow_reboot", 1, g_iniPath) ? 1 : 0;
            pc.status = STATUS_UNKNOWN;
            pc.lastProbeTick = 0;

            for (int i = 0; i < oldCount; ++i) {
                if (StrIEquals(oldComputers[i].section, current)) {
                    pc.status = oldComputers[i].status;
                    if (!pc.mac[0] && oldComputers[i].mac[0]) {
                        SafeCopy(pc.mac, ARRAYSIZE(pc.mac), oldComputers[i].mac);
                    }
                    break;
                }
            }

            if (pc.address[0]) {
                EnterCriticalSection(&g_lock);
                g_computers[g_computerCount++] = pc;
                LeaveCriticalSection(&g_lock);
            }
        }
        current += strlen(current) + 1;
    }

    UpdateOnlineWindowText();
    UpdateTrayTooltip();

    if (g_onlineWnd) {
        if (g_alwaysShowOnline && g_onlinePosValid) {
            SetWindowPos(g_onlineWnd, HWND_TOPMOST, g_onlineSavedPos.x, g_onlineSavedPos.y, 0, 0,
                         SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        } else {
            ShowWindow(g_onlineWnd, g_alwaysShowOnline ? SW_SHOWNOACTIVATE : SW_HIDE);
            if (g_alwaysShowOnline) {
                SetWindowPos(g_onlineWnd, HWND_TOPMOST, 0, 0, 0, 0,
                             SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
            }
        }
    }

    if (g_computerCount == 0) {
        ShowBalloon(APP_NAME, "No computers configured in netmon.ini.", NIIF_WARNING);
    }
    return TRUE;
}

static DWORD WINAPI MonitorThreadProc(LPVOID lpParameter) {
    (void)lpParameter;
    while (!InterlockedCompareExchange(&g_stopMonitor, 0, 0)) {
        DWORD now = GetTickCount();
        for (int i = 0; i < MAX_COMPUTERS; ++i) {
            Computer snapshot;
            BOOL shouldProbe = FALSE;
            int oldStatus;
            int newStatus;
            char newMac[MAX_MAC] = {0};

            ZeroMemory(&snapshot, sizeof(snapshot));
            EnterCriticalSection(&g_lock);
            if (i < g_computerCount) {
                snapshot = g_computers[i];
                if (snapshot.address[0] && (snapshot.lastProbeTick == 0 || (now - snapshot.lastProbeTick) >= snapshot.intervalSec * 1000UL)) {
                    g_computers[i].lastProbeTick = now;
                    shouldProbe = TRUE;
                }
            }
            LeaveCriticalSection(&g_lock);

            if (!shouldProbe) continue;

            newStatus = ProbeComputerStatus(snapshot.address);
            if (newStatus == STATUS_ONLINE && !snapshot.mac[0] && SendArpForMac(snapshot.address, newMac, sizeof(newMac))) {
                EnterCriticalSection(&g_lock);
                if (i < g_computerCount && !g_computers[i].mac[0]) {
                    SafeCopy(g_computers[i].mac, ARRAYSIZE(g_computers[i].mac), newMac);
                    SaveCachedMac(g_computers[i].section, newMac);
                }
                LeaveCriticalSection(&g_lock);
            }

            EnterCriticalSection(&g_lock);
            if (i >= g_computerCount) {
                LeaveCriticalSection(&g_lock);
                continue;
            }
            oldStatus = g_computers[i].status;
            if (oldStatus != newStatus) {
                g_computers[i].status = newStatus;
                LeaveCriticalSection(&g_lock);
                PostMessageA(g_mainWnd, WM_STATUS_CHANGED, (WPARAM)i,
                             (LPARAM)(((oldStatus & 0xFFFF) << 16) | (newStatus & 0xFFFF)));
            } else {
                LeaveCriticalSection(&g_lock);
            }
        }
        Sleep(250);
    }
    return 0;
}

static LRESULT CALLBACK PopupWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    (void)wParam;
    (void)lParam;
    switch (msg) {
    case WM_CLOSE:
        ShowWindow(hwnd, SW_HIDE);
        return 0;
    }
    return DefWindowProcA(hwnd, msg, wParam, lParam);
}

static LRESULT CALLBACK OnlineWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE:
        return 0;
    case WM_NCHITTEST:
        return HTCAPTION;
    case WM_RBUTTONDOWN:
    case WM_RBUTTONUP: {
        POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
        ClientToScreen(hwnd, &pt);
        ShowContextMenuAt(pt.x, pt.y);
        return 0;
    }
    case WM_CONTEXTMENU:
    case WM_NCRBUTTONUP:
        ShowContextMenuAt(GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
        return 0;
    case WM_ERASEBKGND:
        return 1;
    case WM_PAINT: {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        RECT rc;
        Computer list[MAX_COMPUTERS];
        int count = 0;
        int prevOctets[4] = {0, 0, -1, 0};
        int octets[4] = {0, 0, -1, 0};
        int y;
        const int border = 1;
        const int paddingX = 10;
        const int paddingY = 8;
        const int rowHeight = 18;
        const int separatorGap = 8;
        COLORREF activeColor = RGB(0, 0, 0);
        COLORREF inactiveColor = RGB(192, 192, 192);

        GetClientRect(hwnd, &rc);
        FillRect(hdc, &rc, g_onlineBrush ? g_onlineBrush : (HBRUSH)(COLOR_WINDOW + 1));
        FrameRect(hdc, &rc, (HBRUSH)GetStockObject(BLACK_BRUSH));

        if (g_uiFont) SelectObject(hdc, g_uiFont);
        SetBkMode(hdc, TRANSPARENT);

        BuildSortedComputerSnapshot(list, &count);
        y = paddingY;

        if (count <= 0) {
            SetTextColor(hdc, inactiveColor);
            { const char* emptyText = "No computers configured."; TextOutA(hdc, border + paddingX, border + y, emptyText, (int)strlen(emptyText)); }
        } else {
            if (!ParseIpOctets(list[0].address, prevOctets)) {
                prevOctets[2] = -1;
            }

            for (int i = 0; i < count; ++i) {
                const char* name = list[i].name[0] ? list[i].name : list[i].section;

                if (i > 0) {
                    if (!ParseIpOctets(list[i].address, octets)) {
                        octets[2] = -1;
                    }
                    if (octets[2] != prevOctets[2]) {
                        int lineY = border + y + 2;
                        HPEN pen = CreatePen(PS_SOLID, 1, RGB(0, 0, 0));
                        HPEN oldPen = (HPEN)SelectObject(hdc, pen);
                        MoveToEx(hdc, border + paddingX, lineY, NULL);
                        LineTo(hdc, rc.right - border - paddingX, lineY);
                        SelectObject(hdc, oldPen);
                        DeleteObject(pen);
                        y += separatorGap;
                    }
                    prevOctets[2] = octets[2];
                }

                SetTextColor(hdc, list[i].status == STATUS_ONLINE ? activeColor : inactiveColor);
                TextOutA(hdc, border + paddingX, border + y, name, (int)strlen(name));
                y += rowHeight;
            }
        }

        EndPaint(hwnd, &ps);
        return 0;
    }
    case WM_EXITSIZEMOVE:
        SaveOnlineWindowPlacement();
        return 0;
    case WM_CLOSE:
        SaveOnlineWindowPlacement();
        ShowWindow(hwnd, SW_HIDE);
        g_alwaysShowOnline = FALSE;
        return 0;
    case WM_DESTROY:
        SaveOnlineWindowPlacement();
        return 0;
    }
    return DefWindowProcA(hwnd, msg, wParam, lParam);
}

static LRESULT CALLBACK MainWndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == g_taskbarCreatedMsg) {
        AddTrayIcon();
        UpdateTrayTooltip();
        return 0;
    }

    switch (msg) {
    case WM_CREATE:
        return 0;

    case WM_TRAYICON: {
        UINT trayEvent = (UINT)LOWORD(lParam);
        int trayX = GET_X_LPARAM(wParam);
        int trayY = GET_Y_LPARAM(wParam);
        POINT pt;

        if (trayEvent == 0) {
            trayEvent = (UINT)lParam;
        }
        if ((trayX <= 0 && trayY <= 0) || trayX < 0 || trayY < 0) {
            GetCursorPos(&pt);
            trayX = pt.x;
            trayY = pt.y;
        }

        switch (trayEvent) {
        case WM_MOUSEMOVE:
            return 0;
        case WM_RBUTTONUP:
        case WM_CONTEXTMENU:
        case NIN_POPUPOPEN:
            ShowContextMenuAt(trayX, trayY);
            return 0;
        case WM_LBUTTONUP:
        case NIN_SELECT:
        case NIN_KEYSELECT:
        case WM_LBUTTONDBLCLK:
            ToggleOnlineWindowFromTray();
            return 0;
        }
        break;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam)) {
        case ID_MENU_SHOW_ONLINE:
            ToggleOnlineWindow();
            return 0;
        case ID_MENU_RELOAD:
            LoadConfiguration();
            ShowBalloon(APP_NAME, "Configuration reloaded.", NIIF_INFO);
            return 0;
        case ID_MENU_EXIT:
            DestroyWindow(hwnd);
            return 0;
        default:
            if (LOWORD(wParam) >= ID_WAKE_BASE && LOWORD(wParam) < ID_WAKE_BASE + MAX_COMPUTERS) {
                int index = LOWORD(wParam) - ID_WAKE_BASE;
                if (AskActionConfirmation(ACTION_WAKE, index)) {
                    StartActionAsync(ACTION_WAKE, index);
                }
                return 0;
            }
            if (LOWORD(wParam) >= ID_SHUTDOWN_BASE && LOWORD(wParam) < ID_SHUTDOWN_BASE + MAX_COMPUTERS) {
                int index = LOWORD(wParam) - ID_SHUTDOWN_BASE;
                if (AskActionConfirmation(ACTION_SHUTDOWN, index)) {
                    StartActionAsync(ACTION_SHUTDOWN, index);
                }
                return 0;
            }
            if (LOWORD(wParam) >= ID_RDP_BASE && LOWORD(wParam) < ID_RDP_BASE + MAX_COMPUTERS) {
                int index = LOWORD(wParam) - ID_RDP_BASE;
                StartActionAsync(ACTION_RDP, index);
                return 0;
            }
            if (LOWORD(wParam) >= ID_REBOOT_BASE && LOWORD(wParam) < ID_REBOOT_BASE + MAX_COMPUTERS) {
                int index = LOWORD(wParam) - ID_REBOOT_BASE;
                if (AskActionConfirmation(ACTION_REBOOT, index)) {
                    StartActionAsync(ACTION_REBOOT, index);
                }
                return 0;
            }
            break;
        }
        break;

    case WM_STATUS_CHANGED: {
        int index = (int)wParam;
        int packed = (int)lParam;
        short newStatus = (short)(packed & 0xFFFF);
        short oldStatus = (short)((packed >> 16) & 0xFFFF);
        char msgText[256];
        char name[MAX_NAME] = {0};

        UpdateOnlineWindowText();
        UpdateTrayTooltip();

        EnterCriticalSection(&g_lock);
        if (index >= 0 && index < g_computerCount) {
            SafeCopy(name, sizeof(name), g_computers[index].name[0] ? g_computers[index].name : g_computers[index].address);
        }
        LeaveCriticalSection(&g_lock);

        if (oldStatus != STATUS_UNKNOWN && oldStatus != newStatus) {
            snprintf(msgText, sizeof(msgText), "%s is now %s.", name, newStatus == STATUS_ONLINE ? "online" : "offline");
            ShowBalloon(APP_NAME, msgText, NIIF_INFO);
        }
        return 0;
    }

    case WM_DESTROY:
        InterlockedExchange(&g_stopMonitor, 1);
        if (g_monitorThread) {
            WaitForSingleObject(g_monitorThread, 3000);
            CloseHandle(g_monitorThread);
            g_monitorThread = NULL;
        }
        RemoveTrayIcon();
        PostQuitMessage(0);
        return 0;
    }
    return DefWindowProcA(hwnd, msg, wParam, lParam);
}

static BOOL RegisterWindowClasses(void) {
    WNDCLASSA wc;

    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = MainWndProc;
    wc.hInstance = g_hInst;
    wc.lpszClassName = APP_CLASS;
    if (!RegisterClassA(&wc)) return FALSE;

    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = PopupWndProc;
    wc.hInstance = g_hInst;
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = POPUP_CLASS;
    wc.hCursor = LoadCursorA(NULL, IDC_ARROW);
    if (!RegisterClassA(&wc)) return FALSE;

    ZeroMemory(&wc, sizeof(wc));
    wc.lpfnWndProc = OnlineWndProc;
    wc.hInstance = g_hInst;
    wc.hbrBackground = g_onlineBrush ? g_onlineBrush : (HBRUSH)(COLOR_WINDOW + 1);
    wc.lpszClassName = ONLINE_CLASS;
    wc.hCursor = LoadCursorA(NULL, IDC_ARROW);
    if (!RegisterClassA(&wc)) return FALSE;

    return TRUE;
}

int APIENTRY WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
    MSG msg;
    WSADATA wsa;
    (void)hPrevInstance;
    (void)lpCmdLine;
    (void)nCmdShow;

    g_hInst = hInstance;
    g_taskbarCreatedMsg = RegisterWindowMessageA("TaskbarCreated");
    InitializeCriticalSection(&g_lock);
    BuildAppPaths();
    LoadOnlineWindowPlacement();
    g_onlineBrush = CreateSolidBrush(RGB(255, 255, 210));

    if (WSAStartup(MAKEWORD(2, 2), &wsa) != 0) {
        MessageBoxA(NULL, "WSAStartup failed.", APP_NAME, MB_ICONERROR | MB_OK);
        if (g_onlineBrush) {
            DeleteObject(g_onlineBrush);
            g_onlineBrush = NULL;
        }
        DeleteCriticalSection(&g_lock);
        return 1;
    }

    g_uiFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    if (!RegisterWindowClasses()) {
        MessageBoxA(NULL, "Unable to register window classes.", APP_NAME, MB_ICONERROR | MB_OK);
        if (g_onlineBrush) {
            DeleteObject(g_onlineBrush);
            g_onlineBrush = NULL;
        }
        WSACleanup();
        DeleteCriticalSection(&g_lock);
        return 1;
    }

    g_mainWnd = CreateWindowExA(WS_EX_TOOLWINDOW, APP_CLASS, APP_NAME,
                                WS_OVERLAPPED,
                                CW_USEDEFAULT, CW_USEDEFAULT, 200, 100,
                                NULL, NULL, hInstance, NULL);
    if (!g_mainWnd) {
        MessageBoxA(NULL, "Unable to create main window.", APP_NAME, MB_ICONERROR | MB_OK);
        if (g_onlineBrush) {
            DeleteObject(g_onlineBrush);
            g_onlineBrush = NULL;
        }
        WSACleanup();
        DeleteCriticalSection(&g_lock);
        return 1;
    }

    g_popupWnd = CreateWindowExA(WS_EX_TOOLWINDOW | WS_EX_TOPMOST,
                                 POPUP_CLASS, APP_NAME,
                                 WS_POPUP | WS_BORDER,
                                 CW_USEDEFAULT, CW_USEDEFAULT, 420, 120,
                                 g_mainWnd, NULL, hInstance, NULL);

    g_onlineWnd = CreateWindowExA(WS_EX_TOOLWINDOW | WS_EX_TOPMOST | WS_EX_LAYERED,
                                  ONLINE_CLASS, "Online computers",
                                  WS_POPUP,
                                  CW_USEDEFAULT, CW_USEDEFAULT, 220, 80,
                                  NULL, NULL, hInstance, NULL);

    if (!g_popupWnd || !g_onlineWnd) {
        MessageBoxA(NULL, "Unable to create helper windows.", APP_NAME, MB_ICONERROR | MB_OK);
        DestroyWindow(g_mainWnd);
        if (g_onlineBrush) {
            DeleteObject(g_onlineBrush);
            g_onlineBrush = NULL;
        }
        WSACleanup();
        DeleteCriticalSection(&g_lock);
        return 1;
    }

    if (g_onlinePosValid) {
        SetWindowPos(g_onlineWnd, NULL, g_onlineSavedPos.x, g_onlineSavedPos.y, 0, 0,
                     SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
    }
    SetLayeredWindowAttributes(g_onlineWnd, 0, 224, LWA_ALPHA);

    ShowWindow(g_mainWnd, SW_HIDE);

    if (!LoadConfiguration()) {
        DestroyWindow(g_onlineWnd);
        DestroyWindow(g_popupWnd);
        DestroyWindow(g_mainWnd);
        if (g_onlineBrush) {
            DeleteObject(g_onlineBrush);
            g_onlineBrush = NULL;
        }
        WSACleanup();
        DeleteCriticalSection(&g_lock);
        return 1;
    }

    UpdateOnlineWindowText();
    if (g_alwaysShowOnline && g_onlinePosValid) {
        SetWindowPos(g_onlineWnd, HWND_TOPMOST, g_onlineSavedPos.x, g_onlineSavedPos.y, 0, 0,
                     SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
    } else {
        ShowWindow(g_onlineWnd, g_alwaysShowOnline ? SW_SHOWNOACTIVATE : SW_HIDE);
        if (g_alwaysShowOnline) {
            SetWindowPos(g_onlineWnd, HWND_TOPMOST, 0, 0, 0, 0,
                         SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE | SWP_SHOWWINDOW);
        }
    }

    g_monitorThread = CreateThread(NULL, 0, MonitorThreadProc, NULL, 0, NULL);
    if (!g_monitorThread) {
        MessageBoxA(NULL, "Unable to start monitor thread.", APP_NAME, MB_ICONERROR | MB_OK);
        DestroyWindow(g_onlineWnd);
        DestroyWindow(g_popupWnd);
        DestroyWindow(g_mainWnd);
        if (g_onlineBrush) {
            DeleteObject(g_onlineBrush);
            g_onlineBrush = NULL;
        }
        WSACleanup();
        DeleteCriticalSection(&g_lock);
        return 1;
    }

    if (!AddTrayIcon()) {
        MessageBoxA(NULL, "Unable to create tray icon.", APP_NAME, MB_ICONERROR | MB_OK);
        DestroyWindow(g_onlineWnd);
        DestroyWindow(g_popupWnd);
        DestroyWindow(g_mainWnd);
        if (g_onlineBrush) {
            DeleteObject(g_onlineBrush);
            g_onlineBrush = NULL;
        }
        WSACleanup();
        DeleteCriticalSection(&g_lock);
        return 1;
    }

    UpdateTrayTooltip();
    ShowBalloon(APP_NAME, "Monitoring started.", NIIF_INFO);

    while (GetMessageA(&msg, NULL, 0, 0) > 0) {
        TranslateMessage(&msg);
        DispatchMessageA(&msg);
    }

    if (g_onlineBrush) {
        DeleteObject(g_onlineBrush);
        g_onlineBrush = NULL;
    }
    WSACleanup();
    DeleteCriticalSection(&g_lock);
    return (int)msg.wParam;
}
