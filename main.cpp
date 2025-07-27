#include <windows.h>
#include <mmdeviceapi.h>
#include <audioclient.h>
#include <endpointvolume.h>
#include <functiondiscoverykeys_devpkey.h>
#ifndef PKEY_Device_FriendlyName
#define INITGUID
#include <propkey.h>
DEFINE_PROPERTYKEY(PKEY_Device_FriendlyName, 0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0, 14);
#endif
#include <commctrl.h>
#include <vector>
#include <string>
#include <thread>
#include <atomic>
#include <mutex>
#include <map>
#include <comdef.h>

#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "uuid.lib")
#pragma comment(lib, "winmm.lib")
#pragma comment(lib, "avrt.lib")
#pragma comment(lib, "comctl32.lib")

#define MAX_DEVICES 16
#define WM_UPDATE_VOLUME (WM_USER + 1)

struct DeviceInfo {
    std::wstring name;
    std::wstring id;
    bool selected;
    float volume;
    HWND checkbox;
    HWND slider;
};

std::vector<DeviceInfo> g_devices;
std::mutex g_mutex;
std::atomic<bool> g_running{ false };
HWND g_hWnd = nullptr;
HWND g_btnStart = nullptr;
HWND g_btnStop = nullptr;
HWND g_lblStatus = nullptr;
std::thread g_audioThread;
std::wstring g_defaultDeviceId;

void SetDefaultDeviceMute(bool mute) {
    CoInitialize(nullptr);
    IMMDeviceEnumerator* pEnum = nullptr;
    IMMDevice* pDevice = nullptr;
    IAudioEndpointVolume* pVol = nullptr;
    if (SUCCEEDED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pEnum)))) {
        if (SUCCEEDED(pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDevice))) {
            if (SUCCEEDED(pDevice->Activate(__uuidof(IAudioEndpointVolume), CLSCTX_ALL, nullptr, (void**)&pVol))) {
                pVol->SetMute(mute, nullptr);
                pVol->Release();
            }
            pDevice->Release();
        }
        pEnum->Release();
    }
    CoUninitialize();
}

void EnumAudioDevices() {
    g_devices.clear();
    CoInitialize(nullptr);
    IMMDeviceEnumerator* pEnum = nullptr;
    IMMDeviceCollection* pColl = nullptr;
    if (SUCCEEDED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pEnum)))) {
        if (SUCCEEDED(pEnum->EnumAudioEndpoints(eRender, DEVICE_STATE_ACTIVE, &pColl))) {
            UINT count = 0;
            pColl->GetCount(&count);
            for (UINT i = 0; i < count && g_devices.size() < MAX_DEVICES; ++i) {
                IMMDevice* pDev = nullptr;
                if (SUCCEEDED(pColl->Item(i, &pDev))) {
                    LPWSTR id = nullptr;
                    pDev->GetId(&id);
                    IPropertyStore* pProps = nullptr;
                    std::wstring name = L"Unknown";
                    if (SUCCEEDED(pDev->OpenPropertyStore(STGM_READ, &pProps))) {
                        PROPVARIANT varName;
                        PropVariantInit(&varName);
                        if (SUCCEEDED(pProps->GetValue(PKEY_Device_FriendlyName, &varName))) {
                            name = varName.pwszVal;
                            PropVariantClear(&varName);
                        }
                        pProps->Release();
                    }
                    g_devices.push_back({ name, id, false, 1.0f, nullptr, nullptr });
                    CoTaskMemFree(id);
                    pDev->Release();
                }
            }
            // 获取默认设备ID
            IMMDevice* pDefault = nullptr;
            if (SUCCEEDED(pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pDefault))) {
                LPWSTR id = nullptr;
                pDefault->GetId(&id);
                g_defaultDeviceId = id;
                CoTaskMemFree(id);
                pDefault->Release();
            }
            pColl->Release();
        }
        pEnum->Release();
    }
    CoUninitialize();
}

void UpdateDeviceVolumes() {
    std::lock_guard<std::mutex> lock(g_mutex);
    for (auto& dev : g_devices) {
        if (dev.slider) {
            LRESULT pos = SendMessage(dev.slider, TBM_GETPOS, 0, 0);
            dev.volume = pos / 100.0f;
        }
    }
}

void AudioLoop() {
    CoInitialize(nullptr);
    IMMDeviceEnumerator* pEnum = nullptr;
    IMMDevice* pLoopback = nullptr;
    IAudioClient* pClient = nullptr;
    IAudioCaptureClient* pCapture = nullptr;
    WAVEFORMATEX* pwfx = nullptr;
    HANDLE hEvent = CreateEvent(nullptr, FALSE, FALSE, nullptr);
    std::vector<IMMDevice*> outDevices;
    std::vector<IAudioClient*> outClients;
    std::vector<IAudioRenderClient*> outRenderers;

    bool success = false;
    do {
        if (FAILED(CoCreateInstance(__uuidof(MMDeviceEnumerator), nullptr, CLSCTX_ALL, IID_PPV_ARGS(&pEnum)))) break;
        if (FAILED(pEnum->GetDefaultAudioEndpoint(eRender, eConsole, &pLoopback))) break;
        if (FAILED(pLoopback->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void**)&pClient))) break;
        if (FAILED(pClient->GetMixFormat(&pwfx))) break;
        if (FAILED(pClient->Initialize(AUDCLNT_SHAREMODE_SHARED, AUDCLNT_STREAMFLAGS_LOOPBACK | AUDCLNT_STREAMFLAGS_EVENTCALLBACK, 1000000, 0, pwfx, nullptr))) break;
        if (FAILED(pClient->SetEventHandle(hEvent))) break;
        if (FAILED(pClient->GetService(IID_PPV_ARGS(&pCapture)))) break;
        if (FAILED(pClient->Start())) break;

        // 2. 为每个选中的设备创建 AudioClient
        {
            std::lock_guard<std::mutex> lock(g_mutex);
            for (auto& dev : g_devices) {
                if (dev.selected) {
                    IMMDevice* pDev = nullptr;
                    IAudioClient* pOutClient = nullptr;
                    IAudioRenderClient* pRender = nullptr;
                    if (SUCCEEDED(pEnum->GetDevice(dev.id.c_str(), &pDev)) &&
                        SUCCEEDED(pDev->Activate(__uuidof(IAudioClient), CLSCTX_ALL, nullptr, (void**)&pOutClient)) &&
                        SUCCEEDED(pOutClient->Initialize(AUDCLNT_SHAREMODE_SHARED, 0, 1000000, 0, pwfx, nullptr)) &&
                        SUCCEEDED(pOutClient->GetService(IID_PPV_ARGS(&pRender)))) {
                        outDevices.push_back(pDev);
                        outClients.push_back(pOutClient);
                        outRenderers.push_back(pRender);
                        pOutClient->Start();
                    }
                    else {
                        if (pRender) pRender->Release();
                        if (pOutClient) pOutClient->Release();
                        if (pDev) pDev->Release();
                    }
                }
            }
        }

        // 3. 静音默认设备
        SetDefaultDeviceMute(true);

        // 4. 循环采集并输出
        while (g_running) {
            DWORD wait = WaitForSingleObject(hEvent, 100);
            if (wait != WAIT_OBJECT_0) continue;
            UINT32 packetLength = 0;
            if (FAILED(pCapture->GetNextPacketSize(&packetLength))) break;
            while (packetLength) {
                BYTE* pData = nullptr;
                UINT32 numFrames = 0;
                DWORD flags = 0;
                if (FAILED(pCapture->GetBuffer(&pData, &numFrames, &flags, nullptr, nullptr))) break;
                {
                    std::lock_guard<std::mutex> lock(g_mutex);
                    for (size_t i = 0; i < outRenderers.size(); ++i) {
                        UINT32 pad = 0, avail = 0;
                        if (SUCCEEDED(outClients[i]->GetCurrentPadding(&pad))) {
                            avail = pwfx->nSamplesPerSec * pwfx->nBlockAlign / 100; // 10ms
                            avail = avail > pad ? avail - pad : 0;
                            if (avail > numFrames) avail = numFrames;
                            BYTE* pOut = nullptr;
                            if (avail > 0 && SUCCEEDED(outRenderers[i]->GetBuffer(avail, &pOut))) {
                                // 音量调整
                                float vol = g_devices[i].volume;
                                for (UINT32 j = 0; j < avail * pwfx->nBlockAlign; ++j) {
                                    int sample = ((int)pData[j] - 128) * vol + 128;
                                    if (sample < 0) sample = 0;
                                    if (sample > 255) sample = 255;
                                    pOut[j] = (BYTE)sample;
                                }
                                outRenderers[i]->ReleaseBuffer(avail, 0);
                            }
                        }
                    }
                }
                pCapture->ReleaseBuffer(numFrames);
                if (FAILED(pCapture->GetNextPacketSize(&packetLength))) break;
            }
        }

        // 5. 恢复静音
        SetDefaultDeviceMute(false);

        success = true;
    } while (0);

    if (pCapture) pCapture->Release();
    if (pClient) pClient->Release();
    if (pLoopback) pLoopback->Release();
    if (pEnum) pEnum->Release();
    if (pwfx) CoTaskMemFree(pwfx);
    for (auto p : outRenderers) p->Release();
    for (auto p : outClients) p->Release();
    for (auto p : outDevices) p->Release();
    if (hEvent) CloseHandle(hEvent);
    CoUninitialize();
}

void StartAudio() {
    g_running = true;
    g_audioThread = std::thread(AudioLoop);
}

void StopAudio() {
    g_running = false;
    if (g_audioThread.joinable()) g_audioThread.join();
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    switch (msg) {
    case WM_CREATE: {
        EnumAudioDevices();
        int y = 10;
        for (size_t i = 0; i < g_devices.size(); ++i) {
            g_devices[i].checkbox = CreateWindowW(L"BUTTON", g_devices[i].name.c_str(), WS_VISIBLE | WS_CHILD | BS_AUTOCHECKBOX,
                10, y, 300, 24, hWnd, (HMENU)(1000 + i), nullptr, nullptr);
            g_devices[i].slider = CreateWindowW(TRACKBAR_CLASSW, L"", WS_VISIBLE | WS_CHILD | TBS_AUTOTICKS,
                320, y, 200, 24, hWnd, (HMENU)(2000 + i), nullptr, nullptr);
            SendMessage(g_devices[i].slider, TBM_SETRANGE, TRUE, MAKELPARAM(0, 100));
            SendMessage(g_devices[i].slider, TBM_SETPOS, TRUE, 100);
            y += 34;
        }
        g_btnStart = CreateWindowW(L"BUTTON", L"开始", WS_VISIBLE | WS_CHILD, 10, y, 80, 30, hWnd, (HMENU)3000, nullptr, nullptr);
        g_btnStop = CreateWindowW(L"BUTTON", L"停止", WS_VISIBLE | WS_CHILD, 100, y, 80, 30, hWnd, (HMENU)3001, nullptr, nullptr);
        g_lblStatus = CreateWindowW(L"STATIC", L"状态：未启动", WS_VISIBLE | WS_CHILD, 200, y, 200, 30, hWnd, (HMENU)3002, nullptr, nullptr);
        break;
    }
    case WM_COMMAND: {
        int id = LOWORD(wParam);
        if (id >= 1000 && id < 1000 + (int)g_devices.size()) {
            int idx = id - 1000;
            g_devices[idx].selected = (SendMessage(g_devices[idx].checkbox, BM_GETCHECK, 0, 0) == BST_CHECKED);
        }
        else if (id == 3000) { // Start
            UpdateDeviceVolumes();
            SetWindowTextW(g_lblStatus, L"状态：运行中...");
            StartAudio();
        }
        else if (id == 3001) { // Stop
            SetWindowTextW(g_lblStatus, L"状态：已停止");
            StopAudio();
        }
        break;
    }
    case WM_HSCROLL: {
        UpdateDeviceVolumes();
        break;
    }
    case WM_DESTROY:
        StopAudio();
        PostQuitMessage(0);
        break;
    }
    return DefWindowProc(hWnd, msg, wParam, lParam);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE, PWSTR, int nCmdShow) {
    // 最小化自身的控制台窗口（如果有）
    HWND hConsole = GetConsoleWindow();
    if (hConsole) ShowWindow(hConsole, SW_MINIMIZE);
    INITCOMMONCONTROLSEX icc = { sizeof(icc), ICC_BAR_CLASSES };
    InitCommonControlsEx(&icc);

    WNDCLASSW wc = { 0 };
    wc.lpfnWndProc = WndProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = L"TwoSpeakerApp";
    RegisterClassW(&wc);

    HWND hWnd = CreateWindowW(L"TwoSpeakerApp", L"系统声音多设备输出", WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 600, 400, nullptr, nullptr, hInstance, nullptr);
    g_hWnd = hWnd;
    ShowWindow(hWnd, nCmdShow);

    MSG msg;
    while (GetMessage(&msg, nullptr, 0, 0)) {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }
    return 0;
}