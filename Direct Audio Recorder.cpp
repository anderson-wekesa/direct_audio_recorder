#define no_init_all deprecated

#ifndef UNICODE
#define UNICODE
#endif // UNICODE

#include "framework.h"
#include "Direct Audio Recorder.h"
#include "Functions.h"

#pragma comment (lib, "strmiids.lib")

LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

HWND hMainWindow;


int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{

    const wchar_t CLASS_NAME[] = L"DShow Class";

    WNDCLASS wc = {};
    wc.hInstance = hInstance;
    wc.lpfnWndProc = WindowProc;
    wc.lpszClassName = CLASS_NAME;

    if (!RegisterClass(&wc))
    {
        MessageBoxW(NULL, L"Failed to Register Class!", L"Error", MB_OK | MB_ICONERROR);
        return 0;
    }


    hMainWindow = CreateWindowEx(
        0,
        CLASS_NAME,
        L"Direct Video Recorder",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
        NULL,
        NULL,
        hInstance,
        NULL
    );

    if (!hMainWindow)
    {
        MessageBoxW(NULL, L"Failed to Create Window!", L"Error", MB_OK | MB_ICONERROR);
        return 0;
    }

    ShowWindow(hMainWindow, nCmdShow);

    MSG msg = {};

    while (GetMessage(&msg, hMainWindow, NULL, NULL))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }


    return 0;
}


LRESULT CALLBACK WindowProc(HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{

    switch (uMsg)
    {
    case WM_PAINT:
    {
        PAINTSTRUCT ps = {};
        HDC hdc = BeginPaint(hWnd, &ps);

        FillRect(hdc, &ps.rcPaint, (HBRUSH)(COLOR_WINDOW));

        EndPaint(hWnd, &ps);
    }
    break;

    case WM_COMMAND:
    {
        switch (wParam)
        {
            case BTN_START:
            {
                StartCapture();
            }
            break;

            case BTN_STOP:
            {
                StopCapture();
            }
            break;
        }
    }
    break;

    case WM_CREATE:
    {
        Initialize(hWnd);
    }
    break;

    case WM_DESTROY:
    {
        PostQuitMessage(0);
    }
    break;

    }


    return DefWindowProc(hWnd, uMsg, wParam, lParam);
}
