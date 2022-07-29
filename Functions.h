#pragma once

#define BTN_START 1
#define BTN_STOP 2

#include <Windows.h>
#include <dshow.h>
#include <dshowasf.h>
#include <wmsysprf.h>
#include <stdio.h>

void Initialize(HWND hWnd);
void StartCapture();
void StopCapture();
