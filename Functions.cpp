#include "Functions.h"

#pragma comment (lib, "wmvcore.lib")

/*

	char Buffer[10];
	DWORD err = GetLastError();            //Error-Checking Code
	_itoa_s(err, Buffer, 10);
	MessageBoxA(NULL, "Failed!",Buffer , MB_OK | MB_ICONERROR);
	return;

*/

HWND hDisplay;


ICaptureGraphBuilder2* pCaptureBuilder = NULL;
IGraphBuilder* pGraphBuilder = NULL;
IMediaControl* pMediaControl = NULL;
ICreateDevEnum* pDevEnum = NULL;
IEnumMoniker* pEnumMoniker = NULL;
IMoniker* pMoniker = NULL;
IBaseFilter* pAudioCaptureFilter = NULL, *pAsfWriter = NULL, *pSinkFilter = NULL;

IWMProfileManager* pProfileManager = NULL;
IWMProfile* pWMProfile = NULL;
IConfigAsfWriter* pAsfConfig = NULL;

IFileSinkFilter* pFileSink = NULL;
IServiceProvider* pServiceProvider = NULL;
IWMWriterAdvanced2* pWriterAdvanced = NULL;
IWMWriterNetworkSink* pNetworkSink = NULL;


void Initialize(HWND hWnd)
{
//-------------------------------------------------Initialize UI---------------------------------------------------------------//
	CreateWindowExA(0, "Button", "Start", WS_CHILD | WS_VISIBLE, 300, 200, 120, 50, hWnd, (HMENU)BTN_START, NULL, NULL);
	CreateWindowExA(0, "Button", "Stop", WS_CHILD | WS_VISIBLE, 600, 200, 120, 50, hWnd, (HMENU)BTN_STOP, NULL, NULL);

	hDisplay = CreateWindowExA(0, "Edit", NULL, WS_VISIBLE | WS_CHILD | WS_BORDER, 150, 20, 690, 25, hWnd, NULL, NULL, NULL);


//------------------------------------------Initialize COM and Other Interfaces------------------------------------------------//
	HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE);
	if (!SUCCEEDED(hr))
	{
		MessageBoxW(NULL, L"Failed to Initialize COM!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = CoCreateInstance(CLSID_CaptureGraphBuilder2, NULL, CLSCTX_INPROC_SERVER, IID_ICaptureGraphBuilder2, 
		(LPVOID*)&pCaptureBuilder);
	if (!SUCCEEDED(hr))
	{
		MessageBoxW(NULL, L"Failed to Instantiate Capture Graph Builder!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = CoCreateInstance(CLSID_FilterGraph, NULL, CLSCTX_INPROC_SERVER, IID_IGraphBuilder, (LPVOID*)&pGraphBuilder);
	if (!SUCCEEDED(hr))
	{
		MessageBoxW(NULL, L"Failed to Instantiate Filter Graph Builder!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pGraphBuilder->QueryInterface(IID_IMediaControl, (LPVOID*)&pMediaControl);
	if (!SUCCEEDED(hr))
	{
		MessageBoxW(NULL, L"Failed to Instantiate Media Controls!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	pCaptureBuilder -> SetFiltergraph(pGraphBuilder);

//--------------------------------------Enumerate and Select Audio Capture Device----------------------------------------------//
	hr = CoCreateInstance(CLSID_SystemDeviceEnum, NULL, CLSCTX_INPROC_SERVER, IID_ICreateDevEnum,
		(LPVOID*)&pDevEnum);
	if (!SUCCEEDED(hr))
	{
		MessageBoxW(NULL, L"Failed to Instantiate Device Enumerator!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pDevEnum->CreateClassEnumerator(CLSID_AudioInputDeviceCategory, &pEnumMoniker, NULL);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Enumerate Audio Devices!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pEnumMoniker->Next(1, &pMoniker, NULL);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Retrive Audio Device Moniker!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pMoniker->BindToObject(0, 0, IID_IBaseFilter, (LPVOID*)&pAudioCaptureFilter);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Create Audio Capture Filter!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pGraphBuilder->AddFilter(pAudioCaptureFilter, L"Audio Capture Filter");
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Add Filter!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}



//--------------------------------------------------Create Network Sink--------------------------------------------------------//
	
	hr = CoCreateInstance(CLSID_WMAsfWriter, NULL, CLSCTX_INPROC_SERVER, IID_IBaseFilter,
		(LPVOID*)&pAsfWriter);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Create ASF Writer!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pGraphBuilder->AddFilter(pAsfWriter, L"Network Sink");
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Add ASF Writer to Graph!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}



	hr = WMCreateProfileManager(&pProfileManager);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Instantiate Profile Manager!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pProfileManager->LoadProfileByID(WMProfile_V80_128StereoAudio, &pWMProfile);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Load Profile!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pAsfWriter->QueryInterface(IID_IConfigAsfWriter, (LPVOID*)&pAsfConfig);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Query IConfigAsfWriter!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pAsfConfig->ConfigureFilterUsingProfile(pWMProfile);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Configure Filter with Profile!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}


	
	hr = pAsfWriter->QueryInterface(IID_IFileSinkFilter, (LPVOID*)&pFileSink);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Query File Sink Filter!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	hr = pFileSink->SetFileName(L"C:\\Users\\opiyo\\Videos\\dummy.wmv", NULL);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Set File Name!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}


	
	hr = pAsfWriter->QueryInterface(IID_IServiceProvider, (LPVOID*)&pServiceProvider);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Query Service Provider!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	
	//hr = pServiceProvider->QueryService(IID_IWMWriterAdvanced2, &pWriterAdvanced);
	hr = pServiceProvider->QueryService(IID_IWMWriterAdvanced2, IID_IWMWriterAdvanced2,
		(LPVOID*)&pWriterAdvanced);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Query Writer Advanced Service!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}
	

	pWriterAdvanced->SetLiveSource(true);
	pWriterAdvanced->RemoveSink(NULL);

	hr = WMCreateWriterNetworkSink(&pNetworkSink);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Create Network Sink!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	DWORD dwPort = 80; //Actual Port Number Should be for a WinSock Socket. Use HTTP/HTTPS Port
	WCHAR* pHostAddr = (WCHAR*) malloc(100);
	DWORD dwLength = 100;
	pNetworkSink->Open(&dwPort);
	pNetworkSink->GetHostURL(pHostAddr, &dwLength);

	SetWindowTextW(hDisplay, pHostAddr);

	hr = pWriterAdvanced->AddSink(pNetworkSink);
	if (hr != S_OK)
	{
		MessageBoxW(NULL, L"Failed to Add Network Sink!", L"Error", MB_OK | MB_ICONERROR);
		PostQuitMessage(0);
	}

	//MessageBoxW(NULL, L"Initialization Successfull!", L"Confirmation", MB_OK | MB_ICONINFORMATION);
	

}


void StartCapture()
{	
	
	HRESULT	hr = pCaptureBuilder->RenderStream(&PIN_CATEGORY_CAPTURE, &MEDIATYPE_Audio, pAudioCaptureFilter,
		NULL, pAsfWriter);
	if (FAILED(hr))
	{
		MessageBoxW(NULL, L"Failed to Start Audio Capture!", L"Error", MB_OK | MB_ICONERROR);
		return;
	}



	MessageBoxW(NULL, L"Capture Started!", L"Confirmation", MB_OK | MB_ICONINFORMATION);

	pMediaControl->Run();
	
}



void StopCapture()
{
	pMediaControl->Stop();
	pNetworkSink->Close();
}
