/***********************************
 * Copyright (c) 2024 Roger Brown.
 * Licensed under the MIT License.
 ****/

#include<windows.h>
#include <dispsvc_h.h>

static OLECHAR globalModuleFileName[260];
static LONG serverLockCount;
static OLECHAR ServiceName[] = L"AtYourService";
static SERVICE_STATUS gSvcStatus;
static SERVICE_STATUS_HANDLE gSvcStatusHandle;
static HANDLE ghSvcStopEvent = NULL;

typedef struct CHelloWorldData
{
	IUnknown IUnknown;
	IHelloWorld IHelloWorld;
	LONG lUsage;
	IUnknown* lpOuter;
	ITypeInfo* piTypeInfo;
} CHelloWorldData;

typedef struct CClassObjectData
{
	IClassFactory IClassFactory;
	LONG lUsage;
} CClassObjectData;

#define GetBaseObjectPtr(x,y,z)     (x *)(((char *)(void *)z)-(size_t)(char *)(&(((x*)NULL)->y)))

static STDMETHODIMP CHelloWorld_IUnknown_QueryInterface(IUnknown* pThis, REFIID riid, void** ppvObject)
{
	HRESULT hr = E_NOINTERFACE;
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IUnknown, pThis);

	if (IsEqualIID(riid, &IID_IDispatch) || IsEqualIID(riid, &IID_IHelloWorld))
	{
		IUnknown_AddRef(pData->lpOuter);

		*ppvObject = &(pData->IHelloWorld);

		hr = S_OK;
	}
	else
	{
		if (IsEqualIID(riid, &IID_IUnknown))
		{
			InterlockedIncrement(&pData->lUsage);

			*ppvObject = &(pData->IUnknown);

			hr = S_OK;
		}
		else
		{
			*ppvObject = NULL;
		}
	}

	return hr;
}

static STDMETHODIMP_(ULONG) CHelloWorld_IUnknown_AddRef(IUnknown* pThis)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IUnknown, pThis);
	return InterlockedIncrement(&pData->lUsage);
}

static STDMETHODIMP_(ULONG) CHelloWorld_IUnknown_Release(IUnknown* pThis)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IUnknown, pThis);
	LONG res = InterlockedDecrement(&pData->lUsage);

	if (!res)
	{
		if (pData->piTypeInfo) IUnknown_Release(pData->piTypeInfo);
		LocalFree(pData);
	}

	return res;
}

static STDMETHODIMP CHelloWorld_IHelloWorld_QueryInterface(IHelloWorld* pThis, REFIID riid, void** ppvObject)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);
	return IUnknown_QueryInterface(pData->lpOuter, riid, ppvObject);
}

static STDMETHODIMP_(ULONG) CHelloWorld_IHelloWorld_AddRef(IHelloWorld* pThis)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);
	return IUnknown_AddRef(pData->lpOuter);
}

static STDMETHODIMP_(ULONG) CHelloWorld_IHelloWorld_Release(IHelloWorld* pThis)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);
	return IUnknown_Release(pData->lpOuter);
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetTypeInfoCount(IHelloWorld* pThis, UINT* pctinfo)
{
	*pctinfo = 1;
	return S_OK;
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetTypeInfo(IHelloWorld* pThis, UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);
	if (iTInfo) return DISP_E_BADINDEX;
	ITypeInfo_AddRef(pData->piTypeInfo);
	*ppTInfo = pData->piTypeInfo;

	return S_OK;
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetIDsOfNames(IHelloWorld* pThis, REFIID riid, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);

	if (!IsEqualIID(riid, &IID_NULL))
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	return DispGetIDsOfNames(pData->piTypeInfo, rgszNames, cNames, rgDispId);
}

static STDMETHODIMP CHelloWorld_IHelloWorld_Invoke(IHelloWorld* pThis, DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pDispParams, VARIANT* pVarResult, EXCEPINFO* pExcepInfo, UINT* puArgErr)
{
	CHelloWorldData* pData = GetBaseObjectPtr(CHelloWorldData, IHelloWorld, pThis);

	if (!IsEqualIID(riid, &IID_NULL))
	{
		return DISP_E_UNKNOWNINTERFACE;
	}

	return DispInvoke(pThis, pData->piTypeInfo, dispIdMember, wFlags, pDispParams, pVarResult, pExcepInfo, puArgErr);
}

static STDMETHODIMP CHelloWorld_IHelloWorld_GetMessage(IHelloWorld* pThis, int Hint, BSTR* lpMessage)
{
	LPCOLESTR message = NULL;
	OLECHAR buffer[256];
	DWORD dw = (sizeof(buffer) / sizeof(buffer[0])) - 1;
	HRESULT hr = CoImpersonateClient();

	if (SUCCEEDED(hr))
	{
		switch (Hint)
		{
		case -1:
			message = L"Goodbye, Cruel World";
			break;

		case 1:
			message = L"Hello World";
			break;

		case 2:
			if (GetUserNameW(buffer, &dw))
			{
				message = buffer;
			}
			break;

		case 3:
			if (GetComputerNameW(buffer, &dw))
			{
				message = buffer;
			}
			break;

		case 4:
			message = globalModuleFileName;
			break;

		case 5:
			message = ServiceName;
			break;
		}

		*lpMessage = message ? SysAllocString(message) : NULL;

		hr = message ? S_OK : S_FALSE;
	}

	return hr;
}

static IUnknownVtbl CHelloWorld_IUnknownVtbl =
{
	CHelloWorld_IUnknown_QueryInterface,
	CHelloWorld_IUnknown_AddRef,
	CHelloWorld_IUnknown_Release
};

static IHelloWorldVtbl CHelloWorld_IHelloWorldVtbl =
{
	CHelloWorld_IHelloWorld_QueryInterface,
	CHelloWorld_IHelloWorld_AddRef,
	CHelloWorld_IHelloWorld_Release,
	CHelloWorld_IHelloWorld_GetTypeInfoCount,
	CHelloWorld_IHelloWorld_GetTypeInfo,
	CHelloWorld_IHelloWorld_GetIDsOfNames,
	CHelloWorld_IHelloWorld_Invoke,
	CHelloWorld_IHelloWorld_GetMessage
};

static STDMETHODIMP CClassObject_CHelloWorld_IClassFactory_QueryInterface(IClassFactory* pThis, REFIID riid, void** ppvObject)
{
	CClassObjectData* pData = GetBaseObjectPtr(CClassObjectData, IClassFactory, pThis);

	*ppvObject = NULL;

	if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IClassFactory))
	{
		InterlockedIncrement(&pData->lUsage);

		*ppvObject = pThis;

		return S_OK;
	}

	return E_NOINTERFACE;
}

static STDMETHODIMP_(ULONG) CClassObject_CHelloWorld_IClassFactory_AddRef(IClassFactory* pThis)
{
	CClassObjectData* pData = GetBaseObjectPtr(CClassObjectData, IClassFactory, pThis);
	return InterlockedIncrement(&pData->lUsage);
}

static STDMETHODIMP_(ULONG) CClassObject_CHelloWorld_IClassFactory_Release(IClassFactory* pThis)
{
	CClassObjectData* pData = GetBaseObjectPtr(CClassObjectData, IClassFactory, pThis);
	return InterlockedDecrement(&pData->lUsage);
}

static STDMETHODIMP CClassObject_CHelloWorld_IClassFactory_CreateInstance(IClassFactory* pThis, LPVOID punk, REFIID riid, void** ppvObject)
{
	HRESULT hr = E_FAIL;

	if (punk == NULL || IsEqualIID(riid, &IID_IUnknown))
	{
		ITypeLib* piTypeLib = NULL;

		hr = LoadTypeLibEx(globalModuleFileName, REGKIND_NONE, &piTypeLib);

		if (SUCCEEDED(hr))
		{
			CHelloWorldData* pData = LocalAlloc(LMEM_ZEROINIT, sizeof(*pData));

			if (pData)
			{
				IUnknown* p = &(pData->IUnknown);

				pData->IUnknown.lpVtbl = &CHelloWorld_IUnknownVtbl;
				pData->IHelloWorld.lpVtbl = &CHelloWorld_IHelloWorldVtbl;

				pData->lUsage = 1;
				pData->lpOuter = punk ? punk : p;

				hr = ITypeLib_GetTypeInfoOfGuid(piTypeLib, &IID_IHelloWorld, &pData->piTypeInfo);

				if (SUCCEEDED(hr))
				{
					if (punk)
					{
						hr = S_OK;

						*ppvObject = p;
					}
					else
					{
						hr = IUnknown_QueryInterface(p, riid, ppvObject);

						IUnknown_Release(p);
					}
				}
				else
				{
					IUnknown_Release(p);
				}
			}

			if (piTypeLib)
			{
				ITypeLib_Release(piTypeLib);
			}
		}
	}

	return hr;
}

static STDMETHODIMP CClassObject_CHelloWorld_IClassFactory_LockServer(IClassFactory* pThis, BOOL bLock)
{
	if (bLock)
	{
		InterlockedIncrement(&serverLockCount);
	}
	else
	{
		InterlockedDecrement(&serverLockCount);
	}

	return S_OK;
}

static IClassFactoryVtbl CClassObject_CHelloWorld_IClassFactoryVtbl =
{
	CClassObject_CHelloWorld_IClassFactory_QueryInterface,
	CClassObject_CHelloWorld_IClassFactory_AddRef,
	CClassObject_CHelloWorld_IClassFactory_Release,
	CClassObject_CHelloWorld_IClassFactory_CreateInstance,
	CClassObject_CHelloWorld_IClassFactory_LockServer
};

static CClassObjectData CClassObject_CHelloWorld = { {&CClassObject_CHelloWorld_IClassFactoryVtbl }, 0 };

static VOID ReportSvcStatus(DWORD dwCurrentState, DWORD dwWin32ExitCode, DWORD dwWaitHint)
{
	static DWORD dwCheckPoint = 1;

	gSvcStatus.dwCurrentState = dwCurrentState;
	gSvcStatus.dwWin32ExitCode = dwWin32ExitCode;
	gSvcStatus.dwWaitHint = dwWaitHint;

	if (dwCurrentState == SERVICE_START_PENDING)
	{
		gSvcStatus.dwControlsAccepted = 0;
	}
	else
	{
		gSvcStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP;
	}

	if ((dwCurrentState == SERVICE_RUNNING) || (dwCurrentState == SERVICE_STOPPED))
	{
		gSvcStatus.dwCheckPoint = 0;
	}
	else
	{
		gSvcStatus.dwCheckPoint = dwCheckPoint++;
	}

	SetServiceStatus(gSvcStatusHandle, &gSvcStatus);
}

static void WINAPI ServiceCtrlHandler(DWORD dwCtrl)
{
	HANDLE hEvent = ghSvcStopEvent;

	switch (dwCtrl)
	{
	case SERVICE_CONTROL_STOP:
		if (hEvent)
		{
			ReportSvcStatus(SERVICE_STOP_PENDING, NO_ERROR, 0);

			SetEvent(hEvent);
		}
		break;

	case SERVICE_CONTROL_INTERROGATE:
		break;

	default:
		break;
	}
}

static void WINAPI ServiceMain(DWORD argc, LPWSTR* argv)
{
	ghSvcStopEvent = CreateEvent(NULL, TRUE, FALSE, NULL);

	gSvcStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
	gSvcStatus.dwServiceSpecificExitCode = 0;

	gSvcStatusHandle = RegisterServiceCtrlHandlerW(ServiceName, ServiceCtrlHandler);

	__try
	{
		HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED | COINIT_DISABLE_OLE1DDE);

		if (SUCCEEDED(hr))
		{
			__try
			{
				hr = CoInitializeSecurity(NULL,
					-1,
					NULL,
					NULL,
					RPC_C_AUTHN_LEVEL_PKT,
					RPC_C_IMP_LEVEL_IMPERSONATE,
					NULL,
					EOAC_NONE,
					NULL);

				if (SUCCEEDED(hr))
				{
					DWORD dwRegister;

					hr = CoRegisterClassObject(&CLSID_CHelloWorld, (IUnknown*)&(CClassObject_CHelloWorld.IClassFactory), CLSCTX_LOCAL_SERVER, REGCLS_MULTIPLEUSE, &dwRegister);

					if (SUCCEEDED(hr))
					{
						__try
						{
							ReportSvcStatus(SERVICE_RUNNING, NO_ERROR, 0);
							WaitForSingleObject(ghSvcStopEvent, INFINITE);
						}
						__finally
						{
							hr = CoRevokeClassObject(dwRegister);
						}
					}
				}
			}
			__finally
			{
				CoUninitialize();
			}
		}
	}
	__finally
	{
		HANDLE hEvent = ghSvcStopEvent;
		ghSvcStopEvent = NULL;
		ReportSvcStatus(SERVICE_STOPPED, NO_ERROR, 0);
		CloseHandle(hEvent);
	}
}

static SERVICE_TABLE_ENTRYW DispatchTable[] =
{
	{ ServiceName, ServiceMain },
	{ NULL, NULL }
};

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow)
{
	if (!GetModuleFileNameW(hInstance, globalModuleFileName, sizeof(globalModuleFileName) / sizeof(globalModuleFileName[0])))
	{
		return 10;
	}

	return StartServiceCtrlDispatcherW(DispatchTable) ? 0 : GetLastError();
}
