#include <Wia.h>
#include <Wia_xp.h>
#include <windows.h>
#pragma comment (lib, "WiaGuid.lib")


int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
    HRESULT hResult;
    IWiaItem *pItemRoot = NULL;//使应用程序能够查询设备的功能。 IWiaItem还提供对数据传输接口和项属性的访问。 此外，此接口还提供了用于使应用程序能够控制设备的方法。
    IWiaDevMgr *pWiaDevMgr = NULL;//WIA Device Manager
    IEnumWIA_DEV_INFO* pIEnum = nullptr;
    IWiaPropertyStorage* pWiaPropertyStorage = nullptr;
    CoInitialize(NULL);
    __try
    {
        // Create WIA Device Manager instance
        pWiaDevMgr = NULL;
        hResult = CoCreateInstance(CLSID_WiaDevMgr, NULL, CLSCTX_LOCAL_SERVER, IID_IWiaDevMgr, (void**)&pWiaDevMgr);
        if (hResult != S_OK)
        {
            MessageBox(NULL, "Error: CoCreateInstance().", NULL, MB_ICONSTOP);
            __leave;
        }

        // Display a WIA select device dialog
        //hResult = pWiaDevMgr->SelectDeviceDlg(NULL, 0, WIA_SELECT_DEVICE_NODEFAULT, NULL, &pItemRoot);

        
        hResult = pWiaDevMgr->EnumDeviceInfo(WIA_DEVINFO_ENUM_ALL, &pIEnum);
        

        // User canceled
        if (hResult == S_FALSE)
        {
            MessageBox(NULL, "User canceled.", NULL, MB_ICONINFORMATION);
            __leave;
        }
        // No device available
        else if (hResult == WIA_S_NO_DEVICE_AVAILABLE)
        {
            MessageBox(NULL, "No device available.", NULL, MB_ICONINFORMATION);
            __leave;
        }
        // OK, Then ..........
        else if (hResult == S_OK)
        {
            ULONG count = 0;
            pIEnum->GetCount(&count);
            while (hResult == S_OK)
            {
                
                hResult = pIEnum->Next(1, &pWiaPropertyStorage, nullptr);
                if (hResult != S_OK) break;
                PROPSPEC PropSpec[3] = { 0 };
                PROPVARIANT PropVar[3] = { 0 };
                // How many properties are you querying for?
                const ULONG c_nPropertyCount = sizeof(PropSpec) / sizeof(PropSpec[0]);
                // Define which properties you want to read:
                // Device ID.  This is what you would use to create the device.
                PropSpec[0].ulKind = PRSPEC_PROPID;
                PropSpec[0].propid = WIA_DIP_DEV_ID;
                // Device Name
                PropSpec[1].ulKind = PRSPEC_PROPID;
                PropSpec[1].propid = WIA_DIP_DEV_NAME;
                // Device description
                PropSpec[2].ulKind = PRSPEC_PROPID;
                PropSpec[2].propid = WIA_DIP_DEV_DESC;
                // Ask for the property values
                HRESULT hr = pWiaPropertyStorage->ReadMultiple(c_nPropertyCount, PropSpec, PropVar);
                if (hr == S_OK)
                {
                    // Check the return type for the device ID
                    if (VT_BSTR == PropVar[0].vt)
                    {
                        
                        //_tprintf(TEXT("WIA_DIP_DEV_ID: %ws\n"), PropVar[0].bstrVal);
                    }
                    // Check the return type for the device name
                    if (VT_BSTR == PropVar[1].vt)
                    {
                        
                        //_tprintf(TEXT("WIA_DIP_DEV_NAME: %ws\n"), PropVar[1].bstrVal);
                    }
                    // Check the return type for the device description
                    if (VT_BSTR == PropVar[2].vt)
                    {
                        
                        //_tprintf(TEXT("WIA_DIP_DEV_DESC: %ws\n"), PropVar[2].bstrVal);
                    }
                    FreePropVariantArray(c_nPropertyCount, PropVar);
                }
        

                pWiaPropertyStorage->Release();
                pWiaPropertyStorage = nullptr;
            }

        }

    }
    __finally
    {
        // Release COM interface.
        if (pWiaPropertyStorage) pWiaPropertyStorage->Release();
        if (pIEnum) pIEnum->Release();
        if (pItemRoot)
            pItemRoot->Release();

        if (pWiaDevMgr)
            pWiaDevMgr->Release();

        CoUninitialize();
    }
    return 0;
}