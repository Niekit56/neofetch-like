#include <Windows.h>
#include <iostream>
#include <cstdio>
#include <comdef.h>
#include <Wbemidl.h>
#include <intrin.h>
#include <winioctl.h>
#include <string>
#include <tchar.h>
#include <stdio.h>
#include <comdef.h>
#include <memory>
#include <stdexcept>

#pragma comment(lib, "wbemuuid.lib")

class WMIService {
public:
    WMIService() : pLoc(nullptr), pSvc(nullptr), pEnumerator(nullptr) {
        // Initialization of COM library
        HRESULT hres = CoInitializeEx(0, COINIT_MULTITHREADED);
        if (FAILED(hres)) {
            // error handling
        }

        // Creating a pointer to a WMI locator
        hres = CoCreateInstance(
            CLSID_WbemLocator,
            0,
            CLSCTX_INPROC_SERVER,
            IID_IWbemLocator,
            (LPVOID*)&pLoc
        );

        if (FAILED(hres)) {
        }

        // Connection to WMI via the specified locator
        hres = pLoc->ConnectServer(
            _bstr_t(L"ROOT\\CIMV2"), 
            nullptr,                 
            nullptr,                 
            0,                       
            NULL,                 
            0,                       
            0,                       
            &pSvc                    
        );

        if (FAILED(hres)) {
        }

        // Set the authentication level to the default authentication level
        hres = CoSetProxyBlanket(
            pSvc,                       
            RPC_C_AUTHN_WINNT,          
            RPC_C_AUTHZ_NONE,           
            nullptr,                    
            RPC_C_AUTHN_LEVEL_CALL,     
            RPC_C_IMP_LEVEL_IMPERSONATE,
            nullptr,                    
            EOAC_NONE                   
        );

        if (FAILED(hres)) {
        }

        // Executing WMI request
        hres = pSvc->ExecQuery(
            bstr_t("WQL"),
            bstr_t("SELECT * FROM Win32_OperatingSystem"),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            nullptr,
            &pEnumerator
        );

        if (FAILED(hres)) {
        }
    }

    ~WMIService() {
        // Cleaning
        if (pEnumerator) {
            pEnumerator->Release();
            pEnumerator = nullptr; 
        }
        if (pSvc) {
            pSvc->Release();
            pSvc = nullptr; 
        }
        if (pLoc) {
            pLoc->Release();
            pLoc = nullptr; 
        }
        CoUninitialize();
    }


    void GetWindowsVersion() {
        IWbemClassObject* pclsObj = nullptr;
        ULONG uReturn = 0;

        while (pEnumerator) {
            HRESULT hres = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

            if (uReturn == 0)
                break;

            VARIANT vtProp;

            // Getting the value of "Version" property
            hres = pclsObj->Get(L"Version", 0, &vtProp, 0, 0);

            // Split version string
            std::wstring versionString(vtProp.bstrVal);
            size_t pos = versionString.find_last_of(L'.');
            if (pos != std::wstring::npos) {
                versionString = versionString.substr(0, pos); // trim to the last point
            }

            wprintf(L"Version: %s\n", versionString.c_str());
            VariantClear(&vtProp);

            pclsObj->Release();
            pclsObj = nullptr; 
        }
        pEnumerator = nullptr;
    }
private:
    IWbemLocator* pLoc;
    IWbemServices* pSvc;
    IEnumWbemClassObject* pEnumerator;
};

class CPUInfoFetcher {
public:
    CPUInfoFetcher() {
        std::fill_n(CPUInfo, 4, -1);
        std::memset(CPUBrandString, 0, sizeof(CPUBrandString));
        __cpuid(CPUInfo, 0x80000000);
        unsigned int nExIds = CPUInfo[0];
        // Iterate through extended CPUID function IDs
        for (unsigned int i = 0x80000000; i <= nExIds; ++i) {
            __cpuidex(CPUInfo, i, 0);
            if (i == 0x80000002)
                std::memcpy(CPUBrandString, CPUInfo, sizeof(CPUInfo));
            else if (i == 0x80000003)
                std::memcpy(CPUBrandString + 16, CPUInfo, sizeof(CPUInfo));
            else if (i == 0x80000004)
                std::memcpy(CPUBrandString + 32, CPUInfo, sizeof(CPUInfo));
        }
    }
    // Print CPU brand information
    void printCPUInfo() const {
        std::cout << "CPU: " << CPUBrandString << std::endl;
    }

private:
    int CPUInfo[4];
    char CPUBrandString[0x40];
};

class MemoryInfo {
public:
    MemoryInfo() {
        ms.dwLength = sizeof(ms);
        GlobalMemoryStatusEx(&ms);
    }
    // print information
    void printMemoryInfo() const {
        std::cout << "The total amount of physical memory: " << static_cast<int>(ms.ullTotalPhys / 1024 / 1024) << "Mb" << std::endl;
        std::cout << "Free physical memory capacity: " << static_cast<int>(ms.ullAvailPhys / 1024 / 1024) << "Mb" << std::endl;
    }

private:
    MEMORYSTATUSEX ms;
};

class VideoCardInfo {
public:
    VideoCardInfo() {
        // Инициализация COM библиотеки
        CoInitializeEx(0, COINIT_MULTITHREADED);
    }

    ~VideoCardInfo() {
        // Очистка ресурсов при завершении работы с объектом
        CoUninitialize();
    }

    void GetInfo() {
        HRESULT hres;

        // WMI initialization
        IWbemLocator* pLoc = nullptr;
        hres = CoCreateInstance(
            CLSID_WbemLocator,
            nullptr,
            CLSCTX_INPROC_SERVER,
            IID_IWbemLocator,
            reinterpret_cast<LPVOID*>(&pLoc)
        );

        if (FAILED(hres)) {
            std::cerr << "Failed to create IWbemLocator object. Error code: " << hres << std::endl;
            return;
        }

        // Connection to WMI
        IWbemServices* pSvc = nullptr;
        hres = pLoc->ConnectServer(
            _bstr_t(L"ROOT\\CIMv2"),
            nullptr,
            nullptr,
            nullptr,
            0,
            nullptr,
            nullptr,
            &pSvc
        );

        if (FAILED(hres)) {
            std::cerr << "Failed to connect to ROOT\\CIMv2. Error code: " << hres << std::endl;
            pLoc->Release();
            return;
        }

        // Set security for calling WMI methods
        hres = CoSetProxyBlanket(
            pSvc,
            RPC_C_AUTHN_WINNT,
            RPC_C_AUTHZ_NONE,
            nullptr,
            RPC_C_AUTHN_LEVEL_CALL,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            nullptr,
            EOAC_NONE
        );

        if (FAILED(hres)) {
            std::cerr << "Failed to set proxy blanket. Error code: " << hres << std::endl;
            pSvc->Release();
            pLoc->Release();
            return;
        }

        // Executing WMI request
        IEnumWbemClassObject* pEnumerator = nullptr;
        hres = pSvc->ExecQuery(
            bstr_t("WQL"),
            bstr_t("SELECT * FROM Win32_VideoController"),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            nullptr,
            &pEnumerator
        );

        if (FAILED(hres)) {
            std::cerr << "Failed to execute WMI query. Error code: " << hres << std::endl;
            pSvc->Release();
            pLoc->Release();
            return;
        }

        // Get video card data
        IWbemClassObject* pclsObj = nullptr;
        ULONG uReturn = 0;

        while (pEnumerator) {
            HRESULT hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

            if (0 == uReturn) {
                break;
            }

            VARIANT vtProp;

            // Getting the "Name" property (video card name)
            hr = pclsObj->Get(L"Name", 0, &vtProp, 0, 0);
            std::wcout << L"GPU: " << vtProp.bstrVal << std::endl;
            VariantClear(&vtProp);

            // Getting "AdapterRAM" property (amount of video memory)
            hr = pclsObj->Get(L"AdapterRAM", 0, &vtProp, 0, 0);
            std::wcout << L"GPU RAM: " << vtProp.uintVal / (1024 * 1024) << L" MB" << std::endl; // Convert to megabytes
            VariantClear(&vtProp);

            pclsObj->Release();
        }

        // Resource clearing
        pSvc->Release();
        pLoc->Release();
        pEnumerator->Release();
    }
};

class HardDriveInfo {
public:
    HardDriveInfo() {
        // Initialization of COM library
        CoInitializeEx(0, COINIT_MULTITHREADED);

        // WMI initialization
        HRESULT hr = CoCreateInstance(
            CLSID_WbemLocator,
            0,
            CLSCTX_INPROC_SERVER,
            IID_IWbemLocator,
            reinterpret_cast<LPVOID*>(&pLoc)
        );

        if (FAILED(hr)) {
            std::cerr << "Failed to initialize IWbemLocator. Error code: " << hr << std::endl;
            return;
        }

        hr = pLoc->ConnectServer(
            _bstr_t(L"ROOT\\CIMV2"),
            NULL,
            NULL,
            0,
            NULL,
            0,
            0,
            &pSvc
        );

        if (FAILED(hr)) {
            std::cerr << "Не удалось подключиться к серверу WMI. Код ошибки: " << hr << std::endl;
            return;
        }

        // Set the security level for authentication
        CoSetProxyBlanket(
            pSvc,
            RPC_C_AUTHN_WINNT,
            RPC_C_AUTHZ_NONE,
            NULL,
            RPC_C_AUTHN_LEVEL_CALL,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            NULL,
            EOAC_NONE
        );
    }

    ~HardDriveInfo() {
        // Freeing up resources
        if (pSvc != NULL) {
            pSvc->Release();
        }
        if (pLoc != NULL) {
            pLoc->Release();
        }

        // Finalize the COM library
        CoUninitialize();
    }

    void PrintDriveInfo() {
        if (pSvc == NULL) {
            std::cerr << "Failed to get disk information." << std::endl;
            return;
        }

        IEnumWbemClassObject* pEnumerator = NULL;
        HRESULT hr = pSvc->ExecQuery(
            bstr_t("WQL"),
            bstr_t("SELECT * FROM Win32_DiskDrive"),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            NULL,
            &pEnumerator
        );

        if (FAILED(hr)) {
            std::cerr << "Failed to execute WMI request. Error code: " << hr << std::endl;
            return;
        }

        IWbemClassObject* pclsObj = NULL;
        ULONG uReturn = 0;

        while (pEnumerator) {
            hr = pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn);

            if (uReturn == 0) {
                break;
            }

            VARIANT vtProp;

            // Get the "Model" property of the Win32_DiskDrive object
            hr = pclsObj->Get(L"Model", 0, &vtProp, 0, 0);
            wprintf(L"Disk Model: %s\n", vtProp.bstrVal);

            VariantClear(&vtProp);

            pclsObj->Release();
        }

        pEnumerator->Release();
    }

private:
    IWbemLocator* pLoc;
    IWbemServices* pSvc;
};

class MotherboardInfo {
public:
    MotherboardInfo() {
        // COM initialization
        HRESULT hres = CoInitializeEx(0, COINIT_MULTITHREADED);
        if (FAILED(hres)) {
            std::cerr << "Failed to initialize COM library. Error code: 0x" << std::hex << hres << std::endl;
        }

        // Security initialization (required for WMI)
        hres = CoInitializeSecurity(
            NULL,
            -1,
            NULL,
            NULL,
            RPC_C_AUTHN_LEVEL_DEFAULT,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            NULL,
            EOAC_NONE,
            NULL
        );
        // Initialization of COM interface pointer
        hres = CoCreateInstance(
            CLSID_WbemLocator,
            0,
            CLSCTX_INPROC_SERVER,
            IID_IWbemLocator,
            (LPVOID*)&pLoc
        );

        if (FAILED(hres)) {
            std::cerr << "Failed to create IWbemLocator object. Error code: 0x" << std::hex << hres << std::endl;
            CoUninitialize();
        }

        // Connection to WMI
        hres = pLoc->ConnectServer(
            _bstr_t(L"ROOT\\CIMv2"),
            NULL,
            NULL,
            0,
            NULL,
            0,
            0,
            &pSvc
        );

        if (FAILED(hres)) {
            std::cerr << "Failed to connect to WMI. Error code: 0x" << std::hex << hres << std::endl;
            pLoc->Release();
            CoUninitialize();
        }
        // Set security levels on the proxy
        hres = CoSetProxyBlanket(
            pSvc,
            RPC_C_AUTHN_WINNT,
            RPC_C_AUTHZ_NONE,
            NULL,
            RPC_C_AUTHN_LEVEL_CALL,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            NULL,
            EOAC_NONE
        );

        if (FAILED(hres)) {
            std::cerr << "Failed to set proxy blanket. Error code: 0x" << std::hex << hres << std::endl;
            pSvc->Release();
            pLoc->Release();
            CoUninitialize();
        }
    }

    void GetInfo() {
        // WMI query for motherboard information
        IEnumWbemClassObject* pEnumerator = NULL;
        HRESULT hres = pSvc->ExecQuery(
            bstr_t("WQL"),
            bstr_t("SELECT * FROM Win32_BaseBoard"),
            WBEM_FLAG_FORWARD_ONLY | WBEM_FLAG_RETURN_IMMEDIATELY,
            NULL,
            &pEnumerator
        );

        if (FAILED(hres)) {
            std::cerr << "Query for motherboard information failed. Error code: 0x" << std::hex << hres << std::endl;
            return;
        }

        // Retrieve motherboard information from the query
        IWbemClassObject* pclsObj = NULL;
        ULONG uReturn = 0;

        while (pEnumerator->Next(WBEM_INFINITE, 1, &pclsObj, &uReturn) == S_OK) {
            VARIANT vtProp;
            hres = pclsObj->Get(L"Manufacturer", 0, &vtProp, 0, 0);
            wprintf(L"Manufacturer: %s\n", vtProp.bstrVal);
            VariantClear(&vtProp);

            hres = pclsObj->Get(L"Product", 0, &vtProp, 0, 0);
            wprintf(L"Product: %s\n", vtProp.bstrVal);
            VariantClear(&vtProp);

            pclsObj->Release();
            pclsObj = nullptr; 
        }

        // Release the enumerator
        pEnumerator->Release();
        pEnumerator = nullptr; 
    }

private:
    IWbemLocator* pLoc;
    IWbemServices* pSvc;
};

class ConsoleColor {
public:
    static void setGreen() {
        HANDLE hConsole = GetStdHandle(STD_OUTPUT_HANDLE);
        SetConsoleTextAttribute(hConsole, FOREGROUND_GREEN);
    }
};

int main() {
    ConsoleColor::setGreen();
    setlocale(LC_ALL, "eng");
    SetConsoleOutputCP(1251);

    WMIService wmiService;
    wmiService.GetWindowsVersion();
    std::wcout << std::endl;
 
    CPUInfoFetcher cpuInfoFetcher;
    cpuInfoFetcher.printCPUInfo();
    std::wcout << std::endl;

    MemoryInfo memoryInfo;
    memoryInfo.printMemoryInfo();
    std::wcout << std::endl;

    HardDriveInfo hardDriveInfo;
    hardDriveInfo.PrintDriveInfo();
    std::wcout << std::endl;

    VideoCardInfo videoCard;
    videoCard.GetInfo();
    std::wcout << std::endl;

    MotherboardInfo motherboardInfo;
    motherboardInfo.GetInfo();

    getchar();

    return 0;
}