#include "stdafx.h"
#include "InstanaInstrumentation.h"
#include "InstrumentationEngineString.h"
#include "SystemUtils.h"
#include <libloaderapi.h>
#include "Util.h"

// Convenience macro for defining strings.
#define InstrStr(_V) InstanaInstrumentationEngineString _V(m_pStringManager)


#ifdef _WIN64
LPCWSTR k_wszEnteredFunctionProbeName = L"OnEnter";
LPCWSTR k_wszExitedFunctionProbeName = L"OnExit";
#else // Win32
LPCWSTR k_wszEnteredFunctionProbeName = L"OnEnter";
LPCWSTR k_wszExitedFunctionProbeName = L"OnExit";
#endif

// When pumping managed helpers into mscorlib, stick them into this pre-existing mscorlib type
LPCWSTR k_wszHelpersContainerType = L"System.CannotUnloadAppDomainException";

LPCWSTR ManagedTracingApiAssemblyNameClassic = L"Instana.ManagedTracing.Api";
LPCWSTR ManagedTracingAssemblyNameClassic = L"Instana.ManagedTracing.Base";
LPCWSTR ManagedTracingTracerClassNameClassic = L"Instana.ManagedTracing.Base.Tracer";

LPCWSTR ManagedTracingAssemblyNameCore = L"Instana.Tracing.Core";
LPCWSTR ManagedTracingAssemblyNameCoreCommon = L"Instana.Tracing.Core.Common";
LPCWSTR ManagedTracingAssemblyNameCoreTransport = L"Instana.Tracing.Core.Transport";
LPCWSTR ManagedTracingAssemblyNameApi = L"Instana.Tracing.Api";
LPCWSTR ManagedTracingTracerClassNameCore = L"Instana.Tracing.Tracer";


void AssertLogFailure(_In_ const WCHAR* wszError, ...)
{
#ifdef PLATFORM_UNIX
    // Attempting to interpret the var args is very difficult
    // cross-plat without a PAL. We should see about removing
    // the dependency on this in the Common.Lib
    string str;
    HRESULT hr = SystemString::Convert(wszError, str);
    if (!SUCCEEDED(hr))
    {
        str = "Unable to convert wchar string to char string";
    }
    AssertFailed(str.c_str());
#else
    // Since this is a test libary, we will print to standard error and fail
    va_list argptr;
    va_start(argptr, wszError);
    vfwprintf(stderr, wszError, argptr);
#endif
    throw "Assert failed";
}


HRESULT InstanaInstrumentation::Initialize(_In_ IProfilerManager* pProfilerManager)
{
    HRESULT hr;
    m_pProfilerManager = pProfilerManager;
    IfFailRet(m_pProfilerManager->QueryInterface(&m_pStringManager));

//#ifdef PLATFORM_UNIX
//    DWORD retVal = LoadInstrumentationMethodXml(this);
//#else
    // On Windows, xml reading is done in a single-threaded apartment using 
    // COM, so we need to spin up a new thread for it.
    //RCV:- Here we are using xml for the script reading and the other part,
    // Accoridng to Zivan we need something else like json file, instead of the json file

//    CHandle hConfigThread;
//    hConfigThread.Attach(CreateThread(NULL, 0, LoadInstrumentationMethodXml, this, 0, NULL));
//    DWORD retVal = WaitForSingleObject(hConfigThread, 15000);
//#endif

    
    return S_OK;
}

struct ComInitializer
{
#ifndef PLATFORM_UNIX
    HRESULT _hr;
    ComInitializer(DWORD mode = COINIT_APARTMENTTHREADED)
    {
        _hr = CoInitializeEx(NULL, mode);
    }
    ~ComInitializer()
    {
        if (SUCCEEDED(_hr))
        {
            CoUninitialize();
        }
    }
#endif
};
ModuleInfo moduleInfo;
HRESULT InstanaInstrumentation::OnModuleLoaded(_In_ IModuleInfo* pModuleInfo)
{
    HRESULT hr = S_OK;
    CComPtr<IProfilerManagerLogging> spLogger;
    CComQIPtr<IProfilerManager4> pProfilerManager4 = m_pProfilerManager;
    pProfilerManager4->GetGlobalLoggingInstance(&spLogger);
    try
    {
        ModuleID moduleID;
        pModuleInfo->GetModuleID(&moduleID);



        LPCBYTE pbBaseLoadAddr;
        WCHAR wszName[512];
        ULONG cchNameIn = _countof(wszName);
        ULONG cchNameOut;
        AssemblyID assemblyID;
        DWORD dwModuleFlags;
        HRESULT hr = E_FAIL;

        spLogger->LogMessage(_T("Enter!!!!"));

        ICorProfilerInfo4* m_pProfilerInfo;
        try
        {

            hr = m_pProfilerManager->GetCorProfilerInfo((IUnknown**)&m_pProfilerInfo);
            if (FAILED(hr))
            {
                spLogger->LogMessage(_T("Failed!!!!"));
                return hr;
            }
            hr = m_pProfilerInfo->GetModuleInfo2(
                moduleID,
                &pbBaseLoadAddr,
                cchNameIn,
                &cchNameOut,
                wszName,
                &assemblyID,
                &dwModuleFlags);
        }
        catch (...)
        {
            spLogger->LogMessage(_T("Error getting ModuleInfo"));
            return S_OK;
        }
        if (hr != S_OK)
        {
            spLogger->LogMessage(_T("Could not get Module info!"));
        }

        if (FAILED(hr))
        {
            spLogger->LogMessage(_T("GetModuleInfo failed for module {}", moduleID));
            return S_OK;
        }

        // adjust the buffer's size, call again
        if (cchNameIn <= cchNameOut)
        {
            spLogger->LogMessage(_T("The reserved buffer for the module's name was not large enough. Skipping."));
            return S_OK;
        }

        std::string logMessage2("Module loaded:  ");
        std::wstring lm2 = std::wstring(logMessage2.begin(), logMessage2.end());
        lm2 = lm2 + wszName;
        spLogger->LogMessage(lm2.c_str());

        if ((dwModuleFlags & COR_PRF_MODULE_WINDOWS_RUNTIME) != 0)
        {
            // Ignore any Windows Runtime modules.  We cannot obtain writeable metadata
            // interfaces on them or instrument their IL
            spLogger->LogMessage(_T("This is a runtime windows module, skipping"));
            return S_OK;
        }

        AppDomainID appDomainID;
        ModuleID modIDDummy;
        WCHAR assemblyName[255];
        ULONG assemblyNameLength = 0;

        hr = m_pProfilerInfo->GetAssemblyInfo(
            assemblyID,
            _countof(assemblyName),          // cchName,
            &assemblyNameLength,       // pcchName,
            assemblyName,       // szName[] ,
            &appDomainID,
            &modIDDummy);

        if (FAILED(hr))
        {
            spLogger->LogMessage(_T("GetAssemblyInfo failed for module {}", wszName));
            return S_OK;
        }

        LPCWSTR mscorlibName = L"mscorlib";
        if (_wcsicmp(assemblyName, mscorlibName) == 0)
        {
            spLogger->LogMessage(_T("this is mscorlib, noone does anything here..."));
            return S_OK;
        }

        LPCWSTR libName = L"TestClassLibrary";
        if (_wcsicmp(assemblyName, libName) != 0)
        {
            spLogger->LogMessage(_T("this is not TestClassLibrary..."));
            return S_OK;
        }

        bool instrumentationLoaded = false;
        // make sure we have the advices loaded from the config-file
        if (!instrumentationLoaded)
        {
            instrumentationLoaded = true;
        }

        // check whether the module being loaded is one that we are interested in. If not, abort
        /*bool instrumentThisModule = false;
        if (_metaUtil.ContainsAtEnd(wszName, L"Instana.Tracing.Core.dll") ||
            _metaUtil.ContainsAtEnd(wszName, L"Instana.Tracing.Core.Common.dll") ||
            _metaUtil.ContainsAtEnd(wszName, L"Instana.Tracing.Core.Transport.dll") ||
            _metaUtil.ContainsAtEnd(wszName, L"Instana.Tracing.Api.dll") ||
            _metaUtil.ContainsAtEnd(wszName, L"Instana.Tracing.Core.Instrumentation.dll") ||
            _metaUtil.ContainsAtEnd(wszName, L"Instana.ManagedTracing.dll") ||
            _metaUtil.ContainsAtEnd(wszName, L"Instana.ManagedTracing.Api.dll") ||
            _metaUtil.ContainsAtEnd(wszName, L"Instana.ManagedTracing.Base.dll") ||
            _metaUtil.ContainsAtEnd(wszName, L"Instana.ManagedTracing.Modules.dll"))
        {
            return S_OK;
        }*/
        std::string logMessage("Will instrument "); // initialized elsewhere
        std::wstring lm = std::wstring(logMessage.begin(), logMessage.end());
        lm = lm + assemblyName;
        spLogger->LogMessage(lm.c_str());

        CComPtr<IMetaDataEmit> pEmit;
        {
            CComPtr<IUnknown> pUnk;

            hr = m_pProfilerInfo->GetModuleMetaData(moduleID, ofWrite, IID_IMetaDataEmit, &pUnk);
            if (FAILED(hr))
            {
                spLogger->LogMessage(_T("GetModuleMetaData failed for module {}", moduleID));
                return S_OK;
            }

            hr = pUnk->QueryInterface(IID_IMetaDataEmit, (LPVOID*)&pEmit);
            if (FAILED(hr))
            {
                spLogger->LogMessage(_T("QueryInterface IID_IMetaDataEmit failed for module {}", moduleID));
                return S_OK;
            }
        }

        CComPtr<IMetaDataImport> pImport;
        {
            CComPtr<IUnknown> pUnk;

            hr = m_pProfilerInfo->GetModuleMetaData(moduleID, ofRead, IID_IMetaDataImport, &pUnk);
            if (FAILED(hr))
            {
                spLogger->LogMessage(_T("GetModuleMetaData failed for module {}", moduleID));
                return S_OK;
            }

            hr = pUnk->QueryInterface(IID_IMetaDataImport, (LPVOID*)&pImport);
            if (FAILED(hr))
            {
                spLogger->LogMessage(_T("QueryInterface IID_IMetaDataImport failed for module {}", moduleID));
                return S_OK;
            }
        }

        moduleInfo = { 0 };
        moduleInfo.m_pImport = pImport;
        moduleInfo.m_pImport->AddRef();
        //moduleInfo.m_pMethodDefToLatestVersionMap = new MethodDefToLatestVersionMap();

        if (wcscpy_s(moduleInfo.m_wszModulePath, _countof(moduleInfo.m_wszModulePath), wszName) != 0)
        {
            spLogger->LogMessage(_T("Failed to store module path", wszName));
            return S_OK;
        }

        if (wcscpy_s(moduleInfo.m_assemblyName, _countof(moduleInfo.m_assemblyName), assemblyName) != 0)
        {
            spLogger->LogMessage(_T("Failed to store assemblyName {}", assemblyName));
            return S_OK;
        }

        // Add the references to our helper methods.
        CComPtr<IMetaDataAssemblyEmit> pAssemblyEmit;
        {
            CComPtr<IUnknown> pUnk;

            hr = m_pProfilerInfo->GetModuleMetaData(moduleID, ofWrite, IID_IMetaDataAssemblyEmit, &pUnk);
            if (FAILED(hr))
            {
                spLogger->LogMessage(_T("GetModuleMetaData IID_IMetaDataAssemblyEmit failed for module {}", moduleID));
                return S_OK;
            }


            hr = pUnk->QueryInterface(IID_IMetaDataAssemblyEmit, (LPVOID*)&pAssemblyEmit);
            if (FAILED(hr))
            {
                //g_logger->error("IID_IMetaDataEmit: QueryInterface failed for ModuleID {}", moduleID);			LogInfo(L"QueryInterface IID_IMetaDataImport failed for module {}", moduleID);
                spLogger->LogMessage(_T("QueryInterface IID_IMetaDataAssemblyEmit failed for module {}", moduleID));
                return S_OK;
            }
        }

        CComPtr<IMetaDataAssemblyImport> pAssemblyImport;
        {
            CComPtr<IUnknown> pUnk;

            hr = m_pProfilerInfo->GetModuleMetaData(moduleID, ofRead, IID_IMetaDataAssemblyImport, &pUnk);
            if (FAILED(hr))
            {
                spLogger->LogMessage(_T("GetModuleMetaData IID_IMetaDataAssemblyImport failed for module {}", moduleID));
                return S_OK;
            }

            hr = pUnk->QueryInterface(IID_IMetaDataAssemblyImport, (LPVOID*)&pAssemblyImport);
            if (FAILED(hr))
            {
                spLogger->LogMessage(_T("QueryInterface IID_IMetaDataAssemblyImport failed for module {}", moduleID));
                return S_OK;
            }

        }
        bool instrumentThisModule = true;

        moduleInfo.m_pAssemblyImport = pAssemblyImport;
        moduleInfo.m_pAssemblyImport->AddRef();
        bool relevant = AddMemberRefs(pAssemblyImport, pAssemblyEmit, pEmit, &moduleInfo, instrumentThisModule);
        if (relevant)
        {
            spLogger->LogMessage(_T("Add member refs is done"));
            //LogInfo(L"Applying instrumentations to {}", moduleInfo.m_assemblyName);
            //m_moduleIDToInfoMap.Update(moduleID, moduleInfo);
            //_instrumentationCoordinator.registerModule(moduleID, moduleInfo);
            //****************ApplyInstrumentations(moduleID, moduleInfo, instrumentThisModule);


            // Append to the list!

            // If we already rejitted functions in other modules with a matching path, then
            // pre-rejit those functions in this module as well.  This takes care of the case
            // where we rejitted functions in a module loaded in one domain, and we just now
            // loaded the same module (unshared) into another domain.  We must explicitly ask to
            // rejit those functions in this domain's copy of the module, since it's identified
            // by a different ModuleID.

            std::vector<ModuleID> rgModuleIDs;
            std::vector<mdToken> rgMethodDefs;


        }
    }
    catch (std::exception& e)
    {
        spLogger->LogMessage(_T("Exception on loading module: {}", e.what()));
    }
    return S_OK;
}



bool InstanaInstrumentation::AddMemberRefs(IMetaDataAssemblyImport* pAssemblyImport, IMetaDataAssemblyEmit* pAssemblyEmit, IMetaDataEmit* pEmit, ModuleInfo* pModuleInfo, bool explicitlyInstrumented)
{
    CComPtr<IProfilerManagerLogging> spLogger;
    CComQIPtr<IProfilerManager4> pProfilerManager4 = m_pProfilerManager;
    pProfilerManager4->GetGlobalLoggingInstance(&spLogger);

    LPCWSTR mscorlibName = L"mscorlib";
    if (_wcsicmp(pModuleInfo->m_assemblyName, mscorlibName) == 0)
    {
        //LogInfo(L"this is mscorlib, noone does anything here...");
        return false;
    }
    bool prepareInstrumentation = explicitlyInstrumented;
    //assert(pModuleInfo != NULL);

    IMetaDataImport* pImport = pModuleInfo->m_pImport;

    HRESULT hr;

    COR_SIGNATURE sigFunctionEnterProbe[] = {
        IMAGE_CEE_CS_CALLCONV_DEFAULT,      // default calling convention
        0x04,                               // number of arguments == 2
        ELEMENT_TYPE_OBJECT,                  // return type == object
        ELEMENT_TYPE_OBJECT,
        ELEMENT_TYPE_SZARRAY, ELEMENT_TYPE_OBJECT,
        ELEMENT_TYPE_STRING,				// type name
        ELEMENT_TYPE_STRING					// method name
    };

    COR_SIGNATURE sigFunctionExitProbe[] = {
        IMAGE_CEE_CS_CALLCONV_DEFAULT,      // default calling convention
        0x05,                               // number of arguments == 3
        ELEMENT_TYPE_VOID,                  // return type == void
        ELEMENT_TYPE_OBJECT,				// instrumented object
        ELEMENT_TYPE_OBJECT,				// context obtained by entering
        ELEMENT_TYPE_SZARRAY, ELEMENT_TYPE_OBJECT, // parameters
        ELEMENT_TYPE_STRING,				// type name
        ELEMENT_TYPE_STRING					// method name
    };

    mdAssemblyRef rgMSCAssemblyRefs[20];
    ULONG cMSCAssemblyRefsReturned;
    mdAssemblyRef msCorLibRef = NULL;

    HCORENUM hmscEnum = NULL;
    do
    {
        hr = pAssemblyImport->EnumAssemblyRefs(
            &hmscEnum,
            rgMSCAssemblyRefs,
            _countof(rgMSCAssemblyRefs),
            &cMSCAssemblyRefsReturned);

        if (FAILED(hr))
        {
            spLogger->LogMessage(_T("Could not enumerate assemblies"));
            return false;
        }

        if (cMSCAssemblyRefsReturned == 0)
        {
            spLogger->LogMessage(_T("No assemblies have been returned, mscorlib not found?!"));
            return false;
        }

    } while (!FindMscorlibReference(
        pAssemblyImport,
        rgMSCAssemblyRefs,
        cMSCAssemblyRefsReturned,
        &msCorLibRef));
    pAssemblyImport->CloseEnum(hmscEnum);
    hmscEnum = NULL;

    hr = GetTypeRef(L"System.Object", msCorLibRef, pEmit, &(pModuleInfo->m_objectTypeRef));
    hr = GetTypeRef(L"System.Exception", msCorLibRef, pEmit, &(pModuleInfo->m_exceptionTypeRef));
    hr = GetTypeRef(L"System.Int16", msCorLibRef, pEmit, &(pModuleInfo->m_int16TypeRef));
    hr = GetTypeRef(L"System.Int32", msCorLibRef, pEmit, &(pModuleInfo->m_int32TypeRef));
    hr = GetTypeRef(L"System.Int64", msCorLibRef, pEmit, &(pModuleInfo->m_int64TypeRef));
    hr = GetTypeRef(L"System.Single", msCorLibRef, pEmit, &(pModuleInfo->m_float32TypeRef));
    hr = GetTypeRef(L"System.Double", msCorLibRef, pEmit, &(pModuleInfo->m_float64TypeRef));
    hr = GetTypeRef(L"System.UInt16", msCorLibRef, pEmit, &(pModuleInfo->m_uint16TypeRef));
    hr = GetTypeRef(L"System.UInt32", msCorLibRef, pEmit, &(pModuleInfo->m_uint32TypeRef));
    hr = GetTypeRef(L"System.UInt64", msCorLibRef, pEmit, &(pModuleInfo->m_uint64TypeRef));
    hr = GetTypeRef(L"System.Byte", msCorLibRef, pEmit, &(pModuleInfo->m_byteTypeRef));
    hr = GetTypeRef(L"System.SByte", msCorLibRef, pEmit, &(pModuleInfo->m_signedByteTypeRef));
    hr = GetTypeRef(L"System.Boolean", msCorLibRef, pEmit, &(pModuleInfo->m_boolTypeRef));
    hr = GetTypeRef(L"System.IntPtr", msCorLibRef, pEmit, &(pModuleInfo->m_intPtrTypeRef));

    bool g_IsDotNetCoreProcess = false;

    // find out whether the module references a module interesting for inheritance-rules
    // and store it.
    prepareInstrumentation = true;
    if (prepareInstrumentation)
    {
        WCHAR wszLocale[MAX_PATH];
        wcscpy_s(wszLocale, L"");

        // Generate assemblyRef for our api-assembly
        mdAssemblyRef apiAssemblyRef = NULL;
        // 741315d1d1544b22
        BYTE rgbApiPublicKeyToken[] = { 0x74, 0x13, 0x15, 0xd1, 0xd1, 0x54, 0x4b, 0x22 };

        ASSEMBLYMETADATA apiAssemblyMetaData;
        ZeroMemory(&apiAssemblyMetaData, sizeof(apiAssemblyMetaData));
        apiAssemblyMetaData.usMajorVersion = 1;
        apiAssemblyMetaData.usMinorVersion = 0;
        apiAssemblyMetaData.usBuildNumber = 0;
        apiAssemblyMetaData.usRevisionNumber = 0;
        apiAssemblyMetaData.szLocale = wszLocale;
        apiAssemblyMetaData.cbLocale = _countof(wszLocale);


        hr = pAssemblyEmit->DefineAssemblyRef(
            (void*)rgbApiPublicKeyToken,
            sizeof(rgbApiPublicKeyToken),
            (g_IsDotNetCoreProcess == TRUE ? ManagedTracingAssemblyNameApi : ManagedTracingApiAssemblyNameClassic),
            &apiAssemblyMetaData,
            NULL,                   // hash blob
            NULL,                   // cb of hash blob
            0,                      // flags
            &apiAssemblyRef);

        if (g_IsDotNetCoreProcess)
        {
            mdAssemblyRef commonAssemblyRef = NULL;
            BYTE rgbCommonPublicKeyToken[] = { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

            ASSEMBLYMETADATA commonAssemblyMetaData;
            ZeroMemory(&commonAssemblyMetaData, sizeof(commonAssemblyMetaData));
            commonAssemblyMetaData.usMajorVersion = 1;
            commonAssemblyMetaData.usMinorVersion = 0;
            commonAssemblyMetaData.usBuildNumber = 0;
            commonAssemblyMetaData.usRevisionNumber = 0;
            commonAssemblyMetaData.szLocale = wszLocale;
            commonAssemblyMetaData.cbLocale = _countof(wszLocale);

            hr = pAssemblyEmit->DefineAssemblyRef(
                (void*)rgbCommonPublicKeyToken,
                sizeof(rgbCommonPublicKeyToken),
                ManagedTracingAssemblyNameCoreCommon,
                &commonAssemblyMetaData,
                NULL,                   // hash blob
                NULL,                   // cb of hash blob
                0,                      // flags
                &commonAssemblyRef);

            if (hr == S_OK)
            {
                //DEBUGGER.WriteW(L"[PROFILERCALLBACK]\tLoad comon assembly\n");
            }

            mdAssemblyRef transportAssemblyRef = NULL;
            BYTE rgbTransportPublicKeyToken[] = { 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00 };

            ASSEMBLYMETADATA transportAssemblyMetaData;
            ZeroMemory(&transportAssemblyMetaData, sizeof(transportAssemblyMetaData));
            transportAssemblyMetaData.usMajorVersion = 1;
            transportAssemblyMetaData.usMinorVersion = 0;
            transportAssemblyMetaData.usBuildNumber = 0;
            transportAssemblyMetaData.usRevisionNumber = 0;
            transportAssemblyMetaData.szLocale = wszLocale;
            transportAssemblyMetaData.cbLocale = _countof(wszLocale);

            hr = pAssemblyEmit->DefineAssemblyRef(
                (void*)rgbTransportPublicKeyToken,
                sizeof(rgbTransportPublicKeyToken),
                ManagedTracingAssemblyNameCoreTransport,
                &transportAssemblyMetaData,
                NULL,                   // hash blob
                NULL,                   // cb of hash blob
                0,                      // flags
                &transportAssemblyRef);

            if (hr == S_OK)
            {
                //DEBUGGER.WriteW(L"[PROFILERCALLBACK]\tLoad transport assembly\n");
            }
        }

        mdAssemblyRef assemblyRef = NULL;
        mdTypeRef typeRef = mdTokenNil;
        // Generate assemblyRef for our managed-tracing assembly
        BYTE rgbPublicKeyToken[] = { 0x82, 0x6a, 0x13, 0x90, 0x12, 0x13, 0xc9, 0x89 };

        ASSEMBLYMETADATA assemblyMetaData;
        ZeroMemory(&assemblyMetaData, sizeof(assemblyMetaData));
        assemblyMetaData.usMajorVersion = 1;
        assemblyMetaData.usMinorVersion = 0;
        assemblyMetaData.usBuildNumber = 0;
        assemblyMetaData.usRevisionNumber = 0;
        assemblyMetaData.szLocale = wszLocale;
        assemblyMetaData.cbLocale = _countof(wszLocale);


        hr = pAssemblyEmit->DefineAssemblyRef(
            (void*)rgbPublicKeyToken,
            sizeof(rgbPublicKeyToken),
            (g_IsDotNetCoreProcess == TRUE ? ManagedTracingAssemblyNameCore : ManagedTracingAssemblyNameClassic),
            &assemblyMetaData,
            NULL,                   // hash blob
            NULL,                   // cb of hash blob
            0,                      // flags
            &assemblyRef);

        // Generate typeRef to our managed tracer

        hr = pEmit->DefineTypeRefByName(
            assemblyRef,
            (g_IsDotNetCoreProcess == TRUE ? ManagedTracingTracerClassNameCore : ManagedTracingTracerClassNameClassic),
            &typeRef);

        hr = pEmit->DefineMemberRef(
            typeRef,
            k_wszEnteredFunctionProbeName,
            sigFunctionEnterProbe,
            sizeof(sigFunctionEnterProbe),
            &(pModuleInfo->m_mdEnterProbeRef));

        hr = pEmit->DefineMemberRef(
            typeRef,
            k_wszExitedFunctionProbeName,
            sigFunctionExitProbe,
            sizeof(sigFunctionExitProbe),
            &(pModuleInfo->m_mdExitProbeRef));

    }
    return prepareInstrumentation;
}

HRESULT InstanaInstrumentation::GetTypeRef(LPCWSTR typeName, mdToken sourceLibraryReference, IMetaDataEmit* pEmit, mdTypeRef* ptr)
{
    HRESULT hr = pEmit->DefineTypeRefByName(sourceLibraryReference, typeName, ptr);
    return hr;
}

BOOL InstanaInstrumentation::FindMscorlibReference(IMetaDataAssemblyImport* pAssemblyImport, mdAssemblyRef* rgAssemblyRefs, ULONG cAssemblyRefs, mdAssemblyRef* parMscorlib)
{
    HRESULT hr;

    for (ULONG i = 0; i < cAssemblyRefs; i++)
    {
        const void* pvPublicKeyOrToken;
        ULONG cbPublicKeyOrToken;
        WCHAR wszName[512];
        ULONG cchNameReturned;
        ASSEMBLYMETADATA asmMetaData;
        ZeroMemory(&asmMetaData, sizeof(asmMetaData));
        const void* pbHashValue;
        ULONG cbHashValue;
        DWORD asmRefFlags;

        hr = pAssemblyImport->GetAssemblyRefProps(
            rgAssemblyRefs[i],
            &pvPublicKeyOrToken,
            &cbPublicKeyOrToken,
            wszName,
            _countof(wszName),
            &cchNameReturned,
            &asmMetaData,
            &pbHashValue,
            &cbHashValue,
            &asmRefFlags);

        if (FAILED(hr))
        {
            return FALSE;
        }

        LPCWSTR wszContainer = wszName;
        LPCWSTR wszProspectiveEnding = L"mscorlib";
        size_t cchContainer = wcslen(wszContainer);
        size_t cchEnding = wcslen(wszProspectiveEnding);


        if (_wcsicmp(
            wszProspectiveEnding,
            &(wszContainer[cchContainer - cchEnding])) != 0)
        {
            *parMscorlib = rgAssemblyRefs[i];
            return TRUE;
        }

        wszProspectiveEnding = L"netstandard";
        cchEnding = wcslen(wszProspectiveEnding);

        if (_wcsicmp(
            wszProspectiveEnding,
            &(wszContainer[cchContainer - cchEnding])) != 0)
        {
            *parMscorlib = rgAssemblyRefs[i];
            return TRUE;
        }

        wszProspectiveEnding = L"System.Runtime";
        cchEnding = wcslen(wszProspectiveEnding);

        if (_wcsicmp(
            wszProspectiveEnding,
            &(wszContainer[cchContainer - cchEnding])) != 0)
        {
            *parMscorlib = rgAssemblyRefs[i];
            return TRUE;
        }


    }

    return FALSE;
}
