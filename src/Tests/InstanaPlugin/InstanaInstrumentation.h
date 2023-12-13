#pragma once
#include "stdafx.h"
#include "InstanaInstrumentationMethodEntry.h"
#include "ModuleInfo.h"


//RCV:- we got this GUID from the VS Tools -> CreateGUID option
// {892888AB-5DA6-449F-AB2E-9C03AA048FD4}

class __declspec(uuid("892888AB-5DA6-449F-AB2E-9C03AA048FD4"))
    InstanaInstrumentation :
   public IInstrumentationMethod,
   public IInstrumentationMethodExceptionEvents,
   public CModuleRefCount
{
private:

    CComPtr<IProfilerManager> m_pProfilerManager;
    CComPtr<IProfilerStringManager> m_pStringManager;
    vector<shared_ptr<InstanaInstrumentMethodEntry>> m_instrumentMethodEntries;
    tstring m_strBinaryDir;
    shared_ptr<CInjectAssembly> m_spInjectAssembly;

private:
    static const WCHAR TestOutputPathEnvName[];
    static const WCHAR TestScriptFileEnvName[];
    static const WCHAR TestScriptFolder[];

public:
    DEFINE_DELEGATED_REFCOUNT_ADDREF(InstanaInstrumentation);
    DEFINE_DELEGATED_REFCOUNT_RELEASE(InstanaInstrumentation);
    STDMETHOD(QueryInterface)(_In_ REFIID riid, _Out_ void** ppvObject) override
    {
        return ImplQueryInterface(
            static_cast<IInstrumentationMethod*>(this),
            static_cast<IInstrumentationMethodExceptionEvents*>(this),
            riid,
            ppvObject
        );
    }

    // IInstrumentationMethod
public:
    virtual HRESULT STDMETHODCALLTYPE Initialize(_In_ IProfilerManager* pProfilerManager) override;
    virtual HRESULT ShouldInstrumentMethod(_In_ IMethodInfo* pMethodInfo, _In_ BOOL isRejit, _Out_ BOOL* pbInstrument)
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE OnAppDomainCreated(
        _In_ IAppDomainInfo* pAppDomainInfo)
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE OnAppDomainShutdown(
        _In_ IAppDomainInfo* pAppDomainInfo)
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE OnAssemblyLoaded(_In_ IAssemblyInfo* pAssemblyInfo)
    {
        return S_OK;
    }
    virtual HRESULT STDMETHODCALLTYPE OnAssemblyUnloaded(_In_ IAssemblyInfo* pAssemblyInfo)
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE OnModuleLoaded(_In_ IModuleInfo* pModuleInfo) override;

    virtual HRESULT STDMETHODCALLTYPE OnModuleUnloaded(_In_ IModuleInfo* pModuleInfo)
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE OnShutdown()
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE BeforeInstrumentMethod(_In_ IMethodInfo* pMethodInfo, _In_ BOOL isRejit)
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE InstrumentMethod(_In_ IMethodInfo* pMethodInfo, _In_ BOOL isRejit)
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE OnInstrumentationComplete(_In_ IMethodInfo* pMethodInfo, _In_ BOOL isRejit) { return S_OK; }

    virtual HRESULT STDMETHODCALLTYPE AllowInlineSite(_In_ IMethodInfo* pMethodInfoInlinee, _In_ IMethodInfo* pMethodInfoCaller, _Out_ BOOL* pbAllowInline)
    {
        return S_OK;
    }

public:
    virtual HRESULT STDMETHODCALLTYPE ExceptionCatcherEnter(
        _In_ IMethodInfo* pMethodInfo,
        _In_ UINT_PTR   objectId
    )
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionCatcherLeave()
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionSearchCatcherFound(
        _In_ IMethodInfo* pMethodInfo
    )
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionSearchFilterEnter(
        _In_ IMethodInfo* pMethodInfo
    )
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionSearchFilterLeave()
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionSearchFunctionEnter(
        _In_ IMethodInfo* pMethodInfo
    )
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionSearchFunctionLeave()
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionThrown(
        _In_ UINT_PTR thrownObjectId
    )
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionUnwindFinallyEnter(
        _In_ IMethodInfo* pMethodInfo
    )
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionUnwindFinallyLeave()
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionUnwindFunctionEnter(
        _In_ IMethodInfo* pMethodInfo
    )
    {
        return S_OK;
    }

    virtual HRESULT STDMETHODCALLTYPE ExceptionUnwindFunctionLeave()
    {
        return S_OK;
    }
    private:
        bool AddMemberRefs(IMetaDataAssemblyImport* pAssemblyImport, IMetaDataAssemblyEmit* pAssemblyEmit, IMetaDataEmit* pEmit, ModuleInfo* pModuleInfo, bool explicitlyInstrumented);
        HRESULT GetTypeRef(LPCWSTR typeName, mdToken sourceLibraryReference, IMetaDataEmit* pEmit, mdTypeRef* ptr);
        BOOL FindMscorlibReference(IMetaDataAssemblyImport* pAssemblyImport, mdAssemblyRef* rgAssemblyRefs, ULONG cAssemblyRefs, mdAssemblyRef* parMscorlib);

};

