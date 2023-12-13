// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "InstanaInstrumentation.h"

class InstanaCrossPlatFactory : public IClassFactory, public CModuleRefCount
{
private:

public:
    DEFINE_DELEGATED_REFCOUNT_ADDREF(InstanaCrossPlatFactory);
    DEFINE_DELEGATED_REFCOUNT_RELEASE(InstanaCrossPlatFactory);
    STDMETHOD(STDMETHODCALLTYPE) QueryInterface(REFIID riid, PVOID* ppvObject) override
    {
        HRESULT hr = E_NOINTERFACE;

        hr = ImplQueryInterface(
            static_cast<IClassFactory*>(this),
            riid, ppvObject
        );

        return hr;
    }

    // IClassFactory Methods
public:
    STDMETHOD(CreateInstance)(
        _In_ IUnknown* pUnkOuter,
        _In_   REFIID   riid,
        _Out_ void** ppvObject
        ) override
    {
        HRESULT hr = S_OK;

        CComPtr<InstanaInstrumentation> pInstrumentationMethod;
        pInstrumentationMethod.Attach(new InstanaInstrumentation);
        if (pInstrumentationMethod == NULL)
        {
            return E_OUTOFMEMORY;
        }

        IfFailRet(pInstrumentationMethod->QueryInterface(riid, ppvObject));

        return S_OK;
    }

    STDMETHOD(LockServer)(
        _In_ BOOL fLock
        ) override
    {
        return E_NOTIMPL;
    }
};

// Used to determine whether the DLL can be unloaded by OLE.
__control_entrypoint(DllExport)
STDAPI DLLEXPORT(DllCanUnloadNow, 0)(void)
{
    if (CModuleRefCount::GetModuleUsage() == 0)
    {
        return S_OK;
    }

    return S_FALSE;
}

// Returns a class factory to create an object of the requested type.
_Check_return_
STDAPI DLLEXPORT(DllGetClassObject, 12)(_In_ REFCLSID rclsid, _In_ REFIID riid, _Outptr_ LPVOID FAR* ppv)
{
    CComPtr<InstanaCrossPlatFactory> pClassFactory;
    pClassFactory.Attach(new InstanaCrossPlatFactory);
    if (pClassFactory != nullptr)
    {
        return pClassFactory->QueryInterface(riid, ppv);
    }

    return E_NOINTERFACE;
}
