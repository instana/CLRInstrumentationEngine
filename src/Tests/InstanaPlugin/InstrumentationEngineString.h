#pragma once
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License.

#pragma once

#include "stdafx.h"
#include "../Common.Headers/tstring.h"


/// <summary>
/// A class that demonstrates how to free an instrumentation engine string using
/// the IProfilerStringManager
/// </summary>
class InstanaInstrumentationEngineString
{
private:
    CComPtr<IProfilerStringManager> m_pStringManager;

public:
    BSTR m_bstr;
    InstanaInstrumentationEngineString(IProfilerStringManager* pStringManager) : m_bstr(nullptr), m_pStringManager(pStringManager) {}

    // non-copyable,moveable
    InstanaInstrumentationEngineString(InstanaInstrumentationEngineString& other) = delete;
    InstanaInstrumentationEngineString(InstanaInstrumentationEngineString&& other) noexcept
    {
        m_bstr = other.m_bstr;
        m_pStringManager = other.m_pStringManager;

        other.m_bstr = nullptr;
        other.m_pStringManager = nullptr;
    }

    InstanaInstrumentationEngineString& operator=(InstanaInstrumentationEngineString& other) = delete;

    InstanaInstrumentationEngineString& operator=(InstanaInstrumentationEngineString&& other) noexcept
    {
        if (this != &other)
        {
            m_pStringManager->FreeString(m_bstr);
            m_bstr = other.m_bstr;

            other.m_bstr = nullptr;
            other.m_pStringManager = nullptr;
        }

        return *this;
    }

    operator BSTR()
    {
        return m_bstr;
    }

    ~InstanaInstrumentationEngineString()
    {
        if (m_pStringManager != nullptr)
        {
            m_pStringManager->FreeString(m_bstr);
        }
    }

    const BSTR operator()()
    {
        return m_bstr;
    }

    BSTR BStr()
    {
        return m_bstr;
    }
};



