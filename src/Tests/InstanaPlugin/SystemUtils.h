#pragma once
namespace InstanaSystemUtils
{
    HRESULT GetEnvVar(_In_z_ const WCHAR* wszVar, _Inout_ tstring& value);
}
