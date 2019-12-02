
#include "stdafx.h"

#ifdef __linux__
#include <unistd.h>
#include <stdlib.h>
#include <signal.h>
#include <time.h>
#include <errno.h>
#endif

#include <stdio.h>
#include <include/jni.h>
#include <wchar.h>
#include "Utils.h"
#include <string>


jvalue* getParams(JNIEnv* env, tVariant* paParams, int start, int end) {
	if (end < start) {
		return new jvalue[0];
	}
	jvalue *values = new jvalue[end - start];
	for (auto i = start; i < end; i++)
	{
		if (TV_VT(&paParams[i]) == VTYPE_PSTR || TV_VT(&paParams[i]) == VTYPE_PWSTR) {
			values[i - start].l = getjstringFrom1C(env, &paParams[i]);
		}
		else if (TV_VT(&paParams[i]) == VTYPE_I4) {
			values[i - start].i = TV_INT(&paParams[i]);
		}
		else if (TV_VT(&paParams[i]) == VTYPE_BOOL) {
			values[i - start].z = TV_BOOL(&paParams[i]);
		}
		else if (TV_VT(&paParams[i]) == VTYPE_R4) {
			values[i - start].f = TV_R4(&paParams[i]);
		}
	}
	return values;
}

std::string getSignature(JNIEnv* env, tVariant* paParams, int start, int end) {
	std::string signature = "(";
	for (auto i = start; i < end; i++)
	{
		if (TV_VT(&paParams[i]) == VTYPE_PSTR || TV_VT(&paParams[i]) == VTYPE_PWSTR) {
			signature.append("Ljava/lang/String;");
		}
		else if (TV_VT(&paParams[i]) == VTYPE_I4) {
			signature.append("I");
		}
		else if (TV_VT(&paParams[i]) == VTYPE_BOOL) {
			signature.append("Z");
		}
		else if (TV_VT(&paParams[i]) == VTYPE_R4) {
			signature.append("F");
		}
	}
	return signature.append(")");
}

std::wstring JStringToWString(JNIEnv* env, jstring string)
{
	std::wstring value;

	const jchar *raw = env->GetStringChars(string, 0);
	jsize len = env->GetStringLength(string);
	value.assign(raw, raw + len);
	env->ReleaseStringChars(string, raw);

	return value;
}

jstring getjstringFrom1C(JNIEnv* env, tVariant* paParams) {
	jstring jstr = nullptr;

	switch (TV_VT(paParams))
	{
	case VTYPE_PSTR:
		jstr = env->NewStringUTF(paParams->pstrVal);
		break;
	case VTYPE_PWSTR:
		wchar_t* name = TV_WSTR(paParams);
		jstr = env->NewString(reinterpret_cast<jchar*>(name), wcslen(name));

	}
	return jstr;
}

std::string getStdStringFrom1C(tVariant* paParams) {
	std::string resultString;
	char *name = 0;
	size_t size = 0;
	char *mbstr = 0;
	wchar_t* wsTmp = 0;
	switch (TV_VT(paParams))
	{
	case VTYPE_PSTR:
		name = paParams->pstrVal;
		resultString = std::string(name);
		break;
	case VTYPE_PWSTR:
		::convFromShortWchar(&wsTmp, TV_WSTR(paParams));
		size = wcstombs(0, wsTmp, 0) + 1;
		mbstr = new char[size];
		memset(mbstr, 0, size);
		size = wcstombs(mbstr, wsTmp, getLenShortWcharStr(TV_WSTR(paParams)));
		name = mbstr;
		delete[] wsTmp;
		resultString = std::string(name);
		if (mbstr)
			delete[] mbstr;
		break;

	}
	return resultString;
}

//---------------------------------------------------------------------------//
uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint32_t len)
{
    if (!len)
        len = ::wcslen(Source)+1;

    if (!*Dest)
        *Dest = new WCHAR_T[len];

    WCHAR_T* tmpShort = *Dest;
    wchar_t* tmpWChar = (wchar_t*) Source;
    uint32_t res = 0;

    ::memset(*Dest, 0, len*sizeof(WCHAR_T));
    do
    {
        *tmpShort++ = (WCHAR_T)*tmpWChar++;
        ++res;
    }
    while (len-- && *tmpWChar);

    return res;
}
//---------------------------------------------------------------------------//
uint32_t getLenShortWcharStr(const WCHAR_T* Source)
{
	uint32_t res = 0;
	WCHAR_T *tmpShort = (WCHAR_T*)Source;

	while (*tmpShort++)
		++res;

	return res;
}
//---------------------------------------------------------------------------//
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len)
{
    if (!len)
        len = getLenShortWcharStr(Source)+1;

    if (!*Dest)
        *Dest = new wchar_t[len];

    wchar_t* tmpWChar = *Dest;
    WCHAR_T* tmpShort = (WCHAR_T*)Source;
    uint32_t res = 0;

    ::memset(*Dest, 0, len*sizeof(wchar_t));
    do
    {
        *tmpWChar++ = (wchar_t)*tmpShort++;
        ++res;
    }
    while (len-- && *tmpShort);

    return res;
}

