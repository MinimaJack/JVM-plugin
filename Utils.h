#ifndef __UTILS_H__
#define __UTILS_H__

#include "ComponentBase.h"
#include "AddInDefBase.h"
#include "IMemoryManager.h"
#include <string>
#include <include/jni.h>

uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint32_t len = 0);
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len = 0);
uint32_t getLenShortWcharStr(const WCHAR_T* Source);
std::wstring JStringToWString(JNIEnv* env, jstring string);
std::string getStdStringFrom1C(tVariant* paParams);
jstring getjstringFrom1C(JNIEnv* env, tVariant* paParams);
jobject getjdateFrom1C(JNIEnv* env, tVariant* paParams);
std::string getSignature(JNIEnv* env, tVariant* paParams, int start, int end);
jvalue* getParams(JNIEnv* env, tVariant* paParams, int start, int end);
void fromOleDate(struct tm* paTm, double paOADate);

#endif //__UTILS_H__
