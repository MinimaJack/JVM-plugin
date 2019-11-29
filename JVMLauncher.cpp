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
#include "JVMLauncher.h"

#define BASE_ERRNO     7

static jobject s_threadLock;

static IAddInDefBase *gAsyncEvent = NULL;


static wchar_t *g_PropNames[] = { L"IsEnabled", L"javaHome", L"libraryDir" };
static wchar_t *g_MethodNames[] = { L"LaunchInJVM",L"LaunchInJVMP",L"LaunchInJVMPP", L"CallFInJVMB", L"CallFInJVMBP", L"CallFInJVMBPP", L"CallFInJVM", L"CallFInJVMP", L"CallFInJVMPP", L"Disable", L"AddJar" };

static wchar_t *g_PropNamesRu[] = { L"Включен", L"javaHome", L"libraryDir" };
static wchar_t *g_MethodNamesRu[] = { L"LaunchInJVM",L"LaunchInJVMP",L"LaunchInJVMPP", L"CallFInJVMB", L"CallFInJVMBP", L"CallFInJVMBPP", L"CallFInJVM", L"CallFInJVMP", L"CallFInJVMPP", L"Выключить", L"AddJar" };


static void JNICALL Java_Runner_log(JNIEnv *env, jobject thisObj, jstring info) {
	wchar_t *who = L"ComponentNative", *what = L"Java";
	auto wstring = JStringToWString(env, info);
	WCHAR_T *err = 0;

	::convToShortWchar(&err, wstring.c_str());
	gAsyncEvent->ExternalEvent(who, what, err);
	delete[]err;

}

bool JVMLauncher::endCall(JNIEnv* env)
{
	m_JVMEnv = env;
	jobject exh = m_JVMEnv->ExceptionOccurred();
	if (exh) {
		jclass classClass = this->m_JVMEnv->GetObjectClass(exh);
		auto getClassLoaderMethod = this->m_JVMEnv->GetMethodID(classClass, "getLocalizedMessage", "()Ljava/lang/String;");
		auto info = (jstring)this->m_JVMEnv->CallObjectMethod(exh, getClassLoaderMethod);
		if (m_JVMEnv->IsSameObject(info, NULL)) {
			jmethodID mid = env->GetMethodID(classClass, "getClass", "()Ljava/lang/Class;");
			jobject clsObj = env->CallObjectMethod(exh, mid);
			jclass classClass = env->GetObjectClass(clsObj);

			auto getNameMethod = this->m_JVMEnv->GetMethodID(classClass, "getName", "()Ljava/lang/String;");
			info = (jstring)this->m_JVMEnv->CallObjectMethod(clsObj, getNameMethod);
		}

		auto wstring = JStringToWString(env, info);
		WCHAR_T *err = 0;

		::convToShortWchar(&err, wstring.c_str());
		pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", err, 10);
		delete[]err;

		m_JVMEnv->ExceptionClear();

		return false;
	}
	return true;
}

jobjectArray JVMLauncher::JNI_newObjectArray(jsize length, jclass elementClass, jobject initialElement, bool& hasError)
{
	jobjectArray result;
	BEGIN_JAVA
		result = env->NewObjectArray(length, elementClass, initialElement);
	END_JAVA
		return result;
}

jclass JVMLauncher::JNI_findClass(const char* className)
{
	jclass result;
	BEGIN_JAVA
		result = env->FindClass(className);
	END_JAVA
		return result;
}

void JVMLauncher::JNI_callStaticVoidMethod(jclass clazz, jmethodID methodID, bool& hasError, ...)
{
	va_list args;
	va_start(args, methodID);
	JVMLauncher::JNI_callStaticVoidMethodV(clazz, methodID, args, hasError);
	va_end(args);
}

jobject JVMLauncher::JNI_callStaticObjectMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError)
{
	jobject result;
	BEGIN_CALL
		result = (*env).CallStaticObjectMethodA(clazz, methodID, args);
	END_CALL
		return result;
}

jboolean JVMLauncher::JNI_callStaticBooleanMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError)
{
	jboolean result;
	BEGIN_CALL
		result = (*env).CallStaticBooleanMethodA(clazz, methodID, args);
	END_CALL
		return result;
}
jint JVMLauncher::JNI_callStaticIntMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError)
{
	jint result;
	BEGIN_CALL
		result = (*env).CallStaticIntMethodA(clazz, methodID, args);
	END_CALL
		return result;
}
jfloat JVMLauncher::JNI_callStaticFloatMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError)
{
	jfloat result;
	BEGIN_CALL
		result = (*env).CallStaticFloatMethodA(clazz, methodID, args);
	END_CALL
		return result;
}
jdouble JVMLauncher::JNI_callStaticDoubleMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError)
{
	jdouble result;
	BEGIN_CALL
		result = (*env).CallStaticDoubleMethodA(clazz, methodID, args);
	END_CALL
		return result;
}

void JVMLauncher::JNI_callStaticVoidMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError)
{
	BEGIN_CALL
	(*env).CallStaticVoidMethodA(clazz, methodID, args);
	END_CALL
}

void JVMLauncher::JNI_callStaticVoidMethodV(jclass clazz, jmethodID methodID, va_list args, bool& hasError)
{
	BEGIN_CALL
	(*env).CallStaticVoidMethodV(clazz, methodID, args);
	END_CALL
}
jmethodID JVMLauncher::JNI_getStaticMethodID(jclass clazz, const char* name, const char* sig, bool& hasError)
{
	jmethodID result;
	BEGIN_CALL
		result = (*env).GetStaticMethodID(clazz, name, sig);
	END_CALL
		return result;
}

void JVMLauncher::addJar(const char* jarname)
{
	if (m_boolEnabled) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"JVM already launched", 6);
		return;
	}
	m_listOfJars.push_back(jarname);
}

bool JVMLauncher::verify() {
	if (!m_boolEnabled) {
		return true;
	};
	return true;
}

JVMLauncher::JVMLauncher() {
	// Check for JAVA_HOME
	char *pValue = getenv("JAVA_HOME");
	if (pValue != NULL) {
		m_JavaHome = pValue;
	}

	m_ProductLibDir = "d:/";

}

void JVMLauncher::LaunchJVM() {
	if (m_hDllInstance == nullptr) {
		if (m_JavaHome.empty()) {
			pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"Set JAVA_HOME environment", 1);
			return;
		}
		m_JvmDllLocation = m_JavaHome + "/jre/bin/server/jvm.dll";
		m_hDllInstance = LoadLibraryA(m_JvmDllLocation.c_str());
	}
	if (m_hDllInstance == nullptr) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"Cannot find JDK", 1);
		return;
	}
	else {

		if (m_GetCreatedJavaVMs == nullptr) {
			m_GetCreatedJavaVMs = (GetCreatedJavaVMs)GetProcAddress(m_hDllInstance, "JNI_GetCreatedJavaVMs");
		}

		int n;
		jint retval = m_GetCreatedJavaVMs(&m_RunningJVMInstance, 1, (jsize*)&n);

		if (retval == JNI_OK)
		{

			if (n == 0)
			{
				if (m_JVMInstance == nullptr) {
					m_JVMInstance = (CreateJavaVM)GetProcAddress(m_hDllInstance, "JNI_CreateJavaVM");
				}
				std::string strJavaClassPath;
				std::string strJavaLibraryPath;

				strJavaClassPath = "-Djava.class.path=";
				for (std::size_t idx = 0; idx < m_listOfJars.size(); idx++) {
					strJavaClassPath += m_ProductLibDir + "/" + m_listOfJars[idx] + ";";
				}

				strJavaLibraryPath = "-Djava.library.path=";
				strJavaLibraryPath += m_JavaHome + "/lib" + "," + m_JavaHome + "/jre/lib";

				JavaVMOption options[6];
				options[0].optionString = const_cast<char*>(strJavaClassPath.c_str());
				options[1].optionString = const_cast<char*>(strJavaLibraryPath.c_str());
				options[2].optionString = "-Xms128m";
				options[3].optionString = "-Xmx1024m";
				options[4].optionString = "-Xcheck:jni:nonfatal";
				options[5].optionString = "-XX:+DisableAttachMechanism";


				JavaVMInitArgs vm_args;
				vm_args.version = JNI_VERSION_1_8;
				vm_args.options = options;
				vm_args.nOptions = 5;
				vm_args.ignoreUnrecognized = JNI_TRUE;

				jint res = m_JVMInstance(&m_RunningJVMInstance, (void **)&m_JVMEnv, &vm_args);
				if (res != JNI_OK) {
					pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"Could not launch the JVM", 3);
				}
				else {
					m_boolEnabled = true;
					if (!verify()) {
						pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"Cannot verify jar", 8);
						this->JNI_destroyVM();
					}
				}

			}
			else
			{
				m_boolEnabled = true;
			}
		}
	}
}

jint JVMLauncher::JNI_destroyVM() {
	jint result = m_RunningJVMInstance->DestroyJavaVM();
	m_JVMEnv = nullptr;
	s_threadLock = nullptr;
	return result;
}

bool JVMLauncher::validateCall() {
	LaunchJVM();

	if (!m_boolEnabled) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"JVM not running", 5);
		return false;
	}
	if (!valid) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"Load not valid packages", 5);
		return false;
	}
	return true;
}

jclass JVMLauncher::findClassForCall(std::string className) {
	jclass findedClass = nullptr;
	auto val = m_cachedClasses.find(className);
	if (val == m_cachedClasses.end()) {
		jclass neededclass = this->m_JVMEnv->FindClass(className.c_str());
		if (!this->m_JVMEnv->ExceptionCheck()) {
			findedClass = (jclass)this->m_JVMEnv->NewGlobalRef(neededclass);
			m_cachedClasses.insert(std::pair<std::string, jclass>(className, findedClass));
			this->m_JVMEnv->DeleteLocalRef(neededclass);
			JNINativeMethod methods[]{ { "log", "(Ljava/lang/String;)V", (void *)&Java_Runner_log } };  // mapping table
			if (m_JVMEnv->RegisterNatives(findedClass, methods, 1) < 0) {
				if (m_JVMEnv->ExceptionOccurred())                                        // verify if it's ok
					pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L" OOOOOPS: exception when registreing natives handlers", 6);
			}
		}
		else {
			this->m_JVMEnv->ExceptionClear();
		}
	}
	else {
		findedClass = val->second;
	}
	return findedClass;
}

jmethodID JVMLauncher::findMethodForCall(jclass findedclass, bool& hasError) {
	jmethodID findedMethod = nullptr;
	auto val = m_cachedMethod.find(findedclass);
	if (val == m_cachedMethod.end()) {
		hasError = true;
	}
	else {
		findedMethod = val->second;
	}
	return findedMethod;
}

jmethodID JVMLauncher::HasMethodInCache(jclass findedclass) {
	jmethodID findedMethod = nullptr;
	auto val = m_cachedMethod.find(findedclass);
	if (val != m_cachedMethod.end()) {
		findedMethod = val->second;
	}
	return findedMethod;
}

/////////////////////////////////////////////////////////////////////////////
// ILanguageExtenderBase
//---------------------------------------------------------------------------//
bool JVMLauncher::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{
	wchar_t *wsExtension = L"JVMLauncher";
	size_t iActualSize = ::wcslen(wsExtension) + 1;

	if (m_iMemory && m_iMemory->AllocMemory((void**)wsExtensionName, iActualSize * sizeof(wchar_t))) {
		::convToShortWchar(wsExtensionName, wsExtension, iActualSize);
		return true;
	}

	return false;
}
//---------------------------------------------------------------------------//
long JVMLauncher::GetNProps()
{
	return ePropLast;
}
//---------------------------------------------------------------------------//
long JVMLauncher::FindProp(const WCHAR_T* wsPropName)
{
	long plPropNum = -1;
	wchar_t* propName = 0;

	::convFromShortWchar(&propName, wsPropName);
	plPropNum = findName(g_PropNames, propName, ePropLast);

	if (plPropNum == -1)
		plPropNum = findName(g_PropNamesRu, propName, ePropLast);

	delete[] propName;

	return plPropNum;
}
//---------------------------------------------------------------------------//
const WCHAR_T* JVMLauncher::GetPropName(long lPropNum, long lPropAlias)
{
	if (lPropNum >= ePropLast)
		return NULL;

	wchar_t *wsCurrentName = NULL;
	WCHAR_T *wsPropName = NULL;
	int iActualSize = 0;

	switch (lPropAlias)
	{
	case 0: // First language
		wsCurrentName = g_PropNames[lPropNum];
		break;
	case 1: // Second language
		wsCurrentName = g_PropNamesRu[lPropNum];
		break;
	default:
		return 0;
	}

	iActualSize = wcslen(wsCurrentName) + 1;

	if (m_iMemory && wsCurrentName)
	{
		if (m_iMemory->AllocMemory((void**)&wsPropName, iActualSize * sizeof(WCHAR_T)))
			::convToShortWchar(&wsPropName, wsCurrentName, iActualSize);
	}

	return wsPropName;
}
//---------------------------------------------------------------------------//
bool JVMLauncher::GetPropVal(const long lPropNum, tVariant* pvarPropVal)
{
	char *name = 0;
	size_t size = 0;
	wchar_t* wsMsgBuf;

	switch (lPropNum)
	{
	case ePropIsEnabled:
		TV_VT(pvarPropVal) = VTYPE_BOOL;
		TV_BOOL(pvarPropVal) = m_boolEnabled;
		break;
	case ePropJavaHome:
		TV_VT(pvarPropVal) = VTYPE_PWSTR;
		name = const_cast<char*>(this->m_JavaHome.c_str());
		size = mbstowcs(0, name, 0) + 1;
		wsMsgBuf = new wchar_t[size];
		memset(wsMsgBuf, 0, size * sizeof(wchar_t));
		size = mbstowcs(wsMsgBuf, name, size);
		pvarPropVal->wstrLen = size;
		m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, size * sizeof(WCHAR_T));
		::convToShortWchar(&pvarPropVal->pwstrVal, wsMsgBuf, size);
		break;
	case ePropLibraryDir:
		TV_VT(pvarPropVal) = VTYPE_PWSTR;
		name = const_cast<char*>(this->m_ProductLibDir.c_str());
		size = mbstowcs(0, name, 0) + 1;
		wsMsgBuf = new wchar_t[size];
		memset(wsMsgBuf, 0, size * sizeof(wchar_t));
		size = mbstowcs(wsMsgBuf, name, size);
		pvarPropVal->wstrLen = size;
		m_iMemory->AllocMemory((void**)&pvarPropVal->pwstrVal, size * sizeof(WCHAR_T));
		::convToShortWchar(&pvarPropVal->pwstrVal, wsMsgBuf, size);
		break;
	default:
		return false;
	}

	return true;
}
//---------------------------------------------------------------------------//
bool JVMLauncher::SetPropVal(const long lPropNum, tVariant *varPropVal)
{
	switch (lPropNum)
	{
	case ePropIsEnabled:
		if (TV_VT(varPropVal) != VTYPE_BOOL)
			return false;
		m_boolEnabled = TV_BOOL(varPropVal);
		break;
	case ePropJavaHome:
		if (m_boolEnabled) {
			pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"JVM already running", 7);
			return false;
		}
		m_JavaHome = getStdStringFrom1C(varPropVal);
		break;
	case ePropLibraryDir:
		if (m_boolEnabled) {
			pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"JVM already running", 7);
			return false;
		}
		m_ProductLibDir = getStdStringFrom1C(varPropVal);
		break;
	default:
		return false;
	}
	return true;
}
//---------------------------------------------------------------------------//
bool JVMLauncher::IsPropReadable(const long lPropNum)
{
	switch (lPropNum)
	{
	case ePropIsEnabled:
	case ePropJavaHome:
	case ePropLibraryDir:
		return true;
	default:
		return false;
	}

	return false;
}
//---------------------------------------------------------------------------//
bool JVMLauncher::IsPropWritable(const long lPropNum)
{
	switch (lPropNum)
	{
	case ePropJavaHome:
	case ePropLibraryDir:
		return true;
	case ePropIsEnabled:
		return false;
	default:
		return false;
	}

	return false;
}
//---------------------------------------------------------------------------//
long JVMLauncher::GetNMethods()
{
	return eMethLast;
}
//---------------------------------------------------------------------------//
long JVMLauncher::FindMethod(const WCHAR_T* wsMethodName)
{
	long plMethodNum = -1;
	wchar_t* name = 0;

	::convFromShortWchar(&name, wsMethodName);

	plMethodNum = findName(g_MethodNames, name, eMethLast);

	if (plMethodNum == -1)
		plMethodNum = findName(g_MethodNamesRu, name, eMethLast);

	delete[] name;

	return plMethodNum;
}
//---------------------------------------------------------------------------//
const WCHAR_T* JVMLauncher::GetMethodName(const long lMethodNum, const long lMethodAlias)
{
	if (lMethodNum >= eMethLast)
		return NULL;

	wchar_t *wsCurrentName = NULL;
	WCHAR_T *wsMethodName = NULL;
	size_t iActualSize = 0;

	switch (lMethodAlias)
	{
	case 0: // First language
		wsCurrentName = g_MethodNames[lMethodNum];
		break;
	case 1: // Second language
		wsCurrentName = g_MethodNamesRu[lMethodNum];
		break;
	default:
		return 0;
	}

	iActualSize = wcslen(wsCurrentName) + 1;

	if (m_iMemory && wsCurrentName)
	{
		if (m_iMemory->AllocMemory((void**)&wsMethodName, iActualSize * sizeof(WCHAR_T)))
			::convToShortWchar(&wsMethodName, wsCurrentName, iActualSize);
	}

	return wsMethodName;
}
//---------------------------------------------------------------------------//
long JVMLauncher::GetNParams(const long lMethodNum)
{
	switch (lMethodNum)
	{
	case eMethAddJar:
	case eCallAsFuncB:
	case eCallAsProcedure:
		return 1;
	case eCallAsProcedureP:
	case eCallAsFuncBP:
	case eCallAsFunc:
		return 2;
	case eCallAsProcedurePP:
	case eCallAsFuncBPP:
	case eCallAsFuncP:
		return 3;
	case eCallAsFuncPP:
		return 4;
	default:
		return 0;
	}

	return 0;
}

//---------------------------------------------------------------------------//
bool JVMLauncher::HasRetVal(const long lMethodNum)
{
	switch (lMethodNum)
	{
	case eCallAsFuncB:
	case eCallAsFuncBP:
	case eCallAsFuncBPP:
	case eCallAsFunc:
	case eCallAsFuncP:
	case eCallAsFuncPP:
		return true;
	default:
		return false;
	}

	return false;
}


//---------------------------------------------------------------------------//
bool JVMLauncher::CallAsProc(const long lMethodNum,
	tVariant* paParams, const long lSizeArray)
{

	switch (lMethodNum)
	{
	case eCallAsProcedure:
	case eCallAsProcedureP:
	case eCallAsProcedurePP:
		return CallAsFunc(lMethodNum, paParams, paParams, lSizeArray);
	case eMethDisable:
		m_boolEnabled = false;
		break;
	case eMethAddJar:
	{
		if (!lSizeArray)
			return false;
		this->addJar(getStdStringFrom1C(paParams).c_str());
	}
	break;
	default:
		return false;
	}

	return true;
}
//---------------------------------------------------------------------------//
bool JVMLauncher::CallAsFunc(const long lMethodNum,
	tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{
	if (!lSizeArray || !paParams)
		return false;

	bool resultOk = true;
	std::string className = getStdStringFrom1C(paParams);

	if (className.empty()) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"Send empty class name", 10);
		return false;
	}
	if (!validateCall()) {
		return false;
	}

	int lastIndexOfParam = lSizeArray;
	TYPEVAR rt;

	switch (lMethodNum)
	{
	case eCallAsProcedure:
	case eCallAsProcedureP:
	case eCallAsProcedurePP:
		rt = VTYPE_EMPTY;
		break;
	case eCallAsFuncB:
	case eCallAsFuncBP:
	case eCallAsFuncBPP:
		rt = VTYPE_BLOB;
		break;
	case eCallAsFunc:
	case eCallAsFuncP:
	case eCallAsFuncPP:
	{
		lastIndexOfParam = lSizeArray - 1;
		rt = TV_VT(&paParams[lastIndexOfParam]);
	}
	break;
	}

	jint res = m_RunningJVMInstance->GetEnv((void**)&m_JVMEnv, NULL);
	res = m_RunningJVMInstance->AttachCurrentThread((void**)&m_JVMEnv, NULL);

	if (res != JNI_OK) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"Could not attach to the JVM", 4);
		return false;
	}

	jclass findedClass = findClassForCall(className);

	if (findedClass == nullptr) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"Cannot find class", 10);
		resultOk = false;
		goto detach;
	}

	jmethodID methodID;

	if (!(methodID = HasMethodInCache(findedClass))) {
		std::string signature = getSignature(this->m_JVMEnv, paParams, 1, lastIndexOfParam);
		switch (rt) {
		case VTYPE_PSTR: signature.append("Ljava/lang/String;"); break;
		case VTYPE_PWSTR: signature.append("Ljava/lang/String;"); break;
		case VTYPE_BLOB: signature.append("[B"); break;
		case VTYPE_I4: signature.append("I"); break;
		case VTYPE_BOOL: signature.append("Z"); break;
		case VTYPE_R4: signature.append("F"); break;
		case VTYPE_R8: signature.append("D"); break;
		case VTYPE_EMPTY: signature.append("V"); break;
		}
		methodID = JNI_getStaticMethodID(findedClass, "mainInt", signature.c_str(), resultOk);
		if (!resultOk) {
			goto detach;
		}
		m_cachedMethod.insert(std::pair<jclass, jmethodID>(findedClass, methodID));
	}
	else {
		methodID = findMethodForCall(findedClass, resultOk);
	}

	if (!resultOk) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, L"JVMLauncher", L"Cannot find method", 10);
		goto detach;
	}

	jvalue *values = getParams(this->m_JVMEnv, paParams, 1, lastIndexOfParam);
	jboolean resultBoolean = false;
	jstring resultString = nullptr;
	jint resultInt = 0;
	jfloat resultFloat = 0.0f;
	jdouble resultDouble = 0.0;
	jobject resultByteArray = nullptr;
	switch (rt)
	{
	case VTYPE_BOOL:
		resultBoolean = JNI_callStaticBooleanMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_PWSTR:
	case VTYPE_PSTR:
		resultString = (jstring)JNI_callStaticObjectMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_I4:
		resultInt = JNI_callStaticIntMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_R4:
		resultFloat = JNI_callStaticFloatMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_R8:
		resultDouble = JNI_callStaticDoubleMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_BLOB:
		resultByteArray = JNI_callStaticObjectMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_EMPTY:
		this->JNI_callStaticVoidMethodA(findedClass, methodID, values, resultOk);
		break;
	}

	delete[] values;

	if (!resultOk) {
		goto detach;
	}

	if (rt != VTYPE_EMPTY) {
		TV_VT(pvarRetValue) = rt;
	}

	switch (rt)
	{
	case VTYPE_BOOL:
		TV_BOOL(pvarRetValue) = resultBoolean;
		break;
	case VTYPE_PWSTR:
	case VTYPE_PSTR:
	{
		auto wstring = JStringToWString(m_JVMEnv, resultString);
		auto length = wstring.length();
		if (length && m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, length * sizeof(WCHAR_T)))
		{
			memcpy(pvarRetValue->pwstrVal, wstring.c_str(), length * sizeof(WCHAR_T));
			pvarRetValue->strLen = length;
		}
		m_JVMEnv->DeleteLocalRef(resultString);
		TV_VT(pvarRetValue) = VTYPE_PWSTR;
	}
	break;
	case VTYPE_I4:
		TV_INT(pvarRetValue) = resultInt;
		break;
	case VTYPE_R4:
		TV_R4(pvarRetValue) = resultFloat;
		break;
	case VTYPE_R8:
		TV_R8(pvarRetValue) = resultDouble;
		break;
	case VTYPE_BLOB:
	{
		jbyteArray arr = reinterpret_cast<jbyteArray>(resultByteArray);
		auto length = m_JVMEnv->GetArrayLength(arr);
		void *elements = m_JVMEnv->GetPrimitiveArrayCritical(arr, NULL);
		if (length && m_iMemory->AllocMemory((void**)&pvarRetValue->pstrVal, length))
		{
			memcpy(pvarRetValue->pstrVal, elements, length);
			pvarRetValue->strLen = length;
		}
		m_JVMEnv->ReleasePrimitiveArrayCritical(arr, elements, 0);
		m_JVMEnv->DeleteLocalRef(resultByteArray);
	}
	break;
	}

detach:
	m_RunningJVMInstance->DetachCurrentThread();

	return resultOk;
}


/////////////////////////////////////////////////////////////////////////////
// LocaleBase
//---------------------------------------------------------------------------//
bool JVMLauncher::setMemManager(void* memoryManager)
{
	m_iMemory = static_cast<IMemoryManager *>(memoryManager);
	return m_iMemory != nullptr;
}
//---------------------------------------------------------------------------//
void JVMLauncher::addError(uint32_t wcode, const wchar_t* source,
	const wchar_t* descriptor, long code)
{
	if (pAsyncEvent)
	{
		WCHAR_T *err = 0;
		WCHAR_T *descr = 0;

		::convToShortWchar(&err, source);
		::convToShortWchar(&descr, descriptor);

		pAsyncEvent->AddError(wcode, err, descr, code);
		delete[] err;
		delete[] descr;
	}
}
//---------------------------------------------------------------------------//
long JVMLauncher::findName(wchar_t* names[], const wchar_t* name,
	const uint32_t size) const
{
	long ret = -1;
	for (uint32_t i = 0; i < size; i++)
	{
		if (!wcscmp(names[i], name))
		{
			ret = i;
			break;
		}
	}
	return ret;
}

bool JVMLauncher::Init(void* connection)
{
	gAsyncEvent = static_cast<IAddInDefBase *>(connection);
	pAsyncEvent = static_cast<IAddInDefBase *>(connection);
	pAsyncEvent->SetEventBufferDepth(1000);
	return pAsyncEvent != nullptr;
}
