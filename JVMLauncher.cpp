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
#include <algorithm>
#include <utility>
#include "JVMLauncher.h"

#define BASE_ERRNO     7

static jobject s_threadLock;

IAddInDefBase *gAsyncEvent = nullptr;

static wchar_t *g_PropNames[] = { L"IsEnabled", L"javaHome", L"libraryDir" };
static wchar_t *g_MethodNames[] = { L"LaunchInJVM", L"LaunchInJVMP", L"LaunchInJVMPP",
									L"CallFInJVMB", L"CallFInJVMBP", L"CallFInJVMBPP",
									L"CallFInJVM", L"CallFInJVMP", L"CallFInJVMPP", L"CallFInJVMPPP", L"CallFInJVMPPPP",
									L"ClassFunctionCall", L"ClassFunctionCallP", L"ClassFunctionCallPP",
									L"ClassProcedureCall", L"ClassProcedureCallP", L"ClassProcedureCallPP",
									L"Disable", L"AddJar" };

static void JNICALL Java_Runner_log(JNIEnv *env, jobject thisObj, jstring info) {
	if (gAsyncEvent == nullptr) {
		return;
	}

	WCHAR_T *message = 0;
	WCHAR_T *className = 0;
	jstring callingClassName;
	jobject classObject;
	jclass classOfClass;

	jclass cls = env->FindClass("java/lang/Class");
	if (env->IsInstanceOf(thisObj, cls)) {
		classObject = thisObj;
		classOfClass = env->GetObjectClass(thisObj);
	}
	else {
		jclass classClass = env->GetObjectClass(thisObj);

		jmethodID mid = env->GetMethodID(classClass, "getClass", "()Ljava/lang/Class;");
		jobject clsObj = env->CallObjectMethod(thisObj, mid);

		classObject = clsObj;
		classOfClass = env->GetObjectClass(clsObj);
	}

	jmethodID getNameMethod = env->GetMethodID(classOfClass, "getName", "()Ljava/lang/String;");
	callingClassName = (jstring)env->CallObjectMethod(classObject, getNameMethod);

	auto wsCallClassName = JStringToWString(env, callingClassName);
	::convToShortWchar(&className, wsCallClassName.c_str());
	if (env->IsSameObject(info, NULL)) {
		gAsyncEvent->ExternalEvent(JVM_LAUNCHER, className, L"NULL");
	}
	else {
		auto wsMessage = JStringToWString(env, info);
		::convToShortWchar(&message, wsMessage.c_str());
		gAsyncEvent->ExternalEvent(JVM_LAUNCHER, className, message);
	}

	delete[]message;
	delete[]className;

}

bool JVMLauncher::endCall(JNIEnv* env)
{
	m_JVMEnv = env;
	jobject exh = m_JVMEnv->ExceptionOccurred();
	if (exh) {
		jclass classClass = this->m_JVMEnv->GetObjectClass(exh);
		auto getMessageMethod = this->m_JVMEnv->GetMethodID(classClass, "getMessage", "()Ljava/lang/String;");
		auto info = (jstring)this->m_JVMEnv->CallObjectMethod(exh, getMessageMethod);
		auto mid = env->GetMethodID(classClass, "getClass", "()Ljava/lang/Class;");

		jobject clsObj = env->CallObjectMethod(exh, mid);
		jclass classOfClass = env->GetObjectClass(clsObj);
		mid = this->m_JVMEnv->GetMethodID(classOfClass, "getName", "()Ljava/lang/String;");
		auto exeptionName = (jstring)this->m_JVMEnv->CallObjectMethod(clsObj, mid);

		auto wstring = JStringToWString(env, exeptionName);

		if (!m_JVMEnv->IsSameObject(info, NULL)) {
			wstring = wstring.append(L":").append(JStringToWString(env, info));
		};
		WCHAR_T *err = 0;

		::convToShortWchar(&err, wstring.c_str());
		pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, err, 10);
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
		pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"JVM already launched", 6);
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
			pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"Set JAVA_HOME environment", 1);
			return;
		}
		m_JvmDllLocation = m_JavaHome + "/jre/bin/server/jvm.dll";
		m_hDllInstance = LoadLibraryA(m_JvmDllLocation.c_str());
		if (m_hDllInstance == nullptr) {
			m_JvmDllLocation = m_JavaHome + "/bin/server/jvm.dll";
			m_hDllInstance = LoadLibraryA(m_JvmDllLocation.c_str());
		}
	}
	if (m_hDllInstance == nullptr) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"Cannot find JDK", 1);
		return;
	}
	else {

		if (m_GetCreatedJavaVMs == nullptr) {
			m_GetCreatedJavaVMs = reinterpret_cast<GetCreatedJavaVMs>(GetProcAddress(m_hDllInstance, "JNI_GetCreatedJavaVMs"));
		}

		int n;
		jint retval = m_GetCreatedJavaVMs(&m_RunningJVMInstance, 1, (jsize*)&n);

		if (retval == JNI_OK)
		{

			if (n == 0)
			{
				if (m_JVMInstance == nullptr) {
					m_JVMInstance = reinterpret_cast<CreateJavaVM>(GetProcAddress(m_hDllInstance, "JNI_CreateJavaVM"));
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
					pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"Could not launch the JVM", 3);
				}
				else {
					m_boolEnabled = true;
					if (!verify()) {
						pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"Cannot verify jar", 8);
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
		pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"JVM not running", 5);
		return false;
	}
	if (!valid) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"Load not valid packages", 5);
		return false;
	}
	return true;
}

jclass JVMLauncher::findClassForCall(std::string className) {
	jclass findedClass = nullptr;
	std::replace(className.begin(), className.end(), '.', '/');
	auto val = m_cachedClasses.find(className);
	if (val == m_cachedClasses.end()) {
		jclass neededclass = this->m_JVMEnv->FindClass(className.c_str());
		if (!this->m_JVMEnv->ExceptionOccurred()) {
			findedClass = (jclass)this->m_JVMEnv->NewGlobalRef(neededclass);
			m_cachedClasses.insert(std::pair<std::string, jclass>(className, findedClass));
			this->m_JVMEnv->DeleteLocalRef(neededclass);
			JNINativeMethod methods[]{ { "log", "(Ljava/lang/String;)V", (void *)&Java_Runner_log } };
			if (m_JVMEnv->RegisterNatives(findedClass, methods, 1) < 0) {
				if (m_JVMEnv->ExceptionOccurred()) {
					this->m_JVMEnv->ExceptionClear();
				}
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

std::optional<jclassMethodHolder> JVMLauncher::HasClassAndMethodInCache(std::string className) {
	auto val = m_cachedClassesMethod.find(className);
	if (val != m_cachedClassesMethod.end()) {
		return val->second;
	}
	return std::nullopt;
}

/////////////////////////////////////////////////////////////////////////////
// ILanguageExtenderBase
//---------------------------------------------------------------------------//
bool JVMLauncher::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{
	wchar_t *wsExtension = JVM_LAUNCHER;
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
	wsCurrentName = g_PropNames[lPropNum];
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
			pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"JVM already running", 7);
			return false;
		}
		m_JavaHome = getStdStringFrom1C(varPropVal);
		break;
	case ePropLibraryDir:
		if (m_boolEnabled) {
			pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"JVM already running", 7);
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
	wsCurrentName = g_MethodNames[lMethodNum];

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
	case eClassProcedureCall:
		return 1;
	case eCallAsProcedureP:
	case eCallAsFuncBP:
	case eCallAsFunc:
	case eClassFunctionCall:
	case eClassProcedureCallP:
		return 2;
	case eCallAsProcedurePP:
	case eCallAsFuncBPP:
	case eCallAsFuncP:
	case eClassFunctionCallP:
	case eClassProcedureCallPP:
		return 3;
	case eCallAsFuncPP:
	case eClassFunctionCallPP:
		return 4;
	case eCallAsFuncPPP:
		return 5;
	case eCallAsFuncPPPP:
		return 6;
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
	case eCallAsFuncPPP:
	case eCallAsFuncPPPP:
	case eClassFunctionCall:
	case eClassFunctionCallP:
	case eClassFunctionCallPP:
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
	case eClassProcedureCall:
	case eClassProcedureCallP:
	case eClassProcedureCallPP:
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

std::string JVMLauncher::buildSignature(JNIEnv* env, tVariant* paParams, int start, int end, TYPEVAR returnType) {
	std::string signature = getSignature(env, paParams, 1, end);
	switch (returnType) {
	case VTYPE_PSTR: signature.append("Ljava/lang/String;"); break;
	case VTYPE_PWSTR: signature.append("Ljava/lang/String;"); break;
	case VTYPE_BLOB: signature.append("[B"); break;
	case VTYPE_I4: signature.append("I"); break;
	case VTYPE_BOOL: signature.append("Z"); break;
	case VTYPE_R4: signature.append("F"); break;
	case VTYPE_R8: signature.append("D"); break;
	case VTYPE_EMPTY: signature.append("V"); break;
	case VTYPE_DATE: signature.append("Ljava/util/Date;"); break;
	case VTYPE_TM: signature.append("Ljava/util/Date;"); break;
	}
	return signature;
}



//---------------------------------------------------------------------------//
bool JVMLauncher::CallAsFunc(const long lMethodNum,
	tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{
	if (!lSizeArray || !paParams)
		return false;

	bool resultOk = true;
	std::string classFunctionName = getStdStringFrom1C(paParams);

	if (classFunctionName.empty()) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"Send empty class name", 10);
		return false;
	}
	if (!validateCall()) {
		return false;
	}

	int lastIndexOfParam = lSizeArray;
	bool classContainFunctionName = false;
	TYPEVAR returnType;

	switch (lMethodNum)
	{
	case eCallAsProcedure:
	case eCallAsProcedureP:
	case eCallAsProcedurePP:
		returnType = VTYPE_EMPTY;
		break;
	case eCallAsFuncB:
	case eCallAsFuncBP:
	case eCallAsFuncBPP:
		returnType = VTYPE_BLOB;
		break;
	case eCallAsFunc:
	case eCallAsFuncP:
	case eCallAsFuncPP:
	case eCallAsFuncPPP:
	case eCallAsFuncPPPP:
	{
		lastIndexOfParam = lSizeArray - 1;
		returnType = TV_VT(&paParams[lastIndexOfParam]);
	}
	break;
	case eClassFunctionCall:
	case eClassFunctionCallP:
	case eClassFunctionCallPP:
	{
		classContainFunctionName = true;
		lastIndexOfParam = lSizeArray - 1;
		returnType = TV_VT(&paParams[lastIndexOfParam]);
	}
	break;
	case eClassProcedureCall:
	case eClassProcedureCallP:
	case eClassProcedureCallPP:
	{
		classContainFunctionName = true;
		lastIndexOfParam = lSizeArray;
		returnType = VTYPE_EMPTY;
	}
	break;

	}

	jint res = m_RunningJVMInstance->GetEnv((void**)&m_JVMEnv, NULL);
	res = m_RunningJVMInstance->AttachCurrentThread((void**)&m_JVMEnv, NULL);

	if (res != JNI_OK) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"Could not attach to the JVM", 4);
		return false;
	}
	jclass findedClass;
	jmethodID methodID;

	if (classContainFunctionName) {
		auto value = HasClassAndMethodInCache(classFunctionName);
		if (value != std::nullopt) {
			findedClass = value->first;
			methodID = value->second;
		}
		else {
			auto pos = classFunctionName.find_last_of(".");
			auto className = classFunctionName.substr(0, pos);
			findedClass = findClassForCall(className);
			if (findedClass != nullptr) {
				auto methodName = classFunctionName.substr(pos + 1);
				std::string signature = buildSignature(this->m_JVMEnv, paParams, 1, lastIndexOfParam, returnType);
				methodID = JNI_getStaticMethodID(findedClass, methodName.c_str(), signature.c_str(), resultOk);
				if (resultOk) {
					m_cachedClassesMethod.insert(std::make_pair(classFunctionName, std::make_pair(findedClass, methodID)));
				}
			}
		}
	}
	else {
		findedClass = findClassForCall(classFunctionName);
		if (findedClass != nullptr) {
			if (!(methodID = HasMethodInCache(findedClass))) {
				std::string signature = this->buildSignature(this->m_JVMEnv, paParams, 1, lastIndexOfParam, returnType);
				methodID = JNI_getStaticMethodID(findedClass, "mainInt", signature.c_str(), resultOk);
				if (resultOk) {
					m_cachedMethod.insert(std::pair<jclass, jmethodID>(findedClass, methodID));
				}
			}
			else {
				methodID = findMethodForCall(findedClass, resultOk);
			}
		}
	}

	if (findedClass == nullptr) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"Cannot find class", 10);
		resultOk = false;
		goto detach;
	}

	if (methodID == nullptr) {
		pAsyncEvent->AddError(ADDIN_E_FAIL, JVM_LAUNCHER, L"Cannot find method", 10);
		resultOk = false;
		goto detach;
	}

	jvalue *values = getParams(this->m_JVMEnv, paParams, 1, lastIndexOfParam);
	jvalue result;
	switch (returnType)
	{
	case VTYPE_BOOL:
		result.z = JNI_callStaticBooleanMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_PWSTR:
	case VTYPE_PSTR:
	case VTYPE_DATE:
	case VTYPE_TM:
		result.l = JNI_callStaticObjectMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_I4:
		result.i = JNI_callStaticIntMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_R4:
		result.f = JNI_callStaticFloatMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_R8:
		result.d = JNI_callStaticDoubleMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_BLOB:
		result.l = JNI_callStaticObjectMethodA(findedClass, methodID, values, resultOk);
		break;
	case VTYPE_EMPTY:
		this->JNI_callStaticVoidMethodA(findedClass, methodID, values, resultOk);
		break;
	}

	delete[] values;

	if (!resultOk) {
		goto detach;
	}

	if (returnType != VTYPE_EMPTY) {
		TV_VT(pvarRetValue) = returnType;
	}

	switch (returnType)
	{
	case VTYPE_BOOL:
		TV_BOOL(pvarRetValue) = result.z;
		break;
	case VTYPE_PWSTR:
	case VTYPE_PSTR:
	{
		if (!m_JVMEnv->IsSameObject(result.l, NULL)) {
			TV_VT(pvarRetValue) = VTYPE_PWSTR;
			auto wstring = JStringToWString(m_JVMEnv, static_cast<jstring>(result.l));
			auto length = wstring.length();
			if (length && m_iMemory->AllocMemory((void**)&pvarRetValue->pwstrVal, length * sizeof(WCHAR_T)))
			{
				memcpy(pvarRetValue->pwstrVal, wstring.c_str(), length * sizeof(WCHAR_T));
				pvarRetValue->strLen = length;
			}
			m_JVMEnv->DeleteLocalRef(result.l);
		}
		else {
			TV_VT(pvarRetValue) = VTYPE_NULL;
		}
	}
	break;
	case VTYPE_I4:
		TV_INT(pvarRetValue) = result.i;
		break;
	case VTYPE_R4:
		TV_R4(pvarRetValue) = result.f;
		break;
	case VTYPE_R8:
		TV_R8(pvarRetValue) = result.d;
		break;
	case VTYPE_BLOB:
	{
		if (!m_JVMEnv->IsSameObject(result.l, NULL)) {
			jbyteArray arr = static_cast<jbyteArray>(result.l);
			auto length = m_JVMEnv->GetArrayLength(arr);
			void *elements = m_JVMEnv->GetPrimitiveArrayCritical(arr, NULL);
			if (length && m_iMemory->AllocMemory((void**)&pvarRetValue->pstrVal, length))
			{
				memcpy(pvarRetValue->pstrVal, elements, length);
				pvarRetValue->strLen = length;
			}
			m_JVMEnv->ReleasePrimitiveArrayCritical(arr, elements, 0);
			m_JVMEnv->DeleteLocalRef(result.l);
		}
		else {
			TV_VT(pvarRetValue) = VTYPE_NULL;
		}
	}
	break;
	case VTYPE_DATE:
	case VTYPE_TM: {
		if (m_JVMEnv->IsSameObject(result.l, NULL)) {
			TV_VT(pvarRetValue) = VTYPE_NULL;
		}
		else {
			TV_VT(pvarRetValue) = VTYPE_TM;

			jclass date = m_JVMEnv->FindClass("java/util/Date");
			jmethodID getTime = m_JVMEnv->GetMethodID(date, "getTime", "()J");

			time_t rawtime = m_JVMEnv->CallLongMethod(result.l, getTime) / 1000;
			tm* ref = gmtime(&rawtime);
			(pvarRetValue)->tmVal = *ref;
			m_JVMEnv->DeleteLocalRef(result.l);
		}
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
