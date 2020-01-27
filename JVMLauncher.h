#ifndef __JVMLAUNCHER_H__
#define __JVMLAUNCHER_H__

#include "ComponentBase.h"
#include "AddInDefBase.h"
#include "IMemoryManager.h"
#include <string>
#include <vector>
#include <map>
#include <include/jni.h>
#include "Utils.h"
#include <optional>

#define BEGIN_JAVA { JNIEnv* env = this->m_JVMEnv; this->m_JVMEnv = nullptr;
#define END_JAVA this->m_JVMEnv = env; }

#define BEGIN_CALL \
	BEGIN_JAVA

#define END_CALL \
	hasError = this->endCall(env); }

#define JVM_LAUNCHER L"JVMLauncher"

typedef jint(JNICALL *CreateJavaVM)(JavaVM **pvm, void **penv, void *args);
typedef jint(JNICALL * GetCreatedJavaVMs)(JavaVM**, jsize, jsize*);
typedef std::pair<jclass, jmethodID> jclassMethodHolder;

class JVMLauncher : public IComponentBase {
public:

	enum Props
	{
		ePropIsEnabled = 0,
		ePropJavaHome,
		ePropLibraryDir,
		ePropLast      // Always last
	};

	enum Methods
	{
		eCallAsProcedure = 0,
		eCallAsProcedureP,
		eCallAsProcedurePP,
		eCallAsFuncB,
		eCallAsFuncBP,
		eCallAsFuncBPP,
		eCallAsFunc,
		eCallAsFuncP,
		eCallAsFuncPP,
		eCallAsFuncPPP,
		eCallAsFuncPPPP,
		eClassFunctionCall,
		eClassFunctionCallP,
		eClassFunctionCallPP,
		eClassProcedureCall,
		eClassProcedureCallP,
		eClassProcedureCallPP,
		eMethDisable,
		eMethAddJar,
		eMethLast      // Always last
	};

	JVMLauncher(void);
	~JVMLauncher() override {};

	// IInitDoneBase
	bool ADDIN_API Init(void*) override;
	bool ADDIN_API setMemManager(void* mem) override;
	long ADDIN_API GetInfo() override { return 2000; }
	void ADDIN_API Done() override {};
	// ILanguageExtenderBase
	bool ADDIN_API RegisterExtensionAs(WCHAR_T**) override;
	long ADDIN_API GetNProps() override;
	long ADDIN_API FindProp(const WCHAR_T* wsPropName) override;
	const WCHAR_T* ADDIN_API GetPropName(long lPropNum, long lPropAlias) override;
	bool ADDIN_API GetPropVal(const long lPropNum, tVariant* pvarPropVal) override;
	bool ADDIN_API SetPropVal(const long lPropNum, tVariant* varPropVal) override;
	bool ADDIN_API IsPropReadable(const long lPropNum) override;
	bool ADDIN_API IsPropWritable(const long lPropNum) override;
	long ADDIN_API GetNMethods() override;
	long ADDIN_API FindMethod(const WCHAR_T* wsMethodName) override;
	const WCHAR_T* ADDIN_API GetMethodName(const long lMethodNum,
		const long lMethodAlias) override;
	long ADDIN_API GetNParams(const long lMethodNum) override;
	bool ADDIN_API GetParamDefValue(const long lMethodNum, const long lParamNum,
		tVariant *pvarParamDefValue)  override {
		return false;
	};
	bool ADDIN_API HasRetVal(const long lMethodNum) override;
	bool ADDIN_API CallAsProc(const long lMethodNum,
		tVariant* paParams, const long lSizeArray) override;
	bool ADDIN_API CallAsFunc(const long lMethodNum,
		tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray) override;
	// LocaleBase
	void ADDIN_API SetLocale(const WCHAR_T* loc) override {};

private:

	// Attributes
	IAddInDefBase      *pAsyncEvent = nullptr;
	IMemoryManager     *m_iMemory;
	bool                m_boolEnabled = false;
	bool			    valid = true;
	HINSTANCE           m_hDllInstance = nullptr;

	std::string         m_JavaHome;
	std::string         m_ProductLibDir;
	std::string         m_JvmDllLocation;

	CreateJavaVM        m_JVMInstance = nullptr;
	GetCreatedJavaVMs   m_GetCreatedJavaVMs = nullptr;
	JNIEnv             *m_JVMEnv;
	JavaVM             *m_RunningJVMInstance;

	std::vector<std::string> m_listOfJars;
	std::vector<std::string> m_listOfClasses;
	std::map<std::string, jclass> m_cachedClasses;
	std::map<jclass, jmethodID> m_cachedMethod;
	std::map<std::string, jclassMethodHolder> m_cachedClassesMethod;

	void LaunchJVM();
	bool verify();
	jint JNI_destroyVM();
	bool validateCall();
	std::string buildSignature(JNIEnv * env, tVariant * paParams, int start, int end, TYPEVAR returnType);
	jclass findClassForCall(std::string className);
	jmethodID findMethodForCall(jclass findedclass, bool& hasError);
	jmethodID HasMethodInCache(jclass findedclass);
	std::optional<jclassMethodHolder> HasClassAndMethodInCache(std::string classFunctionName);
	long findName(wchar_t* names[], const wchar_t* name, const uint32_t size) const;
	void addError(uint32_t wcode, const wchar_t* source,
		const wchar_t* descriptor, long code);
	void addJar(const char* jarname);
	jclass JNI_findClass(const char* className);
	void JNI_callStaticVoidMethod(jclass clazz, jmethodID methodID, bool& hasError, ...);
	jobjectArray JNI_newObjectArray(jsize length, jclass elementClass, jobject initialElement, bool& hasError);
	void JNI_callStaticVoidMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError);
	jobject JNI_callStaticObjectMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError);
	jboolean JNI_callStaticBooleanMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError);
	jint JNI_callStaticIntMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError);
	jfloat JNI_callStaticFloatMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError);
	jdouble JNI_callStaticDoubleMethodA(jclass clazz, jmethodID methodID, jvalue* args, bool& hasError);
	
	void JNI_callStaticVoidMethodV(jclass clazz, jmethodID methodID, va_list args, bool& hasError);
	jmethodID JNI_getStaticMethodID(jclass clazz, const char* name, const char* sig, bool& hasError);

	bool endCall(JNIEnv* env);
protected:
};

#endif //__JVMLAUNCHER_H__
