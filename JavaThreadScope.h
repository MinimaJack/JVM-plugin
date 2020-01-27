
class JavaThreadScope {
public:
	JavaThreadScope(JVMLauncher* launcher) : attached(false), launcher(launcher){
		jint res = launcher->m_RunningJVMInstance->GetEnv((void**)&(launcher->m_JVMEnv), NULL);
		res = launcher->m_RunningJVMInstance->AttachCurrentThread((void**)&(launcher->m_JVMEnv), NULL);

		if (res != JNI_OK) {
			attached = false;
		}
		else {
			attached = true;
		}

	}

	JavaThreadScope(const JavaThreadScope&) = delete;
	JavaThreadScope& operator=(const JavaThreadScope&) = delete;

	virtual ~JavaThreadScope() {
		if (attached) {
			launcher->m_RunningJVMInstance->DetachCurrentThread();
			attached = false;
		}
	}

	JNIEnv *GetEnv() const { return launcher->m_JVMEnv; }
	bool isAttached() const { return attached; }

private:
	bool attached;
	JVMLauncher* launcher;

};