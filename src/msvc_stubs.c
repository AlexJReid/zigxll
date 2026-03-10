// Stubs for MSVC CRT initialization symbols referenced by msvcrt.lib.
// When building an XLL (DLL loaded into Excel), the host process already
// has the CRT initialized, so these can be safely stubbed out.

int __vcrt_initialize(void) { return 1; }
int __vcrt_uninitialize(int terminating) { (void)terminating; return 1; }
void __vcrt_uninitialize_critical(void) {}
int __vcrt_thread_attach(void) { return 1; }
int __vcrt_thread_detach(void) { return 1; }

int __acrt_initialize(void) { return 1; }
int __acrt_uninitialize(int terminating) { (void)terminating; return 1; }
void __acrt_uninitialize_critical(void) {}
int __acrt_thread_attach(void) { return 1; }
int __acrt_thread_detach(void) { return 1; }

int _is_c_termination_complete(void) { return 0; }
