/**
	Code from generic.h and generic.d
	Ported to the D Programming Language by Laeeth Isharc (2015)
	This module provides the ceremony that must be done for every
	XLL. As long as `myxll` provides the global g_rgWorksheetFuncs
	with the functions to register, things should work.
*/

import xlld;
import core.sys.windows.windows;

__gshared HANDLE g_hInst = null;

// this function must be define in a module compiled with
// the current module
// It's extern(C) so that it can be defined in any module
extern(C) WorksheetFunction[] getWorksheetFunctions() @safe pure nothrow;


extern(Windows)
{
	pragma(lib, "gdi32");
	pragma(lib, "kernel32");
	pragma(lib, "user32");
	pragma(lib, "gdi32");
	pragma(lib, "winspool");
	pragma(lib, "comdlg32");
	pragma(lib, "advapi32");
	pragma(lib, "shell32");
	pragma(lib, "ole32");
	pragma(lib, "oleaut32");
	pragma(lib, "uuid");
	pragma(lib, "odbc32");
	pragma(lib, "xlcall32d");
}

extern(Windows) BOOL DllMain( HANDLE hDLL, DWORD dwReason, LPVOID lpReserved )
{
	import core.runtime;
	import std.c.windows.windows;
	import core.sys.windows.dll;
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH:
		Runtime.initialize();
		// The instance handle passed into DllMain is saved
		// in the global variable g_hInst for later use.
		g_hInst = hDLL;
		dll_process_attach( hDLL, true );
		break;
	case DLL_PROCESS_DETACH:
		Runtime.terminate();
	    dll_process_detach( hDLL, true );
		break;
	case DLL_THREAD_ATTACH:
	    dll_thread_attach( true, true );
		break;
	case DLL_THREAD_DETACH:
		dll_thread_detach( true, true );
		break;
	default:
		break;
	}
	return true;
}


extern(Windows) int xlAutoOpen()
{
	import std.conv;
        import std.algorithm: map;
	import core.runtime:rt_init;

	rt_init(); // move to DllOpen?

        // get name of this XLL, needed to pass to xlfRegister
	static XLOPER12 dllName;
	Excel12f(xlGetName, &dllName, []);

	foreach(row; getWorksheetFunctions.map!(a => a.toStringArray))
		Excel12f(xlfRegister, cast(LPXLOPER12)0, [cast(LPXLOPER12) &dllName] ~ TempStr12(row[]));

	return 1;
}

extern(Windows) LPXLOPER12 xlAddInManagerInfo12(LPXLOPER12 xAction)
{
	static XLOPER12 xInfo, xIntAction;

	//
	// This code coerces the passed-in value to an integer. This is how the
	// code determines what is being requested. If it receives a 1,
	// it returns a string representing the long name. If it receives
	// anything else, it returns a #VALUE! error.
	//

	Excel12f(xlCoerce, &xIntAction,[xAction, TempInt12(xltypeInt)]);

	if (xIntAction.val.w == 1)
	{
		xInfo.xltype = xltypeStr;
		xInfo.val.str = TempStr12("My XLL"w).val.str;
	}
	else
	{
		xInfo.xltype = xltypeErr;
		xInfo.val.err = xlerrValue;
	}

	//Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms
	return cast(LPXLOPER12) &xInfo;
}
