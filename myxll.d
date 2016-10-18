/**
	Code from generic.h and generic.d
	Ported to the D Programming Language by Laeeth Isharc (2015)
	This is the minimum amount of code to get a D function into Excel
*/

import core.sys.windows.windows;
import xlcall;
import framework;
import std.format;
import xlld.wrap;

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

// Global Variables
__gshared HANDLE g_hInst = null;

extern(Windows)
double FuncMulByTwo(double n) {
    return n * 2;
}

extern(Windows) LPXLOPER12 FuncFib (LPXLOPER12 n)
{
	static XLOPER12 xResult;
	XLOPER12 xlt;
	int val, max, error = -1;
	int[2] fib = [1,1];
	switch (n.xltype)
	{
	case xltypeNum:
		max = cast(int)n.val.num;
		if (max < 0)
			error = xlerrValue;
		for (val = 3; val <= max; val++)
		{
			fib[val%2] += fib[(val+1)%2];
		}
		xResult.xltype = xltypeNum;
		xResult.val.num = fib[(val+1)%2];
		break;
	case xltypeSRef:
		error = Excel12f(xlCoerce, &xlt, [n, TempInt12(xltypeInt)]);
		if (!error)
		{
			error = -1;
			max = xlt.val.w;
			if (max < 0)
				error = xlerrValue;
			for (val = 3; val <= max; val++)
			{
				fib[val%2] += fib[(val+1)%2];
			}
			xResult.xltype = xltypeNum;
			xResult.val.num = fib[(val+1)%2];
		}
		Excel12f(xlFree, cast(LPXLOPER12)0, [&xlt]);
		break;
	default:
		error = xlerrValue;
		break;
	}

	if ( error != - 1 )
	{
		xResult.xltype = xltypeErr;
		xResult.val.err = error;
	}

	//Word of caution - returning static XLOPERs/XLOPER12s is not thread safe
	//for UDFs declared as thread safe, use alternate memory allocation mechanisms
    return cast(LPXLOPER12) &xResult;
}

__gshared wstring[][] g_rgWorksheetFuncs =
[
	[ "FuncFib"w,
		"UU"w,
		"FuncFib"w,
		"Compute to..."w,
		"1"w,
		"MyXLL"w,
		""w,
		""w,
		"Number to compute to"w
		"Computes the nth fibonacci number"w,
	],
	[ "FuncMulByTwo"w,
		"BB"w,
		"FuncMulByTwo"w,
		"Multiply By two"w,
		"1"w,
		"MyXLL"w,
		""w,
		""w,
		"Number to multiply"w
		"Multiplies by 2"w,
	],
];


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
	import core.runtime:rt_init;
	rt_init();
	static XLOPER12 xDLL; 	   // name of this DLL //

	/**
	   In the following block of code the name of the XLL is obtained by
	   calling xlGetName. This name is used as the first argument to the
	   REGISTER function to specify the name of the XLL. Next, the XLL loops
	   through the g_rgWorksheetFuncs[] table, and the g_rgCommandFuncs[]
	   tableregistering each function in the table using xlfRegister.
	   Functions must be registered before you can add a menu item.
	*/

	Excel12f(xlGetName, &xDLL, []);

	foreach(row;g_rgWorksheetFuncs)
		Excel12f(xlfRegister, cast(LPXLOPER12)0, [cast(LPXLOPER12) &xDLL] ~ TempStr12(row[]));

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
