/**
	Code from generic.h and generic.d
	Ported to the D Programming Language by Laeeth Isharc (2015)

	what works: everything except fDance doesn't change active cell, and fShowDialog

	-----
	File:				GENERIC.H

   Purpose:			Header file for Generic.c

   Platform:    Microsoft Windows

   Updated by Microsoft Product Support Services, Windows Developer Support.
   From the Microsoft Excel Developer's Kit, Version 14
   Copyright (c) 1996-2010 Microsoft Corporation. All rights reserved.
 */
/**
   File:        GENERIC.C

   Purpose:     Template for creating XLLs for Microsoft Excel.

                This file contains sample code you can use as
                a template for writing your own Microsoft Excel XLLs.
                An XLL is a DLL that stands alone, that is, you
                can open it by choosing the Open command from the
                File menu. This code demonstrates many of the features
                of the Microsoft Excel C API.

                When you open GENERIC.XLL, it
                creates a new Generic menu with four
                commands:

                    Dialog...          displays a Microsoft Excel dialog box
                    Dance              moves the selection around
                                       until you press ESC
                    Native Dialog...   displays a Windows dialog box
                    Exit               Closes GENERIC.XLL and
                                       removes the menu

                GENERIC.XLL also provides three functions,
                Func1, FuncSum and FuncFib, which can be used whenever
                GENERIC.XLL is open.

                GENERIC.XLL can also be added with the
                Add-in Manager.

                This file uses the framework library
                (FRMWRK32.LIB).

   Platform:    Microsoft Windows

   Functions:
                DllMain
                xlAutoOpen
                xlAutoClose
                lpstricmp
                xlAutoRegister12
                xlAutoAdd
                xlAutoRemove
                xlAddInManagerInfo12
                DIALOGMsgProc
                ExcelCursorProc
                HookExcelWindow
                UnhookExcelWindow
                fShowDialog
                GetHwnd
                Func1
                FuncSum
                fDance
                fDialog
                fExit
                FuncFib

*/
import win32.winuser:PostMessage,CallWindowProc,GetWindowLongPtr,SetWindowLongPtr,DialogBox;
//import std.c.windows.windows;
import core.sys.windows.windows;
import xlcall;
import framework;
import core.stdc.wchar_ : wcslen;
import core.stdc.wctype:towlower;
import std.format;
import xlld.wrap;

enum GWLP_WNDPROC=-4;
enum MAXWORD = 0xFFFF;
debug=0;
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
	//pragma(lib, "odbccp32");
	//pragma(lib, "msvcrt32");

	enum GMEM_MOVEABLE = 0x02;
	void* GlobalAlloc(uint, size_t);
   void* GlobalLock(void*);
   bool GlobalUnlock(void*);
	void cwCenter(HWND, int);
	//INT_PTR /*CALLBACK*/ DIALOGMsgProc(HWND hWndDlg, UINT message, WPARAM wParam, LPARAM lParam);
}

//   identifier for controls

enum FREE_SPACE                  =104;
enum EDIT                        =101;
enum TEST_EDIT                   =106;


/**
   Later, the instance handle is required to create dialog boxes.
   g_hInst holds the instance handle passed in by DllMain so that it is
   available for later use. hWndMain is used in several routines to
   store Microsoft Excel's hWnd. This is used to attach dialog boxes as
   children of Microsoft Excel's main window. A buffer is used to store
   the free space that DIALOGMsgProc will put into the dialog box.
 */

// Global Variables

__gshared HWND g_hWndMain = null;
__gshared HANDLE g_hInst = null;
wchar[20] g_szBuffer = ""w;


/**
   Syntax of the Register Command:
        REGISTER(module_text, procedure, type_text, function_text,
                 argument_text, macro_type, category, shortcut_text,
                 help_topic, function_help, argument_help1, argument_help2,...)


   g_rgWorksheetFuncs will use only the first 11 arguments of
   the Register function.

   This is a table of all the worksheet functions exported by this module.
   These functions are all registered (in xlAutoOpen) when you
   open the XLL. Before every string, leave a space for the
   byte count. The format of this table is the same as
   arguments two through eleven of the REGISTER function.
   g_rgWorksheetFuncsRows define the number of rows in the table. The
   g_rgWorksheetFuncsCols represents the number of columns in the table.
*/
enum g_rgWorksheetFuncsRows =5;
enum g_rgWorksheetFuncsCols =10;

__gshared wstring[g_rgWorksheetFuncsCols][g_rgWorksheetFuncsRows] g_rgWorksheetFuncs =
[
	[ "Func1"w,                                     // Procedure
		"UU"w,                                  // type_text
		"Func1"w,                               // function_text
		"Arg"w,                                 // argument_text
		"1"w,                                   // macro_type
		"Generic Add-In"w,                      // category
		""w,                                    // shortcut_text
		""w,                                    // help_topic
		"Always returns the string 'Func1'"w,   // function_help
		"Argument ignored"w                     // argument_help1
	],
	[ "FuncSum"w,
		"UUUUUUUUUUUUUUUUUUUUUUUUUUUUUU"w, // up to 255 args in Excel 2007 and later,
										   // upto 29 args in Excel 2003 and earlier versions
		"FuncSum"w,
		"number1,number2,..."w,
		"1"w,
		"Generic Add-In"w,
		""w,
		""w,
		"Adds the arguments"w,
		"Number1,number2,... are 1 to 29 arguments for which you want to sum."w
	],
	[ "lastErrorMessage"w,
		"Q"w, // up to 255 args in Excel 2007 and later,
										   // upto 29 args in Excel 2003 and earlier versions
		"lastErrorMessage"w,
		""w,
		"1"w,
		"Generic Add-In"w,
		""w,
		""w,
		"Return last D error message"w,
		""w,
	],
	[ "WrapSquare3"w,
		"QQQQQQQQQQQQQQQQQQQQQQQQQQQQQQ"w, // up to 255 args in Excel 2007 and later,
										   // upto 29 args in Excel 2003 and earlier versions
		"WrapSquare3"w,
		"number1,number2,..."w,
		"1"w,
		"Generic Add-In"w,
		""w,
		""w,
		"Sum of squares of the arguments"w,
		"Number1,number2,... are 1 to 29 arguments for which you want to sum."w
	],
	[ "FuncFib"w,
		"UU"w,
		"FuncFib"w,
		"Compute to..."w,
		"1"w,
		"Generic Add-In"w,
		""w,
		""w,
		"Number to compute to"w
		"Computes the nth fibonacci number"w,
	],
];

/**
   g_rgCommandFuncs

   This is a table of all the command functions exported by this module.
   These functions are all registered (in xlAutoOpen) when you
   open the XLL. Before every string, leave a space for the
   byte count. The format of this table is the same as
   arguments two through eight of the REGISTER function.
   g_rgFuncsRows define the number of rows in the table. The
   g_rgCommandFuncsCols represents the number of columns in the table.
*/
enum g_rgCommandFuncsRows = 1;
enum g_rgCommandFuncsCols = 7;

__gshared wstring g_rgCommandFuncs[g_rgCommandFuncsRows][g_rgCommandFuncsCols] =
[
	[ "fDialog"w,                   // Procedure
		"A"w,                   // type_text
		"fDialog"w,             // function_text
		""w,                    // argument_text
		"2"w,                   // macro_type
		"Generic Add-In"w,      // category
		"l"w                    // shortcut_text
	],
];


/**
   DllMain()

   Purpose:

        Windows calls DllMain, for both initialization and termination.
  		It also makes calls on both a per-process and per-thread basis,
  		so several initialization calls can be made if a process is multithreaded.

        This function is called when the DLL is first loaded, with a dwReason
        of DLL_PROCESS_ATTACH.

   Parameters:

        HANDLE hDLL         Module handle.
        DWORD dwReason,     Reason for call
        LPVOID lpReserved   Reserved

   Returns:
        The function returns true (1) to indicate success. If, during
        per-process initialization, the function returns zero,
        the system cancels the process.

   Comments:

   History:  Date       Author        Reason
*/

extern(Windows) BOOL /*APIENTRY*/ DllMain( HANDLE hDLL, DWORD dwReason, LPVOID lpReserved )
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

/**
   xlAutoOpen()

   Purpose:
        Microsoft Excel call this function when the DLL is loaded.

        Microsoft Excel uses xlAutoOpen to load XLL files.
        When you open an XLL file, the only action
        Microsoft Excel takes is to call the xlAutoOpen function.

        More specifically, xlAutoOpen is called:

         - when you open this XLL file from the File menu,
         - when this XLL is in the XLSTART directory, and is
           automatically opened when Microsoft Excel starts,
         - when Microsoft Excel opens this XLL for any other reason, or
         - when a macro calls REGISTER(), with only one argument, which is the
           name of this XLL.

        xlAutoOpen is also called by the Add-in Manager when you add this XLL
        as an add-in. The Add-in Manager first calls xlAutoAdd, then calls
        REGISTER("EXAMPLE.XLL"), which in turn calls xlAutoOpen.

        xlAutoOpen should:

         - register all the functions you want to make available while this
           XLL is open,

         - add any menus or menu items that this XLL supports,

         - perform any other initialization you need, and

         - return 1 if successful, or return 0 if your XLL cannot be opened.

   Parameters:

   Returns:

        int         1 on success, 0 on failure

   Comments:

   History:  Date       Author        Reason
*/

extern(Windows) int /*WINAPI*/ xlAutoOpen()
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


/**
   xlAddInManagerInfo12()

   Purpose:

        This function is called by the Add-in Manager to find the long name
        of the add-in. If xAction = 1, this function should return a string
        containing the long name of this XLL, which the Add-in Manager will use
        to describe this XLL. If xAction = 2 or 3, this function should return
        #VALUE!.

   Parameters:

        LPXLOPER12 xAction  What information you want. One of:
                              1 = the long name of the
                                  add-in
                              2 = reserved
                              3 = reserved

   Returns:

        LPXLOPER12          The long name or #VALUE!.

   Comments:

   History:  Date       Author        Reason
*/

extern(Windows) LPXLOPER12 /*WINAPI*/ xlAddInManagerInfo12(LPXLOPER12 xAction)
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


/**
   FuncFib()

   Purpose:

        A sample function that computes the nth Fibonacci number.
        Features a call to several wrapper functions.

   Parameters:

        LPXLOPER12 n    int to compute to

   Returns:

        LPXLOPER12      nth Fibonacci number

   Comments:

   History:  Date       Author        Reason
*/

extern(Windows) LPXLOPER12 /*WINAPI*/ FuncFib (LPXLOPER12 n)
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
