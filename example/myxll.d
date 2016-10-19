/**
	Code from generic.h and generic.d
	Ported to the D Programming Language by Laeeth Isharc (2015)
	This is the minimum amount of code to get a D function into Excel
*/
module myxll;

import xlld;


extern(C) export double FuncMulByTwo(double n) {
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

// for what these mean see:
// https://msdn.microsoft.com/en-us/library/office/bb687900.aspx
__gshared wstring[][] g_rgWorksheetFuncs =
[
    [ "FuncFib"w, //Procedure
		"UU"w, //TypeText
		"FuncFib"w, //FunctionText
		"Compute to..."w, //ArgumentText
		"1"w, //MacroType
		"MyXLL"w, //Category
		""w, //ShortcutText
		""w, //HelpTopic
		"Number to compute to"w // FunctionHelp
		"Computes the nth fibonacci number"w, //ArgumentHelp1
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
