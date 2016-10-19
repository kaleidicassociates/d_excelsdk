/**
	Code from generic.h and generic.d
	Ported to the D Programming Language by Laeeth Isharc (2015)
	This is an example of how to write D functions that can
	be called from Excel.
	The getWorksheetFunctions function returns the necessary
	binding information
*/
module example.myxll;

import xlld;

mixin(implGetWorksheetFunctionsString!__MODULE__);

// extern(C) export means it doesn't have to be explicitly
// added to the .def file
extern(C) export double FuncMulByTwo(double n) {
    return n * 2;
}

// extern(C) export means it doesn't have to be explicitly
// added to the .def file
extern(C) export LPXLOPER12 FuncFib (LPXLOPER12 n)
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
