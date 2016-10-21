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

extern(Windows) double FuncAddEverything(LPXLOPER12 arg) {
    import std.algorithm: fold;
    return getValues(arg).fold!((a, b) => a + b)(0.0);
}

// calling this recursively fails miserably, even if it's
// also extern(Windows). I don't know why
private double[] getValues(LPXLOPER12 arg) {
    switch(arg.xltype) {

    case xltypeNum:
        return [arg.val.num];

    case xltypeSRef:
        XLOPER12 array;
        Excel12f(xlCoerce, &array, [arg]);
        scope(exit) Excel12f(xlFree, null, [&array]);

        double[] ret;
        const rows = array.val.array.rows;
        const columns = array.val.array.columns;
        auto values = array.val.array.lparray[0 .. (rows * columns)];

        foreach(const row; 0 .. rows) {
            foreach(const col; 0 .. columns) {
                XLOPER12 val;
                Excel12f(xlCoerce, &val, [&values[row * columns + col]]);
                ret ~= val.val.num;
                Excel12f(xlFree, null, [&val]);
            }
        }
        return ret;

    default:
        return [-1];
    }
}

// extern(Windows) means it has to be explicitly added
// to the .def file
// Because of that and the only double -> double reflection functionality,
// it won't appear in Excel
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
