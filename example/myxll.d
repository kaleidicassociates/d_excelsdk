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
//mixin(wrapWorksheetFunctionsString!"xlld.test_d_funcs");

extern(Windows) LPXLOPER12 FuncAddEverything(LPXLOPER12 arg) {
    static import xlld.test_d_funcs;
    static XLOPER12 ret;
    static XLOPER12[100] opers;

    if(arg.xltype != xltypeSRef) {
        ret.xltype = xltypeErr;
        ret.val.err = -1;
        return &ret;
    }

    XLOPER12 array;
    Excel12f(xlCoerce, &array, [arg]);
    scope(exit) Excel12f(xlFree, null, [&array]);

    const rows = array.val.array.rows;
    const columns = array.val.array.columns;
    auto values = array.val.array.lparray[0 .. (rows * columns)];

    ret.xltype = xltypeNum;
    ret.val.num = 0;

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. columns) {
            XLOPER12 val;
            Excel12f(xlCoerce, &val, [&values[row * columns + col]]);
            scope(exit) Excel12f(xlFree, null, [&val]);

            if(val.xltype != xltypeNum) {
                ret.xltype = xltypeErr;
                ret.val.err = -1;
                return &ret;
            }

            ret.val.num += val.val.num;
        }
    }

    return &ret;
}


mixin(implGetWorksheetFunctionsString!(__MODULE__));
