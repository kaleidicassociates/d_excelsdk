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
    import std.traits: ReturnType, Parameters;
    import std.conv: text;

    alias wrappedFunc = xlld.test_d_funcs.FuncAddEverything;
    alias InputType = Parameters!wrappedFunc[0];

    static assert(Parameters!wrappedFunc.length == 1,
                  text("Illegal number of parameters, only 1 supported, not ",
                       Parametes!wrappedFunc.length));

    static XLOPER12 ret;

    if(arg.xltype != dlangToXlOperInputType!InputType) {
        ret.xltype = xltypeErr;
        ret.val.err = -1;
        return &ret;
    }

    XLOPER12 realArg;
    Excel12f(xlCoerce, &realArg, [arg]);
    scope(exit) Excel12f(xlFree, null, [&realArg]);

    if(realArg.xltype != dlangToXlOperType!InputType) {
        ret.xltype = xltypeErr;
        ret.val.err = -1;
        return &ret;
    }

    ret.xltype = dlangToXlOperType!(ReturnType!wrappedFunc);
    ret.val.num = 0;

    const rows = realArg.val.array.rows;
    const columns = realArg.val.array.columns;

    // the type expected by the D function being wrapped
    InputType forwardArg;
    forwardArg.length = rows;
    foreach(ref col; forwardArg) col.length = columns;

    foreach(const row; 0 .. rows) {
        foreach(const col; 0 .. columns) {
            XLOPER12 val;
            Excel12f(xlCoerce, &val, [&realArg.val.array.lparray[row * columns + col]]);
            scope(exit) Excel12f(xlFree, null, [&val]);

            if(val.xltype == xltypeNum)
                forwardArg[row][col] = val.val.num;
            else
                forwardArg[row][col] = 0;
        }
    }

    wrappedFunc(forwardArg);
    ret = toXlOper(wrappedFunc(forwardArg));

    return &ret;
}


mixin(implGetWorksheetFunctionsString!(__MODULE__));
