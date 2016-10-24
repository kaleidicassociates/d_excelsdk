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
    import xlld.xl: coerce, free, convertInput;
    import std.traits: ReturnType, Parameters;
    import std.conv: text;

    alias wrappedFunc = xlld.test_d_funcs.FuncAddEverything;
    alias InputType = Parameters!wrappedFunc[0];

    static assert(Parameters!wrappedFunc.length == 1,
                  text("Illegal number of parameters, only 1 supported, not ",
                       Parameters!wrappedFunc.length));

    static XLOPER12 ret;

    // must 1st convert argument to the "real" type.
    // 2D arrays are passed in as SRefs, for instance
    XLOPER12 realArg;
    try {
        realArg = convertInput!InputType(arg);
    } catch(Exception ex) {
        ret.xltype = xltypeErr;
        ret.val.err = -1;
        return &ret;
    }

    scope(exit) free(&realArg);

    const rows = realArg.val.array.rows;
    const columns = realArg.val.array.columns;

    // the type expected by the D function being wrapped
    auto forwardArg = fromXlOper!InputType(&realArg);

    wrappedFunc(forwardArg);
    ret = toXlOper(wrappedFunc(forwardArg));

    return &ret;
}



mixin(implGetWorksheetFunctionsString!(__MODULE__));
