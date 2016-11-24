/**
 Wraps calls to xlXXX "functions" via the Excel4/Excel12 functions
 */
module xlld.xl;

import xlld.xlcall;
import xlld.wrap;

version(unittest) {
    XLOPER12 coerce(LPXLOPER12 oper) nothrow @nogc {

        XLOPER12 ret;
        ret = *oper;

        switch(oper.xltype) {

        case xltypeSRef:
            ret.xltype = xltypeMulti;
            return ret;

        case xltypeNum:
            ret.xltype = xltypeNum;
            return ret;

        default:
            return ret;
        }
    }

    void free(LPXLOPER12 oper) nothrow @nogc {

    }
} else {
    import xlld.framework;
    XLOPER12 coerce(LPXLOPER12 oper) nothrow {
        XLOPER12 coerced;
        LPXLOPER12[1] arg = [oper];
        Excel12f(xlCoerce, &coerced, arg);
        return coerced;
    }

    void free(LPXLOPER12 oper) nothrow {
        LPXLOPER12[1] arg = [oper];
        Excel12f(xlFree, null, arg);
    }
}
