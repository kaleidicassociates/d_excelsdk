/**
 Wraps calls to xlXXX "functions" via the Excel4/Excel12 functions
 */
module xlld.xl;

import xlld.xlcall;
import xlld.wrap;

version(unittest) {
    XLOPER12 coerce(LPXLOPER12 oper) {

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

    void free(LPXLOPER12 oper) {

    }
} else {
    import xlld.framework;
    XLOPER12 coerce(LPXLOPER12 oper) {
        XLOPER12 coerced;
        Excel12f(xlCoerce, &coerced, [oper]);
        return coerced;
    }

    void free(LPXLOPER12 oper) {
        Excel12f(xlFree, null, [oper]);
    }
}
