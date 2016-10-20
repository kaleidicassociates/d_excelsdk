/**
 Only exists to test the module reflection functionality
 */
module xlld.test_module;

version(unittest):

import xlld.xlcall;

// extern(C) export means it doesn't have to be explicitly
// added to the .def file
extern(C) export double FuncMulByTwo(double n) {
    return n * 2;
}

extern(C) export double FuncFP12(FP12* cells) {
    return 0;
}


extern(C) export LPXLOPER12 FuncFib (LPXLOPER12 n) {
    return LPXLOPER12.init;
}
