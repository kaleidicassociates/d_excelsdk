/**
 Only exists to test the module reflection functionality
 */
module xlld.test_module;

version(unittest):

// extern(C) export means it doesn't have to be explicitly
// added to the .def file
extern(C) export double FuncMulByTwo(double n) {
    return n * 2;
}
