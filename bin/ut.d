//xlld.framework,xlld.memorymanager,xlld.memorypool,xlld.worksheet,xlld.wrap,xlld.xlcall,xlld.xlcallcpp,xlld.xll
//Automatically generated by unit_threaded.gen_ut_main, do not edit by hand.
import std.stdio;
import unit_threaded;

int main(string[] args)
{
    writeln("\nAutomatically generated file bin/ut.d");
    writeln(`Running unit tests from dirs ["."]`);
    return runTests!(
                     "xlld.worksheet",
                     "xlld.traits",
                     "xlld.newwrap",
                     )
                     (args);
}
