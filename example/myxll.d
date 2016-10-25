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
mixin(wrapAll!(__MODULE__, "xlld.test_d_funcs"));
version(main) {
    void main(string[] args) {

        import std.stdio: File;
        import std.exception: enforce;
        import std.path: stripExtension;

        enforce(args.length >= 2 && args.length <= 4,
                "Usage: " ~ args[0] ~ " [file_name] <lib_name> <description>");

        immutable fileName = args[1];
        immutable libName = args.length > 2 ? args[2] : fileName.stripExtension ~ ".xll";
        immutable description = args.length > 3 ? args[3] : "Simple D add-in to Excel";

        auto file = File(fileName, "w");
        foreach(stmt; dllDefFile!__MODULE__(libName, description).statements)
            file.writeln(stmt.toString);
    }
}
