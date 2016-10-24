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

mixin(wrapModuleFunctionStr("xlld.test_d_funcs", "FuncAddEverything"));
mixin(wrapModuleFunctionStr("xlld.test_d_funcs", "FuncAllLengths"));


mixin(implGetWorksheetFunctionsString!(__MODULE__));
