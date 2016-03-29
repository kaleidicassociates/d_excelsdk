module genXllDescr;
import xlltypes;
template descr(alias exportedFunction) {
	enum descr = mixin(descr_!exportedFunction);
	static assert(is(typeof(descr) == wstring[]));
}

wchar typeText(T)(T t) {
	// are C and F the same ?
	// F is modified in place
	// C is not
	static if (is(T == LxOper)) {

	} else static if (is(T == double)) {
		return "B"w;
	} else static if (is(T == double*)) {
		return "E"w;
	} else static if (is(T == Boolean)) {
		return "A"w;
	} else static if (is(T == Boolean*)) {
		return "L"w;
	} else static if (is(T == const char*)) {
		return "C"w;
	} else static if (is(T == char*)) {
		return "F"w;
	}
}

string descr_(alias exportedFunction)() {
	import std.traits;
	static assert(hasUDA!(exportedFunction, Xll));



	string result = "[";
	/* Output looks like :
	[ "Func1"w,                                     // Procedure
		"UU"w,                                  // type_text
		"Func1"w,                               // function_text
		"Arg"w,                                 // argument_text
		"1"w,                                   // macro_type
		"Generic Add-In"w,                      // category
		""w,                                    // shortcut_text
		""w,                                    // help_topic
		"Always returns the string 'Func1'"w,   // function_help
		"Argument ignored"w                     // argument_help1
	]
*/
	 
}