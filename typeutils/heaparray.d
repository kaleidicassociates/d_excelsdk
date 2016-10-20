module typeUtils.heapArray;


T[] TempHeapSlice(alias TempAllocFunc, T, Array)(Array u) {
	static assert(is(typeof(TempAllocFunc(typeof(Array.length.init).init)) == void*), "Allocation function does not have required signature");

	static assert(is(typeof(U[0]) : T) &&
		is(typeof(*Array.init.ptr) : T) &&
		is(typeof(Array.init.length) : size_t)
		, "This overload can only be used with Slices or Arrays"
	);
	
	enum sliceSize = T[].sizeof;
	
	(T[])* array = TempAllocFunc(u.length * T.sizeof + sliceSize);
//	*array.length = u.length;
	*array = (cast(T*)((cast(void*)array) + sliceSize))[0 .. u.length];
	assert(&(*array).ptr == array + size_t);
	
	import core.stdc.string : memcpy;
	memcpy(((*array).ptr), u.ptr, u.length);
	
	return *array;
}

T[] TempHeapSlice(alias TempAllocFunc, T, Range)(Range r) {
	static assert(is(typeof(TempAllocFunc(typeof(U.length.init).init)) == void*), "Allocation function does not have required signature");

	static assert(is(typeof(Range[0]) : T) &&
		!is(typeof(Range.init.ptr)) &&
		is(typeof(Range.init.length) : size_t)
		, "This overload can only be used with RandomAccsessRanges"
	);
	
	enum sliceSize = T[].sizeof;
	
	(T[])* array = TempAllocFunc(r.length * T.sizeof + sliceSize);
//	*array.length = u.length;
	*array = (cast(T*)((cast(void*)array) + sliceSize))[0 .. r.length];
	assert(&(*array).ptr == array + size_t);
	
	foreach(i;0 .. r.length) {
		(*array)[i] = elm[i];
	}
		
	return *array;
}

// arrayUtils 

