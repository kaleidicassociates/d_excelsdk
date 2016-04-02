module xlltypes;

struct Boolean {
	short _value;
	@property bool asBool() pure {
		return _value != 0;
	}

	@property void asBool(bool val) pure {
		val ? _value = 1 : _value = 0;
	}

	alias this asBool;
}

struct CountedString4 {
	private ubyte* _data;

	@property string _string() {
		return cast(string) _data[1 .. length + 1];
	} 

	@property ubyte length() {
		return _data[0];
	}
	
	alias _string this;
}
