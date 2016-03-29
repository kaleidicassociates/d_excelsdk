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