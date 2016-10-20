# A Makefile for running unit tests on this codebase Using dub isn't
# an easy option since some of the files only compile on Windows. This
# Makefile makes it possible to develop the pure non-Windows-dependent
# functionality in Linux

all: test

UT_DIR=~/.dub/packages/unit-threaded-0.6.30/unit-threaded
UT_LIB=$(UT_DIR)/libunit-threaded.a
UT_SRC=$(UT_DIR)/dub.json

test: bin/ut
	$^

bin/ut: $(UT_LIB) bin/ut.o
	dmd -of$@ $^

bin/ut.o: bin/ut.d xlld/worksheet.d xlld/traits.d xlld/test_module.d
	dmd -c -of$@ -I$(UT_DIR)/source -unittest -g -debug $^

$(UT_LIB): $(UT_SRC)
	cd $(dir $^); dub build

$(UT_SRC):
	dub fetch unit-threaded --version=0.6.30
