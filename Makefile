# A Makefile for running unit tests on this codebase Using dub isn't
# an easy option since some of the files only compile on Windows. This
# Makefile makes it possible to develop the pure non-Windows-dependent
# functionality in Linux

all: test

LIB_UT=~/.dub/packages/unit-threaded-0.6.30/unit-threaded/libunit-threaded.a
UT_FILES=$(addprefix ~/.dub/packages/unit-threaded-0.6.30/unit-threaded/source/unit_threaded/, asserts.d dub.d integration.d meta.d options.d reflection.d runtime.d testcase.d testsuite.d attrs.d factory.d io.d mock.d package.d runner.d should.d uda.d)

test: bin/ut
	$^

bin/ut: bin/ut.o $(LIB_UT)
	dmd -of$@ $^

bin/ut.o: bin/ut.d xlld/worksheet.d xlld/traits.d xlld/test_module.d
	dmd -c -of$@ -I~/.dub/packages/unit-threaded-0.6.30/unit-threaded/source -unittest -g -debug $^

$(LIB_UT): ~/.dub/packages/unit-threaded-0.6.30/unit-threaded/dub.json
	dub fetch unit-threaded --version=0.6.30; cd $^; dub build
