@echo off

rem Write the .def file automatically
dub run -c def --nodeps -- myxll32.def || goto :error

rem Build the XLL object file
dmd -c -g -m32 -Isource -ofmyxll.obj example/myxll.d source/xlld/xll.d source/xlld/worksheet.d source/xlld/memorymanager.d source/xlld/memorypool.d source/xlld/xlcall.d source/xlld/xlcallcpp.d source/xlld/framework.d source/xlld/xl.d source/xlld/wrap.d source/xlld/test_d_funcs.d || goto :error

rem Link the final XLL to be loaded by Excel
dmd -m32 -ofmyxll32.xll -L/IMPLIB myxll.obj myxll32.def d_excelsdk.lib xlcall32d.lib -g -map || goto :error

echo.
echo.
echo Build successful
echo.
exit /b 0

:error
echo Failed with error #%errorlevel%.
exit /b %errorlevel%
