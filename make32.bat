@echo off

rem Write the .def file automatically
dub run -c def --nodeps -- myxll32.def || goto :error

rem Build the XLL object file
dmd -m32 -c -ofmyxll.obj -debug -g -w -version=Have_d_excelsdk -Isource source\xlld\dummy.d source\xlld\framework.d source\xlld\memorymanager.d source\xlld\memorypool.d source\xlld\package.d source\xlld\test_d_funcs.d source\xlld\test_xl_funcs.d source\xlld\traits.d source\xlld\worksheet.d source\xlld\wrap.d source\xlld\xl.d source\xlld\xlcall.d source\xlld\xlcallcpp.d source\xlld\xll.d example\myxll.d -vcolumns

rem Link the final XLL to be loaded by Excel
dmd -ofmyxll32.xll -L/IMPLIB myxll.obj myxll32.def d_excelsdk.lib xlcall32d.lib || goto :error

echo.
echo.
echo Build successful
echo.
exit /b 0

:error
echo Failed with error #%errorlevel%.
exit /b %errorlevel%
