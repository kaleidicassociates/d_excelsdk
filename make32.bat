dmd -c -g -m32 -Isource -ofmyxll.obj example/myxll.d source/xlld/xll.d source/xlld/worksheet.d source/xlld/memorymanager.d source/xlld/memorypool.d source/xlld/xlcall.d source/xlld/xlcallcpp.d source/xlld/framework.d source/xlld/xl.d source/xlld/wrap.d source/xlld/test_d_funcs.d
dmd -m32 -ofmyxll32.xll -L/IMPLIB myxll.obj myxll32.def xlcall32d.lib -g -map
