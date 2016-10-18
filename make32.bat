dmd -c -g -m32 -ofmyxll.obj myxll.d  memorymanager.d memorypool.d xlcall.d xlcallcpp.d framework.d wrap.d
dmd -m32 -ofmyxll32.xll -L/IMPLIB myxll.obj generic32.def xlcall32d.lib -g -map
