#include <ruby.h>
#include <ruby/encoding.h>
#include <windows.h>

void Init_ole(){
  VALUE mWin32 = rb_define_module("Win32");
  VALUE cOle = rb_define_class_under(mWin32, "OLE", rb_cObject);
}
