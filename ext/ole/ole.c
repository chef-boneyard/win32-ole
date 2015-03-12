#include <ruby.h>
#include <ruby/encoding.h>
#include <windows.h>
#include <objbase.h>

static VALUE ole_initialize(int argc, VALUE* argv, VALUE self){
  HRESULT hr;
  VALUE v_server, v_host;
  wchar_t* server;
  void* dispatch;
  int length;
  CLSID clsid;
  rb_encoding* server_encoding;
  rb_econv_t* ec;
  const int replaceflags = ECONV_UNDEF_REPLACE | ECONV_INVALID_REPLACE;

  rb_scan_args(argc, argv, "11", &v_server, &v_host);

  SafeStringValue(v_server);

  server_encoding = rb_enc_get(v_server);
  ec = rb_econv_open(rb_enc_name(server_encoding), "UTF-8", replaceflags);
  v_server = rb_econv_str_convert(ec, v_server, ECONV_PARTIAL_INPUT);
  rb_econv_close(ec);

  // Make our server a wide string for later functions
  length = MultiByteToWideChar(CP_UTF8, 0, RSTRING_PTR(v_server), -1, NULL, 0);
  server = (wchar_t*)ruby_xmalloc(MAX_PATH * sizeof(wchar_t));

  if (!MultiByteToWideChar(CP_UTF8, 0, RSTRING_PTR(v_server), -1, server, length)){
    ruby_xfree(server);
    rb_raise(rb_eSystemCallError, "MultibyteToWideChar", GetLastError());
  }

  if (!NIL_P(v_host))
    SafeStringValue(v_host);

  // Attempt to get a CLSID using from both ProgID and String
  hr = CLSIDFromProgID(server, &clsid);

  if (FAILED(hr)){
    hr = CLSIDFromString(server, &clsid);

    if (FAILED(hr))
      rb_raise(rb_eSystemCallError, "CLSIDFromString", hr);
  }

  hr = CoInitialize(NULL);

  if (FAILED(hr))
    rb_raise(rb_eSystemCallError, "CoInitialize", hr);

  //hr = CoCreateInstance(
  //  &clsid,
  //  NULL,
  //  CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
  //  IID_IDispatch,
  //  dispatch
  //);

  //if (FAILED(hr))
  //  rb_raise(rb_eSystemCallError, "CoCreateInstance", hr);

  return self;
}

void Init_ole(){
  VALUE mWin32 = rb_define_module("Win32");
  VALUE cOle = rb_define_class_under(mWin32, "OLE", rb_cObject);

  rb_define_method(cOle, "initialize", ole_initialize, -1);
}
