#include <ruby.h>
#include <windows.h>

static VALUE ole_initialize(int argc, VALUE* argv, VALUE self){
  HRESULT hr;
  VALUE v_server, v_host;
  wchar_t* server;
  wchar_t* host;
  IDispatch* dispatch;
  int length;
  CLSID clsid;

  rb_scan_args(argc, argv, "11", &v_server, &v_host);

  SafeStringValue(v_server);

  // Make our server a wide string for later functions
  length = MultiByteToWideChar(CP_UTF8, 0, RSTRING_PTR(v_server), -1, NULL, 0);
  server = (wchar_t*)ruby_xmalloc(MAX_PATH * sizeof(wchar_t));

  if (!MultiByteToWideChar(CP_UTF8, 0, RSTRING_PTR(v_server), -1, server, length)){
    ruby_xfree(server);
    rb_raise(rb_eSystemCallError, "MultibyteToWideChar", GetLastError());
  }

  // TODO: Treat host as nil if argument is same as local computer name or localhost
  if (!NIL_P(v_host)){
    SafeStringValue(v_host);

    length = MultiByteToWideChar(CP_UTF8, 0, RSTRING_PTR(v_host), -1, NULL, 0);
    host = (wchar_t*)ruby_xmalloc(MAX_PATH * sizeof(wchar_t));

    if (!MultiByteToWideChar(CP_UTF8, 0, RSTRING_PTR(v_host), -1, host, length)){
      ruby_xfree(host);
      rb_raise(rb_eSystemCallError, "MultibyteToWideChar", GetLastError());
    }
  }

  // Attempt to get a CLSID using from both ProgID and String
  hr = CLSIDFromProgID(server, &clsid);

  if (FAILED(hr)){
    hr = CLSIDFromString(server, &clsid);

    if (FAILED(hr))
      rb_raise(rb_eSystemCallError, "CLSIDFromString", hr);
  }

  ruby_xfree(server); // Don't need this any more

  hr = CoInitialize(NULL);

  if (FAILED(hr))
    rb_raise(rb_eSystemCallError, "CoInitialize", hr);

  if (NIL_P(v_host)){
    hr = CoCreateInstance(
      &clsid,
      NULL,
      CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER,
      &IID_IDispatch,
      &dispatch
    );
  }
  else{
    COSERVERINFO info;
    MULTI_QI mqi;

    memset(&info, 0, sizeof(COSERVERINFO));
    memset(&mqi, 0, sizeof(MULTI_QI));

    mqi.pIID = &IID_IDispatch;
    mqi.pItf = NULL;

    info.pwszName = host;
    info.pAuthInfo = NULL;

    hr = CoCreateInstanceEx(
      &clsid,
      NULL,
      CLSCTX_INPROC_SERVER | CLSCTX_REMOTE_SERVER,
      &info,
      1,
      &mqi
    );
  }

  if (FAILED(hr))
    rb_raise(rb_eSystemCallError, "CoCreateInstance", hr);

  return self;
}

static VALUE ole_close(VALUE self){
  CoUninitialize();
  return self;
}

void Init_ole(){
  VALUE mWin32 = rb_define_module("Win32");
  VALUE cOle = rb_define_class_under(mWin32, "OLE", rb_cObject);

  rb_define_method(cOle, "initialize", ole_initialize, -1);
  rb_define_method(cOle, "close", ole_close, 0);
}
