// Native HTTP bootstrap runs on a dedicated worker inside the WPS process.
// Full mode can start its verified sibling service; no shell or Agent is invoked.
#define UNICODE
#define _UNICODE
#define WIN32_LEAN_AND_MEAN
#define COBJMACROS
#include <winsock2.h>
#include <windows.h>
#include <winhttp.h>
#include <wincrypt.h>
#ifdef OFFICEBRIDGE_FULL_BOOTSTRAP
#include <iphlpapi.h>
#endif
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>
#include <oleauto.h>
#include "native-config.h"
#include "native-registry.h"
#ifndef NATIVE_LISTEN_PORT
#define NATIVE_LISTEN_PORT 43191
#define NATIVE_ADDIN_NAME "OfficeBridgeXllProbe"
#define NATIVE_ADDIN_URL "http://127.0.0.1:43191/probe/"
#define NATIVE_ADDIN_VERSION "5.0.0"
#define NATIVE_DEPLOY_JSON "{\"cmd\":\"enable\",\"name\":\"OfficeBridgeXllProbe\",\"addonType\":\"et\",\"online\":\"true\",\"url\":\"http://127.0.0.1:43191/probe/\",\"version\":\"5.0.0\",\"customDomain\":\"\"}"
#define NATIVE_INVOKE_JSON "{\"name\":\"OfficeBridgeXllProbe\",\"function\":\"OfficeBridgeProbeOpen\",\"info\":\"{}\",\"jsPluginsXml\":\"\"}"
#endif
#ifdef OFFICEBRIDGE_FULL_BOOTSTRAP
#define NATIVE_TITLE L"办公桥"
#define NATIVE_MENU L"办公桥"
#else
#define NATIVE_TITLE L"OfficeBridge Native Probe"
#define NATIVE_MENU L"办公桥验证"
#endif


static HMODULE image;
static volatile LONG opening,stop_server,reads,deployment,invocation,page_loads,registration,activation_error,relay_http;
static char page_token[65],last_address[32];
static DWORD relay_port=58890,deployment_budget=120000,server_port=NATIVE_LISTEN_PORT;
static BOOL quiet_activation;

static void diagnostic(const char *stage,DWORD code) {
  WCHAR dir[32768],path[32768];DWORD n=GetEnvironmentVariableW(L"LOCALAPPDATA",dir,29000);
  if(!n||n>28000)return;
  wcscat(dir,L"\\OfficeBridge");CreateDirectoryW(dir,0);
  DWORD attributes=GetFileAttributesW(dir);if(attributes==INVALID_FILE_ATTRIBUTES||(attributes&FILE_ATTRIBUTE_REPARSE_POINT))return;
  wcscat(dir,L"\\XllNativeDiagnostics");CreateDirectoryW(dir,0);
  attributes=GetFileAttributesW(dir);if(attributes==INVALID_FILE_ATTRIBUTES||(attributes&FILE_ATTRIBUTE_REPARSE_POINT))return;
  swprintf(path,32768,L"%ls\\startup-%lu.txt",dir,GetCurrentProcessId());
  HANDLE file=CreateFileW(path,FILE_APPEND_DATA,FILE_SHARE_READ|FILE_SHARE_WRITE,0,OPEN_ALWAYS,FILE_ATTRIBUTE_NORMAL|FILE_FLAG_OPEN_REPARSE_POINT,0);
  if(file==INVALID_HANDLE_VALUE)return;
  BY_HANDLE_FILE_INFORMATION info;
  if(GetFileInformationByHandle(file,&info)&&!(info.dwFileAttributes&FILE_ATTRIBUTE_REPARSE_POINT)) {
    char line[256];DWORD written=0;int length=snprintf(line,sizeof(line),"{\"stage\":\"%s\",\"code\":%lu,\"pid\":%lu}\r\n",stage,code,GetCurrentProcessId());
    if(length>0&&length<(int)sizeof(line))WriteFile(file,line,(DWORD)length,&written,0);
  }CloseHandle(file);
}
static int last_rows,last_columns;
static char host[32]="127.0.0.1:43191";
static char origin[64]="http://127.0.0.1:43191";

// The system COM object owns asynchronous I/O; no callback targets this DLL.
// Polling does not resend the POST. Abort runs on this activation worker only.
static VARIANT argument_number(LONG value) {VARIANT v;VariantInit(&v);V_VT(&v)=VT_I4;V_I4(&v)=value;return v;}
static VARIANT argument_bool(BOOL value) {VARIANT v;VariantInit(&v);V_VT(&v)=VT_BOOL;V_BOOL(&v)=value?VARIANT_TRUE:VARIANT_FALSE;return v;}
static VARIANT argument_text(const WCHAR *value) {VARIANT v;VariantInit(&v);V_VT(&v)=VT_BSTR;V_BSTR(&v)=SysAllocString(value);return v;}
static HRESULT http_method(IDispatch *object,const WCHAR *name,WORD flags,VARIANT *args,UINT count,VARIANT *result) {
  LPOLESTR member=(LPOLESTR)name;DISPID id=0,put=DISPID_PROPERTYPUT;HRESULT hr=S_OK;EXCEPINFO exception={0};UINT argument_error=0;
  for(UINT i=0;i<count;i++)if(V_VT(&args[i])==VT_BSTR&&!V_BSTR(&args[i]))hr=E_OUTOFMEMORY;
  if(SUCCEEDED(hr))hr=IDispatch_GetIDsOfNames(object,&IID_NULL,&member,1,LOCALE_INVARIANT,&id);
  DISPPARAMS parameters={args,(flags&DISPATCH_PROPERTYPUT)?&put:0,count,(flags&DISPATCH_PROPERTYPUT)?1:0};
  if(SUCCEEDED(hr))hr=IDispatch_Invoke(object,id,&IID_NULL,LOCALE_INVARIANT,flags,&parameters,result,&exception,&argument_error);
  if(hr==DISP_E_EXCEPTION&&exception.scode)hr=exception.scode;
  SysFreeString(exception.bstrSource);SysFreeString(exception.bstrDescription);SysFreeString(exception.bstrHelpFile);
  for(UINT i=0;i<count;i++)VariantClear(&args[i]);
  return hr;
}
static DWORD http_error(HRESULT hr) {return HRESULT_FACILITY(hr)==FACILITY_WIN32?(DWORD)HRESULT_CODE(hr):(DWORD)hr;}
static BOOL exchange_at(DWORD port,const WCHAR *path,const char *body,char *reply,DWORD capacity,DWORD budget,DWORD *http_status) {
  BOOL ok=FALSE,initialized=FALSE;DWORD error=0;HRESULT hr;CLSID clsid;IDispatch *request=0;VARIANT args[4],result;
  WCHAR url[256],text[2300];ULONGLONG deadline=GetTickCount64()+budget;VariantInit(&result);
  *http_status=0;
  if(stop_server){SetLastError(ERROR_CANCELLED);return FALSE;}
  hr=CoInitializeEx(0,COINIT_MULTITHREADED);if(FAILED(hr)){error=http_error(hr);goto cleanup;}initialized=TRUE;
  hr=CLSIDFromProgID(L"WinHttp.WinHttpRequest.5.1",&clsid);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  hr=CoCreateInstance(&clsid,0,CLSCTX_INPROC_SERVER,&IID_IDispatch,(void**)&request);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  args[0]=argument_number(1);hr=http_method(request,L"SetProxy",DISPATCH_METHOD,args,1,0);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  args[0]=argument_number((LONG)budget);args[1]=argument_number(3000);args[2]=argument_number(1000);args[3]=argument_number(1000);
  hr=http_method(request,L"SetTimeouts",DISPATCH_METHOD,args,4,0);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  swprintf(url,256,L"http://127.0.0.1:%lu%ls",port,path);
  args[0]=argument_bool(TRUE);args[1]=argument_text(url);args[2]=argument_text(body?L"POST":L"GET");
  hr=http_method(request,L"Open",DISPATCH_METHOD,args,3,0);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  // WinHttpRequestOption_EnableRedirects=6; never leave the local relay.
  args[0]=argument_bool(FALSE);args[1]=argument_number(6);
  hr=http_method(request,L"Option",DISPATCH_PROPERTYPUT,args,2,0);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  args[0]=argument_number(2);hr=http_method(request,L"SetAutoLogonPolicy",DISPATCH_METHOD,args,1,0);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  if(body&&!MultiByteToWideChar(CP_UTF8,MB_ERR_INVALID_CHARS,body,-1,text,2300)){error=GetLastError();goto cleanup;}
  if(stop_server){error=ERROR_CANCELLED;goto cleanup;}
  args[0]=argument_text(body?text:L"");hr=http_method(request,L"Send",DISPATCH_METHOD,args,1,0);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  for(;;) {
    if(stop_server){error=ERROR_CANCELLED;goto cleanup;}
    if(GetTickCount64()>=deadline){error=ERROR_WINHTTP_TIMEOUT;goto cleanup;}
    args[0]=argument_number(0);hr=http_method(request,L"WaitForResponse",DISPATCH_METHOD,args,1,&result);
    if(FAILED(hr)){error=http_error(hr);goto cleanup;}
    if(V_VT(&result)!=VT_BOOL){error=ERROR_INVALID_DATA;goto cleanup;}
    BOOL finished=V_BOOL(&result)==VARIANT_TRUE;VariantClear(&result);
    if(finished)break;Sleep(100);
  }
  if(stop_server){error=ERROR_CANCELLED;goto cleanup;}
  if(GetTickCount64()>=deadline){error=ERROR_WINHTTP_TIMEOUT;goto cleanup;}
  hr=http_method(request,L"Status",DISPATCH_PROPERTYGET,0,0,&result);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  if(V_VT(&result)==VT_I4)*http_status=(DWORD)V_I4(&result);
  if(V_VT(&result)!=VT_I4||V_I4(&result)!=200){if(V_VT(&result)==VT_I4)diagnostic("relay_http_status",(DWORD)V_I4(&result));error=ERROR_WINHTTP_INVALID_SERVER_RESPONSE;goto cleanup;}
  VariantClear(&result);hr=http_method(request,L"ResponseText",DISPATCH_PROPERTYGET,0,0,&result);if(FAILED(hr)){error=http_error(hr);goto cleanup;}
  if(V_VT(&result)!=VT_BSTR){error=ERROR_INVALID_DATA;goto cleanup;}
  UINT length=SysStringLen(V_BSTR(&result));
  if(length>=capacity){error=ERROR_INSUFFICIENT_BUFFER;goto cleanup;}
  if(length&&wmemchr(V_BSTR(&result),0,length)){error=ERROR_INVALID_DATA;goto cleanup;}
  if(!length)reply[0]=0;
  else {int bytes=WideCharToMultiByte(CP_UTF8,0,V_BSTR(&result),(int)length,reply,(int)capacity-1,0,0);if(!bytes){error=GetLastError();goto cleanup;}reply[bytes]=0;}
  ok=TRUE;
cleanup:
  VariantClear(&result);
  if(request){if(!ok)http_method(request,L"Abort",DISPATCH_METHOD,0,0,0);IDispatch_Release(request);}
  if(initialized)CoUninitialize();SetLastError(error);return ok;
}
static BOOL exchange(const WCHAR *path,const char *body,char *reply,DWORD capacity,DWORD budget,DWORD *http_status) {
  return exchange_at(relay_port,path,body,reply,capacity,budget,http_status);
}
static void encoded(const char *json,char *out,DWORD capacity) {
  CryptBinaryToStringA((const BYTE*)json,(DWORD)strlen(json),CRYPT_STRING_BASE64|CRYPT_STRING_NOCRLF,out,&capacity);
}
static void activation_failure(const WCHAR *stage,DWORD code,DWORD http_status) {
  InterlockedExchange(&activation_error,(LONG)code);InterlockedExchange(&relay_http,(LONG)http_status);
  diagnostic("activation_failure",code);
  if(stop_server||quiet_activation)return;
  WCHAR message[512];
  if(!wcscmp(stage,L"registration_conflict"))
    swprintf(message,512,L"办公桥引导已启动。WPS 中已有办公桥记录，但地址、版本、启用状态或记录内容未通过核对。未覆盖安装，未请求打开侧栏，没有修改表格。");
  else if(!wcscmp(stage,L"registration_query"))
    swprintf(message,512,L"办公桥引导已启动，但无法确认 WPS 中的加载项安装状态：HTTP %lu，错误码 %lu。未重新安装或请求打开侧栏，没有修改表格。",http_status,code);
  else if(code==ERROR_WINHTTP_TIMEOUT&&!wcscmp(stage,L"deployment_request"))
    swprintf(message,512,L"办公桥引导已启动。等待 WPS 安装确认超时：请检查是否存在“信任并安装”窗口。未自动重试或请求打开侧栏，没有修改表格。");
  else if(!wcscmp(stage,L"deployment_request")||!wcscmp(stage,L"deployment_response"))
    swprintf(message,512,L"办公桥引导已启动，但 WPS 安装响应未确认：HTTP %lu，错误码 %lu。若已安装，可从“" NATIVE_MENU L"”菜单手动打开。未自动重试，没有修改表格。",http_status,code);
  else
    swprintf(message,512,L"办公桥引导已启动，WPS 中的安装已确认，但打开侧栏的请求未确认：HTTP %lu，错误码 %lu。可从“" NATIVE_MENU L"”菜单手动打开。未自动重试，没有修改表格。",http_status,code);
  MessageBoxW(0,message,NATIVE_TITLE,MB_OK|MB_ICONWARNING);
}
static DWORD WINAPI activate(void *unused) {
  (void)unused;char body[2048],response[2048],request[2200];DWORD status=0;
  if(stop_server)return 0;
  char *installed=(char*)HeapAlloc(GetProcessHeap(),0,65536);
  if(!installed){InterlockedExchange(&registration,-2);activation_failure(L"registration_query",ERROR_NOT_ENOUGH_MEMORY,0);return 0;}
  BOOL queried=exchange(L"/publishlist","",installed,65536,5000,&status);DWORD query_error=GetLastError();
  #ifdef OFFICEBRIDGE_FULL_BOOTSTRAP
  int existing=queried?officebridge_registry_match_online(installed,NATIVE_ADDIN_NAME,NATIVE_ADDIN_URL,NATIVE_ADDIN_VERSION):-1;
  #else
  int existing=queried?officebridge_registry_match(installed,NATIVE_ADDIN_NAME,NATIVE_ADDIN_URL,NATIVE_ADDIN_VERSION):-1;
  #endif
  HeapFree(GetProcessHeap(),0,installed);
  InterlockedExchange(&relay_http,(LONG)status);
  if(!queried){InterlockedExchange(&registration,-2);activation_failure(L"registration_query",query_error,status);return 0;}
  if(stop_server)return 0;
  if(existing<0){InterlockedExchange(&registration,-1);activation_failure(L"registration_conflict",ERROR_INVALID_DATA,status);return 0;}
  if(existing==1) {
    InterlockedExchange(&registration,2);InterlockedExchange(&deployment,3);
    diagnostic("registration_reused",0);
  } else {
    InterlockedExchange(&registration,1);
    encoded(NATIVE_DEPLOY_JSON,body,sizeof(body));
    InterlockedExchange(&deployment,2);
    if(!exchange(L"/deployaddons/runParams",body,response,sizeof(response),deployment_budget,&status)){DWORD code=GetLastError();InterlockedExchange(&deployment,-1);activation_failure(L"deployment_request",code,status);return 0;}
    InterlockedExchange(&relay_http,(LONG)status);
    if(stop_server)return 0;
    if(strcmp(response,"OK")&&response[0]){InterlockedExchange(&deployment,-2);activation_failure(L"deployment_response",ERROR_INVALID_DATA,status);return 0;}
    InterlockedExchange(&deployment,1);
  }
  if(stop_server)return 0;
  encoded(NATIVE_INVOKE_JSON,body,sizeof(body));
  snprintf(request,sizeof(request),"ksowebstartupet://%s",body);
  if(stop_server)return 0;
  if(exchange(L"/et/runParams",request,response,sizeof(response),30000,&status)) {
    InterlockedExchange(&relay_http,(LONG)status);
    if(stop_server)return 0;
    if(strcmp(response,"OK")&&response[0]){InterlockedExchange(&invocation,-1);activation_failure(L"invocation",ERROR_INVALID_DATA,status);return 0;}
    InterlockedExchange(&invocation,1);diagnostic("open_requested",status);
  } else {DWORD code=GetLastError();InterlockedExchange(&invocation,-1);activation_failure(L"invocation",code,status);}
  return 0;
}
static ULONGLONG request_deadline;
static BOOL send_all(SOCKET socket,const char *data,int length) {
  while(length>0){if(stop_server||GetTickCount64()>=request_deadline)return FALSE;int n=send(socket,data,length,0);if(n<=0)return FALSE;data+=n;length-=n;}return TRUE;
}
static void reply(SOCKET socket,int code,const char *type,const char *body,int length) {
  char header[768];
  int n=snprintf(header,sizeof(header),"HTTP/1.1 %d %s\r\nContent-Type: %s\r\nContent-Length: %d\r\nConnection: close\r\nCache-Control: no-store\r\nX-Content-Type-Options: nosniff\r\nContent-Security-Policy: default-src 'self'; connect-src 'self'; frame-ancestors 'self'\r\n\r\n",code,code==200?"OK":"Rejected",type,length);
  if(send_all(socket,header,n))send_all(socket,body,length);
}
static void json_reply(SOCKET socket,int code,const char *body){reply(socket,code,"application/json",body,(int)strlen(body));}
static const char *header_value(char *request,const char *name) {
  static char value[4096];size_t length=strlen(name);char *line=strstr(request,"\r\n");
  while(line&&line[2]&&line[2]!='\r') {
    line+=2;char *end=strstr(line,"\r\n");if(!end)break;
    if(!_strnicmp(line,name,length)&&line[length]==':') {
      char *start=line+length+1;while(start<end&&(*start==' '||*start=='\t'))start++;
      size_t size=(size_t)(end-start);if(size>=sizeof(value))return "";
      memcpy(value,start,size);value[size]=0;return value;
    }
    line=end;
  }return "";
}
static BOOL valid_address(const char *address) {
  const char *p=address;
  for(int part=0;part<2;part++) {
    int letters=0;while(*p>='A'&&*p<='Z'){letters++;p++;}if(letters<1||letters>3||*p<'1'||*p>'9')return FALSE;
    int digits=0;while(*p>='0'&&*p<='9'){digits++;p++;}if(digits>7)return FALSE;
    if(!*p)return TRUE;if(part||*p++!=':')return FALSE;
  }return FALSE;
}

#ifdef OFFICEBRIDGE_FULL_BOOTSTRAP
static DWORD native_backend_port=NATIVE_BACKEND_PORT;
// Only the HTTP worker owns these handles. The launch worker publishes fields
// before its interlocked state transition: 0 untouched, 1 queued, 2 spawned,
// 3 failed. The process handle is retained for observation, never termination.
static volatile LONG native_launch_state;
static HANDLE native_launch_thread,native_process;
static DWORD native_launch_error,native_process_pid,native_exit_code;
static BOOL native_exit_known;
#include "native-runtime.h"
static BOOL native_field(const cJSON *root,const char *key,const char *value) {
  const cJSON *field=cJSON_GetObjectItemCaseSensitive(root,key);
  return cJSON_IsString(field)&&!strcmp(field->valuestring,value);
}
// Read-only fixed-target probe. Full API streaming remains in the Node server;
// no Agent or spreadsheet operation is ever proxied through this serial server.
// Read the listener table instead of waiting for a failed TCP connection.
// Closed local ports can otherwise take a full connect timeout on Windows.
static int native_listener_state(ULONG family) {
  DWORD size=0;DWORD code=GetExtendedTcpTable(0,&size,FALSE,family,TCP_TABLE_OWNER_PID_LISTENER,0);
  if(code!=ERROR_INSUFFICIENT_BUFFER||size>4*1024*1024)return -1;
  void *table=HeapAlloc(GetProcessHeap(),0,size);if(!table)return -1;
  code=GetExtendedTcpTable(table,&size,FALSE,family,TCP_TABLE_OWNER_PID_LISTENER,0);
  int state=code==NO_ERROR?0:-1;
  if(code==NO_ERROR&&family==AF_INET) {
    MIB_TCPTABLE_OWNER_PID *rows=(MIB_TCPTABLE_OWNER_PID*)table;
    for(DWORD i=0;i<rows->dwNumEntries;i++)if(ntohs((u_short)rows->table[i].dwLocalPort)==native_backend_port&&
      (rows->table[i].dwLocalAddr==0||rows->table[i].dwLocalAddr==htonl(INADDR_LOOPBACK))){state=1;break;}
  } else if(code==NO_ERROR&&family==AF_INET6) {
    MIB_TCP6TABLE_OWNER_PID *rows=(MIB_TCP6TABLE_OWNER_PID*)table;BYTE any[16]={0},loopback[16]={0};loopback[15]=1;
    for(DWORD i=0;i<rows->dwNumEntries;i++)if(ntohs((u_short)rows->table[i].dwLocalPort)==native_backend_port&&
      (!memcmp(rows->table[i].ucLocalAddr,any,16)||!memcmp(rows->table[i].ucLocalAddr,loopback,16))){state=1;break;}
  }
  HeapFree(GetProcessHeap(),0,table);return state;
}
static int native_backend_port_state(void) {
  int ipv4=native_listener_state(AF_INET);if(ipv4)return ipv4;
  return native_listener_state(AF_INET6);
}
static int native_backend_state(void) {
  int connection=native_backend_port_state();if(connection!=1)return connection;
  char result[2048];DWORD status=0;
  if(!exchange_at(native_backend_port,L"/api/v1/health",0,result,sizeof(result),1500,&status))
    return -1;
  if(!officebridge_registry_strings_safe(result))return -1;
  cJSON *root=cJSON_ParseWithOpts(result,0,1);
  BOOL matched=root&&cJSON_IsObject(root)&&officebridge_registry_tree_safe(root,0)&&
    native_field(root,"service","office-bridge")&&native_field(root,"status","ok")&&native_field(root,"protocolVersion",NATIVE_BRIDGE_PROTOCOL_VERSION);
  cJSON_Delete(root);return matched?1:-1;
}
static char *native_request_body(SOCKET socket,char *request,char *end,int count) {
  char *tail=0;const char *header=header_value(request,"Content-Length");
  unsigned long length=strtoul(header,&tail,10);
  char *body=end+4;int available=count-(int)(body-request);
  if(!*header||!tail||*tail||length<1||length>1024||(body-request)+length>=8192)return 0;
  while(available<(int)length) {
    if(stop_server||GetTickCount64()>=request_deadline)return 0;
    int size=recv(socket,body+available,(int)length-available,0);if(size<=0)return 0;available+=size;
  }
  if(memchr(body,0,length))return 0;body[length]=0;return body;
}
static void native_refresh_launch(void) {
  if(InterlockedCompareExchange(&native_launch_state,0,0)!=2)return;
  DWORD waited=WaitForSingleObject(native_process,0);
  if(waited==WAIT_OBJECT_0) {
    DWORD code=0;
    if(GetExitCodeProcess(native_process,&code)){native_exit_code=code;native_exit_known=TRUE;native_launch_error=0;}
    else {native_launch_error=GetLastError();native_exit_known=FALSE;}
    InterlockedExchange(&native_launch_state,3);
  } else if(waited==WAIT_FAILED) {
    native_launch_error=GetLastError();native_exit_known=FALSE;
    InterlockedExchange(&native_launch_state,3);
  }
}
static DWORD WINAPI native_start_service(void *unused) {
  (void)unused;
  WCHAR command[32832],directory[32768];STARTUPINFOW startup={0};PROCESS_INFORMATION process={0};
  startup.cb=sizeof(startup);
  wcscpy(directory,native_runtime_path);WCHAR *separator=wcsrchr(directory,L'\\');
  if(!separator){native_launch_error=ERROR_BAD_PATHNAME;InterlockedExchange(&native_launch_state,3);return 0;}
  if(separator==directory+2)separator[1]=0;else *separator=0;
  int length=swprintf(command,32832,L"\"%ls\" --no-open --port %lu",native_runtime_path,native_backend_port);
  if(length<0||length>=32767){native_launch_error=ERROR_FILENAME_EXCED_RANGE;InterlockedExchange(&native_launch_state,3);return 0;}
  if(stop_server){native_launch_error=ERROR_CANCELLED;InterlockedExchange(&native_launch_state,3);return 0;}
  // Ordinary same-token process creation, no shell, elevation, broker or policy
  // changes. Both executable and arguments are fixed by this compiled package.
  if(!CreateProcessW(native_runtime_path,command,0,0,FALSE,CREATE_NO_WINDOW,0,directory,&startup,&process)) {
    native_launch_error=GetLastError();diagnostic("service_create_failed",native_launch_error);
    InterlockedExchange(&native_launch_state,3);return 0;
  }
  CloseHandle(process.hThread);native_process=process.hProcess;native_process_pid=process.dwProcessId;
  native_launch_error=0;InterlockedExchange(&native_launch_state,2);diagnostic("service_created",process.dwProcessId);return 0;
}
static void native_status_reply(SOCKET socket,int backend) {
  char response[512];native_refresh_launch();LONG launch=InterlockedCompareExchange(&native_launch_state,0,0);
  // A second compatible launcher can win between our first probe and child exit.
  if(backend==0&&launch==3)backend=native_backend_state();
  const char *state=backend==1?"ready":backend<0?"conflict":launch==3?"failed":launch?"starting":"offline";
  if(backend==0&&launch==3&&native_exit_known)
    snprintf(response,sizeof(response),"{\"state\":\"failed\",\"exitCode\":%lu,\"pid\":%lu}",native_exit_code,native_process_pid);
  else if(backend==0&&launch==3)
    snprintf(response,sizeof(response),"{\"state\":\"failed\",\"errorCode\":%lu}",native_launch_error);
  else if(native_process_pid)
    snprintf(response,sizeof(response),"{\"state\":\"%s\",\"entryUrl\":\"http://127.0.0.1:%lu/wps/taskpane.html\",\"pid\":%lu}",state,native_backend_port,native_process_pid);
  else
    snprintf(response,sizeof(response),"{\"state\":\"%s\",\"entryUrl\":\"http://127.0.0.1:%lu/wps/taskpane.html\"}",state,native_backend_port);
  json_reply(socket,200,response);
}
static void native_launch_reply(SOCKET socket,BOOL explicit_retry) {
  native_refresh_launch();LONG state=InterlockedCompareExchange(&native_launch_state,0,0);
  if(state==1||state==2||(state==3&&!explicit_retry)){json_reply(socket,409,"{\"error\":\"already_requested\"}");return;}
  if(native_launch_thread) {
    if(WaitForSingleObject(native_launch_thread,0)!=WAIT_OBJECT_0){json_reply(socket,409,"{\"error\":\"already_requested\"}");return;}
    CloseHandle(native_launch_thread);native_launch_thread=0;
  }
  if(native_process){CloseHandle(native_process);native_process=0;}
  DWORD code=verify_native_runtime();
  if(code){diagnostic("runtime_verification",code);json_reply(socket,409,"{\"error\":\"runtime_unavailable\"}");return;}
  native_launch_error=0;native_process_pid=0;native_exit_known=FALSE;native_exit_code=0;
  InterlockedExchange(&native_launch_state,1);
  native_launch_thread=CreateThread(0,0,native_start_service,0,0,0);
  if(!native_launch_thread) {
    native_launch_error=GetLastError();InterlockedExchange(&native_launch_state,3);
    native_status_reply(socket,0);return;
  }
  json_reply(socket,202,"{\"state\":\"starting\"}");
}
static void native_launch_release(void) {
  // Join only on the server worker. A started shared backend outlives this DLL.
  if(native_launch_thread){WaitForSingleObject(native_launch_thread,INFINITE);CloseHandle(native_launch_thread);native_launch_thread=0;}
  if(native_process){CloseHandle(native_process);native_process=0;}
}

#endif
static void handle(SOCKET socket) {
  request_deadline=GetTickCount64()+3000;
  char request[8192],method[12],path[256],response[2048];int count=0;char *end=0;
  while(count<(int)sizeof(request)-1) {
    if(stop_server||GetTickCount64()>=request_deadline)return;
    int n=recv(socket,request+count,(int)sizeof(request)-count-1,0);if(n<=0)return;count+=n;request[count]=0;
    end=strstr(request,"\r\n\r\n");if(end)break;
  }
  if(!end||sscanf(request,"%11s %255s",method,path)!=2){json_reply(socket,400,"{\"error\":\"invalid_request\"}");return;}
  char *query=strchr(path,'?');if(query)*query=0;
  if(strcmp(header_value(request,"Host"),host)){json_reply(socket,403,"{\"error\":\"invalid_host\"}");return;}
  const char *sent_origin=header_value(request,"Origin");if(*sent_origin&&strcmp(sent_origin,origin)){json_reply(socket,403,"{\"error\":\"invalid_origin\"}");return;}
  BOOL authenticated=!strcmp(header_value(request,"X-OfficeBridge-Xll-Probe"),NATIVE_TOKEN);
#ifdef OFFICEBRIDGE_FULL_BOOTSTRAP
  if(!strcmp(path,"/native/health")&&!strcmp(method,"GET")) {
    json_reply(socket,200,"{\"service\":\"officebridge-wps-bootstrap\",\"status\":\"ok\"}");return;
  }
  if(!strcmp(path,"/native/bridge-status")&&!strcmp(method,"GET")) {
    native_status_reply(socket,native_backend_state());return;
  }
  if(!strcmp(path,"/native/start")&&!strcmp(method,"POST")) {
    if(strcmp(header_value(request,"Origin"),origin)||strcmp(header_value(request,"X-OfficeBridge-Probe-Token"),page_token)){json_reply(socket,403,"{\"error\":\"unauthorized\"}");return;}
    char *body=native_request_body(socket,request,end,count);
    if(!body||!officebridge_registry_strings_safe(body)){json_reply(socket,400,"{\"error\":\"invalid_request\"}");return;}
    cJSON *parsed=cJSON_ParseWithOpts(body,0,1);
    const cJSON *retry=parsed?cJSON_GetObjectItemCaseSensitive(parsed,"retry"):0;
    BOOL valid=parsed&&cJSON_IsObject(parsed)&&officebridge_registry_tree_safe(parsed,0)&&cJSON_GetArraySize(parsed)==1&&cJSON_IsBool(retry);
    BOOL explicit_retry=valid&&cJSON_IsTrue(retry);cJSON_Delete(parsed);
    if(!valid){json_reply(socket,400,"{\"error\":\"invalid_request\"}");return;}
    int state=native_backend_state();
    if(state==1){native_status_reply(socket,1);return;}
    if(state<0){json_reply(socket,409,"{\"error\":\"port_conflict\"}");return;}
    native_launch_reply(socket,explicit_retry);return;
  }
#endif

  if(!strcmp(path,"/probe/health")&&!strcmp(method,"GET")) {
    if(!authenticated){json_reply(socket,403,"{\"error\":\"unauthorized\"}");return;}
    snprintf(response,sizeof(response),"{\"service\":\"officebridge-native-xll-probe\",\"version\":6,\"mode\":\"in-process-thread\",\"pid\":%lu,\"registration\":%ld,\"activationError\":%ld,\"relayHttpStatus\":%ld,\"deployment\":%ld,\"invocation\":%ld,\"reads\":%ld,\"pageLoads\":%ld,\"address\":\"%s\",\"rows\":%d,\"columns\":%d}",GetCurrentProcessId(),registration,activation_error,relay_http,deployment,invocation,reads,page_loads,last_address,last_rows,last_columns);
    json_reply(socket,200,response);return;
  }
  if(!strcmp(path,"/probe/stop")&&!strcmp(method,"POST")) {
    if(!authenticated){json_reply(socket,403,"{\"error\":\"unauthorized\"}");return;}json_reply(socket,200,"{\"ok\":true}");InterlockedExchange(&stop_server,1);return;
  }
  if(!strcmp(path,"/probe/session")&&!strcmp(method,"GET")) {snprintf(response,sizeof(response),"{\"token\":\"%s\"}",page_token);json_reply(socket,200,response);return;}
  if(!strcmp(path,"/probe/read-result")&&!strcmp(method,"POST")) {
    if(strcmp(header_value(request,"Origin"),origin)||strcmp(header_value(request,"X-OfficeBridge-Probe-Token"),page_token)){json_reply(socket,403,"{\"error\":\"unauthorized\"}");return;}
    int length=atoi(header_value(request,"Content-Length"));char *body=end+4;int available=count-(int)(body-request);
    if(length<1||length>1024||(body-request)+length>=(int)sizeof(request)){json_reply(socket,400,"{\"error\":\"invalid_result\"}");return;}
    while(available<length){if(stop_server||GetTickCount64()>=request_deadline)return;int n=recv(socket,body+available,length-available,0);if(n<=0)return;available+=n;}body[length]=0;
    char address[32];int rows=0,columns=0,consumed=0;
    if(sscanf(body,"{\"address\":\"%31[A-Z0-9:]\",\"rows\":%d,\"columns\":%d}%n",address,&rows,&columns,&consumed)!=3||consumed!=length||!valid_address(address)||rows<1||rows>20||columns<1||columns>10){json_reply(socket,400,"{\"error\":\"invalid_result\"}");return;}
    strcpy(last_address,address);last_rows=rows;last_columns=columns;InterlockedIncrement(&reads);json_reply(socket,200,"{\"ok\":true}");return;
  }
  if(!strcmp(method,"GET"))for(unsigned int i=0;i<sizeof(native_assets)/sizeof(native_assets[0]);i++) {
    if(strcmp(path,native_assets[i].path))continue;
    HRSRC resource=FindResourceW(image,MAKEINTRESOURCEW(native_assets[i].id),RT_RCDATA);
    HGLOBAL loaded=resource?LoadResource(image,resource):0;void *data=loaded?LockResource(loaded):0;
    if(!data){json_reply(socket,500,"{\"error\":\"asset_unavailable\"}");return;}
    if(!strcmp(path,"/probe/taskpane.html"))InterlockedIncrement(&page_loads);
    reply(socket,200,native_assets[i].type,(const char*)data,(int)SizeofResource(image,resource));return;
  }
  json_reply(socket,404,"{\"error\":\"not_found\"}");
}
static DWORD run_server(BOOL verify_only,const WCHAR **stage) {
  *stage=L"random";
  HCRYPTPROV provider=0;BYTE random[32];
  if(!CryptAcquireContextW(&provider,0,0,PROV_RSA_AES,CRYPT_VERIFYCONTEXT))return GetLastError();
  if(!CryptGenRandom(provider,sizeof(random),random)){DWORD code=GetLastError();CryptReleaseContext(provider,0);return code;}
  CryptReleaseContext(provider,0);for(int i=0;i<32;i++)sprintf(page_token+i*2,"%02x",random[i]);
  *stage=L"winsock";WSADATA winsock;int startup=WSAStartup(MAKEWORD(2,2),&winsock);if(startup)return (DWORD)startup;
  *stage=L"socket";SOCKET listener=socket(AF_INET,SOCK_STREAM,IPPROTO_TCP);
  if(listener==INVALID_SOCKET){DWORD code=(DWORD)WSAGetLastError();WSACleanup();return code;}
  BOOL exclusive=TRUE;
  *stage=L"socket_options";
  if(setsockopt(listener,SOL_SOCKET,SO_EXCLUSIVEADDRUSE,(char*)&exclusive,sizeof(exclusive))){DWORD code=(DWORD)WSAGetLastError();closesocket(listener);WSACleanup();return code;}
  struct sockaddr_in address={0};address.sin_family=AF_INET;address.sin_port=htons((u_short)server_port);address.sin_addr.s_addr=htonl(INADDR_LOOPBACK);
  *stage=L"bind";
  if(bind(listener,(struct sockaddr*)&address,sizeof(address))){DWORD code=(DWORD)WSAGetLastError();closesocket(listener);WSACleanup();return code;}
  *stage=L"listen";
  if(listen(listener,16)){DWORD code=(DWORD)WSAGetLastError();closesocket(listener);WSACleanup();return code;}
  diagnostic("thread_listening",server_port);
  HANDLE activating=0;DWORD code=0;
  if(!verify_only&&!stop_server){activating=CreateThread(0,0,activate,0,0,0);if(!activating){*stage=L"activation_thread";code=GetLastError();}}
  while(!code&&!stop_server) {
    fd_set pending;FD_ZERO(&pending);FD_SET(listener,&pending);struct timeval timeout={0,250000};
    int selected=select(0,&pending,0,0,&timeout);
    if(selected==SOCKET_ERROR){*stage=L"select";code=(DWORD)WSAGetLastError();break;}if(!selected)continue;
    SOCKET client=accept(listener,0,0);if(client==INVALID_SOCKET)continue;
    DWORD milliseconds=1000;setsockopt(client,SOL_SOCKET,SO_RCVTIMEO,(char*)&milliseconds,sizeof(milliseconds));setsockopt(client,SOL_SOCKET,SO_SNDTIMEO,(char*)&milliseconds,sizeof(milliseconds));
    handle(client);closesocket(client);
  }
  closesocket(listener);WSACleanup();
  // This wait is on the server worker, never on the WPS UI or in DllMain.
  if(activating){WaitForSingleObject(activating,INFINITE);CloseHandle(activating);}
#ifdef OFFICEBRIDGE_FULL_BOOTSTRAP
  native_launch_release();native_runtime_release();
#endif
  return code;
}
static DWORD WINAPI server_thread(void *argument) {
  BOOL verify_only=argument==(void*)1;const WCHAR *stage=L"startup";
  DWORD code=run_server(verify_only,&stage);
  diagnostic(code?"thread_failed":"thread_stopped",code);
  if(code&&!verify_only&&!quiet_activation&&!stop_server){WCHAR message[384];swprintf(message,384,L"办公桥引导线程运行失败：%ls，错误码 %lu。没有修改表格。",stage,code);MessageBoxW(0,message,NATIVE_TITLE,MB_OK|MB_ICONERROR);}
  InterlockedExchange(&opening,0);
  FreeLibraryAndExitThread(image,code);
}
static int start_worker(int mode,DWORD test_port,DWORD test_budget,DWORD listen_port) {
  if(InterlockedCompareExchange(&opening,1,0)){if(stop_server){SetLastError(ERROR_BUSY);return 0;}return 1;}
  InterlockedExchange(&stop_server,0);InterlockedExchange(&deployment,0);InterlockedExchange(&invocation,0);InterlockedExchange(&reads,0);InterlockedExchange(&page_loads,0);
  InterlockedExchange(&registration,0);InterlockedExchange(&activation_error,0);InterlockedExchange(&relay_http,0);
  last_address[0]=0;last_rows=last_columns=0;
  quiet_activation=mode==2;relay_port=mode==2?test_port:58890;deployment_budget=mode==2?test_budget:120000;
  server_port=mode?listen_port:NATIVE_LISTEN_PORT;snprintf(host,sizeof(host),"127.0.0.1:%lu",server_port);snprintf(origin,sizeof(origin),"http://%s",host);
  #ifdef OFFICEBRIDGE_FULL_BOOTSTRAP
  InterlockedExchange(&native_launch_state,0);native_launch_error=0;native_process_pid=0;native_exit_known=FALSE;if(!mode)native_backend_port=NATIVE_BACKEND_PORT;
  #endif
  HMODULE retained=0;
  if(!GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,(LPCWSTR)&start_worker,&retained)){InterlockedExchange(&opening,0);return 0;}
  image=retained;
  HANDLE thread=CreateThread(0,0,server_thread,(void*)(INT_PTR)mode,0,0);
  if(!thread){DWORD code=GetLastError();InterlockedExchange(&opening,0);FreeLibrary(retained);SetLastError(code);return 0;}
  CloseHandle(thread);return 1;
}
__declspec(dllexport) int __cdecl OfficeBridgeProbeTestStart(void){return start_worker(1,0,0,43191);}
__declspec(dllexport) int __cdecl OfficeBridgeProbeTestStartPort(DWORD port){if(port<1||port>65535||port==58890){SetLastError(ERROR_INVALID_PARAMETER);return 0;}return start_worker(1,0,0,port);}
__declspec(dllexport) int __cdecl OfficeBridgeProbeTestStartRelay(DWORD port,DWORD budget,DWORD listen_port){if(port<1||port>65535||port==43191||port==58890||listen_port<1||listen_port>65535||listen_port==port||listen_port==58890||budget<100||budget>120000){SetLastError(ERROR_INVALID_PARAMETER);return 0;}return start_worker(2,port,budget,listen_port);}
#ifdef OFFICEBRIDGE_FULL_BOOTSTRAP
__declspec(dllexport) int __cdecl OfficeBridgeNativeTestStart(DWORD port,DWORD backend) {
  if(port<1||port>65535||port==58890||backend<1||backend>65535||backend==58890||port==backend||opening){SetLastError(ERROR_INVALID_PARAMETER);return 0;}
  native_backend_port=backend;return start_worker(1,0,0,port);
}
#endif
__declspec(dllexport) int __cdecl OfficeBridgeProbeTestState(void){return (int)InterlockedCompareExchange(&opening,0,0);}
__declspec(dllexport) int __cdecl xlAutoOpen(void) {
  int result=start_worker(0,0,0,NATIVE_LISTEN_PORT);
  if(!result){WCHAR message[192];swprintf(message,192,L"办公桥无法创建验证工作线程，错误码 %lu。",GetLastError());MessageBoxW(0,message,NATIVE_TITLE,MB_OK|MB_ICONERROR);}
  return result;
}
__declspec(dllexport) int __cdecl xlAutoAdd(void){return xlAutoOpen();}
__declspec(dllexport) int __cdecl xlAutoClose(void){InterlockedExchange(&stop_server,1);return 1;}
__declspec(dllexport) int __cdecl xlAutoRemove(void){InterlockedExchange(&stop_server,1);return 1;}
BOOL WINAPI DllMain(HINSTANCE instance,DWORD reason,LPVOID reserved){(void)reserved;if(reason==DLL_PROCESS_ATTACH){image=instance;DisableThreadLibraryCalls(instance);}return TRUE;}
