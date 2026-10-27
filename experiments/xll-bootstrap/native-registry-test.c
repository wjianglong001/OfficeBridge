/* Pure synthetic parser checks. No network, WPS, files or Agent access. */
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include "native-registry.h"

#define MATCH "{\"name\":\"OfficeBridgeXllProbe\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43191/probe/\",\"version\":\"5.0.0\"}"
#define NATIVE_NAME "OfficeBridgeNative"
#define NATIVE_URL "http://127.0.0.1:43192/bootstrap/"
#define NATIVE_VERSION "1.0.0"
#define NATIVE_MATCH "{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":\"1.0.0\"}"
#define NATIVE_EMPTY_VERSION "{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":\"\"}"
static unsigned int checks;
static void check(const char *label,const char *json,int expected) {
  int actual=officebridge_registry_state(json);checks++;
  if(actual!=expected){fprintf(stderr,"FAIL: %s (actual=%d expected=%d)\n",label,actual,expected);exit(1);}
}
static void check_native(const char *label,const char *json,int expected) {
  int actual=officebridge_registry_match(json,NATIVE_NAME,NATIVE_URL,NATIVE_VERSION);checks++;
  if(actual!=expected){fprintf(stderr,"FAIL: %s (actual=%d expected=%d)\n",label,actual,expected);exit(1);}
}
static void check_native_online(const char *label,const char *json,int expected) {
  int actual=officebridge_registry_match_online(json,NATIVE_NAME,NATIVE_URL,NATIVE_VERSION);checks++;
  if(actual!=expected){fprintf(stderr,"FAIL: %s (actual=%d expected=%d)\n",label,actual,expected);exit(1);}
}
static void check_parameters(const char *label,const char *name,const char *url,const char *version) {
  int actual=officebridge_registry_match("[" NATIVE_MATCH "]",name,url,version);checks++;
  if(actual!=-1){fprintf(stderr,"FAIL: %s (actual=%d expected=-1)\n",label,actual);exit(1);}
}
int main(void) {
  char buffer[4096],nested[256],oversized[65538];
  unsigned int i;
  check("null input",0,-1);
  check("empty input","",-1);
  check("missing","[]",0);
  check("unrelated record","[{\"name\":\"unrelated\",\"meta\":[1,true,null]}]",0);
  check("exact registration","[" MATCH "]",1);
  check("whitespace"," \r\n[" MATCH "]\t ",1);
  check("unknown optional metadata","[{\"name\":\"unrelated\",\"meta\":{\"anything\":\"value\"}}," MATCH "]",1);
  check("duplicate name records","[" MATCH "," MATCH "]",-1);
  check("root object",MATCH,-1);
  check("primitive entry","[null]",-1);
  check("record lacks name","[{}]",-1);
  check("name wrong type","[{\"name\":true}]",-1);
  check("case-sensitive name key","[{\"Name\":\"OfficeBridgeXllProbe\"}]",-1);
  check("case-sensitive name value","[{\"name\":\"officebridgexllprobe\"}]",0);
  check("duplicate foreign name before filtering","[{\"name\":\"unrelated\",\"name\":\"OfficeBridgeXllProbe\"}]",-1);
  check("decoded duplicate keys","[{\"name\":\"unrelated\",\"\\u006eame\":\"OfficeBridgeXllProbe\"}]",-1);
  check("nested duplicate keys","[{\"name\":\"unrelated\",\"meta\":{\"a\":1,\"a\":2}}]",-1);
  check("truncated input","[" MATCH,-1);
  check("trailing data","[" MATCH "]false",-1);
  check("trailing garbage","[" MATCH "] x",-1);
  check("non-JSON control whitespace","[\v]",-1);
  check("escaped NUL name","[{\"name\":\"OfficeBridgeXllProbe\\u0000other\"}]",-1);
  check("escaped NUL key","[{\"name\\u0000other\":\"unrelated\"}]",-1);
  check("escaped NUL unrelated value","[{\"name\":\"unrelated\",\"meta\":\"\\u0000\"}]",-1);
  check("escaped slash then NUL","[{\"name\":\"unrelated\",\"meta\":\"\\\\\\u0000\"}]",-1);
  check("literal slash-u text","[{\"name\":\"unrelated\",\"meta\":\"\\\\u0000\"}]",0);
  check("unknown escape","[{\"name\":\"bad\\x00\"}]",-1);
  check("short unicode","[{\"name\":\"bad\\u00\"}]",-1);
  check("unpaired surrogate","[{\"name\":\"bad\\uD800\"}]",-1);
  check("escaped exact name","[{\"name\":\"\\u004ffficeBridgeXllProbe\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43191/probe/\",\"version\":\"5.0.0\"}]",1);
  {
    const char *keys[]={"addonType","online","enable","url","version"};
    const char *good[]={"et","true","enable_dev","http://127.0.0.1:43191/probe/","5.0.0"};
    for(i=0;i<5;i++) {
      const char *start=strstr(MATCH,keys[i]);
      size_t prefix=(size_t)(start-MATCH);
      snprintf(buffer,sizeof(buffer),"[%.*s%s%s]",(int)prefix,MATCH,"unexpected",start+strlen(keys[i]));
      check("missing required exact key",buffer,-1);
      start=strstr(MATCH,good[i]);prefix=(size_t)(start-MATCH);
      snprintf(buffer,sizeof(buffer),"[%.*s%s%s]",(int)prefix,MATCH,"wrong",start+strlen(good[i]));
      check("mismatched exact field",buffer,-1);
    }
  }
  check("boolean online not accepted","[{\"name\":\"OfficeBridgeXllProbe\",\"addonType\":\"et\",\"online\":true,\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43191/probe/\",\"version\":\"5.0.0\"}]",-1);
  check("boolean enable not accepted","[{\"name\":\"OfficeBridgeXllProbe\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":true,\"url\":\"http://127.0.0.1:43191/probe/\",\"version\":\"5.0.0\"}]",-1);
  check("duplicate enabled key","[{\"name\":\"OfficeBridgeXllProbe\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"enable\":\"false\",\"url\":\"http://127.0.0.1:43191/probe/\",\"version\":\"5.0.0\"}]",-1);
  for(i=0;i<40;i++)nested[i]='[';
  nested[40]='0';for(i=0;i<40;i++)nested[41+i]=']';nested[81]=0;
  snprintf(buffer,sizeof(buffer),"[{\"name\":\"unrelated\",\"meta\":%s}]",nested);
  check("nesting limit",buffer,-1);
  memset(oversized,' ',sizeof(oversized)-1);oversized[0]='[';oversized[1]=']';oversized[sizeof(oversized)-1]=0;
  check("size defense",oversized,-1);
  check_native("native missing","[]",0);
  check_native("native exact","[" NATIVE_MATCH "]",1);
  check_native("probe entry does not satisfy native namespace","[" MATCH "]",0);
  check("native entry does not satisfy probe namespace","[" NATIVE_MATCH "]",0);
  check_native("native coexists with probe","[" MATCH "," NATIVE_MATCH "]",1);
  check("probe coexists with native","[" MATCH "," NATIVE_MATCH "]",1);
  check_native("duplicate native registrations","[" NATIVE_MATCH "," NATIVE_MATCH "]",-1);
  check_native("native wrong source URL","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43191/probe/\",\"version\":\"1.0.0\"}]",-1);
  check_native("native wrong version","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":\"5.0.0\"}]",-1);
  check_native("native disabled","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"false\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":\"1.0.0\"}]",-1);
  check_native("native boolean online refused","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":true,\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":\"1.0.0\"}]",-1);
  check_native("native escaped NUL refused","[{\"name\":\"OfficeBridgeNative\\u0000probe\"}]",-1);
  check_native("native decoded duplicate name key refused","[{\"name\":\"OfficeBridgeNative\",\"\\u006eame\":\"unrelated\"}]",-1);
  check_parameters("null expected name",0,NATIVE_URL,NATIVE_VERSION);
  check_parameters("empty expected name","",NATIVE_URL,NATIVE_VERSION);
  check_parameters("null expected URL",NATIVE_NAME,0,NATIVE_VERSION);
  check_parameters("empty expected URL",NATIVE_NAME,"",NATIVE_VERSION);
  check_parameters("null expected version",NATIVE_NAME,NATIVE_URL,0);
  check_parameters("empty expected version",NATIVE_NAME,NATIVE_URL,"");
  check_native_online("online exact version remains accepted","[" NATIVE_MATCH "]",1);
  check_native_online("online observed empty string version accepted","[" NATIVE_EMPTY_VERSION "]",1);
  check_native("strict native matcher still rejects empty version","[" NATIVE_EMPTY_VERSION "]",-1);
  check_native_online("online namespace ignores probe","[" MATCH "]",0);
  check_native_online("online empty version coexists with probe","[" MATCH "," NATIVE_EMPTY_VERSION "]",1);
  check_native_online("duplicate online empty version records rejected","[" NATIVE_EMPTY_VERSION "," NATIVE_EMPTY_VERSION "]",-1);
  check_native_online("wrong nonempty online version rejected","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":\"5.0.0\"}]",-1);
  check_native_online("missing online version rejected","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\"}]",-1);
  check_native_online("null online version rejected","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":null}]",-1);
  check_native_online("numeric online version rejected","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":0}]",-1);
  check_native_online("boolean online version rejected","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":false}]",-1);
  check_native_online("object online version rejected","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":{}}]",-1);
  check_native_online("online version NUL is not empty","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":\"\\u0000other\"}]",-1);
  check_native_online("duplicate online version keys rejected","[{\"name\":\"OfficeBridgeNative\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43192/bootstrap/\",\"version\":\"\",\"version\":\"1.0.0\"}]",-1);
  {
    const char *good[]={"et","true","enable_dev","http://127.0.0.1:43192/bootstrap/"};
    for(i=0;i<4;i++) {
      const char *start=strstr(NATIVE_EMPTY_VERSION,good[i]);
      size_t prefix=(size_t)(start-NATIVE_EMPTY_VERSION);
      snprintf(buffer,sizeof(buffer),"[%.*s%s%s]",(int)prefix,NATIVE_EMPTY_VERSION,"wrong",start+strlen(good[i]));
      check_native_online("empty version never relaxes another field",buffer,-1);
    }
  }
  check("probe wrapper still rejects empty version","[{\"name\":\"OfficeBridgeXllProbe\",\"addonType\":\"et\",\"online\":\"true\",\"enable\":\"enable_dev\",\"url\":\"http://127.0.0.1:43191/probe/\",\"version\":\"\"}]",-1);
  printf("PASS: %u synthetic registry checks; no WPS, network or Agent access.\n",checks);
  return 0;
}
