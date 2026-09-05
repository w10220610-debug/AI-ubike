from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit.components.v1 as components


SERVER_SYNC_COMPONENT_HTML = r'''<!doctype html><html><head><meta charset="utf-8"><style>html,body{width:1px;height:1px;margin:0;overflow:hidden}</style></head><body>
<button id="serverSyncButton" type="button" hidden>更新</button>
<script>
(()=>{
 const API=1; let args={}; let timer=null; let busy=false;
 function send(type,data={}){window.parent.postMessage({isStreamlitMessage:true,type,...data},"*");}
 function setValue(value){send("streamlit:setComponentValue",{value});}
 function setHeight(){send("streamlit:setFrameHeight",{height:1});}
 function host(state,detail={}){window.parent.postMessage({source:"ubike-server-sync",type:"ubike:sync-state",state,...detail},"*");}
 function schedule(){if(timer!==null)clearTimeout(timer);if(!args.auto_refresh)return;const ms=Math.max(15,Number(args.auto_refresh_seconds||60))*1000;timer=setTimeout(()=>trigger(false),ms);}
 function trigger(force){if(busy)return;busy=true;host("busy");setValue({action:"refresh",force:Boolean(force),nonce:Date.now()});setTimeout(()=>{busy=false;schedule();},45000);}
 document.getElementById("serverSyncButton").onclick=()=>trigger(true);
 window.addEventListener("message",event=>{
   if(!event.data)return;
   if(event.data.type==="ubike:manual-sync"){trigger(true);return;}
   if(event.data.type!=="streamlit:render")return;
   args=event.data.args||{};busy=false;
   if(args.last_status==="success")host("success",{station_count:Number(args.station_count||0)});
   else if(args.last_status==="error")host("error",{message:String(args.error_message||"請稍後再試")});
   schedule();setHeight();
 });
 send("streamlit:componentReady",{apiVersion:API});setHeight();
})();
</script></body></html>'''


_SERVER_SYNC_COMPONENT = None


def get_server_sync_component():
    global _SERVER_SYNC_COMPONENT
    if _SERVER_SYNC_COMPONENT is not None:
        return _SERVER_SYNC_COMPONENT
    component_dir = Path(tempfile.gettempdir()) / "youbike_server_sync_component_v29"
    component_dir.mkdir(parents=True, exist_ok=True)
    index_path = component_dir / "index.html"
    try:
        if not index_path.exists() or index_path.read_text(encoding="utf-8") != SERVER_SYNC_COMPONENT_HTML:
            index_path.write_text(SERVER_SYNC_COMPONENT_HTML, encoding="utf-8")
    except OSError:
        pass
    _SERVER_SYNC_COMPONENT = components.declare_component(
        "youbike_server_sync_v29",
        path=str(component_dir),
    )
    return _SERVER_SYNC_COMPONENT


def render_server_sync_trigger(
    *,
    key: str,
    last_status: str = "",
    station_count: int = 0,
    error_message: str = "",
    auto_refresh_seconds: int = 60,
):
    component = get_server_sync_component()
    return component(
        key=key,
        default=None,
        auto_refresh=True,
        auto_refresh_seconds=max(15, int(auto_refresh_seconds)),
        last_status=last_status,
        station_count=max(0, int(station_count)),
        error_message=str(error_message or ""),
    )
