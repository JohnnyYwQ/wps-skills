"""Bounded local-channel and process-layer experiments. No WPS calls."""
import json
from pathlib import Path
import socket
import subprocess
import sys
import time
import uuid

from wps_skills.core.action_runtime import ControllerContext
from wps_skills.windows.bridge_types import BackendActionFailure
from wps_skills.windows.owned_process import WindowsOwnedProcessLauncher
from wps_skills.windows.powershell_bridge import JsonLineBridgeTransport,PowerShellBridge

MODES=['positive','invalid_json','duplicate_key','nonfinite','wrong_id','wrong_outcome',
       'missing_data','extra_field','bad_encoding','exit','delay','partial']


class NetworkInput:
    def __init__(self,connection,channel):self.connection,self.channel=connection,channel
    def write(self,text):
        raw=text.encode('utf-8')
        if self.channel=='tcp':self.connection.sendall(raw)
        else:self.connection.send_bytes(raw)
        return len(text)
    def flush(self):pass
    def close(self):
        if self.channel=='tcp':self.connection.shutdown(socket.SHUT_WR)
        else:self.connection.close()


class NamedOutput:
    def __init__(self,connection):self.connection=connection
    def readline(self):
        try:return self.connection.recv_bytes().decode('utf-8')
        except EOFError:return ''


class ConnectedProcess:
    def __init__(self,process,connection,channel):
        self.process=process;self.stdin=NetworkInput(connection,channel)
        self.stdout=connection.makefile('r',encoding='utf-8',errors='strict',newline='\n') if channel=='tcp' else NamedOutput(connection)
    def poll(self):return self.process.poll()
    def wait(self,timeout=None):return self.process.wait(timeout=timeout)


def exchange_case(experiment,channel,mode,label,*,messages=1,size=32):
    from .runner import require,write,read
    directory=experiment.root/'protocol'/label;directory.mkdir(parents=True,exist_ok=False)
    config={'channel':channel,'mode':mode,'ready':str(directory/'ready.json'),
            'address':'\\\\.\\pipe\\wps-design-'+uuid.uuid4().hex}
    write(directory/'config.json',config)
    launcher=WindowsOwnedProcessLauncher(cleanup_timeout_seconds=3,graceful_timeout_seconds=.3,terminate_timeout_seconds=.3)
    process=None;connection=None;transport=None
    started=time.perf_counter();result={}
    try:
        process=launcher.start_bridge([sys.executable,str(experiment.kit/'design_value/protocol_child.py'),'--config',str(directory/'config.json')])
        end=time.monotonic()+10
        while not (directory/'ready.json').exists():
            require(process.poll() is None and time.monotonic()<end,'Echo peer did not become ready',channel=channel,mode=mode)
            time.sleep(.01)
        ready=read(directory/'ready.json');spawn_seconds=time.perf_counter()-started
        target=process
        if channel=='tcp':connection=socket.create_connection(tuple(ready['address']),timeout=3);connection.settimeout(None)
        elif channel=='named_pipe':
            from multiprocessing.connection import Client
            connection=Client(ready['address'],family='AF_PIPE',authkey=None)
        if connection is not None:target=ConnectedProcess(process,connection,channel)
        transport=JsonLineBridgeTransport(process=target,liveness_timeout_seconds=2,close_timeout_seconds=.5)
        bridge=PowerShellBridge(transport=transport);times=[]
        for index in range(messages):
            payload=('中文😀é\\引号"\n'*((size//9)+1))[:size]
            arguments={'payload':payload,'index':index,'empty':None,'nested':{'flag':True}}
            context=ControllerContext(uuid.uuid4().hex,'design-protocol',time.monotonic()+(.3 if mode=='delay' else 3))
            call=time.perf_counter()
            try:
                value=bridge.execute('echo',arguments,context)
                require(mode=='positive' and value==arguments,'Invalid response accepted or payload changed',mode=mode,value=value)
                outcome='succeeded';code=None
            except BackendActionFailure as error:
                require(mode!='positive' and error.outcome=='unknown' and error.code=='RESPONSE_LOST','Protocol failure mapping incorrect',mode=mode,outcome=error.outcome,code=error.code)
                outcome=error.outcome;code=error.code
            times.append(time.perf_counter()-call)
        result={'channel':channel,'mode':mode,'sizeChars':size,'messages':messages,'outcome':outcome,'code':code,
                'spawnReadySeconds':spawn_seconds,'exchangeSeconds':times,'peerPid':process.pid}
    finally:
        if transport is not None:result['gracefulTransportClose']=transport.close()
        cleanup=launcher.close()
        if connection is not None:connection.close()
        result['cleanup']=[{'pid':c.pid,'released':c.released,'steps':list(c.cleanup_steps)} for c in cleanup]
        result['totalSeconds']=time.perf_counter()-started
        write(directory/'result.json',result)
        require(all(c.released for c in cleanup),'Protocol process cleanup failed',result=result)
    require(result['totalSeconds']<15,'Protocol case exceeded declared bound',result=result)
    return result


def protocol(experiment,app,label):
    from .runner import write
    if app!='word':return {'state':'not_applicable','reason':'Shared transport measured once, no application dependency.'}
    rows=[]
    for channel in ('anonymous','tcp','named_pipe'):
        for mode in MODES:
            rows.append(exchange_case(experiment,channel,mode,label+'-'+channel+'-'+mode))
    # Alternate channel order across paired rounds; separate first-response and warm exchange samples.
    for trial in range(10):
        channels=['anonymous','tcp','named_pipe'];channels=channels[trial%3:]+channels[:trial%3]
        for channel in channels:
            for size in (32,4096):rows.append(exchange_case(experiment,channel,'positive',label+'-cost-'+str(trial)+'-'+channel+'-'+str(size),messages=20,size=size))
    write(experiment.root/'measurements'/(label+'.json'),rows)
    return {'cases':len(rows),'faultCases':3*(len(MODES)-1),'channelCandidates':['anonymous','tcp','named_pipe'],
            'rawFile':'measurements/'+label+'.json','nativeWps':False,'serverLanguage':'same Python for all channels'}


def startup(experiment,app,label):
    from .runner import require,write
    if app!='word':return {'state':'not_applicable','reason':'Process launch independent of application.'}
    directory=experiment.root/'measurements'/(label+'_中文 空格');directory.mkdir(parents=True)
    payload={'内容':'中文与 emoji 😀','path':'C:\\离线 文件\\任务.json','nested':[True,None,3.5]}
    write(directory/'input.json',payload)
    write(directory/'config.json',{'launchProbe':True,'input':str(directory/'input.json')})
    direct=[sys.executable,str(experiment.kit/'design_value/protocol_child.py'),'--config',str(directory/'config.json')]
    wrapped=[experiment.ps,'-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-File',str(experiment.kit/'resources/outer_python.ps1'),
             '-Python',sys.executable,'-Script',str(experiment.kit/'design_value/protocol_child.py'),'-Config',str(directory/'config.json')]
    rows=[]
    for trial in range(20):
        for mode in (('direct','powershell') if trial%2==0 else ('powershell','direct')):
            start=time.perf_counter();cp=subprocess.run(direct if mode=='direct' else wrapped,capture_output=True,text=True,encoding='utf-8',timeout=15,cwd=directory)
            elapsed=time.perf_counter()-start
            require(cp.returncode==0,'Launch probe failed',mode=mode,stderr=cp.stderr)
            observed=json.loads(cp.stdout)
            require(observed['payload']==payload and observed['executable']==sys.executable,'Launch changed input or interpreter',observed=observed)
            rows.append({'trial':trial,'mode':mode,'seconds':elapsed,'observation':observed})
    write(experiment.root/'measurements'/(label+'.json'),rows)
    return {'pairs':20,'samples':rows,'nativeWps':False,'inputEncoding':'UTF-8 JSON file; Unicode and spaces in input path'}
