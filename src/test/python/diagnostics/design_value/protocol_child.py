"""Pure echo peer; no COM. Same JSON protocol over three local transport candidates."""
import argparse
import json
import os
from pathlib import Path
import socket
import sys
import time


def main():
    p=argparse.ArgumentParser();p.add_argument('--config',type=Path,required=True);args=p.parse_args()
    c=json.loads(args.config.read_text(encoding='utf-8-sig'))
    if c.get('launchProbe'):
        payload=json.loads(Path(c['input']).read_text(encoding='utf-8-sig'))
        print(json.dumps({'payload':payload,'python':sys.version,'executable':sys.executable,'cwd':os.getcwd()},ensure_ascii=True));return
    channel=c['channel'];address=None;resources=[]
    if channel=='anonymous':
        read=sys.stdin.buffer.readline;emit=lambda b:(sys.stdout.buffer.write(b),sys.stdout.buffer.flush())
    elif channel=='tcp':
        server=socket.socket();server.bind(('127.0.0.1',0));server.listen(1);address=server.getsockname()
    else:
        from multiprocessing.connection import Listener
        address=c['address'];server=Listener(address,family='AF_PIPE',authkey=None)
    Path(c['ready']).write_text(json.dumps({'pid':os.getpid(),'address':address}))
    if channel=='tcp':
        client,_=server.accept();server.close();stream=client.makefile('rwb');resources=[stream,client]
        read=stream.readline;emit=lambda b:(stream.write(b),stream.flush())
    elif channel=='named_pipe':
        client=server.accept();server.close();resources=[client]
        def read():
            try:return client.recv_bytes()
            except EOFError:return b''
        emit=client.send_bytes
    try:
        while True:
            line=read()
            if not line:break
            request=json.loads(line.decode('utf-8'));mode=c['mode']
            value={'requestId':request['requestId'],'outcome':'succeeded','data':request['arguments']}
            if mode=='exit':return
            if mode=='delay':time.sleep(3)
            if mode=='wrong_id':value['requestId']='foreign-request'
            if mode=='wrong_outcome':value['outcome']='maybe'
            if mode=='missing_data':del value['data']
            if mode=='extra_field':value['unsolicited']=True
            if mode=='nonfinite':value['data']={'nonfinite':float('nan')}
            raw=json.dumps(value,ensure_ascii=False,separators=(',',':')).encode('utf-8')+b'\n'
            if mode=='invalid_json':raw=b'{broken}\n'
            # Keep all other fields legal so a different check cannot hide a missing rejection.
            if mode=='duplicate_key':raw=b'{"requestId":'+json.dumps(request['requestId']).encode('utf-8')+b','+raw[1:]
            if mode=='bad_encoding':raw=b'\xff\xfe\n'
            if mode=='partial':raw=raw[:len(raw)//2]
            emit(raw)
            if mode!='positive':return
    finally:
        for r in resources:r.close()

if __name__=='__main__':main()
