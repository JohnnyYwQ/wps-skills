"""Summarize only hash-verified experiment snapshots; keep raw measurements linked."""
from collections import Counter,defaultdict
import hashlib
import json
from pathlib import Path
import statistics


def stats(values,scale=1):
    values=[v*scale for v in values]
    return {'n':len(values),'median':statistics.median(values),'min':min(values),'max':max(values)}


def summarize(runs,output):
    summary={'scope':'Host-independent, instrumented Windows experiments. No Agent completion-rate claim.',
             'runs':[],'schema':{},'startup':{},'taskCost':{},'readbackCost':{},'protocol':{}}
    costs=defaultdict(list);readback=[]
    for root in runs:
        root=Path(root).resolve();collection=json.loads((root/'collection.json').read_text(encoding='utf-8-sig'))
        if collection.get('verified') is not True:raise ValueError('Unverified collection: '+str(root))
        target=root/'collected'
        manifest=json.loads((target/'files.json').read_text(encoding='utf-8-sig'))
        for name,expected in manifest.items():
            if hashlib.sha256((target/name).read_bytes()).hexdigest()!=expected:raise ValueError('Evidence changed: '+name)
        data=json.loads((target/'results/report.json').read_text(encoding='utf-8-sig'))
        states=Counter();responses=Counter()
        for path in (target/'results/calls').glob('*/response.json'):
            value=json.loads(path.read_text(encoding='utf-8-sig'));responses[value['outcome']]+=1
        for row in data['experiments']:
            observations=row.get('observations')
            state='not_applicable' if isinstance(observations,dict) and observations.get('state')=='not_applicable' else row['state']
            states[state]+=1
            if state!='passed':continue
            group,app,trial=row['id'].rsplit('-',2)
            if group=='schema':
                samples=observations['samples'];measured={}
                for mode in ('selected','all'):
                    rows=[s for s in samples if s['mode']==mode]
                    measured[mode]={'bytes':rows[0]['bytes'],'actionCount':rows[0]['actionCount'],'timeMs':stats([r['seconds'] for r in rows],1000)}
                measured['byteReductionPercent']=100*(1-measured['selected']['bytes']/measured['all']['bytes'])
                summary['schema'][app]=measured
            elif group=='startup' and app=='word':
                samples=observations['samples']
                for mode in ('direct','powershell'):summary['startup'][mode]=stats([r['seconds'] for r in samples if r['mode']==mode],1000)
                pairs=defaultdict(dict)
                for r in samples:pairs[r['trial']][r['mode']]=r['seconds']
                summary['startup']['pairedPowerShellMinusDirectMs']=stats([r['powershell']-r['direct'] for r in pairs.values()],1000)
            elif group=='task_cost':
                for r in observations:costs[(app,r['contentActions'])].append(r)
            elif group=='readback_cost' and app=='excel':
                readback.extend(dict(r,pairId=root.name+'-'+trial+'-'+str(r['pair'])) for r in observations)
            elif group=='protocol' and app=='word':
                rows=json.loads((target/'results'/observations['rawFile']).read_text(encoding='utf-8-sig'))
                result={'faultCases':sum(r['mode']!='positive' for r in rows),'unknownFaults':sum(r['mode']!='positive' and r['outcome']=='unknown' for r in rows),'allOwnedProcessesReleased':all(c['released'] for r in rows for c in r['cleanup']),'channels':{}}
                for channel in ('anonymous','tcp','named_pipe'):
                    result['channels'][channel]={}
                    for size in (32,4096):
                        samples=[r for r in rows if r['channel']==channel and r['messages']==20 and r['sizeChars']==size]
                        result['channels'][channel][str(size)]={'processes':len(samples),'verifiedMessages':sum(r['messages'] for r in samples),
                          'spawnReadyMs':stats([r['spawnReadySeconds'] for r in samples],1000),
                          'firstResponseMs':stats([r['exchangeSeconds'][0] for r in samples],1000),
                          'warmRoundTripPerProcessMedianMs':stats([statistics.median(r['exchangeSeconds'][1:]) for r in samples],1000)}
                summary['protocol'][root.name]=result
        summary['runs'].append({'run':root.name,'state':data['state'],'groups':dict(states),'taskResponseCalls':dict(responses),'verifiedFiles':collection['count'],'archiveSha256':collection['sha256'],'startedUtc':data['startedUtc'],'endedUtc':data['endedUtc'],'evidence':str(target)})
    for (app,count),rows in costs.items():
        summary['taskCost'].setdefault(app,{})[str(count)]={'wallSeconds':stats([r['wallSeconds'] for r in rows]),
          'receiptPublishMs':stats([r['phaseMs']['task.receipt_publish'] for r in rows]),'inclusiveActionMs':stats([r['phaseMs']['action.execute'] for r in rows])}
    if readback:
        for mode in (False,True):summary['readbackCost']['optimistic' if mode else 'native']=stats([r['seconds'] for r in readback if r['optimistic']==mode],1000)
        pairs=defaultdict(dict)
        for r in readback:pairs[r['pairId']][r['optimistic']]=r['seconds']
        summary['readbackCost']['pairedNativeMinusOptimisticMs']=stats([r[False]-r[True] for r in pairs.values()],1000)
    output=Path(output);output.parent.mkdir(parents=True,exist_ok=True)
    output.write_text(json.dumps(summary,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    print('Verified and summarized',len(runs),'runs:',output)
    return summary
