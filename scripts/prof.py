import json,glob,collections,sys
f=sys.argv[2] if len(sys.argv)>2 else sorted(glob.glob('scripts/out/*.cpuprofile'))[-1]
p=json.load(open(f))
nodes={n['id']:n for n in p['nodes']}
parent={}
for n in p['nodes']:
    for c in n.get('children',[]): parent[c]=n['id']
self=collections.Counter()
for s,d in zip(p['samples'],p['timeDeltas']): self[s]+=d
agg=collections.Counter()
for i,t in self.items():
    cf=nodes[i]['callFrame']; agg[(cf['functionName'],cf['url'].split('/')[-1],cf['lineNumber'])]+=t
print('--- self time')
for k,v in agg.most_common(int(sys.argv[1]) if len(sys.argv)>1 else 14): print(round(v/1000),'ms',k)
# inclusive time by function
incl=collections.Counter()
for i,t in self.items():
    seen=set(); cur=i
    while cur is not None:
        cf=nodes[cur]['callFrame']; k=(cf['functionName'],cf['url'].split('/')[-1],cf['lineNumber'])
        if k not in seen: incl[k]+=t; seen.add(k)
        cur=parent.get(cur)
print('--- inclusive')
for k,v in incl.most_common(24): print(round(v/1000),'ms',k)
