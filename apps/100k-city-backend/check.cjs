const fs=require('fs'),vm=require('vm'),assert=require('assert');
const html=fs.readFileSync(fs.existsSync(__dirname+'/100k-city.html')?__dirname+'/100k-city.html':__dirname+'/../100k-city.html','utf8');
const js=html.match(/<script>([\s\S]*)<\/script>/)[1];new vm.Script(js);
const c=vm.createContext({console,crypto:require('crypto').webcrypto,Math,Date});
vm.runInContext(js.split('/* ============ 街ビュー')[0],c);
function run(s){return vm.runInContext(s,c);}
run('newGame()');
assert(run('populationModel(G,40).natural<0'));
assert(run('(()=>{const b={...G,IND:90,FAM:85,INF:85,WEL:85,DX:80,serviceLag:80};return populationModel(b,85).social>populationModel({...b,INF:20},85).social})()'));
assert(run('(()=>{G.FUND=1;G.DEBT=200;applyFx({fund:-4});return G.FUND===0&&G.DEBT===203})()'));
assert(run('(()=>{newGame();recordDecision("event",{eventId:0});return G.decisions.length===0})()'));
const strategies={industry:[.65,.08,.08,.12,.05,.02],balanced:[.22,.24,.23,.18,.10,.03],care:[.12,.18,.18,.42,.08,.02]};
for(const [name,weights]of Object.entries(strategies)){
 const results=[];
 for(let seed=1;seed<=100;seed++){
 run(`newGame();G.rand=rngf(${seed});G.cyclePhase=${seed%6};`);
 for(let y=0;y<20;y++){
 run(`(()=>{const p=preview().policy;const keys=['ind','fam','inf','wel','dx','promo'];const a=Object.fromEntries(keys.map((k,i)=>[k,Math.floor(p*${JSON.stringify(weights)}[i])]));step(a,null,'full');})()`);
 assert(run('Number.isFinite(G.POP)&&G.POP>0&&Number.isFinite(G.FUND)'));
 if(!run('G.alive'))break;
 }
 results.push(run('G.POP'));
 }
 console.log(name,Math.round(results.reduce((a,b)=>a+b)/results.length),Math.min(...results),Math.max(...results));
}
console.log('Syntax, demographic dependencies, paid event costs, opt-out and 300 trajectories passed.');
assert(run(`(()=>{newGame();G.FUND=50;G.commitments=[{label:'test',years:2,cost:2,fx:{fam:5}}];const a={ind:10,fam:10,inf:8,wel:8,dx:5,promo:5};step(a,null,'full');if(G.commitments.length!==1||G.commitments[0].years!==1)return false;step(a,null,'full');return G.commitments.length===0;})()`));
// Every event choice must execute without NaN, including renovation with an empty reserve.
run(`YEAR_EVENTS.forEach((ev,id)=>ev.ch.forEach(choice=>{newGame();G.year=10;G.FUND=0;G.renew={hall:'degraded'};const first=Object.keys(BUILDINGS)[0];G.renew={[first]:'degraded'};applyFx(choice.fx);if(choice.run)choice.run();if(!Number.isFinite(G.FUND)||!Number.isFinite(G.DEBT))throw Error('event '+id);}));`);
new vm.Script(fs.readFileSync(__dirname+'/Code.gs','utf8'));
const backend=vm.createContext({console,Date,Set,JSON});
vm.runInContext(fs.readFileSync(__dirname+'/Code.gs','utf8'),backend);
const valid={version:'city-3.1',noticeVersion:'2026-09-v2',id:require('crypto').randomUUID(),demographics:{age:'na',gender:'na'},result:{year:1},decisions:[{type:'execution',year:0,allocation:{ind:1,fam:1,inf:1,wel:1,dx:1,promo:1}}]};
valid.playedAt=new Date().toISOString();
backend.payload=valid;vm.runInContext('validate(payload)',backend);
assert.equal(vm.runInContext("fiscalYear('2026-03-31T14:59:59Z')",backend),2025);
assert.equal(vm.runInContext("fiscalYear('2026-03-31T15:00:00Z')",backend),2026);
backend.payload={...valid,playedAt:'invalid'};assert.throws(()=>vm.runInContext('validate(payload)',backend));
backend.payload={...valid,noticeVersion:'invalid'};assert.throws(()=>vm.runInContext('validate(payload)',backend));
backend.payload={...valid,decisions:[...valid.decisions,...valid.decisions]};assert.throws(()=>vm.runInContext('validate(payload)',backend));
console.log('All event choices, delayed commitments and backend consent/duplicate-year validation passed.');
run(`newGame(); var ctx={P:57,fi:preview().fi,prevWel:8,alloc:{ind:10,fam:10,inf:8,wel:8,dx:5,promo:5}};`);
assert(run('!tallyVotes(judgeBudget(ctx)).pass'),'Default budget must require negotiation');
assert(run('JSON.stringify(judgeBudget(ctx))===JSON.stringify(judgeBudget(ctx))'),'Preview must be stable');
assert(run(`(()=>{const p=amendmentPlan(ctx,'fukushi','spread');return p&&judgeBudget({...ctx,alloc:p}).fukushi.ok&&Object.values(p).reduce((a,b)=>a+b,0)<=57;})()`));
assert(run(`amendmentPlan({...ctx,P:10},'fukushi','promo')===null`));
assert(run(`(()=>{const a={ind:8,fam:10,inf:16,wel:14,dx:4,promo:0};return tallyVotes(judgeBudget({...ctx,alloc:a})).pass;})()`),'A feasible welfare/local coalition must exist');
assert(run(`(()=>{const before=JSON.stringify(G.bias);judgeBudget(ctx);judgeBudget(ctx);return before===JSON.stringify(G.bias);})()`));
console.log('Council rejection, feasible coalition, funded amendment and deterministic previews passed.');
backend.rows=[['header'],...Array.from({length:40},(_,i)=>[new Date(),String(i),'city-3.1','na','na',JSON.stringify({playedAt:i<20?'2025-06-01T00:00:00Z':'2026-06-01T00:00:00Z',decisions:valid.decisions})])];
vm.runInContext('sheet=()=>({getDataRange:()=>({getValues:()=>rows})});json=v=>v;',backend);
const yearly=vm.runInContext("doGet({parameter:{year:'2025'}})",backend);
assert.equal(yearly.count,20);assert.equal(yearly.annual.length,2);assert.equal(yearly.annual[0].year,2026);
assert(vm.runInContext("doGet({parameter:{year:'2024'}}).hidden",backend));
console.log('Japan fiscal-year boundary, year filtering and annual comparison passed.');
