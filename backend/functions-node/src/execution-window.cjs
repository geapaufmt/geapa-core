'use strict';
// Reuse the Core window contract, never a second authorization implementation.
const {createDomain}=require('./domain.cjs');
const domain=createDomain(()=>{throw Error('WINDOW_UNEXPECTED_IO');});
const windowLive=(value,now)=>JSON.parse(domain.run('(x)=>f4ExecutorWindow_().live(x.value,x.now)',{value,now}));
module.exports={windowLive};
