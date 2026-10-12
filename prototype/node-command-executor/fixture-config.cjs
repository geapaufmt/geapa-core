'use strict';
const p=require('node:path'),{PROJECT:P,ROOT}=require('./sdk.cjs');
const DIR=__dirname,CORE='1M3GpuWwyRkV24nbHnopzj1q4J-Ea7zHv5eCG-ngOw0qR5lWDMvvv8UfD',BASE='projects/'+P+'/databases/(default)/documents';
module.exports={DIR,P,CORE,ROOT,BASE,save:(name,v)=>require('./session.cjs').save(name,v)};
