'use strict';
// Server build/runtime selection only; never selected by HTTP payload or claims.
const profiles=Object.freeze({
 DEV:Object.freeze({environment:'DEV',projectId:'geapa-convergencia-dev',origin:'https://geapa-convergencia-dev.web.app',executor:'CF_GEN2_DEV_EDIT_V1',serviceAccount:'geapa-command-dev@geapa-convergencia-dev.iam.gserviceaccount.com',functionName:'geapaCommandDevPilot',prefix:'act-dev/'}),
 PROD:Object.freeze({environment:'PROD',projectId:'portal-geapa',projectNumber:'913092907585',origin:'https://portal-geapa.web.app',executor:'CF_GEN2_PROD_V1',serviceAccount:'geapa-command-prod@portal-geapa.iam.gserviceaccount.com',functionName:'geapaCommand',prefix:'act-prod/'})
});
const fs=require('fs'),path=require('path'),manifestPath=path.resolve(__dirname,'../release-manifest.json');
const locked=fs.existsSync(manifestPath)?JSON.parse(fs.readFileSync(manifestPath)).environment||'DEV':null;
function select(env=process.env){const key=env.GEAPA_TARGET_ENVIRONMENT||locked||'DEV';if(locked&&key!==locked)throw Object.assign(Error('DESTINATION_INVALID'),{code:'DESTINATION_INVALID'});if(!Object.hasOwn(profiles,key))throw Object.assign(Error('DESTINATION_INVALID'),{code:'DESTINATION_INVALID'});return profiles[key];}
module.exports={select,profile:select()};
