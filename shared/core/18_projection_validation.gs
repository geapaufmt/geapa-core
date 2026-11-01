// F4A: native synchronous port of F3.5 shared/core/projection-validator.cjs. No external module loading.
function f4Validation_() {
  if (f4Validation_.value_) return f4Validation_.value_;
  return (f4Validation_.value_ = (function () {
'use strict';
const definitions=f4Schemas_().definitions;
function jsonValue(value){if(value?.toDate)return value.toDate().toISOString();if(Array.isArray(value))return value.map(jsonValue);if(value&&typeof value==='object')return Object.fromEntries(Object.entries(value).map(([k,v])=>[k,jsonValue(v)]));return value;}
function check(s,v){if(s.anyOf)return s.anyOf.some(x=>check(x,v));if(Object.prototype.hasOwnProperty.call(s,'const')&&v!==s.const)return false;if(s.enum&&!s.enum.includes(v))return false;if(s.type){const valid=s.type==='null'?v===null:s.type==='array'?Array.isArray(v):s.type==='integer'?Number.isSafeInteger(v):s.type==='object'?!!v&&typeof v==='object'&&!Array.isArray(v):typeof v===s.type;if(!valid)return false;}if(s.type==='object'){if(s.required.some(k=>!Object.prototype.hasOwnProperty.call(v,k)))return false;if(s.additionalProperties===false&&Object.keys(v).some(k=>!Object.prototype.hasOwnProperty.call(s.properties,k)))return false;for(const[k,x]of Object.entries(v))if(s.properties[k]&&!check(s.properties[k],x))return false;}if(s.type==='array'){if(v.length<(s.minItems??0)||v.length>(s.maxItems??Infinity)||!v.every(x=>check(s.items,x)))return false;}if(s.type==='string'){if(v.length>(s.maxLength??Infinity))return false;if(s.pattern&&!new RegExp(s.pattern).test(v))return false;if(s.format==='date-time'&&(!/^\d{4}-\d\d-\d\dT/.test(v)||!Number.isFinite(Date.parse(v))))return false;}return true;}
function validateProjection(kind,value){if(!definitions[kind]||!check(definitions[kind],jsonValue(value))){const {fail}=f4Core_();fail('SCHEMA_INVALID');}return value;}
return {validateProjection};


  })());
}

// Explicit version dispatcher for slice A. Never falls back from unknown schemas.
function f4ValidationM13_(){return {validateProjection:function(value,control){
 const K=f4MemberContracts_();
 if(Object.prototype.hasOwnProperty.call(K.versions,value?.schemaVersion)){f4MemberAuthority_().validate(control);
   if(value.schemaVersion==='safe-member-function-v13'&&value.scopeKey===null)K.need(Object.prototype.hasOwnProperty.call(control.nullableFunctionApprovals,value.idVigencia));
   if(value.schemaVersion==='member-function-index-v13')K.need(value.authorityEpoch===control.epoch);
 }
 if(['safe-member-function-v12','safe-member-function-v13'].includes(value?.schemaVersion))return K.dto(value);
 if(['member-function-index-v12','member-function-index-v13'].includes(value?.schemaVersion))return K.index(value);
 if(['member-function-proof-v12','member-function-proof-v13'].includes(value?.schemaVersion))return K.proof(value,control);
 K.need(false);
}};}
