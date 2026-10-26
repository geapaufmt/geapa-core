// M63 adapter contract, local only. Google verification is not replaced by JWT decoding.
function f4MemberIdentityM13_(){
 const C=f4Core_(),N=C.need;
 function createLocal(options){
  C.exact(options,['config','decode','lookupToken','lookupAccount']);C.destination(options.config);
  N((options.config.environment==='SYNTHETIC_LOCAL' || f4OperationalDestination_(options.config))&&[options.decode,options.lookupToken,options.lookupAccount].every(v=>typeof v==='function'),'DESTINATION_INVALID');
  const projectId=options.config.projectId;
  function verify(token,now,enrollment=false){
   N(typeof token==='string'&&token.length>0&&token.length<=16384&&token.split('.').length===3&&Number.isSafeInteger(now),'AUTH_BINDING_INVALID');
   const jwt=options.decode(token);C.exact(jwt,['header','payload']);const h=jwt.header,p=jwt.payload;
   N(h&&h.alg==='RS256'&&C.id(h.kid)&&p&&p.iss==='https://securetoken.google.com/'+projectId&&p.aud===projectId&&C.id(p.sub)&&p.sub.length<=128,'AUTH_BINDING_INVALID');
   for(const field of ['iat','exp','auth_time'])N(Number.isSafeInteger(p[field])&&p[field]>0,'AUTH_BINDING_INVALID');
   N(p.auth_time<=p.iat&&p.iat*1000<=now&&now<p.exp*1000&&p.exp>p.iat,'AUTH_BINDING_INVALID');
   // This must be accounts:lookup with the ID token, not an unverified subject lookup.
   const verified=options.lookupToken(token,projectId);
   N(verified&&verified.status===200&&Array.isArray(verified.users)&&verified.users.length===1&&verified.users[0].localId===p.sub,'AUTH_BINDING_INVALID');
   const account=options.lookupAccount(p.sub,projectId);
   N(account&&account.status===200&&Array.isArray(account.users)&&account.users.length===1,'AUTH_BINDING_INVALID');
   const user=account.users[0];N(user.localId===p.sub&&user.disabled===false&&typeof user.validSince==='string'&&/^\d+$/.test(user.validSince),'AUTH_BINDING_INVALID');
   const since=Number(user.validSince);N(Number.isSafeInteger(since)&&p.auth_time>=since&&p.iat>=since,'ACCESS_REVOKED');
   if(enrollment){
    // The verified Google account supplies the email. A browser form or an
    // unverified JWT claim never binds an invitation to a person.
    N(user.emailVerified===true&&typeof user.email==='string'&&/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(user.email)&&verified.users[0].emailVerified===true&&typeof verified.users[0].email==='string'&&verified.users[0].email.toLowerCase()===user.email.toLowerCase(),'AUTH_BINDING_INVALID');
    return Object.freeze({uid:p.sub,expiresAt:p.exp*1000,projectId,email:user.email.trim().toLowerCase(),emailVerified:true});
   }
   return Object.freeze({uid:p.sub,expiresAt:p.exp*1000});
  }
  return Object.freeze({verify:(token,now)=>verify(token,now),verifyEnrollment:(token,now)=>verify(token,now,true)});
 }
 // No implicit real adapter/session. Physical D01 must prove the Google integration.
 return Object.freeze({createLocal});
}
