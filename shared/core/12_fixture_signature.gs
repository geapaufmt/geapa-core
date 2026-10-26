// Fixture signature verifier only. It does not validate Firebase production JWTs.
function f4Ed25519_() {
 if (f4Ed25519_.value_) return f4Ed25519_.value_;
 const P=(BigInt("1")<<BigInt("255"))-BigInt("19"),L=(BigInt("1")<<BigInt("252"))+BigInt("27742317777372353535851937790883648493");
 const mod=x=>((x%P)+P)%P;
 function pow(a,n){let r=BigInt("1");a=mod(a);while(n){if(n&BigInt("1"))r=mod(r*a);a=mod(a*a);n>>=BigInt("1")}return r}
 const D=mod(-BigInt("121665")*pow(BigInt("121666"),P-BigInt("2"))),SQRT=pow(BigInt("2"),(P-BigInt("1"))/BigInt("4")),ZERO=[BigInt("0"),BigInt("1"),BigInt("1"),BigInt("0")];
 function bytes(s){return Utilities.base64DecodeWebSafe(s).map(x=>x&255)}
 function number(a){let n=BigInt("0");for(let i=a.length-1;i>=0;i--)n=(n<<BigInt("8"))+BigInt(a[i]);return n}
 function point(a){if(a.length!==32)throw Error('TOKEN_SIGNATURE_INVALID');let y=number(a)&((BigInt("1")<<BigInt("255"))-BigInt("1")),sign=BigInt(a[31]>>7);if(y>=P)throw Error('TOKEN_SIGNATURE_INVALID');let y2=mod(y*y),xx=mod((y2-BigInt("1"))*pow(D*y2+BigInt("1"),P-BigInt("2"))),x=pow(xx,(P+BigInt("3"))/BigInt("8"));if(mod(x*x)!==xx)x=mod(x*SQRT);if(mod(x*x)!==xx||(x===BigInt("0")&&sign===BigInt("1")))throw Error('TOKEN_SIGNATURE_INVALID');if((x&BigInt("1"))!==sign)x=P-x;return [x,y,BigInt("1"),mod(x*y)]}
 function add(p,q){const a=mod((p[1]-p[0])*(q[1]-q[0])),b=mod((p[1]+p[0])*(q[1]+q[0])),c=mod(BigInt("2")*D*p[3]*q[3]),d=mod(BigInt("2")*p[2]*q[2]),e=b-a,f=d-c,g=d+c,h=b+a;return [mod(e*f),mod(g*h),mod(f*g),mod(e*h)]}
 function mul(p,n){let r=ZERO;while(n){if(n&BigInt("1"))r=add(r,p);p=add(p,p);n>>=BigInt("1")}return r}
 const eq=(a,b)=>mod(a[0]*b[2])===mod(b[0]*a[2])&&mod(a[1]*b[2])===mod(b[1]*a[2]);
 const B=point([88,...Array(31).fill(102)]),keys=new Map();
 function verify(message,key,signature){try{if(typeof key!=='string'||!/^[A-Za-z0-9_-]{43}$/.test(key))return false;const sig=bytes(signature),kb=bytes(key);if(sig.length!==64||kb.length!==32)return false;const s=number(sig.slice(32));if(s>=L)return false;let a=keys.get(key);if(!a){a=point(kb);if(eq(a,ZERO)||!eq(mul(a,L),ZERO))return false;keys.set(key,a)}const r=point(sig.slice(0,32));if(!eq(mul(r,L),ZERO))return false;const input=sig.slice(0,32).concat(kb,Utilities.newBlob(message).getBytes().map(x=>x&255));const h=number(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_512,input.map(x=>x>127?x-256:x)).map(x=>x&255))%L;return eq(mul(B,s),add(r,mul(a,h)))}catch{return false}}
 function decodeText(s){return Utilities.newBlob(Utilities.base64DecodeWebSafe(s)).getDataAsString('UTF-8')}
 return (f4Ed25519_.value_={verify,decodeText});
}
