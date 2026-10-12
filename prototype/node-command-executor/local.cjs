'use strict';
const current=require('./test-support/current/local.cjs'),{ROOT}=require('./sdk.cjs');
module.exports={...current,runtime:changes=>current.runtime({'coreDevValidation/ACT_FUNCTIONAL_20261008':ROOT,...changes})};
