function test_core_rolesConfig_debug() {
  Logger.log(JSON.stringify(core_debugInstitutionalRolesConfig_(), null, 2));
}

function test_projection_debug() {
  Logger.log(JSON.stringify(core_debugCurrentInstitutionalProjection_(), null, 2));
}

function test_sync_all_derived_fields() {
  Logger.log(JSON.stringify(core_syncMembersCurrentDerivedFields_(), null, 2));
}