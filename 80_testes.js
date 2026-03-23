function test_core_rolesConfig_debug() {
  Logger.log(JSON.stringify(core_debugInstitutionalRolesConfig_(), null, 2));
}

function test_core_rolesConfig_byKey() {
  Logger.log(JSON.stringify(
    core_findInstitutionalRoleByKey_('SECRETARIO_GERAL'),
    null,
    2
  ));
}

function test_core_rolesConfig_byPublicName() {
  Logger.log(JSON.stringify(
    core_findInstitutionalRoleByPublicName_('Secretário(a) Geral'),
    null,
    2
  ));
}

function test_core_rolesConfig_byAlias() {
  Logger.log(JSON.stringify(
    core_findInstitutionalRoleByAnyName_('Secretario Geral'),
    null,
    2
  ));
}

function test_core_rolesConfig_secretaria() {
  Logger.log(JSON.stringify(
    core_getInstitutionalRolesByEmailGroup_('SECRETARIA'),
    null,
    2
  ));
}

function test_core_rolesConfig_comunicacao() {
  Logger.log(JSON.stringify(
    core_getInstitutionalRolesByEmailGroup_('COMUNICACAO'),
    null,
    2
  ));
}

function test_core_rolesConfig_eventos() {
  Logger.log(JSON.stringify(
    core_getInstitutionalRolesByEmailGroup_('EVENTOS'),
    null,
    2
  ));
}

function test_core_rolesConfig_allActive() {
  Logger.log(JSON.stringify(
    core_getInstitutionalRolesActive_(),
    null,
    2
  ));
}

function test_core_rolesConfig_clearCacheAndDebug() {
  core_rolesConfigCacheClear_();
  Logger.log(JSON.stringify(core_debugInstitutionalRolesConfig_(), null, 2));
}

function test_projection_debug() {
  Logger.log(JSON.stringify(core_debugCurrentInstitutionalProjection_(), null, 2));
}

function test_projection_secretaria_html() {
  Logger.log(core_getCurrentContactsHtmlByEmailGroup_('SECRETARIA'));
}

function test_projection_sync_members() {
  Logger.log(JSON.stringify(core_syncMembersCurrentInstitutionalRoles_(), null, 2));
}

function test_projection_sync_members_and_debug() {
  Logger.log(JSON.stringify(core_syncMembersCurrentInstitutionalRoles_(), null, 2));
  Logger.log(JSON.stringify(core_debugCurrentInstitutionalProjection_(), null, 2));
}

function test_memberIdentity_byRga() {
  Logger.log(JSON.stringify(core_memberIdentityFindByAny_('SEU_RGA_AQUI'), null, 2));
}

function test_memberIdentity_byEmail() {
  Logger.log(JSON.stringify(core_memberIdentityFindByAny_('email@exemplo.com'), null, 2));
}

function test_sync_all_derived_fields() {
  Logger.log(JSON.stringify(core_syncMembersCurrentDerivedFields_(), null, 2));
}