function f4Schemas_() { return {
  "$schema": "https://json-schema.org/draft/2020-12/schema",
  "scope": "Somente contratos locais de projeções server-only; validação do futuro produtor obrigatória antes de qualquer escrita",
  "timestampRepresentation": "ISO para validação documental; Timestamp nativo no Firestore",
  "definitions": {
    "portal": {
      "type": "object",
      "properties": {
        "schemaVersion": {
          "const": "portal-authz-v1"
        },
        "uid": {
          "type": "string"
        },
        "idPessoa": {
          "type": "string"
        },
        "ativo": {
          "type": "boolean"
        },
        "podeAcessarPortal": {
          "type": "boolean"
        },
        "stale": {
          "type": "boolean"
        },
        "bindingVersion": {
          "type": "integer"
        },
        "accessEpoch": {
          "type": "integer"
        },
        "sourceRevision": {
          "type": "integer"
        },
        "policyVersion": {
          "type": "integer"
        },
        "startsAt": {
          "type": "string",
          "format": "date-time"
        },
        "validUntil": {
          "type": "string",
          "format": "date-time"
        },
        "technicalAccount": {
          "type": "boolean"
        }
      },
      "required": [
        "schemaVersion",
        "uid",
        "idPessoa",
        "ativo",
        "podeAcessarPortal",
        "stale",
        "bindingVersion",
        "accessEpoch",
        "sourceRevision",
        "policyVersion",
        "startsAt",
        "validUntil",
        "technicalAccount"
      ],
      "additionalProperties": false
    },
    "status": {
      "type": "object",
      "properties": {
        "schemaVersion": {
          "const": "safe-status-v1"
        },
        "ownerUid": {
          "type": "string"
        },
        "state": {
          "type": "string"
        },
        "egressEndsAtExclusive": {
          "anyOf": [
            {
              "type": "string",
              "format": "date-time"
            },
            {
              "type": "null"
            }
          ]
        }
      },
      "required": [
        "schemaVersion",
        "ownerUid",
        "state",
        "egressEndsAtExclusive"
      ],
      "additionalProperties": false
    },
    "access": {
      "type": "object",
      "properties": {
        "schemaVersion": {
          "const": "portal-access-status-v1"
        },
        "status": {
          "enum": [
            "ACTIVE",
            "BINDING_REQUIRED",
            "STALE",
            "SUSPENDED",
            "EGRESS",
            "REVOKED"
          ]
        },
        "updatedAt": {
          "type": "string",
          "format": "date-time"
        },
        "validUntil": {
          "type": "string",
          "format": "date-time"
        }
      },
      "required": [
        "schemaVersion",
        "status",
        "updatedAt",
        "validUntil"
      ],
      "additionalProperties": false
    },
    "justification": {
      "type": "object",
      "properties": {
        "schemaVersion": {
          "const": "safe-own-justification-v1"
        },
        "ownerUid": {
          "type": "string"
        },
        "justificationId": {
          "type": "string"
        },
        "activityId": {
          "type": "string"
        },
        "status": {
          "enum": [
            "PREVIA",
            "ENVIADA",
            "PENDENTE",
            "EM_ANALISE",
            "AJUSTE_SOLICITADO",
            "DEFERIDA",
            "ABONADA",
            "INDEFERIDA",
            "CANCELADA"
          ]
        },
        "submittedAt": {
          "type": "string",
          "format": "date-time"
        },
        "updatedAt": {
          "type": "string",
          "format": "date-time"
        },
        "reason": {
          "type": "string",
          "maxLength": 8000
        },
        "analysisSummary": {
          "type": "string",
          "maxLength": 500
        }
      },
      "required": [
        "schemaVersion",
        "ownerUid",
        "justificationId",
        "activityId",
        "status",
        "submittedAt",
        "updatedAt",
        "reason",
        "analysisSummary"
      ],
      "additionalProperties": false
    },
    "publicPage": {
      "type": "object",
      "properties": {
        "id": {
          "type": "string"
        },
        "schemaVersion": {
          "const": "geapa-public-v2.1"
        },
        "published": {
          "const": true
        },
        "sanitized": {
          "const": true
        },
        "title": {
          "type": "string",
          "maxLength": 240
        },
        "subtitle": {
          "type": "string"
        },
        "sections": {
          "type": "array",
          "items": {
            "type": "object",
            "properties": {
              "id": {
                "type": "string"
              },
              "type": {
                "enum": [
                  "TEXT",
                  "HERO",
                  "CTA",
                  "IMAGE_TEXT",
                  "DOCUMENT_LIST"
                ]
              },
              "title": {
                "type": "string"
              },
              "subtitle": {
                "type": "string"
              },
              "text": {
                "type": "string"
              },
              "imageUrl": {
                "type": "string"
              },
              "buttonText": {
                "type": "string"
              },
              "buttonUrl": {
                "type": "string"
              },
              "order": {
                "type": "integer"
              },
              "published": {
                "const": true
              }
            },
            "required": [
              "id",
              "type",
              "title",
              "subtitle",
              "text",
              "imageUrl",
              "buttonText",
              "buttonUrl",
              "order",
              "published"
            ],
            "additionalProperties": false
          },
          "minItems": 1,
          "maxItems": 30
        },
        "sourceHash": {
          "type": "string"
        },
        "sourceUpdatedAt": {
          "type": "string",
          "format": "date-time"
        },
        "updatedAt": {
          "type": "string",
          "format": "date-time"
        },
        "revision": {
          "type": "integer"
        }
      },
      "required": [
        "id",
        "schemaVersion",
        "published",
        "sanitized",
        "title",
        "subtitle",
        "sections",
        "sourceHash",
        "sourceUpdatedAt",
        "updatedAt",
        "revision"
      ],
      "additionalProperties": false
    }
  }
}; }

// Separate registry: old definitions are byte/behavior compatible, M13 is explicit opt-in.
function f4SchemasM13_(){const K=f4MemberContracts_();return {scope:'LOCAL_CONTRACTS_ONLY',versions:K.versions};}
