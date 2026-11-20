"use strict";

const test = require("node:test");
const assert = require("node:assert/strict");
const core = require("../dashboard-core.js");

const baseUrl = "https://tenant.example/sites/demo/_api/";

function makeUrlValidator() {
  return input => {
    const url = new URL(input, baseUrl);
    if (url.origin !== new URL(baseUrl).origin ||
        !url.pathname.startsWith(new URL(baseUrl).pathname)) {
      throw new Error("URL fuera del sitio permitido.");
    }
    return url.href;
  };
}

function makePagedReader(pages) {
  const calls = [];
  const readAllPages = input => core.readAllPages(input, {
    validateUrl: makeUrlValidator(),
    requestPage: async url => {
      calls.push(url);
      const payload = pages.get(url);
      if (!payload) throw new Error(`Pagina no simulada: ${url}`);
      return { payload };
    }
  });
  return { calls, readAllPages };
}

test("miembros directos y grupo M365 se expanden, paginan y deduplican", async () => {
  const groupId = "11111111-2222-4333-8444-555555555555";
  const usersUrl = `${baseUrl}web/AssociatedMemberGroup/Users?$select=Id,Title,LoginName,PrincipalType&$top=200`;
  const groupUrl = `${baseUrl}SP.Directory.DirectorySession/Group('${groupId}')/Members?$select=PrincipalName,DisplayName,Mail&$top=200`;
  const groupNext = `${groupUrl}&$skiptoken=page-2`;
  const { calls, readAllPages } = makePagedReader(new Map([
    [usersUrl, { value: [
      { PrincipalType: 1, LoginName: "i:0#.f|membership|direct@example.test", Title: "Directa" },
      { PrincipalType: 4, LoginName: `c:0o.c|federateddirectoryclaimprovider|{${groupId}}`, Title: "Equipo" }
    ] }],
    [groupUrl, {
      value: [{ PrincipalName: "member@example.test", DisplayName: "Miembro" }],
      "@odata.nextLink": groupNext
    }],
    [groupNext, { value: [
      { PrincipalName: "MEMBER@example.test", DisplayName: "Duplicado" },
      { PrincipalName: "second@example.test", DisplayName: "Segundo" }
    ] }]
  ]));

  const members = await core.loadMembers({ baseUrl, readAllPages });
  assert.deepEqual(members.map(member => member.normalized), [
    "direct@example.test",
    "member@example.test",
    "second@example.test"
  ]);
  assert.equal(members.filter(member => member.normalized === "member@example.test").length, 1);
  assert.equal(members.find(member => member.normalized === "direct@example.test").source, "member");
  assert.equal(members.find(member => member.normalized === "member@example.test").source, "m365-group");
  assert.equal(calls.some(url => url.includes("siteusers") || url.includes("ensureuser")), false);
});

test("miembros rechaza principales ajenos y tipos de grupo no soportados", async () => {
  const usersUrl = `${baseUrl}web/AssociatedMemberGroup/Users?$select=Id,Title,LoginName,PrincipalType&$top=200`;
  const { readAllPages } = makePagedReader(new Map([
    [usersUrl, { value: [{ PrincipalType: 8, LoginName: "c:0t.c|tenant|unsupported" }] }]
  ]));
  await assert.rejects(
    core.loadMembers({ baseUrl, readAllPages }),
    /Tipo de principal no soportado/
  );
  assert.throws(
    () => core.requireAssignableMember("outsider@example.test", [
      core.mapDirectoryMember({ Mail: "member@example.test", Title: "Miembro" }, "member")
    ]),
    /no pertenece al grupo/
  );
});

test("paginacion sigue nextLink del mismo sitio y rechaza enlaces externos", async () => {
  const firstUrl = `${baseUrl}web/lists/items?$top=1`;
  const secondUrl = `${baseUrl}web/lists/items?$skiptoken=2`;
  const pages = new Map([
    [firstUrl, { value: [{ Id: 1 }], "@odata.nextLink": secondUrl }],
    [secondUrl, { value: [{ Id: 2 }] }]
  ]);
  const { calls, readAllPages } = makePagedReader(pages);
  assert.deepEqual(await readAllPages(firstUrl), [{ Id: 1 }, { Id: 2 }]);
  assert.equal(calls.length, 2);

  const external = makePagedReader(new Map([
    [firstUrl, { value: [{ Id: 1 }], "@odata.nextLink": "https://attacker.example/next" }]
  ]));
  await assert.rejects(external.readAllPages(firstUrl), /URL fuera del sitio permitido/);
});

test("ETag se extrae de los formatos REST usados y no se inventa", () => {
  assert.equal(core.getEtag({ "@odata.etag": '"v2"' }), '"v2"');
  assert.equal(core.getEtag({ __metadata: { etag: '"v1"' } }), '"v1"');
  assert.equal(core.getEtag({}), null);
  assert.equal(core.encodedODataLiteral("client's #1%.pdf"), "'client%27%27s%20%231%25.pdf'");
});

test("sincronizacion escribe solo diferencias, usa ETags exactos y limpia null", async () => {
  const documents = [
    { Id: 1, OportunidadId: 10, AsignadoAId: 5, Title: "uno.pdf", Observaciones: "keep", "@odata.etag": '"doc-1"' },
    { Id: 2, OportunidadId: 10, AsignadoAId: 7, Title: "dos.pdf", Observaciones: "keep", "@odata.etag": '"doc-2"' },
    { Id: 3, OportunidadId: 20, AsignadoAId: 99, Title: "historico.pdf", Observaciones: "keep", "@odata.etag": '"doc-3"' }
  ];
  const opportunities = [{ id: 10, assigneeId: 7 }, { id: 20, assigneeId: null }];
  const writes = [];
  const options = {
    opportunities,
    documentsUrl: `${baseUrl}web/lists/docs/items`,
    readAllPages: async () => documents,
    mergeDocument: async (id, fields, etag) => {
      writes.push({ id, fields, etag });
      Object.assign(documents.find(item => item.Id === id), fields);
    },
    readDocument: async id => documents.find(item => item.Id === id)
  };

  const first = await core.synchronizeDocumentAssignments(options);
  assert.deepEqual(first, { total: 3, updated: 2 });
  assert.deepEqual(writes, [
    { id: 1, fields: { AsignadoAId: 7 }, etag: '"doc-1"' },
    { id: 3, fields: { AsignadoAId: null }, etag: '"doc-3"' }
  ]);
  assert.equal(documents[0].Title, "uno.pdf");
  assert.equal(documents[0].Observaciones, "keep");
  assert.equal(documents[2].AsignadoAId, null);

  assert.deepEqual(await core.synchronizeDocumentAssignments(options), { total: 3, updated: 0 });
  assert.equal(writes.length, 2);
});

test("fallo parcial se detiene y el reintento omite lo ya sincronizado", async () => {
  const documents = [
    { Id: 11, OportunidadId: 10, AsignadoAId: 1, "@odata.etag": '"first"' },
    { Id: 12, OportunidadId: 10, AsignadoAId: 2, "@odata.etag": '"second"' }
  ];
  const writes = [];
  let failSecond = true;
  const options = {
    opportunities: [{ id: 10, assigneeId: 7 }],
    documentsUrl: `${baseUrl}web/lists/docs/items`,
    readAllPages: async () => documents,
    mergeDocument: async (id, fields, etag) => {
      writes.push({ id, etag });
      if (id === 12 && failSecond) {
        failSecond = false;
        throw new Error("HTTP 412");
      }
      Object.assign(documents.find(item => item.Id === id), fields);
    },
    readDocument: async id => documents.find(item => item.Id === id)
  };

  await assert.rejects(
    core.synchronizeDocumentAssignments(options),
    error => error instanceof core.SynchronizationError &&
      error.updated === 1 && error.documentId === 12 &&
      /HTTP 412/.test(error.cause.message)
  );
  assert.deepEqual(writes, [{ id: 11, etag: '"first"' }, { id: 12, etag: '"second"' }]);

  const retried = await core.synchronizeDocumentAssignments(options);
  assert.deepEqual(retried, { total: 2, updated: 1 });
  assert.deepEqual(writes, [
    { id: 11, etag: '"first"' },
    { id: 12, etag: '"second"' },
    { id: 12, etag: '"second"' }
  ]);
});

test("carga lee la asignacion mas reciente antes de crear metadatos", async () => {
  const order = [];
  let savedAssignee;
  const response = await core.createDocumentWithCurrentAssignment({
    opportunityId: 54,
    readOpportunity: async id => {
      order.push(`read:${id}`);
      return { Id: id, AsignadoAId: null };
    },
    createDocument: async assigneeId => {
      order.push("create");
      savedAssignee = assigneeId;
      return { Id: 91 };
    }
  });
  assert.deepEqual(order, ["read:54", "create"]);
  assert.equal(savedAssignee, null);
  assert.deepEqual(response, { Id: 91 });
});

test("validaciones documentales aplican limites inclusivos y motivo de rechazo", () => {
  assert.equal(core.validateDocumentFile({
    name: `${"a".repeat(176)}.pdf`,
    size: 20 * 1024 * 1024
  }, 20 * 1024 * 1024, 180).length, 180);
  assert.throws(() => core.validateDocumentFile({
    name: "x.pdf",
    size: 20 * 1024 * 1024 + 1
  }, 20 * 1024 * 1024, 180), /20 MiB/);
  assert.throws(() => core.validateDocumentFile({
    name: `${"a".repeat(177)}.pdf`,
    size: 1
  }, 20 * 1024 * 1024, 180), /180 caracteres/);
  assert.throws(() => core.validateDocumentFile({ name: "bad?.pdf", size: 1 }, 100, 180));
  assert.throws(() => core.validateDocumentFile({ name: "empty.pdf", size: 0 }, 100, 180));

  assert.throws(
    () => core.validateDocumentReview("Rechazado", "  ", ["Pendiente", "Validado", "Rechazado"]),
    /motivo/
  );
  assert.throws(
    () => core.validateDocumentReview("Validado", "x".repeat(2001), ["Pendiente", "Validado", "Rechazado"]),
    /2000 caracteres/
  );
  assert.deepEqual(
    core.validateDocumentReview("Pendiente", " nota ", ["Pendiente", "Validado", "Rechazado"]),
    { status: "Pendiente", observations: "nota" }
  );
  assert.deepEqual(
    core.buildDocumentReviewFields("Pendiente", " nota ", ["Pendiente", "Validado", "Rechazado"], ""),
    {
      EstadoValidacion: "Pendiente",
      Observaciones: "nota",
      ValidadoPor: null,
      FechaValidacion: null
    }
  );
  assert.deepEqual(
    core.buildDocumentReviewFields(
      "Validado",
      "",
      ["Pendiente", "Validado", "Rechazado"],
      "i:0#.f|membership|reviewer@example.test",
      "2026-11-18T18:00:00.000Z"
    ),
    {
      EstadoValidacion: "Validado",
      Observaciones: "",
      ValidadoPor: "reviewer@example.test",
      FechaValidacion: "2026-11-18T18:00:00.000Z"
    }
  );
});

test("fallo de binario compensa solo el registro provisional y reporta limpieza", async () => {
  let deleted = 0;
  const failedUpload = await core.uploadAttachmentWithCompensation(
    91,
    async id => { assert.equal(id, 91); throw new Error("fallo binario"); },
    async id => { deleted = id; return { confirmed: true }; }
  );
  assert.equal(deleted, 91);
  assert.equal(failedUpload.uploaded, false);
  assert.equal(failedUpload.cleanupConfirmed, true);
  assert.match(failedUpload.error.message, /fallo binario/);

  const failedCleanup = await core.uploadAttachmentWithCompensation(
    92,
    async () => { throw new Error("fallo binario"); },
    async () => ({ confirmed: false, error: new Error("registro presente") })
  );
  assert.equal(failedCleanup.cleanupConfirmed, false);
  assert.match(failedCleanup.cleanupError.message, /registro presente/);

  const uploaded = await core.uploadAttachmentWithCompensation(
    93,
    async () => {},
    async () => { throw new Error("no debe compensar"); }
  );
  assert.equal(uploaded.uploaded, true);
  assert.equal(uploaded.cleanupConfirmed, null);
});

test("descarga autenticada conserva los bytes recibidos", async () => {
  const expected = Uint8Array.from([0, 1, 2, 255, 13, 10]);
  let requestOptions;
  const blob = await core.fetchAttachmentBlob("/attachment", {
    accessToken: "mock-token",
    fetchImpl: async (url, options) => {
      assert.equal(url, "/attachment");
      requestOptions = options;
      return new Response(expected, { status: 200 });
    },
    onHttpError: async () => new Error("No esperado")
  });
  assert.equal(requestOptions.headers.Authorization, "Bearer mock-token");
  assert.equal(requestOptions.headers.Accept, "application/octet-stream");
  assert.deepEqual([...new Uint8Array(await blob.arrayBuffer())], [...expected]);
});
