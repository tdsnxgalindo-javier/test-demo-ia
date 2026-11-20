(function (root, factory) {
  const core = factory();
  if (typeof module === "object" && module.exports) {
    module.exports = core;
  } else {
    root.DashboardCore = core;
  }
})(globalThis, function () {
  "use strict";

  class SynchronizationError extends Error {
    constructor(updated, documentId, cause) {
      super("Se interrumpio la sincronizacion de responsables.");
      this.name = "SynchronizationError";
      this.updated = updated;
      this.documentId = documentId;
      this.cause = cause instanceof Error ? cause : new Error(String(cause));
    }
  }

  function getEtag(item) {
    return item?.["odata.etag"] ||
      item?.["@odata.etag"] ||
      item?.__metadata?.etag ||
      item?.etag ||
      null;
  }

  function encodedODataLiteral(value) {
    const escaped = String(value).replace(/'/g, "''");
    const encoded = encodeURIComponent(escaped).replace(/[!'()*]/g, character =>
      `%${character.charCodeAt(0).toString(16).toUpperCase()}`
    );
    return `'${encoded}'`;
  }

  function unwrap(payload) {
    return payload?.d || payload;
  }

  function pageValues(payload) {
    const data = unwrap(payload);
    if (Array.isArray(data?.value)) return data.value;
    if (Array.isArray(data?.results)) return data.results;
    if (Array.isArray(data?.d?.results)) return data.d.results;
    return [];
  }

  function nextPageUrl(payload) {
    const data = unwrap(payload);
    return data?.["@odata.nextLink"] || data?.["odata.nextLink"] ||
      data?.d?.__next || data?.__next || null;
  }

  async function readAllPages(input, { requestPage, validateUrl }) {
    if (typeof requestPage !== "function" || typeof validateUrl !== "function") {
      throw new TypeError("La paginacion requiere requestPage y validateUrl.");
    }
    let next = validateUrl(input);
    const visited = new Set();
    const values = [];
    while (next) {
      const validated = validateUrl(next);
      if (visited.has(validated)) {
        throw new Error("SharePoint devolvio un enlace de paginacion repetido.");
      }
      visited.add(validated);
      const response = await requestPage(validated);
      const payload = response && Object.prototype.hasOwnProperty.call(response, "payload")
        ? response.payload
        : response;
      values.push(...pageValues(payload));
      const link = nextPageUrl(payload);
      next = link ? validateUrl(link) : null;
    }
    return values;
  }

  function accountFromLogin(raw) {
    if (typeof raw !== "string") return "";
    const value = raw.trim();
    if (!value) return "";
    return value.slice(value.lastIndexOf("|") + 1).trim();
  }

  function normalizedAccount(account) {
    return accountFromLogin(account).toLocaleLowerCase("en-US");
  }

  function groupClaimGuid(loginName) {
    if (typeof loginName !== "string") return null;
    const match = loginName.match(/^c:0o\.c\|federateddirectoryclaimprovider\|\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?$/i);
    return match ? match[1] : null;
  }

  function firstString(record, keys) {
    for (const key of keys) {
      const value = record?.[key];
      if (typeof value === "string" && value.trim()) return value.trim();
    }
    return "";
  }

  function mapDirectoryMember(person, source) {
    const accountValue = firstString(person, [
      "principalName", "PrincipalName", "loginName", "LoginName", "mail", "Mail", "Email"
    ]);
    const account = accountFromLogin(accountValue);
    if (!account) {
      throw new Error("El grupo contiene una persona sin cuenta utilizable; no se ofrecera.");
    }
    const displayName = firstString(person, [
      "displayName", "DisplayName", "Title", "title", "Name", "name"
    ]);
    if (!displayName) {
      throw new Error(`La cuenta ${account} no tiene nombre visible; no se ofrecera.`);
    }
    return { account, normalized: normalizedAccount(account), displayName, source };
  }

  async function loadMembers({ baseUrl, readAllPages: readPages }) {
    const usersUrl = baseUrl +
      "web/AssociatedMemberGroup/Users?$select=Id,Title,LoginName,PrincipalType&$top=200";
    const associatedUsers = await readPages(usersUrl);
    const people = [];
    for (const principal of associatedUsers) {
      const principalType = Number(principal.PrincipalType);
      const loginName = firstString(principal, ["LoginName", "loginName", "Name", "name"]);
      if (principalType === 1) {
        people.push(mapDirectoryMember({
          LoginName: loginName,
          Title: firstString(principal, ["Title", "title"]),
          Mail: firstString(principal, ["Email", "email", "Mail", "mail"])
        }, "member"));
        continue;
      }
      const groupId = groupClaimGuid(loginName);
      if (groupId) {
        const groupUrl = baseUrl +
          `SP.Directory.DirectorySession/Group('${groupId}')/Members?$select=PrincipalName,DisplayName,Mail&$top=200`;
        const groupMembers = await readPages(groupUrl);
        if (!groupMembers.length) {
          throw new Error(`El grupo Microsoft 365 ${groupId} no contiene miembros enumerables.`);
        }
        for (const member of groupMembers) {
          people.push(mapDirectoryMember(member, "m365-group"));
        }
        continue;
      }
      throw new Error(`Tipo de principal no soportado en el grupo de miembros (PrincipalType ${principal.PrincipalType}).`);
    }

    const unique = new Map();
    for (const person of people) {
      const key = normalizedAccount(person.account);
      if (!key) throw new Error("El grupo devolvio una cuenta vacia.");
      if (!unique.has(key)) unique.set(key, person);
    }
    const members = [...unique.values()].sort((a, b) =>
      a.displayName.localeCompare(b.displayName, "es", { sensitivity: "base" })
    );
    if (!members.length) {
      throw new Error("El grupo de miembros no devolvio personas asignables.");
    }
    return members;
  }

  function requireAssignableMember(account, members) {
    const member = members.find(item =>
      item.normalized === normalizedAccount(account)
    );
    if (!member) {
      throw new Error("La persona seleccionada ya no pertenece al grupo de miembros; no se guardo el cambio.");
    }
    return member;
  }

  function lookupId(value) {
    if (value === null || value === undefined || value === "") return null;
    const id = Number(value);
    if (!Number.isSafeInteger(id) || id <= 0) {
      throw new Error("SharePoint devolvio un ID de responsable no valido.");
    }
    return id;
  }

  async function synchronizeDocumentAssignments({
    opportunities,
    documentsUrl,
    readAllPages: readPages,
    mergeDocument,
    readDocument
  }) {
    const documents = await readPages(documentsUrl);
    const opportunityById = new Map(opportunities.map(opportunity => [
      opportunity.id,
      opportunity
    ]));
    let updated = 0;

    for (const document of documents) {
      const documentId = Number(document.Id ?? document.ID);
      if (!Number.isSafeInteger(documentId) || documentId <= 0) {
        throw new SynchronizationError(updated, document.Id ?? document.ID ?? "desconocido",
          new Error("El documento no devolvio un ID valido."));
      }
      const opportunityId = Number(document.OportunidadId);
      const opportunity = opportunityById.get(opportunityId);
      if (!opportunity) {
        throw new SynchronizationError(
          updated,
          documentId,
          new Error(`No se encontro la oportunidad ${document.OportunidadId} del documento.`)
        );
      }
      let documentAssignee;
      let sourceAssignee;
      try {
        documentAssignee = lookupId(document.AsignadoAId);
        sourceAssignee = lookupId(opportunity.assigneeId);
      } catch (error) {
        throw new SynchronizationError(updated, documentId, error);
      }
      if (documentAssignee === sourceAssignee) continue;

      const etag = getEtag(document);
      if (!etag) {
        throw new SynchronizationError(
          updated,
          documentId,
          new Error("El documento no devolvio ETag; no se escribio.")
        );
      }
      try {
        await mergeDocument(documentId, { AsignadoAId: sourceAssignee }, etag);
        updated += 1;
        const confirmed = await readDocument(
          documentId,
          "$select=Id,OportunidadId,AsignadoAId"
        );
        const confirmedAssignee = lookupId(confirmed.AsignadoAId);
        if (confirmedAssignee !== sourceAssignee) {
          throw new Error("La lectura posterior no confirma la asignacion esperada.");
        }
      } catch (error) {
        throw new SynchronizationError(updated, documentId, error);
      }
    }
    return { total: documents.length, updated };
  }

  function validateDocumentFile(file, maxBytes, maxNameLength) {
    if (!file) throw new Error("Selecciona un archivo.");
    if (!Number.isSafeInteger(file.size) || file.size <= 0) {
      throw new Error("El archivo debe tener contenido y un tamaño valido.");
    }
    if (file.size > maxBytes) {
      throw new Error("El archivo supera el limite de 20 MiB.");
    }
    const name = file.name;
    if (!name || name.length > maxNameLength) {
      throw new Error(`El nombre del archivo es obligatorio y admite hasta ${maxNameLength} caracteres.`);
    }
    if (/["*:<>?\/\\|]/.test(name) || /[\u0000-\u001f]/.test(name) ||
        /^[. ]|[. ]$/.test(name)) {
      throw new Error("El nombre contiene caracteres no permitidos o empieza/termina con punto o espacio.");
    }
    return name;
  }

  function validateDocumentReview(status, observations, statuses) {
    if (!statuses.includes(status)) {
      throw new Error("Selecciona un estado de validacion valido.");
    }
    const text = String(observations ?? "");
    if (text.length > 2000) {
      throw new Error("Las observaciones admiten hasta 2000 caracteres.");
    }
    if (status === "Rechazado" && !text.trim()) {
      throw new Error("Indica el motivo del rechazo en Observaciones.");
    }
    return { status, observations: text.trim() };
  }

  function buildDocumentReviewFields(status, observations, statuses, reviewerAccount, validatedAt) {
    const review = validateDocumentReview(status, observations, statuses);
    const fields = {
      EstadoValidacion: review.status,
      Observaciones: review.observations
    };
    if (review.status === "Pendiente") {
      fields.ValidadoPor = null;
      fields.FechaValidacion = null;
      return fields;
    }

    const account = accountFromLogin(reviewerAccount);
    if (!account) throw new Error("Falta la cuenta actual del revisor.");
    const date = validatedAt ? new Date(validatedAt) : new Date();
    if (!Number.isFinite(date.getTime())) throw new Error("La fecha de revision no es valida.");
    fields.ValidadoPor = account;
    fields.FechaValidacion = date.toISOString();
    return fields;
  }

  async function createDocumentWithCurrentAssignment({
    opportunityId,
    readOpportunity,
    createDocument
  }) {
    const opportunity = await readOpportunity(opportunityId);
    if (!opportunity) throw new Error(`No se encontro la oportunidad ${opportunityId}.`);
    const assigneeId = lookupId(opportunity.AsignadoAId);
    return createDocument(assigneeId);
  }

  async function uploadAttachmentWithCompensation(documentId, upload, compensate) {
    try {
      await upload(documentId);
      return { uploaded: true, cleanupConfirmed: null, error: null, cleanupError: null };
    } catch (error) {
      try {
        const cleanup = await compensate(documentId);
        return {
          uploaded: false,
          cleanupConfirmed: cleanup?.confirmed === true,
          error,
          cleanupError: cleanup?.error || null
        };
      } catch (cleanupError) {
        return {
          uploaded: false,
          cleanupConfirmed: false,
          error,
          cleanupError
        };
      }
    }
  }

  async function fetchAttachmentBlob(url, { accessToken, fetchImpl, onHttpError }) {
    const response = await fetchImpl(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        Accept: "application/octet-stream"
      },
      credentials: "omit",
      redirect: "error"
    });
    if (!response.ok) {
      throw await onHttpError(response);
    }
    return response.blob();
  }

  return {
    SynchronizationError,
    accountFromLogin,
    buildDocumentReviewFields,
    createDocumentWithCurrentAssignment,
    encodedODataLiteral,
    fetchAttachmentBlob,
    firstString,
    getEtag,
    groupClaimGuid,
    loadMembers,
    mapDirectoryMember,
    normalizedAccount,
    readAllPages,
    requireAssignableMember,
    synchronizeDocumentAssignments,
    uploadAttachmentWithCompensation,
    validateDocumentFile,
    validateDocumentReview
  };
});
