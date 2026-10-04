/**
 * PAC → OJS Submission Proxy
 * Cloudflare Worker — Phase 1
 *
 * Accepts a POST from index.html (hosted on GitHub Pages) containing
 * PAC form data, creates a bare submission in OJS, adds authors, and
 * returns the OJS author workflow URL so the author can finish there.
 *
 * Deployment: paste into Cloudflare Dashboard → Workers & Pages → Create → Worker.
 * DO NOT hardcode secrets here — set them in the Worker's Settings → Variables:
 *
 * Secrets (encrypted, set in Cloudflare dashboard):
 *   OJS_TOKEN       — OJS API token for the PAC admin/service account
 *
 * Plain-text environment variables (set in Cloudflare dashboard):
 *   ALLOWED_ORIGIN  — e.g. https://sgarrettroe.github.io
 *   OJS_BASE_URL    — https://pac.pogil.org/index.php/pac/api/v1
 *
 * This file contains no secrets and is safe to commit to a public repository.
 */

// Map PAC form review_kind / activity_idea readiness → OJS sectionId
const SECTION_IDS = {
  traditional_review:          13,
  collaborative_peer_feedback: 12,
  classroom_testing:           15,
  activity_idea:                3,
};

export default {
  async fetch(request, env) {

    const origin = request.headers.get('Origin') || '';
    const allowedOrigin = env.ALLOWED_ORIGIN || '';

    // ── CORS preflight ──────────────────────────────────────────────
    if (request.method === 'OPTIONS') {
      return corsResponse(null, 204, origin, allowedOrigin);
    }

    // ── Only accept POST ────────────────────────────────────────────
    if (request.method !== 'POST') {
      return corsResponse({ error: 'Method not allowed' }, 405, origin, allowedOrigin);
    }

    // ── Parse PAC form payload ──────────────────────────────────────
    let pac;
    try {
      pac = await request.json();
    } catch {
      return corsResponse({ error: 'Invalid JSON body' }, 400, origin, allowedOrigin);
    }

    const base    = env.OJS_BASE_URL;
    const headers = {
      'Authorization': `Bearer ${env.OJS_TOKEN}`,
      'Content-Type':  'application/json',
    };

    // ── Determine sectionId from review_kind (or activity_idea) ────
    const reviewKind = pac.review_kind || '';
    const readiness  = pac.readiness   || '';
    const isIdea     = readiness === 'activity_idea';
    const sectionId  = isIdea
      ? SECTION_IDS.activity_idea
      : (SECTION_IDS[reviewKind] || SECTION_IDS.traditional_review);

    const title = pac.title || '(Untitled)';

    // ── Step 1: Create bare submission (OJS 3.3 does NOT auto-create ──
    // a publication when you POST to /submissions — that changed in 3.4.
    // We create the submission first, then create the publication separately.)
    let submissionRes, submission;
    try {
      submissionRes = await fetch(`${base}/submissions`, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          locale: 'en_US',
          sectionId,
          submissionProgress: 1,
        }),
      });
      submission = await submissionRes.json();
    } catch (e) {
      return corsResponse({ error: 'OJS unreachable', detail: e.message }, 502, origin, allowedOrigin);
    }

    if (!submissionRes.ok) {
      return corsResponse({
        error:  'OJS rejected submission creation',
        status: submissionRes.status,
        detail: submission,
      }, 502, origin, allowedOrigin);
    }

    const submissionId = submission.id;
    if (!submissionId) {
      return corsResponse({ error: 'OJS response missing submission id', detail: submission }, 502, origin, allowedOrigin);
    }

    // ── Step 2: Create publication for the submission ─────────────────
    // In OJS 3.3, POST /submissions does not create a publication.
    let pubRes, publication;
    try {
      pubRes = await fetch(`${base}/submissions/${submissionId}/publications`, {
        method: 'POST',
        headers,
        body: JSON.stringify({
          locale:   'en_US',
          sectionId,
          title:    { en_US: title },
          abstract: { en_US: buildAbstract(pac) },
        }),
      });
      publication = await pubRes.json();
    } catch (e) {
      return corsResponse({ error: 'OJS unreachable creating publication', detail: e.message }, 502, origin, allowedOrigin);
    }

    if (!pubRes.ok) {
      return corsResponse({
        error:  'OJS rejected publication creation',
        status: pubRes.status,
        detail: publication,
      }, 502, origin, allowedOrigin);
    }

    const publicationId = publication.id;
    if (!publicationId) {
      return corsResponse({ error: 'OJS response missing publication id', detail: publication }, 502, origin, allowedOrigin);
    }

    // ── Step 3: Add authors ───────────────────────────────────────────
    // OJS 3.3 uses bracket notation for locale fields, e.g. "givenName[en]",
    // NOT the object form {"givenName": {"en_US": "..."}} used in 3.4+.
    const authors = pac.authors || [];
    const authorErrors = [];

    for (let i = 0; i < authors.length; i++) {
      const a = authors[i];
      if (!a.author_name && !a.author_email) continue;

      const nameParts = splitName(a.author_name || '');
      const authorPayload = {
        'givenName[en_US]':  nameParts.given,
        'familyName[en_US]': nameParts.family,
        email:               a.author_email || '',
        seq:                 i,
        primaryContact:      i === 0,
        userGroupId:         31,  // PAC OJS "Author" role (confirmed from existing submissions)
      };

      const authorRes = await fetch(
        `${base}/submissions/${submissionId}/publications/${publicationId}/contributors`,
        { method: 'POST', headers, body: JSON.stringify(authorPayload) }
      );

      if (!authorRes.ok) {
        const err = await authorRes.json().catch(() => ({}));
        authorErrors.push({ author: a.author_name, status: authorRes.status, detail: err });
      }
    }

    // ── Step 3: Return result ───────────────────────────────────────
    const workflowUrl = submission.urlAuthorWorkflow
      || `https://pac.pogil.org/index.php/pac/authorDashboard/submission/${submissionId}`;

    return corsResponse({
      ok:           true,
      submissionId,
      publicationId,
      workflowUrl,
      authorErrors: authorErrors.length ? authorErrors : undefined,
      note: authorErrors.length
        ? 'Submission created but some authors could not be added — check authorErrors.'
        : 'Submission created successfully.',
    }, 201, origin, allowedOrigin);
  }
};

// ── Helpers ─────────────────────────────────────────────────────────────────

function corsResponse(body, status, origin, allowedOrigin) {
  const isAllowed = !allowedOrigin || origin === allowedOrigin || allowedOrigin === '*';
  const headers = {
    'Content-Type':                'application/json',
    'Access-Control-Allow-Origin': isAllowed ? (origin || '*') : 'null',
    'Access-Control-Allow-Methods':'POST, OPTIONS',
    'Access-Control-Allow-Headers':'Content-Type',
  };
  return new Response(
    body != null ? JSON.stringify(body, null, 2) : null,
    { status, headers }
  );
}

function splitName(fullName) {
  // Simple split: everything before the last space = given, last word = family.
  // Handles "First Last", "First Middle Last", single-word names.
  const parts = fullName.trim().split(/\s+/);
  if (parts.length === 1) return { given: parts[0], family: '' };
  return {
    given:  parts.slice(0, -1).join(' '),
    family: parts[parts.length - 1],
  };
}

function buildAbstract(pac) {
  // TODO: Replace this placeholder with the real abstract field once the PAC
  // submission form captures it directly from the author. For now, this encodes
  // key PAC metadata into the OJS abstract field so reviewers have context
  // while editors act on the submission. Once a dedicated abstract/description
  // field exists in structure.xlsx, use that value here instead:
  //   return pac.abstract || '';
  const lines = [];
  if (pac.readiness)   lines.push(`Readiness: ${pac.readiness}`);
  if (pac.review_kind) lines.push(`Review kind: ${pac.review_kind}`);
  if (pac.discipline)  lines.push(`Discipline: ${pac.discipline}`);
  if (pac.course)      lines.push(`Course: ${pac.course}`);
  const objectives = (pac.content_objectives || []).map(o => o.objective_text).filter(Boolean);
  if (objectives.length) lines.push(`Content objectives: ${objectives.join('; ')}`);
  return lines.join('\n');
}
