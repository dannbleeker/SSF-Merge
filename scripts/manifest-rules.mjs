/**
 * The manifest rules that can be checked without asking Microsoft.
 *
 * `office-addin-manifest validate` is the authority and CI runs it in a job of
 * its own. It calls a Microsoft SERVICE, so it cannot run in a sandbox with no
 * route out — which is not hypothetical, it is this development environment —
 * and a sibling project went the whole life of its repo with manifests nothing
 * had ever validated. These are the small number of rules whose violation costs
 * something specific, so they hold even when the validator is unreachable.
 *
 * Every one of them is a thing that has actually been shipped somewhere, or is
 * one edit away from being.
 */
import { withoutXmlComments } from "./without-prose.mjs";

export const REQUIRED_ID = "43ebbbac-44ad-42b2-a582-0ef079093e6c";

/** Between the four generated files, which are meant to differ in one thing. */
export function isProd(name) {
  return name.includes("-prod");
}

function xmlRules(raw, name, out) {
  const text = withoutXmlComments(raw);
  if (!text.includes("<OfficeApp")) {
    out.push(`${name} is not an Office add-in manifest`);
    return;
  }
  const version = /<Version>([^<]+)<\/Version>/.exec(text)?.[1];
  if (!version) out.push(`${name} has no <Version>`);
  else if (Number(version.split(".")[0]) < 1) {
    // "Manifest Version Too Low: The manifest has unsupported version number
    // less than 1.0." An ERROR, and a sibling project shipped it in four
    // manifests for months with a fully green suite.
    out.push(`${name} has <Version>${version}</Version>, which is below 1.0 and Office rejects outright`);
  }
  const id = /<Id>([^<]+)<\/Id>/.exec(text)?.[1];
  if (id !== REQUIRED_ID) {
    // A changed GUID is a DIFFERENT add-in: every existing sideload is
    // orphaned, and the user has to remove the old entry by hand with nothing
    // anywhere saying why.
    out.push(`${name} carries id ${id ?? "(none)"}, not ${REQUIRED_ID} — every existing sideload would be orphaned`);
  }
  if (/<Requirements>/.test(text)) {
    // The load-bearing omission. A host that does not meet a DECLARED
    // requirement set does not show the add-in at all: no ribbon entry, no
    // error, nothing to report. checkFloor says which version is missing.
    out.push(`${name} declares <Requirements>; the PowerPointApi floor is checked at runtime by checkFloor instead`);
  }
  if (!/<Permissions>ReadWriteDocument<\/Permissions>/.test(text)) {
    out.push(`${name} does not ask for ReadWriteDocument, which inserting slides needs`);
  }
}

function jsonRules(text, name, out) {
  let doc;
  try {
    doc = JSON.parse(text);
  } catch {
    out.push(`${name} is not valid JSON`);
    return;
  }
  if (doc.id !== REQUIRED_ID) {
    out.push(
      `${name} carries id ${doc.id ?? "(none)"}, not ${REQUIRED_ID} — every existing sideload would be orphaned`,
    );
  }
  if (Number(String(doc.version).split(".")[0]) < 1) {
    out.push(`${name} has version ${doc.version}, which is below 1.0 and Office rejects outright`);
  }
  for (const extension of doc.extensions ?? []) {
    // `capabilities`, not the whole block. `requirements` also carries
    // `scopes` — which host the add-in runs in, the JSON spelling of the XML's
    // <Hosts> — and `formFactors`. Refusing the block outright forbade a
    // correct declaration under a message about a floor that only the one field
    // is about, and for a while the unified manifest did not say it was a
    // PowerPoint add-in because of it.
    if (extension.requirements?.capabilities) {
      out.push(
        `${name} declares requirement-set capabilities; the PowerPointApi floor is checked at runtime by checkFloor instead`,
      );
    }
  }
  const asked = (doc.authorization?.permissions?.resourceSpecific ?? []).map((p) => p.name);
  if (!asked.includes("Document.ReadWrite.User")) {
    out.push(`${name} does not ask for Document.ReadWrite.User, which inserting slides needs`);
  }
}

/** Every URL a manifest points at, in the order it names them. */
export function urlsIn(text) {
  return [...withoutXmlComments(text).matchAll(/https?:\/\/[^"'\s<>]+/g)].map((m) => m[0].replace(/&amp;/g, "&"));
}

/**
 * What is wrong with this manifest, as sentences. Empty means nothing is.
 *
 * `name` decides which format is read and whether the localhost rule applies —
 * a production manifest can be perfectly CURRENT and full of localhost, if the
 * origin the generator was handed ever stops being the production one.
 */
export function checkManifest(text, name) {
  const out = [];
  if (name.endsWith(".json")) jsonRules(text, name, out);
  else xmlRules(text, name, out);

  // HTTPS, which Office requires of every address it fetches. A production
  // manifest on `http://` fails Microsoft's validator, and sideloaded anyway it
  // fails the way this file's other rules describe: no ribbon entry, no error,
  // nothing to report.
  //
  // The namespace declarations are stripped first and that is the whole
  // difficulty. `xmlns="http://schemas.microsoft.com/..."` is an IDENTIFIER,
  // not an address — it is never fetched, it is http by definition, and it may
  // not be changed. A rule that read them would fire on every manifest ever
  // written and would have been deleted rather than fixed.
  //
  // `PROD_ORIGIN` is one constant in `manifest-source.mjs`, which is this
  // file's own test for whether a rule is worth having: one edit away.
  const addresses = urlsIn(text.replace(/xmlns(:[A-Za-z0-9_-]+)?="[^"]*"/g, ""));
  const isLocal = (u) => u.includes("localhost") || u.includes("127.0.0.1");
  const insecure = addresses.filter((u) => u.startsWith("http://") && !isLocal(u));
  if (isProd(name) && insecure.length > 0) {
    out.push(`${name} is a production manifest and points at an insecure address: ${insecure[0]}`);
  }

  const local = urlsIn(text).filter(isLocal);
  if (isProd(name) && local.length > 0) {
    out.push(`${name} is a production manifest and points at localhost: ${local[0]}`);
  }
  if (!isProd(name) && local.length === 0) {
    // The other direction, and it is the one that wastes an afternoon: a dev
    // manifest quietly pointing at production means every edit is tested
    // against the last deploy.
    out.push(`${name} is a development manifest and points at no local origin`);
  }
  return out;
}
