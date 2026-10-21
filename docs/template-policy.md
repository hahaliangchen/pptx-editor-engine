# Template HTML semantics and native PPTX persistence

The AI edits content only through `get_design` / `apply_html`. A template is an editable design starting point; page roles, examples and layout names are optional references. There is no slot-filling generation route.

## Persistent minimal record

Each slide may carry one `p:extLst/p:ext` with URI `{70D18A51-144B-49B9-A917-03139173DE67}`. Its `design:template` child (`urn:pptx-editor:template:1`) contains versioned JSON:

```json
{"version":1,"page":{"role":"cover","layoutName":"Brand cover","note":"Use a short title"},"nodes":{"3":{"purpose":"fixed","name":"Logo","note":"Preserve original brand"},"6":{"purpose":"content","name":"Title"}}}
```

This stores semantics and source shape identity only, never a geometry copy or complete design document. Unknown keys, invalid purposes/versions and references to nonexistent source shapes are rejected. Strings are capped at 4096 UTF-16 code units. Native edits validate the original record, then remove records only for actually deleted shapes; copy/duplicate remaps policies with the copied shape IDs. Undo restores the source record because every export starts from the unmodified package. The existing `@用途` shape annotations remain readable; role-only annotations do not unlock images/backgrounds.

Saving without any change does not add a record or rewrite untouched XML. A real edit or annotation writes policies alongside the PPTX; reopening needs no saved JSON/HTML. Slide duplication copies its resources and record, with a distinct slide part and consequently distinct design IDs. Office applications retaining custom slide extension elements should preserve the record; independent Office/WPS resave compatibility still needs product acceptance testing.

## HTML contract

- Objects: `data-purpose="fixed|content|reference"`, optional `data-note`.
- Body: optional `data-page-role`, `data-layout-name`, `data-page-note`.
- `data-id`, `data-bind`, `data-role`, `data-name`, `data-locked`, required/constraint flags remain host-owned. Removing an HTML lock flag never authorizes an edit.
- Generation cannot change semantic attributes. Template authoring displays the exact before/after proposal; authorization is bound to page ID, source hash, revision and change values.
- `DesignBridge.updateTemplateSemantics(page, edits)` is a host-only human-form API using the same HTML conversion. It is not an AI tool.
- `DesignBridge` accepts `{mode:"template-authoring",confirmSemanticChanges(proposal)}` as its fourth constructor argument. No callback means no AI annotation grant. Concurrent edits invalidate a pending confirmation.

Metadata-only changes increment revision, are saved and refresh the projection, even if all pixels remain unchanged.

## Effective protection

Original unannotated images/complex assets, named brand objects and full-page backgrounds default fixed. Explicit image placeholders/reference annotations can be mutable. Repeated short top headers are conservatively preserved until author review. New AI-created objects are content and retain that purpose after saving. Fixed group policy applies to descendants. Technical hidden/unpreviewable and master backdrop locks cannot be relaxed via ordinary semantic edits.

Fixed objects cannot change content, geometry, style, parent membership, visibility or overlapping paint order. New/changed foreground objects cannot newly cover fixed regions. A full-slide backdrop is not a full-slide forbidden area: master foreground objects supply smaller protected regions. Existing legal overlap is retained when unchanged. Transparent text is checked using browser-measured visible glyph rectangles (Canvas ink metrics anchored to DOM line positions), so a wide empty text box does not itself cover a logo. Filled/bordered boxes and assets still use their full bounds. If ink metrics are unavailable, measurement conservatively falls back to DOM Range bounds. These values come from the host browser, never HTML attributes, and are not persisted. This is not a general pixel-accurate occlusion engine; final PPT rendering still requires visual review.

Reference groups may be deleted/replaced as a whole if no fixed/technically locked members remain. Retained groups preserve membership. Images are standalone previews with original XML/vector assets kept intact when unchanged. An explicitly replaceable complex asset is replaced as a complete image object, never edited as unsupported XML internals.

## Validation

`tests/design-template.test.ts` uses a copy of the actual Unicom/UniAI template (ordinary image logo at the cover's upper right). Coverage includes image replacement/move/delete, ancestor clipping, foreground cover, exact annotation grants, metadata-only save/reopen, 20-page copies, human-form annotation, concurrent confirmation invalidation, fixed groups, reference group replacement, native delete/duplicate and undo/redo, and malformed records. Existing property-difference tests remain applicable.

The direct solid slide background (`p:bg`, projected as body background color) remains adjustable; it has no separate fixed/content control in this version. Fixed background image/shape objects retain their object policy. Do not describe this as all backgrounds being locked. A full-slide opaque raster containing a logo cannot expose that internal logo as a separate protected region; split the brand mark into an independent asset during template authoring or review its placement manually.

Final permission review also checks root `<html>` CSS, individual CSS transforms and clipping/masks; root opacity cannot hide protected assets. Newly added `data-kind` nodes without `data-id` receive generated IDs but cannot declare privileged semantic metadata before author confirmation.

The engine tests do not replace actual published-template → model generation → final-PPT visual acceptance. The original template fixture is not modified.

## Bounded page reads

`get_design()` returns the document overview: `revision`, `sourceHash`, `canvas`, `totalPages`, diagnostics, and all `pages` with stable IDs and optional semantics. `get_design({page})` returns only the requested page's `page`, `pageId`, `semantics`, `semanticsDefined`, `nodes`, `assets`, and complete HTML, together with the same document identity and `totalPages`. It does not repeat the overview `pages` array. Read the overview once when selecting or ordering pages; read a single page when editing it. Asset references and current-page protection metadata are never omitted.
