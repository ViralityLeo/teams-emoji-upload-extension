# Teams Emoji Uploader Feature Roadmap

Status: proposed. Updated: 2026-10-14.

Goal: make emoji uploads easier, library management safer, and browser support simpler. Phases define suggested delivery order; no release dates committed. All features below remain planned.

## Current baseline

Repository README describes Chrome and Firefox builds, bulk uploads, tenant browsing, Slackmojis imports, filename conflict suffixes, and exact-byte SHA-256 duplicate checks. Tenant hash cache currently lasts up to 24 hours. Visual similarity, including differently encoded GIFs, needs separate detection.

## Phase 1 Platform and usability

- [ ] **Rewrite in cross-platform extension framework.** Select framework after small Chrome/Firefox build proof. Reuse existing upload, session discovery, tenant isolation, and deduplication logic. Completion: both browser packages build from shared source; uploads, viewer, and Slackmojis imports retain current behavior. Other browser targets require separate validation.
- [ ] **UI refresh.** Unify popup and viewer styling; clarify tenant selection, progress, errors, and upload results. Completion: keyboard-accessible controls, readable loading/error states, and clear uploaded/skipped/renamed outcomes across both browsers.
- [ ] **Longer-lasting cache.** Extend reusable image hashes beyond current 24-hour window while refreshing live emoji metadata. Use bounded storage, expiry, and explicit invalidation when emoji content changes. Completion: unchanged images avoid repeat downloads after 24 hours; deleted/changed entries reconcile correctly; users can clear cache. Never reuse stale permissions or session tokens as cache data.

## Phase 2 Automatic naming

Deliver one shared naming pipeline for local uploads and Slackmojis imports: detect convention, normalize basename, apply prefix once, validate final name, then resolve conflicts.

- [ ] **Automatic prefix and filename separator style detector.** Infer common prefix and separator from selected tenant's existing names. Completion: detected convention shown with representative examples; mixed or sparse libraries require explicit selection instead of silent guessing.
- [ ] **Automatic filename normalizer.** Strip final file extension; normalize whitespace and separators; handle unsupported characters against verified Teams naming rules. Completion: before/after preview available; empty names and collisions handled deterministically; repeated normalization leaves name unchanged.
- [ ] **Automatic filename prefix setting.** Apply detected or user-selected prefix, with tenant-scoped saved preference and override. Completion: existing prefix never doubled; preview shows final name; disabling prefix preserves normalization.

## Phase 3 Attribution and library management

- [ ] **Who created emoji attribution.** Investigate creator metadata and available identity lookup. Show creator on emoji cards/details when exposed and resolvable. Completion: unavailable identity shown as unknown; never infer creator from filename or cache ownership.
- [ ] **Emoji deletion rights detection and logic.** Determine deletion eligibility for current user and selected tenant from available permission signals; handle server rejection explicitly. Completion: delete controls reflect known rights, unknown rights remain explicit, and tenant/session changes invalidate permission state. Server response remains authoritative.
- [ ] **Rename emoji by deleting then reuploading.** Depends on deletion rights and recoverable original image bytes. Completion: validate new name and conflicts before deletion; preserve original bytes and metadata needed for recovery; require explicit confirmation; report each operation's result.

Rename sequence: save the original image, validate the replacement, confirm deletion, delete the original, then reupload under the new name. If reupload fails, offer restoration using the saved original. Restoration can also fail; retain recovery data and explain the remaining state. Preserve GIF animation. Warn that replacement may change emoji identity, attribution, and references in existing messages; verify actual Teams behavior before release.

## Phase 4 Image quality

- [ ] **Similarity duplicate detection, including GIFs.** Keep exact-byte checks as fast first pass; add visual comparison for resized/reencoded static images and decoded GIF frames. Account for animation timing and frame differences. Completion: known similar images identified across encodings; distinct animations sharing first frame remain distinguishable; candidate matches show side-by-side previews. Users decide whether to skip; similarity never triggers automatic deletion. Bound processing cost for large libraries and files.
- [ ] **Broken transparent background detection and fix.** Flag likely opaque backgrounds or transparency artifacts; distinguish intentional backgrounds through user review. Completion: before/after preview on light and dark backgrounds; fixes preserve subject edges, output validity, and GIF frame timing. Keep original file available; apply fixes only after user accepts preview. Full subject reconstruction remains outside initial scope.

## Phase 5 Native Teams integration and instant reload

- [ ] **Native Teams app integration and instant reload.** First validate feasible integration paths, authentication, required installation, and refresh behavior on Windows and macOS. Browser extension support alone does not establish desktop app support. Build smallest working path only after proof succeeds.

Completion: upload from supported native Teams workflow; newly uploaded emoji becomes available without full app restart where demonstrated refresh mechanism permits it. Document measured propagation delay, supported app/OS versions, and required setup. If instant refresh is unavailable, expose refresh/sync status and fallback; keep instant reload marked incomplete. Treat native integration and instant reload as separate acceptance checks.

## Delivery gates

- Keep framework migration behavior-preserving before stacking feature changes.
- Ship naming pipeline before rename; ship deletion permission handling before destructive library actions.
- Confirm endpoint capabilities for creator attribution, deletion, rename recovery, and native refresh before committing implementation scope.
- Verify each phase against Chrome and Firefox, multiple tenants, expired sessions, network failures, and animated files where relevant.
- Prioritize working increments. Add libraries or companion services only when existing browser/platform capabilities cannot meet acceptance criteria.

## Decisions needed before implementation

- Framework choice and browser targets beyond Chrome/Firefox.
- Cache retention duration, storage budget, and content-change invalidation signals.
- Teams naming constraints and convention-detection confidence threshold.
- Creator metadata availability and deletion permission signals.
- Similarity thresholds, GIF comparison budget, and representative validation files.
- Native integration mechanism and achievable refresh latency on each supported OS.
