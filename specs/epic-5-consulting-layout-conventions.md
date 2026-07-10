-- allium: 3
--
-- Epic 5 — Bounded consulting-format layout conventions
--
-- These conventions extend the deterministic tidy pass (normalize_layout in
-- scripts/render.py, shared verbatim by renderer and verifier). They tidy
-- hand-estimated geometry; they never change structure.
--
-- FIDELITY RATIONALE (the drift anchor for this epic):
-- Each convention is a BOUNDED INTENT-MAPPING, not a redesign. Sloppy spacing
-- of three near-equal columns MEANS equal columns; a chip drawn touching a box
-- MEANS a flush header; three tall frames near a 1/3+2/3 split MEAN that split.
-- Every mapping carries an explicit numeric tolerance (in config below). Inside
-- the tolerance we snap to the drawn intent; outside it we leave the drawing
-- exactly as-is. That boundedness is what keeps this "mapping, not redesign"
-- and consistent with the NORTH STAR: faithful conversion, not reinterpretation.
--
-- NOTE ON FILE FORM: sibling specs (epic-1, epic-4) are prose checklists. This
-- file is written as an Allium specification because the change is behavioural
-- and benefits from precise, checkable constructs.

------------------------------------------------------------
-- External Entities
------------------------------------------------------------

-- The renderer and the verifier scripts. Both consume the identical tidy pass;
-- that they never disagree is the core cross-host guarantee (Epic 4).
external entity RenderHost {
    role: renderer | verifier
}

------------------------------------------------------------
-- Value Types
------------------------------------------------------------

-- Geometry in normalized 0..1 slide coordinates, origin top-left. w/x are
-- fractions of slide width, h/y fractions of slide height.
value Box {
    x: Decimal
    y: Decimal
    w: Decimal
    h: Decimal

    left: x
    right: x + w
    top: y
    bottom: y + h
    center_x: x + w / 2
    center_y: y + h / 2
}

value TableRow {
    cells: List<String>
}

------------------------------------------------------------
-- Contracts
------------------------------------------------------------

-- The tidy pass as a pure, deterministic transformation. The obligation that
-- renderer and verifier apply the identical function lives here.
contract TidyPass {
    normalize: (layout: Layout) -> Layout

    @invariant Determinism
        -- For an identical input layout, normalize produces byte-identical
        -- geometry. The renderer and the verifier therefore never disagree
        -- about where an element landed.

    @invariant Purity
        -- No clock, no randomness, no host- or run-dependent state. The pass
        -- depends only on the input layout and the config tolerances.

    @invariant StructurePreserving
        -- Element count and connectivity are unchanged, with the single
        -- exception of table expansion, which is itself deterministic and
        -- separately verified. No element is added, merged or rearranged
        -- beyond the bounded geometric snaps specified below.

    @invariant NativeShapesOnly
        -- Every element in the output is a native, editable PowerPoint shape.
        -- In particular a table is expanded into native shapes and never
        -- becomes a native PowerPoint table object.
}

------------------------------------------------------------
-- Enumerations
------------------------------------------------------------

enum ElementType {
    rect | rounded_rect | ellipse | diamond | circle
    | chevron_right | chevron_down
    | triangle_up | triangle_right | triangle_down | triangle_left
    | text
    | line | arrow | double_arrow
}

enum Orientation { row | column }

-- Canonical consulting frame splits. `none` means no canon matched within
-- tolerance (e.g. a 50/50 two-frame layout), which leaves the frames untouched.
enum FrameCanon {
    one_third_two_thirds
    | two_thirds_one_third
    | equal_thirds
    | quarter_half_quarter
    | none
}

------------------------------------------------------------
-- Entities and Variants
------------------------------------------------------------

entity Element {
    id: String?
    element_type: ElementType
    box: Box
    text: String?

    -- Connector endpoints reference element ids (never a table id; see
    -- invariant ConnectorsNeverReferenceTable).
    from_id: String?
    to_id: String?

    layout: Layout

    is_box: element_type in {rect, rounded_rect, ellipse, diamond, circle}
    is_connector: element_type in {line, arrow, double_arrow}
    is_circle: element_type = circle
    is_gap_shape: element_type in {
        triangle_up, triangle_right, triangle_down, triangle_left,
        chevron_right, chevron_down
    }

    -- A circle renders round when its width and height cover the same physical
    -- distance on the 16:9 canvas. The tidy pass must preserve this.
    is_round: box.w * config.slide_width_in = box.h * config.slide_height_in
}

-- New in Epic 5. A table is expanded deterministically into native shapes;
-- it is never rendered as a native PowerPoint table.
entity Table {
    id: String
    box: Box
    columns: List<String>?          -- header strings, optional
    rows: List<TableRow>            -- each row is a list of cell strings
    row_headers: Boolean           -- when true, the first cell of each row is a header box
    layout: Layout

    body_row_count: rows.count

    -- Deterministic expansion count: header boxes (one per column header),
    -- one text/header box per body cell, and one horizontal separator per
    -- body row (between and below body rows). Computed the same way in the
    -- renderer and the verifier.
    expanded_element_count: table_element_count(this)   -- black box, deterministic
}

-- A detected cluster of 3+ aligned boxes (or column-groups clustered by
-- centre) considered for equidistant distribution. member_boxes are ordered
-- along the axis; each may be a single box or the bounding box of a clustered
-- column-group.
entity DistributionGroup {
    orientation: Orientation
    member_boxes: List<Box>
    exempt_gap_indices: Set<Integer>   -- gaps occupied by a drawn shape or crossed by a connector

    member_count: member_boxes.count

    -- Edge-to-edge gaps in axis order, then the subset subject to the bound.
    edge_gaps: gap_sizes(member_boxes, orientation)        -- black box
    unexempt_gaps: gaps_excluding(edge_gaps, exempt_gap_indices)   -- black box

    -- Similar spacing: the largest bound-subject gap is no more than
    -- gap_ratio_bound times the smallest. Exempt gaps do not disqualify.
    gaps_within_bound:
        unexempt_gaps.count > 0
        and max_of(unexempt_gaps) <= config.gap_ratio_bound * min_of(unexempt_gaps)

    qualifies_for_equidistant:
        member_count >= config.min_group_size and gaps_within_bound
}

-- A pair of boxes drawn touching or nearly touching, candidates for a flush snap
-- (e.g. a header chip sitting on top of its body box).
entity FlushPair {
    box_a: Box
    box_b: Box
    orientation: Orientation            -- axis along which the two are adjacent

    edge_distance: edge_gap(box_a, box_b, orientation)              -- black box, normalized to the axis' slide dimension
    perpendicular_overlap: overlap_fraction(box_a, box_b, orientation)   -- black box, 0..1

    should_snap:
        edge_distance <= config.flush_edge_tolerance
        and perpendicular_overlap >= config.flush_overlap_min
}

-- A set of 2-3 tall frames (column-groups) considered for a canonical
-- consulting split. frame_boxes are ordered left to right.
entity FrameSet {
    frame_boxes: List<Box>
    content_box: Box

    frame_count: frame_boxes.count
    union_height_fraction: union_height(frame_boxes) / content_box.h       -- black box
    coverage_width_fraction: coverage_width(frame_boxes) / content_box.w   -- black box
    non_overlapping: frames_non_overlapping(frame_boxes)                   -- black box
    width_fractions: frame_width_fractions(frame_boxes, content_box)       -- black box, List<Decimal>

    -- The nearest canonical split whose width fractions are all within
    -- frame_ratio_tolerance, or `none` if no canon matches (e.g. 50/50).
    nearest_canon: canon_within(width_fractions, config.frame_ratio_tolerance)   -- black box -> FrameCanon

    matches_canon:
        frame_count >= config.min_frames
        and frame_count <= config.max_frames
        and union_height_fraction >= config.frame_min_union_height
        and coverage_width_fraction >= config.frame_min_coverage_width
        and non_overlapping
        and nearest_canon != none
}

entity Layout {
    status: raw | tidied
    elements: Element with layout = this
    tables: Table with layout = this
    content_box: Box

    transitions status {
        raw -> tidied
        terminal: tidied
    }

    -- Detected candidates for each convention. Detection is part of the
    -- deterministic pass; the decision predicates live on the entities above.
    distribution_groups: detect_distribution_groups(this)   -- black box -> Set<DistributionGroup>
    flush_pairs: detect_flush_pairs(this)                   -- black box -> Set<FlushPair>
    frame_sets: detect_frame_sets(this)                     -- black box -> Set<FrameSet>
}

------------------------------------------------------------
-- Config
------------------------------------------------------------

config {
    -- Equidistant distribution
    min_group_size: Integer = 3           -- 3+ aligned members required
    gap_ratio_bound: Decimal = 2.0        -- largest bound-subject gap <= 2x smallest

    -- Flush snap
    flush_edge_tolerance: Decimal = 0.015 -- edge distance <= 1.5% of slide dimension
    flush_overlap_min: Decimal = 0.5      -- >= 50% overlap on the perpendicular axis

    -- Canonical frame ratios
    min_frames: Integer = 2
    max_frames: Integer = 3
    frame_min_union_height: Decimal = 0.55    -- union frame height >= ~55% of content height
    frame_min_coverage_width: Decimal = 0.70  -- frames jointly cover >= ~70% of content width
    frame_ratio_tolerance: Decimal = 0.08     -- width fractions within 8 percentage points of canon

    -- Canvas geometry (16:9), used to keep circles round through the pass.
    slide_width_in: Decimal = 13.333
    slide_height_in: Decimal = 7.5
}

------------------------------------------------------------
-- Defaults
------------------------------------------------------------
-- Acceptance fixtures. The decision predicates above make each case checkable.

-- SCR three-column case: three columns, gaps 0.04 and 0.06 (ratio 1.5 <= 2),
-- no exempt gaps. Expected: qualifies_for_equidistant = true -> gaps equalized.
default DistributionGroup scr_three_column = {
    orientation: row,
    member_boxes: [
        { x: 0.10, y: 0.30, w: 0.20, h: 0.14 },
        { x: 0.34, y: 0.30, w: 0.20, h: 0.14 },
        { x: 0.60, y: 0.30, w: 0.20, h: 0.14 }
    ],
    exempt_gap_indices: {}
}

-- Clustered-timeline non-case: two boxes clustered (gap 0.02), one far (gap
-- 0.20), ratio 10 > 2, no exemption. Expected: qualifies_for_equidistant =
-- false -> the deliberate clustering is left untouched.
default DistributionGroup clustered_timeline = {
    orientation: row,
    member_boxes: [
        { x: 0.10, y: 0.30, w: 0.15, h: 0.14 },
        { x: 0.27, y: 0.30, w: 0.15, h: 0.14 },
        { x: 0.62, y: 0.30, w: 0.15, h: 0.14 }
    ],
    exempt_gap_indices: {}
}

-- 50/50 non-case: two frames of equal width. Width fractions ~0.5/0.5 match no
-- canon. Expected: matches_canon = false -> the frames are left untouched.
default FrameSet fifty_fifty = {
    frame_boxes: [
        { x: 0.05, y: 0.15, w: 0.44, h: 0.70 },
        { x: 0.51, y: 0.15, w: 0.44, h: 0.70 }
    ],
    content_box: { x: 0.05, y: 0.15, w: 0.90, h: 0.70 }
}

------------------------------------------------------------
-- Rules
------------------------------------------------------------

-- The single deterministic tidy pass. Each convention only fires inside its
-- bounded tolerance; outside it, geometry is left exactly as drawn. The
-- detailed geometric mechanics are deferred to normalize_layout.
rule TidyLayout {
    when: NormalizeLayout(host, layout)
    requires: layout.status = raw
    ensures:
        layout.status = tidied

        -- Table expansion: deterministic, native shapes only.
        for table in layout.tables:
            TableExpansion.expand(table)

        -- Equidistant distribution: only similarly-spaced 3+ groups, outer
        -- span fixed, shapes in gaps re-centred; out-of-bound gaps untouched.
        for group in layout.distribution_groups:
            if group.qualifies_for_equidistant:
                GapDistribution.equalize(group)

        -- Flush snap: near-touching, well-overlapped pairs snap flush.
        for pair in layout.flush_pairs:
            if pair.should_snap:
                FlushSnap.snap(pair)

        -- Canonical frame ratios: 2-3 tall frames near a canon snap to the
        -- exact ratio with uniform gutters; 50/50 (no canon) stays untouched.
        for frames in layout.frame_sets:
            if frames.matches_canon:
                FrameRatioSnap.snap(frames)

    @guidance
        -- Table expansion produces: one bold header box per column header;
        -- one borderless, left-aligned text element per body cell (or a header
        -- box for the first cell of each row when row_headers is true); and one
        -- horizontal separator line per body row, between and below body rows.
        -- Columns are equal width, rows equal height. A table is NEVER a native
        -- PowerPoint table, and a connector referencing a table id is a
        -- validation error (see ConnectorsNeverReferenceTable).
        --
        -- Every snap is bounded by config tolerances: within tolerance is drawn
        -- intent (mapping); outside tolerance is deliberate (left as-is).
        --
        -- The verifier re-applies this identical pass, so render and verify
        -- never disagree about the tidied geometry.
}

------------------------------------------------------------
-- Invariants
------------------------------------------------------------

-- Circles remain circular through every snap in the pass.
invariant CirclesStayCircular {
    for e in Elements:
        e.is_circle implies e.is_round
}

-- Connectors may never reference a table's id. A table is expanded into
-- separate native shapes, so there is no single shape for a connector to glue
-- to; such a reference is a validation error.
invariant ConnectorsNeverReferenceTable {
    for c in Elements:
        for t in Tables:
            c.is_connector implies
                (c.from_id = null or c.from_id != t.id)
                and (c.to_id = null or c.to_id != t.id)
}

------------------------------------------------------------
-- Actor Declarations
------------------------------------------------------------

-- The tidy pass is consumed by two code parties: the renderer and the verifier.
actor TidyConsumer {
    identified_by: RenderHost where role in {renderer, verifier}
}

------------------------------------------------------------
-- Surfaces
------------------------------------------------------------

surface LayoutNormalization {
    facing host: TidyConsumer
    context layout: Layout

    exposes:
        layout.status
        layout.elements
        layout.tables

    provides:
        NormalizeLayout(host, layout) when layout.status = raw

    contracts:
        demands TidyPass

    @guarantee RenderVerifyAgree
        -- The renderer and the verifier invoke the identical tidy pass and
        -- therefore compute identical tidied geometry. The verifier's `passed`
        -- verdict reflects agreement, never a re-derivation.

    @guarantee NativeShapesOnly
        -- The tidied output contains only native, editable PowerPoint shapes,
        -- including every element a table expands into.

    @guarantee BoundedIntentMapping
        -- Each convention snaps only within an explicit numeric tolerance,
        -- mapping drawn intent; geometry outside tolerance is left as drawn.
        -- This is mapping, not redesign.
}

------------------------------------------------------------
-- Deferred Specifications
------------------------------------------------------------

deferred TableExpansion.expand    -- see: scripts/render.py (deterministic table expansion, mirrored in verify.py)
deferred GapDistribution.equalize -- see: scripts/render.py normalize_layout (equidistant distribution)
deferred FlushSnap.snap           -- see: scripts/render.py normalize_layout (flush snap)
deferred FrameRatioSnap.snap      -- see: scripts/render.py normalize_layout (canonical frame ratios)

------------------------------------------------------------
-- Resolved Questions (2026-07-10, decided with the implementation)
------------------------------------------------------------

-- RESOLVED: application order is fixed as: circle fix -> canvas fit ->
-- row/column alignment -> size unification -> flush snap -> frame-ratio snap
-- -> equidistant distribution (horizontal, then vertical) -> circle re-fix.
-- Flush snaps survive because distribution and frame snaps move column-groups
-- as units: a chip and its body box always share a column-group, so every
-- later horizontal move applies the same delta to both. Flush is vertical;
-- horizontal distribution is orthogonal to it. Frame snaps scale all boxes of
-- a group by the same factor about the group's left edge, preserving flush
-- and internal alignment.

-- RESOLVED: column-group clustering reuses the base tidy-pass tolerances
-- verbatim (COL_TOL 0.035 horizontal, ROW_TOL 0.05 vertical). No new
-- threshold is introduced.

-- RESOLVED: the frame gates are exact config defaults: union height >= 0.55
-- of content height, joint coverage >= 0.70 of content width. No fuzz.
