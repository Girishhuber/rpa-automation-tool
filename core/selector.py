
from __future__ import annotations
import time
from dataclasses import dataclass, field
from typing import Optional, Any
import re


ATTR_SCORES: dict[str, int] = {
    "automation_id":  100,
    "aria_label":      70,
    "xpath_id":        75,
    "css_id":          75,
    "name":            80,
    "control_type":    60,
    "xpath_text":      50,
    "placeholder":     55,
    "class_name":      40,
    "tag_name":        30,
    "sibling_index":   20,
    "ancestor_chain":  35,
    "bbox":            10,
    "screen_coord":     5,
    "process_name":    45,
    "window_title":    50,
}


_UNSTABLE_PATTERNS = [
    re.compile(r"^[0-9]{6,}$"),               # S-1: pure numeric 6+ digits
    re.compile(r"\b[a-f0-9]{8,}\b"),          # SEL-4: hex hash (whole word)
    re.compile(r"^_\w+\d{4,}$"),              # underscore + 4+ digit suffix
    re.compile(r"-[a-f0-9]{6,}$"),            # CSS module hash suffix (e.g. "btn-a1b2c3")
    re.compile(r"_[a-f0-9]{6,}$"),            # JS module hash suffix
]

# S-1: whitelist of known stable short numeric IDs (Windows dialog resource IDs)
_KNOWN_STABLE_IDS = {
    "1148", "1001", "1000", "100", "101", "1", "2", "3", "4",
    "200", "201", "300", "1003", "1004", "1005",
}


def _is_unstable(value: str) -> bool:
    """
    Return True if value looks like a dynamic/generated identifier that will
    not survive across sessions or UI reloads.
    """
    if not value:
        return True
    if value in _KNOWN_STABLE_IDS:
        return False
    # Very short values are stable (single-char button names, etc.)
    if len(value) <= 3 and value.isalnum():
        return False
    for pat in _UNSTABLE_PATTERNS:
        if pat.search(value):
            return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# Data classes
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class SelectorAttribute:
    name:  str
    value: Any
    score: int = 0

    def __post_init__(self):
        if self.score == 0:
            self.score = ATTR_SCORES.get(self.name, 10)


@dataclass
class SelectorStrategy:
    name:        str
    attributes:  list[SelectorAttribute]
    total_score: int = 0
    backend:     str = "uia"

    # SEL-2: per-strategy success tracking
    successes:  int = 0
    failures:   int = 0

    def __post_init__(self):
        self.total_score = sum(a.score for a in self.attributes)

    @property
    def success_rate(self) -> float:
        total = self.successes + self.failures
        return self.successes / total if total > 0 else 0.5

    def record_success(self) -> None:
        self.successes += 1

    def record_failure(self) -> None:
        self.failures += 1

    def to_dict(self) -> dict:
        return {
            "name":        self.name,
            "backend":     self.backend,
            "total_score": self.total_score,
            "attributes":  [{"name": a.name, "value": a.value, "score": a.score}
                            for a in self.attributes],
            "successes":   self.successes,
            "failures":    self.failures,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "SelectorStrategy":
        attrs = [SelectorAttribute(a["name"], a["value"], a.get("score", 0))
                 for a in d.get("attributes", [])]
        s = cls(name=d["name"], attributes=attrs, backend=d.get("backend", "uia"))
        s.total_score = d.get("total_score", s.total_score)
        s.successes   = d.get("successes", 0)
        s.failures    = d.get("failures",  0)
        return s


@dataclass
class AnchorElement:
    """
    SEL-3: A nearby stable element that can help disambiguate the primary target.
    Example: label "Username:" to the left of an input field.
    """
    direction:     str            # "left", "right", "above", "below", "parent"
    name:          Optional[str]  = None
    control_type:  Optional[str]  = None
    automation_id: Optional[str]  = None
    offset_x:      int = 0        # pixel offset from primary element
    offset_y:      int = 0

    def to_dict(self) -> dict:
        return {
            "direction":     self.direction,
            "name":          self.name,
            "control_type":  self.control_type,
            "automation_id": self.automation_id,
            "offset_x":      self.offset_x,
            "offset_y":      self.offset_y,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "AnchorElement":
        return cls(
            direction     = d.get("direction", "left"),
            name          = d.get("name"),
            control_type  = d.get("control_type"),
            automation_id = d.get("automation_id"),
            offset_x      = d.get("offset_x", 0),
            offset_y      = d.get("offset_y", 0),
        )


# SEL-2: Current schema version — bump when Selector fields change
SELECTOR_SCHEMA_VERSION = "2.1"

# SEL-1: How many seconds before confidence starts decaying
_DECAY_HALF_LIFE_DAYS = 30.0
_DECAY_HALF_LIFE_S    = _DECAY_HALF_LIFE_DAYS * 86400


@dataclass
class Selector:
    """
    Complete element identity. Stored alongside every recorded event.

    SEL-1: recorded_at + time_decay_score() for confidence decay.
    SEL-2: selector_version + per-strategy success_rate tracking.
    SEL-3: anchor_elements for disambiguation.
    """
    strategies:       list[SelectorStrategy]
    process_name:     Optional[str] = None
    window_title:     Optional[str] = None
    semantic_role:    Optional[str] = None
    is_editable:      bool  = False
    is_clickable:     bool  = True
    confidence:       float = 0.0

    # Positioning
    screen_x:         Optional[int] = None
    screen_y:         Optional[int] = None

    # SEL-2: versioning + replay history
    selector_version: str = SELECTOR_SCHEMA_VERSION
    recorded_at:      float = field(default_factory=time.time)  # SEL-1: epoch seconds
    last_used_at:     Optional[float] = None
    replay_successes: int = 0
    replay_failures:  int = 0
    last_strategy_used: Optional[str] = None

    # SEL-3: anchor elements for disambiguation
    anchor_elements:  list[AnchorElement] = field(default_factory=list)


    @property
    def success_rate(self) -> float:
        """SEL-2: lifetime replay success rate."""
        total = self.replay_successes + self.replay_failures
        return self.replay_successes / total if total > 0 else 1.0

    def time_decay_score(self) -> float:
        """
        SEL-1: Confidence penalty based on age.
        Returns multiplier 1.0 (fresh) → ~0.5 (one half-life old) → ~0.25 (two).
        Uses exponential decay: score = 0.5^(age_seconds / half_life).
        """
        age_s = time.time() - self.recorded_at
        if age_s <= 0:
            return 1.0
        return 0.5 ** (age_s / _DECAY_HALF_LIFE_S)

    def effective_confidence(self) -> float:
        """Base confidence × time decay × success rate."""
        return self.confidence * self.time_decay_score() * max(self.success_rate, 0.1)

    def best_strategy(self) -> Optional[SelectorStrategy]:
        """ADV-1 compatible: return strategy with highest success_rate, else first."""
        if not self.strategies:
            return None
        by_rate = sorted(self.strategies, key=lambda s: s.success_rate, reverse=True)
        return by_rate[0]

    def strategy_names(self) -> list[str]:
        return [s.name for s in self.strategies]

    def ordered_strategies(self) -> list[SelectorStrategy]:
        """Return strategies sorted by success_rate DESC (for adaptive ordering)."""
        return sorted(self.strategies, key=lambda s: s.success_rate, reverse=True)

    def record_replay(self, strategy_used: str, success: bool) -> None:
        """SEL-2: Update per-selector and per-strategy stats."""
        if success:
            self.replay_successes += 1
        else:
            self.replay_failures += 1
        self.last_strategy_used = strategy_used
        self.last_used_at = time.time()
        # Update per-strategy stats
        for s in self.strategies:
            if s.name == strategy_used:
                if success:
                    s.record_success()
                else:
                    s.record_failure()
                break

    def to_dict(self) -> dict:
        return {
            "selector_version":  self.selector_version,
            "strategies":        [s.to_dict() for s in self.strategies],
            "process_name":      self.process_name,
            "window_title":      self.window_title,
            "semantic_role":     self.semantic_role,
            "is_editable":       self.is_editable,
            "is_clickable":      self.is_clickable,
            "confidence":        self.confidence,
            "screen_x":          self.screen_x,
            "screen_y":          self.screen_y,
            "recorded_at":       self.recorded_at,
            "last_used_at":      self.last_used_at,
            "replay_successes":  self.replay_successes,
            "replay_failures":   self.replay_failures,
            "last_strategy_used": self.last_strategy_used,
            "anchor_elements":   [a.to_dict() for a in self.anchor_elements],
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Selector":
        return cls(
            selector_version   = d.get("selector_version", "1.0"),
            strategies         = [SelectorStrategy.from_dict(s) for s in d.get("strategies", [])],
            process_name       = d.get("process_name"),
            window_title       = d.get("window_title"),
            semantic_role      = d.get("semantic_role"),
            is_editable        = d.get("is_editable", False),
            is_clickable       = d.get("is_clickable", True),
            confidence         = d.get("confidence", 0.0),
            screen_x           = d.get("screen_x"),
            screen_y           = d.get("screen_y"),
            recorded_at        = d.get("recorded_at", time.time()),
            last_used_at       = d.get("last_used_at"),
            replay_successes   = d.get("replay_successes", 0),
            replay_failures    = d.get("replay_failures",  0),
            last_strategy_used = d.get("last_strategy_used"),
            anchor_elements    = [AnchorElement.from_dict(a) for a in d.get("anchor_elements", [])],
        )

    def is_stale(self, threshold_days: float = 90.0) -> bool:
        """SEL-1: True if selector is older than threshold_days."""
        age_days = (time.time() - self.recorded_at) / 86400
        return age_days > threshold_days

    def needs_refresh(self) -> bool:
        """SEL-2: True if success_rate is low enough to warrant re-recording."""
        if self.replay_successes + self.replay_failures < 5:
            return False   # not enough data
        return self.success_rate < 0.5


class SelectorBuilder:

    @staticmethod
    def from_uia(
        automation_id:   Optional[str],
        name:            Optional[str],
        control_type:    Optional[str],
        class_name:      Optional[str],
        window_title:    Optional[str],
        process_name:    Optional[str],
        bbox:            Optional[Any],
        ancestor_chain:  list[str],
        sibling_index:   Optional[int],
        screen_x:        Optional[int],
        screen_y:        Optional[int],
        dpi_scale:       float = 1.0,
        capabilities:    Optional[dict] = None,
        anchor_elements: Optional[list[AnchorElement]] = None,
    ) -> Selector:
        caps = capabilities or {}
        raw: list[SelectorAttribute] = []

        aid_stable = automation_id and not _is_unstable(automation_id)
        if aid_stable:
            raw.append(SelectorAttribute("automation_id", automation_id))

        if name and name.strip():
            raw.append(SelectorAttribute("name", name.strip()[:120]))

        if control_type:
            raw.append(SelectorAttribute("control_type", control_type))

        if class_name and not _is_unstable(class_name):
            raw.append(SelectorAttribute("class_name", class_name))

        if ancestor_chain:
            raw.append(SelectorAttribute("ancestor_chain", ancestor_chain[:3]))

        if sibling_index is not None:
            raw.append(SelectorAttribute("sibling_index", sibling_index))

        if bbox:
            raw.append(SelectorAttribute("bbox", {
                "left": bbox.left, "top": bbox.top,
                "right": bbox.right, "bottom": bbox.bottom,
                "dpi": dpi_scale,
            }))

        if screen_x is not None:
            raw.append(SelectorAttribute("screen_coord", {"x": screen_x, "y": screen_y}))

        raw.sort(key=lambda a: a.score, reverse=True)

        strategies: list[SelectorStrategy] = []

        if aid_stable:
            strict_attrs = [a for a in raw if a.name == "automation_id"]
            if window_title:
                strict_attrs = strict_attrs + [SelectorAttribute("window_title", window_title, score=50)]
            elif process_name:
                # S-3: process_name when window_title absent
                strict_attrs = strict_attrs + [SelectorAttribute("process_name", process_name, score=45)]
            strategies.append(SelectorStrategy(name="strict", attributes=strict_attrs, backend="uia"))

        # 2: SEMANTIC — name + ctrl_type + window
        sem_attrs = [a for a in raw if a.name in ("name", "control_type")]
        if len(sem_attrs) >= 2:
            ctx = ([SelectorAttribute("window_title", window_title, score=50)] if window_title else
                   [SelectorAttribute("process_name", process_name, score=45)] if process_name else [])
            strategies.append(SelectorStrategy(name="semantic",
                                               attributes=sem_attrs + ctx, backend="uia"))

        # 3: RELAXED — name only
        name_attrs = [a for a in raw if a.name == "name"]
        if name_attrs:
            strategies.append(SelectorStrategy(name="relaxed", attributes=name_attrs, backend="uia"))

        # 4: ANCESTOR
        anc_attrs = [a for a in raw if a.name in ("ancestor_chain", "control_type", "class_name")]
        if anc_attrs and ancestor_chain:
            strategies.append(SelectorStrategy(name="ancestor", attributes=anc_attrs, backend="uia"))

        # 5: POSITIONAL
        pos_attrs = [a for a in raw if a.name in ("bbox", "screen_coord")]
        if pos_attrs:
            strategies.append(SelectorStrategy(name="positional", attributes=pos_attrs, backend="uia"))

        max_possible = 100 + 80 + 60
        actual       = sum(a.score for a in raw[:3])
        confidence   = min(actual / max_possible, 1.0)

        return Selector(
            strategies      = strategies,
            process_name    = process_name,
            window_title    = window_title,
            semantic_role   = _detect_semantic_role_uia(control_type, name, caps),
            is_editable     = caps.get("is_editable", False),
            is_clickable    = caps.get("is_clickable", True),
            confidence      = confidence,
            screen_x        = screen_x,
            screen_y        = screen_y,
            recorded_at     = time.time(),
            anchor_elements = anchor_elements or [],
        )

    @staticmethod
    def from_browser(
        xpath:           Optional[str],
        css_selector:    Optional[str],
        element_id:      Optional[str],
        aria_label:      Optional[str],
        aria_role:       Optional[str],
        inner_text:      Optional[str],
        tag_name:        Optional[str],
        name_attr:       Optional[str],
        placeholder:     Optional[str],
        href:            Optional[str],
        tab_url:         Optional[str],
        screen_x:        Optional[int],
        screen_y:        Optional[int],
        anchor_elements: Optional[list[AnchorElement]] = None,
    ) -> Selector:
        raw: list[SelectorAttribute] = []

        # S-2: xpath_id detection — handles @id= with both quote styles
        def _xpath_has_id(xp: str, eid: str) -> bool:
            if not xp or not eid:
                return False
            return (f'@id="{eid}"' in xp or f"@id='{eid}'" in xp)

        if xpath and element_id and _xpath_has_id(xpath, element_id):
            raw.append(SelectorAttribute("xpath_id", xpath, score=75))
        elif xpath:
            is_positional = bool(re.match(r"^/html(/body)?(/[a-z]+\[\d+\])+$", xpath))
            raw.append(SelectorAttribute("xpath_text", xpath, score=40 if is_positional else 55))

        if css_selector and element_id and css_selector.startswith("#"):
            raw.append(SelectorAttribute("css_id", css_selector, score=75))

        if aria_label:
            raw.append(SelectorAttribute("aria_label", aria_label, score=70))

        if name_attr and not _is_unstable(name_attr):
            raw.append(SelectorAttribute("name", name_attr, score=65))

        if placeholder:
            raw.append(SelectorAttribute("placeholder", placeholder, score=55))

        if inner_text and inner_text.strip():
            raw.append(SelectorAttribute("name", inner_text.strip()[:100], score=50))

        if tag_name:
            raw.append(SelectorAttribute("tag_name", tag_name, score=30))

        if screen_x is not None:
            raw.append(SelectorAttribute("screen_coord", {"x": screen_x, "y": screen_y}, score=5))

        raw.sort(key=lambda a: a.score, reverse=True)

        strategies: list[SelectorStrategy] = []

        id_attrs = [a for a in raw if a.name in ("xpath_id", "css_id")]
        if id_attrs:
            strategies.append(SelectorStrategy(name="strict",   attributes=id_attrs, backend="browser"))

        aria_attrs = [a for a in raw if a.name == "aria_label"]
        if aria_attrs:
            strategies.append(SelectorStrategy(name="aria",     attributes=aria_attrs, backend="browser"))

        sem_attrs = [a for a in raw if a.name in ("name", "placeholder", "tag_name")]
        if sem_attrs:
            strategies.append(SelectorStrategy(name="semantic", attributes=sem_attrs[:3], backend="browser"))

        xpath_attrs = [a for a in raw if a.name == "xpath_text"]
        if xpath_attrs:
            strategies.append(SelectorStrategy(name="xpath",    attributes=xpath_attrs, backend="browser"))

        pos_attrs = [a for a in raw if a.name == "screen_coord"]
        if pos_attrs:
            strategies.append(SelectorStrategy(name="positional", attributes=pos_attrs, backend="browser"))

        max_possible = 75 + 70
        actual       = sum(a.score for a in raw[:2])
        confidence   = min(actual / max_possible, 1.0) if max_possible > 0 else 0.0

        return Selector(
            strategies      = strategies,
            window_title    = tab_url,
            semantic_role   = _detect_semantic_role_browser(tag_name, aria_role, aria_label),
            is_editable     = tag_name in ("input", "textarea", "select") if tag_name else False,
            is_clickable    = True,
            confidence      = confidence,
            screen_x        = screen_x,
            screen_y        = screen_y,
            recorded_at     = time.time(),
            anchor_elements = anchor_elements or [],
        )



def _detect_semantic_role_uia(
    control_type: Optional[str],
    name:         Optional[str],
    caps:         dict,
) -> Optional[str]:
    if not control_type:
        return None
    ct = control_type.lower()
    if "button"   in ct:             return "button"
    if ct in ("edit", "document"):   return "input"
    if "checkbox" in ct:             return "checkbox"
    if "radio"    in ct:             return "radio"
    if "combobox" in ct:             return "dropdown"
    if ct in ("dataitem", "cell"):   return "cell"
    if "tab"      in ct:             return "tab"
    if "menu"     in ct:             return "menu"
    if "list"     in ct:             return "list"
    if "tree"     in ct:             return "tree"
    if "pane"     in ct:             return "pane"
    return "generic"


def _detect_semantic_role_browser(
    tag:   Optional[str],
    role:  Optional[str],
    label: Optional[str],
) -> Optional[str]:
    if role:
        r = role.lower()
        if "button"   in r: return "button"
        if "checkbox" in r: return "checkbox"
        if "input"    in r: return "input"
        if "link"     in r: return "link"
        if "cell"     in r: return "cell"
        if "menu"     in r: return "menu"
    if tag:
        return {
            "button":   "button",
            "input":    "input",
            "select":   "dropdown",
            "textarea": "input",
            "a":        "link",
            "td":       "cell",
            "th":       "cell",
        }.get(tag.lower(), "generic")
    return "generic"
