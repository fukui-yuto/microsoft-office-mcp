"""Design tokens for PowerPoint presentations.

All color, font, and size decisions are expressed as Enum tokens.
No raw RGB values should leak into tool arguments.
"""

from enum import Enum


# ---------------------------------------------------------------------------
# Theme names
# ---------------------------------------------------------------------------

class ThemeName(str, Enum):
    """Available presentation themes."""
    BUSINESS = "business"
    MINIMAL = "minimal"
    WARM = "warm"


# ---------------------------------------------------------------------------
# Color roles (max 7 per theme)
# ---------------------------------------------------------------------------

class TokenRole(str, Enum):
    """Semantic color roles within a theme."""
    BG_PRIMARY = "bg_primary"
    BG_SECONDARY = "bg_secondary"
    TEXT_PRIMARY = "text_primary"
    TEXT_SECONDARY = "text_secondary"
    ACCENT = "accent"
    BORDER = "border"
    SURFACE = "surface"


# Each palette: 7 colors only.  Derived from Material / Carbon / Keynote.
PALETTES: dict[ThemeName, dict[TokenRole, tuple[int, int, int]]] = {
    ThemeName.BUSINESS: {
        TokenRole.BG_PRIMARY: (255, 255, 255),
        TokenRole.BG_SECONDARY: (245, 247, 250),
        TokenRole.TEXT_PRIMARY: (23, 23, 23),
        TokenRole.TEXT_SECONDARY: (107, 114, 128),
        TokenRole.ACCENT: (37, 99, 235),
        TokenRole.BORDER: (229, 231, 235),
        TokenRole.SURFACE: (249, 250, 251),
    },
    ThemeName.MINIMAL: {
        TokenRole.BG_PRIMARY: (255, 255, 255),
        TokenRole.BG_SECONDARY: (250, 250, 250),
        TokenRole.TEXT_PRIMARY: (28, 28, 28),
        TokenRole.TEXT_SECONDARY: (115, 115, 115),
        TokenRole.ACCENT: (15, 15, 15),
        TokenRole.BORDER: (230, 230, 230),
        TokenRole.SURFACE: (245, 245, 245),
    },
    ThemeName.WARM: {
        TokenRole.BG_PRIMARY: (255, 252, 248),
        TokenRole.BG_SECONDARY: (254, 243, 232),
        TokenRole.TEXT_PRIMARY: (41, 37, 36),
        TokenRole.TEXT_SECONDARY: (120, 113, 108),
        TokenRole.ACCENT: (194, 65, 12),
        TokenRole.BORDER: (231, 229, 228),
        TokenRole.SURFACE: (250, 245, 240),
    },
}


# ---------------------------------------------------------------------------
# Font tokens (fixed set — no arbitrary font names)
# ---------------------------------------------------------------------------

class FontToken(str, Enum):
    """Font tokens mapped to concrete font families."""
    TITLE = "title"
    BODY = "body"
    MONO = "mono"


FONT_MAP: dict[FontToken, str] = {
    FontToken.TITLE: "BIZ UDPGothic",
    FontToken.BODY: "BIZ UDPGothic",
    FontToken.MONO: "Consolas",
}


# ---------------------------------------------------------------------------
# Size tokens (fixed typographic scale)
# ---------------------------------------------------------------------------

class SizeToken(str, Enum):
    """Typographic size scale."""
    XL = "xl"   # 36 pt — slide title
    LG = "lg"   # 24 pt — section heading
    MD = "md"   # 18 pt — body text
    SM = "sm"   # 14 pt — caption
    XS = "xs"   # 11 pt — footnote

SIZE_MAP: dict[SizeToken, float] = {
    SizeToken.XL: 36.0,
    SizeToken.LG: 24.0,
    SizeToken.MD: 18.0,
    SizeToken.SM: 14.0,
    SizeToken.XS: 11.0,
}

# Allowed range for validation
FONT_SIZE_MIN: float = SIZE_MAP[SizeToken.XS]   # 11 pt
FONT_SIZE_MAX: float = SIZE_MAP[SizeToken.XL]    # 36 pt


# ---------------------------------------------------------------------------
# Constraint constants
# ---------------------------------------------------------------------------

MAX_BULLETS = 5
MAX_ACCENT_PER_SLIDE = 1
MAX_COLORS_PER_SLIDE = 6
MAX_COMPARISON_ITEMS = 4


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_color(theme: ThemeName, role: TokenRole) -> tuple[int, int, int]:
    """Resolve a token to its RGB value."""
    return PALETTES[theme][role]


def get_font(token: FontToken) -> str:
    """Resolve a font token to its family name."""
    return FONT_MAP[token]


def get_size(token: SizeToken) -> float:
    """Resolve a size token to pt value."""
    return SIZE_MAP[token]
