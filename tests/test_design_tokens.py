"""Tests for design_tokens module."""

from microsoft_office.design_tokens import (
    FONT_MAP,
    FONT_SIZE_MAX,
    FONT_SIZE_MIN,
    PALETTES,
    SIZE_MAP,
    FontToken,
    SizeToken,
    ThemeName,
    TokenRole,
    get_color,
    get_font,
    get_size,
)


class TestPalettes:
    def test_all_themes_have_all_roles(self):
        for theme in ThemeName:
            palette = PALETTES[theme]
            for role in TokenRole:
                assert role in palette, f"{theme.value} missing {role.value}"

    def test_palette_has_max_7_colors(self):
        for theme in ThemeName:
            assert len(PALETTES[theme]) <= 7, f"{theme.value} has >7 colors"

    def test_rgb_values_in_range(self):
        for theme in ThemeName:
            for role, (r, g, b) in PALETTES[theme].items():
                assert 0 <= r <= 255, f"{theme.value}.{role.value} R={r}"
                assert 0 <= g <= 255, f"{theme.value}.{role.value} G={g}"
                assert 0 <= b <= 255, f"{theme.value}.{role.value} B={b}"


class TestFonts:
    def test_all_tokens_mapped(self):
        for token in FontToken:
            assert token in FONT_MAP

    def test_font_names_non_empty(self):
        for token, name in FONT_MAP.items():
            assert len(name) > 0, f"{token.value} has empty font name"


class TestSizes:
    def test_all_tokens_mapped(self):
        for token in SizeToken:
            assert token in SIZE_MAP

    def test_sizes_are_positive(self):
        for token, size in SIZE_MAP.items():
            assert size > 0, f"{token.value} size is {size}"

    def test_scale_is_descending(self):
        ordered = [SizeToken.XL, SizeToken.LG, SizeToken.MD, SizeToken.SM, SizeToken.XS]
        sizes = [SIZE_MAP[t] for t in ordered]
        for i in range(len(sizes) - 1):
            assert sizes[i] > sizes[i + 1], "Size scale not descending"

    def test_min_max_match_scale(self):
        assert FONT_SIZE_MIN == SIZE_MAP[SizeToken.XS]
        assert FONT_SIZE_MAX == SIZE_MAP[SizeToken.XL]


class TestGetters:
    def test_get_color(self):
        c = get_color(ThemeName.BUSINESS, TokenRole.ACCENT)
        assert c == (37, 99, 235)

    def test_get_font(self):
        assert get_font(FontToken.TITLE) == "BIZ UDPGothic"

    def test_get_size(self):
        assert get_size(SizeToken.XL) == 36.0
