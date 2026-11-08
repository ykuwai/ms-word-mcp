"""LaTeX to UnicodeMath. Every expected string here was built up by Word for
Mac 16.114 and read back with Linearize as the same structure (the letters
come back in math italic)."""

import pytest

from ms_word_mcp.tools.latex import LatexError, latex_to_unicodemath as convert


@pytest.mark.parametrize("latex, unicodemath", [
    (r"\frac{a}{b}", "a/b"),
    (r"E=mc^2", "E=mc^2"),
    (r"\sqrt{x^2+1}", "√(x^2+1)"),
    (r"\sqrt[3]{x}", "√(3&x)"),
    (r"\int_0^\infty e^{-x^2}\,dx = \frac{\sqrt{\pi}}{2}", "∫_0^∞▒〖e^(-x^2)\u2009dx〗=(√π)/2"),
    (r"\sum_{i=1}^{n} i = \frac{n(n+1)}{2}", "∑_(i=1)^n▒i=(n(n+1))/2"),
    (r"\alpha+\beta\le\gamma", "α+β≤γ"),
    (r"\left(\frac{a}{b}\right)^2", "〖(a/b)〗^2"),
    (r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}", "(■(a&b@c&d))"),
    (r"\hat{x}+\vec{v}", "x̂+v⃗"),
    (r"\sin^2\theta+\cos^2\theta=1", "sin^2⁡θ+cos^2⁡θ=1"),
    (r"\lim_{x\to 0}\frac{\sin x}{x}=1", "lim_(x→0)⁡(sin⁡x)/x=1"),
    (r"f(x)=\begin{cases} 1 & x>0 \\ 0 & \text{otherwise} \end{cases}",
     'f(x)={█(1&x>0@0&"otherwise")'),
    (r"\begin{aligned} a &= b+c \\ &= d \end{aligned}", "█(a&=b+c@&=d)"),
    (r"$$x^2$$", "x^2"),
])
def test_conversions_word_builds(latex, unicodemath):
    assert convert(latex) == unicodemath


def test_symbols_come_from_words_own_list():
    from ms_word_mcp.tools import latex

    table = latex.symbols()
    assert table["rightarrow"] == "→" and table["hbar"] == "ℏ"


@pytest.mark.parametrize("bad", [r"\frac{a}", r"\nosuchcommand x", r"{x", r"\left( x",
                                 r"\begin{tabular}x\end{tabular}"])
def test_what_it_cannot_read_is_named(bad):
    with pytest.raises(LatexError):
        convert(bad)


def test_a_control_space_is_a_space():
    assert convert(r"x = \frac{1}{2},\ -2") == "x=1/2, -2"
