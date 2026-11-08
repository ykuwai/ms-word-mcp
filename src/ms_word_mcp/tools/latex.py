"""LaTeX to UnicodeMath, so Word can build an equation from either.

Word builds equations from UnicodeMath when code puts them in; it reads LaTeX
only through a ribbon setting that its object model does not reach, and its
Math AutoCorrect (\\alpha -> α) does not act on text code inserts (both
measured, docs/PROBES.md). So LaTeX is turned into UnicodeMath here, and
Word builds that.

The control words that stand for one symbol (\\alpha, \\rightarrow, \\le ...)
are looked up in Word's own Math AutoCorrect list, saved from Word in
docs/reference/math-autocorrect.json, with a few common ones built in for
when that file is missing. The structures follow the UnicodeMath
specification (Unicode Technical Note 28): a/b fractions, √(n&x) roots,
^( ) and _( ) scripts, ∑_(..)^(..)▒ operators, ■( & @ ) matrices.
"""

from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path

# Used when docs/reference/math-autocorrect.json is not there (an installed
# package has no docs/). Values from Word's Math AutoCorrect list.
_FALLBACK_SYMBOLS = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ϵ",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "pi": "π", "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ", "sigma": "σ",
    "varsigma": "ς", "tau": "τ", "upsilon": "υ", "phi": "ϕ", "varphi": "φ",
    "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ", "Pi": "Π",
    "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ", "Omega": "Ω",
    "infty": "∞", "partial": "∂", "nabla": "∇", "pm": "±", "mp": "∓",
    "times": "×", "div": "÷", "cdot": "⋅", "cdots": "⋯", "ldots": "…", "dots": "…",
    "vdots": "⋮", "ddots": "⋱", "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥",
    "ne": "≠", "neq": "≠", "approx": "≈", "equiv": "≡", "sim": "∼", "simeq": "≃",
    "propto": "∝", "in": "∈", "notin": "∉", "subset": "⊂", "supset": "⊃",
    "subseteq": "⊆", "supseteq": "⊇", "cup": "∪", "cap": "∩", "forall": "∀",
    "exists": "∃", "neg": "¬", "wedge": "∧", "vee": "∨", "to": "→",
    "rightarrow": "→", "leftarrow": "←", "Rightarrow": "⇒", "Leftarrow": "⇐",
    "leftrightarrow": "↔", "Leftrightarrow": "⇔", "mapsto": "↦", "ll": "≪",
    "gg": "≫", "angle": "∠", "perp": "⊥", "parallel": "∥", "circ": "∘",
    "degree": "°", "hbar": "ℏ", "ell": "ℓ", "Re": "ℜ", "Im": "ℑ", "aleph": "ℵ",
    "emptyset": "∅", "therefore": "∴", "because": "∵", "prime": "′",
    "sum": "∑", "prod": "∏", "coprod": "∐", "int": "∫", "iint": "∬", "iiint": "∭",
    "oint": "∮", "bigcup": "⋃", "bigcap": "⋂",
    "langle": "⟨", "rangle": "⟩", "lfloor": "⌊", "rfloor": "⌋", "lceil": "⌈",
    "rceil": "⌉", "lbrace": "{", "rbrace": "}", "vert": "|", "Vert": "‖",
    "quad": " ", "qquad": "  ", ",": " ", ";": " ",
    ":": " ", "!": "",
}

# Large operators: the operand follows ▒ in UnicodeMath.
_NARY = {"sum", "prod", "coprod", "int", "iint", "iiint", "oint", "bigcup", "bigcap"}

# Function names Word typesets upright when they are written as plain words.
_FUNCTIONS = {"sin", "cos", "tan", "cot", "sec", "csc", "arcsin", "arccos", "arctan",
              "sinh", "cosh", "tanh", "log", "ln", "lg", "exp", "lim", "max", "min",
              "sup", "inf", "det", "dim", "ker", "deg", "arg", "gcd", "Pr", "mod"}

# Accents, as the combining character UnicodeMath puts after the base.
_ACCENTS = {"hat": "̂", "widehat": "̂", "tilde": "̃", "widetilde": "̃",
            "bar": "̅", "overline": "̅", "vec": "⃗", "dot": "̇",
            "ddot": "̈", "check": "̌", "acute": "́", "grave": "̀",
            "breve": "̆"}

_MATRIX_ENVIRONMENTS = {"matrix": ("", ""), "pmatrix": ("(", ")"), "bmatrix": ("[", "]"),
                        "Bmatrix": ("{", "}"), "vmatrix": ("|", "|"), "Vmatrix": ("‖", "‖"),
                        "cases": ("{", "")}

_REFERENCE = Path(__file__).resolve().parents[3] / "docs" / "reference" / "math-autocorrect.json"


@lru_cache(maxsize=1)
def symbols() -> dict[str, str]:
    """Control word (without the backslash) -> symbol, Word's list first."""
    table = dict(_FALLBACK_SYMBOLS)
    try:
        saved = json.loads(_REFERENCE.read_text(encoding="utf-8"))["entries"]
    except (OSError, ValueError, KeyError):
        return table
    for name, value in saved.items():
        if name.startswith("\\") and isinstance(value, str) and len(value) <= 3:
            table[name[1:]] = value
    return table


class LatexError(ValueError):
    """LaTeX this converter does not understand; the message says where."""


def _group(text: str) -> str:
    """Wrap in parentheses unless it is one character, so a/b, x^2 stay plain."""
    text = text.strip()
    if len(text) == 1 or (text.isalnum() and len(text) <= 1):
        return text
    return f"({text})"


class _Parser:
    def __init__(self, source: str):
        self.s = source
        self.i = 0

    def error(self, message: str) -> LatexError:
        return LatexError(f"{message} at position {self.i} of the LaTeX: "
                          f"...{self.s[max(0, self.i - 15):self.i + 15]}...")

    def peek(self) -> str:
        return self.s[self.i] if self.i < len(self.s) else ""

    def skip_space(self) -> None:
        while self.peek().isspace():
            self.i += 1

    def control_word(self) -> str:
        """The name after a backslash (the backslash already consumed)."""
        start = self.i
        if self.peek().isalpha():
            while self.peek().isalpha():
                self.i += 1
        else:
            self.i += 1  # \, \; \{ \} \\ and the like
        return self.s[start:self.i]

    def argument(self) -> str:
        """A {group} or a single token, converted."""
        self.skip_space()
        if self.peek() == "{":
            self.i += 1
            out = self.sequence(until="}")
            if self.peek() != "}":
                raise self.error("A { is not closed")
            self.i += 1
            return out
        if not self.peek():
            raise self.error("An argument is missing")
        return self.atom()

    def optional(self) -> str | None:
        self.skip_space()
        if self.peek() != "[":
            return None
        self.i += 1
        out = self.sequence(until="]")
        self.i += 1
        return out

    def scripts(self) -> str:
        """Any ^ and _ that follow, in UnicodeMath."""
        out = ""
        while True:
            self.skip_space()
            ch = self.peek()
            if ch not in ("^", "_"):
                return out
            self.i += 1
            out += ch + _group(self.argument())

    def atom(self) -> str:
        ch = self.peek()
        if ch == "\\":
            self.i += 1
            return self.command(self.control_word())
        if ch == "{":
            self.i += 1
            out = self.sequence(until="}")
            if self.peek() != "}":
                raise self.error("A { is not closed")
            self.i += 1
            return _group(out)
        self.i += 1
        if ch == "'":
            return "′"
        return ch

    def command(self, name: str) -> str:
        table = symbols()
        if name == " ":
            return " "  # "\ ", a space in math
        if name == "frac" or name == "dfrac" or name == "tfrac":
            num, den = self.argument(), self.argument()
            return f"{_group(num)}/{_group(den)}"
        if name == "binom":
            top, bottom = self.argument(), self.argument()
            return f"({top}¦{bottom})"
        if name == "sqrt":
            index = self.optional()
            body = self.argument()
            return f"√({index}&{body})" if index else f"√{_group(body)}"
        if name in ("text", "mathrm", "textrm", "operatorname", "mbox"):
            self.skip_space()
            if self.peek() != "{":
                raise self.error(f"\\{name} needs {{...}}")
            start = self.i + 1
            depth = 0
            while self.i < len(self.s):
                if self.s[self.i] == "{":
                    depth += 1
                elif self.s[self.i] == "}":
                    depth -= 1
                    if depth == 0:
                        break
                self.i += 1
            words = self.s[start:self.i]
            self.i += 1
            if name in ("operatorname",) or words in _FUNCTIONS:
                return words + "⁡"
            return f'"{words}"'
        if name in ("mathbf", "boldsymbol", "mathit", "mathcal", "mathbb", "mathsf"):
            return self.argument()  # Word's math alphabets are not reached from code
        if name in _ACCENTS:
            return _group(self.argument()) + _ACCENTS[name]
        if name == "left":
            self.skip_space()
            opening = self.delimiter()
            body = self.sequence(until="\\right")
            if not self.s.startswith("\\right", self.i):
                raise self.error("\\left has no \\right")
            self.i += len("\\right")
            self.skip_space()
            closing = self.delimiter()
            return f"〖{opening}{body}{closing}〗" if opening or closing else body
        if name == "begin":
            return self.environment()
        if name in _NARY:
            symbol = table.get(name, name)
            limits = self.scripts()
            self.skip_space()
            operand = self.operand()
            # The operand is grouped with 〖 〗, which Word does not draw; ( )
            # would show as parentheses (seen on the page).
            body = operand.strip()
            if len(body) > 1:
                body = f"〖{body}〗"
            return f"{symbol}{limits}▒{body}"
        if name in _FUNCTIONS:
            return name + self.scripts() + "⁡"
        if name in table:
            return table[name]
        raise self.error(f"\\{name} is not a command this converter knows; write it in "
                         "UnicodeMath instead")

    def delimiter(self) -> str:
        ch = self.peek()
        self.i += 1
        if ch == ".":
            return ""
        if ch == "\\":
            name = self.control_word()
            if name in ("{", "}"):
                return name
            return symbols().get(name, "")
        return ch

    def operand(self) -> str:
        """What a large operator applies to: the rest up to + - = or the end."""
        parts = []
        while self.i < len(self.s):
            self.skip_space()
            ch = self.peek()
            if ch in ("+", "-", "=", "}", "&", "") or self.s.startswith("\\\\", self.i) \
                    or self.s.startswith("\\right", self.i) \
                    or self.s.startswith("\\end", self.i):
                break
            parts.append(self.atom() + self.scripts())
        return "".join(parts)

    def environment(self) -> str:
        self.skip_space()
        if self.peek() != "{":
            raise self.error("\\begin needs {name}")
        end = self.s.index("}", self.i)
        name = self.s[self.i + 1:end]
        self.i = end + 1
        if name not in _MATRIX_ENVIRONMENTS and name not in ("aligned", "align", "align*",
                                                            "gathered", "split"):
            raise self.error(f"The environment {name} is not one this converter knows")
        rows, row, cell_start = [], [], True
        current = []
        while True:
            if self.i >= len(self.s):
                raise self.error(f"\\begin{{{name}}} has no \\end")
            if self.s.startswith("\\end", self.i):
                self.i += len("\\end")
                self.skip_space()
                close = self.s.index("}", self.i)
                self.i = close + 1
                break
            if self.s.startswith("\\\\", self.i):
                self.i += 2
                row.append("".join(current))
                rows.append(row)
                row, current = [], []
                continue
            if self.peek() == "&":
                self.i += 1
                row.append("".join(current))
                current = []
                continue
            current.append(self.atom() + self.scripts())
        if current or row:
            row.append("".join(current))
            rows.append(row)
        if name in _MATRIX_ENVIRONMENTS:
            opening, closing = _MATRIX_ENVIRONMENTS[name]
            body = "@".join("&".join(cell.strip() for cell in r) for r in rows)
            if name == "cases":
                return f"{{█({body})"
            return f"{opening}■({body}){closing}" if opening else f"■({body})"
        # Aligned equations: an equation array, & marks the alignment point.
        return "█(" + "@".join("&".join(cell.strip() for cell in r) for r in rows) + ")"

    def sequence(self, until: str = "") -> str:
        parts = []
        while self.i < len(self.s):
            if until and self.s.startswith(until, self.i):
                break
            ch = self.peek()
            if ch.isspace():
                self.i += 1
                if parts and parts[-1] != " ":
                    parts.append(" ")
                continue
            if ch == "}" and until != "}":
                raise self.error("A } has no {")
            parts.append(self.atom() + self.scripts())
        return "".join(parts).strip()


def latex_to_unicodemath(latex: str) -> str:
    """UnicodeMath for `latex`. Raises LatexError naming what it cannot read."""
    source = latex.strip()
    for wrapper in (("$$", "$$"), ("\\[", "\\]"), ("$", "$"), ("\\(", "\\)")):
        if source.startswith(wrapper[0]) and source.endswith(wrapper[1]) \
                and len(source) > len(wrapper[0]) + len(wrapper[1]):
            source = source[len(wrapper[0]):-len(wrapper[1])].strip()
            break
    parser = _Parser(source)
    out = parser.sequence()
    if parser.i < len(source):
        raise parser.error("Unexpected text")
    return out
