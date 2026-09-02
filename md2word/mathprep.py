"""Normalize AI Markdown math so Pandoc can emit Word OMML equations.

AI tools often emit LaTeX-like fragments such as (\\alpha), (\\hat{x}),
or bracket display blocks without $...$ delimiters. This module wraps those
fragments in standard TeX math delimiters. Unicode flattening is kept only
for the plain-text clipboard fallback.
"""

from __future__ import annotations

import re


_LATEX_COMMANDS: list[tuple[str, str]] = [
    (r"\leqslant", "≤"),
    (r"\geqslant", "≥"),
    (r"\rightarrow", "→"),
    (r"\leftarrow", "←"),
    (r"\Rightarrow", "⇒"),
    (r"\Leftarrow", "⇐"),
    (r"\Leftrightarrow", "⇔"),
    (r"\leftrightarrow", "↔"),
    (r"\subseteq", "⊆"),
    (r"\supseteq", "⊇"),
    (r"\subset", "⊂"),
    (r"\supset", "⊃"),
    (r"\emptyset", "∅"),
    (r"\varnothing", "∅"),
    (r"\partial", "∂"),
    (r"\nabla", "∇"),
    (r"\infty", "∞"),
    (r"\approx", "≈"),
    (r"\neq", "≠"),
    (r"\ne", "≠"),
    (r"\leq", "≤"),
    (r"\geq", "≥"),
    (r"\times", "×"),
    (r"\cdot", "·"),
    (r"\cdots", "⋯"),
    (r"\ldots", "…"),
    (r"\dots", "…"),
    (r"\quad", " "),
    (r"\qquad", "  "),
    (r"\alpha", "α"),
    (r"\beta", "β"),
    (r"\gamma", "γ"),
    (r"\delta", "δ"),
    (r"\epsilon", "ε"),
    (r"\varepsilon", "ε"),
    (r"\zeta", "ζ"),
    (r"\eta", "η"),
    (r"\theta", "θ"),
    (r"\vartheta", "ϑ"),
    (r"\iota", "ι"),
    (r"\kappa", "κ"),
    (r"\lambda", "λ"),
    (r"\mu", "μ"),
    (r"\nu", "ν"),
    (r"\xi", "ξ"),
    (r"\pi", "π"),
    (r"\rho", "ρ"),
    (r"\sigma", "σ"),
    (r"\tau", "τ"),
    (r"\upsilon", "υ"),
    (r"\phi", "φ"),
    (r"\varphi", "φ"),
    (r"\chi", "χ"),
    (r"\psi", "ψ"),
    (r"\omega", "ω"),
    (r"\Gamma", "Γ"),
    (r"\Delta", "Δ"),
    (r"\Theta", "Θ"),
    (r"\Lambda", "Λ"),
    (r"\Xi", "Ξ"),
    (r"\Pi", "Π"),
    (r"\Sigma", "Σ"),
    (r"\Upsilon", "Υ"),
    (r"\Phi", "Φ"),
    (r"\Psi", "Ψ"),
    (r"\Omega", "Ω"),
    (r"\sum", "∑"),
    (r"\prod", "∏"),
    (r"\int", "∫"),
    (r"\pm", "±"),
    (r"\mp", "∓"),
    (r"\in", "∈"),
    (r"\notin", "∉"),
    (r"\cup", "∪"),
    (r"\cap", "∩"),
    (r"\forall", "∀"),
    (r"\exists", "∃"),
    (r"\ell", "ℓ"),
    (r"\hbar", "ℏ"),
    (r"\mathbb{R}", "ℝ"),
    (r"\mathbb{N}", "ℕ"),
    (r"\mathbb{Z}", "ℤ"),
    (r"\mathbb{Q}", "ℚ"),
    (r"\mathbb{C}", "ℂ"),
    (r"\mathbf", ""),
    (r"\mathrm", ""),
    (r"\mathit", ""),
    (r"\mathcal", ""),
    (r"\operatorname", ""),
    (r"\text", ""),
    (r"\left", ""),
    (r"\right", ""),
    (r"\big", ""),
    (r"\Big", ""),
    (r"\bigg", ""),
    (r"\Bigg", ""),
    (r"\,", " "),
    (r"\;", " "),
    (r"\:", " "),
    (r"\!", ""),
    (r"\%", "%"),
    (r"\&", "&"),
    (r"\_", "_"),
    (r"\{", "{"),
    (r"\}", "}"),
]

_SUP_MAP = {
    "0": "⁰", "1": "¹", "2": "²", "3": "³", "4": "⁴",
    "5": "⁵", "6": "⁶", "7": "⁷", "8": "⁸", "9": "⁹",
    "+": "⁺", "-": "⁻", "=": "⁼", "(": "⁽", ")": "⁾",
    "n": "ⁿ", "i": "ⁱ", "T": "ᵀ",
}

_SUB_MAP = {
    "0": "₀", "1": "₁", "2": "₂", "3": "₃", "4": "₄",
    "5": "₅", "6": "₆", "7": "₇", "8": "₈", "9": "₉",
    "+": "₊", "-": "₋", "=": "₌", "(": "₍", ")": "₎",
    "a": "ₐ", "e": "ₑ", "o": "ₒ", "x": "ₓ",
    "i": "ᵢ", "j": "ⱼ", "k": "ₖ", "n": "ₙ", "m": "ₘ",
    "p": "ₚ", "r": "ᵣ", "s": "ₛ", "t": "ₜ", "u": "ᵤ", "v": "ᵥ",
}

_CHUNK_RE = re.compile(r"§§MD2WORD_CHUNK_(\d+)§§")


def _map_script(s: str, mapping: dict[str, str], prefix: str) -> str:
    """Map every char, or keep TeX-like prefix form if any char is unmapped."""
    out: list[str] = []
    for ch in s:
        if ch in mapping:
            out.append(mapping[ch])
        elif ch.isalpha() or ch.isdigit() or ch in "+-()=":
            return f"{prefix}{s}" if len(s) == 1 else f"{prefix}{{{s}}}"
        else:
            out.append(ch)
    return "".join(out)


def _to_sup(s: str) -> str:
    return _map_script(s, _SUP_MAP, "^")


def _to_sub(s: str) -> str:
    return _map_script(s, _SUB_MAP, "_")


def _apply_accents(expr: str) -> str:
    def comb(mark: str):
        def repl(m: re.Match[str]) -> str:
            body = m.group(1)
            if not body:
                return ""
            return body[0] + mark + body[1:]

        return repl

    patterns = [
        (r"\\widehat\{([^{}]+)\}", "\u0302"),
        (r"\\hat\{([^{}]+)\}", "\u0302"),
        (r"\\bar\{([^{}]+)\}", "\u0304"),
        (r"\\tilde\{([^{}]+)\}", "\u0303"),
        (r"\\vec\{([^{}]+)\}", "\u20D7"),
        (r"\\dot\{([^{}]+)\}", "\u0307"),
        (r"\\widehat\s+([A-Za-z])", "\u0302"),
        (r"\\hat\s+([A-Za-z])", "\u0302"),
        (r"\\bar\s+([A-Za-z])", "\u0304"),
        (r"\\tilde\s+([A-Za-z])", "\u0303"),
        (r"\\vec\s+([A-Za-z])", "\u20D7"),
    ]
    for pat, mark in patterns:
        expr = re.sub(pat, comb(mark), expr)
    return expr


def latex_math_to_unicode(expr: str) -> str:
    if not expr:
        return ""
    s = expr.strip().strip("$")
    s = re.sub(r"^\\\(|\\\)$", "", s)
    s = re.sub(r"^\\\[|\\\]$", "", s)

    s = re.sub(r"\\sqrt\[([^{}\]]+)\]\{([^{}]+)\}", r"√[\1](\2)", s)
    s = re.sub(r"\\sqrt\{([^{}]+)\}", r"√(\1)", s)
    s = re.sub(r"\\sqrt\s*([0-9]+(?:\.[0-9]+)?)", r"√(\1)", s)
    s = re.sub(r"\\sqrt\b", "√", s)

    s = _apply_accents(s)

    def frac_repl(m: re.Match[str]) -> str:
        a, b = m.group(1), m.group(2)
        if re.fullmatch(r"[A-Za-z0-9α-ωΑ-Ω]+", a) and re.fullmatch(
            r"[A-Za-z0-9α-ωΑ-Ω]+", b
        ):
            return f"{a}/{b}"
        return f"({a})/({b})"

    s = re.sub(r"\\frac\{([^{}]+)\}\{([^{}]+)\}", frac_repl, s)
    s = re.sub(r"\\dfrac\{([^{}]+)\}\{([^{}]+)\}", frac_repl, s)

    for _ in range(3):
        ns = re.sub(r"\^\{([^{}]+)\}", lambda m: _to_sup(m.group(1)), s)
        ns = re.sub(r"_\{([^{}]+)\}", lambda m: _to_sub(m.group(1)), ns)
        ns = re.sub(r"\^([A-Za-z0-9+\-=()])", lambda m: _to_sup(m.group(1)), ns)
        ns = re.sub(r"_([A-Za-z0-9+\-=()])", lambda m: _to_sub(m.group(1)), ns)
        if ns == s:
            break
        s = ns

    for cmd, rep in sorted(_LATEX_COMMANDS, key=lambda x: len(x[0]), reverse=True):
        if cmd in s:
            s = s.replace(cmd, rep)

    for _ in range(4):
        ns = re.sub(r"\{([^{}]*)\}", r"\1", s)
        if ns == s:
            break
        s = ns

    s = re.sub(r"[ \t]{2,}", "  ", s)
    s = re.sub(r"\s+,", ",", s)
    s = re.sub(r",\s*", ", ", s)
    s = re.sub(r"(?<![<>!=])=(?!=)", " = ", s)
    s = re.sub(r"(?<=[\wα-ωΑ-Ω)])\+(?=[\wα-ωΑ-Ω(√])", " + ", s)
    s = re.sub(r"(?<=[\wα-ωΑ-Ω)])<(?=[\wα-ωΑ-Ω(√])", " < ", s)
    s = re.sub(r"(?<=[\wα-ωΑ-Ω)])>(?=[\wα-ωΑ-Ω(√])", " > ", s)
    s = re.sub(r"[ \t]{2,}", " ", s)
    s = re.sub(r"\s+([,.;:])", r"\1", s)
    return s.strip()


def _protect_segments(
    text: str,
    chunks: list[str] | None = None,
    include_math: bool = False,
) -> tuple[str, list[str]]:
    if chunks is None:
        chunks = []

    def keep(m: re.Match[str]) -> str:
        chunks.append(m.group(0))
        return f"§§MD2WORD_CHUNK_{len(chunks) - 1}§§"

    text = re.sub(r"```[\s\S]*?```", keep, text)
    text = re.sub(r"`[^`\n]+`", keep, text)
    if include_math:
        text = re.sub(r"\$\$[\s\S]*?\$\$", keep, text)
        text = re.sub(r"(?<!\$)\$(?!\$)(?:\\.|[^$\n])+\$", keep, text)
    return text, chunks


def _restore_segments(text: str, chunks: list[str]) -> str:
    def restore(m: re.Match[str]) -> str:
        return chunks[int(m.group(1))]

    return _CHUNK_RE.sub(restore, text)


def _looks_like_math(body: str) -> bool:
    return bool(
        re.search(
            r"\\[A-Za-z]+|_|\^|\\frac|\\sum|\\sqrt|\\rho|\\alpha|\\beta|\\gamma|"
            r"\\Gamma|[=<>≤≥≈≠]|[α-ωΑ-Ω]|\\times|\\cdot|\\left|\\right",
            body,
        )
    )


def _as_display(body: str) -> str:
    inner = body.strip()
    if inner.startswith("$$") and inner.endswith("$$"):
        return "\n\n" + inner + "\n\n"
    return "\n\n$$\n" + inner + "\n$$\n\n"


def _as_inline(body: str) -> str:
    inner = body.strip().strip("$")
    return "$" + inner + "$"


def prepare_markdown_math(text: str) -> str:
    """Wrap AI-style math in $...$ / $$...$$ so Pandoc can emit OMML."""
    if not text:
        return text

    text, chunks = _protect_segments(text, include_math=True)

    text = re.sub(
        r"\\begin\{(?:equation|align|gather|displaymath)\*?\}([\s\S]*?)\\end\{(?:equation|align|gather|displaymath)\*?\}",
        lambda m: _as_display(m.group(1)),
        text,
    )
    text = re.sub(
        r"\\\[(.+?)\\\]",
        lambda m: _as_display(m.group(1)),
        text,
        flags=re.S,
    )
    text = re.sub(
        r"\\\((.+?)\\\)",
        lambda m: _as_inline(m.group(1)),
        text,
        flags=re.S,
    )

    def bracket_block(m: re.Match[str]) -> str:
        body = m.group(1).strip()
        if not body or not _looks_like_math(body):
            return m.group(0)
        return _as_display(body)

    text = re.sub(
        r"(?m)^[ \t]*\[[ \t]*\n([\s\S]*?)\n[ \t]*\][ \t]*$",
        bracket_block,
        text,
    )

    def bracket_line(m: re.Match[str]) -> str:
        body = m.group(1).strip()
        if not _looks_like_math(body):
            return m.group(0)
        return _as_display(body).strip()

    text = re.sub(r"(?m)^[ \t]*\[[ \t]*([^\]\n]+)[ \t]*\][ \t]*$", bracket_line, text)

    # Protect newly created / existing math before wrapping parentheticals.
    text, chunks = _protect_segments(text, chunks, include_math=True)

    def paren_math(m: re.Match[str]) -> str:
        inner = m.group(1)
        if not re.search(r"\\[A-Za-z]+|[_^]", inner):
            return m.group(0)
        return _as_inline(inner)

    text = re.sub(r"\(([^()\n]{1,200})\)", paren_math, text)

    def bare_latex_line(m: re.Match[str]) -> str:
        line = m.group(0)
        stripped = line.strip()
        if stripped.startswith("$$") or stripped.startswith("$"):
            return line
        if re.search(r"\\[A-Za-z]{2,}", stripped) and _looks_like_math(stripped):
            return _as_display(stripped).strip()
        return line

    text = re.sub(r"(?m)^[ \t]*\\[A-Za-z].+$", bare_latex_line, text)

    text = _restore_segments(text, chunks)
    text = re.sub(r"\n{4,}", "\n\n\n", text)
    return text


def strip_thematic_breaks(text: str) -> str:
    """Remove Markdown horizontal rules (--- / *** / ___) used as section dividers."""
    text = re.sub(r"(?m)^[ \t]*(?:-{3,}|\*{3,}|_{3,})[ \t]*$", "", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text


def plain_text_from_markdown(md: str) -> str:
    text = prepare_markdown_math(md)

    def display_plain(m: re.Match[str]) -> str:
        return "\n" + latex_math_to_unicode(m.group(1)) + "\n"

    def inline_plain(m: re.Match[str]) -> str:
        return latex_math_to_unicode(m.group(1))

    text, chunks = _protect_segments(text)
    text = re.sub(r"\$\$(.+?)\$\$", display_plain, text, flags=re.S)
    text = re.sub(r"(?<!\$)\$(?!\$)(.+?)(?<!\$)\$(?!\$)", inline_plain, text, flags=re.S)
    text = _restore_segments(text, chunks)

    text = re.sub(r"(?m)^#{1,6}\s+", "", text)
    text = re.sub(r"\*\*([^*]+)\*\*", r"\1", text)
    text = re.sub(r"\*([^*]+)\*", r"\1", text)
    text = re.sub(r"`([^`]+)`", r"\1", text)
    text = re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)
    text = re.sub(r"(?m)^>\s?", "", text)
    text = re.sub(r"(?m)^(\s*)[-*+]\s+", r"\1• ", text)
    return text.strip() + "\n"
