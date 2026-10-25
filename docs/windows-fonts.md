# Headless fonts on Windows

Windows resolves fonts directly from TTF/OTF/TTC metadata in `%WINDIR%/Fonts` and
`%LOCALAPPDATA%/Microsoft/Windows/Fonts`; no `fc-match` installation is needed.
The resolver uses the same open-source Office-font substitutions and fallback
families as the browser editor, and chooses the closest weight/italic face.

Install Noto Sans SC (or Noto Sans CJK SC) for Chinese plus Carlito, Caladea or
Liberation for Latin text. Microsoft Office fonts are not selected as substitutes.
An installation with no supported open-source fonts fails with instructions.

Optionally set `PPT_FONT_DIRS` to a directory containing these fonts before
starting the renderer, using `;` between directories on Windows (`:` on Unix).
Set the same variable in the AI PPT Go backend for matching browser rendering.
When this variable is set, Unix also uses native discovery. Otherwise Unix
continues to use its existing fontconfig resolver. Restart after installing fonts.

The Node resolver caches file matches. It reads names from the font files and
ignores unreadable/corrupt files; renaming an Office font to an allowed family
does not make it eligible.
