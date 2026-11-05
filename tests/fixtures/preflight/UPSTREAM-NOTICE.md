Copied from superdoc-macros at b6fa9dc1dde5a82ea2028ea582845454d083686b.
`office-macros.docm` is the upstream `with-macros.docm` without modifications.
The upstream package and its generator are MIT licensed; see LICENSE-superdoc-macros.
The embedded Office-produced VBA project has the BSD notice below.

## Local variants

- `office-macros-repair.docm` uses the same Office-produced project. Its first
  `word/document.xml` run has `<w:rPr><w:bCs/></w:rPr>` inserted immediately after
  `<w:r>`, to exercise the existing complex-script repair while preserving VBA.
- `auto-open-macros.docm` uses that modified document and a **synthetic** VBA
  project, generated with `buildVbaProjectBin` from upstream
  `tests/vba-fixtures.ts` at the pinned revision above (Git blob
  `5ca707ced1505adb2547b24a7fed06ff8f620a12`). Parameters are `codePage: 1255`,
  `moduleName: 'Module1'`, `projectLines: ['Module=Module1']`, `layout: 'mini'`.
  Its source, encoded with that generator's `encodeCp1255`, is:

  ```vb
  Attribute VB_Name = "Module1"
  Sub AutoOpen()
  ' בדיקת תעבורה בלבד
  End Sub
  ```

  Every source line, including the last, ends with CRLF. There is no macro body;
  the test parses the auto-run name, Hebrew source and warnings without execution.

Both variants are repackaged with Python `zipfile`, DEFLATE level 6, original
entry order and timestamp `(2026, 1, 1, 0, 0, 0)`. No other uncompressed entry is
modified. These are small transport/repair fixtures, not real user documents.

The original upstream fixture notice follows unchanged.

# Test fixtures

Binary inputs for the VBA import tests (`tests/vba-import.test.ts`).

## `vbaProject.bin`

A genuine, Excel-produced VBA project containing one code module
(`Module1.say_hello`) plus the usual `ThisWorkbook` and `Sheet1` document
modules.

Taken from the test suite of [XlsxWriter](https://github.com/jmcnamara/XlsxWriter)
(`xlsxwriter/test/comparison/xlsx_files/vbaProject02.bin`), which is
distributed under the BSD 2-Clause License:

> Copyright (c) 2013-2025, John McNamara <jmcnamara@cpan.org>
>
> Redistribution and use in source and binary forms, with or without
> modification, are permitted provided that the following conditions are met:
>
> 1. Redistributions of source code must retain the above copyright notice,
>    this list of conditions and the following disclaimer.
> 2. Redistributions in binary form must reproduce the above copyright notice,
>    this list of conditions and the following disclaimer in the documentation
>    and/or other materials provided with the distribution.
>
> THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
> AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
> IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
> ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
> LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
> CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
> SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
> INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
> CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
> ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
> POSSIBILITY OF SUCH DAMAGE.

Why a real file rather than a generated one: the container's allocation
tables, the mini-stream layout, the declared code page and the compressed
module streams are all shaped the way Office actually writes them. A fixture
we generated ourselves would only prove the readers agree with our own
writer. Both are used — see `tests/vba-fixtures.ts` for the synthetic
projects, which cover the cases this file cannot reach (a Hebrew code page,
an auto-run macro).

The macro it contains is inert: an empty `Sub say_hello()`. Nothing in the
toolkit executes VBA.

## `with-macros.docm`

A minimal macro-enabled Word package wrapping the `vbaProject.bin` above:
`[Content_Types].xml` with the `macroEnabled.main` override, a one-paragraph
`word/document.xml`, and the `vbaProject` relationship that points at the
macro part. Every entry is deflate-compressed, so reading it exercises
inflation rather than a stored-entry passthrough.

Written for this repository.

## `no-macros.docx`

The same minimal package with no macro part — the negative case for
`extractVbaFromDocx`, which must answer `no-macros` rather than fail.

Written for this repository.
