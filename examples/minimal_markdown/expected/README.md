# expected/

This directory explains how to generate and inspect the **reference visual
output** of the `minimal_markdown` example on your own machine. For privacy
reasons, the repository does not commit any screenshots produced from a real
school template — generate them locally instead.

## How to generate the preview

On Windows with Microsoft Word and pywin32 installed:

```bash
python examples/minimal_markdown/run_example.py
```

This produces `final.pdf` and `final.preview.png` inside
`examples/minimal_markdown/`. Without Word (macOS / Linux), the run stops
at `final.docx`, which you can still open in Word / LibreOffice to verify.

To preview specific pages of a real thesis template run:

```bash
python examples/minimal_markdown/run_example.py \
    --template /path/to/your-real-thesis-template.docx \
    --preview-pages "20"
```

`--preview-pages` accepts any page range or list understood by
`render_pdf_preview.py` (e.g. `1-4`, `7,12,13,14`).

## What a good body page looks like

A representative body page of `final.pdf` should show, all at once:

- an embedded PNG figure from `assets/pipeline_diagram.png`,
- a generated figure caption (centered, with a number prefix),
- Chinese/English body paragraphs in the template's body style,
- a display equation that survives the round-trip,
- a Markdown table converted into a three-line table,
- a second caption-like block for caption-rule regression coverage.

## pandoc baseline (optional comparison)

To see why `pandoc --reference-doc` alone is not enough, run the same
`sample.md` directly through pandoc against the same template and compare:

```bash
pandoc examples/minimal_markdown/sample.md \
    --resource-path examples/minimal_markdown \
    --reference-doc /path/to/your-real-thesis-template.docx \
    -o pandoc_baseline.docx
python skills/texpdf2word/scripts/finalize_word_docx.py \
    pandoc_baseline.docx --pdf
```

Typical differences: the pandoc baseline is more likely to split the
figure/equation/table/caption region across pages, drop the `图 X.Y` caption
number prefix, and leave tables borderless. It can reuse style definitions,
but it does not understand the target template's semantic layout rules, nor
does it apply this project's adaptive three-line-table and caption/body
remapping logic.

## Privacy notes

- Do not commit your real `.docx` templates, generated `final.pdf`,
  or preview screenshots that contain a real institution's cover pages,
  logos, or signatures into a public fork. They are covered by `.gitignore`
  (e.g. `final.preview.png`, `sample_template.docx`).
- `examples/minimal_markdown/assets/pipeline_diagram.png` is the only image
  tracked in git: a self-made, generic flow diagram used as test input.
