# Female Disease Prevalence

> ## Big Idea
> **The same female-skewed health burden that shows up in Zambia's diabetes numbers runs through five of the most disabling diseases of modern life — and the world still designs medicine, research and screening as if it didn't.**

A research-to-presentation pipeline. Starts from the Diabesity Zambia STEPS-2017 finding (women carrying more non-communicable disease risk than men), broadens the question to *which other diseases skew female*, and ends in a polished 10-slide .pptx built to present in 3 minutes.

---

## 3-minute story summary

1. **Hook.** Zambia's 2017 STEPS survey: women were 24% more likely than men to carry multiple non-communicable disease risk factors; female diabetes prevalence 3.0% vs 2.1% in men. Sounds local. Isn't.
2. **Tension.** The "default patient" in medical research has historically been male. But the disease burden does not split evenly.
3. **Reveal.** Five conditions, female-to-male prevalence ratio:
   - Lupus (autoimmune) — **9:1**
   - Osteoporosis (musculoskeletal) — **4.5:1**
   - Migraine (neurological) — **3:1**
   - Alzheimer's (neurodegenerative) — **2:1**
   - Major depression (mental health) — **2:1**
4. **So-what.** Make sex-disaggregated reporting the default in every NCD surveillance round — starting with Zambia's next STEPS — so the burden women carry stops hiding in the totals.

The full narrative is in [`docs/story_structure.md`](docs/story_structure.md).

---

## How to run

```bash
pip install python-pptx
python scripts/build_presentation.py
```

The script reads [`data/disease_research.json`](data/disease_research.json) and writes [`output/FemaleDiseasPrevalence.pptx`](output/FemaleDiseasPrevalence.pptx).

To re-run the layout QA pass:

```bash
python scripts/qa_inspect.py
```

### A note on the build tool

The original brief specified `pptxgenjs` (Node). Node was not available on the build machine, so the script was implemented in Python with [`python-pptx`](https://python-pptx.readthedocs.io/). The output structure, palette, and layout rules are identical to the brief's specification. See [`docs/qa_notes.md`](docs/qa_notes.md) for the QA approach.

---

## Repo layout

```
female-disease-prevalence/
├── README.md                       ← you are here
├── data/
│   ├── disease_research.json       ← machine-readable research
│   └── disease_research.md         ← human-readable research
├── docs/
│   ├── story_structure.md          ← Big Idea + 3-minute outline
│   ├── qa_notes.md                 ← QA approach + findings + fixes
│   └── qa_notes_raw.txt            ← raw shape-geometry dump
├── scripts/
│   ├── build_presentation.py       ← build_presentation (python-pptx)
│   └── qa_inspect.py               ← geometry-level QA
└── output/
    └── FemaleDiseasPrevalence.pptx ← the deck
```

---

## Data sources

| # | Disease | Citation | URL |
|---|---------|----------|-----|
| 1 | Lupus (SLE) | *Sex- and gender-related differences in systemic lupus erythematosus: a scoping review*, PMC 2024 | https://pmc.ncbi.nlm.nih.gov/articles/PMC12204902/ |
| 2 | Osteoporosis | Sarafrazi N et al. *Osteoporosis or Low Bone Mass in Older Adults: United States, 2017–2018.* CDC NCHS Data Brief No. 405, 2021 | https://www.cdc.gov/nchs/products/databriefs/db405.htm |
| 3 | Migraine | Buse DC et al. *Sex differences in the prevalence, symptoms, and associated features of migraine — AMPP Study*, PubMed 2013 | https://pubmed.ncbi.nlm.nih.gov/23808666/ |
| 4 | Alzheimer's | *Alzheimer's disease as a women's health challenge*, PMC 2024 | https://pmc.ncbi.nlm.nih.gov/articles/PMC11106001/ |
| 5 | Depression | Albert PR. *Why is depression more prevalent in women?* J. Psychiatry Neurosci., PMC 2015 | https://pmc.ncbi.nlm.nih.gov/articles/PMC4478054/ |
| H | Zambia hook | Pengpid S, Peltzer K. *Prevalence and correlates of multiple non-communicable disease risk factors among adults in Zambia: STEPS 2017*, PMC 2021 | https://www.ncbi.nlm.nih.gov/pmc/articles/PMC7864270/ |

---

## Credits

Research, narrative and build by the author. Storytelling structure follows Cole Nussbaumer Knaflic's *Storytelling with Data* (Big Idea + 3-minute story).
