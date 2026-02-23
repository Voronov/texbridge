# ==============================================================================
# Makefile for LaTeX <-> DOCX conversion
# ==============================================================================
#
# Usage:
#   make docx          — Convert LaTeX → DOCX (with images)
#   make latex          — Convert DOCX → LaTeX
#   make pdf            — Compile LaTeX → PDF (requires xelatex)
#   make extract        — Extract ZIP into src/
#   make roundtrip      — Full round-trip: LaTeX → DOCX → LaTeX → DOCX
#   make check          — Round-trip + comparison report
#   make clean          — Remove build artifacts
#   make clean-all      — Remove everything (build + extracted src)
#
# ==============================================================================

# --- Auto-detect ZIP file ---
ZIP         := $(wildcard *.zip)
ifeq ($(words $(ZIP)),0)
  $(error No .zip file found in project root)
endif
ifneq ($(words $(ZIP)),1)
  $(error Multiple .zip files found: $(ZIP). Keep only one or set ZIP= explicitly)
endif

# --- Derive name from ZIP ---
NAME        := $(basename $(ZIP))
SRC_DIR     := src
BUILD_DIR   := build
PREP_DIR    := $(BUILD_DIR)/preprocessed
MAIN_TEX    := $(SRC_DIR)/main.tex
PREP_TEX    := $(PREP_DIR)/main.tex
OUT_DOCX    := $(BUILD_DIR)/$(NAME).docx
OUT_PDF     := $(BUILD_DIR)/$(NAME).pdf
OUT_TEX_DIR := $(BUILD_DIR)/latex
FIX_SCRIPT  := scripts/fix_roundtrip.py
PREP_SCRIPT := scripts/preprocess_tex.py
REF_DOCX    := templates/reference.docx
CROSSREF_CFG:= pandoc-crossref.yaml

# --- Pandoc flags ---
PANDOC_COMMON := --wrap=none

PANDOC_TO_DOCX := $(PANDOC_COMMON) \
	--from=latex \
	--to=docx \
	--resource-path=.:sections \
	--extract-media=media \
	--number-sections \
	--reference-doc=../../$(REF_DOCX) \
	--filter pandoc-crossref \
	--metadata-file=../../$(CROSSREF_CFG)

PANDOC_TO_LATEX := $(PANDOC_COMMON) \
	--from=docx \
	--to=latex \
	--extract-media=$(OUT_TEX_DIR)/media \
	--wrap=none

# ==============================================================================
# Targets
# ==============================================================================

.PHONY: docx latex pdf extract clean clean-all help roundtrip check preprocess

help: ## Show this help
	@echo "Detected: $(ZIP) → $(NAME)"
	@echo ""
	@echo "Available targets:"
	@echo "  make extract    — Extract $(ZIP) into $(SRC_DIR)/"
	@echo "  make docx       — LaTeX → DOCX (output: $(NAME).docx)"
	@echo "  make latex      — DOCX → LaTeX + fixes (output: $(OUT_TEX_DIR)/$(NAME).tex)"
	@echo "  make roundtrip  — Full round-trip: LaTeX → DOCX → LaTeX → DOCX"
	@echo "  make check      — Round-trip + comparison report"
	@echo "  make pdf        — LaTeX → PDF  (output: $(OUT_PDF), requires xelatex)"
	@echo "  make clean      — Remove $(BUILD_DIR)/"
	@echo "  make clean-all  — Remove $(BUILD_DIR)/ and $(SRC_DIR)/"

# --- Extract ZIP ---
extract: $(MAIN_TEX)

$(MAIN_TEX): $(ZIP)
	@echo "📦 Extracting $(ZIP) → $(SRC_DIR)/"
	@mkdir -p $(SRC_DIR)
	unzip -o $(ZIP) -d $(SRC_DIR)
	@touch $@

# --- Pre-process LaTeX for pandoc compatibility ---
preprocess: $(PREP_TEX)

$(PREP_TEX): $(MAIN_TEX) $(wildcard $(SRC_DIR)/sections/*.tex)
	@echo "🔧 Pre-processing LaTeX for pandoc compatibility..."
	python3 $(PREP_SCRIPT) $(SRC_DIR) $(PREP_DIR)

# --- LaTeX → DOCX ---
docx: $(OUT_DOCX)

$(OUT_DOCX): $(PREP_TEX) $(wildcard $(SRC_DIR)/*.png $(SRC_DIR)/*.jpg)
	@echo "📝 Converting LaTeX → DOCX"
	@mkdir -p $(BUILD_DIR)
	cd $(PREP_DIR) && pandoc main.tex -o ../../$(OUT_DOCX) $(PANDOC_TO_DOCX) 2>&1 | grep -v "^$$" || true
	cp $(OUT_DOCX) $(NAME).docx
	@echo "✅ Output: $(NAME).docx"

# --- DOCX → LaTeX ---
latex: $(OUT_DOCX)
	@echo "📝 Converting DOCX → LaTeX"
	@mkdir -p $(OUT_TEX_DIR)
	pandoc $(OUT_DOCX) -o $(OUT_TEX_DIR)/$(NAME)_raw.tex $(PANDOC_TO_LATEX)
	@echo "🔧 Applying round-trip fixes..."
	python3 $(FIX_SCRIPT) $(OUT_TEX_DIR)/$(NAME)_raw.tex $(OUT_TEX_DIR)/$(NAME).tex --src-dir=$(SRC_DIR) --media-dir=$(OUT_TEX_DIR)/media
	@echo "✅ Output: $(OUT_TEX_DIR)/$(NAME).tex (raw: $(OUT_TEX_DIR)/$(NAME)_raw.tex)"

# --- LaTeX → PDF (requires xelatex) ---
pdf: $(OUT_PDF)

$(OUT_PDF): $(MAIN_TEX) $(wildcard $(SRC_DIR)/sections/*.tex)
	@echo "📄 Compiling LaTeX → PDF (xelatex)"
	@mkdir -p $(BUILD_DIR)
	cd $(SRC_DIR) && xelatex -output-directory=../$(BUILD_DIR) main.tex
	cd $(SRC_DIR) && xelatex -output-directory=../$(BUILD_DIR) main.tex
	@echo "✅ Output: $(OUT_PDF)"

# --- Full round-trip test: LaTeX → DOCX → LaTeX(fixed) → DOCX ---
roundtrip: latex
	@echo "🔄 Re-exporting fixed LaTeX → DOCX for verification"
	cd $(OUT_TEX_DIR) && pandoc $(NAME).tex -o ../$(NAME)_roundtrip.docx \
		--from=latex --to=docx --resource-path=. --extract-media=media_rt --wrap=none
	@echo "✅ Round-trip DOCX: $(BUILD_DIR)/$(NAME)_roundtrip.docx"

# --- Quick check: compare original vs round-trip DOCX ---
check: roundtrip
	@echo ""
	@echo "📊 Comparison:"
	@echo "  Original DOCX:    $$(du -h $(OUT_DOCX) | cut -f1)"
	@echo "  Round-trip DOCX:  $$(du -h $(BUILD_DIR)/$(NAME)_roundtrip.docx | cut -f1)"
	@echo ""
	@echo "  Original LaTeX equations:    $$(cat $(SRC_DIR)/sections/*.tex | grep -c 'begin{equation}')"
	@echo "  Round-trip LaTeX equations:   $$(grep -c 'begin{equation}' $(OUT_TEX_DIR)/$(NAME).tex)"
	@echo "  Original LaTeX labels:        $$(cat $(SRC_DIR)/sections/*.tex | grep -v '^\s*%%*' | grep -c 'label{fig\|label{tab')"
	@echo "  Round-trip LaTeX labels:       $$(grep -c 'label{fig\|label{tab' $(OUT_TEX_DIR)/$(NAME).tex)"
	@echo "  Original images:              $$(ls $(SRC_DIR)/*.png $(SRC_DIR)/*.jpg 2>/dev/null | wc -l | tr -d ' ')"
	@echo "  Round-trip images:             $$(ls $(OUT_TEX_DIR)/images/ 2>/dev/null | wc -l | tr -d ' ')"
	@echo ""

# --- Clean ---
clean:
	@echo "🧹 Cleaning $(BUILD_DIR)/"
	rm -rf $(BUILD_DIR)

clean-all: clean
	@echo "🧹 Cleaning $(SRC_DIR)/"
	rm -rf $(SRC_DIR)
