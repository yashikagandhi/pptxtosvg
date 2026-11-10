package pptxtosvg

import (
	"math"
	"sync"
)

// ====================================================================
// Render config for PPTX-to-SVG: only layout constants from OOXML spec.
// All per-shape values (font size, color, anchor, fill) come from PPT XML
// (shape, theme, layout, presentation.xml).
// ====================================================================

// RenderConfig holds layout parameters and document-level defaults from XML.
// DefaultFontSizePx is set from presentation.xml defaultTextStyle when available.
type RenderConfig struct {
	TextLayout        TextLayoutConfig
	DefaultFontSizePx int // from presentation.xml defaultTextStyle (set during conversion; 0 if absent)
}

// TextLayoutConfig holds OOXML-spec and typographic constants used for
// line height, wrap width, and default text insets (bodyPr lIns/rIns/tIns/bIns).
type TextLayoutConfig struct {
	LineHeightFactor float64 // multiplier on font size when lnSpc not in XML (1.2)
	CharWidthRatio   float64 // approximate char-width / font-size for wrap (~0.48)
	MinCharsPerLine  int     // floor for wrap calculation
	DefaultInsets    Insets  // OOXML default body insets in EMU (91440, 91440, 45720, 45720)
}

// Insets stores text-box insets in EMU (from OOXML spec).
type Insets struct {
	Left   int
	Right  int
	Top    int
	Bottom int
}

// --------------- helpers ---------------

func (cfg *RenderConfig) lineHeight(fontSize int) int {
	lh := math.Round(float64(fontSize) * cfg.TextLayout.LineHeightFactor)
	if lh < 1 {
		lh = float64(fontSize) + 4
	}
	return int(lh)
}

func (cfg *RenderConfig) maxCharsPerLine(widthPx, fontSize int) int {
	if fontSize <= 0 {
		fontSize = 12
	}
	charW := cfg.TextLayout.CharWidthRatio * float64(fontSize)
	if charW < 1 {
		charW = 1
	}
	n := int(float64(widthPx) / charW)
	if n < cfg.TextLayout.MinCharsPerLine {
		n = cfg.TextLayout.MinCharsPerLine
	}
	return n
}

func (cfg *RenderConfig) insetsPx() (lIns, rIns, tIns, bIns int) {
	lIns = int(math.Round(float64(cfg.TextLayout.DefaultInsets.Left) / emuPerPx))
	rIns = int(math.Round(float64(cfg.TextLayout.DefaultInsets.Right) / emuPerPx))
	tIns = int(math.Round(float64(cfg.TextLayout.DefaultInsets.Top) / emuPerPx))
	bIns = int(math.Round(float64(cfg.TextLayout.DefaultInsets.Bottom) / emuPerPx))
	return
}

// OOXML default text body insets (EMU): left/right 91440, top/bottom 45720.
var defaultOOXMLInsets = Insets{Left: 91440, Right: 91440, Top: 45720, Bottom: 45720}

func defaultRenderConfig() *RenderConfig {
	return &RenderConfig{
		TextLayout: TextLayoutConfig{
			LineHeightFactor: 1.2,
			CharWidthRatio:   0.48,
			MinCharsPerLine:  8,
			DefaultInsets:    defaultOOXMLInsets,
		},
	}
}

var (
	globalRenderConfig *RenderConfig
	renderConfigOnce   sync.Once
)

func getRenderConfig() *RenderConfig {
	renderConfigOnce.Do(func() {
		globalRenderConfig = defaultRenderConfig()
	})
	return globalRenderConfig
}
