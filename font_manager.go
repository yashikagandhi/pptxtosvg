package pptxtosvg

import (
	"io/ioutil"
	"path/filepath"
	"strings"
	"sync"

	"golang.org/x/image/font"
	"golang.org/x/image/font/sfnt"
	"golang.org/x/image/math/fixed"
)

type FontKey struct {
	Family string
	Weight string
}

var (
	fontRegistry   map[FontKey]*sfnt.Font
	fontRegistryMu sync.RWMutex
	fontsLoaded    bool
)

func LoadFontsOnce(fontsDir string) {
	fontRegistryMu.Lock()
	defer fontRegistryMu.Unlock()
	if fontsLoaded {
		return
	}
	fontRegistry = make(map[FontKey]*sfnt.Font)
	files, err := ioutil.ReadDir(fontsDir)
	if err != nil {
		return
	}
	var buf sfnt.Buffer
	for _, f := range files {
		ext := strings.ToLower(filepath.Ext(f.Name()))
		if ext != ".ttf" && ext != ".otf" {
			continue
		}
		b, err := ioutil.ReadFile(filepath.Join(fontsDir, f.Name()))
		if err != nil {
			continue
		}
		ft, err := sfnt.Parse(b)
		if err != nil {
			continue
		}
		family, err := ft.Name(&buf, sfnt.NameIDFamily)
		if err != nil {
			continue
		}
		subFamily, err := ft.Name(&buf, sfnt.NameIDSubfamily)
		if err != nil {
			continue
		}

		key := FontKey{
			Family: strings.ToLower(family),
			Weight: strings.ToLower(subFamily),
		}
		fontRegistry[key] = ft
	}
	fontsLoaded = true
}

// GetStringWidth computes the accurate visual width of a text string in pixels based on the TTF metrics.
// If the font cannot be found, it falls back to a 0.48 multiplier approximation.
func GetStringWidth(text string, family string, fontSize int, isBold bool, isItalic bool) float64 {
	fontRegistryMu.RLock()
	defer fontRegistryMu.RUnlock()

	if !fontsLoaded || len(fontRegistry) == 0 {
		return float64(len(text)) * 0.48 * float64(fontSize)
	}

	searchFam := strings.ToLower(family)

	// Normalize PPT theme font identifiers (e.g. "+mn-lt" -> "calibri", "+mj-lt" -> "calibri light" roughly)
	if strings.Contains(searchFam, "+mn-lt") {
		searchFam = "calibri"
	}
	if strings.Contains(searchFam, "+mj-lt") {
		searchFam = "calibri light"
	}

	// Try to resolve the closest font variant
	var weight string
	if isBold && isItalic {
		weight = "bold italic"
	} else if isBold {
		weight = "bold"
	} else if isItalic {
		weight = "italic"
	} else {
		weight = "regular"
	}
	// Fallback mappings since subfamilies may differ like "Normal", "Negreta"
	ft, ok := fontRegistry[FontKey{searchFam, weight}]
	if !ok && weight != "regular" {
		ft, ok = fontRegistry[FontKey{searchFam, "regular"}]
	}
	if !ok {
		ft, ok = fontRegistry[FontKey{searchFam, "normal"}]
	}

	// Default to Calibri if totally unfound
	if !ok {
		ft, ok = fontRegistry[FontKey{"calibri", weight}]
		if !ok {
			ft, ok = fontRegistry[FontKey{"calibri", "regular"}]
		}
	}
	if !ok {
		// Absolute fallback
		for _, f := range fontRegistry {
			ft = f
			break
		}
		if ft == nil {
			return float64(len(text)) * 0.48 * float64(fontSize)
		}
	}

	var buf sfnt.Buffer
	var totalWidth fixed.Int26_6
	ppem := fixed.I(fontSize) // pixels per em (equals font size in px at standard 96dpi)

	for _, r := range text {
		idx, err := ft.GlyphIndex(&buf, r)
		if err != nil {
			continue // skip missing glyphs
		}
		adv, err := ft.GlyphAdvance(&buf, idx, ppem, font.HintingNone)
		if err == nil {
			totalWidth += adv
		}
	}

	// Convert fixed 26.6 layout back to float64 pixel length cleanly (divide by 64)
	return float64(totalWidth) / 64.0
}

// GetFontBaselineOffsetPct computes the baseline offset percentage (Ascent / (Ascent + abs(Descent))) for a given font.
// Returns 0.78 as a fallback if the font cannot be resolved or metrics are invalid.
func GetFontBaselineOffsetPct(family string, isBold bool, isItalic bool) float64 {
	fontRegistryMu.RLock()
	defer fontRegistryMu.RUnlock()

	if !fontsLoaded || len(fontRegistry) == 0 {
		return 0.78
	}

	searchFam := strings.ToLower(family)

	// Normalize PPT theme font identifiers
	if strings.Contains(searchFam, "+mn-lt") {
		searchFam = "calibri"
	} else if strings.Contains(searchFam, "+mj-lt") {
		searchFam = "calibri light"
	}

	var weight string
	if isBold && isItalic {
		weight = "bold italic"
	} else if isBold {
		weight = "bold"
	} else if isItalic {
		weight = "italic"
	} else {
		weight = "regular"
	}

	ft, ok := fontRegistry[FontKey{searchFam, weight}]
	if !ok && weight != "regular" {
		ft, ok = fontRegistry[FontKey{searchFam, "regular"}]
	}
	if !ok {
		ft, ok = fontRegistry[FontKey{searchFam, "normal"}]
	}

	if !ok {
		ft, ok = fontRegistry[FontKey{"calibri", weight}]
		if !ok {
			ft, ok = fontRegistry[FontKey{"calibri", "regular"}]
		}
	}
	if !ok {
		// Absolute fallback
		for _, f := range fontRegistry {
			ft = f
			break
		}
		if ft == nil {
			return 0.78
		}
	}

	var buf sfnt.Buffer
	metrics, err := ft.Metrics(&buf, fixed.I(100), font.HintingNone)
	if err != nil {
		return 0.78
	}

	asc := float64(metrics.Ascent)
	desc := float64(metrics.Descent)
	if desc < 0 {
		desc = -desc
	}
	total := asc + desc
	if total <= 0 || asc <= 0 {
		return 0.78
	}
	return asc / total
}
