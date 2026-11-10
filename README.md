# pptxtosvg

A pure Go library to convert Microsoft PowerPoint presentations (`.pptx`) into Scalable Vector Graphics (`.svg`), with speaker notes extraction to Markdown.

## Features

- **Pure Go / Zero CGO**: Runs anywhere Go runs without external runtime dependencies (no LibreOffice, no Python, no Node.js).
- **Standards-Compliant OOXML Parsing**: Parses presentation slides, master slides, layout slides, themes, and shapes according to ECMA-376 OOXML specifications.
- **Rich Vector Rendering**:
  - Shapes, paths, curves, polygons, and custom geometries.
  - Linear/radial gradients, pattern fills, and drop shadows.
  - Word art, rich text runs, formatting (bold, italic, underline, strikethrough, subscript, superscript).
  - Rotations, flips (horizontal/vertical), and hierarchical group transforms.
- **Tables & Charts**:
  - Comprehensive table rendering with theme and custom table styles.
  - Bar, column, and series charts rendered directly into clean SVG.
- **Exact Font Metrics & Text Wrapping**:
  - Optional TTF/OTF font loading for pixel-perfect glyph metrics and text wrapping.
  - Graceful fallback text-width heuristics when font files are not provided.
- **Speaker Notes**:
  - Extracts slide notes into formatted Markdown.

## Installation

```bash
go get github.com/yashikagandhi/pptxtosvg
```

## Quick Start

### Convert a PPTX file

```go
package main

import (
	"fmt"
	"log"
	"os"

	"github.com/yashikagandhi/pptxtosvg"
)

func main() {
	// Optional: load TTF fonts for exact text measurements
	pptxtosvg.LoadFontsOnce("./fonts")

	// Convert all slides in a PPTX file
	results, err := pptxtosvg.ConvertPPTXFile("presentation.pptx")
	if err != nil {
		log.Fatalf("Conversion failed: %v", err)
	}

	for _, slide := range results {
		filename := fmt.Sprintf("slide%d.svg", slide.SlideNum)
		os.WriteFile(filename, []byte(slide.SVG), 0644)

		if slide.NotesMD != "" {
			notesFile := fmt.Sprintf("notes%d.md", slide.SlideNum)
			os.WriteFile(notesFile, []byte(slide.NotesMD), 0644)
		}
	}
}
```

### Convert from memory (`[]byte`) or `zip.Reader`

```go
// From byte slice (e.g. HTTP upload handler)
results, err := pptxtosvg.ConvertPPTXBytes(fileBytes)

// Or from an existing zip.Reader
results, err := pptxtosvg.ConvertPPTX(zipReader)
```

### Low-level Slide Conversion

```go
// Convert an individual slide's raw XML
svg := pptxtosvg.ConvertSlideToSVG(slideXMLBytes, slideNum, zipReader)

// Extract notes for an individual slide
notes := pptxtosvg.ExtractSlideNotes(slideNum, zipReader)
```

## License

Apache License 2.0. See [LICENSE](LICENSE) for details.