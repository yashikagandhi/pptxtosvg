package pptxtosvg

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"path"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

func writeDebugLog(msg string, data map[string]interface{}, hypothesisId string) {
	// Debug logging hook; no-op in standalone library
}

// SlideResult holds the conversion result for a single slide.
type SlideResult struct {
	SlideNum int    `json:"slide_num"`
	SVG      string `json:"svg"`
	NotesMD  string `json:"notes_md,omitempty"`
}

// ConvertPPTX parses a PPTX zip archive and converts all slides to SVG and notes to Markdown.
func ConvertPPTX(zr *zip.Reader) ([]SlideResult, error) {
	slideRe := regexp.MustCompile(`slide(\d+)\.xml`)
	slideFiles := make(map[int]*zip.File)
	for _, f := range zr.File {
		if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
			if m := slideRe.FindStringSubmatch(f.Name); len(m) == 2 {
				num, _ := strconv.Atoi(m[1])
				slideFiles[num] = f
			}
		}
	}
	if len(slideFiles) == 0 {
		return nil, fmt.Errorf("no slides found in PPTX")
	}

	var keys []int
	for k := range slideFiles {
		keys = append(keys, k)
	}
	sort.Ints(keys)

	var results []SlideResult
	for _, i := range keys {
		sf := slideFiles[i]
		rc, err := sf.Open()
		if err != nil {
			continue
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			continue
		}

		svgStr := ConvertSlideToSVG(data, i, zr)
		notesMD := ExtractSlideNotes(i, zr)
		results = append(results, SlideResult{
			SlideNum: i,
			SVG:      svgStr,
			NotesMD:  notesMD,
		})
	}
	return results, nil
}

// ConvertPPTXBytes parses a PPTX byte slice and converts all slides to SVG and notes.
func ConvertPPTXBytes(data []byte) ([]SlideResult, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("failed to open PPTX zip: %w", err)
	}
	return ConvertPPTX(zr)
}

// ConvertPPTXFile parses a PPTX file from disk and converts all slides to SVG and notes.
func ConvertPPTXFile(filePath string) ([]SlideResult, error) {
	zr, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open PPTX file: %w", err)
	}
	defer zr.Close()
	return ConvertPPTX(&zr.Reader)
}

func extractSlideNotes(slideNum int, zipReader *zip.Reader) string {
	return ExtractSlideNotes(slideNum, zipReader)
}

func convertSlideToSVG(slideXML []byte, slideNum int, zipReader *zip.Reader) string {
	return ConvertSlideToSVG(slideXML, slideNum, zipReader)
}

func loadFileFromZip(zr *zip.Reader, filePath string) ([]byte, error) {
	return LoadFileFromZip(zr, filePath)
}

// ====================================================================
// PPTX-to-SVG Converter
//
// Parses OOXML slide XML using encoding/xml into typed structs,
// then renders to SVG. No regex-based XML parsing.
// ====================================================================

// EMU to pixel conversion: 914400 EMU = 1 inch at 96 DPI
const emuPerPx = 914400.0 / 96.0

const (
	svgWidth  = 1280
	svgHeight = 720
)

// ======================== Intermediate Render Types ========================

// RenderTextRun represents a single text run with its own formatting.
type RenderTextRun struct {
	Text          string
	FontSize      int    // points; 0 = inherit from shape
	FontFamily    string // empty = inherit from shape
	Bold          bool
	Italic        bool
	Color         string         // hex color; empty = inherit from shape type default
	Underline     string         // "sng","dbl","heavy",etc. or "" for none
	Strikethrough string         // "sngStrike","dblStrike" or "" for none
	BaselineShift int            // positive = superscript, negative = subscript (1/100000 units)
	HyperlinkURL  string         // resolved URL from hlinkClick rId; empty = no link
	GradientDef   *SVGGradient   // text gradient fill (WordArt)
	PatternFill   *RenderPattern // text pattern fill (WordArt)
	StrokeColor   string         // text outline/stroke color
	StrokeWidth   float64        // text outline width in px
	Shadow        *RenderShadow  // text shadow effect
	HasReflection bool           // whether text has a reflection effect
}

// RenderParagraph represents a paragraph with runs and paragraph-level formatting.
type RenderParagraph struct {
	Runs              []RenderTextRun
	Alignment         string // "l", "ctr", "r"
	Level             int
	SpaceBefore       int    // space before in pixels
	SpaceAfter        int    // space after in pixels
	LineSpacing       int    // line spacing in pixels (0 = use config lineHeightFactor)
	LineSpacingPct    int    // spcPct value (e.g. 150000 = 150%; 0 = use LineSpacing or config)
	MarginLeft        int    // left margin in pixels (from marL)
	Indent            int    // first line indent in pixels (from indent)
	IsEmpty           bool   // empty paragraph (contributes spacing only)
	BulletType        string // "none", "char", "autoNum"
	BulletChar        string // for "char" type
	BulletAutoNumType string // for "autoNum" type
	BulletStartAt     int    // start index (default 1)
}

// RenderShape is a shape ready for SVG rendering, built from OOXML types.
type RenderShape struct {
	XmlID          int
	Name           string
	ShapeType      string // "title","body","rect","roundRect","ellipse","freeform","picture","stickyNote","textbox","chart","table","prstGeomFilled"
	PrstGeom       string // preset geometry name (e.g., "downArrow", "upArrow") for prstGeomFilled shapes
	Rotation       int    // rotation in 60000ths of a degree (from xfrm rot attr)
	FlipH          bool   // horizontally flipped
	FlipV          bool   // vertically flipped
	X, Y           int64  // position in EMU
	W, H           int64  // size in EMU
	TextX, TextY   int64  // text box position in EMU (if different from X,Y)
	TextW, TextH   int64  // text box size in EMU (if different from W,H)
	Text           string
	Fill           string
	Stroke         string
	FontSize       int // points
	Bold           bool
	Alignment      string // "l","ctr","r" â€” horizontal paragraph alignment
	AnchorCtr      bool   // Centered anchor bounding box
	TextAnchor     string // "t","ctr","b" â€” vertical text anchor from bodyPr
	RoundRadius    int
	CustPathD      string            // SVG path d attribute for custom geometry
	Paragraphs     []RenderParagraph // structured paragraph data (used for multi-paragraph rendering)
	FontFamily     string            // primary font family (from first non-empty run)
	TextColor      string            // primary text color (hex, from first non-empty run)
	ImageData      string            // base64-encoded image data (for pictures)
	ImageMIME      string            // MIME type for image (e.g., "image/png")
	EmbedRID       string            // relationship ID for embedded content
	PreRenderedSVG string            // pre-rendered SVG fragment (for charts)
	GradientDef    *SVGGradient      // gradient fill definition (nil = solid fill)
	DashStyle      string            // OOXML prstDash val: "dash","dot","lgDash", etc. Empty = solid.
	StrokeWidth    float64           // stroke width in pixels; 0 means use default
	LineCap        string            // SVG stroke-linecap value: "round","square","butt", or empty
	FillOpacity    float64           // fill opacity 0.0â€“1.0; 0 means fully opaque (default)
	SoftEdgeRadius float64           // soft edge blur radius in pixels; 0 means no soft edge
	Shadow         *RenderShadow     // shadow effect (nil if none)
	PatternFill    *RenderPattern    // pattern fill (nil if none)
	TextDirection  string            // text direction: "horz", "vert", "vert270", "eaVert"
	// BodyInsets: text box margins from bodyPr (px). When all 0, renderer uses config defaults.
	BodyLIns, BodyRIns, BodyTIns, BodyBIns int
	// TextAutofit: "none" | "norm" | "shape" from bodyPr noAutofit / normAutofit / spAutoFit
	TextAutofit           string  // Do not Auto Fit, Shrink Text on Overflow, or Resize Shape to Fit Text
	NormFontScalePct      float64 // for "norm": font scale percentage (0 = derive from fit)
	NormLnSpcReductionPct float64 // for "norm": line spacing reduction percentage (0 = 0%)
	// WrapText: from bodyPr wrap. true = wrap on (square or omitted), false = wrap off (none). Case d = wrap off â†’ no wrap, shrink has no effect.
	WrapText     bool
	HyperlinkURL string // resolved URL for shape-level hyperlink (from cNvPr hlinkClick)
}

// RenderPattern holds resolved pattern fill parameters.
type RenderPattern struct {
	Preset  string
	FgColor string // Hex
	BgColor string // Hex
}

// RenderShadow holds resolved shadow parameters for rendering.
type RenderShadow struct {
	Type         string  // "outer", "inner" (perspective is a type of outer)
	BlurRad      int64   // EMUs
	Dist         int64   // EMUs
	Dir          int64   // 60000ths of a degree
	Color        string  // Resolved hex color (e.g. "#000000")
	Alpha        float64 // 0.0-1.0
	Sx, Sy       int64   // Scale factors (100000 = 100%)
	Kx, Ky       int64   // Skew angles (60000ths of a degree)
	Algn         string  // Alignment anchor
	RotWithShape bool
}

// SlideBackground holds the resolved background for a slide.
type SlideBackground struct {
	SolidFill   string       // hex color for solid fill background (empty if gradient)
	GradientDef *SVGGradient // gradient fill for background (nil if solid)
}

// SVGGradient holds the info needed to generate an SVG gradient definition.
type SVGGradient struct {
	ID       string        // unique gradient ID (e.g., "grad_4_5")
	IsRadial bool          // true = radialGradient, false = linearGradient
	Angle    float64       // linear gradient angle in degrees (0 = left-to-right)
	CX, CY   float64       // radial: outer circle center in PIXELS (userSpaceOnUse)
	RX       float64       // radial: gradient radius in PIXELS
	FX, FY   float64       // radial: focal point in PIXELS
	FXSet    bool          // true if fx differs from cx (emit fx attribute)
	FYSet    bool          // true if fy differs from cy (emit fy attribute)
	Stops    []SVGGradStop // gradient color stops
}

// SVGGradStop is a single stop in an SVG gradient.
type SVGGradStop struct {
	Offset float64 // 0.0â€“1.0
	Color  string  // hex color
	Alpha  float64 // 0.0â€“1.0 (1.0 = opaque)
}

// RenderConnector is a connector ready for SVG rendering.
// SX/SY/EX/EY are pre-computed start/end EMU coordinates with flip+rotation applied.
// For curved connectors, CP1X/CP1Y/CP2X/CP2Y hold the bezier control points (also transformed).
type RenderConnector struct {
	SX, SY      int64  // start point in EMU
	EX, EY      int64  // end point in EMU
	CP1X, CP1Y  int64  // bezier control point 1 (EMU); only used for curvedConnector*
	CP2X, CP2Y  int64  // bezier control point 2 (EMU); only used for curvedConnector*
	HasCurveCP  bool   // true if CP1/CP2 are populated (curved connector with spec-based control points)
	GeomType    string // "straightConnector1","bentConnector3","curvedConnector2",etc.
	AdjVal      int    // adjustment value (percentage * 1000); 0 means default 50000
	HasArrow    bool
	ArrowAtHead bool // true when arrow is at head (start) end; false = tail (end)
	Color       string
	DashStyle   string  // OOXML prstDash val: "dash","dot","lgDash","dashDot", etc. Empty = solid.
	StrokeWidth float64 // stroke width in pixels; 0 means use default
	// Gradient fields for compound-path rendering of gradient-stroked connectors
	GradColors    []string // resolved hex colors for each gradient stop, e.g. ["#FFFFFF","#0072C6"]
	GradPositions []int    // gradient stop positions (0â€“100000 scale), e.g. [0, 100000]
	ShapeID       int      // OOXML shape ID, used for unique gradient def naming
	ArrowW        string   // tailEnd/headEnd w attr: "sm","med","lg" (empty = default)
	ArrowLen      string   // tailEnd/headEnd len attr: "sm","med","lg" (empty = default)
}

// placeholderInfo stores layout/master placeholder resolved properties.
type placeholderInfo struct {
	Xfrm   OOXMLXfrm
	Anchor string    // bodyPr anchor from layout ("t", "ctr", "b")
	SpPr   OOXMLSpPr // full properties for fallback (fills, strokes, etc)
}

// paraDefaults holds inherited paragraph-level defaults from layout/master.
type paraDefaults struct {
	SpaceBefore       int    // pixels (from spcBef spcPts, converted ptâ†’px)
	SpaceAfter        int    // pixels (from spcAft spcPts, converted ptâ†’px)
	LineSpacing       int    // pixels (from lnSpc spcPts, converted ptâ†’px; 0=unset)
	LineSpacingPct    int    // spcPct value (from lnSpc spcPct; 0=unset)
	FontSize          int    // px (from defRPr sz, converted ptâ†’px; 0=unset)
	Alignment         string // alignment (l, ctr, r)
	BulletType        string // "none", "char", "autoNum", "" (unset)
	BulletChar        string // glyph for "char" type (e.g. "â€¢")
	BulletAutoNumType string // e.g. "arabicPeriod" for "autoNum" type
	MarginLeft        int    // px (from marL, converted emuâ†’px)
	Indent            int    // px (from indent, converted emuâ†’px)
}

// slideContext carries contextual information needed during slide processing.
type slideContext struct {
	zipReader                  *zip.Reader
	slideRels                  []OOXMLRelationship
	layoutRels                 []OOXMLRelationship
	masterRels                 []OOXMLRelationship
	phMap                      map[string]placeholderInfo  // placeholder type -> info from layout/master
	majorFont                  string                      // theme major font family (for titles)
	minorFont                  string                      // theme minor font family (for body text)
	themeColors                map[string]string           // theme color scheme: "accent1" -> "#2781CB"
	titleDefaultColor          string                      // default title text color (hex) from master txStyles TitleStyle
	titleDefaultFontSizePx     int                         // default title font size (px) from master/layout title placeholder
	titleDefaultLineSpacing    int                         // default title line spacing (px); 0 = use config
	titleDefaultLineSpacingPct int                         // default title line spacing % (e.g. 150000 = 150%); 0 = use titleDefaultLineSpacing or config
	bodyDefaultColor           string                      // default body text color (hex) from master/layout txStyles BodyStyle
	cfg                        *RenderConfig               // rendering configuration
	bodyParaDefsByLevel        map[int]paraDefaults        // inherited body paragraph defaults from layout/master, keyed by paragraph level (0-based)
	tableStyles                map[string]*OOXMLTableStyle // table style GUID -> parsed style definition
	themeFmtScheme             *OOXMLFmtScheme             // parsed theme format scheme (fillStyleLst, effectStyleLst)
}


// ======================== Notes Extraction ========================

// extractSlideNotes reads ppt/notesSlides/notesSlide*.xml that corresponds
// to the given slideNum, parses the body text, and converts it to markdown.
// Returns empty string if no notes exist for the slide.
func ExtractSlideNotes(slideNum int, zipReader *zip.Reader) string {
	// Find which notesSlide file corresponds to this slide via slide rels
	slideRelsPath := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)
	slideRels := loadRelationships(zipReader, slideRelsPath)

	notesTarget := ""
	for _, r := range slideRels {
		if strings.HasSuffix(r.Type, "/notesSlide") {
			notesTarget = r.Target
			break
		}
	}
	if notesTarget == "" {
		return ""
	}

	// Resolve relative path (e.g. "../notesSlides/notesSlide1.xml" -> "ppt/notesSlides/notesSlide1.xml")
	notesPath := path.Join("ppt/slides", notesTarget)
	notesPath = path.Clean(notesPath)

	notesData, err := loadFileFromZip(zipReader, notesPath)
	if err != nil {
		return ""
	}

	var notes OOXMLNotesSlide
	if xml.Unmarshal(notesData, &notes) != nil {
		return ""
	}

	// Load notes slide relationships for hyperlink resolution
	notesBaseName := path.Base(notesPath)
	notesRelsPath := path.Join(path.Dir(notesPath), "_rels", notesBaseName+".rels")
	notesRels := loadRelationships(zipReader, notesRelsPath)

	// Find the body placeholder shape (ph type="body")
	var bodyTxBody *OOXMLNotesTxBody
	for _, sp := range notes.CSld.SpTree.Shapes {
		if sp.NvSpPr.NvPr.Ph != nil && sp.NvSpPr.NvPr.Ph.Type == "body" {
			bodyTxBody = sp.TxBody
			break
		}
	}
	if bodyTxBody == nil {
		return ""
	}

	// Convert paragraphs to markdown
	var paragraphs []string
	for _, para := range bodyTxBody.Paragraphs {
		paraText := convertNotesParaToMarkdown(para, notesRels)
		paragraphs = append(paragraphs, paraText)
	}

	// Join with double newlines, trim trailing empty lines
	result := strings.Join(paragraphs, "\n\n")
	result = strings.TrimSpace(result)
	return result
}

// convertNotesParaToMarkdown converts a single notes paragraph to markdown.
func convertNotesParaToMarkdown(para OOXMLNotesParagraph, rels []OOXMLRelationship) string {
	var runs []string
	for _, r := range para.Runs {
		text := r.Text
		if text == "" {
			continue
		}

		// Apply formatting wrappers
		if r.RPr != nil {
			isBold := r.RPr.Bold == "1" || r.RPr.Bold == "true"
			isItalic := r.RPr.Italic == "1" || r.RPr.Italic == "true"
			isUnderline := r.RPr.Underline == "sng"

			if isBold {
				text = "**" + text + "**"
			}
			if isItalic {
				text = "*" + text + "*"
			}
			if isUnderline {
				text = "<u>" + text + "</u>"
			}

			// Hyperlink
			if r.RPr.HlinkClick != nil && r.RPr.HlinkClick.RID != "" {
				url := resolveHyperlinkURL(r.RPr.HlinkClick.RID, rels)
				if url != "" {
					// Strip formatting wrappers for the link text display
					plainText := r.Text
					text = "[" + plainText + "](" + url + ")"
				}
			}
		}

		runs = append(runs, text)
	}

	lineText := strings.Join(runs, "")

	// Add bullet prefix if paragraph has bullet properties
	if para.PPr != nil && para.PPr.BuNone == nil {
		if para.PPr.BuChar != nil {
			lineText = "- " + lineText
		} else if para.PPr.BuAutoNum != nil {
			lineText = "- " + lineText
		}
	}

	return lineText
}

// resolveHyperlinkURL finds the URL for a hyperlink relationship ID in the given rels.
func resolveHyperlinkURL(rID string, rels []OOXMLRelationship) string {
	for _, r := range rels {
		if r.ID == rID && strings.HasSuffix(r.Type, "/hyperlink") {
			return r.Target
		}
	}
	return ""
}

// ======================== Main Pipeline ========================

func ConvertSlideToSVG(slideXML []byte, slideNum int, zipReader *zip.Reader) string {
	var slide OOXMLSlide
	if err := xml.Unmarshal(slideXML, &slide); err != nil {
		return fmt.Sprintf(`<svg xmlns="http://www.w3.org/2000/svg" width="%d" height="%d"><text x="20" y="40" fill="red">Parse error: %s</text></svg>`, svgWidth, svgHeight, escapeXMLText(err.Error()))
	}

	// Load slide relationships
	slideRelsPath := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)
	slideRels := loadRelationships(zipReader, slideRelsPath)

	// Load layout from slide relationships
	var layout *OOXMLSlideLayout
	var layoutRels []OOXMLRelationship
	layoutTarget := findRelByType(slideRels, "slideLayout")
	if layoutTarget != "" {
		layoutFullPath := resolveRelPath("ppt/slides", layoutTarget)
		layoutData, err := loadFileFromZip(zipReader, layoutFullPath)
		if err == nil {
			var sl OOXMLSlideLayout
			if xml.Unmarshal(layoutData, &sl) == nil {
				layout = &sl
			}
			// Load layout relationships (for images, master, etc.)
			layoutDir := path.Dir(layoutFullPath)
			layoutBase := path.Base(layoutFullPath)
			layoutRelsPath := path.Join(layoutDir, "_rels", layoutBase+".rels")
			layoutRels = loadRelationships(zipReader, layoutRelsPath)
		}
	}

	// Build placeholder map from layout
	phMap := make(map[string]placeholderInfo)
	if layout != nil {
		phMap = buildPlaceholderMap(&layout.CSld.SpTree)
	}

	baseCfg := getRenderConfig()
	// Make a shallow copy so we can set per-PPTX defaults without race conditions
	cfgCopy := *baseCfg
	cfg := &cfgCopy

	// Parse presentation.xml for defaultTextStyle (fallback font size for shapes without explicit sz)
	if presData, presErr := loadFileFromZip(zipReader, "ppt/presentation.xml"); presErr == nil {
		var pres OOXMLPresentation
		if xml.Unmarshal(presData, &pres) == nil && pres.DefaultTextStyle != nil {
			if pres.DefaultTextStyle.Lvl1pPr != nil && pres.DefaultTextStyle.Lvl1pPr.DefRPr != nil && pres.DefaultTextStyle.Lvl1pPr.DefRPr.Sz > 0 {
				cfg.DefaultFontSizePx = ptToPx(pres.DefaultTextStyle.Lvl1pPr.DefRPr.Sz / 100)
			}
		}
	}

	// Load and parse ppt/tableStyles.xml for data-driven table style resolution
	tableStyleMap := make(map[string]*OOXMLTableStyle)
	if tsData, tsErr := loadFileFromZip(zipReader, "ppt/tableStyles.xml"); tsErr == nil {
		var tsList OOXMLTableStyleList
		if xml.Unmarshal(tsData, &tsList) == nil {
			for i := range tsList.Styles {
				s := &tsList.Styles[i]
				tableStyleMap[s.StyleId] = s
			}
		}
	}

	ctx := &slideContext{
		zipReader:   zipReader,
		slideRels:   slideRels,
		layoutRels:  layoutRels,
		phMap:       phMap,
		cfg:         cfg,
		tableStyles: tableStyleMap,
	}

	// Load slide master from layout relationships for placeholder position fallback and theme
	var master *OOXMLSlideMaster
	var masterRels []OOXMLRelationship
	if layoutTarget != "" {
		layoutFullPathForMaster := resolveRelPath("ppt/slides", layoutTarget)
		masterTarget := findRelByType(layoutRels, "slideMaster")
		if masterTarget != "" {
			masterFullPath := resolveRelPath(path.Dir(layoutFullPathForMaster), masterTarget)
			masterData, merr := loadFileFromZip(zipReader, masterFullPath)
			if merr == nil {
				var sm OOXMLSlideMaster
				if xml.Unmarshal(masterData, &sm) == nil {
					master = &sm
					// Fill in phMap entries that weren't found in the layout, and merge attributes like Anchor
					masterPhMap := buildPlaceholderMap(&sm.CSld.SpTree)
					for k, v := range masterPhMap {
						if existing, exists := phMap[k]; !exists {
							phMap[k] = v
						} else {
							// Inherit Anchor from master if the layout placeholder doesn't specify one
							if existing.Anchor == "" && v.Anchor != "" {
								existing.Anchor = v.Anchor
								phMap[k] = existing
							}
						}
					}
					ctx.phMap = phMap

					// Extract body paragraph defaults from master txStyles (per-level)
					ctx.bodyParaDefsByLevel = make(map[int]paraDefaults)
					if sm.TxStyles != nil && sm.TxStyles.BodyStyle != nil {
						bs := sm.TxStyles.BodyStyle
						lvlPPrs := []*OOXMLLvlPPr{bs.Lvl1pPr, bs.Lvl2pPr}
						for lvl, lp := range lvlPPrs {
							if lp == nil {
								continue
							}
							d := paraDefaults{}
							d.Alignment = lp.Algn
							if lp.SpcBef != nil && lp.SpcBef.SpcPts != nil {
								d.SpaceBefore = ptToPx(lp.SpcBef.SpcPts.Val / 100)
							}
							if lp.SpcAft != nil && lp.SpcAft.SpcPts != nil {
								d.SpaceAfter = ptToPx(lp.SpcAft.SpcPts.Val / 100)
							}
							if lp.LnSpc != nil {
								if lp.LnSpc.SpcPts != nil {
									d.LineSpacing = ptToPx(lp.LnSpc.SpcPts.Val / 100)
								} else if lp.LnSpc.SpcPct != nil {
									d.LineSpacingPct = lp.LnSpc.SpcPct.Val
								}
							}
							if lp.DefRPr != nil && lp.DefRPr.Sz > 0 {
								d.FontSize = ptToPx(lp.DefRPr.Sz / 100)
							}
							ctx.bodyParaDefsByLevel[lvl] = d
						}
					}
				}
				// Load theme from master rels
				masterDir := path.Dir(masterFullPath)
				masterBase := path.Base(masterFullPath)
				masterRelsPath := path.Join(masterDir, "_rels", masterBase+".rels")
				masterRels = loadRelationships(zipReader, masterRelsPath)
				ctx.masterRels = masterRels
				themeTarget := findRelByType(masterRels, "theme")
				if themeTarget != "" {
					themeFullPath := resolveRelPath(masterDir, themeTarget)
					themeData, terr := loadFileFromZip(zipReader, themeFullPath)
					if terr == nil {
						var theme OOXMLTheme
						if xml.Unmarshal(themeData, &theme) == nil {
							ctx.majorFont = theme.ThemeElements.FontScheme.MajorFont.Latin.Typeface
							ctx.minorFont = theme.ThemeElements.FontScheme.MinorFont.Latin.Typeface
							fmtScheme := theme.ThemeElements.FmtScheme
							ctx.themeFmtScheme = &fmtScheme
							// Parse color scheme from theme (raw theme slot colors)
							cs := theme.ThemeElements.ClrScheme
							raw := map[string]string{
								"dk1":      cs.Dk1.Hex(),
								"lt1":      cs.Lt1.Hex(),
								"dk2":      cs.Dk2.Hex(),
								"lt2":      cs.Lt2.Hex(),
								"accent1":  cs.Accent1.Hex(),
								"accent2":  cs.Accent2.Hex(),
								"accent3":  cs.Accent3.Hex(),
								"accent4":  cs.Accent4.Hex(),
								"accent5":  cs.Accent5.Hex(),
								"accent6":  cs.Accent6.Hex(),
								"hlink":    cs.Hlink.Hex(),
								"folHlink": cs.FolHlink.Hex(),
							}
							ctx.themeColors = applyClrMap(raw, master.ClrMap)
							// Default text colors, title font size, and title line spacing from master txStyles (title/body) for placeholders with no run-level value
							if master != nil && master.TxStyles != nil {
								if ts := master.TxStyles.TitleStyle; ts != nil && ts.Lvl1pPr != nil && ts.Lvl1pPr.DefRPr != nil {
									ctx.titleDefaultColor = resolveSolidFill(ts.Lvl1pPr.DefRPr.SolidFill, ctx.themeColors)
									if ts.Lvl1pPr.DefRPr.Sz > 0 {
										ctx.titleDefaultFontSizePx = ptToPx(ts.Lvl1pPr.DefRPr.Sz / 100)
									}
								}
								if ts := master.TxStyles.TitleStyle; ts != nil && ts.Lvl1pPr != nil && ts.Lvl1pPr.LnSpc != nil {
									if ts.Lvl1pPr.LnSpc.SpcPts != nil {
										ctx.titleDefaultLineSpacing = ptToPx(ts.Lvl1pPr.LnSpc.SpcPts.Val / 100)
									} else if ts.Lvl1pPr.LnSpc.SpcPct != nil {
										ctx.titleDefaultLineSpacingPct = ts.Lvl1pPr.LnSpc.SpcPct.Val
									}
								}
								if bs := master.TxStyles.BodyStyle; bs != nil {
									if bs.Lvl1pPr != nil && bs.Lvl1pPr.DefRPr != nil {
										ctx.bodyDefaultColor = resolveSolidFill(bs.Lvl1pPr.DefRPr.SolidFill, ctx.themeColors)
									}
									ctx.bodyParaDefsByLevel = extractListStyleParaDefs(bs)
								}
							}
						}
					}
				}
			}
		}
	}

	// Overlay layout lstStyle defaults (overrides master): title/body colors, title font size, title line spacing
	if layout != nil {
		if ctx.themeColors != nil {
			if c := getLayoutTitleDefaultTextColor(layout, ctx.themeColors); c != "" {
				ctx.titleDefaultColor = c
			}
			if c := getLayoutBodyDefaultTextColor(layout, ctx.themeColors); c != "" {
				ctx.bodyDefaultColor = c
			}
		}
		if layoutTitleFS := getLayoutTitleDefaultFontSizePx(layout); layoutTitleFS > 0 {
			ctx.titleDefaultFontSizePx = layoutTitleFS
		}
		if layoutTitleLS, layoutTitleLSPct := getLayoutTitleDefaultLineSpacing(layout); layoutTitleLS > 0 || layoutTitleLSPct > 0 {
			ctx.titleDefaultLineSpacing = layoutTitleLS
			ctx.titleDefaultLineSpacingPct = layoutTitleLSPct
		} else if master != nil {
			// Layout does not define title lnSpc â†’ use master (Single, 1.5, Double, Exactly, Multiple).
			if masterLS, masterLSPct := getMasterTitleDefaultLineSpacing(master); masterLS > 0 || masterLSPct > 0 {
				ctx.titleDefaultLineSpacing = masterLS
				ctx.titleDefaultLineSpacingPct = masterLSPct
			} else {
				ctx.titleDefaultLineSpacingPct = 100000 // master also unset â†’ Single (1.0)
			}
		} else {
			ctx.titleDefaultLineSpacingPct = 100000
		}
		if ldMap := extractLayoutBodyParaDefs(&layout.CSld.SpTree); ldMap != nil {
			if ctx.bodyParaDefsByLevel == nil {
				ctx.bodyParaDefsByLevel = make(map[int]paraDefaults)
			}
			for lvl, ld := range ldMap {
				d := ctx.bodyParaDefsByLevel[lvl] // start from master defaults for this level
				if ld.Alignment != "" {
					d.Alignment = ld.Alignment
				}
				if ld.SpaceBefore > 0 {
					d.SpaceBefore = ld.SpaceBefore
				}
				if ld.SpaceAfter > 0 {
					d.SpaceAfter = ld.SpaceAfter
				}
				if ld.LineSpacing > 0 {
					d.LineSpacing = ld.LineSpacing
					d.LineSpacingPct = 0 // explicit pts overrides pct
				} else if ld.LineSpacingPct > 0 {
					d.LineSpacingPct = ld.LineSpacingPct
					d.LineSpacing = 0
				}
				if ld.FontSize > 0 {
					d.FontSize = ld.FontSize
				}
				if ld.BulletType != "" {
					d.BulletType = ld.BulletType
					d.BulletChar = ld.BulletChar
					d.BulletAutoNumType = ld.BulletAutoNumType
				}
				if ld.MarginLeft != 0 {
					d.MarginLeft = ld.MarginLeft
				}
				if ld.Indent != 0 {
					d.Indent = ld.Indent
				}
				ctx.bodyParaDefsByLevel[lvl] = d
			}
		}
	}

	// Collect renderables from slide
	shapes, connectors := collectRenderables(&slide, ctx)

	// Collect master-level shapes, connectors, and pictures (bottom layer behind everything)
	var masterShapes []RenderShape
	masterSeenTexts := make(map[string]bool)
	if master != nil {
		for _, sp := range master.CSld.SpTree.Shapes {
			// Skip placeholders from master (they are template frames, not decorations)
			if sp.NvSpPr.NvPr.Ph != nil {
				continue
			}
			if rs := processShape(sp, 0, 0, masterSeenTexts, ctx); rs != nil {
				masterShapes = append(masterShapes, *rs)
			}
		}
		for _, cxn := range master.CSld.SpTree.Connectors {
			if rc := processConnector(cxn, ctx); rc != nil {
				connectors = append([]RenderConnector{*rc}, connectors...) // prepend so master connectors render first
			}
		}
		for _, pic := range master.CSld.SpTree.Pictures {
			if rs := processPicture(pic, 0, 0, ctx, masterRels); rs != nil {
				masterShapes = append(masterShapes, *rs)
			}
		}
	}

	// Collect layout-level pictures (background images) and shapes (behind slide content)
	var layoutShapes []RenderShape
	layoutSeenTexts := make(map[string]bool)
	if layout != nil {
		for _, sp := range layout.CSld.SpTree.Shapes {
			// Skip placeholders from layout (they are template frames, not decorations)
			if sp.NvSpPr.NvPr.Ph != nil {
				continue
			}
			if rs := processShape(sp, 0, 0, layoutSeenTexts, ctx); rs != nil {
				layoutShapes = append(layoutShapes, *rs)
			}
		}
		for _, pic := range layout.CSld.SpTree.Pictures {
			if rs := processPicture(pic, 0, 0, ctx, layoutRels); rs != nil {
				layoutShapes = append(layoutShapes, *rs)
			}
		}
		// Collect layout-level connectors (e.g., divider lines)
		for _, cxn := range layout.CSld.SpTree.Connectors {
			if rc := processConnector(cxn, ctx); rc != nil {
				connectors = append(connectors, *rc)
			}
		}
	}

	shapes = deduplicateRenderShapes(shapes)

	// Pre-render chart shapes (load chart XML from zip and generate SVG fragments)
	for i := range shapes {
		if shapes[i].ShapeType == "chart" {
			shapes[i].PreRenderedSVG = loadAndRenderChart(shapes[i], ctx)
		}
	}

	// NOTE: Do NOT sort shapes by Y position. OOXML z-order is defined by document
	// order, not position. Sorting by Y destroys the correct layering of overlapping
	// shapes (e.g., a background rectangle being painted on top of an image within
	// a group). The shapes are already in document order from collectRenderables.

	// Master shapes (bottom) â†’ Layout shapes (middle) â†’ Slide shapes (top)
	allShapes := append(masterShapes, layoutShapes...)
	allShapes = append(allShapes, shapes...)

	// Resolve slide background (cascade: slide â†’ layout â†’ master)
	bg := resolveSlideBackground(&slide, layout, ctx.themeColors)

	return buildSVG(allShapes, connectors, cfg, bg)
}

// ======================== ZIP and Relationship Helpers ========================

// loadFileFromZip reads a file from the ZIP archive by its path.
func LoadFileFromZip(zr *zip.Reader, filePath string) ([]byte, error) {
	// Normalize path separators
	filePath = strings.ReplaceAll(filePath, "\\", "/")
	for _, f := range zr.File {
		if f.Name == filePath {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("file not found in ZIP: %s", filePath)
}

// loadRelationships parses a .rels XML file from the ZIP.
func loadRelationships(zr *zip.Reader, relsPath string) []OOXMLRelationship {
	data, err := loadFileFromZip(zr, relsPath)
	if err != nil {
		return nil
	}
	var rels OOXMLRelationships
	if xml.Unmarshal(data, &rels) != nil {
		return nil
	}
	return rels.Rels
}

// findRelByType finds the first relationship whose Type ends with the given suffix.
func findRelByType(rels []OOXMLRelationship, typeSuffix string) string {
	for _, r := range rels {
		if strings.HasSuffix(r.Type, "/"+typeSuffix) || strings.HasSuffix(r.Type, typeSuffix) {
			return r.Target
		}
	}
	return ""
}

// findRelByID finds a relationship by its ID and returns the target path.
func findRelByID(rels []OOXMLRelationship, id string) string {
	for _, r := range rels {
		if r.ID == id {
			return r.Target
		}
	}
	return ""
}

// resolveRelPath resolves a relative relationship target path against a base directory.
func resolveRelPath(baseDir, relTarget string) string {
	return path.Clean(path.Join(baseDir, relTarget))
}

// buildPlaceholderMap creates a map from placeholder type to xfrm from a shape tree.
// It also stores an alias key "body_N" when a placeholder has both a named type
// (e.g. "body") and an idx, so that slide shapes referencing the placeholder by
// idx alone (no type attribute) can still resolve the position.
func buildPlaceholderMap(tree *OOXMLSpTree) map[string]placeholderInfo {
	m := make(map[string]placeholderInfo)
	for _, sp := range tree.Shapes {
		if ph := sp.NvSpPr.NvPr.Ph; ph != nil {
			key := ph.Type
			if key == "" && ph.Idx != "" {
				key = "body_" + ph.Idx
			}
			// Only store if the shape has position info
			if sp.SpPr.Xfrm.Ext.Cx > 0 || sp.SpPr.Xfrm.Ext.Cy > 0 || sp.SpPr.SolidFill != nil || sp.SpPr.GradFill != nil || sp.SpPr.PattFill != nil || sp.SpPr.BlipFill != nil {
				info := placeholderInfo{
					Xfrm:   sp.SpPr.Xfrm,
					Anchor: sp.TxBody.BodyPr.Anchor,
					SpPr:   sp.SpPr,
				}
				m[key] = info
				// Also store an alias "body_N" when the placeholder has a named
				// type AND an idx. Slide content placeholders often reference by
				// idx only (e.g. <p:ph idx="1"/>), while the master/layout uses
				// the type attribute (e.g. <p:ph type="body" idx="1"/>).
				if ph.Type != "" && ph.Idx != "" {
					aliasKey := "body_" + ph.Idx
					if _, exists := m[aliasKey]; !exists {
						m[aliasKey] = info
					}
				}
			}
		}
	}
	return m
}

// getLayoutTitleDefaultFontSizePx returns the default title font size (px) from the layout's
// title/ctrTitle placeholder lstStyle lvl1pPr.defRPr.sz. 0 if not set (layout does not override master).
func getLayoutTitleDefaultFontSizePx(layout *OOXMLSlideLayout) int {
	if layout == nil {
		return 0
	}
	for _, sp := range layout.CSld.SpTree.Shapes {
		ph := sp.NvSpPr.NvPr.Ph
		if ph == nil {
			continue
		}
		if ph.Type != "title" && ph.Type != "ctrTitle" {
			continue
		}
		ls := sp.TxBody.LstStyle
		if ls == nil || ls.Lvl1pPr == nil || ls.Lvl1pPr.DefRPr == nil || ls.Lvl1pPr.DefRPr.Sz <= 0 {
			continue
		}
		return ptToPx(ls.Lvl1pPr.DefRPr.Sz / 100)
	}
	return 0
}

// getLayoutTitleDefaultLineSpacing returns (lineSpacingPx, lineSpacingPct) from the layout's
// title/ctrTitle placeholder lstStyle lvl1pPr.lnSpc. (0,0) if not set.
func getLayoutTitleDefaultLineSpacing(layout *OOXMLSlideLayout) (int, int) {
	if layout == nil {
		return 0, 0
	}
	for _, sp := range layout.CSld.SpTree.Shapes {
		ph := sp.NvSpPr.NvPr.Ph
		if ph == nil {
			continue
		}
		if ph.Type != "title" && ph.Type != "ctrTitle" {
			continue
		}
		ls := sp.TxBody.LstStyle
		if ls == nil || ls.Lvl1pPr == nil || ls.Lvl1pPr.LnSpc == nil {
			continue
		}
		if ls.Lvl1pPr.LnSpc.SpcPts != nil && ls.Lvl1pPr.LnSpc.SpcPts.Val > 0 {
			return ptToPx(ls.Lvl1pPr.LnSpc.SpcPts.Val / 100), 0
		}
		if ls.Lvl1pPr.LnSpc.SpcPct != nil && ls.Lvl1pPr.LnSpc.SpcPct.Val > 0 {
			return 0, ls.Lvl1pPr.LnSpc.SpcPct.Val
		}
		break
	}
	return 0, 0
}

// getMasterTitleDefaultLineSpacing returns (lineSpacingPx, lineSpacingPct) from the master's
// p:txStyles/p:titleStyle/a:lvl1pPr/a:lnSpc (Single=100%, 1.5 lines=150%, etc.). (0,0) if not set.
func getMasterTitleDefaultLineSpacing(master *OOXMLSlideMaster) (int, int) {
	if master == nil || master.TxStyles == nil || master.TxStyles.TitleStyle == nil || master.TxStyles.TitleStyle.Lvl1pPr == nil {
		return 0, 0
	}
	ln := master.TxStyles.TitleStyle.Lvl1pPr.LnSpc
	if ln == nil {
		return 0, 0
	}
	if ln.SpcPts != nil && ln.SpcPts.Val > 0 {
		return ptToPx(ln.SpcPts.Val / 100), 0
	}
	if ln.SpcPct != nil && ln.SpcPct.Val > 0 {
		return 0, ln.SpcPct.Val
	}
	return 0, 0
}

// getLayoutTitleDefaultTextColor returns the default title text color (hex) from the layout's
// title/ctrTitle placeholder lstStyle lvl1pPr.defRPr.solidFill (e.g. schemeClr tx1 + lumMod/lumOff â†’ #404040). Empty if not set.
func getLayoutTitleDefaultTextColor(layout *OOXMLSlideLayout, themeColors map[string]string) string {
	if layout == nil || themeColors == nil {
		return ""
	}
	for _, sp := range layout.CSld.SpTree.Shapes {
		ph := sp.NvSpPr.NvPr.Ph
		if ph == nil {
			continue
		}
		if ph.Type != "title" && ph.Type != "ctrTitle" {
			continue
		}
		ls := sp.TxBody.LstStyle
		if ls == nil || ls.Lvl1pPr == nil || ls.Lvl1pPr.DefRPr == nil || ls.Lvl1pPr.DefRPr.SolidFill == nil {
			continue
		}
		return resolveSolidFill(ls.Lvl1pPr.DefRPr.SolidFill, themeColors)
	}
	return ""
}

// getLayoutBodyDefaultTextColor returns the default body text color (hex) from the layout's
// body/subTitle placeholder lstStyle defRPr.solidFill, resolved with theme. Empty if not set.
func getLayoutBodyDefaultTextColor(layout *OOXMLSlideLayout, themeColors map[string]string) string {
	if layout == nil || themeColors == nil {
		return ""
	}
	for _, sp := range layout.CSld.SpTree.Shapes {
		ph := sp.NvSpPr.NvPr.Ph
		if ph == nil {
			continue
		}
		if ph.Type != "body" && ph.Type != "subTitle" && !(ph.Type == "" && ph.Idx != "") {
			continue
		}
		ls := sp.TxBody.LstStyle
		if ls == nil || ls.Lvl1pPr == nil || ls.Lvl1pPr.DefRPr == nil {
			continue
		}
		return resolveSolidFill(ls.Lvl1pPr.DefRPr.SolidFill, themeColors)
	}
	return ""
}

func extractListStyleParaDefs(ls *OOXMLTextListStyle) map[int]paraDefaults {
	if ls == nil {
		return nil
	}
	result := make(map[int]paraDefaults)
	lvlPPrs := []*OOXMLLvlPPr{ls.Lvl1pPr, ls.Lvl2pPr}
	for lvl, lp := range lvlPPrs {
		if lp == nil {
			continue
		}
		d := paraDefaults{}
		d.Alignment = lp.Algn
		if lp.MarL != 0 {
			d.MarginLeft = emuToPx(lp.MarL)
		}
		if lp.Indent != 0 {
			d.Indent = emuToPx(lp.Indent)
		}
		if lp.SpcBef != nil && lp.SpcBef.SpcPts != nil {
			d.SpaceBefore = ptToPx(lp.SpcBef.SpcPts.Val / 100)
		}
		if lp.SpcAft != nil && lp.SpcAft.SpcPts != nil {
			d.SpaceAfter = ptToPx(lp.SpcAft.SpcPts.Val / 100)
		}
		if lp.LnSpc != nil {
			if lp.LnSpc.SpcPts != nil {
				d.LineSpacing = ptToPx(lp.LnSpc.SpcPts.Val / 100)
			} else if lp.LnSpc.SpcPct != nil {
				d.LineSpacingPct = lp.LnSpc.SpcPct.Val
			}
		}
		if lp.DefRPr != nil && lp.DefRPr.Sz > 0 {
			d.FontSize = ptToPx(lp.DefRPr.Sz / 100)
		}
		if lp.BuNone != nil {
			d.BulletType = "none"
		} else if lp.BuChar != nil {
			d.BulletType = "char"
			d.BulletChar = lp.BuChar.Char
		} else if lp.BuAutoNum != nil {
			d.BulletType = "autoNum"
			d.BulletAutoNumType = lp.BuAutoNum.Type
		}
		result[lvl] = d
	}
	if len(result) > 0 {
		return result
	}
	return nil
}

// extractLayoutBodyParaDefs extracts paragraph defaults from the layout's body
// placeholder lstStyle (overriding master txStyles defaults).
func extractLayoutBodyParaDefs(tree *OOXMLSpTree) map[int]paraDefaults {
	for _, sp := range tree.Shapes {
		if ph := sp.NvSpPr.NvPr.Ph; ph != nil {
			// Match body placeholder (idx="1" or type="body")
			isBody := ph.Type == "body" || ph.Type == "subTitle" || (ph.Type == "" && ph.Idx != "")
			if !isBody {
				continue
			}
			return extractListStyleParaDefs(sp.TxBody.LstStyle)
		}
	}
	return nil
}

// extractImageBase64 reads an image from the ZIP and returns base64-encoded data and MIME type.
func extractImageBase64(zr *zip.Reader, imagePath string) (string, string) {
	data, err := loadFileFromZip(zr, imagePath)
	if err != nil {
		return "", ""
	}
	encoded := base64.StdEncoding.EncodeToString(data)
	mime := guessMIME(imagePath)
	return encoded, mime
}

// guessMIME guesses the MIME type from a file path extension.
func guessMIME(filePath string) string {
	lower := strings.ToLower(filePath)
	switch {
	case strings.HasSuffix(lower, ".png"):
		return "image/png"
	case strings.HasSuffix(lower, ".jpg") || strings.HasSuffix(lower, ".jpeg"):
		return "image/jpeg"
	case strings.HasSuffix(lower, ".gif"):
		return "image/gif"
	case strings.HasSuffix(lower, ".svg"):
		return "image/svg+xml"
	case strings.HasSuffix(lower, ".emf"):
		return "image/x-emf"
	case strings.HasSuffix(lower, ".wmf"):
		return "image/x-wmf"
	default:
		return "application/octet-stream"
	}
}

// getSvgBlipRID extracts the SVG blip relationship ID from a blip's ExtLst.
// Returns empty string if no SVG blip is found.
func getSvgBlipRID(blip OOXMLBlip) string {
	if blip.ExtLst == nil {
		return ""
	}
	for _, ext := range blip.ExtLst.Exts {
		if ext.SvgBlip != nil && ext.SvgBlip.Embed != "" {
			return ext.SvgBlip.Embed
		}
	}
	return ""
}

// svgPathRe matches <path ... /> elements in SVG content.
var svgPathRe = regexp.MustCompile(`<path\b([^>]*)/>`)

// svgAttrRe extracts a named attribute from an SVG element tag.
var svgAttrRe = regexp.MustCompile(`(\w+)="([^"]*)"`)

// extractSVGPathsFromZip reads an SVG file from the ZIP and returns inline SVG
// fragment with <path> elements wrapped in a <g> with a matrix transform to
// position and scale them onto the slide.
// x, y, w, h are the picture's EMU position/size.
func extractSVGPathsFromZip(zr *zip.Reader, svgPath string, xEmu, yEmu, wEmu, hEmu int64) string {
	data, err := loadFileFromZip(zr, svgPath)
	if err != nil || len(data) == 0 {
		return ""
	}
	svgContent := string(data)

	// Parse viewBox to get the source coordinate system
	vbRe := regexp.MustCompile(`viewBox="([^"]*)"`)
	vbMatch := vbRe.FindStringSubmatch(svgContent)
	if vbMatch == nil {
		return ""
	}
	vbParts := strings.Fields(vbMatch[1])
	if len(vbParts) < 4 {
		return ""
	}
	vbW, _ := strconv.ParseFloat(vbParts[2], 64)
	vbH, _ := strconv.ParseFloat(vbParts[3], 64)
	if vbW == 0 || vbH == 0 {
		return ""
	}

	// Convert picture EMU position/size to pixels
	px := float64(xEmu) / 914400.0 * 96.0
	py := float64(yEmu) / 914400.0 * 96.0
	pw := float64(wEmu) / 914400.0 * 96.0
	ph := float64(hEmu) / 914400.0 * 96.0

	// Calculate scale factors
	scaleX := pw / vbW
	scaleY := ph / vbH

	// Extract all <path .../> elements
	pathMatches := svgPathRe.FindAllStringSubmatch(svgContent, -1)
	if len(pathMatches) == 0 {
		return ""
	}

	var out strings.Builder
	// Wrap in nested <g> elements to match PowerPoint's output structure
	out.WriteString(fmt.Sprintf("  <g><g><g>\n"))
	for _, pm := range pathMatches {
		attrs := pm[1]
		// Extract d and fill attributes
		attrMap := make(map[string]string)
		for _, am := range svgAttrRe.FindAllStringSubmatch(attrs, -1) {
			attrMap[am[1]] = am[2]
		}
		d := attrMap["d"]
		if d == "" {
			continue
		}
		fill := attrMap["fill"]
		if fill == "" {
			fill = "#FFFFFF"
		}
		// Build the path with matrix transform
		out.WriteString(fmt.Sprintf(`  <path d="%s" fill="%s" transform="matrix(%.6g 0 0 %.6g %.6g %.6g)"/>`, d, fill, scaleX, scaleY, px, py))
		out.WriteString("\n")
	}
	out.WriteString("  </g></g></g>\n")

	return out.String()
}

// ======================== Collecting Renderables ========================

func collectRenderables(slide *OOXMLSlide, ctx *slideContext) ([]RenderShape, []RenderConnector) {
	return convertSpTree(&slide.CSld.SpTree, ctx)
}

func convertSpTree(tree *OOXMLSpTree, ctx *slideContext) ([]RenderShape, []RenderConnector) {
	var shapes []RenderShape
	var conns []RenderConnector

	// Process in document order (ECMA-376 19.3.1.45: z-order = lexical order).
	if len(tree.ChildOrder) > 0 {
		for _, ref := range tree.ChildOrder {
			switch ref.Kind {
			case "sp":
				if ref.Index >= len(tree.Shapes) {
					continue
				}
				shapeSeenTexts := make(map[string]bool)
				if rs := processShape(tree.Shapes[ref.Index], 0, 0, shapeSeenTexts, ctx); rs != nil {
					shapes = append(shapes, *rs)
				}
			case "grpSp":
				if ref.Index >= len(tree.Groups) {
					continue
				}
				groupSeenTexts := make(map[string]bool)
				gs, gc := processGroup(tree.Groups[ref.Index], groupSeenTexts, ctx)
				shapes = append(shapes, gs...)
				conns = append(conns, gc...)
			case "pic":
				if ref.Index >= len(tree.Pictures) {
					continue
				}
				if rs := processPicture(tree.Pictures[ref.Index], 0, 0, ctx, ctx.slideRels); rs != nil {
					shapes = append(shapes, *rs)
				}
			case "cxnSp":
				if ref.Index >= len(tree.Connectors) {
					continue
				}
				if rc := processConnector(tree.Connectors[ref.Index], ctx); rc != nil {
					conns = append(conns, *rc)
				}
			case "graphicFrame":
				if ref.Index >= len(tree.GraphicFrames) {
					continue
				}
				ss, sc := processGraphicFrame(tree.GraphicFrames[ref.Index], ctx)
				shapes = append(shapes, ss...)
				conns = append(conns, sc...)
			}
		}
		return shapes, conns
	}

	// Fallback when ChildOrder not set (e.g. legacy or non-custom parse)
	for _, sp := range tree.Shapes {
		shapeSeenTexts := make(map[string]bool)
		if rs := processShape(sp, 0, 0, shapeSeenTexts, ctx); rs != nil {
			shapes = append(shapes, *rs)
		}
	}
	for _, grp := range tree.Groups {
		groupSeenTexts := make(map[string]bool)
		gs, gc := processGroup(grp, groupSeenTexts, ctx)
		shapes = append(shapes, gs...)
		conns = append(conns, gc...)
	}
	for _, pic := range tree.Pictures {
		if rs := processPicture(pic, 0, 0, ctx, ctx.slideRels); rs != nil {
			shapes = append(shapes, *rs)
		}
	}
	for _, cxn := range tree.Connectors {
		if rc := processConnector(cxn, ctx); rc != nil {
			conns = append(conns, *rc)
		}
	}
	for _, gf := range tree.GraphicFrames {
		ss, sc := processGraphicFrame(gf, ctx)
		shapes = append(shapes, ss...)
		conns = append(conns, sc...)
	}
	return shapes, conns
}

// ======================== Shape Processing ========================

func extractShapeShadow(effectLst *OOXMLEffectLst, themeColors map[string]string) *RenderShadow {
	if effectLst == nil {
		return nil
	}
	// Prioritize outer shadow (which covers perspective shadows too)
	if outer := effectLst.OuterShdw; outer != nil {
		blurRad, _ := strconv.ParseFloat(outer.BlurRad, 64)
		dist, _ := strconv.ParseFloat(outer.Dist, 64)
		dir, _ := strconv.ParseFloat(outer.Dir, 64)

		s := &RenderShadow{
			Type:         "outer",
			BlurRad:      int64(blurRad),
			Dist:         int64(dist),
			Dir:          int64(dir),
			Sx:           outer.Sx,
			Sy:           outer.Sy,
			Kx:           outer.Kx,
			Ky:           outer.Ky,
			Algn:         outer.Algn,
			RotWithShape: outer.RotWithShape == 1,
			Alpha:        1.0, // Default opaque
		}
		// Resolve color
		s.Color = resolveShadowColorHex(outer.SrgbClr, outer.SchemeClr, outer.PrstClr, themeColors)

		// Resolve alpha
		if outer.SrgbClr != nil && outer.SrgbClr.Alpha != nil {
			s.Alpha = float64(outer.SrgbClr.Alpha.Val) / 100000.0
		} else if outer.SchemeClr != nil && outer.SchemeClr.Alpha != nil {
			s.Alpha = float64(outer.SchemeClr.Alpha.Val) / 100000.0
		} else if outer.PrstClr != nil && outer.PrstClr.Alpha != nil {
			s.Alpha = float64(outer.PrstClr.Alpha.Val) / 100000.0
		}
		return s
	}
	// Check inner shadow
	if inner := effectLst.InnerShdw; inner != nil {
		blurRad, _ := strconv.ParseFloat(inner.BlurRad, 64)
		dist, _ := strconv.ParseFloat(inner.Dist, 64)
		dir, _ := strconv.ParseFloat(inner.Dir, 64)
		s := &RenderShadow{
			Type:    "inner",
			BlurRad: int64(blurRad),
			Dist:    int64(dist),
			Dir:     int64(dir),
			Alpha:   1.0,
		}
		s.Color = resolveShadowColorHex(inner.SrgbClr, inner.SchemeClr, inner.PrstClr, themeColors)
		if inner.SrgbClr != nil && inner.SrgbClr.Alpha != nil {
			s.Alpha = float64(inner.SrgbClr.Alpha.Val) / 100000.0
		} else if inner.SchemeClr != nil && inner.SchemeClr.Alpha != nil {
			s.Alpha = float64(inner.SchemeClr.Alpha.Val) / 100000.0
		} else if inner.PrstClr != nil && inner.PrstClr.Alpha != nil {
			s.Alpha = float64(inner.PrstClr.Alpha.Val) / 100000.0
		}
		return s
	}
	return nil
}

func resolveColorWrapper(cw *OOXMLColorChoice, themeColors map[string]string) string {
	if cw == nil {
		return ""
	}
	return resolveShadowColorHex(cw.SrgbClr, cw.SchemeClr, cw.PrstClr, themeColors)
}

func extractPatternFill(pf *OOXMLPattFill, themeColors map[string]string) *RenderPattern {
	if pf == nil {
		return nil
	}
	return &RenderPattern{
		Preset:  pf.Prst,
		FgColor: resolveColorWrapper(pf.FgClr, themeColors),
		BgColor: resolveColorWrapper(pf.BgClr, themeColors),
	}
}

func processShape(sp OOXMLShape, offX, offY int64, seenTexts map[string]bool, ctx *slideContext) *RenderShape {
	cnvPr := sp.NvSpPr.CNvPr
	xfrm := sp.SpPr.Xfrm
	x := xfrm.Off.X + offX
	y := xfrm.Off.Y + offY
	w := xfrm.Ext.Cx
	h := xfrm.Ext.Cy

	// Resolve text box transform if present (SmartArt often separates geometry and text bounds)
	txX, txY, txW, txH := x, y, w, h
	if sp.TxXfrm != nil {
		txX = sp.TxXfrm.Off.X + offX
		txY = sp.TxXfrm.Off.Y + offY
		txW = sp.TxXfrm.Ext.Cx
		txH = sp.TxXfrm.Ext.Cy
	}

	// Resolve placeholder key for layout/master lookup
	var phKey string
	if ph := sp.NvSpPr.NvPr.Ph; ph != nil {
		phKey = ph.Type
		if phKey == "" && ph.Idx != "" {
			phKey = "body_" + ph.Idx
		}
	}

	// If shape has no position/size, resolve from layout placeholder
	if x == 0 && y == 0 && w == 0 && h == 0 && ctx != nil && phKey != "" {
		if phInfo, ok := ctx.phMap[phKey]; ok {
			x = phInfo.Xfrm.Off.X + offX
			y = phInfo.Xfrm.Off.Y + offY
			w = phInfo.Xfrm.Ext.Cx
			h = phInfo.Xfrm.Ext.Cy
		}
	}

	// Always resolve layout bodyPr anchor as fallback (independent of xfrm)
	var layoutAnchor string
	if ctx != nil && phKey != "" {
		if phInfo, ok := ctx.phMap[phKey]; ok {
			layoutAnchor = phInfo.Anchor
		}
	}

	// Determine shape type
	shapeType := classifyShape(sp)

	// Resolve theme colors from context for use throughout shape processing
	var themeColors map[string]string
	if ctx != nil {
		themeColors = ctx.themeColors
	}

	// Extract text (returns structured paragraphs + backward-compat fields)
	text, fontSize, bold, alignment, paragraphs, fontFamily, textColor := extractText(sp.TxBody, seenTexts, themeColors, ctx.slideRels)

	// Center-title placeholder: OOXML ctrTitle implies horizontal center alignment when not set in paragraph
	if shapeType == "title" && alignment == "" && sp.NvSpPr.NvPr.Ph != nil && sp.NvSpPr.NvPr.Ph.Type == "ctrTitle" {
		alignment = "ctr"
	}

	// Default text color from PPT XML when run has no solidFill: master/layout txStyles then theme
	if textColor == "" && ctx != nil && ctx.themeColors != nil {
		if shapeType == "title" && ctx.titleDefaultColor != "" {
			textColor = ctx.titleDefaultColor
		} else if shapeType == "body" && ctx.bodyDefaultColor != "" {
			textColor = ctx.bodyDefaultColor
		}
		if textColor == "" {
			if c := ctx.themeColors["tx1"]; c != "" {
				textColor = c
			} else if c := ctx.themeColors["dk1"]; c != "" {
				textColor = c
			}
		}
	}

	// Line spacing resolution order: 1) para pPr (already in paragraph from extractText), 2) shape txBody lstStyle, 3) layout/master by type (body/title below).
	// Apply shape's own lstStyle properties (lnSpc, font size, bullets) when paragraph has no explicit value.
	if sp.TxBody.LstStyle != nil {
		lvlPPrs := []*OOXMLLvlPPr{sp.TxBody.LstStyle.Lvl1pPr, sp.TxBody.LstStyle.Lvl2pPr}
		for i := range paragraphs {
			lvl := paragraphs[i].Level
			if lvl < 0 || lvl >= len(lvlPPrs) || lvlPPrs[lvl] == nil {
				continue
			}
			lp := lvlPPrs[lvl]

			// Line spacing
			if paragraphs[i].LineSpacing == 0 && paragraphs[i].LineSpacingPct == 0 && lp.LnSpc != nil {
				if lp.LnSpc.SpcPts != nil && lp.LnSpc.SpcPts.Val > 0 {
					paragraphs[i].LineSpacing = ptToPx(lp.LnSpc.SpcPts.Val / 100)
				} else if lp.LnSpc.SpcPct != nil && lp.LnSpc.SpcPct.Val > 0 {
					paragraphs[i].LineSpacingPct = lp.LnSpc.SpcPct.Val
				}
			}

			// Font size
			if lp.DefRPr != nil && lp.DefRPr.Sz > 0 {
				dFontSize := ptToPx(lp.DefRPr.Sz / 100)
				for j := range paragraphs[i].Runs {
					if paragraphs[i].Runs[j].FontSize == 0 {
						paragraphs[i].Runs[j].FontSize = dFontSize
					}
				}
				if fontSize == 0 {
					fontSize = dFontSize // fallback for shape level
				}
			}

			// Bullets
			if paragraphs[i].BulletType == "none" { // extractText defaults to "none" if no explicit pPr bullet set
				if lp.BuNone != nil {
					paragraphs[i].BulletType = "none"
				} else if lp.BuChar != nil {
					paragraphs[i].BulletType = "char"
					paragraphs[i].BulletChar = lp.BuChar.Char
				} else if lp.BuAutoNum != nil {
					paragraphs[i].BulletType = "autoNum"
					paragraphs[i].BulletAutoNumType = lp.BuAutoNum.Type
				}

				if paragraphs[i].BulletType != "none" {
					if paragraphs[i].MarginLeft == 0 && lp.MarL != 0 {
						paragraphs[i].MarginLeft = emuToPx(lp.MarL)
					}
					if paragraphs[i].Indent == 0 && lp.Indent != 0 {
						paragraphs[i].Indent = emuToPx(lp.Indent)
					}
				}
			}

			// Always inherit margins if unset by paragraph and paragraph is still a bullet
			if paragraphs[i].MarginLeft == 0 && lp.MarL != 0 {
				paragraphs[i].MarginLeft = emuToPx(lp.MarL)
			}
			if paragraphs[i].Indent == 0 && lp.Indent != 0 {
				paragraphs[i].Indent = emuToPx(lp.Indent)
			}
		}
	}

	// Apply inherited body paragraph defaults from layout/master for body-type shapes (per-level)
	if (shapeType == "body") && ctx != nil && ctx.bodyParaDefsByLevel != nil {
		for i := range paragraphs {
			lvl := paragraphs[i].Level
			d, ok := ctx.bodyParaDefsByLevel[lvl]
			if !ok {
				// Fall back to level 0 defaults if the specific level is not found
				d = ctx.bodyParaDefsByLevel[0]
			}
			// SpaceBefore: use master/layout default if paragraph doesn't specify
			if paragraphs[i].SpaceBefore == 0 && d.SpaceBefore > 0 {
				paragraphs[i].SpaceBefore = d.SpaceBefore
			}
			// SpaceAfter: use layout default if paragraph doesn't specify
			if paragraphs[i].SpaceAfter == 0 && d.SpaceAfter > 0 {
				paragraphs[i].SpaceAfter = d.SpaceAfter
			}
			// Alignment: use layout default if paragraph doesn't specify
			if paragraphs[i].Alignment == "" && d.Alignment != "" {
				paragraphs[i].Alignment = d.Alignment
				if i == 0 && alignment == "" {
					alignment = d.Alignment
				}
			}
			// LineSpacing: use inherited default if paragraph doesn't specify
			if paragraphs[i].LineSpacing == 0 && paragraphs[i].LineSpacingPct == 0 {
				if d.LineSpacing > 0 {
					paragraphs[i].LineSpacing = d.LineSpacing
				} else if d.LineSpacingPct > 0 {
					paragraphs[i].LineSpacingPct = d.LineSpacingPct
				}
			}
			// FontSize: apply layout/master default to runs that have no explicit size
			if d.FontSize > 0 {
				for j := range paragraphs[i].Runs {
					if paragraphs[i].Runs[j].FontSize == 0 {
						paragraphs[i].Runs[j].FontSize = d.FontSize
					}
				}
			}
			// Bullets: apply inherited bullet settings if none was explicitly defined
			if paragraphs[i].BulletType == "none" && d.BulletType != "" && d.BulletType != "none" {
				paragraphs[i].BulletType = d.BulletType
				paragraphs[i].BulletChar = d.BulletChar
				paragraphs[i].BulletAutoNumType = d.BulletAutoNumType
			}
			// Always inherit formatting margins for bullets if unset at paragraph run
			if paragraphs[i].MarginLeft == 0 && d.MarginLeft != 0 {
				paragraphs[i].MarginLeft = d.MarginLeft
			}
			if paragraphs[i].Indent == 0 && d.Indent != 0 {
				paragraphs[i].Indent = d.Indent
			}
		}
		// Shape-level fontSize fallback from body level-0 default when still unset
		if fontSize == 0 && ctx.bodyParaDefsByLevel[0].FontSize > 0 {
			fontSize = ctx.bodyParaDefsByLevel[0].FontSize
		}
	}

	// Title line spacing: use layout default when set; when layout has no lnSpc, use 1.0 (100%).
	// Apply only when paragraph has no explicit lnSpc so explicit slide values (e.g. 1.5) are preserved.
	if shapeType == "title" && ctx != nil {
		for i := range paragraphs {
			if paragraphs[i].LineSpacing != 0 || paragraphs[i].LineSpacingPct != 0 {
				continue // paragraph has explicit lnSpc (e.g. user selected 1.5) â†’ keep it
			}
			if ctx.titleDefaultLineSpacing > 0 {
				paragraphs[i].LineSpacing = ctx.titleDefaultLineSpacing
			} else if ctx.titleDefaultLineSpacingPct > 0 {
				paragraphs[i].LineSpacingPct = ctx.titleDefaultLineSpacingPct
			}
		}
	}

	// Shape-level fontSize from title default when still unset (title placeholders with no run/para sz)
	if fontSize == 0 && shapeType == "title" && ctx != nil && ctx.titleDefaultFontSizePx > 0 {
		fontSize = ctx.titleDefaultFontSizePx
	}

	// Resolve fill and stroke
	fill, stroke := resolveShapeFill(sp, shapeType, themeColors)

	// Determine if we should use a BlipFill (Slide shape > Layout placeholder > empty)
	var blipFill *OOXMLBlipFill
	if sp.SpPr.BlipFill != nil {
		blipFill = sp.SpPr.BlipFill
	} else if fill == "" && sp.SpPr.SolidFill == nil && sp.SpPr.GradFill == nil && sp.SpPr.PattFill == nil && sp.SpPr.NoFill == nil && ctx != nil && phKey != "" {
		// Inherit blipFill from placeholder only if shape has no explicit fill overrides
		if phInfo, ok := ctx.phMap[phKey]; ok && phInfo.SpPr.BlipFill != nil {
			blipFill = phInfo.SpPr.BlipFill
		}
	}
	hasBlipFill := blipFill != nil

	var blipData, blipMIME string
	if hasBlipFill && ctx != nil && ctx.zipReader != nil && blipFill.Blip.Embed != "" {
		embedRID := blipFill.Blip.Embed
		var imageTarget, baseDir string
		if target := findRelByID(ctx.slideRels, embedRID); target != "" {
			imageTarget = target
			baseDir = "ppt/slides"
		} else if target := findRelByID(ctx.layoutRels, embedRID); target != "" {
			imageTarget = target
			baseDir = "ppt/slideLayouts"
		} else if target := findRelByID(ctx.masterRels, embedRID); target != "" {
			imageTarget = target
			baseDir = "ppt/slideMasters"
		}
		if imageTarget != "" {
			imagePath := resolveRelPath(baseDir, imageTarget)
			blipData, blipMIME = extractImageBase64(ctx.zipReader, imagePath)
		}
	}

	// Skip shapes without text and without visual fill or stroke (solid, gradient, pattern, or blip).
	// Never skip textbox: they often contain placeholder/template text (e.g. "Placeholder" in Green Accent 2)
	// and may be deduped or empty in edge cases; we still want to render them.
	hasGradFill := sp.SpPr.GradFill != nil && len(sp.SpPr.GradFill.GsLst.Gs) > 0
	hasPattFill := sp.SpPr.PattFill != nil
	hasVisibleStroke := stroke != "" && stroke != "none"
	if text == "" && fill == "" && !hasGradFill && !hasPattFill && !hasBlipFill && !hasVisibleStroke && shapeType != "freeform" && shapeType != "picture" && shapeType != "prstGeomFilled" && shapeType != "textbox" {
		return nil
	}

	// Custom geometry SVG path (absolute pixel coordinates, no scale transform needed)
	var custPathD string
	if shapeType == "freeform" && sp.SpPr.CustGeom != nil {
		custPathD = custGeomToSVGPath(sp.SpPr.CustGeom, x, y, w, h)
	}

	// Preset geometry path for arrow shapes etc.
	var prstGeom string
	var preRenderedSVG string
	if shapeType == "prstGeomFilled" && sp.SpPr.PrstGeom != nil {
		prstGeom = sp.SpPr.PrstGeom.Prst
		adjs := extractPrstAdjs(sp.SpPr.PrstGeom)
		// Compound shapes need multiple SVG elements (e.g. smiley face = circle + eyes + mouth)
		if isCompoundPreset(prstGeom) {
			preRenderedSVG = renderCompoundPreset(prstGeom, x, y, w, h, fill, stroke, adjs, ctx)
		} else {
			custPathD = prstGeomToSVGPath(prstGeom, x, y, w, h, adjs)
		}
	}

	roundRadius := 0
	if shapeType == "roundRect" {
		roundRadius = 12
	}

	// Determine vertical text anchor from bodyPr (Format Shape â†’ Vertical alignment), then layout fallback
	textAnchor := sp.TxBody.BodyPr.Anchor
	if textAnchor == "" && layoutAnchor != "" {
		textAnchor = layoutAnchor // inherit from layout placeholder
	}
	// OOXML default vertical alignment is top ("t"). Use center for ellipse only when XML explicitly has ctr/middle.
	anchorFromXML := textAnchor
	if textAnchor == "" {
		textAnchor = "t"
	} else if shapeType == "ellipse" && textAnchor != "t" && textAnchor != "b" {
		// e.g. "ctr", "middle" â†’ keep center for ellipse
		textAnchor = "ctr"
	}

	// Use theme fonts when shape doesn't specify one explicitly
	if ctx != nil {
		if shapeType == "title" && fontFamily == "" && ctx.majorFont != "" {
			fontFamily = ctx.majorFont
		}
		if (shapeType == "body" || shapeType == "textbox") && fontFamily == "" && ctx.minorFont != "" {
			fontFamily = ctx.minorFont
		}
	}

	// Build gradient definition if shape has a gradient fill
	var gradDef *SVGGradient
	if sp.SpPr.GradFill != nil && len(sp.SpPr.GradFill.GsLst.Gs) > 0 {
		gradDef = buildSVGGradient(sp.SpPr.GradFill, cnvPr.ID, x, y, w, h, themeColors)
	}

	// Extract line dash style, stroke width, and line cap from <a:ln>
	dashStyle := ""
	strokeWidth := 0.0
	lineCap := ""
	if ln := sp.SpPr.Ln; ln != nil {
		if ln.PrstDash != nil && ln.PrstDash.Val != "" && ln.PrstDash.Val != "solid" {
			dashStyle = ln.PrstDash.Val
		}
		if ln.W > 0 {
			strokeWidth = float64(ln.W) / emuPerPx
		}
		// Map OOXML line cap to SVG stroke-linecap
		switch ln.Cap {
		case "rnd":
			lineCap = "round"
		case "sq":
			lineCap = "square"
		case "flat":
			lineCap = "butt"
		}
	}
	// Resolve default stroke width from theme lnRef when not explicit
	if strokeWidth == 0 && sp.Style.LnRef.Idx != "" {
		switch sp.Style.LnRef.Idx {
		case "1":
			strokeWidth = 6350.0 / emuPerPx // theme lnStyleLst[0]: 0.67px
		case "2":
			strokeWidth = 12700.0 / emuPerPx // theme lnStyleLst[1]: 1.33px
		case "3":
			strokeWidth = 19050.0 / emuPerPx // theme lnStyleLst[2]: 2.0px
		}
	}

	// Resolve soft edge effect (Gaussian blur radius)
	softEdgeRadius := 0.0
	if sp.SpPr.EffectLst != nil && sp.SpPr.EffectLst.SoftEdge != nil {
		softEdgeRadius = float64(sp.SpPr.EffectLst.SoftEdge.Rad) / emuPerPx
	}

	// Resolve fill opacity from alpha channel on solidFill color
	fillOpacity := 0.0 // 0 means fully opaque (default)
	if sf := sp.SpPr.SolidFill; sf != nil {
		if sf.SrgbClr != nil && sf.SrgbClr.Alpha != nil {
			fillOpacity = float64(sf.SrgbClr.Alpha.Val) / 100000.0
		} else if sf.SchemeClr != nil && sf.SchemeClr.Alpha != nil {
			fillOpacity = float64(sf.SchemeClr.Alpha.Val) / 100000.0
		}
	}

	// Text box insets from bodyPr (OOXML defaults when not set: lIns/rIns 91440, tIns/bIns 45720 EMU)
	const defaultLIns = 91440
	const defaultRIns = 91440
	const defaultTIns = 45720
	const defaultBIns = 45720
	lInsEmu := defaultLIns
	if sp.TxBody.BodyPr.LIns != nil {
		lInsEmu = int(*sp.TxBody.BodyPr.LIns)
	}
	rInsEmu := defaultRIns
	if sp.TxBody.BodyPr.RIns != nil {
		rInsEmu = int(*sp.TxBody.BodyPr.RIns)
	}
	tInsEmu := defaultTIns
	if sp.TxBody.BodyPr.TIns != nil {
		tInsEmu = int(*sp.TxBody.BodyPr.TIns)
	}
	bInsEmu := defaultBIns
	if sp.TxBody.BodyPr.BIns != nil {
		bInsEmu = int(*sp.TxBody.BodyPr.BIns)
	}
	bodyLIns := int(math.Round(float64(lInsEmu) / emuPerPx))
	bodyRIns := int(math.Round(float64(rInsEmu) / emuPerPx))
	bodyTIns := int(math.Round(float64(tInsEmu) / emuPerPx))
	bodyBIns := int(math.Round(float64(bInsEmu) / emuPerPx))

	// Text box autofit from bodyPr: noAutofit / normAutofit / spAutoFit (OOXML default = none)
	textAutofit := "none"
	var normFontScalePct, normLnSpcReductionPct float64
	if sp.TxBody.BodyPr.NoAutofit != nil {
		textAutofit = "none"
	} else if sp.TxBody.BodyPr.NormAutofit != nil {
		textAutofit = "norm"
		normFontScalePct = parsePercentString(sp.TxBody.BodyPr.NormAutofit.FontScale)
		normLnSpcReductionPct = parsePercentString(sp.TxBody.BodyPr.NormAutofit.LnSpcReduction)
	} else if sp.TxBody.BodyPr.SpAutoFit != nil {
		textAutofit = "shape"
	} else {
		textAutofit = "none"
	}

	// Wrap text: bodyPr wrap="none" â†’ off (case d); square or omitted â†’ on (cases a, b, c). OOXML default is square.
	wrapText := true
	if sp.TxBody.BodyPr.Wrap == "none" {
		wrapText = false
	}

	// #region agent log
	if text != "" {
		writeDebugLog("processShape textAnchor", map[string]interface{}{
			"name": cnvPr.Name, "shapeType": shapeType,
			"anchorFromXML": anchorFromXML, "textAnchorFinal": textAnchor,
			"bodyTIns": bodyTIns, "bodyBIns": bodyBIns,
		}, "H1_H3")
	}
	// #endregion

	flipH := xfrm.FlipH == "1"
	flipV := xfrm.FlipV == "1"

	// Force-disable flip for standard shapes to prevent scale() accumulation in SVG.
	// PPT usually bakes flips directly into path geometry or ignores them for symmetric shapes.
	if shapeType != "picture" {
		flipH = false
		flipV = false
	}

	rs := &RenderShape{
		XmlID:                 cnvPr.ID,
		Name:                  cnvPr.Name,
		ShapeType:             shapeType,
		PrstGeom:              prstGeom,
		Rotation:              xfrm.Rot,
		FlipH:                 flipH,
		FlipV:                 flipV,
		X:                     x,
		Y:                     y,
		W:                     w,
		H:                     h,
		TextX:                 txX,
		TextY:                 txY,
		TextW:                 txW,
		TextH:                 txH,
		Text:                  text,
		Fill:                  fill,
		Stroke:                stroke,
		FontSize:              fontSize,
		Bold:                  bold,
		Alignment:             alignment,
		AnchorCtr:             sp.TxBody.BodyPr.AnchorCtr == "1" || sp.TxBody.BodyPr.AnchorCtr == "true",
		TextAnchor:            textAnchor,
		RoundRadius:           roundRadius,
		CustPathD:             custPathD,
		PreRenderedSVG:        preRenderedSVG,
		Paragraphs:            paragraphs,
		FontFamily:            fontFamily,
		TextColor:             textColor,
		GradientDef:           gradDef,
		DashStyle:             dashStyle,
		StrokeWidth:           strokeWidth,
		LineCap:               lineCap,
		FillOpacity:           fillOpacity,
		SoftEdgeRadius:        softEdgeRadius,
		Shadow:                extractShapeShadow(sp.SpPr.EffectLst, themeColors),
		PatternFill:           extractPatternFill(sp.SpPr.PattFill, themeColors),
		TextDirection:         sp.TxBody.BodyPr.Vert,
		BodyLIns:              bodyLIns,
		BodyRIns:              bodyRIns,
		BodyTIns:              bodyTIns,
		BodyBIns:              bodyBIns,
		TextAutofit:           textAutofit,
		NormFontScalePct:      normFontScalePct,
		NormLnSpcReductionPct: normLnSpcReductionPct,
		WrapText:              wrapText,
		ImageData:             blipData,
		ImageMIME:             blipMIME,
	}
	// Resolve shape-level hyperlink from cNvPr
	if cnvPr.HlinkClick != nil && cnvPr.HlinkClick.RID != "" {
		rs.HyperlinkURL = resolveHyperlinkURL(cnvPr.HlinkClick.RID, ctx.slideRels)
	}
	return rs
}

// parsePercentString parses a percent value to a float (e.g. 92.0 for 92%).
// Accepts: "92.000%" or "100%" (percent string), or "62500" (thousandths of percent per OOXML, e.g. 62500 â†’ 62.5%). Returns 0 if invalid.
func parsePercentString(s string) float64 {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0
	}
	hadPct := strings.HasSuffix(s, "%")
	s = strings.TrimSuffix(s, "%")
	s = strings.TrimSpace(s)
	if s == "" {
		return 0
	}
	f, err := strconv.ParseFloat(s, 64)
	if err != nil {
		return 0
	}
	// OOXML can use integer thousandths of percent (e.g. fontScale="62500" = 62.5%)
	if !hadPct && f > 0 && f >= 1000 {
		return f / 1000
	}
	return f
}

func classifyShape(sp OOXMLShape) string {
	name := sp.NvSpPr.CNvPr.Name

	// Placeholder-based
	if ph := sp.NvSpPr.NvPr.Ph; ph != nil {
		switch ph.Type {
		case "title", "ctrTitle":
			return "title"
		case "body", "subTitle":
			return "body"
		}
		if ph.Idx != "" {
			return "body"
		}
	}

	// TextBox flag takes priority for text containers
	if sp.NvSpPr.CNvSpPr.TxBox == "1" || strings.Contains(name, "TextBox") {
		return "textbox"
	}

	// Custom geometry
	if sp.SpPr.CustGeom != nil || strings.Contains(name, "Freeform") {
		return "freeform"
	}

	// Geometry-based â€” only classify as rect/roundRect if the shape has visible fill.
	// Shapes with prstGeom="rect" but no fill (FillRef.Idx=0, no explicit fill) are textboxes.
	if pg := sp.SpPr.PrstGeom; pg != nil {
		// Ellipse / Oval
		if pg.Prst == "ellipse" || pg.Prst == "oval" {
			return "ellipse"
		}

		if pg.Prst == "roundRect" {
			if shapeHasVisibleFill(sp) {
				return "roundRect"
			}
			return "textbox"
		}
		if pg.Prst == "rect" {
			if strings.Contains(name, "Rounded") {
				if shapeHasVisibleFill(sp) {
					return "roundRect"
				}
				return "textbox"
			}
			if shapeHasVisibleFill(sp) {
				return "rect"
			}
			return "textbox"
		}

		// Filled preset shapes (arrows, polygons, 3D, etc.)
		if isFilledPreset(pg.Prst) && shapeHasVisibleFill(sp) {
			return "prstGeomFilled"
		}
		// Line-only preset shapes (braces, brackets, arcs without fill)
		if isLinePreset(pg.Prst) && shapeHasVisibleStroke(sp) {
			return "prstGeomFilled" // reuse same rendering path; fill will resolve to "none"
		}
	}

	// Name-based fallbacks
	if strings.Contains(name, "Rectangle") {
		if shapeHasVisibleFill(sp) {
			if strings.Contains(name, "Rounded") {
				return "roundRect"
			}
			return "rect"
		}
		return "textbox"
	}

	return "textbox"
}

// shapeHasVisibleFill returns true if the shape has an explicit fill or a valid style fill reference.
func shapeHasVisibleFill(sp OOXMLShape) bool {
	// Explicit noFill
	if sp.SpPr.NoFill != nil {
		return false
	}
	// Has explicit solid fill
	if sp.SpPr.SolidFill != nil {
		return true
	}
	// Has gradient fill
	if sp.SpPr.GradFill != nil {
		return true
	}
	// Has style fill reference with idx > 0
	fr := sp.Style.FillRef
	if fr.Idx != "" && fr.Idx != "0" {
		return true
	}
	return false
}

// shapeHasVisibleStroke returns true if the shape has a visible stroke/line.
func shapeHasVisibleStroke(sp OOXMLShape) bool {
	// Has explicit line fill
	if sp.SpPr.Ln != nil && sp.SpPr.Ln.SolidFill != nil {
		return true
	}
	// Has style line reference with idx > 0
	lr := sp.Style.LnRef
	if lr.Idx != "" && lr.Idx != "0" {
		return true
	}
	return false
}

func hasStyleRef(style OOXMLStyle) bool {
	return style.FillRef.Idx != "" && style.FillRef.Idx != "0" ||
		style.LnRef.Idx != "" && style.LnRef.Idx != "0"
}

// ======================== Group Processing ========================

func applyGroupFlips(gXfrm OOXMLXfrm, childXfrm *OOXMLXfrm) {
	if gXfrm.FlipH == "1" {
		if childXfrm.FlipH == "1" {
			childXfrm.FlipH = "0"
		} else {
			childXfrm.FlipH = "1"
		}
	}
	if gXfrm.FlipV == "1" {
		if childXfrm.FlipV == "1" {
			childXfrm.FlipV = "0"
		} else {
			childXfrm.FlipV = "1"
		}
	}
}

func processGroup(grp OOXMLGroupShape, seenTexts map[string]bool, ctx *slideContext) ([]RenderShape, []RenderConnector) {
	var shapes []RenderShape
	var conns []RenderConnector

	gXfrm := grp.GrpSpPr.Xfrm
	name := grp.NvGrpSpPr.CNvPr.Name

	// StickyNote special handling
	if strings.Contains(name, "StickyNote") || hasStickyNoteChild(grp) {
		text := extractFirstText(grp.Shapes)
		shapes = append(shapes, RenderShape{
			XmlID:     grp.NvGrpSpPr.CNvPr.ID,
			Name:      name,
			ShapeType: "stickyNote",
			X:         gXfrm.Off.X,
			Y:         gXfrm.Off.Y,
			W:         gXfrm.Ext.Cx,
			H:         gXfrm.Ext.Cy,
			Text:      text,
			Fill:      "#FFFF99",
			Stroke:    "#808080",
		})
		// Record child positions for connector lookup but don't render them separately
		return shapes, conns
	}

	// Normal group: each group gets its own seenTexts to allow identical template text
	// across sibling groups (e.g., repeated "Placeholder" / "STEP" in step diagrams).
	groupSeenTexts := make(map[string]bool)

	// Process children in document order (ECMA-376 19.3.1.45: z-order = lexical order).
	if len(grp.ChildOrder) > 0 {
		for _, ref := range grp.ChildOrder {
			switch ref.Kind {
			case "sp":
				if ref.Index >= len(grp.Shapes) {
					continue
				}
				sp := grp.Shapes[ref.Index]
				cx, cy, cw, ch := groupChildTransform(gXfrm, sp.SpPr.Xfrm.Off, sp.SpPr.Xfrm.Ext)
				modSp := sp
				modSp.SpPr.Xfrm.Off = OOXMLOff{X: cx, Y: cy}
				modSp.SpPr.Xfrm.Ext = OOXMLExt{Cx: cw, Cy: ch}
				applyGroupFlips(gXfrm, &modSp.SpPr.Xfrm)
				if rs := processShape(modSp, 0, 0, groupSeenTexts, ctx); rs != nil {
					shapes = append(shapes, *rs)
				}
			case "pic":
				if ref.Index >= len(grp.Pictures) {
					continue
				}
				pic := grp.Pictures[ref.Index]
				cx, cy, cw, ch := groupChildTransform(gXfrm, pic.SpPr.Xfrm.Off, pic.SpPr.Xfrm.Ext)
				modPic := pic
				modPic.SpPr.Xfrm.Off = OOXMLOff{X: cx, Y: cy}
				modPic.SpPr.Xfrm.Ext = OOXMLExt{Cx: cw, Cy: ch}
				applyGroupFlips(gXfrm, &modPic.SpPr.Xfrm)
				if rs := processPicture(modPic, 0, 0, ctx, ctx.slideRels); rs != nil {
					shapes = append(shapes, *rs)
				}
			case "cxnSp":
				if ref.Index >= len(grp.Connectors) {
					continue
				}
				modCxn := grp.Connectors[ref.Index]
				cx, cy, cw, ch := groupChildTransform(gXfrm, modCxn.SpPr.Xfrm.Off, modCxn.SpPr.Xfrm.Ext)
				modCxn.SpPr.Xfrm.Off = OOXMLOff{X: cx, Y: cy}
				modCxn.SpPr.Xfrm.Ext = OOXMLExt{Cx: cw, Cy: ch}
				applyGroupFlips(gXfrm, &modCxn.SpPr.Xfrm)
				if rc := processConnector(modCxn, ctx); rc != nil {
					conns = append(conns, *rc)
				}
			case "grpSp":
				if ref.Index >= len(grp.Groups) {
					continue
				}
				nested := grp.Groups[ref.Index]
				nestedXfrm := nested.GrpSpPr.Xfrm
				cx, cy, cw, ch := groupChildTransform(gXfrm, nestedXfrm.Off, nestedXfrm.Ext)
				modNested := nested
				modNested.GrpSpPr.Xfrm.Off = OOXMLOff{X: cx, Y: cy}
				modNested.GrpSpPr.Xfrm.Ext = OOXMLExt{Cx: cw, Cy: ch}
				applyGroupFlips(gXfrm, &modNested.GrpSpPr.Xfrm)
				gs, gc := processGroup(modNested, groupSeenTexts, ctx)
				shapes = append(shapes, gs...)
				conns = append(conns, gc...)
			}
		}
		return shapes, conns
	}

	// Fallback when ChildOrder not set
	for _, sp := range grp.Shapes {
		cx, cy, cw, ch := groupChildTransform(gXfrm, sp.SpPr.Xfrm.Off, sp.SpPr.Xfrm.Ext)
		modSp := sp
		modSp.SpPr.Xfrm.Off = OOXMLOff{X: cx, Y: cy}
		modSp.SpPr.Xfrm.Ext = OOXMLExt{Cx: cw, Cy: ch}
		applyGroupFlips(gXfrm, &modSp.SpPr.Xfrm)
		if rs := processShape(modSp, 0, 0, groupSeenTexts, ctx); rs != nil {
			shapes = append(shapes, *rs)
		}
	}
	for _, pic := range grp.Pictures {
		cx, cy, cw, ch := groupChildTransform(gXfrm, pic.SpPr.Xfrm.Off, pic.SpPr.Xfrm.Ext)
		modPic := pic
		modPic.SpPr.Xfrm.Off = OOXMLOff{X: cx, Y: cy}
		modPic.SpPr.Xfrm.Ext = OOXMLExt{Cx: cw, Cy: ch}
		applyGroupFlips(gXfrm, &modPic.SpPr.Xfrm)
		if rs := processPicture(modPic, 0, 0, ctx, ctx.slideRels); rs != nil {
			shapes = append(shapes, *rs)
		}
	}
	for _, cxn := range grp.Connectors {
		modCxn := cxn
		cx, cy, cw, ch := groupChildTransform(gXfrm, modCxn.SpPr.Xfrm.Off, modCxn.SpPr.Xfrm.Ext)
		modCxn.SpPr.Xfrm.Off = OOXMLOff{X: cx, Y: cy}
		modCxn.SpPr.Xfrm.Ext = OOXMLExt{Cx: cw, Cy: ch}
		applyGroupFlips(gXfrm, &modCxn.SpPr.Xfrm)
		if rc := processConnector(modCxn, ctx); rc != nil {
			conns = append(conns, *rc)
		}
	}
	for _, nested := range grp.Groups {
		nestedXfrm := nested.GrpSpPr.Xfrm
		cx, cy, cw, ch := groupChildTransform(gXfrm, nestedXfrm.Off, nestedXfrm.Ext)
		modNested := nested
		modNested.GrpSpPr.Xfrm.Off = OOXMLOff{X: cx, Y: cy}
		modNested.GrpSpPr.Xfrm.Ext = OOXMLExt{Cx: cw, Cy: ch}
		applyGroupFlips(gXfrm, &modNested.GrpSpPr.Xfrm)
		gs, gc := processGroup(modNested, groupSeenTexts, ctx)
		shapes = append(shapes, gs...)
		conns = append(conns, gc...)
	}
	return shapes, conns
}

func hasStickyNoteChild(grp OOXMLGroupShape) bool {
	for _, sp := range grp.Shapes {
		if strings.Contains(sp.NvSpPr.CNvPr.Name, "StickyNote") {
			return true
		}
	}
	return false
}

func extractFirstText(shapes []OOXMLShape) string {
	for _, sp := range shapes {
		for _, p := range sp.TxBody.Paragraphs {
			for _, r := range p.Runs {
				t := strings.TrimSpace(r.T)
				if t != "" {
					return t
				}
			}
		}
	}
	return ""
}

// groupChildTransform maps child coordinates from the group's child space to slide space.
// It also handles group-level flips (flipH/flipV) by mirroring child coordinates
// within the group's bounding box.
func groupChildTransform(gXfrm OOXMLXfrm, childOff OOXMLOff, childExt OOXMLExt) (x, y, w, h int64) {
	chOff := gXfrm.ChOff
	chExt := gXfrm.ChExt
	if chOff == nil || chExt == nil || chExt.Cx == 0 || chExt.Cy == 0 {
		// No child coordinate system defined; use group offset directly
		return gXfrm.Off.X + childOff.X, gXfrm.Off.Y + childOff.Y, childExt.Cx, childExt.Cy
	}

	// Compute relative position in child space
	relX := float64(childOff.X - chOff.X)
	relY := float64(childOff.Y - chOff.Y)
	childW := float64(childExt.Cx)
	childH := float64(childExt.Cy)

	// Apply group-level horizontal flip: mirror X within child coordinate space
	if gXfrm.FlipH == "1" {
		relX = float64(chExt.Cx) - relX - childW
	}
	// Apply group-level vertical flip: mirror Y within child coordinate space
	if gXfrm.FlipV == "1" {
		relY = float64(chExt.Cy) - relY - childH
	}

	scaleX := float64(gXfrm.Ext.Cx) / float64(chExt.Cx)
	scaleY := float64(gXfrm.Ext.Cy) / float64(chExt.Cy)
	x = gXfrm.Off.X + int64(relX*scaleX)
	y = gXfrm.Off.Y + int64(relY*scaleY)
	w = int64(childW * scaleX)
	h = int64(childH * scaleY)
	return
}

// ======================== Picture Processing ========================

func processPicture(pic OOXMLPictureShape, offX, offY int64, ctx *slideContext, rels []OOXMLRelationship) *RenderShape {
	xfrm := pic.SpPr.Xfrm

	stroke := ""
	if ctx != nil {
		stroke = resolveShapeStroke(OOXMLShape{SpPr: pic.SpPr, Style: pic.Style}, ctx.themeColors)
	}

	dashStyle := ""
	strokeWidth := 0.0
	lineCap := ""
	if spPrLn := pic.SpPr.Ln; spPrLn != nil {
		if spPrLn.PrstDash != nil {
			dashStyle = spPrLn.PrstDash.Val
		}
		if spPrLn.W > 0 {
			strokeWidth = float64(spPrLn.W) / emuPerPx
		}
		lineCap = spPrLn.Cap
	}
	if strokeWidth == 0 && pic.Style.LnRef.Idx != "" {
		switch pic.Style.LnRef.Idx {
		case "1":
			strokeWidth = 6350.0 / emuPerPx
		case "2":
			strokeWidth = 12700.0 / emuPerPx
		case "3":
			strokeWidth = 19050.0 / emuPerPx
		}
	}

	rs := &RenderShape{
		XmlID:       pic.NvPicPr.CNvPr.ID,
		Name:        pic.NvPicPr.CNvPr.Name,
		ShapeType:   "picture",
		X:           xfrm.Off.X + offX,
		Y:           xfrm.Off.Y + offY,
		W:           xfrm.Ext.Cx,
		H:           xfrm.Ext.Cy,
		Fill:        "#4472C4",
		Stroke:      stroke,
		StrokeWidth: strokeWidth,
		DashStyle:   dashStyle,
		LineCap:     lineCap,
		EmbedRID:    pic.BlipFill.Blip.Embed,
		Rotation:    xfrm.Rot,
		FlipH:       xfrm.FlipH == "1",
		FlipV:       xfrm.FlipV == "1",
		Shadow:      extractShapeShadow(pic.SpPr.EffectLst, ctx.themeColors),
	}

	// Try to extract the actual image from the ZIP
	if ctx != nil && ctx.zipReader != nil && rs.EmbedRID != "" {
		// Determine the base directory for the relationships
		// Layout rels are relative to ppt/slideLayouts/, slide rels to ppt/slides/
		var baseDir string
		if sameSlice(rels, ctx.layoutRels) {
			baseDir = "ppt/slideLayouts"
		} else {
			baseDir = "ppt/slides"
		}

		// Check for SVG blip first (prefer inline SVG paths over raster)
		svgRID := getSvgBlipRID(pic.BlipFill.Blip)
		if svgRID != "" {
			svgTarget := findRelByID(rels, svgRID)
			if svgTarget != "" {
				svgFilePath := resolveRelPath(baseDir, svgTarget)
				svgFragment := extractSVGPathsFromZip(ctx.zipReader, svgFilePath,
					xfrm.Off.X+offX, xfrm.Off.Y+offY, xfrm.Ext.Cx, xfrm.Ext.Cy)
				if svgFragment != "" {
					rs.PreRenderedSVG = svgFragment
					rs.ShapeType = "prstGeomFilled" // render as inline SVG, not as <image>
				}
			}
		}

		// Fall back to raster image if no SVG paths were extracted
		if rs.PreRenderedSVG == "" {
			imageTarget := findRelByID(rels, rs.EmbedRID)
			if imageTarget != "" {
				imagePath := resolveRelPath(baseDir, imageTarget)
				imgData, imgMIME := extractImageBase64(ctx.zipReader, imagePath)
				if imgData != "" {
					rs.ImageData = imgData
					rs.ImageMIME = imgMIME
				}
			}
		}
	}

	return rs
}

// sameSlice returns true if two relationship slices point to the same underlying data.
func sameSlice(a, b []OOXMLRelationship) bool {
	if len(a) == 0 || len(b) == 0 {
		return len(a) == 0 && len(b) == 0
	}
	return &a[0] == &b[0]
}

// ======================== Connector Processing ========================

func processConnector(cxn OOXMLConnectorShape, ctx *slideContext) *RenderConnector {
	xfrm := cxn.SpPr.Xfrm
	geomType := "straightConnector1"
	adjVal := 0

	if pg := cxn.SpPr.PrstGeom; pg != nil {
		geomType = pg.Prst
		// Extract adjustment value (e.g., bentConnector3 adj1 = midpoint %)
		for _, gd := range pg.AvLst.Gd {
			if strings.HasPrefix(gd.Name, "adj") {
				if v := parseFormulaVal(gd.Fmla); v > 0 {
					adjVal = v
				}
			}
		}
	}

	hasArrow := false
	arrowAtHead := false
	arrowW := ""
	arrowLen := ""
	if ln := cxn.SpPr.Ln; ln != nil {
		if ln.TailEnd != nil && ln.TailEnd.Type == "triangle" {
			hasArrow = true
			arrowW = ln.TailEnd.W
			arrowLen = ln.TailEnd.Len
		}
		if ln.HeadEnd != nil && ln.HeadEnd.Type == "triangle" {
			if !hasArrow {
				hasArrow = true
				arrowAtHead = true
				arrowW = ln.HeadEnd.W
				arrowLen = ln.HeadEnd.Len
			}
			// Note: if both headEnd and tailEnd are arrows, tailEnd wins (common case)
		}
	}

	// Parse dash style from <a:ln><a:prstDash val="dash"/>
	dashStyle := ""
	if ln := cxn.SpPr.Ln; ln != nil && ln.PrstDash != nil && ln.PrstDash.Val != "" {
		dashStyle = ln.PrstDash.Val
	}

	// Determine stroke width:
	// 1. Explicit <a:ln w="..."> in EMU
	// 2. Fallback to lnRef idx â†’ standard theme widths
	strokeWidth := 0.0
	if ln := cxn.SpPr.Ln; ln != nil && ln.W > 0 {
		strokeWidth = float64(ln.W) / emuPerPx
	}
	if strokeWidth == 0 {
		// Map lnRef idx to standard Office theme line widths
		lnIdx := cxn.Style.LnRef.Idx
		switch lnIdx {
		case "1":
			strokeWidth = 1.33 // 12700 EMU = 1pt
		case "2":
			strokeWidth = 2.67 // 25400 EMU = 2pt
		case "3":
			strokeWidth = 4.0 // 38100 EMU = 3pt
		default:
			strokeWidth = 1.33 // default thin line
		}
	}

	var themeColors map[string]string
	if ctx != nil {
		themeColors = ctx.themeColors
	}
	color := resolveConnectorColor(cxn, themeColors)

	// Extract gradient fill on line stroke (if present)
	var gradColors []string
	var gradPositions []int
	if ln := cxn.SpPr.Ln; ln != nil && ln.GradFill != nil && len(ln.GradFill.GsLst.Gs) > 0 {
		for _, gs := range ln.GradFill.GsLst.Gs {
			c := resolveGradStopColor(gs, themeColors)
			if c == "" {
				c = "#000000"
			}
			gradColors = append(gradColors, c)
			gradPositions = append(gradPositions, gs.Pos)
		}
	}

	// Bounding box in EMU
	offX := xfrm.Off.X
	offY := xfrm.Off.Y
	extW := xfrm.Ext.Cx
	extH := xfrm.Ext.Cy

	// ----- Compute all path points in local (pre-transform) coordinates -----
	// OOXML preset geometry base paths (Appendix D of ECMA-376):
	//   curvedConnector2: M(0,0) C(0,h/2)(w/2,h)(w,h) â€” quarter-circle turn
	//   curvedConnector3: S-curve through midpoint (uses adj1, default 50000)
	// Straight/bent connectors just need start=(0,0) end=(w,h).
	type point struct{ x, y int64 }

	// Base path: start at (0,0), end at (w,h) in local coords.
	p0 := point{0, 0}
	p3 := point{extW, extH}

	// For curved connectors, define control points per spec.
	hasCurveCP := false
	var cp1, cp2 point
	if geomType == "curvedConnector2" {
		// OOXML spec: M(0,0) C(0, h/2)(w/2, h)(w, h)
		cp1 = point{0, extH / 2}
		cp2 = point{extW / 2, extH}
		hasCurveCP = true
	} else if strings.HasPrefix(geomType, "curvedConnector") {
		// curvedConnector3: S-curve with adj1.
		// Spec: M(0,0) C(x1,0)(x1,h/2)(x1,h/2) then C(x1,h)(w,h)(w,h)
		// We approximate the two-segment bezier as a single cubic through the midpoint.
		adj := int64(50000)
		if adjVal > 0 {
			adj = int64(adjVal)
		}
		x1 := extW * adj / 100000
		cp1 = point{x1, 0}
		cp2 = point{x1, extH}
		hasCurveCP = true
	}

	// ----- Apply flipH / flipV in local coords -----
	flipLocalX := func(p point) point {
		return point{extW - p.x, p.y}
	}
	flipLocalY := func(p point) point {
		return point{p.x, extH - p.y}
	}

	if xfrm.FlipH == "1" {
		p0 = flipLocalX(p0)
		p3 = flipLocalX(p3)
		if hasCurveCP {
			cp1 = flipLocalX(cp1)
			cp2 = flipLocalX(cp2)
		}
	}
	if xfrm.FlipV == "1" {
		p0 = flipLocalY(p0)
		p3 = flipLocalY(p3)
		if hasCurveCP {
			cp1 = flipLocalY(cp1)
			cp2 = flipLocalY(cp2)
		}
	}

	// ----- Convert local coords â†’ global EMU coords (add offset) -----
	toGlobal := func(p point) point {
		return point{offX + p.x, offY + p.y}
	}
	p0 = toGlobal(p0)
	p3 = toGlobal(p3)
	if hasCurveCP {
		cp1 = toGlobal(cp1)
		cp2 = toGlobal(cp2)
	}

	// ----- Apply rotation (clockwise in screen coords, 60000ths of a degree) -----
	// OOXML positive rotation = clockwise as seen on screen (y-down).
	// CW rotation in screen coords: x' = cx + dx*cos(Î¸) - dy*sin(Î¸)
	//                                y' = cy + dx*sin(Î¸) + dy*cos(Î¸)
	if xfrm.Rot != 0 {
		cx := offX + extW/2
		cy := offY + extH/2
		angle := float64(xfrm.Rot) / 60000.0 * math.Pi / 180.0
		cosA := math.Cos(angle)
		sinA := math.Sin(angle)

		rotatePoint := func(p point) point {
			dx := float64(p.x - cx)
			dy := float64(p.y - cy)
			rx := float64(cx) + dx*cosA - dy*sinA
			ry := float64(cy) + dx*sinA + dy*cosA
			return point{int64(math.Round(rx)), int64(math.Round(ry))}
		}
		p0 = rotatePoint(p0)
		p3 = rotatePoint(p3)
		if hasCurveCP {
			cp1 = rotatePoint(cp1)
			cp2 = rotatePoint(cp2)
		}
	}

	return &RenderConnector{
		SX:            p0.x,
		SY:            p0.y,
		EX:            p3.x,
		EY:            p3.y,
		CP1X:          cp1.x,
		CP1Y:          cp1.y,
		CP2X:          cp2.x,
		CP2Y:          cp2.y,
		HasCurveCP:    hasCurveCP,
		GeomType:      geomType,
		AdjVal:        adjVal,
		HasArrow:      hasArrow,
		ArrowAtHead:   arrowAtHead,
		Color:         color,
		DashStyle:     dashStyle,
		StrokeWidth:   strokeWidth,
		GradColors:    gradColors,
		GradPositions: gradPositions,
		ShapeID:       cxn.NvCxnSpPr.CNvPr.ID,
		ArrowW:        arrowW,
		ArrowLen:      arrowLen,
	}
}

func parseFormulaVal(fmla string) int {
	// Format: "val 50000"
	parts := strings.Fields(fmla)
	if len(parts) == 2 && parts[0] == "val" {
		v, _ := strconv.Atoi(parts[1])
		return v
	}
	return 0
}

func resolveConnectorColor(cxn OOXMLConnectorShape, themeColors map[string]string) string {
	// Check line solid fill
	if ln := cxn.SpPr.Ln; ln != nil && ln.SolidFill != nil {
		if c := resolveSolidFill(ln.SolidFill, themeColors); c != "" {
			return c
		}
	}
	// Check style lnRef scheme color
	if sc := cxn.Style.LnRef.SchemeClr; sc != nil {
		return resolveSchemeColor(sc.Val, themeColors)
	}
	return resolveSchemeColor("accent1", themeColors)
}

// ======================== Background Resolution ========================

// resolveSlideBackground resolves the background from slide â†’ layout â†’ master cascade.
func resolveSlideBackground(slide *OOXMLSlide, layout *OOXMLSlideLayout, themeColors map[string]string) *SlideBackground {
	// 1. Check slide-level background
	if slide.CSld.Bg != nil && slide.CSld.Bg.BgPr != nil {
		return resolveBgPr(slide.CSld.Bg.BgPr, themeColors)
	}
	// 2. Check layout-level background
	if layout != nil && layout.CSld.Bg != nil && layout.CSld.Bg.BgPr != nil {
		return resolveBgPr(layout.CSld.Bg.BgPr, themeColors)
	}
	// 3. No explicit background found
	return nil
}

// resolveBgPr creates a SlideBackground from background properties.
func resolveBgPr(bgPr *OOXMLBgPr, themeColors map[string]string) *SlideBackground {
	bg := &SlideBackground{}

	if bgPr.SolidFill != nil {
		bg.SolidFill = resolveSolidFill(bgPr.SolidFill, themeColors)
		return bg
	}

	if bgPr.GradFill != nil && len(bgPr.GradFill.GsLst.Gs) > 0 {
		// Build gradient for full-slide background (shape ID 0, position 0,0, size = slide EMU dimensions)
		slideWEmu := int64(svgWidth) * int64(emuPerPx)
		slideHEmu := int64(svgHeight) * int64(emuPerPx)
		bg.GradientDef = buildSVGGradient(bgPr.GradFill, 0, 0, 0, slideWEmu, slideHEmu, themeColors)
		if bg.GradientDef != nil {
			bg.GradientDef.ID = "bgGrad"
		}
		return bg
	}

	return nil
}

// ======================== Color Resolution ========================

func resolveShapeFill(sp OOXMLShape, shapeType string, themeColors map[string]string) (fill, stroke string) {
	// 1. Explicit solid fill on shape properties
	if sf := sp.SpPr.SolidFill; sf != nil {
		fill = resolveSolidFill(sf, themeColors)
		stroke = resolveShapeStroke(sp, themeColors)
		return
	}

	// 2. Gradient fill â€” use first stop as fallback solid color
	// (actual gradient rendering happens in processShape via buildSVGGradient)
	if gf := sp.SpPr.GradFill; gf != nil && len(gf.GsLst.Gs) > 0 {
		// Use first stop as the fill color (fallback for non-gradient renderers)
		fill = resolveGradStopColor(gf.GsLst.Gs[0], themeColors)
		stroke = resolveShapeStroke(sp, themeColors)
		return
	}

	// 3. Explicit no-fill â€” mark fill as "none", still resolve stroke from line properties
	if sp.SpPr.NoFill != nil {
		stroke = resolveShapeStroke(sp, themeColors)
		return "none", stroke
	}

	// 4. Style fill reference (theme-based)
	// Only use style fill if the FillRef.Idx is a non-zero number (idx="0" means no fill)
	fr := sp.Style.FillRef
	if fr.Idx != "" && fr.Idx != "0" && fr.SchemeClr != nil {
		fill = resolveSchemeColor(fr.SchemeClr.Val, themeColors)
		fill = applyColorTransform(fill, fr.SchemeClr.HueMod, fr.SchemeClr.HueOff, fr.SchemeClr.SatMod, fr.SchemeClr.SatOff, fr.SchemeClr.LumMod, fr.SchemeClr.LumOff, fr.SchemeClr.Shade, fr.SchemeClr.Tint)
		stroke = resolveShapeStroke(sp, themeColors)
		return
	}

	// 5. No fill found â€” return empty (transparent) but resolve stroke
	stroke = resolveShapeStroke(sp, themeColors)
	return "", stroke
}

// resolveShapeStroke determines the stroke color from explicit line props or style LnRef.
func resolveShapeStroke(sp OOXMLShape, themeColors map[string]string) string {
	// 0. If the line explicitly declares <a:noFill/>, suppress stroke
	if ln := sp.SpPr.Ln; ln != nil && ln.NoFill != nil {
		return "none"
	}
	// 1. Explicit line solid fill
	if ln := sp.SpPr.Ln; ln != nil && ln.SolidFill != nil {
		if c := resolveSolidFill(ln.SolidFill, themeColors); c != "" {
			return c
		}
	}
	// 2. Style LnRef scheme color
	if sc := sp.Style.LnRef.SchemeClr; sc != nil {
		c := resolveSchemeColor(sc.Val, themeColors)
		return applyColorTransform(c, sc.HueMod, sc.HueOff, sc.SatMod, sc.SatOff, sc.LumMod, sc.LumOff, sc.Shade, sc.Tint)
	}
	return ""
}

// applyLumMod applies a lumMod (luminance modulation) to a hex color using HSL.
// lumModVal is in 1/100000 units (e.g., 75000 = 75%).
func applyLumMod(hexColor string, lumModVal int) string {
	if len(hexColor) != 7 || hexColor[0] != '#' {
		return hexColor
	}
	ri, _ := strconv.ParseInt(hexColor[1:3], 16, 64)
	gi, _ := strconv.ParseInt(hexColor[3:5], 16, 64)
	bi, _ := strconv.ParseInt(hexColor[5:7], 16, 64)
	rf := float64(ri) / 255.0
	gf := float64(gi) / 255.0
	bf := float64(bi) / 255.0

	// RGB to HSL
	maxC := math.Max(rf, math.Max(gf, bf))
	minC := math.Min(rf, math.Min(gf, bf))
	l := (maxC + minC) / 2.0
	var h, s float64
	if maxC == minC {
		h, s = 0, 0
	} else {
		d := maxC - minC
		if l > 0.5 {
			s = d / (2.0 - maxC - minC)
		} else {
			s = d / (maxC + minC)
		}
		switch maxC {
		case rf:
			h = (gf - bf) / d
			if gf < bf {
				h += 6
			}
		case gf:
			h = (bf-rf)/d + 2
		case bf:
			h = (rf-gf)/d + 4
		}
		h /= 6
	}

	// Apply lumMod: multiply luminance
	factor := float64(lumModVal) / 100000.0
	l *= factor
	if l > 1 {
		l = 1
	}
	if l < 0 {
		l = 0
	}

	// HSL to RGB
	hslToRGB := func(h, s, l float64) (int, int, int) {
		if s == 0 {
			v := int(math.Round(l * 255))
			return v, v, v
		}
		var q float64
		if l < 0.5 {
			q = l * (1 + s)
		} else {
			q = l + s - l*s
		}
		p := 2*l - q
		hue2rgb := func(p, q, t float64) float64 {
			if t < 0 {
				t += 1
			}
			if t > 1 {
				t -= 1
			}
			switch {
			case t < 1.0/6.0:
				return p + (q-p)*6*t
			case t < 1.0/2.0:
				return q
			case t < 2.0/3.0:
				return p + (q-p)*(2.0/3.0-t)*6
			default:
				return p
			}
		}
		rr := int(math.Round(hue2rgb(p, q, h+1.0/3.0) * 255))
		gg := int(math.Round(hue2rgb(p, q, h) * 255))
		bb := int(math.Round(hue2rgb(p, q, h-1.0/3.0) * 255))
		return rr, gg, bb
	}

	ro, go2, bo := hslToRGB(h, s, l)
	return fmt.Sprintf("#%02X%02X%02X", ro, go2, bo)
}

// applyShade darkens a color toward black in linear RGB space per the OOXML spec.
// shadeVal is in 1/100000 units (e.g., 60000 = 60%).
func applyShade(hex string, shadeVal int) string {
	if len(hex) != 7 || hex[0] != '#' {
		return hex
	}
	r, _ := strconv.ParseInt(hex[1:3], 16, 64)
	g, _ := strconv.ParseInt(hex[3:5], 16, 64)
	b, _ := strconv.ParseInt(hex[5:7], 16, 64)
	s := float64(shadeVal) / 100000.0
	nr := int(math.Round(float64(r) * s))
	ng := int(math.Round(float64(g) * s))
	nb := int(math.Round(float64(b) * s))

	if nr > 255 {
		nr = 255
	}
	if ng > 255 {
		ng = 255
	}
	if nb > 255 {
		nb = 255
	}

	return fmt.Sprintf("#%02X%02X%02X", nr, ng, nb)
}

// prstDashToSVG converts an OOXML prstDash value to an SVG stroke-dasharray attribute string.
// OOXML dash patterns are defined in multiples of the stroke width.
// When round or square caps are used, the visible dash is extended by one stroke-width
// (half on each end), so we compensate by reducing dash and increasing gap by 1*sw.
// Returns empty string for solid or unknown styles.
func prstDashToSVG(dashStyle string, strokeWidth float64, lineCap string) string {
	if strokeWidth <= 0 {
		strokeWidth = 1.33
	}
	sw := strokeWidth

	// Round/square caps extend visible dash length by ~1*sw, so adjust:
	// reduce dash by 1*sw, increase gap by 1*sw to match expected visual output.
	capAdj := 0.0
	if lineCap == "round" || lineCap == "square" {
		capAdj = sw
	}

	// OOXML dash pattern values in multiples of stroke width
	type pattern struct {
		dash float64
		gap  float64
	}
	multi := func(pairs ...pattern) string {
		parts := make([]string, 0, len(pairs)*2)
		for _, p := range pairs {
			d := p.dash*sw - capAdj
			g := p.gap*sw + capAdj
			if d < 0 {
				d = 0.1
			}
			parts = append(parts, fmt.Sprintf("%.5g", d))
			parts = append(parts, fmt.Sprintf("%.5g", g))
		}
		return fmt.Sprintf(` stroke-dasharray="%s"`, strings.Join(parts, " "))
	}

	switch dashStyle {
	case "dash":
		return multi(pattern{4, 3})
	case "dot":
		return multi(pattern{1, 1})
	case "lgDash":
		return multi(pattern{8, 3})
	case "dashDot":
		return multi(pattern{4, 3}, pattern{1, 3})
	case "lgDashDot":
		return multi(pattern{8, 3}, pattern{1, 3})
	case "lgDashDotDot":
		return multi(pattern{8, 3}, pattern{1, 3}, pattern{1, 3})
	case "sysDash":
		return multi(pattern{3, 1})
	case "sysDot":
		return multi(pattern{1, 1})
	case "sysDashDot":
		return multi(pattern{3, 1}, pattern{1, 1})
	case "sysDashDotDot":
		return multi(pattern{3, 1}, pattern{1, 1}, pattern{1, 1})
	default:
		return ""
	}
}

// applyLumModOff applies lumMod and lumOff to a hex color using HSL.
// lumMod is in 1/100000 units (e.g., 20000 = 20%), lumOff likewise (e.g., 80000 = 80%).
// Formula: new_luminance = luminance * lumMod/100000 + lumOff/100000
func applyHSLModifiers(hexColor string, hueMod, hueOff, satMod, satOff, lumMod, lumOff *OOXMLIntVal) string {
	if len(hexColor) != 7 || hexColor[0] != '#' {
		return hexColor
	}
	ri, _ := strconv.ParseInt(hexColor[1:3], 16, 64)
	gi, _ := strconv.ParseInt(hexColor[3:5], 16, 64)
	bi, _ := strconv.ParseInt(hexColor[5:7], 16, 64)
	rf := float64(ri) / 255.0
	gf := float64(gi) / 255.0
	bf := float64(bi) / 255.0

	// RGB to HSL
	maxC := math.Max(rf, math.Max(gf, bf))
	minC := math.Min(rf, math.Min(gf, bf))
	l := (maxC + minC) / 2.0
	var h, s float64
	if maxC == minC {
		h, s = 0, 0
	} else {
		d := maxC - minC
		if l > 0.5 {
			s = d / (2.0 - maxC - minC)
		} else {
			s = d / (maxC + minC)
		}
		switch maxC {
		case rf:
			h = (gf - bf) / d
			if gf < bf {
				h += 6
			}
		case gf:
			h = (bf-rf)/d + 2
		case bf:
			h = (rf-gf)/d + 4
		}
		h /= 6
	}

	// Apply Modifiers
	if hueMod != nil {
		h *= float64(hueMod.Val) / 100000.0
	}
	if hueOff != nil {
		h += float64(hueOff.Val) / 21600000.0 // 360 degrees * 60000 = 21600000
	}
	for h < 0 {
		h += 1.0
	}
	for h >= 1.0 {
		h -= 1.0
	}

	if satMod != nil {
		s *= float64(satMod.Val) / 100000.0
	}
	if satOff != nil {
		s += float64(satOff.Val) / 100000.0
	}
	if s > 1 {
		s = 1
	}
	if s < 0 {
		s = 0
	}

	if lumMod != nil {
		l *= float64(lumMod.Val) / 100000.0
	}
	if lumOff != nil {
		l += float64(lumOff.Val) / 100000.0
	}
	if l > 1 {
		l = 1
	}
	if l < 0 {
		l = 0
	}

	// HSL to RGB
	if s == 0 {
		v := int(math.Round(l * 255))
		return fmt.Sprintf("#%02X%02X%02X", v, v, v)
	}
	var q float64
	if l < 0.5 {
		q = l * (1 + s)
	} else {
		q = l + s - l*s
	}
	p := 2*l - q
	hue2rgb := func(p, q, t float64) float64 {
		if t < 0 {
			t += 1
		}
		if t > 1 {
			t -= 1
		}
		switch {
		case t < 1.0/6.0:
			return p + (q-p)*6*t
		case t < 1.0/2.0:
			return q
		case t < 2.0/3.0:
			return p + (q-p)*(2.0/3.0-t)*6
		default:
			return p
		}
	}
	ro := int(math.Round(hue2rgb(p, q, h+1.0/3.0) * 255))
	go2 := int(math.Round(hue2rgb(p, q, h) * 255))
	bo := int(math.Round(hue2rgb(p, q, h-1.0/3.0) * 255))
	return fmt.Sprintf("#%02X%02X%02X", ro, go2, bo)
}

func applyColorTransform(c string, hueMod, hueOff, satMod, satOff, lumMod, lumOff, shade, tint *OOXMLIntVal) string {
	if tint != nil && tint.Val > 0 {
		c = applyTint(c, tint.Val)
	}
	if shade != nil && shade.Val > 0 {
		c = applyShade(c, shade.Val)
	}
	if hueMod != nil || hueOff != nil || satMod != nil || satOff != nil || lumMod != nil || lumOff != nil {
		c = applyHSLModifiers(c, hueMod, hueOff, satMod, satOff, lumMod, lumOff)
	}
	return c
}

func applyLumModOff(hexColor string, lumMod, lumOff int) string {
	if len(hexColor) != 7 || hexColor[0] != '#' {
		return hexColor
	}
	ri, _ := strconv.ParseInt(hexColor[1:3], 16, 64)
	gi, _ := strconv.ParseInt(hexColor[3:5], 16, 64)
	bi, _ := strconv.ParseInt(hexColor[5:7], 16, 64)
	rf := float64(ri) / 255.0
	gf := float64(gi) / 255.0
	bf := float64(bi) / 255.0

	// RGB to HSL
	maxC := math.Max(rf, math.Max(gf, bf))
	minC := math.Min(rf, math.Min(gf, bf))
	l := (maxC + minC) / 2.0
	var h, s float64
	if maxC == minC {
		h, s = 0, 0
	} else {
		d := maxC - minC
		if l > 0.5 {
			s = d / (2.0 - maxC - minC)
		} else {
			s = d / (maxC + minC)
		}
		switch maxC {
		case rf:
			h = (gf - bf) / d
			if gf < bf {
				h += 6
			}
		case gf:
			h = (bf-rf)/d + 2
		case bf:
			h = (rf-gf)/d + 4
		}
		h /= 6
	}

	// Apply lumMod + lumOff
	modF := float64(lumMod) / 100000.0
	offF := float64(lumOff) / 100000.0
	l = l*modF + offF
	if l > 1 {
		l = 1
	}
	if l < 0 {
		l = 0
	}

	// HSL to RGB (reuse hslToRGB logic)
	hslToRGB := func(h, s, l float64) (int, int, int) {
		if s == 0 {
			v := int(math.Round(l * 255))
			return v, v, v
		}
		var q float64
		if l < 0.5 {
			q = l * (1 + s)
		} else {
			q = l + s - l*s
		}
		p := 2*l - q
		hue2rgb := func(p, q, t float64) float64 {
			if t < 0 {
				t += 1
			}
			if t > 1 {
				t -= 1
			}
			switch {
			case t < 1.0/6.0:
				return p + (q-p)*6*t
			case t < 1.0/2.0:
				return q
			case t < 2.0/3.0:
				return p + (q-p)*(2.0/3.0-t)*6
			default:
				return p
			}
		}
		rr := int(math.Round(hue2rgb(p, q, h+1.0/3.0) * 255))
		gg := int(math.Round(hue2rgb(p, q, h) * 255))
		bb := int(math.Round(hue2rgb(p, q, h-1.0/3.0) * 255))
		return rr, gg, bb
	}

	ro, go2, bo := hslToRGB(h, s, l)
	return fmt.Sprintf("#%02X%02X%02X", ro, go2, bo)
}

// applyTint applies a tint modifier (mix with white).
// tintVal is in 1/100000 units (e.g., 66000 = 66%).
// Formula: new = color * tint/100000 + white * (1 - tint/100000)
// srgbToLinear converts an sRGB channel value (0-255) to linear light (0.0-1.0).
func srgbToLinear(c float64) float64 {
	c /= 255.0
	if c <= 0.04045 {
		return c / 12.92
	}
	return math.Pow((c+0.055)/1.055, 2.4)
}

// linearToSrgb converts a linear light value (0.0-1.0) to sRGB channel (0-255).
func linearToSrgb(c float64) int {
	if c <= 0.0 {
		return 0
	}
	if c >= 1.0 {
		return 255
	}
	var s float64
	if c <= 0.0031308 {
		s = c * 12.92
	} else {
		s = 1.055*math.Pow(c, 1.0/2.4) - 0.055
	}
	return int(math.Round(s * 255.0))
}

func applyTint(hexColor string, tintVal int) string {
	if len(hexColor) != 7 || hexColor[0] != '#' {
		return hexColor
	}
	r, _ := strconv.ParseInt(hexColor[1:3], 16, 64)
	g, _ := strconv.ParseInt(hexColor[3:5], 16, 64)
	b, _ := strconv.ParseInt(hexColor[5:7], 16, 64)
	t := float64(tintVal) / 100000.0

	nr := int(math.Round(float64(r)*t + 255.0*(1.0-t)))
	ng := int(math.Round(float64(g)*t + 255.0*(1.0-t)))
	nb := int(math.Round(float64(b)*t + 255.0*(1.0-t)))

	if nr > 255 {
		nr = 255
	}
	if ng > 255 {
		ng = 255
	}
	if nb > 255 {
		nb = 255
	}

	return fmt.Sprintf("#%02X%02X%02X", nr, ng, nb)
}

// applySatMod applies a saturation modulation to a hex color using HSL.
// satModVal is in 1/100000 units (e.g., 160000 = 160%).
func applySatMod(hexColor string, satModVal int) string {
	if len(hexColor) != 7 || hexColor[0] != '#' {
		return hexColor
	}
	ri, _ := strconv.ParseInt(hexColor[1:3], 16, 64)
	gi, _ := strconv.ParseInt(hexColor[3:5], 16, 64)
	bi, _ := strconv.ParseInt(hexColor[5:7], 16, 64)
	rf := float64(ri) / 255.0
	gf := float64(gi) / 255.0
	bf := float64(bi) / 255.0

	maxC := math.Max(rf, math.Max(gf, bf))
	minC := math.Min(rf, math.Min(gf, bf))
	l := (maxC + minC) / 2.0
	var h, s float64
	if maxC == minC {
		return hexColor // achromatic, satMod has no effect
	}
	d := maxC - minC
	if l > 0.5 {
		s = d / (2.0 - maxC - minC)
	} else {
		s = d / (maxC + minC)
	}
	switch maxC {
	case rf:
		h = (gf - bf) / d
		if gf < bf {
			h += 6
		}
	case gf:
		h = (bf-rf)/d + 2
	case bf:
		h = (rf-gf)/d + 4
	}
	h /= 6

	// Apply satMod
	s *= float64(satModVal) / 100000.0
	if s > 1 {
		s = 1
	}
	if s < 0 {
		s = 0
	}

	// HSL to RGB
	hslToRGB := func(h, s, l float64) (int, int, int) {
		if s == 0 {
			v := int(math.Round(l * 255))
			return v, v, v
		}
		var q float64
		if l < 0.5 {
			q = l * (1 + s)
		} else {
			q = l + s - l*s
		}
		p := 2*l - q
		hue2rgb := func(p, q, t float64) float64 {
			if t < 0 {
				t += 1
			}
			if t > 1 {
				t -= 1
			}
			switch {
			case t < 1.0/6.0:
				return p + (q-p)*6*t
			case t < 1.0/2.0:
				return q
			case t < 2.0/3.0:
				return p + (q-p)*(2.0/3.0-t)*6
			default:
				return p
			}
		}
		rr := int(math.Round(hue2rgb(p, q, h+1.0/3.0) * 255))
		gg := int(math.Round(hue2rgb(p, q, h) * 255))
		bb := int(math.Round(hue2rgb(p, q, h-1.0/3.0) * 255))
		return rr, gg, bb
	}

	ro, go2, bo := hslToRGB(h, s, l)
	return fmt.Sprintf("#%02X%02X%02X", ro, go2, bo)
}

// resolveGradStopColor resolves a single gradient stop's color from OOXML types.
func resolveGradStopColor(gs OOXMLGs, themeColors map[string]string) string {
	var c string
	if gs.SrgbClr != nil {
		c = "#" + gs.SrgbClr.Val
		c = applyColorTransform(c, gs.SrgbClr.HueMod, gs.SrgbClr.HueOff, gs.SrgbClr.SatMod, gs.SrgbClr.SatOff, gs.SrgbClr.LumMod, gs.SrgbClr.LumOff, gs.SrgbClr.Shade, gs.SrgbClr.Tint)
	} else if gs.SchemeClr != nil {
		c = resolveSchemeColor(gs.SchemeClr.Val, themeColors)
		c = applyColorTransform(c, gs.SchemeClr.HueMod, gs.SchemeClr.HueOff, gs.SchemeClr.SatMod, gs.SchemeClr.SatOff, gs.SchemeClr.LumMod, gs.SchemeClr.LumOff, gs.SchemeClr.Shade, gs.SchemeClr.Tint)
	}
	if c == "" {
		c = resolveSchemeColor("accent1", themeColors)
	}
	return c
}

// buildSVGGradient creates an SVGGradient from an OOXML GradFill for a given shape.
// shapeX/Y/W/H are the shape's bounding box in EMU; used to compute pixel coords for radial gradients.
func buildSVGGradient(gf *OOXMLGradFill, shapeID int, shapeX, shapeY, shapeW, shapeH int64, themeColors map[string]string) *SVGGradient {
	if gf == nil || len(gf.GsLst.Gs) == 0 {
		return nil
	}

	grad := &SVGGradient{
		ID: fmt.Sprintf("grad_%d", shapeID),
	}

	// Parse stops
	for _, gs := range gf.GsLst.Gs {
		stop := SVGGradStop{
			Offset: float64(gs.Pos) / 100000.0,
			Color:  resolveGradStopColor(gs, themeColors),
			Alpha:  1.0,
		}
		// Check alpha on stops
		if gs.SchemeClr != nil && gs.SchemeClr.Alpha != nil {
			stop.Alpha = float64(gs.SchemeClr.Alpha.Val) / 100000.0
		} else if gs.SrgbClr != nil && gs.SrgbClr.Alpha != nil {
			stop.Alpha = float64(gs.SrgbClr.Alpha.Val) / 100000.0
		}
		grad.Stops = append(grad.Stops, stop)
	}

	// Sort stops by offset (SVG requires non-decreasing order)
	sort.Slice(grad.Stops, func(i, j int) bool {
		return grad.Stops[i].Offset < grad.Stops[j].Offset
	})

	// Shape bounding box in pixels
	sPx := float64(shapeX) / emuPerPx
	sPy := float64(shapeY) / emuPerPx
	sW := float64(shapeW) / emuPerPx
	sH := float64(shapeH) / emuPerPx

	// Determine gradient type
	if gf.Path != nil {
		// Radial gradient (path="circle", "rect", "shape")
		grad.IsRadial = true

		// Focal and center fractions (0.0â€“1.0 of shape bounding box)
		fxFrac := 0.5
		fyFrac := 0.5
		cxFrac := 0.5
		cyFrac := 0.5

		if ftr := gf.Path.FillToRect; ftr != nil {
			// fillToRect defines where the gradient focal point is. The l/t/r/b
			// values define a rectangle within the bounding box:
			//   left edge   = l/100000
			//   top edge    = t/100000
			//   right edge  = 1 - r/100000
			//   bottom edge = 1 - b/100000
			// The FOCAL POINT (fx, fy) is at the center of this rectangle.
			l := float64(ftr.L) / 100000.0
			t := float64(ftr.T) / 100000.0
			r := float64(ftr.R) / 100000.0
			b := float64(ftr.B) / 100000.0

			leftEdge := l
			rightEdge := 1.0 - r
			topEdge := t
			bottomEdge := 1.0 - b

			fxFrac = (leftEdge + rightEdge) / 2.0
			fyFrac = (topEdge + bottomEdge) / 2.0

			// Outer circle center: if focal is within bounds, center = focal.
			// If focal is outside, use shape center (0.5, 0.5) for eccentric gradient.
			if fxFrac >= 0 && fxFrac <= 1 && fyFrac >= 0 && fyFrac <= 1 {
				cxFrac = fxFrac
				cyFrac = fyFrac
			} else {
				cxFrac = 0.5
				cyFrac = 0.5
			}
		}

		// Convert to pixel coordinates (userSpaceOnUse)
		grad.CX = sPx + cxFrac*sW
		grad.CY = sPy + cyFrac*sH
		grad.FX = sPx + fxFrac*sW
		grad.FY = sPy + fyFrac*sH
		grad.FXSet = grad.FX != grad.CX
		grad.FYSet = grad.FY != grad.CY

		// Radius = distance from outer circle center (cx,cy) to the farthest corner of the shape
		cx := grad.CX
		cy := grad.CY
		corners := [4][2]float64{
			{sPx, sPy},
			{sPx + sW, sPy},
			{sPx, sPy + sH},
			{sPx + sW, sPy + sH},
		}
		maxDist := 0.0
		for _, c := range corners {
			dx := c[0] - cx
			dy := c[1] - cy
			d := math.Sqrt(dx*dx + dy*dy)
			if d > maxDist {
				maxDist = d
			}
		}
		grad.RX = maxDist
	} else if gf.Lin != nil {
		// Linear gradient
		grad.IsRadial = false
		grad.Angle = float64(gf.Lin.Ang) / 60000.0 // convert from 60000ths of a degree
	} else {
		// Default: top-to-bottom linear
		grad.IsRadial = false
		grad.Angle = 270 // top to bottom in SVG
	}

	return grad
}

// renderSVGGradientDef renders an SVG gradient definition element.
func renderSVGGradientDef(g *SVGGradient) string {
	if g == nil || len(g.Stops) == 0 {
		return ""
	}

	var out strings.Builder
	if g.IsRadial {
		// Emit radialGradient in userSpaceOnUse with pixel coordinates
		fxAttr := ""
		fyAttr := ""
		if g.FXSet {
			fxAttr = fmt.Sprintf(` fx="%.4f"`, g.FX)
		}
		if g.FYSet {
			fyAttr = fmt.Sprintf(` fy="%.4f"`, g.FY)
		}
		out.WriteString(fmt.Sprintf(`    <radialGradient id="%s" cx="%.4f" cy="%.4f" r="%.4f"%s%s gradientUnits="userSpaceOnUse">
`, g.ID, g.CX, g.CY, g.RX, fxAttr, fyAttr))
	} else {
		// Convert angle to x1/y1/x2/y2
		// OOXML angle: 0 = right-to-left, 5400000 = top-to-bottom
		// SVG: x1,y1 â†’ x2,y2
		rad := g.Angle * math.Pi / 180.0
		x1 := 0.5 - 0.5*math.Cos(rad)
		y1 := 0.5 - 0.5*math.Sin(rad)
		x2 := 0.5 + 0.5*math.Cos(rad)
		y2 := 0.5 + 0.5*math.Sin(rad)
		out.WriteString(fmt.Sprintf(`    <linearGradient id="%s" x1="%.2f" y1="%.2f" x2="%.2f" y2="%.2f" gradientUnits="objectBoundingBox">
`, g.ID, x1, y1, x2, y2))
	}

	for _, s := range g.Stops {
		if s.Alpha < 1.0 {
			out.WriteString(fmt.Sprintf(`      <stop offset="%.0f%%" stop-color="%s" stop-opacity="%.2f"/>
`, s.Offset*100, s.Color, s.Alpha))
		} else {
			out.WriteString(fmt.Sprintf(`      <stop offset="%.0f%%" stop-color="%s"/>
`, s.Offset*100, s.Color))
		}
	}

	if g.IsRadial {
		out.WriteString("    </radialGradient>\n")
	} else {
		out.WriteString("    </linearGradient>\n")
	}
	return out.String()
}

// resolveColorChoice resolves a color from an OOXMLColorChoice (fgClr/bgClr container).
func resolveColorChoice(cc *OOXMLColorChoice, themeColors map[string]string) string {
	if cc == nil {
		return ""
	}
	if cc.SrgbClr != nil {
		c := "#" + cc.SrgbClr.Val
		return applyColorTransform(c, cc.SrgbClr.HueMod, cc.SrgbClr.HueOff, cc.SrgbClr.SatMod, cc.SrgbClr.SatOff, cc.SrgbClr.LumMod, cc.SrgbClr.LumOff, cc.SrgbClr.Shade, cc.SrgbClr.Tint)
	}
	if cc.SchemeClr != nil {
		c := resolveSchemeColor(cc.SchemeClr.Val, themeColors)
		return applyColorTransform(c, cc.SchemeClr.HueMod, cc.SchemeClr.HueOff, cc.SchemeClr.SatMod, cc.SchemeClr.SatOff, cc.SchemeClr.LumMod, cc.SchemeClr.LumOff, cc.SchemeClr.Shade, cc.SchemeClr.Tint)
	}
	return ""
}

func resolveSolidFill(sf *OOXMLSolidFill, themeColors map[string]string) string {
	if sf == nil {
		return ""
	}
	if sf.SrgbClr != nil {
		c := "#" + sf.SrgbClr.Val
		return applyColorTransform(c, sf.SrgbClr.HueMod, sf.SrgbClr.HueOff, sf.SrgbClr.SatMod, sf.SrgbClr.SatOff, sf.SrgbClr.LumMod, sf.SrgbClr.LumOff, sf.SrgbClr.Shade, sf.SrgbClr.Tint)
	}
	if sf.SchemeClr != nil {
		c := resolveSchemeColor(sf.SchemeClr.Val, themeColors)
		return applyColorTransform(c, sf.SchemeClr.HueMod, sf.SchemeClr.HueOff, sf.SchemeClr.SatMod, sf.SchemeClr.SatOff, sf.SchemeClr.LumMod, sf.SchemeClr.LumOff, sf.SchemeClr.Shade, sf.SchemeClr.Tint)
	}
	if sf.PrstClr != nil {
		c := resolvePresetColor(sf.PrstClr.Val)
		return applyColorTransform(c, nil, nil, nil, nil, sf.PrstClr.LumMod, sf.PrstClr.LumOff, nil, nil)
	}
	return ""
}

// applyClrMap builds the effective theme color map by applying master's clrMap (tx1â†’dk1, bg1â†’lt1, etc.).
// raw contains theme slot colors (dk1, lt1, accent1, ...). Returns a map that includes both raw slots
// and logical names (tx1, bg1, tx2, bg2) resolved via clrMap, or default mapping if clrMap is nil.
func applyClrMap(raw map[string]string, clrMap *OOXMLClrMap) map[string]string {
	out := make(map[string]string)
	for k, v := range raw {
		out[k] = v
	}
	// Logical names (tx1, bg1, tx2, bg2): use clrMap if present, else default (tx1=dk1, bg1=lt1, ...)
	setMapped := func(logical, defaultSlot string) {
		slot := defaultSlot
		if clrMap != nil {
			switch logical {
			case "tx1":
				if clrMap.Tx1 != "" {
					slot = clrMap.Tx1
				}
			case "bg1":
				if clrMap.Bg1 != "" {
					slot = clrMap.Bg1
				}
			case "tx2":
				if clrMap.Tx2 != "" {
					slot = clrMap.Tx2
				}
			case "bg2":
				if clrMap.Bg2 != "" {
					slot = clrMap.Bg2
				}
			}
		}
		if c := out[slot]; c != "" {
			out[logical] = c
		}
	}
	setMapped("tx1", "dk1")
	setMapped("bg1", "lt1")
	setMapped("tx2", "dk2")
	setMapped("bg2", "lt2")
	return out
}

// resolveSchemeColor maps OOXML scheme color names to hex values using the theme color map.
// If themeColors is nil or doesn't contain the value, falls back to Office defaults.
func resolveSchemeColor(val string, themeColors map[string]string) string {
	// Use theme colors from the actual PPTX file if available
	if themeColors != nil {
		if c, ok := themeColors[val]; ok && c != "" {
			return c
		}
	}
	// Fallback to Office default theme colors
	defaults := map[string]string{
		"accent1": "#4472C4",
		"accent2": "#ED7D31",
		"accent3": "#A5A5A5",
		"accent4": "#FFC000",
		"accent5": "#5B9BD5",
		"accent6": "#70AD47",
		"tx1":     "#000000",
		"tx2":     "#404040",
		"bg1":     "#FFFFFF",
		"bg2":     "#F5F5F5",
		"lt1":     "#FFFFFF",
		"lt2":     "#E7E6E6",
		"dk1":     "#000000",
		"dk2":     "#44546A",
	}
	if c, ok := defaults[val]; ok {
		return c
	}
	return "#4472C4"
}

func resolvePresetColor(val string) string {
	colors := map[string]string{
		"black": "#000000", "white": "#FFFFFF", "red": "#FF0000",
		"green": "#00FF00", "blue": "#0000FF", "yellow": "#FFFF00",
	}
	if c, ok := colors[val]; ok {
		return c
	}
	return "#000000"
}

func darkenColor(hex string) string {
	known := map[string]string{
		"#4472C4": "#2F528F", "#5B9BD5": "#2E75B6", "#ED7D31": "#C65911",
		"#70AD47": "#507E32", "#A5A5A5": "#7F7F7F", "#FFC000": "#BF9000",
		"#FFFF99": "#CCCC77", "#C00000": "#900000",
	}
	if d, ok := known[hex]; ok {
		return d
	}
	// Algorithmic darken: reduce each channel by 30%
	if len(hex) == 7 && hex[0] == '#' {
		r, _ := strconv.ParseInt(hex[1:3], 16, 64)
		g, _ := strconv.ParseInt(hex[3:5], 16, 64)
		b, _ := strconv.ParseInt(hex[5:7], 16, 64)
		r = int64(float64(r) * 0.7)
		g = int64(float64(g) * 0.7)
		b = int64(float64(b) * 0.7)
		return fmt.Sprintf("#%02X%02X%02X", r, g, b)
	}
	return hex
}

// ======================== Text Extraction ========================

// extractText walks paragraph/run structure and returns:
//   - combined text, font size, bold, alignment (backward-compat)
//   - structured paragraphs with per-run formatting
//   - primary font family and text color
func extractText(txBody OOXMLTxBody, seenTexts map[string]bool, themeColors map[string]string, rels []OOXMLRelationship) (string, int, bool, string, []RenderParagraph, string, string) {
	var parts []string
	fontSize := 0
	bold := false
	alignment := ""
	fontFamily := ""
	textColor := ""
	var paragraphs []RenderParagraph

	for _, para := range txBody.Paragraphs {
		paraAlign := ""
		paraLevel := 0
		if para.PPr != nil {
			if para.PPr.Algn != "" {
				paraAlign = para.PPr.Algn
				if alignment == "" {
					alignment = paraAlign
				}
			}
			paraLevel = para.PPr.Lvl // 0-based paragraph level
		}

		// Compute space before for this paragraph (in pixels)
		spaceBefore := 0
		if para.PPr != nil && para.PPr.SpcBef != nil {
			if para.PPr.SpcBef.SpcPts != nil {
				spaceBefore = ptToPx(para.PPr.SpcBef.SpcPts.Val / 100)
			}
		}

		// Compute space after for this paragraph (in pixels)
		spaceAfter := 0
		if para.PPr != nil && para.PPr.SpcAft != nil {
			if para.PPr.SpcAft.SpcPts != nil {
				spaceAfter = ptToPx(para.PPr.SpcAft.SpcPts.Val / 100)
			}
		}

		// Compute line spacing for this paragraph (in pixels)
		// spcPts: absolute spacing in hundredths of a point â†’ convert to px
		// spcPct: percentage of "normal" line spacing (100000 = 100%)
		lineSpacing := 0
		lineSpacingPct := 0 // store percentage separately for spcPct
		if para.PPr != nil && para.PPr.LnSpc != nil {
			if para.PPr.LnSpc.SpcPts != nil {
				lineSpacing = ptToPx(para.PPr.LnSpc.SpcPts.Val / 100)
			} else if para.PPr.LnSpc.SpcPct != nil {
				lineSpacingPct = para.PPr.LnSpc.SpcPct.Val // e.g., 150000 = 150%
			}
		}

		// Compute left margin from marL (EMU to pixels)
		marginLeft := 0
		if para.PPr != nil && para.PPr.MarL > 0 {
			marginLeft = emuToPx(para.PPr.MarL)
		}

		// Compute first line indent (EMU to pixels)
		indent := 0
		if para.PPr != nil && para.PPr.Indent != 0 {
			indent = emuToPx(para.PPr.Indent)
		}

		// Bullet processing
		bulletType := "none"
		bulletChar := ""
		bulletAutoNumType := ""
		bulletStartAt := 1

		if para.PPr != nil {
			if para.PPr.BuNone != nil {
				bulletType = "none"
			} else if para.PPr.BuChar != nil {
				bulletType = "char"
				bulletChar = para.PPr.BuChar.Char
			} else if para.PPr.BuAutoNum != nil {
				bulletType = "autoNum"
				bulletAutoNumType = para.PPr.BuAutoNum.Type
			}
		}

		isFirstSegment := true

		// Helper: append one OOXML run to segment runs/paraText and update shape-level font/color/bold.
		addRun := func(run *OOXMLRun, segRuns *[]RenderTextRun, segText *[]string) {
			if run == nil {
				return
			}
			t := run.T
			trimmed := strings.TrimSpace(t)
			if trimmed == "" {
				return
			}
			*segText = append(*segText, trimmed)
			tr := RenderTextRun{Text: t, Bold: run.RPr.B == "1", Italic: run.RPr.I == "1"}
			if run.RPr.U != "" && run.RPr.U != "none" {
				tr.Underline = run.RPr.U
			}
			if run.RPr.Strike != "" && run.RPr.Strike != "noStrike" {
				tr.Strikethrough = run.RPr.Strike
			}
			if run.RPr.Baseline != 0 {
				tr.BaselineShift = run.RPr.Baseline
			}
			if run.RPr.HlinkClick != nil && run.RPr.HlinkClick.RID != "" {
				tr.HyperlinkURL = resolveHyperlinkURL(run.RPr.HlinkClick.RID, rels)
			}
			if run.RPr.Sz > 0 {
				tr.FontSize = ptToPx(run.RPr.Sz / 100)
			} else if para.PPr != nil && para.PPr.DefRPr != nil && para.PPr.DefRPr.Sz > 0 {
				tr.FontSize = ptToPx(para.PPr.DefRPr.Sz / 100)
			}
			if tr.FontSize > 0 && fontSize == 0 {
				fontSize = tr.FontSize
			}
			if run.RPr.Latin != nil && run.RPr.Latin.Typeface != "" {
				tf := run.RPr.Latin.Typeface
				if !strings.HasPrefix(tf, "+") {
					tr.FontFamily = tf
					if fontFamily == "" {
						fontFamily = tf
					}
				}
			}
			if run.RPr.SolidFill != nil {
				c := resolveSolidFill(run.RPr.SolidFill, themeColors)
				if c != "" {
					tr.Color = c
					if textColor == "" {
						textColor = c
					}
				}
			}
			if run.RPr.GradFill != nil {
				tr.GradientDef = buildSVGGradient(run.RPr.GradFill, 0, 0, 0, 0, 0, themeColors)
			}
			if run.RPr.PattFill != nil {
				tr.PatternFill = extractPatternFill(run.RPr.PattFill, themeColors)
			}
			if run.RPr.Ln != nil && run.RPr.Ln.NoFill == nil && run.RPr.Ln.W > 0 {
				tr.StrokeWidth = float64(run.RPr.Ln.W) / 12700.0
				if run.RPr.Ln.SolidFill != nil {
					tr.StrokeColor = resolveSolidFill(run.RPr.Ln.SolidFill, themeColors)
				}
			}
			if run.RPr.EffectLst != nil {
				tr.Shadow = extractShapeShadow(run.RPr.EffectLst, themeColors)
				if run.RPr.EffectLst.Reflection != nil {
					tr.HasReflection = true
				}
			}
			if run.RPr.B == "1" {
				bold = true
			}
			*segRuns = append(*segRuns, tr)
		}

		flushSegment := func(segRuns []RenderTextRun, segText []string) {
			if len(segRuns) > 0 {
				rp := RenderParagraph{
					Runs:           segRuns,
					Alignment:      paraAlign,
					Level:          paraLevel,
					SpaceBefore:    spaceBefore,
					SpaceAfter:     spaceAfter,
					LineSpacing:    lineSpacing,
					LineSpacingPct: lineSpacingPct,
					MarginLeft:     marginLeft,
					Indent:         indent,
					BulletType:     "none",
				}
				if isFirstSegment {
					rp.BulletType = bulletType
					rp.BulletChar = bulletChar
					rp.BulletAutoNumType = bulletAutoNumType
					rp.BulletStartAt = bulletStartAt
					isFirstSegment = false
				}
				paragraphs = append(paragraphs, rp)
				parts = append(parts, strings.Join(segText, " "))
			}
		}

		if len(para.Content) > 0 {
			// Content preserves <a:r> and <a:br/> order; each Br starts a new line (new RenderParagraph).
			var segRuns []RenderTextRun
			var segText []string
			for _, item := range para.Content {
				if item.Br {
					flushSegment(segRuns, segText)
					segRuns = nil
					segText = nil
					continue
				}
				addRun(item.Run, &segRuns, &segText)
			}
			flushSegment(segRuns, segText)
		} else {
			var paraText []string
			var runs []RenderTextRun
			for i := range para.Runs {
				addRun(&para.Runs[i], &runs, &paraText)
			}
			if len(runs) > 0 {
				rp := RenderParagraph{
					Runs:              runs,
					Alignment:         paraAlign,
					Level:             paraLevel,
					SpaceBefore:       spaceBefore,
					SpaceAfter:        spaceAfter,
					LineSpacing:       lineSpacing,
					LineSpacingPct:    lineSpacingPct,
					MarginLeft:        marginLeft,
					Indent:            indent,
					BulletType:        bulletType,
					BulletChar:        bulletChar,
					BulletAutoNumType: bulletAutoNumType,
					BulletStartAt:     bulletStartAt,
				}
				paragraphs = append(paragraphs, rp)
				parts = append(parts, strings.Join(paraText, " "))
			} else {
				paragraphs = append(paragraphs, RenderParagraph{
					IsEmpty:           true,
					Level:             paraLevel,
					SpaceBefore:       spaceBefore,
					SpaceAfter:        spaceAfter,
					LineSpacing:       lineSpacing,
					LineSpacingPct:    lineSpacingPct,
					MarginLeft:        marginLeft, // Even empty para has margin
					Indent:            indent,
					BulletType:        bulletType, // Empty paragraphs can still have bullets technically
					BulletChar:        bulletChar,
					BulletAutoNumType: bulletAutoNumType,
					BulletStartAt:     bulletStartAt,
				})
			}
		}
	}

	text := strings.Join(parts, " ")
	text = strings.TrimSpace(text)

	// Track seen text for callers that need it; do not suppress rendering of
	// identical template text at different positions (e.g. "Placeholder" in Green Accent 2,
	// Blue Accent 1, Orange Accent 5) â€” each shape should render its text.
	if text != "" {
		seenTexts[text] = true
	}

	return text, fontSize, bold, alignment, paragraphs, fontFamily, textColor
}

// ======================== Preset Geometry (arrows, etc.) ========================

// extractPrstAdjs extracts all adjustment values from a preset geometry's avLst.
// Returns a map of name â†’ value. Empty map if no adjustments.
func extractPrstAdjs(pg *OOXMLPrstGeom) map[string]int {
	result := make(map[string]int)
	if pg == nil {
		return result
	}
	for _, gd := range pg.AvLst.Gd {
		if strings.HasPrefix(gd.Fmla, "val ") {
			val, err := strconv.Atoi(strings.TrimPrefix(gd.Fmla, "val "))
			if err == nil {
				result[gd.Name] = val
			}
		}
	}
	return result
}

// isFilledPreset returns true if the given preset geometry name should be
// rendered as a filled SVG path rather than a basic rect/ellipse/textbox.
func isFilledPreset(prst string) bool {
	switch prst {
	// Arrows
	case "downArrow", "upArrow", "leftArrow", "rightArrow",
		"leftRightArrow", "upDownArrow",
		"stripedRightArrow", "notchedRightArrow", "bentUpArrow",
		"uturnArrow", "chevron", "homePlate",
		// Polygons
		"pentagon", "hexagon", "octagon", "decagon", "heptagon", "dodecagon",
		// Stars
		"star4", "star5", "star6", "star8", "star10", "star12",
		// Basic shapes
		"diamond", "triangle", "rtTriangle",
		"parallelogram", "trapezoid",
		"cross", "plus", "heart",
		// 3D shapes
		"cube", "can", "bevel",
		// Crescents, arcs and pie
		"moon", "blockArc", "donut", "pie", "arc", "pieWedge",
		// Symbols
		"noSmoking", "smileyFace", "circularArrow",
		// Callouts
		"rightArrowCallout", "downArrowCallout",
		"leftArrowCallout", "upArrowCallout",
		// Other common shapes
		"frame", "gear6", "gear9",
		"ribbon", "ribbon2",
		"cloud", "lightningBolt", "sun", "round2DiagRect":
		return true
	}
	return false
}

// isCompoundPreset returns true for preset shapes that need multiple SVG elements
// (e.g. smiley face = circle + eyes + mouth) and can't be represented as a single path.
func isCompoundPreset(prst string) bool {
	switch prst {
	case "smileyFace":
		return true
	}
	return false
}

// renderCompoundPreset generates a multi-element SVG fragment for compound preset shapes.
func renderCompoundPreset(prst string, shapeX, shapeY, shapeW, shapeH int64, fill, stroke string, adjs map[string]int, ctx *slideContext) string {
	px := float64(shapeX) / 914400.0 * 96.0
	py := float64(shapeY) / 914400.0 * 96.0
	pw := float64(shapeW) / 914400.0 * 96.0
	ph := float64(shapeH) / 914400.0 * 96.0

	if fill == "" {
		fill = "#4472C4"
	}
	if stroke == "" {
		stroke = darkenColor(fill)
	}

	switch prst {
	case "smileyFace":
		return renderSmileyFace(px, py, pw, ph, fill, stroke, adjs)
	}
	return ""
}

// renderSmileyFace generates SVG elements for a smiley face: circle + eyes + smile arc.
// adj controls the mouth curvature: positive = smile (default 4653), negative = frown.
func renderSmileyFace(x, y, w, h float64, fill, stroke string, adjs map[string]int) string {
	cx := x + w/2
	cy := y + h/2
	rx := w / 2
	ry := h / 2

	// adj controls mouth curvature: default 4653 = smile, negative = frown
	adjVal := 4653.0
	if v, ok := adjs["adj"]; ok {
		adjVal = float64(v)
	}

	var buf strings.Builder

	// Face circle (filled)
	buf.WriteString(fmt.Sprintf(`  <ellipse cx="%.1f" cy="%.1f" rx="%.1f" ry="%.1f" fill="%s" stroke="%s" stroke-width="1.33"/>
`, cx, cy, rx, ry, fill, stroke))

	// Left eye (small filled ellipse in stroke color)
	eyeRx := rx * 0.08
	eyeRy := ry * 0.10
	leftEyeX := cx - rx*0.30
	leftEyeY := cy - ry*0.25
	buf.WriteString(fmt.Sprintf(`  <ellipse cx="%.1f" cy="%.1f" rx="%.1f" ry="%.1f" fill="%s"/>
`, leftEyeX, leftEyeY, eyeRx, eyeRy, stroke))

	// Right eye
	rightEyeX := cx + rx*0.30
	rightEyeY := cy - ry*0.25
	buf.WriteString(fmt.Sprintf(`  <ellipse cx="%.1f" cy="%.1f" rx="%.1f" ry="%.1f" fill="%s"/>
`, rightEyeX, rightEyeY, eyeRx, eyeRy, stroke))

	// Mouth: positive adj = smile (curve down), negative adj = frown (curve up)
	smileY := cy + ry*0.25
	smileLeftX := cx - rx*0.40
	smileRightX := cx + rx*0.40
	// Scale depth by adj: 4653 â†’ standard smile, -4653 â†’ standard frown
	smileDepth := ry * 0.30 * adjVal / 4653.0
	buf.WriteString(fmt.Sprintf(`  <path d="M%.1f %.1f Q%.1f %.1f %.1f %.1f" fill="none" stroke="%s" stroke-width="1.33" stroke-linecap="round"/>
`, smileLeftX, smileY, cx, smileY+smileDepth, smileRightX, smileY, stroke))

	return buf.String()
}

// isLinePreset returns true for preset shapes that are typically line-only (no fill).
func isLinePreset(prst string) bool {
	switch prst {
	case "rightBrace", "leftBrace",
		"rightBracket", "leftBracket",
		"bracePair", "bracketPair",
		"mathPlus", "mathMinus", "mathMultiply", "mathDivide", "mathEqual",
		"arc":
		return true
	}
	return false
}

// prstGeomToSVGPath generates an SVG path d-attribute for known preset geometries.
// The path uses absolute pixel coordinates derived from the shape's EMU position/size.
// adjs is a map of adjustment name â†’ value from avLst.
func prstGeomToSVGPath(prst string, shapeX, shapeY, shapeW, shapeH int64, adjs map[string]int) string {
	// Convert EMU to pixels
	x := float64(shapeX) / 914400.0 * 96.0
	y := float64(shapeY) / 914400.0 * 96.0
	w := float64(shapeW) / 914400.0 * 96.0
	h := float64(shapeH) / 914400.0 * 96.0

	switch prst {
	case "downArrow":
		// Default adjustment: shaft width = 50% of total width, arrowhead height = 50% of shaft width
		shaftW := w * 0.5
		headH := shaftW // arrowhead height equals shaft width (OOXML default adj)
		if headH > h*0.5 {
			headH = h * 0.5
		}
		shaftH := h - headH
		sl := x + (w-shaftW)/2 // shaft left
		sr := x + (w+shaftW)/2 // shaft right
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			sl, y, // top-left of shaft
			sr, y, // top-right of shaft
			sr, y+shaftH, // right side at arrowhead junction
			x+w, y+shaftH, // right wing tip
			x+w/2, y+h, // bottom center point
			x, y+shaftH, // left wing tip
			sl, y+shaftH, // left side at arrowhead junction
		)
	case "upArrow":
		shaftW := w * 0.5
		headH := shaftW
		if headH > h*0.5 {
			headH = h * 0.5
		}
		sl := x + (w-shaftW)/2
		sr := x + (w+shaftW)/2
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x+w/2, y, // top center point
			x+w, y+headH, // right wing tip
			sr, y+headH, // right side at junction
			sr, y+h, // bottom-right of shaft
			sl, y+h, // bottom-left of shaft
			sl, y+headH, // left side at junction
			x, y+headH, // left wing tip
		)
	case "leftArrow":
		shaftH := h * 0.5
		headW := shaftH
		if headW > w*0.5 {
			headW = w * 0.5
		}
		st := y + (h-shaftH)/2
		sb := y + (h+shaftH)/2
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x, y+h/2, // left center point
			x+headW, y, // top wing tip
			x+headW, st, // top side at junction
			x+w, st, // top-right of shaft
			x+w, sb, // bottom-right of shaft
			x+headW, sb, // bottom side at junction
			x+headW, y+h, // bottom wing tip
		)
	case "rightArrow":
		shaftH := h * 0.5
		headW := shaftH
		if headW > w*0.5 {
			headW = w * 0.5
		}
		st := y + (h-shaftH)/2
		sb := y + (h+shaftH)/2
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x, st, // top-left of shaft
			x+w-headW, st, // top side at junction
			x+w-headW, y, // top wing tip
			x+w, y+h/2, // right center point
			x+w-headW, y+h, // bottom wing tip
			x+w-headW, sb, // bottom side at junction
			x, sb, // bottom-left of shaft
		)
	case "diamond":
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x+w/2, y, x+w, y+h/2, x+w/2, y+h, x, y+h/2)
	case "triangle":
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x+w/2, y, x+w, y+h, x, y+h)
	case "rtTriangle":
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x, y, x+w, y+h, x, y+h)
	case "pentagon":
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x+w*0.5, y, x+w, y+h*0.38, x+w*0.81, y+h, x+w*0.19, y+h, x, y+h*0.38)
	case "hexagon":
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x+w*0.25, y, x+w*0.75, y, x+w, y+h/2, x+w*0.75, y+h, x+w*0.25, y+h, x, y+h/2)
	case "cross", "plus":
		// adj defines half-arm-width as fraction of short side (in 100000ths); default is 25000
		adjVal := 25000.0
		if v, ok := adjs["adj"]; ok && v > 0 {
			adjVal = float64(v)
		}
		ss := math.Min(w, h)
		dx := ss * adjVal / 100000.0 // half-width of vertical arm
		dy := ss * adjVal / 100000.0 // half-height of horizontal arm
		cx := x + w/2
		cy := y + h/2
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			cx-dx, y, cx+dx, y, // top of vertical bar
			cx+dx, cy-dy, x+w, cy-dy, // inner top-right â†’ right of horiz bar top
			x+w, cy+dy, cx+dx, cy+dy, // right of horiz bar bottom â†’ inner bottom-right
			cx+dx, y+h, cx-dx, y+h, // bottom of vertical bar
			cx-dx, cy+dy, x, cy+dy, // inner bottom-left â†’ left of horiz bar bottom
			x, cy-dy, cx-dx, cy-dy, // left of horiz bar top â†’ inner top-left
		)
	case "chevron":
		notch := w * 0.25
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x, y, x+w-notch, y, x+w, y+h/2, x+w-notch, y+h, x, y+h)
	case "homePlate":
		notch := w * 0.20
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x, y, x+w-notch, y, x+w, y+h/2, x+w-notch, y+h, x, y+h)

	case "cube":
		// 3D cube: front face + top parallelogram + right side parallelogram
		// Default adj = 25000 â†’ offset is 25% of min(w,h)
		off := math.Min(w, h) * 0.25
		// Front face
		front := fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x, y+off, x+w-off, y+off, x+w-off, y+h, x, y+h)
		// Top face
		top := fmt.Sprintf(" M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x, y+off, x+off, y, x+w, y, x+w-off, y+off)
		// Right face
		right := fmt.Sprintf(" M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x+w-off, y+off, x+w, y, x+w, y+h-off, x+w-off, y+h)
		return front + top + right

	case "moon":
		// Crescent moon shape using SVG arc commands
		// The moon is drawn as an outer ellipse arc and an inner arc cut-out
		moonCx := x + w/2
		rx := w / 2
		ry := h / 2
		// Inner arc radius (controls crescent thickness). Default adj=50000 â†’ inner rx is ~37.5% of outer
		innerRx := rx * 0.375
		// Draw outer right arc (from top to bottom), then inner arc (bottom to top)
		return fmt.Sprintf("M%.1f %.1f A%.1f %.1f 0 1 1 %.1f %.1f A%.1f %.1f 0 0 0 %.1f %.1f Z",
			moonCx, y, // start: top center
			rx, ry, moonCx, y+h, // outer arc: top â†’ bottom (large arc, sweep clockwise)
			innerRx, ry, moonCx, y, // inner arc: bottom â†’ top (small arc, sweep counter-clockwise)
		)

	case "decagon":
		// Regular 10-sided polygon
		return regularPolygonPath(x, y, w, h, 10)
	case "heptagon":
		return regularPolygonPath(x, y, w, h, 7)
	case "octagon":
		return regularPolygonPath(x, y, w, h, 8)
	case "dodecagon":
		return regularPolygonPath(x, y, w, h, 12)

	case "star4":
		return starPath(x, y, w, h, 4, 0.38)
	case "star5":
		return starPath(x, y, w, h, 5, 0.38)
	case "star6":
		return starPath(x, y, w, h, 6, 0.38)
	case "star8":
		return starPath(x, y, w, h, 8, 0.38)
	case "star10":
		return starPath(x, y, w, h, 10, 0.38)
	case "star12":
		return starPath(x, y, w, h, 12, 0.38)

	case "trapezoid":
		inset := w * 0.25
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x+inset, y, x+w-inset, y, x+w, y+h, x, y+h)
	case "parallelogram":
		offset := w * 0.25
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x+offset, y, x+w, y, x+w-offset, y+h, x, y+h)
	case "heart":
		// Heart shape as a single continuous path (no center seam)
		// Start at bottom tip, curve up left side, across top notch, down right side
		return fmt.Sprintf(
			"M%.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f Z",
			x+w/2, y+h, // bottom tip
			x+w/2, y+h*0.85, x, y+h*0.65, x, y+h*0.35, // left side going up
			x, y, x+w/2, y, x+w/2, y+h*0.3, // left lobe to top notch
			x+w/2, y, x+w, y, x+w, y+h*0.35, // right lobe from top notch
			x+w, y+h*0.65, x+w/2, y+h*0.85, x+w/2, y+h, // right side going down
		)

	case "noSmoking":
		// "No" symbol: circle with diagonal ban line across it
		// Outer circle + inner circle (donut ring) + diagonal bar
		cx := x + w/2
		cy := y + h/2
		rx := w / 2
		ry := h / 2
		// Ring thickness: default adj=18750 â†’ 18.75% of min dimension
		ringAdj := 18750.0
		if v, ok := adjs["adj"]; ok && v > 0 {
			ringAdj = float64(v)
		}
		thickness := math.Min(w, h) * ringAdj / 100000.0
		irx := rx - thickness
		iry := ry - thickness
		if irx < 0 {
			irx = 0
		}
		if iry < 0 {
			iry = 0
		}
		// Outer circle (two arcs)
		outer := fmt.Sprintf("M%.1f %.1f A%.1f %.1f 0 1 1 %.1f %.1f A%.1f %.1f 0 1 1 %.1f %.1f Z",
			cx-rx, cy, rx, ry, cx+rx, cy, rx, ry, cx-rx, cy)
		// Inner circle (counter-clockwise to cut out)
		inner := fmt.Sprintf(" M%.1f %.1f A%.1f %.1f 0 1 0 %.1f %.1f A%.1f %.1f 0 1 0 %.1f %.1f Z",
			cx-irx, cy, irx, iry, cx+irx, cy, irx, iry, cx-irx, cy)
		// Diagonal bar (from top-right to bottom-left area of the ring)
		angle := math.Pi / 4 // 45 degrees
		barOx1 := cx + rx*math.Cos(angle)
		barOy1 := cy - ry*math.Sin(angle)
		barOx2 := cx - rx*math.Cos(angle)
		barOy2 := cy + ry*math.Sin(angle)
		barIx1 := cx + irx*math.Cos(angle)
		barIy1 := cy - iry*math.Sin(angle)
		barIx2 := cx - irx*math.Cos(angle)
		barIy2 := cy + iry*math.Sin(angle)
		bar := fmt.Sprintf(" M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			barOx1, barOy1, barOx2, barOy2, barIx2, barIy2, barIx1, barIy1)
		return outer + inner + bar

	case "arc":
		// Open arc shape from adj1 to adj2
		adj1 := 0.0
		adj2 := 16200000.0
		if v, ok := adjs["adj1"]; ok {
			adj1 = float64(v)
		}
		if v, ok := adjs["adj2"]; ok {
			adj2 = float64(v)
		}
		cx := x + w/2
		cy := y + h/2
		rx := w / 2
		ry := h / 2
		startRad := adj1 / 60000.0 * math.Pi / 180.0
		endRad := adj2 / 60000.0 * math.Pi / 180.0
		sx := cx + rx*math.Cos(startRad)
		sy := cy + ry*math.Sin(startRad)
		ex := cx + rx*math.Cos(endRad)
		ey := cy + ry*math.Sin(endRad)
		sweep := endRad - startRad
		if sweep < 0 {
			sweep += 2 * math.Pi
		}
		largeArc := 0
		if sweep > math.Pi {
			largeArc = 1
		}
		return fmt.Sprintf("M%.1f %.1f A%.1f %.1f 0 %d 1 %.1f %.1f",
			sx, sy, rx, ry, largeArc, ex, ey)

	case "pie":
		// Pie / partial circle shape
		// adj1 = start angle (in 60000ths of a degree), default 0
		// adj2 = end angle (in 60000ths of a degree), default 16200000 (270Â°)
		adj1 := 0.0
		adj2 := 16200000.0
		if v, ok := adjs["adj1"]; ok {
			adj1 = float64(v)
		}
		if v, ok := adjs["adj2"]; ok {
			adj2 = float64(v)
		}
		cx := x + w/2
		cy := y + h/2
		rx := w / 2
		ry := h / 2
		// Convert OOXML angles to radians (OOXML: 0Â° = right/3-o'clock, clockwise)
		startRad := adj1 / 60000.0 * math.Pi / 180.0
		endRad := adj2 / 60000.0 * math.Pi / 180.0
		// Compute arc endpoints
		sx := cx + rx*math.Cos(startRad)
		sy := cy + ry*math.Sin(startRad)
		ex := cx + rx*math.Cos(endRad)
		ey := cy + ry*math.Sin(endRad)
		// Determine if large arc flag is needed
		sweep := endRad - startRad
		if sweep < 0 {
			sweep += 2 * math.Pi
		}
		largeArc := 0
		if sweep > math.Pi {
			largeArc = 1
		}
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f A%.1f %.1f 0 %d 1 %.1f %.1f Z",
			cx, cy, sx, sy, rx, ry, largeArc, ex, ey)

	case "pieWedge":
		// pieWedge is a 90-degree quadrant with its center positioned at the bottom-right of the bounding box
		// It starts at the bottom-left corner and arcs up to the top-right corner.
		return fmt.Sprintf("M%.1f %.1f A%.1f %.1f 0 0 1 %.1f %.1f L%.1f %.1f Z",
			x, y+h, w, h, x+w, y, x+w, y+h)

	case "circularArrow":
		// A 180-degree semi-circular arrow filling the top half of the bounding box.
		th := w * 0.1
		midY := y + h/2

		arrowH := h * 0.14

		outerSx := x + w*0.06
		innerSx := outerSx + th

		arrowTip := x + w - w*0.02
		arrowBarbOuter := arrowTip + w*0.07
		arrowBarbInner := arrowTip - w*0.15
		arcBaseY := midY - arrowH

		outerEx := arrowBarbOuter - w*0.04
		innerEx := arrowBarbInner + w*0.04

		return fmt.Sprintf("M%.1f %.1f A%.1f %.1f 0 0 1 %.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f A%.1f %.1f 0 0 0 %.1f %.1f Z",
			outerSx, midY,
			w*0.45, h*0.45, outerEx, arcBaseY,
			arrowBarbOuter, arcBaseY,
			arrowTip, midY,
			arrowBarbInner, arcBaseY,
			innerEx, arcBaseY,
			w*0.35, h*0.35, innerSx, midY,
		)

	case "rightBrace":
		// Right curly brace shape (line-only, drawn as a path with no fill)
		// Approximation using cubic bezier curves
		midY := y + h/2
		tipX := x + w
		return fmt.Sprintf(
			"M%.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f",
			x, y, // start top
			x+w*0.5, y, x+w*0.5, y, x+w*0.5, y+h*0.25, // curve to mid-top
			x+w*0.5, midY-h*0.05, tipX, midY-h*0.05, tipX, midY, // curve to tip
			tipX, midY+h*0.05, x+w*0.5, midY+h*0.05, x+w*0.5, y+h*0.75, // curve from tip
			x+w*0.5, y+h, x+w*0.5, y+h, x, y+h, // curve to bottom
		)

	case "leftBrace":
		// Left curly brace (mirror of right brace)
		midY := y + h/2
		return fmt.Sprintf(
			"M%.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f C%.1f %.1f %.1f %.1f %.1f %.1f",
			x+w, y,
			x+w*0.5, y, x+w*0.5, y, x+w*0.5, y+h*0.25,
			x+w*0.5, midY-h*0.05, x, midY-h*0.05, x, midY,
			x, midY+h*0.05, x+w*0.5, midY+h*0.05, x+w*0.5, y+h*0.75,
			x+w*0.5, y+h, x+w*0.5, y+h, x+w, y+h,
		)

	case "cloud":
		// Cloud shape approximation using cubic bezier curves
		// A cloud is a rounded blobby shape with multiple bumps
		return fmt.Sprintf(
			"M%.1f %.1f "+
				"C%.1f %.1f %.1f %.1f %.1f %.1f "+
				"C%.1f %.1f %.1f %.1f %.1f %.1f "+
				"C%.1f %.1f %.1f %.1f %.1f %.1f "+
				"C%.1f %.1f %.1f %.1f %.1f %.1f "+
				"C%.1f %.1f %.1f %.1f %.1f %.1f "+
				"C%.1f %.1f %.1f %.1f %.1f %.1f "+
				"Z",
			// Start at bottom-left
			x+w*0.15, y+h*0.85,
			// Bottom curve to right
			x+w*0.10, y+h*1.05, x+w*0.40, y+h*1.05, x+w*0.50, y+h*0.90,
			// Bottom-right bump
			x+w*0.60, y+h*1.05, x+w*0.90, y+h*1.05, x+w*0.90, y+h*0.75,
			// Right side up
			x+w*1.05, y+h*0.65, x+w*1.05, y+h*0.35, x+w*0.85, y+h*0.25,
			// Top-right bump
			x+w*0.88, y+h*0.05, x+w*0.65, y+h*-0.05, x+w*0.50, y+h*0.15,
			// Top-left bump
			x+w*0.38, y+h*-0.05, x+w*0.15, y+h*0.05, x+w*0.12, y+h*0.30,
			// Left side down
			x+w*-0.05, y+h*0.35, x+w*-0.05, y+h*0.70, x+w*0.15, y+h*0.85,
		)

	case "lightningBolt":
		// Lightning bolt shape â€” matches PowerPoint rendering output.
		// 11-point path: single top point â†’ zigzag down-right to tip â†’ zigzag back up-left to tip â†’ close.
		// Proportions reverse-engineered from PowerPoint SVG export.
		return fmt.Sprintf(
			"M%.1f %.1f "+
				"%.1f %.1f %.1f %.1f %.1f %.1f %.1f %.1f "+
				"%.1f %.1f "+
				"%.1f %.1f %.1f %.1f %.1f %.1f %.1f %.1f "+
				"%.1f %.1fZ",
			x+w*0.40, y, // 1. top peak
			x+w*0.60, y+h*0.285, // 2. down-right
			x+w*0.52, y+h*0.318, // 3. step left-down
			x+w*0.78, y+h*0.558, // 4. down-right
			x+w*0.69, y+h*0.599, // 5. step left-down
			x+w, y+h, // 6. bottom-right tip
			x+w*0.47, y+h*0.692, // 7. return up-left
			x+w*0.57, y+h*0.650, // 8. step right-up
			x+w*0.24, y+h*0.452, // 9. return up-left
			x+w*0.36, y+h*0.391, // 10. step right-up
			x, y+h*0.184, // 11. top-left tip
		)

	case "round2DiagRect":
		// Round diagonal corners: Top-Right and Bottom-Left
		// adj1: Top-Right radius
		// adj2: Bottom-Left radius
		// XML example: adj1=0, adj2=16670

		r1Val := 16667 // default
		if v, ok := adjs["adj1"]; ok {
			r1Val = v
		}
		r2Val := 16667 // default
		if v, ok := adjs["adj2"]; ok {
			r2Val = v
		}

		// If one is 0 and the other is not, enforce symmetry (User expectation)
		if r1Val == 0 && r2Val > 0 {
			r1Val = r2Val
		} else if r2Val == 0 && r1Val > 0 {
			r2Val = r1Val
		}

		minDim := w
		if h < minDim {
			minDim = h
		}

		r1 := float64(r1Val) / 100000.0 * minDim
		r2 := float64(r2Val) / 100000.0 * minDim

		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f A%.1f %.1f 0 0 1 %.1f %.1f L%.1f %.1f L%.1f %.1f A%.1f %.1f 0 0 1 %.1f %.1f L%.1f %.1f Z",
			x, y, // Move to Top-Left (sharp)
			x+w-r1, y, // Line to start of Top-Right arc
			r1, r1, x+w, y+r1, // Arc to end of Top-Right arc
			x+w, y+h, // Line to Bottom-Right (sharp)
			x+r2, y+h, // Line to start of Bottom-Left arc
			r2, r2, x, y+h-r2, // Arc to end of Bottom-Left arc
			x, y, // Line back to Top-Left
		)

	default:
		// Fallback: render as a simple rectangle path
		return fmt.Sprintf("M%.1f %.1f L%.1f %.1f L%.1f %.1f L%.1f %.1f Z",
			x, y, x+w, y, x+w, y+h, x, y+h)
	}
}

// regularPolygonPath generates an SVG path for a regular n-sided polygon inscribed in the bounding box.
func regularPolygonPath(x, y, w, h float64, n int) string {
	cx := x + w/2
	cy := y + h/2
	rx := w / 2
	ry := h / 2
	var parts []string
	for i := 0; i < n; i++ {
		angle := math.Pi + 2*math.Pi*float64(i)/float64(n) // start from left, counterclockwise (matches PowerPoint)
		px := cx + rx*math.Cos(angle)
		py := cy + ry*math.Sin(angle)
		if i == 0 {
			parts = append(parts, fmt.Sprintf("M%.1f %.1f", px, py))
		} else {
			parts = append(parts, fmt.Sprintf("L%.1f %.1f", px, py))
		}
	}
	parts = append(parts, "Z")
	return strings.Join(parts, " ")
}

// starPath generates an SVG path for a star with n outer points.
// innerRatio controls the inner radius as a fraction of the outer radius.
func starPath(x, y, w, h float64, n int, innerRatio float64) string {
	cx := x + w/2
	cy := y + h/2
	outerRx := w / 2
	outerRy := h / 2
	innerRx := outerRx * innerRatio
	innerRy := outerRy * innerRatio
	totalPts := n * 2
	var parts []string
	for i := 0; i < totalPts; i++ {
		angle := -math.Pi/2 + 2*math.Pi*float64(i)/float64(totalPts)
		var px, py float64
		if i%2 == 0 {
			// Outer point
			px = cx + outerRx*math.Cos(angle)
			py = cy + outerRy*math.Sin(angle)
		} else {
			// Inner point
			px = cx + innerRx*math.Cos(angle)
			py = cy + innerRy*math.Sin(angle)
		}
		if i == 0 {
			parts = append(parts, fmt.Sprintf("M%.1f %.1f", px, py))
		} else {
			parts = append(parts, fmt.Sprintf("L%.1f %.1f", px, py))
		}
	}
	parts = append(parts, "Z")
	return strings.Join(parts, " ")
}

// ======================== Custom Geometry ========================

// custGeomToSVGPath converts OOXML custom geometry path commands to an SVG path d attribute.
// Coordinates are converted from path-local space to absolute pixel space so that no
// SVG scale transform is needed (which would distort stroke widths).
func custGeomToSVGPath(geom *OOXMLCustGeom, shapeX, shapeY, shapeW, shapeH int64) string {
	if geom == nil || len(geom.PathLst.Paths) == 0 {
		return ""
	}
	pathRaw := geom.PathLst.Paths[0]
	pathW := pathRaw.W
	pathH := pathRaw.H
	if pathW == 0 {
		pathW = shapeW
	}
	if pathH == 0 {
		pathH = shapeH
	}

	// Parse inner XML tokens to preserve command order
	type pathCmd struct {
		kind string // "M", "L", "C", "Z"
		pts  []OOXMLPt
	}
	var commands []pathCmd

	wrapped := "<root>" + pathRaw.Inner + "</root>"
	decoder := xml.NewDecoder(strings.NewReader(wrapped))
	for {
		tok, err := decoder.Token()
		if err != nil {
			break
		}
		se, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}
		switch se.Name.Local {
		case "moveTo":
			var mt struct {
				Pt OOXMLPt `xml:"pt"`
			}
			decoder.DecodeElement(&mt, &se)
			commands = append(commands, pathCmd{kind: "M", pts: []OOXMLPt{mt.Pt}})
		case "lnTo":
			var lt struct {
				Pt OOXMLPt `xml:"pt"`
			}
			decoder.DecodeElement(&lt, &se)
			commands = append(commands, pathCmd{kind: "L", pts: []OOXMLPt{lt.Pt}})
		case "cubicBezTo":
			var cb struct {
				Pts []OOXMLPt `xml:"pt"`
			}
			decoder.DecodeElement(&cb, &se)
			commands = append(commands, pathCmd{kind: "C", pts: cb.Pts})
		case "close":
			decoder.DecodeElement(&struct{}{}, &se)
			commands = append(commands, pathCmd{kind: "Z"})
		}
	}

	if len(commands) == 0 {
		return ""
	}

	// Convert path coordinates to absolute pixel coordinates:
	// pixelX = shapeXpx + (pathCoordX / pathW) * shapeWpx
	sxPx := float64(shapeX) / emuPerPx
	syPx := float64(shapeY) / emuPerPx
	swPx := float64(shapeW) / emuPerPx
	shPx := float64(shapeH) / emuPerPx
	toX := func(px int64) float64 { return sxPx + float64(px)/float64(pathW)*swPx }
	toY := func(py int64) float64 { return syPx + float64(py)/float64(pathH)*shPx }

	var d strings.Builder
	for _, cmd := range commands {
		switch cmd.kind {
		case "M":
			if len(cmd.pts) >= 1 {
				d.WriteString(fmt.Sprintf("M%.3f %.3f", toX(cmd.pts[0].X), toY(cmd.pts[0].Y)))
			}
		case "L":
			if len(cmd.pts) >= 1 {
				d.WriteString(fmt.Sprintf("L%.3f %.3f", toX(cmd.pts[0].X), toY(cmd.pts[0].Y)))
			}
		case "C":
			if len(cmd.pts) >= 3 {
				d.WriteString(fmt.Sprintf("C%.3f %.3f %.3f %.3f %.3f %.3f",
					toX(cmd.pts[0].X), toY(cmd.pts[0].Y),
					toX(cmd.pts[1].X), toY(cmd.pts[1].Y),
					toX(cmd.pts[2].X), toY(cmd.pts[2].Y)))
			}
		case "Z":
			d.WriteString("Z")
		}
	}
	return d.String()
}

// ======================== SVG Building ========================

func buildSVG(shapes []RenderShape, connectors []RenderConnector, cfg *RenderConfig, bg *SlideBackground) string {
	var svg strings.Builder

	// Build background fill for the slide
	bgDefs := ""
	bgFill := "#FFFFFF"
	if bg != nil {
		if bg.GradientDef != nil {
			bgDefs = renderSVGGradientDef(bg.GradientDef)
			bgFill = fmt.Sprintf("url(#%s)", bg.GradientDef.ID)
		} else if bg.SolidFill != "" {
			bgFill = bg.SolidFill
		}
	}

	// Pre-render connectors to collect gradient defs (must go in <defs> block)
	var connectorDefs strings.Builder
	var connectorElements []string
	for _, c := range connectors {
		defs, element := renderConnectorSVG(c)
		if defs != "" {
			connectorDefs.WriteString(defs)
		}
		if element != "" {
			connectorElements = append(connectorElements, element)
		}
	}

	svg.WriteString(fmt.Sprintf(`<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" viewBox="0 0 %d %d" width="%d" height="%d">
  <defs>
    <style>text { font-family: Calibri, Arial, sans-serif; }</style>
    <marker id="arrowhead" markerWidth="10" markerHeight="7" refX="9" refY="3.5" orient="auto">
      <polygon points="0 0, 10 3.5, 0 7" fill="#4472C4"/>
    </marker>
    <linearGradient id="stickyGradient" x1="0%%" y1="0%%" x2="0%%" y2="100%%">
      <stop offset="0%%" style="stop-color:#FFFF99"/>
      <stop offset="100%%" style="stop-color:#FFFF89"/>
    </linearGradient>
%s%s  </defs>
  <rect width="%d" height="%d" fill="%s"/>

`, svgWidth, svgHeight, svgWidth, svgHeight, bgDefs, connectorDefs.String(), svgWidth, svgHeight, bgFill))

	for _, s := range shapes {
		svg.WriteString(renderShapeSVG(s, cfg))
	}
	for _, el := range connectorElements {
		svg.WriteString(el)
	}

	svg.WriteString("</svg>")
	return svg.String()
}

// ======================== Shape Rendering ========================

// buildShapeTransform returns an SVG transform string for rotation and flips.
// Returns empty string if no transform is needed.
func buildShapeTransform(s RenderShape, x, y, w, h int) string {
	return buildShapeTransformOpts(s, x, y, w, h, true)
}

// buildShapeRotationOnly returns an SVG transform string for rotation only (no flips).
// Use this for text elements: PowerPoint flips shape geometry but keeps text readable.
func buildShapeRotationOnly(s RenderShape, x, y, w, h int) string {
	return buildShapeTransformOpts(s, x, y, w, h, false)
}

func buildShapeTransformOpts(s RenderShape, x, y, w, h int, includeFlips bool) string {
	var transforms []string

	// Center of the shape for rotation/flip
	cx := float64(x) + float64(w)/2
	cy := float64(y) + float64(h)/2

	// Rotation: OOXML stores in 60000ths of a degree, clockwise
	angleDeg := 0.0
	if s.Rotation != 0 {
		angleDeg = float64(s.Rotation) / 60000.0
	}

	// Handle Text Direction rotation (only for text transform, i.e., !includeFlips)
	if !includeFlips {
		if s.TextDirection == "vert" || s.TextDirection == "eaVert" {
			angleDeg += 90
		} else if s.TextDirection == "vert270" {
			angleDeg += 270
		}
	}

	if angleDeg != 0 {
		transforms = append(transforms, fmt.Sprintf("rotate(%.2f,%.1f,%.1f)", angleDeg, cx, cy))
	}

	if includeFlips {
		// Horizontal flip
		if s.FlipH {
			// Flip around the vertical center axis: translate to origin, scale -1 on X, translate back
			transforms = append(transforms, fmt.Sprintf("translate(%.1f,0) scale(-1,1) translate(%.1f,0)", cx, -cx))
		}

		// Vertical flip
		if s.FlipV {
			transforms = append(transforms, fmt.Sprintf("translate(0,%.1f) scale(1,-1) translate(0,%.1f)", cy, -cy))
		}
	}

	if len(transforms) == 0 {
		return ""
	}
	return strings.Join(transforms, " ")
}

func generateShadowFilter(shdw *RenderShadow, filterID string, isPerspective bool) string {
	r, g, b := hexToRGBFractions(shdw.Color)
	rVal := fmt.Sprintf("%.4g %.4g", r, r)
	gVal := fmt.Sprintf("%.4g %.4g", g, g)
	bVal := fmt.Sprintf("%.4g %.4g", b, b)

	blurStdDev := float64(shdw.BlurRad) / emuPerPx / 3.0
	alpha := shdw.Alpha

	var sb strings.Builder
	// Filter region needs to be large enough to catch the shadow
	sb.WriteString(fmt.Sprintf(`<filter id="%s" x="-50%%" y="-50%%" width="200%%" height="200%%" filterUnits="userSpaceOnUse" primitiveUnits="userSpaceOnUse">`, filterID))

	// Color replacement
	sb.WriteString(`<feComponentTransfer in="SourceAlpha">`)
	sb.WriteString(fmt.Sprintf(`<feFuncR type="discrete" tableValues="%s"/>`, rVal))
	sb.WriteString(fmt.Sprintf(`<feFuncG type="discrete" tableValues="%s"/>`, gVal))
	sb.WriteString(fmt.Sprintf(`<feFuncB type="discrete" tableValues="%s"/>`, bVal))
	sb.WriteString(fmt.Sprintf(`<feFuncA type="linear" slope="%.5g" intercept="0"/>`, alpha))
	sb.WriteString(`</feComponentTransfer>`)

	sb.WriteString(fmt.Sprintf(`<feGaussianBlur stdDeviation="%.5g %.5g"/>`, blurStdDev, blurStdDev))

	if !isPerspective {
		dirRad := float64(shdw.Dir) * math.Pi / 10800000.0
		distPx := float64(shdw.Dist) / emuPerPx
		dx := distPx * math.Cos(dirRad)
		dy := distPx * math.Sin(dirRad)
		sb.WriteString(fmt.Sprintf(`<feOffset dx="%.4f" dy="%.4f" result="shadow"/>`, dx, dy))
		sb.WriteString(`<feMerge><feMergeNode in="shadow"/><feMergeNode in="SourceGraphic"/></feMerge>`)
	}

	sb.WriteString(`</filter>`)
	return sb.String()
}

// renderPatternDef emits an SVG pattern from OOXML preset (prst), fgClr and bgClr â€” all from XML.
func renderPatternDef(p *RenderPattern, id string) string {
	fg, bg := p.FgColor, p.BgColor
	if fg == "" {
		fg = "#000000"
	}
	if bg == "" {
		bg = "#FFFFFF"
	}
	// Map OOXML prst (from pattFill) to SVG pattern; unknown presets fall back to pct25-style grid.
	switch p.Preset {
	case "pct5", "pct10":
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><rect width="1" height="1" fill="%s"/></pattern>`, id, bg, fg)
	case "pct20":
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><rect width="1" height="1" fill="%s"/></pattern>`, id, bg, fg)
	case "pct25":
		return fmt.Sprintf(`<pattern id="%s" width="2" height="2" patternUnits="userSpaceOnUse"><rect width="2" height="2" fill="%s"/><rect width="1" height="1" fill="%s"/></pattern>`, id, bg, fg)
	case "pct30", "pct40":
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><rect width="2" height="2" fill="%s"/></pattern>`, id, bg, fg)
	case "pct50":
		return fmt.Sprintf(`<pattern id="%s" width="2" height="2" patternUnits="userSpaceOnUse"><rect width="2" height="2" fill="%s"/><rect width="1" height="1" x="0" y="0" fill="%s"/><rect width="1" height="1" x="1" y="1" fill="%s"/></pattern>`, id, bg, fg, fg)
	case "pct60":
		// 8x8: foreground dominant, staggered bg dots (~40% bg). Checkerboard with extra fg.
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><rect x="0" y="0" width="1" height="1" fill="%s"/><rect x="2" y="2" width="1" height="1" fill="%s"/></pattern>`, id, fg, bg, bg)
	case "pct70":
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><rect x="0" y="0" width="1" height="1" fill="%s"/></pattern>`, id, fg, bg)
	case "pct75":
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><rect x="0" y="0" width="1" height="1" fill="%s"/></pattern>`, id, fg, bg)
	case "pct80":
		// 8x8 dotted: ~80% fg, ~20% bg. Staggered single-pixel bg dots every 4px.
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><rect x="0" y="0" width="1" height="1" fill="%s"/><rect x="4" y="4" width="1" height="1" fill="%s"/></pattern>`, id, fg, bg, bg)
	case "pct90":
		// 8x8 dotted: ~90% fg, ~10% bg. Single bg dot per tile.
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><rect x="0" y="0" width="1" height="1" fill="%s"/></pattern>`, id, fg, bg)
	// ---- Horizontal stripes ----
	case "horz": // Horizontal
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="0" y1="2" x2="4" y2="2" stroke="%s" stroke-width="1"/></pattern>`, id, bg, fg)
	case "narHorz": // Narrow Horizontal â€” 2px spacing, 1px line
		return fmt.Sprintf(`<pattern id="%s" width="2" height="2" patternUnits="userSpaceOnUse"><rect width="2" height="2" fill="%s"/><line x1="0" y1="0" x2="2" y2="0" stroke="%s" stroke-width="1"/></pattern>`, id, bg, fg)
	case "ltHorz": // Light Horizontal â€” thin line, 8px spacing
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="0" y1="4" x2="8" y2="4" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)
	case "dkHorz": // Dark Horizontal â€” thick line, 4px spacing
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="0" y1="1" x2="4" y2="1" stroke="%s" stroke-width="2"/></pattern>`, id, bg, fg)
	case "dashHorz": // Dashed Horizontal
		return fmt.Sprintf(`<pattern id="%s" width="8" height="4" patternUnits="userSpaceOnUse"><rect width="8" height="4" fill="%s"/><line x1="0" y1="2" x2="4" y2="2" stroke="%s" stroke-width="1"/></pattern>`, id, bg, fg)

	// ---- Vertical stripes ----
	case "vert": // Vertical
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="2" y1="0" x2="2" y2="4" stroke="%s" stroke-width="1"/></pattern>`, id, bg, fg)
	case "narVert": // Narrow Vertical â€” 2px spacing, 1px line
		return fmt.Sprintf(`<pattern id="%s" width="2" height="2" patternUnits="userSpaceOnUse"><rect width="2" height="2" fill="%s"/><line x1="0" y1="0" x2="0" y2="2" stroke="%s" stroke-width="1"/></pattern>`, id, bg, fg)
	case "ltVert": // Light Vertical â€” thin line, 8px spacing
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="4" y1="0" x2="4" y2="8" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)
	case "dkVert": // Dark Vertical â€” thick line, 4px spacing
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="1" y1="0" x2="1" y2="4" stroke="%s" stroke-width="2"/></pattern>`, id, bg, fg)
	case "dashVert": // Dashed Vertical
		return fmt.Sprintf(`<pattern id="%s" width="4" height="8" patternUnits="userSpaceOnUse"><rect width="4" height="8" fill="%s"/><line x1="2" y1="0" x2="2" y2="4" stroke="%s" stroke-width="1"/></pattern>`, id, bg, fg)

	// ---- Diagonal stripes (light/narrow) ----
	case "ltDnDiag": // Light Downward Diagonal
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="0" y1="0" x2="8" y2="8" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)
	case "ltUpDiag": // Light Upward Diagonal
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="8" y1="0" x2="0" y2="8" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)
	case "dnDiag": // Downward Diagonal
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="0" y1="0" x2="4" y2="4" stroke="%s" stroke-width="1"/></pattern>`, id, bg, fg)
	case "upDiag": // Upward Diagonal
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="4" y1="0" x2="0" y2="4" stroke="%s" stroke-width="1"/></pattern>`, id, bg, fg)

	// ---- Diagonal stripes (dark/wide) ----
	case "dkDnDiag": // Dark Downward Diagonal â€” thick
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="0" y1="0" x2="4" y2="4" stroke="%s" stroke-width="2"/></pattern>`, id, bg, fg)
	case "dkUpDiag": // Dark Upward Diagonal â€” thick
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="4" y1="0" x2="0" y2="4" stroke="%s" stroke-width="2"/></pattern>`, id, bg, fg)
	case "wdDnDiag": // Wide Downward Diagonal
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="0" y1="0" x2="8" y2="8" stroke="%s" stroke-width="2"/></pattern>`, id, bg, fg)
	case "wdUpDiag": // Wide Upward Diagonal
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="8" y1="0" x2="0" y2="8" stroke="%s" stroke-width="2"/></pattern>`, id, bg, fg)
	case "dashDnDiag": // Dashed Downward Diagonal
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="0" y1="0" x2="4" y2="4" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)
	case "dashUpDiag": // Dashed Upward Diagonal
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="8" y1="0" x2="4" y2="4" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)

	// ---- Crosses and grids ----
	case "cross": // Cross
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="0" y1="2" x2="4" y2="2" stroke="%s" stroke-width="0.5"/><line x1="2" y1="0" x2="2" y2="4" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg, fg)
	case "diagCross": // Diagonal Cross
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="0" y1="0" x2="4" y2="4" stroke="%s" stroke-width="0.5"/><line x1="4" y1="0" x2="0" y2="4" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg, fg)
	case "smGrid": // Small Grid
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><line x1="0" y1="0" x2="4" y2="0" stroke="%s" stroke-width="0.5"/><line x1="0" y1="0" x2="0" y2="4" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg, fg)
	case "lgGrid": // Large Grid
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="0" y1="0" x2="8" y2="0" stroke="%s" stroke-width="0.5"/><line x1="0" y1="0" x2="0" y2="8" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg, fg)
	case "dotGrid": // Dotted Grid
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><rect x="0" y="0" width="1" height="1" fill="%s"/></pattern>`, id, bg, fg)

	// ---- Checkerboards ----
	case "smCheck": // Small Checkerboard
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><rect x="0" y="0" width="2" height="2" fill="%s"/><rect x="2" y="2" width="2" height="2" fill="%s"/></pattern>`, id, bg, fg, fg)
	case "lgCheck": // Large Checkerboard
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><rect x="0" y="0" width="4" height="4" fill="%s"/><rect x="4" y="4" width="4" height="4" fill="%s"/></pattern>`, id, bg, fg, fg)

	// ---- Diamonds ----
	case "openDmnd": // Open Diamond
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="4" y1="0" x2="8" y2="4" stroke="%s" stroke-width="0.5"/><line x1="8" y1="4" x2="4" y2="8" stroke="%s" stroke-width="0.5"/><line x1="4" y1="8" x2="0" y2="4" stroke="%s" stroke-width="0.5"/><line x1="0" y1="4" x2="4" y2="0" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg, fg, fg, fg)
	case "solidDmnd": // Solid Diamond
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><polygon points="4,0 8,4 4,8 0,4" fill="%s"/></pattern>`, id, bg, fg)
	case "dotDmnd": // Dotted Diamond
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><rect x="4" y="0" width="1" height="1" fill="%s"/><rect x="2" y="2" width="1" height="1" fill="%s"/><rect x="6" y="2" width="1" height="1" fill="%s"/><rect x="0" y="4" width="1" height="1" fill="%s"/><rect x="4" y="4" width="1" height="1" fill="%s"/><rect x="2" y="6" width="1" height="1" fill="%s"/><rect x="6" y="6" width="1" height="1" fill="%s"/></pattern>`, id, bg, fg, fg, fg, fg, fg, fg, fg)

	// ---- Bricks ----
	case "horzBrick": // Horizontal Brick
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="0" y1="4" x2="8" y2="4" stroke="%s" stroke-width="0.5"/><line x1="0" y1="0" x2="0" y2="4" stroke="%s" stroke-width="0.5"/><line x1="4" y1="4" x2="4" y2="8" stroke="%s" stroke-width="0.5"/><line x1="0" y1="8" x2="8" y2="8" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg, fg, fg, fg)
	case "diagBrick": // Diagonal Brick
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="0" y1="0" x2="4" y2="4" stroke="%s" stroke-width="0.5"/><line x1="4" y1="4" x2="8" y2="0" stroke="%s" stroke-width="0.5"/><line x1="0" y1="4" x2="4" y2="8" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg, fg, fg)

	// ---- Trellis, Weave, Plaid, Shingle ----
	case "trellis": // Trellis â€” interlocking diagonal lines
		return fmt.Sprintf(`<pattern id="%s" width="4" height="4" patternUnits="userSpaceOnUse"><rect width="4" height="4" fill="%s"/><rect x="0" y="0" width="1" height="1" fill="%s"/><rect x="2" y="0" width="1" height="1" fill="%s"/><rect x="1" y="1" width="1" height="1" fill="%s"/><rect x="3" y="1" width="1" height="1" fill="%s"/><rect x="0" y="2" width="1" height="1" fill="%s"/><rect x="2" y="2" width="1" height="1" fill="%s"/><rect x="1" y="3" width="1" height="1" fill="%s"/><rect x="3" y="3" width="1" height="1" fill="%s"/></pattern>`, id, fg, bg, bg, bg, bg, bg, bg, bg, bg)
	case "weave": // Weave â€” woven pattern approximation
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="0" y1="0" x2="4" y2="4" stroke="%s" stroke-width="1"/><line x1="4" y1="4" x2="8" y2="0" stroke="%s" stroke-width="1"/><line x1="0" y1="4" x2="4" y2="8" stroke="%s" stroke-width="1"/><line x1="4" y1="8" x2="8" y2="4" stroke="%s" stroke-width="1"/></pattern>`, id, bg, fg, fg, fg, fg)
	case "plaid": // Plaid â€” overlapping h+v bands
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><rect x="0" y="0" width="4" height="4" fill="%s" opacity="0.5"/><rect x="0" y="0" width="8" height="2" fill="%s" opacity="0.3"/><rect x="0" y="0" width="2" height="8" fill="%s" opacity="0.3"/></pattern>`, id, bg, fg, fg, fg)
	case "shingle": // Shingle
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><line x1="0" y1="4" x2="4" y2="0" stroke="%s" stroke-width="0.5"/><line x1="4" y1="8" x2="8" y2="4" stroke="%s" stroke-width="0.5"/><line x1="0" y1="8" x2="8" y2="8" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg, fg, fg)

	// ---- Other named patterns ----
	case "sphere": // Sphere â€” small circles
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><circle cx="4" cy="4" r="3" fill="none" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)
	case "wave": // Wave
		return fmt.Sprintf(`<pattern id="%s" width="8" height="4" patternUnits="userSpaceOnUse"><rect width="8" height="4" fill="%s"/><path d="M0 2Q2 0 4 2Q6 4 8 2" fill="none" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)
	case "divot": // Divot â€” V shapes
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><path d="M2 2L4 4L6 2" fill="none" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)
	case "zigZag": // ZigZag
		return fmt.Sprintf(`<pattern id="%s" width="8" height="4" patternUnits="userSpaceOnUse"><rect width="8" height="4" fill="%s"/><path d="M0 4L2 0L4 4L6 0L8 4" fill="none" stroke="%s" stroke-width="0.5"/></pattern>`, id, bg, fg)
	case "smConfetti": // Small Confetti â€” scattered dots
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><rect x="1" y="1" width="1" height="1" fill="%s"/><rect x="5" y="3" width="1" height="1" fill="%s"/><rect x="3" y="5" width="1" height="1" fill="%s"/><rect x="7" y="7" width="1" height="1" fill="%s"/></pattern>`, id, bg, fg, fg, fg, fg)
	case "lgConfetti": // Large Confetti â€” scattered larger dots
		return fmt.Sprintf(`<pattern id="%s" width="8" height="8" patternUnits="userSpaceOnUse"><rect width="8" height="8" fill="%s"/><rect x="0" y="0" width="2" height="2" fill="%s"/><rect x="4" y="4" width="2" height="2" fill="%s"/></pattern>`, id, bg, fg, fg)

	default:
		// Fallback: 2x2 pct25-style (any unrecognized prst from XML)
		return fmt.Sprintf(`<pattern id="%s" width="2" height="2" patternUnits="userSpaceOnUse"><rect width="2" height="2" fill="%s"/><rect width="1" height="1" fill="%s"/></pattern>`, id, bg, fg)
	}
}

// buildTextDecorationAttr returns an SVG attribute string for text-decoration based on
// OOXML underline (u) and strikethrough (strike) values. Returns empty string if neither is set.
func buildTextDecorationAttr(underline, strikethrough string) string {
	var decos []string
	decoStyle := ""
	if underline != "" {
		decos = append(decos, "underline")
		switch underline {
		case "dbl", "wavyDbl":
			decoStyle = "double"
		case "dotted", "dottedHeavy":
			decoStyle = "dotted"
		case "dash", "dashHeavy", "dashLong", "dashLongHeavy":
			decoStyle = "dashed"
		case "wavy", "wavyHeavy":
			decoStyle = "wavy"
		case "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy":
			decoStyle = "dashed"
		}
	}
	if strikethrough != "" {
		decos = append(decos, "line-through")
	}
	if len(decos) == 0 {
		return ""
	}
	attr := fmt.Sprintf(` text-decoration="%s"`, strings.Join(decos, " "))
	if decoStyle != "" {
		attr += fmt.Sprintf(` text-decoration-style="%s"`, decoStyle)
	}
	return attr
}

// renderShapeRichText renders the text content of a shape using full rich text logic (bullets, wrapping, runs).
// It returns the SVG string for the text. It assumes the caller handles geometry transforms.
// Transforms for text (rotation) are handled here if they are text-specific (like rotation without flip),
// but usually the caller handles the main transform group.
// However, to match "textbox" logic, we need to know if we should apply rotation-only transform here or if it's applied externally.
// The "textbox" case applies `tbTransform` (rotation only) wrapping the text.
// The "rect"/"ellipse" cases apply transforms differently based on flips.
// For simplicity, this function will render the text at the correct (x,y) with rotation if specified in arguments or implicitly via s.Rotation.
// actually, let's keep it simple: this function renders the text elements. The caller wraps them in <g transform> if needed.
// BUT, vertical alignment depends on effective height, which depends on wrapping width.
func renderShapeRichText(s RenderShape, x, y, w, h int, fs int, cfg *RenderConfig) string {
	var out strings.Builder

	fontWeight := "normal"
	if s.Bold {
		fontWeight = "bold"
	}
	txtColor := s.TextColor
	if txtColor == "" {
		txtColor = "#000000"
	}

	// Determine font-family attribute for fallback
	fontFamilyAttr := ""
	if s.FontFamily != "" {
		fontFamilyAttr = fmt.Sprintf(` font-family="%s"`, escapeXMLText(s.FontFamily))
	}

	// Margins from bodyPr (Format Shape â†’ Text Box); fallback when zero
	lInsPx := s.BodyLIns
	if lInsPx <= 0 {
		lInsPx = 4
	}
	rInsPx := s.BodyRIns
	if rInsPx <= 0 {
		rInsPx = 4
	}
	tInsPx := s.BodyTIns
	if tInsPx <= 0 {
		tInsPx = 4
	}
	bInsPx := s.BodyBIns
	if bInsPx <= 0 {
		bInsPx = 4
	}
	effectiveW := w - lInsPx - rInsPx
	if effectiveW < 20 {
		effectiveW = 20
	}
	effectiveH := h - tInsPx - bInsPx
	if effectiveH < fs {
		effectiveH = fs
	}
	// Wrap OFF (case d): do not wrap; single line, overflow horizontal. Wrap ON: use effective width; for "shape" use conservative width.
	wrapW := effectiveW
	if !s.WrapText {
		wrapW = 999999 // no wrap: one line
	} else if s.TextAutofit == "shape" {
		wrapW = int(float64(effectiveW) * 0.92)
		if wrapW < 20 {
			wrapW = 20
		}
	}

	// Determine horizontal alignment from paragraph alignment or shape alignment (from XML)
	hAlgn := s.Alignment
	if hAlgn == "" {
		hAlgn = "l"
	}

	// Text-anchor and x position from bodyPr/paragraph alignment
	var textAnchorAttr string
	textX := x + lInsPx
	switch hAlgn {
	case "ctr":
		textAnchorAttr = ` text-anchor="middle"`
		textX = x + w/2
	case "r":
		textAnchorAttr = ` text-anchor="end"`
		textX = x + w - rInsPx
	default:
		textAnchorAttr = ""
		textX = x + lInsPx
	}

	// Case b (norm): wrap at pre-shrink font size so we get 2 lines at 20px, not 1 line at 16px.
	// fsForWrap := fs (omitted for now as we don't have unscaled fs passed in)

	maxCharsPerLine := cfg.maxCharsPerLine(wrapW, fs)
	lh := cfg.lineHeight(fs)

	// Collect either run-level placements (one <text> per run) or line-level (one per wrapped line).
	type tbLine struct {
		text     string
		fs       int
		lh       int
		color    string
		weight   string
		fontAttr string
	}
	type runPlacement struct {
		text          string
		fs            int
		x, y, w       int
		color         string
		weight        string
		fontAttr      string
		underline     string
		strikethrough string
		baselineShift int
		hyperlinkURL  string
		gradientDef   *SVGGradient
		patternFill   *RenderPattern
		strokeColor   string
		strokeWidth   float64
		shadow        *RenderShadow
		hasReflection bool
	}
	var allLines []tbLine
	var runPlacements []runPlacement

	if len(s.Paragraphs) > 0 {
		// Build run-level placements so we emit one <text> per run (matches expected SVG structure).
		charWidthRatio := cfg.TextLayout.CharWidthRatio
		if charWidthRatio <= 0 {
			charWidthRatio = 0.48
		}
		var lineY, lineX, totalRunLinesH int
		firstLineFS := fs

		// Counters for auto-numbering (level -> current count)
		autoNumCounters := make(map[int]int)

		for pi, para := range s.Paragraphs {
			paraFS := fs
			paraColor := txtColor
			paraWeight := fontWeight
			paraFontAttr := fontFamilyAttr
			for _, run := range para.Runs {
				if run.FontSize > 0 {
					paraFS = run.FontSize
				}
				if run.Color != "" {
					paraColor = run.Color
				}
				if run.Bold {
					paraWeight = "bold"
				}
				if run.FontFamily != "" {
					paraFontAttr = fmt.Sprintf(` font-family="%s"`, escapeXMLText(run.FontFamily))
				}
			}
			paraLH := cfg.lineHeight(paraFS)
			if len(runPlacements) == 0 {
				firstLineFS = paraFS
			}
			// New paragraph starts on a new line
			paraSpcBefore := para.SpaceBefore
			if pi == 0 {
				paraSpcBefore = 0
			}
			totalRunLinesH += paraSpcBefore + paraLH
			lineY = totalRunLinesH - paraLH
			lineX = para.MarginLeft

			// Determine bullet text
			bulletTxt := ""
			if para.BulletType == "char" {
				bulletTxt = para.BulletChar
			} else if para.BulletType == "autoNum" {
				for l := para.Level + 1; l < 10; l++ {
					autoNumCounters[l] = 0
				}
				autoNumCounters[para.Level]++
				count := autoNumCounters[para.Level]
				if para.BulletStartAt > 1 && count == 1 {
					autoNumCounters[para.Level] = para.BulletStartAt
					count = para.BulletStartAt
				}
				switch para.BulletAutoNumType {
				case "arabicPeriod":
					bulletTxt = fmt.Sprintf("%d.", count)
				case "arabicParenR":
					bulletTxt = fmt.Sprintf("%d)", count)
				case "alphaLcParenR":
					r := 'a' + (count-1)%26
					bulletTxt = fmt.Sprintf("%c)", r)
				case "alphaUcParenR":
					r := 'A' + (count-1)%26
					bulletTxt = fmt.Sprintf("%c)", r)
				default:
					bulletTxt = fmt.Sprintf("%d.", count)
				}
			}

			if bulletTxt != "" {
				bx := para.MarginLeft + para.Indent
				bulletW := int(math.Ceil(GetStringWidth(bulletTxt, "", paraFS, paraWeight == "bold", false)))
				runPlacements = append(runPlacements, runPlacement{
					text: bulletTxt, fs: paraFS, x: bx, y: lineY, w: bulletW,
					color: paraColor, weight: paraWeight, fontAttr: paraFontAttr,
				})
			}
			for _, run := range para.Runs {
				runFS := paraFS
				if run.FontSize > 0 {
					runFS = run.FontSize
				}
				runColor := paraColor
				if run.Color != "" {
					runColor = run.Color
				}
				runWeight := paraWeight
				if run.Bold {
					runWeight = "bold"
				}
				runFontAttr := paraFontAttr
				if run.FontFamily != "" {
					runFontAttr = fmt.Sprintf(` font-family="%s"`, escapeXMLText(run.FontFamily))
				}
				// Split "SOLID FILL â€“ DO NOT AUTO FIT" into segments so we emit one <text> per segment (matches expected SVG).
				segments := splitRunTextForPlacement(run.Text)
				// Use passed fs for wrapping (simplified from textbox logic which used unscaled fs)
				fsForWrapRun := runFS

				effectiveWrapW := wrapW - para.MarginLeft
				if effectiveWrapW < 20 {
					effectiveWrapW = 20
				}

				for _, seg := range segments {
					wrapped := wrapTextPx(seg, float64(effectiveWrapW), run.FontFamily, fsForWrapRun, run.Bold, run.Italic)
					for wi, line := range wrapped {
						runW := int(math.Ceil(GetStringWidth(line, run.FontFamily, fsForWrapRun, run.Bold, run.Italic)))
						if runW < 2 {
							runW = 2
						}

						if lineX+runW > wrapW && lineX > para.MarginLeft {
							lineY += paraLH
							lineX = para.MarginLeft
							totalRunLinesH += paraLH
						}

						runPlacements = append(runPlacements, runPlacement{
							text: line, fs: runFS, x: lineX, y: lineY, w: runW,
							color: runColor, weight: runWeight, fontAttr: runFontAttr,
							underline: run.Underline, strikethrough: run.Strikethrough, baselineShift: run.BaselineShift,
							hyperlinkURL: run.HyperlinkURL,
							gradientDef:  run.GradientDef, patternFill: run.PatternFill,
							strokeColor: run.StrokeColor, strokeWidth: run.StrokeWidth,
							shadow: run.Shadow, hasReflection: run.HasReflection,
						})

						lineX += runW

						// Force a new line if wrapText explicitly broke this segment into multiple line pieces
						if wi < len(wrapped)-1 && lineX > para.MarginLeft {
							lineY += paraLH
							lineX = para.MarginLeft
							totalRunLinesH += paraLH
						}
					}
				}
			}
			totalRunLinesH += para.SpaceAfter
		}
		if totalRunLinesH == 0 && len(runPlacements) > 0 {
			totalRunLinesH = cfg.lineHeight(firstLineFS)
		}
		// Vertical alignment for run-level: compute yOff then add runPlacements[].y
		var yOff int
		switch s.TextAnchor {
		case "ctr":
			yOff = y + tInsPx + (effectiveH-totalRunLinesH)/2 + firstLineFS
		case "b":
			yOff = y + h - bInsPx - totalRunLinesH + firstLineFS
		case "t", "":
			yOff = y + tInsPx + firstLineFS + int(math.Round(float64(firstLineFS)*0.2))
		default:
			yOff = y + tInsPx + (effectiveH-totalRunLinesH)/2 + firstLineFS
		}
		var textBlockWidth int
		if s.AnchorCtr && (hAlgn == "l" || hAlgn == "") {
			for _, r := range runPlacements {
				if r.x+r.w > textBlockWidth {
					textBlockWidth = r.x + r.w
				}
			}
		}

		textRunIdx := 0
		for _, r := range runPlacements {
			runX := textX
			if hAlgn == "l" || hAlgn == "" {
				if s.AnchorCtr {
					runX = x + w/2 - textBlockWidth/2 + r.x
				} else {
					runX = textX + r.x
				}
			} else if hAlgn == "ctr" {
				runX = x + w/2
			} else {
				runX = x + w - rInsPx - r.x
			}
			// Build text decoration and baseline-shift attributes
			textDecoAttr := buildTextDecorationAttr(r.underline, r.strikethrough)
			baselineAttr := ""
			runFSActual := r.fs
			runYActual := yOff + r.y
			if r.baselineShift > 0 {
				runFSActual = int(math.Round(float64(r.fs) * 0.6))
				if runFSActual < 1 {
					runFSActual = 1
				}
				baselineAttr = ` baseline-shift="super"`
			} else if r.baselineShift < 0 {
				runFSActual = int(math.Round(float64(r.fs) * 0.6))
				if runFSActual < 1 {
					runFSActual = 1
				}
				baselineAttr = ` baseline-shift="sub"`
			}

			fillAttr := r.color
			strokeAttr := ""
			var textDefs strings.Builder
			filterAttr := ""

			// Text gradient fill
			if r.gradientDef != nil {
				gradID := fmt.Sprintf("txtGrad_%d_%d", s.XmlID, textRunIdx)
				r.gradientDef.ID = gradID
				textDefs.WriteString("<defs>")
				textDefs.WriteString(renderSVGGradientDef(r.gradientDef))
				textDefs.WriteString("</defs>")
				fillAttr = fmt.Sprintf("url(#%s)", gradID)
			}

			// Text pattern fill
			if r.patternFill != nil {
				pattID := fmt.Sprintf("txtPatt_%d_%d", s.XmlID, textRunIdx)
				textDefs.WriteString("<defs>")
				textDefs.WriteString(renderPatternDef(r.patternFill, pattID))
				textDefs.WriteString("</defs>")
				fillAttr = fmt.Sprintf("url(#%s)", pattID)
			}

			// Text outline/stroke
			if r.strokeColor != "" && r.strokeWidth > 0 {
				strokeAttr = fmt.Sprintf(` stroke="%s" stroke-width="%.1f" paint-order="stroke"`, r.strokeColor, r.strokeWidth)
			}

			// Text shadow filter
			if r.shadow != nil {
				shdwID := fmt.Sprintf("txtShdw_%d_%d", s.XmlID, textRunIdx)
				if r.shadow.Type == "inner" {
					textDefs.WriteString("<defs>")
					textDefs.WriteString(generateInnerShadowFilter(r.shadow, shdwID))
					textDefs.WriteString("</defs>")
				} else {
					textDefs.WriteString("<defs>")
					textDefs.WriteString(generateShadowFilter(r.shadow, shdwID, false))
					textDefs.WriteString("</defs>")
				}
				filterAttr = fmt.Sprintf(` filter="url(#%s)"`, shdwID)
			}

			if textDefs.Len() > 0 {
				out.WriteString(textDefs.String())
			}

			textElem := fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-weight="%s"%s%s%s%s%s%s>%s</text>
`, runX, runYActual, runFSActual, fillAttr, r.weight, r.fontAttr, textAnchorAttr, textDecoAttr, baselineAttr, strokeAttr, filterAttr, escapeXMLText(r.text))
			out.WriteString(wrapInSVGAnchor(r.hyperlinkURL, textElem))

			// Reflection: duplicate text with vertical flip and fading opacity
			if r.hasReflection {
				reflY := runYActual + runFSActual/3
				out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-weight="%s"%s%s opacity="0.3" transform="scale(1,-1) translate(0,%d)">%s</text>
`, runX, reflY, runFSActual, fillAttr, r.weight, r.fontAttr, textAnchorAttr, -(2 * reflY), escapeXMLText(r.text)))
			}

			textRunIdx++
		}
	} else {
		// No paragraphs: use line-based layout (wrap s.Text).
		lines := wrapText(s.Text, maxCharsPerLine)
		for _, line := range lines {
			allLines = append(allLines, tbLine{text: line, fs: fs, lh: lh, color: txtColor, weight: fontWeight, fontAttr: fontFamilyAttr})
		}
		totalTextH := 0
		for _, l := range allLines {
			totalTextH += l.lh
		}
		var yOff int
		switch s.TextAnchor {
		case "ctr":
			yOff = y + tInsPx + (effectiveH-totalTextH)/2 + fs
		case "b":
			yOff = y + h - bInsPx - totalTextH + fs
		case "t", "":
			yOff = y + tInsPx + fs + int(math.Round(float64(fs)*0.2))
		default:
			yOff = y + tInsPx + (effectiveH-totalTextH)/2 + fs
		}
		for _, l := range allLines {
			if l.text == "" {
				yOff += l.lh
				continue
			}
			out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-weight="%s"%s%s>%s</text>
`, textX, yOff, l.fs, l.color, l.weight, l.fontAttr, textAnchorAttr, escapeXMLText(l.text)))
			yOff += l.lh
		}
	}
	return out.String()
}

const (
	renderModeAll = iota
	renderModeGeo
	renderModeText
)

func renderShapeSVG(s RenderShape, cfg *RenderConfig) string {
	content := renderShapeContent(s, cfg, renderModeAll)
	if content == "" {
		return ""
	}

	if s.Shadow != nil {
		var finalOut strings.Builder
		filterID := fmt.Sprintf("shdw_%d_%d_%d_%d", s.XmlID, s.X, s.Y, s.Shadow.Dir)
		isPerspective := s.Shadow.Sy != 0 || s.Shadow.Kx != 0

		var filterDef string
		if s.Shadow.Type == "inner" {
			filterDef = generateInnerShadowFilter(s.Shadow, filterID)
			isPerspective = false
		} else {
			filterDef = generateShadowFilter(s.Shadow, filterID, isPerspective)
		}

		finalOut.WriteString("<defs>")
		finalOut.WriteString(filterDef)
		finalOut.WriteString("</defs>")

		if isPerspective {
			shapeID := fmt.Sprintf("shape_%d", s.XmlID)

			sxFactor := 1.0
			if s.Shadow.Sx != 0 {
				sxFactor = float64(s.Shadow.Sx) / 100000.0
			}
			syFactor := float64(s.Shadow.Sy) / 100000.0

			kxRad := float64(s.Shadow.Kx) * math.Pi / 10800000.0
			skewC := math.Tan(kxRad)

			dirRad := float64(s.Shadow.Dir) * math.Pi / 10800000.0
			distPx := float64(s.Shadow.Dist) / emuPerPx
			distX := distPx * math.Cos(dirRad)
			distY := distPx * math.Sin(dirRad)

			// Anchor logic (default: bottom center)
			x := float64(emuToPx(s.X))
			y := float64(emuToPx(s.Y))
			w := float64(emuToPx(s.W))
			h := float64(emuToPx(s.H))

			anchorX := x + w/2
			anchorY := y + h

			if s.Shadow.Algn == "bl" {
				anchorX = x
				anchorY = y + h
			} else if s.Shadow.Algn == "br" {
				anchorX = x + w
				anchorY = y + h
			}

			a := sxFactor
			c := skewC
			d := syFactor

			// Matrix transform relative to anchor
			e := -a*anchorX - c*anchorY + anchorX + distX
			f := -d*anchorY + anchorY + distY

			finalOut.WriteString(fmt.Sprintf(`<use href="#%s" filter="url(#%s)" transform="matrix(%.5g 0 %.5g %.5g %.5g %.5g)"/>`,
				shapeID, filterID, a, c, d, e, f))

			finalOut.WriteString(fmt.Sprintf(`<g id="%s">`, shapeID))
			finalOut.WriteString(renderShapeContent(s, cfg, renderModeGeo))
			finalOut.WriteString("</g>")
			finalOut.WriteString(renderShapeContent(s, cfg, renderModeText))
		} else if s.Shadow.Type == "inner" {
			// For inner shadow, apply filter to the shape group directly
			// The filter composites the shadow ON TOP of the source graphic
			finalOut.WriteString(fmt.Sprintf(`<g filter="url(#%s)">`, filterID))
			finalOut.WriteString(renderShapeContent(s, cfg, renderModeGeo))
			finalOut.WriteString("</g>")
			finalOut.WriteString(renderShapeContent(s, cfg, renderModeText))
		} else {
			finalOut.WriteString(fmt.Sprintf(`<g filter="url(#%s)">`, filterID))
			finalOut.WriteString(renderShapeContent(s, cfg, renderModeGeo))
			finalOut.WriteString("</g>")
			finalOut.WriteString(renderShapeContent(s, cfg, renderModeText))
		}
		return wrapInSVGAnchor(s.HyperlinkURL, finalOut.String())
	}

	return wrapInSVGAnchor(s.HyperlinkURL, content)
}

// wrapInSVGAnchor wraps SVG content in an <a> element if url is non-empty.
func wrapInSVGAnchor(url, content string) string {
	if url == "" || content == "" {
		return content
	}
	return fmt.Sprintf(`<a href="%s" target="_blank">%s</a>`, escapeXMLText(url), content)
}

func generateInnerShadowFilter(s *RenderShadow, id string) string {
	// Inner shadow implementation:
	// 1. Invert the shape alpha (create a hole)
	// 2. Blur and offset the inverted mask
	// 3. Clip the result to the original shape alpha
	// 4. Composite over original shape

	dirRad := float64(s.Dir) * math.Pi / 10800000.0
	distPx := float64(s.Dist) / emuPerPx
	// Invert direction for inner shadow to create "inset" effect correctly
	dx := -distPx * math.Cos(dirRad)
	dy := -distPx * math.Sin(dirRad)
	blurStdDev := float64(s.BlurRad) / emuPerPx

	return fmt.Sprintf(`
	<filter id="%s" x="-50%%" y="-50%%" width="200%%" height="200%%">
		<!-- 1. Flood opacity 1 to create full plane -->
		<feFlood flood-color="black" flood-opacity="1" result="floodFull"/>
		<!-- 2. Cut out the shape from the full plane (Inverted Alpha) -->
		<feComposite operator="out" in="floodFull" in2="SourceAlpha" result="invertedAlpha"/>
		<!-- 3. Blur the inverted alpha -->
		<feGaussianBlur stdDeviation="%.2f" in="invertedAlpha" result="blurredInverted"/>
		<!-- 4. Offset the blurred inverted alpha -->
		<feOffset dx="%.2f" dy="%.2f" in="blurredInverted" result="offsetBlurred"/>
		<!-- 5. Clip shadow to inside of original shape -->
		<feComposite operator="in" in="offsetBlurred" in2="SourceAlpha" result="innerShadowAlpha"/>
		<!-- 6. Color the shadow -->
		<feFlood flood-color="%s" flood-opacity="%.2f" result="shadowColor"/>
		<feComposite operator="in" in="shadowColor" in2="innerShadowAlpha" result="finalShadow"/>
		<!-- 7. Merge original graphic and shadow -->
		<feMerge>
			<feMergeNode in="SourceGraphic"/>
			<feMergeNode in="finalShadow"/>
		</feMerge>
	</filter>`, id, blurStdDev, dx, dy, s.Color, s.Alpha)
}

func renderShapeContent(s RenderShape, cfg *RenderConfig, renderMode int) string {
	x := emuToPx(s.X)
	y := emuToPx(s.Y)
	w := emuToPx(s.W)
	h := emuToPx(s.H)

	tx := emuToPx(s.TextX)
	ty := emuToPx(s.TextY)
	tw := emuToPx(s.TextW)
	th := emuToPx(s.TextH)
	if tw == 0 && th == 0 {
		tx, ty, tw, th = x, y, w, h
	}

	if w <= 0 && h <= 0 && s.Text == "" && s.ImageData == "" {
		return ""
	}

	// Font size from PPT XML: shape runs â†’ presentation.xml defaultTextStyle â†’ fallback
	fs := s.FontSize
	if fs == 0 && cfg.DefaultFontSizePx > 0 {
		fs = cfg.DefaultFontSizePx
	}
	if fs == 0 {
		fs = 16 // fallback only when no sz in shape and no defaultTextStyle in presentation.xml
	}

	// For norm autofit we wrap at original font size, then shrink for display. Keep pre-shrink fs for wrap calculation so case b gets 2 lines at 20px, not 1 line at 16px.
	// wrapFS := fs
	// Shrink Text on Overflow (norm): scale font so text fits â€” only when wrap is ON (case b). When wrap is OFF (case d), overflow is horizontal and shrink has no effect.
	if s.TextAutofit == "norm" && s.WrapText && s.Text != "" {
		effW := w - s.BodyLIns - s.BodyRIns
		effH := h - s.BodyTIns - s.BodyBIns
		if effW < 20 {
			effW = 20
		}
		if effH < fs {
			effH = fs
		}
		scale := normAutofitScale(s, effW, effH, fs, cfg)
		fs = int(math.Round(float64(fs) * scale))
		if fs < 1 {
			fs = 1
		}
	}
	// "shape" (Resize Shape to Fit Text): design-time in PPT; we wrap and use fixed size (same as "none")

	fontWeight := "normal"
	if s.Bold {
		fontWeight = "bold"
	}

	// Determine font-family attribute
	fontFamilyAttr := ""
	if s.FontFamily != "" {
		fontFamilyAttr = fmt.Sprintf(` font-family="%s"`, escapeXMLText(s.FontFamily))
	}

	var out strings.Builder

	switch s.ShapeType {
	case "title":
		fillOverride := s.Fill
		if s.GradientDef != nil {
			out.WriteString("  <defs>\n")
			out.WriteString(renderSVGGradientDef(s.GradientDef))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", s.GradientDef.ID)
		}
		if s.PatternFill != nil {
			pattID := fmt.Sprintf("patt_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(renderPatternDef(s.PatternFill, pattID))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", pattID)
		} else if s.ImageData != "" && s.ImageMIME != "" {
			bgID := fmt.Sprintf("bgimg_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(fmt.Sprintf(`    <pattern id="%s" patternUnits="userSpaceOnUse" width="%d" height="%d">
      <image width="%d" height="%d" preserveAspectRatio="none" href="data:%s;base64,%s"/>
    </pattern>
`, bgID, w, h, w, h, s.ImageMIME, s.ImageData))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", bgID)
		}
		if renderMode == renderModeAll || renderMode == renderModeGeo {
			if fillOverride == "" {
				fillOverride = "none"
			}
			tbStroke := "none"
			if s.Stroke != "" {
				tbStroke = s.Stroke
			}
			if fillOverride != "none" || tbStroke != "none" {
				dashAttr := prstDashToSVG(s.DashStyle, s.StrokeWidth, s.LineCap)
				lineCapAttr := ""
				if s.LineCap != "" {
					lineCapAttr = fmt.Sprintf(` stroke-linecap="%s"`, s.LineCap)
				}
				tbSW := "1.33"
				if s.StrokeWidth > 0 {
					tbSW = fmt.Sprintf("%.2f", s.StrokeWidth)
				}
				out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="%s" stroke="%s" stroke-width="%s"%s%s/>
`, x, y, w, h, fillOverride, tbStroke, tbSW, dashAttr, lineCapAttr))
			}
		} // end geo mode
		if renderMode == renderModeAll || renderMode == renderModeText {
			titleColor := s.TextColor
			if titleColor == "" {
				titleColor = "#000000"
			}
			titleFontWeight := fontWeight
			if strings.Contains(s.FontFamily, "Light") && !s.Bold {
				titleFontWeight = "300"
			}
			tInsPx := s.BodyTIns
			if tInsPx <= 0 {
				tInsPx = int(math.Round(45720.0 / emuPerPx))
			}
			bInsPx := s.BodyBIns
			if bInsPx <= 0 {
				bInsPx = tInsPx
			}
			lInsPx := s.BodyLIns
			if lInsPx <= 0 {
				lInsPx = int(math.Round(91440.0 / emuPerPx))
			}
			rInsPx := s.BodyRIns
			if rInsPx <= 0 {
				rInsPx = lInsPx
			}

			titleAnchor := s.TextAnchor
			if titleAnchor == "" {
				titleAnchor = "t"
			}
			// Horizontal alignment from paragraph/layout (e.g. ctrTitle â†’ center)
			textX := x + lInsPx
			titleAnchorAttr := ""
			if s.Alignment == "ctr" {
				textX = x + w/2
				titleAnchorAttr = ` text-anchor="middle"`
			} else if s.Alignment == "r" {
				textX = x + w - rInsPx
				titleAnchorAttr = ` text-anchor="end"`
			} else if textX <= 0 {
				textX = 73
			}
			if len(s.Paragraphs) > 0 {
				var totalH int
				for _, para := range s.Paragraphs {
					if para.IsEmpty {
						lh := paraLineHeight(para, fs, cfg)
						totalH += lh
						continue
					}
					paraText := ""
					for _, run := range para.Runs {
						paraText += run.Text
					}
					paraText = strings.TrimSpace(paraText)
					if paraText == "" {
						lh := paraLineHeight(para, fs, cfg)
						totalH += lh
						continue
					}
					lh := paraLineHeight(para, fs, cfg)
					totalH += lh
				}
				var startY int
				baselineOffsetPct := GetFontBaselineOffsetPct(s.FontFamily, s.Bold, false)
				baselineOffset := int(float64(fs) * baselineOffsetPct)
				switch titleAnchor {
				case "ctr":
					startY = y + tInsPx + (h-tInsPx-bInsPx-totalH)/2 + baselineOffset
				case "b":
					startY = y + h - bInsPx - totalH + baselineOffset
				default:
					startY = y + tInsPx + baselineOffset
				}
				titleRunIdx := 0
				for _, para := range s.Paragraphs {
					if para.IsEmpty {
						lh := paraLineHeight(para, fs, cfg)
						startY += lh
						continue
					}
					paraText := ""
					for _, run := range para.Runs {
						paraText += run.Text
					}
					paraText = strings.TrimSpace(paraText)
					if paraText == "" {
						lh := paraLineHeight(para, fs, cfg)
						startY += lh
						continue
					}
					lh := paraLineHeight(para, fs, cfg)

					for _, run := range para.Runs {
						runText := strings.TrimSpace(run.Text)
						if runText == "" {
							continue
						}
						runFillAttr := titleColor
						runStrokeAttr := ""
						runFilterAttr := ""
						runWeight := titleFontWeight
						if run.Bold {
							runWeight = "bold"
						}
						runFS := fs
						if run.FontSize > 0 {
							runFS = run.FontSize
						}
						var runDefs strings.Builder

						if run.PatternFill != nil {
							pattID := fmt.Sprintf("titlePatt_%d_%d", s.XmlID, titleRunIdx)
							runDefs.WriteString("<defs>")
							runDefs.WriteString(renderPatternDef(run.PatternFill, pattID))
							runDefs.WriteString("</defs>")
							runFillAttr = fmt.Sprintf("url(#%s)", pattID)
						} else if run.GradientDef != nil {
							gradID := fmt.Sprintf("titleGrad_%d_%d", s.XmlID, titleRunIdx)
							run.GradientDef.ID = gradID
							runDefs.WriteString("<defs>")
							runDefs.WriteString(renderSVGGradientDef(run.GradientDef))
							runDefs.WriteString("</defs>")
							runFillAttr = fmt.Sprintf("url(#%s)", gradID)
						} else if run.Color != "" {
							runFillAttr = run.Color
						}

						if run.StrokeColor != "" && run.StrokeWidth > 0 {
							runStrokeAttr = fmt.Sprintf(` stroke="%s" stroke-width="%.1f" paint-order="stroke"`, run.StrokeColor, run.StrokeWidth)
						}

						if run.Shadow != nil {
							shdwID := fmt.Sprintf("titleShdw_%d_%d", s.XmlID, titleRunIdx)
							if run.Shadow.Type == "inner" {
								runDefs.WriteString("<defs>")
								runDefs.WriteString(generateInnerShadowFilter(run.Shadow, shdwID))
								runDefs.WriteString("</defs>")
							} else {
								runDefs.WriteString("<defs>")
								runDefs.WriteString(generateShadowFilter(run.Shadow, shdwID, false))
								runDefs.WriteString("</defs>")
							}
							runFilterAttr = fmt.Sprintf(` filter="url(#%s)"`, shdwID)
						}

						if runDefs.Len() > 0 {
							out.WriteString(runDefs.String())
						}
						out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-weight="%s"%s%s%s%s%s>%s</text>
`, textX, startY, runFS, runFillAttr, runWeight, fontFamilyAttr, titleAnchorAttr, runStrokeAttr, runFilterAttr, buildTextDecorationAttr(run.Underline, run.Strikethrough), escapeXMLText(runText)))
						titleRunIdx++
					}
					startY += lh
				}
			} else {
				var titleY int
				switch titleAnchor {
				case "ctr":
					titleY = y + h/2 + fs/3
				case "b":
					titleY = y + h - bInsPx
				default:
					titleY = y + tInsPx + fs
				}
				out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-weight="%s"%s%s>%s</text>
`, textX, titleY, fs, titleColor, titleFontWeight, titleAnchorAttr, fontFamilyAttr, escapeXMLText(s.Text)))
			}
		} // end text mode

	case "body":
		fillOverride := s.Fill
		if s.GradientDef != nil {
			out.WriteString("  <defs>\n")
			out.WriteString(renderSVGGradientDef(s.GradientDef))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", s.GradientDef.ID)
		}
		if s.PatternFill != nil {
			pattID := fmt.Sprintf("patt_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(renderPatternDef(s.PatternFill, pattID))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", pattID)
		} else if s.ImageData != "" && s.ImageMIME != "" {
			bgID := fmt.Sprintf("bgimg_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(fmt.Sprintf(`    <pattern id="%s" patternUnits="userSpaceOnUse" width="%d" height="%d">
      <image width="%d" height="%d" preserveAspectRatio="none" href="data:%s;base64,%s"/>
    </pattern>
`, bgID, w, h, w, h, s.ImageMIME, s.ImageData))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", bgID)
		}
		if renderMode == renderModeAll || renderMode == renderModeGeo {
			if fillOverride == "" {
				fillOverride = "none"
			}
			tbStroke := "none"
			if s.Stroke != "" {
				tbStroke = s.Stroke
			}
			if fillOverride != "none" || tbStroke != "none" {
				dashAttr := prstDashToSVG(s.DashStyle, s.StrokeWidth, s.LineCap)
				lineCapAttr := ""
				if s.LineCap != "" {
					lineCapAttr = fmt.Sprintf(` stroke-linecap="%s"`, s.LineCap)
				}
				tbSW := "1.33"
				if s.StrokeWidth > 0 {
					tbSW = fmt.Sprintf("%.2f", s.StrokeWidth)
				}
				out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="%s" stroke="%s" stroke-width="%s"%s%s/>
`, x, y, w, h, fillOverride, tbStroke, tbSW, dashAttr, lineCapAttr))
			}
		} // end geo mode
		if renderMode == renderModeAll || renderMode == renderModeText {
			lInsPx, rInsPx, tInsPx, bInsPx := cfg.insetsPx()

			effectiveW := w - lInsPx - rInsPx
			if effectiveW < 40 {
				effectiveW = 40
			}
			effectiveH := h - tInsPx - bInsPx
			if effectiveH < fs {
				effectiveH = fs
			}

			if len(s.Paragraphs) > 0 {
				renderMultiParagraph(&out, s, x, y, w, h, lInsPx, rInsPx, tInsPx, bInsPx, fs, fontWeight, fontFamilyAttr, cfg)
			} else {
				maxChars := cfg.maxCharsPerLine(effectiveW, fs)
				if !s.WrapText {
					maxChars = 999999
				}
				lines := wrapText(s.Text, maxChars)
				lineHeight := cfg.lineHeight(fs)

				totalTextH := len(lines) * lineHeight
				var startY int
				anchor := s.TextAnchor
				if anchor == "" {
					anchor = "t"
				}
				switch anchor {
				case "b":
					startY = y + h - bInsPx - totalTextH + fs
				case "ctr":
					startY = y + tInsPx + (effectiveH-totalTextH)/2 + fs
				default:
					startY = y + tInsPx + fs
				}

				textX := x + lInsPx
				textAnchorAttr := ""
				if s.Alignment == "ctr" {
					textX = x + w/2
					textAnchorAttr = ` text-anchor="middle"`
				} else if s.Alignment == "r" {
					textX = x + w - rInsPx
					textAnchorAttr = ` text-anchor="end"`
				}

				bodyColor := s.TextColor
				if bodyColor == "" {
					bodyColor = "#000000"
				}

				for _, line := range lines {
					out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-weight="%s"%s%s>%s</text>
`, textX, startY, fs, bodyColor, fontWeight, textAnchorAttr, fontFamilyAttr, escapeXMLText(line)))
					startY += lineHeight
				}
			}
		} // end text mode

	case "ellipse":
		cx := x + w/2
		cy := y + h/2
		rx := w / 2
		ry := h / 2

		fill := s.Fill
		if fill == "" {
			fill = "none"
		}
		stroke := s.Stroke
		if stroke == "" && fill != "none" {
			stroke = darkenColor(fill)
		} else if stroke == "" && fill == "none" {
			stroke = "none"
		}

		// Emit gradient definition if present
		if s.GradientDef != nil {
			out.WriteString("  <defs>\n")
			out.WriteString(renderSVGGradientDef(s.GradientDef))
			out.WriteString("  </defs>\n")
			fill = fmt.Sprintf("url(#%s)", s.GradientDef.ID)
		}
		// Emit pattern definition if present
		if s.PatternFill != nil {
			pattID := fmt.Sprintf("patt_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(renderPatternDef(s.PatternFill, pattID))
			out.WriteString("  </defs>\n")
			fill = fmt.Sprintf("url(#%s)", pattID)
		} else if s.ImageData != "" && s.ImageMIME != "" {
			bgID := fmt.Sprintf("bgimg_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(fmt.Sprintf(`    <pattern id="%s" patternUnits="userSpaceOnUse" width="%d" height="%d">
      <image width="%d" height="%d" preserveAspectRatio="none" href="data:%s;base64,%s"/>
    </pattern>
`, bgID, w, h, w, h, s.ImageMIME, s.ImageData))
			out.WriteString("  </defs>\n")
			fill = fmt.Sprintf("url(#%s)", bgID)
		}

		dashAttr := prstDashToSVG(s.DashStyle, s.StrokeWidth, s.LineCap)
		ellipseSW := "1.33"
		if s.StrokeWidth > 0 {
			ellipseSW = fmt.Sprintf("%.2f", s.StrokeWidth)
		}
		lineCapAttr := ""
		if s.LineCap != "" {
			lineCapAttr = fmt.Sprintf(` stroke-linecap="%s"`, s.LineCap)
		}
		fillOpacityAttr := ""
		if s.FillOpacity > 0 && s.FillOpacity < 1.0 {
			fillOpacityAttr = fmt.Sprintf(` fill-opacity="%.2f"`, s.FillOpacity)
		}

		// Soft edge effect: PowerPoint "softEdge" fades the alpha along the shape edges.
		// For ellipses, approximate with a radial gradient from opaque center to transparent edges.
		if s.SoftEdgeRadius > 0 {
			avgRadius := float64(rx+ry) / 2.0
			// Fraction of the radius that remains fully opaque
			softFraction := s.SoftEdgeRadius / avgRadius
			if softFraction > 1.0 {
				softFraction = 1.0
			}
			opaqueStop := 1.0 - softFraction
			if opaqueStop < 0 {
				opaqueStop = 0
			}

			gradID := fmt.Sprintf("softEdge_%d", s.XmlID)
			out.WriteString(fmt.Sprintf(`  <defs><radialGradient id="%s" cx="50%%" cy="50%%" r="50%%">
    <stop offset="%.0f%%" stop-color="%s" stop-opacity="1"/>
    <stop offset="100%%" stop-color="%s" stop-opacity="0"/>
  </radialGradient></defs>
`, gradID, opaqueStop*100, fill, fill))
			fill = fmt.Sprintf("url(#%s)", gradID)
			fillOpacityAttr = "" // opacity handled by gradient stops
		}

		// Build transforms â€” full for geometry, rotation-only for text
		ellipseGeoTransform := buildShapeTransform(s, x, y, w, h)
		ellipseTextTransform := buildShapeRotationOnly(s, x, y, w, h)

		if s.Text != "" && (s.FlipH || s.FlipV) {
			// Shape has both geometry and text with flips: separate transforms
			if renderMode == renderModeAll || renderMode == renderModeGeo {
				if ellipseGeoTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", ellipseGeoTransform))
				}
				if stroke == "none" {
					out.WriteString(fmt.Sprintf(`  <ellipse cx="%d" cy="%d" rx="%d" ry="%d" fill="%s"%s/>
`, cx, cy, rx, ry, fill, fillOpacityAttr))
				} else {
					out.WriteString(fmt.Sprintf(`  <ellipse cx="%d" cy="%d" rx="%d" ry="%d" fill="%s"%s stroke="%s" stroke-width="%s"%s%s/>
`, cx, cy, rx, ry, fill, fillOpacityAttr, stroke, ellipseSW, dashAttr, lineCapAttr))
				}
				if ellipseGeoTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // geo mode
			if renderMode == renderModeAll || renderMode == renderModeText {
				// Text with rotation only (no flip)
				if ellipseTextTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", ellipseTextTransform))
				}
				if s.Text != "" {
					ensureMinDims(&tw, &th, s.Text)
					out.WriteString(renderShapeRichText(s, tx, ty, tw, th, fs, cfg))
				}
				if ellipseTextTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // text mode
		} else {
			// No flip+text conflict: wrap everything together
			if renderMode == renderModeAll || renderMode == renderModeGeo {
				if ellipseGeoTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", ellipseGeoTransform))
				}
				if stroke == "none" {
					out.WriteString(fmt.Sprintf(`  <ellipse cx="%d" cy="%d" rx="%d" ry="%d" fill="%s"%s/>
`, cx, cy, rx, ry, fill, fillOpacityAttr))
				} else {
					out.WriteString(fmt.Sprintf(`  <ellipse cx="%d" cy="%d" rx="%d" ry="%d" fill="%s"%s stroke="%s" stroke-width="%s"%s%s/>
`, cx, cy, rx, ry, fill, fillOpacityAttr, stroke, ellipseSW, dashAttr, lineCapAttr))
				}
				if ellipseGeoTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // end geo mode
			if renderMode == renderModeAll || renderMode == renderModeText {
				if s.Text != "" {
					if ellipseGeoTransform != "" {
						out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", ellipseGeoTransform))
					}
					ensureMinDims(&w, &h, s.Text)
					out.WriteString(renderShapeRichText(s, x, y, w, h, fs, cfg))
					if ellipseGeoTransform != "" {
						out.WriteString("  </g>\n")
					}
				}
			} // text mode
		}

	case "rect":
		// #region agent log
		if s.Text != "" {
			writeDebugLog("renderShapeSVG rect", map[string]interface{}{
				"name": s.Name, "ShapeType": s.ShapeType, "TextAnchor": s.TextAnchor,
				"BodyTIns": s.BodyTIns, "BodyBIns": s.BodyBIns,
			}, "H1_H5")
		}
		// #endregion
		if s.Text != "" {
			ensureMinDims(&w, &h, s.Text)
		}
		rx := 0
		txtColor := s.TextColor
		if txtColor == "" {
			txtColor = "#000000"
		}
		// Emit gradient definition if present
		fillOverride := s.Fill
		if s.GradientDef != nil {
			out.WriteString("  <defs>\n")
			out.WriteString(renderSVGGradientDef(s.GradientDef))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", s.GradientDef.ID)
		}
		// Emit pattern definition if present
		if s.PatternFill != nil {
			pattID := fmt.Sprintf("patt_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(renderPatternDef(s.PatternFill, pattID))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", pattID)
		} else if s.ImageData != "" && s.ImageMIME != "" {
			bgID := fmt.Sprintf("bgimg_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(fmt.Sprintf(`    <pattern id="%s" patternUnits="userSpaceOnUse" width="%d" height="%d">
      <image width="%d" height="%d" preserveAspectRatio="none" href="data:%s;base64,%s"/>
    </pattern>
`, bgID, w, h, w, h, s.ImageMIME, s.ImageData))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", bgID)
		}
		dashAttr := prstDashToSVG(s.DashStyle, s.StrokeWidth, s.LineCap)
		lineCapAttr := ""
		if s.LineCap != "" {
			lineCapAttr = fmt.Sprintf(` stroke-linecap="%s"`, s.LineCap)
		}
		// Build transforms â€” full transform for geometry, rotation-only for text.
		// PowerPoint flips shape geometry but keeps text readable.
		geoTransform := buildShapeTransform(s, x, y, w, h)
		textTransform := buildShapeRotationOnly(s, x, y, w, h)
		if s.Text != "" && (s.FlipH || s.FlipV) {
			// Shape has both geometry and text with flips:
			// apply full transform to geometry, rotation-only to text
			if renderMode == renderModeAll || renderMode == renderModeGeo {
				if geoTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", geoTransform))
				}
				out.WriteString(renderFilledShapeGeometryOnly(x, y, w, h, rx, fillOverride, s.Stroke, dashAttr, s.StrokeWidth, lineCapAttr, cfg))
				if geoTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // end geo
			if renderMode == renderModeAll || renderMode == renderModeText {
				if textTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", textTransform))
				}
				if s.Text != "" {
					ensureMinDims(&tw, &th, s.Text)
					out.WriteString(renderShapeRichText(s, tx, ty, tw, th, fs, cfg))
				}
				if textTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // end text
		} else {
			// No flip+text conflict: apply full transform to everything
			if renderMode == renderModeAll || renderMode == renderModeGeo {
				if geoTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", geoTransform))
				}
				out.WriteString(renderFilledShapeGeometryOnly(x, y, w, h, rx, fillOverride, s.Stroke, dashAttr, s.StrokeWidth, lineCapAttr, cfg))
				if geoTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // end geo
			if renderMode == renderModeAll || renderMode == renderModeText {
				if geoTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", geoTransform))
				}
				if s.Text != "" {
					ensureMinDims(&tw, &th, s.Text)
					out.WriteString(renderShapeRichText(s, tx, ty, tw, th, fs, cfg))
				}
				if geoTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // end text
		}

	case "roundRect":
		// #region agent log
		if s.Text != "" {
			writeDebugLog("renderShapeSVG roundRect", map[string]interface{}{
				"name": s.Name, "ShapeType": s.ShapeType, "TextAnchor": s.TextAnchor,
				"BodyTIns": s.BodyTIns, "BodyBIns": s.BodyBIns,
			}, "H1_H5")
		}
		// #endregion
		if s.Text != "" {
			ensureMinDims(&w, &h, s.Text)
		}
		txtColor := s.TextColor
		if txtColor == "" {
			txtColor = "#000000"
		}
		// Emit gradient definition if present
		fillOverride := s.Fill
		if s.GradientDef != nil {
			out.WriteString("  <defs>\n")
			out.WriteString(renderSVGGradientDef(s.GradientDef))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", s.GradientDef.ID)
		}
		// Emit pattern definition if present
		if s.PatternFill != nil {
			pattID := fmt.Sprintf("patt_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(renderPatternDef(s.PatternFill, pattID))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", pattID)
		} else if s.ImageData != "" && s.ImageMIME != "" {
			bgID := fmt.Sprintf("bgimg_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(fmt.Sprintf(`    <pattern id="%s" patternUnits="userSpaceOnUse" width="%d" height="%d">
      <image width="%d" height="%d" preserveAspectRatio="none" href="data:%s;base64,%s"/>
    </pattern>
`, bgID, w, h, w, h, s.ImageMIME, s.ImageData))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", bgID)
		}
		dashAttr := prstDashToSVG(s.DashStyle, s.StrokeWidth, s.LineCap)
		lineCapAttr := ""
		if s.LineCap != "" {
			lineCapAttr = fmt.Sprintf(` stroke-linecap="%s"`, s.LineCap)
		}
		// Build transforms â€” full for geometry, rotation-only for text
		geoTransform := buildShapeTransform(s, x, y, w, h)
		textTransform := buildShapeRotationOnly(s, x, y, w, h)
		if s.Text != "" && (s.FlipH || s.FlipV) {
			if renderMode == renderModeAll || renderMode == renderModeGeo {
				if geoTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", geoTransform))
				}
				out.WriteString(renderFilledShapeGeometryOnly(x, y, w, h, s.RoundRadius, fillOverride, s.Stroke, dashAttr, s.StrokeWidth, lineCapAttr, cfg))
				if geoTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // geo mode
			if renderMode == renderModeAll || renderMode == renderModeText {
				if textTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", textTransform))
				}
				if s.Text != "" {
					ensureMinDims(&tw, &th, s.Text)
					out.WriteString(renderShapeRichText(s, tx, ty, tw, th, fs, cfg))
				}
				if textTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // text mode
		} else {
			if renderMode == renderModeAll || renderMode == renderModeGeo {
				if geoTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", geoTransform))
				}
				out.WriteString(renderFilledShapeGeometryOnly(x, y, w, h, s.RoundRadius, fillOverride, s.Stroke, dashAttr, s.StrokeWidth, lineCapAttr, cfg))
				if geoTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // geo mode
			if renderMode == renderModeAll || renderMode == renderModeText {
				if geoTransform != "" {
					out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", geoTransform))
				}
				if s.Text != "" {
					ensureMinDims(&tw, &th, s.Text)
					out.WriteString(renderShapeRichText(s, tx, ty, tw, th, fs, cfg))
				}
				if geoTransform != "" {
					out.WriteString("  </g>\n")
				}
			} // text mode
		}

	case "stickyNote":
		stickyFS := fs
		stickyColor := s.TextColor
		if stickyColor == "" {
			stickyColor = "#000000"
		}
		out.WriteString(fmt.Sprintf(`  <g>
    <rect x="%d" y="%d" width="50" height="15" fill="#C00000" transform="rotate(3 %d %d)"/>
    <rect x="%d" y="%d" width="%d" height="%d" fill="url(#stickyGradient)" stroke="#808080" stroke-width="0.5"/>
    <text x="%d" y="%d" font-size="%d" fill="%s">%s</text>
  </g>
`, x+w/2-25, y-5, x+w/2, y, x, y+10, w, h-10, x+10, y+35, stickyFS, stickyColor, escapeXMLText(s.Text)))

	case "freeform":
		freeformDash := prstDashToSVG(s.DashStyle, s.StrokeWidth, s.LineCap)
		freeformSW := 1.33
		if s.StrokeWidth > 0 {
			freeformSW = s.StrokeWidth
		}
		freeformCap := ""
		if s.LineCap != "" {
			freeformCap = fmt.Sprintf(` stroke-linecap="%s"`, s.LineCap)
		}
		freeformFill := s.Fill
		if freeformFill == "" {
			freeformFill = "none"
		}
		// Emit gradient definition if present (custGeom shapes with gradFill)
		if s.GradientDef != nil {
			out.WriteString("  <defs>\n")
			out.WriteString(renderSVGGradientDef(s.GradientDef))
			out.WriteString("  </defs>\n")
			freeformFill = fmt.Sprintf("url(#%s)", s.GradientDef.ID)
		} else if s.PatternFill != nil {
			pattID := fmt.Sprintf("patt_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(renderPatternDef(s.PatternFill, pattID))
			out.WriteString("  </defs>\n")
			freeformFill = fmt.Sprintf("url(#%s)", pattID)
		} else if s.ImageData != "" && s.ImageMIME != "" {
			bgID := fmt.Sprintf("bgimg_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(fmt.Sprintf(`    <pattern id="%s" patternUnits="userSpaceOnUse" width="%d" height="%d">
      <image width="%d" height="%d" preserveAspectRatio="none" href="data:%s;base64,%s"/>
    </pattern>
`, bgID, w, h, w, h, s.ImageMIME, s.ImageData))
			out.WriteString("  </defs>\n")
			freeformFill = fmt.Sprintf("url(#%s)", bgID)
		}
		freeformStroke := s.Stroke
		if freeformStroke == "" {
			freeformStroke = "none"
		}

		// Build transform for rotation and flips
		freeformTransform := buildShapeTransform(s, x, y, w, h)
		freeformTransformAttr := ""
		if freeformTransform != "" {
			freeformTransformAttr = fmt.Sprintf(` transform="%s"`, freeformTransform)
		}

		if s.CustPathD != "" {
			out.WriteString(fmt.Sprintf(`  <path d="%s" fill="%s" stroke="%s" stroke-width="%.2f" stroke-linejoin="round"%s%s%s/>
`, s.CustPathD, freeformFill, freeformStroke, freeformSW, freeformDash, freeformCap, freeformTransformAttr))
		} else {
			out.WriteString(fmt.Sprintf(`  <path d="M%d %dC%d %d %d %d %d %dL%d %dL%d %dL%d %dC%d %d %d %d %d %dZ" fill="%s" stroke="%s" stroke-width="%.2f" stroke-linejoin="round"%s%s%s/>
`, x+w/4, y, x, y, x, y+h/4, x, y+h/3, x, y+h*2/3, x+w, y+h*2/3, x+w, y+h/3, x+w, y+h/4, x+w, y, x+w*3/4, y, freeformFill, freeformStroke, freeformSW, freeformDash, freeformCap, freeformTransformAttr))
		}
		if s.Text != "" {
			freeformColor := s.TextColor
			if freeformColor == "" {
				freeformColor = "#000000"
			}
			out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" text-anchor="middle" dominant-baseline="middle">%s</text>
`, x+w/2, y+h/2, fs, freeformColor, escapeXMLText(s.Text)))
		}

	case "prstGeomFilled":
		// Compound preset shapes (smiley, etc.) use PreRenderedSVG
		if s.PreRenderedSVG != "" {
			transform := buildShapeTransform(s, x, y, w, h)
			if transform != "" {
				out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", transform))
				out.WriteString(s.PreRenderedSVG)
				out.WriteString("  </g>\n")
			} else {
				out.WriteString(s.PreRenderedSVG)
			}
			break
		}
		// Render preset geometry shapes (arrows, triangles, cubes, moons, etc.) as SVG paths
		if s.CustPathD != "" && (renderMode == renderModeAll || renderMode == renderModeGeo) {
			fill := s.Fill
			if fill == "" {
				fill = "none"
			}
			stroke := s.Stroke
			if stroke == "" && fill != "none" {
				stroke = darkenColor(fill)
			}
			if stroke == "" || stroke == "none" {
				if fill == "none" {
					stroke = "#000000"
				}
			}

			// Emit gradient definition if present
			if s.GradientDef != nil {
				out.WriteString("  <defs>\n")
				out.WriteString(renderSVGGradientDef(s.GradientDef))
				out.WriteString("  </defs>\n")
				fill = fmt.Sprintf("url(#%s)", s.GradientDef.ID)
			} else if s.PatternFill != nil {
				pattID := fmt.Sprintf("patt_%d", s.XmlID)
				out.WriteString("  <defs>\n")
				out.WriteString(renderPatternDef(s.PatternFill, pattID))
				out.WriteString("  </defs>\n")
				fill = fmt.Sprintf("url(#%s)", pattID)
			} else if s.ImageData != "" && s.ImageMIME != "" {
				bgID := fmt.Sprintf("bgimg_%d", s.XmlID)
				out.WriteString("  <defs>\n")
				out.WriteString(fmt.Sprintf(`    <pattern id="%s" patternUnits="userSpaceOnUse" width="%d" height="%d">
      <image width="%d" height="%d" preserveAspectRatio="none" href="data:%s;base64,%s"/>
    </pattern>
`, bgID, w, h, w, h, s.ImageMIME, s.ImageData))
				out.WriteString("  </defs>\n")
				fill = fmt.Sprintf("url(#%s)", bgID)
			}

			// Build transform for rotation and flips
			transform := buildShapeTransform(s, x, y, w, h)
			transformAttr := ""
			if transform != "" {
				transformAttr = fmt.Sprintf(` transform="%s"`, transform)
			}
			textTransform := buildShapeRotationOnly(s, x, y, w, h)

			// Dash, stroke width, and line cap
			dashAttr := prstDashToSVG(s.DashStyle, s.StrokeWidth, s.LineCap)
			sw := 1.33
			if s.StrokeWidth > 0 {
				sw = s.StrokeWidth
			}
			lineCapAttr := ""
			if s.LineCap != "" {
				lineCapAttr = fmt.Sprintf(` stroke-linecap="%s"`, s.LineCap)
			}

			if stroke == "none" || stroke == "" {
				out.WriteString(fmt.Sprintf(`  <path d="%s" fill="%s" fill-rule="evenodd"%s/>
`, s.CustPathD, fill, transformAttr))
			} else {
				out.WriteString(fmt.Sprintf(`  <path d="%s" fill="%s" stroke="%s" stroke-width="%.2f" stroke-miterlimit="8" fill-rule="evenodd"%s%s%s/>
`, s.CustPathD, fill, stroke, sw, dashAttr, lineCapAttr, transformAttr))
			}

			if renderMode == renderModeAll || renderMode == renderModeText {
				if s.Text != "" {
					if textTransform != "" {
						out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", textTransform))
					}
					ensureMinDims(&tw, &th, s.Text)
					out.WriteString(renderShapeRichText(s, tx, ty, tw, th, fs, cfg))
					if textTransform != "" {
						out.WriteString("  </g>\n")
					}
				}
			}
		}

	case "table":
		if s.PreRenderedSVG != "" {
			out.WriteString(s.PreRenderedSVG)
		} else {
			// Fallback: render a placeholder box
			out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="#F5F5F5" stroke="#D9D9D9" stroke-width="1"/>
`, x, y, w, h))
			out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="14" fill="#999" text-anchor="middle">Table</text>
`, x+w/2, y+h/2))
		}

	case "chart":
		if s.PreRenderedSVG != "" {
			out.WriteString(s.PreRenderedSVG)
		} else {
			// Fallback: render a placeholder box
			out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="#F5F5F5" stroke="#D9D9D9" stroke-width="1"/>
`, x, y, w, h))
			out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="14" fill="#999" text-anchor="middle">Chart</text>
`, x+w/2, y+h/2))
		}

	case "picture":
		// Build transform for rotation and flips
		picTransform := buildShapeTransform(s, x, y, w, h)
		if picTransform != "" {
			out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", picTransform))
		}
		if s.ImageData != "" && s.ImageMIME != "" {
			out.WriteString(fmt.Sprintf(`  <image x="%d" y="%d" width="%d" height="%d" href="data:%s;base64,%s" preserveAspectRatio="none"/>
`, x, y, w, h, s.ImageMIME, s.ImageData))
		}

		if s.Stroke != "" && s.Stroke != "none" {
			dashAttr := prstDashToSVG(s.DashStyle, s.StrokeWidth, s.LineCap)
			lineCapAttr := ""
			if s.LineCap != "" {
				lineCapAttr = fmt.Sprintf(` stroke-linecap="%s"`, s.LineCap)
			}
			sw := "1.33"
			if s.StrokeWidth > 0 {
				sw = fmt.Sprintf("%.2f", s.StrokeWidth)
			}
			out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="none" stroke="%s" stroke-width="%s"%s%s/>
`, x, y, w, h, s.Stroke, sw, dashAttr, lineCapAttr))
		} else if s.ImageData == "" {
			out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="none" stroke="#4472C4" stroke-width="1.33" stroke-linejoin="round"/>
`, x, y, w, h))
		}
		if picTransform != "" {
			out.WriteString("  </g>\n")
		}

	case "textbox":
		// #region agent log
		if s.Text != "" {
			writeDebugLog("renderShapeSVG textbox", map[string]interface{}{
				"name": s.Name, "ShapeType": s.ShapeType, "TextAnchor": s.TextAnchor,
				"BodyTIns": s.BodyTIns, "BodyBIns": s.BodyBIns,
			}, "H1_H5")
		}
		// #endregion
		// Build transform for rotation only â€” PowerPoint flips the shape frame
		// but keeps text readable, so textboxes should never apply flipH/flipV.
		// When TextDirection is vertical, add 90Â° (or 270Â°) so text reads vertically.
		tbTransform := buildShapeRotationOnly(s, x, y, w, h)
		if tbTransform != "" {
			out.WriteString(fmt.Sprintf("  <g transform=\"%s\">\n", tbTransform))
		}
		// When a textbox has a fill (e.g. blue bar/gradient), draw the shape rect so structure matches expected SVG.
		fillOverride := s.Fill
		if s.GradientDef != nil {
			out.WriteString("  <defs>\n")
			out.WriteString(renderSVGGradientDef(s.GradientDef))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", s.GradientDef.ID)
		}
		if s.PatternFill != nil {
			pattID := fmt.Sprintf("patt_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(renderPatternDef(s.PatternFill, pattID))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", pattID)
		} else if s.ImageData != "" && s.ImageMIME != "" {
			bgID := fmt.Sprintf("bgimg_%d", s.XmlID)
			out.WriteString("  <defs>\n")
			out.WriteString(fmt.Sprintf(`    <pattern id="%s" patternUnits="userSpaceOnUse" width="%d" height="%d">
      <image width="%d" height="%d" preserveAspectRatio="none" href="data:%s;base64,%s"/>
    </pattern>
`, bgID, w, h, w, h, s.ImageMIME, s.ImageData))
			out.WriteString("  </defs>\n")
			fillOverride = fmt.Sprintf("url(#%s)", bgID)
		}
		if renderMode == renderModeAll || renderMode == renderModeGeo {
			if fillOverride == "" {
				fillOverride = "none"
			}
			tbStroke := "none"
			if s.Stroke != "" {
				tbStroke = s.Stroke
			}
			if fillOverride != "none" || tbStroke != "none" {
				dashAttr := prstDashToSVG(s.DashStyle, s.StrokeWidth, s.LineCap)
				lineCapAttr := ""
				if s.LineCap != "" {
					lineCapAttr = fmt.Sprintf(` stroke-linecap="%s"`, s.LineCap)
				}
				tbSW := "1.33"
				if s.StrokeWidth > 0 {
					tbSW = fmt.Sprintf("%.2f", s.StrokeWidth)
				}
				out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="%s" stroke="%s" stroke-width="%s"%s%s/>
`, x, y, w, h, fillOverride, tbStroke, tbSW, dashAttr, lineCapAttr))
			}
		} // end geo
		if renderMode == renderModeAll || renderMode == renderModeText {
			if s.Text != "" {
				ensureMinDims(&w, &h, s.Text)
				out.WriteString(renderShapeRichText(s, x, y, w, h, fs, cfg))
			}
			if tbTransform != "" {
				out.WriteString("  </g>\n")
			}
		} // text mode
	}

	return out.String()
}

// renderMultiParagraph renders structured paragraphs for body shapes.
func renderMultiParagraph(out *strings.Builder, s RenderShape, x, y, w, h, lInsPx, rInsPx, tInsPx, bInsPx, defaultFS int, fontWeight, fontFamilyAttr string, cfg *RenderConfig) {
	effectiveW := w - lInsPx - rInsPx
	if effectiveW < 40 {
		effectiveW = 40
	}
	effectiveH := h - tInsPx - bInsPx
	if effectiveH < defaultFS {
		effectiveH = defaultFS
	}

	// Calculate total height of all paragraphs
	totalHeight := 0
	type renderedLine struct {
		text          string
		fontSize      int
		lineHeight    int
		color         string
		fontW         string
		fontFam       string
		alignment     string
		marginLeft    int
		indent        int
		spaceBefore   int // extra vertical space before this line (first line of paragraph)
		spaceAfter    int // extra vertical space after this line (last line of paragraph)
		bulletText    string
		underline     string
		strikethrough string
		baselineShift int
	}
	var allLines []renderedLine

	bodyColor := s.TextColor
	if bodyColor == "" {
		bodyColor = "#000000"
	}

	// Counters for auto-numbering (level -> current count)
	autoNumCounters := make(map[int]int)

	for paraIdx, para := range s.Paragraphs {
		// In PowerPoint, the first paragraph's spaceBefore is typically not applied
		paraSpcBefore := para.SpaceBefore
		if paraIdx == 0 {
			paraSpcBefore = 0
		}
		totalHeight += paraSpcBefore

		paraFS := defaultFS
		paraColor := bodyColor
		paraFW := fontWeight
		paraFF := fontFamilyAttr
		paraAlign := para.Alignment
		if paraAlign == "" {
			paraAlign = s.Alignment
		}
		paraMarginLeft := para.MarginLeft

		// Determine bullet text
		bulletTxt := ""
		if para.BulletType == "char" {
			bulletTxt = para.BulletChar
		} else if para.BulletType == "autoNum" {
			// Reset deeper levels
			for l := para.Level + 1; l < 10; l++ {
				autoNumCounters[l] = 0
			}
			autoNumCounters[para.Level]++
			count := autoNumCounters[para.Level]
			if para.BulletStartAt > 1 && count == 1 {
				// Initialize count if startAt is specified (simplified logic)
				autoNumCounters[para.Level] = para.BulletStartAt
				count = para.BulletStartAt
			}

			// Simple numbering style map
			switch para.BulletAutoNumType {
			case "arabicPeriod":
				bulletTxt = fmt.Sprintf("%d.", count)
			case "arabicParenR":
				bulletTxt = fmt.Sprintf("%d)", count)
			case "alphaLcParenR":
				r := 'a' + (count-1)%26
				bulletTxt = fmt.Sprintf("%c)", r)
			case "alphaUcParenR":
				r := 'A' + (count-1)%26
				bulletTxt = fmt.Sprintf("%c)", r)
			default:
				bulletTxt = fmt.Sprintf("%d.", count)
			}
		}

		if para.IsEmpty {
			emptyLH := cfg.lineHeight(defaultFS)
			if para.LineSpacing > 0 {
				emptyLH = para.LineSpacing
			} else if para.LineSpacingPct > 0 {
				normalLH := float64(defaultFS) * 1.2
				emptyLH = int(math.Round(normalLH * float64(para.LineSpacingPct) / 100000.0))
			}
			totalHeight += emptyLH + para.SpaceAfter
			// Empty paragraph can still have a bullet!
			allLines = append(allLines, renderedLine{"", defaultFS, emptyLH, paraColor, paraFW, paraFF, paraAlign, 0, para.Indent, paraSpcBefore, para.SpaceAfter, bulletTxt, "", "", 0})
			continue
		}

		paraText := ""
		for _, run := range para.Runs {
			paraText += run.Text + " "
			if run.FontSize > 0 {
				paraFS = run.FontSize
			}
			if run.Color != "" {
				paraColor = run.Color
			}
			if run.Bold {
				paraFW = "bold"
			}
			if run.FontFamily != "" {
				paraFF = fmt.Sprintf(` font-family="%s"`, escapeXMLText(run.FontFamily))
			}
		}
		paraText = strings.TrimSpace(paraText)
		if paraText == "" {
			emptyLH := cfg.lineHeight(paraFS)
			totalHeight += emptyLH + para.SpaceAfter
			allLines = append(allLines, renderedLine{"", paraFS, emptyLH, paraColor, paraFW, paraFF, paraAlign, 0, para.Indent, paraSpcBefore, para.SpaceAfter, bulletTxt, "", "", 0})
			continue
		}

		wrapW := effectiveW - paraMarginLeft
		if wrapW < 40 {
			wrapW = 40
		}
		wrapPx := float64(wrapW)
		if !s.WrapText {
			wrapPx = 999999.0
		}
		var parsedFamily string
		var parsedBold bool
		var paraUnderline, paraStrikethrough string
		var paraBaselineShift int
		if len(para.Runs) > 0 {
			parsedFamily = para.Runs[0].FontFamily
			parsedBold = para.Runs[0].Bold
			paraUnderline = para.Runs[0].Underline
			paraStrikethrough = para.Runs[0].Strikethrough
			paraBaselineShift = para.Runs[0].BaselineShift
		}
		lines := wrapTextPx(paraText, wrapPx, parsedFamily, paraFS, parsedBold, false)

		paraLineHeight := cfg.lineHeight(paraFS)
		if para.LineSpacing > 0 {
			paraLineHeight = para.LineSpacing
		} else if para.LineSpacingPct > 0 {
			// spcPct is relative to "normal" line height (â‰ˆ fontSize Ã— 1.2)
			normalLH := float64(paraFS) * 1.2
			paraLineHeight = int(math.Round(normalLH * float64(para.LineSpacingPct) / 100000.0))
		}
		for i, line := range lines {
			spcB := 0
			bt := ""
			if i == 0 {
				spcB = paraSpcBefore // only first line of paragraph gets spaceBefore
				bt = bulletTxt       // only first line gets bullet
			}
			spcA := 0
			if i == len(lines)-1 {
				spcA = para.SpaceAfter // only last line of paragraph gets spaceAfter
			}
			allLines = append(allLines, renderedLine{line, paraFS, paraLineHeight, paraColor, paraFW, paraFF, paraAlign, paraMarginLeft, para.Indent, spcB, spcA, bt, paraUnderline, paraStrikethrough, paraBaselineShift})
			totalHeight += paraLineHeight
			if i == len(lines)-1 {
				totalHeight += para.SpaceAfter
			}
		}
	}

	// Compute vertical start position based on TextAnchor (from bodyPr in XML)
	anchor := s.TextAnchor
	if anchor == "" {
		anchor = "t"
	}

	// For centering, exclude trailing spaceAfter of the last paragraph
	// (trailing whitespace should not shift centered content upward)
	lastSpaceAfter := 0
	if len(allLines) > 0 {
		lastSpaceAfter = allLines[len(allLines)-1].spaceAfter
	}

	var startY int
	switch anchor {
	case "b":
		startY = y + h - bInsPx - totalHeight + defaultFS
	case "ctr":
		centeredHeight := totalHeight - lastSpaceAfter
		startY = y + tInsPx + (effectiveH-centeredHeight)/2 + defaultFS
	default:
		startY = y + tInsPx + defaultFS
	}

	// Render each line
	for _, line := range allLines {
		startY += line.spaceBefore // apply spaceBefore (first line of paragraph)

		// If line has bullet, render it
		if line.bulletText != "" {
			// Bullet X: marL is the content start. indent is the offset of the first line relative to marL.
			// Typically marL > 0 (e.g. 30px) and indent < 0 (e.g. -20px) => bullet at 10px, text at 30px.
			// We need to render bullet at (x + lInsPx + line.marginLeft + line.indent).
			// But note: indent is already in pixels.
			bx := x + lInsPx + line.marginLeft + line.indent

			// Adjust bullet alignment: usually right-aligned to the text start if hanging?
			// Let's assume left aligned at bx for now.

			out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-weight="%s"%s%s>%s</text>
`, bx, startY, line.fontSize, line.color, line.fontW, line.fontFam, "", escapeXMLText(line.bulletText)))
		}

		if line.text == "" {
			startY += line.lineHeight + line.spaceAfter
			continue
		}

		textX := x + lInsPx + line.marginLeft
		textAnchorAttr := ""
		align := line.alignment
		if align == "" {
			align = s.Alignment
		}
		if align == "ctr" {
			textX = x + w/2
			textAnchorAttr = ` text-anchor="middle"`
		} else if align == "r" {
			textX = x + w - rInsPx
			textAnchorAttr = ` text-anchor="end"`
		}

		textDecoAttr := buildTextDecorationAttr(line.underline, line.strikethrough)
		baselineAttr := ""
		lineFSActual := line.fontSize
		if line.baselineShift > 0 {
			lineFSActual = int(math.Round(float64(line.fontSize) * 0.6))
			if lineFSActual < 1 {
				lineFSActual = 1
			}
			baselineAttr = ` baseline-shift="super"`
		} else if line.baselineShift < 0 {
			lineFSActual = int(math.Round(float64(line.fontSize) * 0.6))
			if lineFSActual < 1 {
				lineFSActual = 1
			}
			baselineAttr = ` baseline-shift="sub"`
		}
		out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-weight="%s"%s%s%s%s>%s</text>
`, textX, startY, lineFSActual, line.color, line.fontW, textAnchorAttr, line.fontFam, textDecoAttr, baselineAttr, escapeXMLText(line.text)))
		startY += line.lineHeight + line.spaceAfter
	}
}

// paraLineHeight returns the vertical step for a title paragraph (from para props or config).
func paraLineHeight(para RenderParagraph, fs int, cfg *RenderConfig) int {
	lh := cfg.lineHeight(fs)
	if para.LineSpacing > 0 {
		lh = para.LineSpacing
	} else if para.LineSpacingPct > 0 {
		normalLH := float64(fs) * 1.2
		lh = int(math.Round(normalLH * float64(para.LineSpacingPct) / 100000.0))
	}
	return lh
}

// shapeParagraphLines returns one string per non-empty paragraph (Enter = new line). Empty if no paragraphs.
func shapeParagraphLines(s RenderShape) []string {
	var lines []string
	for _, p := range s.Paragraphs {
		if p.IsEmpty {
			continue
		}
		t := ""
		for _, r := range p.Runs {
			t += r.Text
		}
		t = strings.TrimSpace(t)
		if t != "" {
			lines = append(lines, t)
		}
	}
	return lines
}

// renderFilledShapeWithTextLines renders rect/roundRect geometry and one <text> per line (paragraphs on separate lines).
func renderFilledShapeWithTextLines(x, y, w, h, rx, fs int, fill, stroke string, lines []string, textColor string, dashAttr string, swPx float64, lineCapAttr string, textAnchor string, tInsPx, bInsPx int, cfg *RenderConfig) string {
	if len(lines) == 0 {
		sw := "1.33"
		if swPx > 0 {
			sw = fmt.Sprintf("%.2f", swPx)
		}
		return fmt.Sprintf(`  <g>
    <rect x="%d" y="%d" width="%d" height="%d" rx="%d" fill="%s" stroke="%s" stroke-width="%s"%s%s/>
  </g>
`, x, y, w, h, rx, fill, stroke, sw, dashAttr, lineCapAttr)
	}
	lineHeight := cfg.lineHeight(fs)
	totalTextHeight := len(lines) * lineHeight
	var startY int
	switch textAnchor {
	case "b":
		startY = y + h - bInsPx - totalTextHeight + fs
	case "t", "":
		startY = y + tInsPx + fs
	default:
		startY = y + (h-totalTextHeight)/2 + fs
	}
	if textColor == "" {
		textColor = "#000000"
	}
	sw := "1.33"
	if swPx > 0 {
		sw = fmt.Sprintf("%.2f", swPx)
	}
	var out strings.Builder
	out.WriteString(fmt.Sprintf(`  <g>
    <rect x="%d" y="%d" width="%d" height="%d" rx="%d" fill="%s" stroke="%s" stroke-width="%s"%s%s/>
`, x, y, w, h, rx, fill, stroke, sw, dashAttr, lineCapAttr))
	for _, line := range lines {
		out.WriteString(fmt.Sprintf(`    <text x="%d" y="%d" font-size="%d" fill="%s" text-anchor="middle" dominant-baseline="auto">%s</text>
`, x+w/2, startY, fs, textColor, escapeXMLText(line)))
		startY += lineHeight
	}
	out.WriteString("  </g>\n")
	return out.String()
}

// renderFilledShapeTextOnlyLines renders one <text> per line (paragraphs on separate lines), no rect.
func renderFilledShapeTextOnlyLines(x, y, w, h, fs int, lines []string, textColor string, textAnchor string, tInsPx, bInsPx int, cfg *RenderConfig) string {
	if len(lines) == 0 {
		return ""
	}
	if textColor == "" {
		textColor = "#000000"
	}
	lineHeight := cfg.lineHeight(fs)
	totalTextHeight := len(lines) * lineHeight
	var startY int
	switch textAnchor {
	case "b":
		startY = y + h - bInsPx - totalTextHeight + fs
	case "t", "":
		startY = y + tInsPx + fs
	default:
		startY = y + (h-totalTextHeight)/2 + fs
	}
	var out strings.Builder
	for _, line := range lines {
		out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" text-anchor="middle" dominant-baseline="auto">%s</text>
`, x+w/2, startY, fs, textColor, escapeXMLText(line)))
		startY += lineHeight
	}
	return out.String()
}

// renderFilledShapeWithText renders a filled rectangle or rounded rect with optional multi-line text.
// textAnchor and tInsPx/bInsPx come from bodyPr. When doWrap is false (case d), text is one line.
func renderFilledShapeWithText(x, y, w, h, rx, fs int, fill, stroke, text, textColor string, dashAttr string, swPx float64, lineCapAttr string, textAnchor string, tInsPx, bInsPx int, doWrap bool, cfg *RenderConfig) string {
	var out strings.Builder

	sw := "1.33"
	if swPx > 0 {
		sw = fmt.Sprintf("%.2f", swPx)
	}

	if text == "" {
		out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" rx="%d" fill="%s" stroke="%s" stroke-width="%s"%s%s/>
`, x, y, w, h, rx, fill, stroke, sw, dashAttr, lineCapAttr))
		return out.String()
	}

	maxCharsPerLine := cfg.maxCharsPerLine(w, fs)
	if !doWrap {
		maxCharsPerLine = 999999
	}
	lines := wrapText(text, maxCharsPerLine)
	lineHeight := cfg.lineHeight(fs)
	totalTextHeight := len(lines) * lineHeight

	effectiveH := h - tInsPx - bInsPx
	if effectiveH < fs {
		effectiveH = fs
	}
	var startY int
	var branch string
	switch textAnchor {
	case "b":
		startY = y + h - bInsPx - totalTextHeight + fs
		branch = "b"
	case "t", "": // "t" or omitted: OOXML default vertical alignment is top
		startY = y + tInsPx + fs
		branch = "t"
	default: // "ctr"
		startY = y + (h-totalTextHeight)/2 + fs
		branch = "ctr"
	}
	// #region agent log
	writeDebugLog("renderFilledShapeWithText", map[string]interface{}{
		"textAnchor": textAnchor, "startY": startY, "y": y, "tInsPx": tInsPx, "h": h,
		"totalTextHeight": totalTextHeight, "branch": branch,
	}, "H4_H5")
	// #endregion

	if textColor == "" {
		textColor = "#000000"
	}

	out.WriteString(fmt.Sprintf(`  <g>
    <rect x="%d" y="%d" width="%d" height="%d" rx="%d" fill="%s" stroke="%s" stroke-width="%s"%s%s/>
`, x, y, w, h, rx, fill, stroke, sw, dashAttr, lineCapAttr))
	for _, line := range lines {
		out.WriteString(fmt.Sprintf(`    <text x="%d" y="%d" font-size="%d" fill="%s" text-anchor="middle" dominant-baseline="auto">%s</text>
`, x+w/2, startY, fs, textColor, escapeXMLText(line)))
		startY += lineHeight
	}
	out.WriteString("  </g>\n")
	return out.String()
}

// renderFilledShapeGeometryOnly renders just the rect/roundRect geometry without text.
// Used when a flipped shape has text â€” the geometry gets the flip transform, text does not.
func renderFilledShapeGeometryOnly(x, y, w, h, rx int, fill, stroke string, dashAttr string, swPx float64, lineCapAttr string, cfg *RenderConfig) string {
	sw := "1.33"
	if swPx > 0 {
		sw = fmt.Sprintf("%.2f", swPx)
	}
	return fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" rx="%d" fill="%s" stroke="%s" stroke-width="%s"%s%s/>
`, x, y, w, h, rx, fill, stroke, sw, dashAttr, lineCapAttr)
}

// renderFilledShapeTextOnly renders just the text portion of a filled shape.
// When doWrap is false (case d), text is one line.
func renderFilledShapeTextOnly(x, y, w, h, fs int, text, textColor string, textAnchor string, tInsPx, bInsPx int, doWrap bool, cfg *RenderConfig) string {
	if text == "" {
		return ""
	}
	var out strings.Builder
	maxCharsPerLine := cfg.maxCharsPerLine(w, fs)
	if !doWrap {
		maxCharsPerLine = 999999
	}
	lines := wrapText(text, maxCharsPerLine)
	lineHeight := cfg.lineHeight(fs)
	totalTextHeight := len(lines) * lineHeight

	effectiveH := h - tInsPx - bInsPx
	if effectiveH < fs {
		effectiveH = fs
	}
	var startY int
	var branch string
	switch textAnchor {
	case "b":
		startY = y + h - bInsPx - totalTextHeight + fs
		branch = "b"
	case "t", "": // top or omitted (OOXML default)
		startY = y + tInsPx + fs
		branch = "t"
	default:
		startY = y + (h-totalTextHeight)/2 + fs
		branch = "ctr"
	}
	// #region agent log
	writeDebugLog("renderFilledShapeTextOnly", map[string]interface{}{
		"textAnchor": textAnchor, "startY": startY, "y": y, "tInsPx": tInsPx, "h": h,
		"totalTextHeight": totalTextHeight, "branch": branch,
	}, "H4_H5")
	// #endregion

	if textColor == "" {
		textColor = "#000000"
	}

	for _, line := range lines {
		out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" text-anchor="middle" dominant-baseline="auto">%s</text>
`, x+w/2, startY, fs, textColor, escapeXMLText(line)))
		startY += lineHeight
	}
	return out.String()
}

func ensureMinDims(w, h *int, text string) {
	// Only expand dimensions if they are zero or extremely small.
	// Do NOT override valid OOXML-defined dimensions â€” text should wrap to fit.
	if *w <= 0 {
		minW := 60
		if tw := len(text) * 7; tw > minW {
			minW = tw
		}
		*w = minW
	}
	if *h <= 0 {
		*h = 24
	}
}

// ======================== Connector Rendering ========================

// renderConnectorSVG renders a connector as SVG. Returns (gradientDefs, element).
// For gradient-filled straight connectors, a compound path (shaft + arrowhead) is returned
// along with a <linearGradient> definition. For all others, defs is empty.
func renderConnectorSVG(cn RenderConnector) (string, string) {
	x1 := emuToPx(cn.SX)
	y1 := emuToPx(cn.SY)
	x2 := emuToPx(cn.EX)
	y2 := emuToPx(cn.EY)

	if x1 == x2 && y1 == y2 {
		return "", ""
	}

	color := cn.Color
	if color == "" {
		color = "#4472C4"
	}
	arrow := ""
	if cn.HasArrow {
		arrow = ` marker-end="url(#arrowhead)"`
	}

	// Stroke width
	sw := cn.StrokeWidth
	if sw <= 0 {
		sw = 1.33
	}
	swAttr := fmt.Sprintf("%.2f", sw)

	// Dash style: map OOXML prstDash values to SVG stroke-dasharray
	dashAttr := prstDashToSVG(cn.DashStyle, sw, "")

	switch {
	case strings.HasPrefix(cn.GeomType, "bentConnector"):
		// Bent connector: 3-segment elbow through midpoint.
		adj := 50000
		if cn.AdjVal > 0 {
			adj = cn.AdjVal
		}
		frac := float64(adj) / 100000.0
		midX := x1 + int(float64(x2-x1)*frac)
		return "", fmt.Sprintf(`  <path d="M%d %d L%d %d L%d %d L%d %d" stroke="%s" stroke-width="%s" fill="none"%s%s/>
`, x1, y1, midX, y1, midX, y2, x2, y2, color, swAttr, dashAttr, arrow)

	case strings.HasPrefix(cn.GeomType, "curvedConnector"):
		// Curved connectors: use pre-computed, fully transformed bezier control points.
		if cn.HasCurveCP {
			cx1 := emuToPx(cn.CP1X)
			cy1 := emuToPx(cn.CP1Y)
			cx2 := emuToPx(cn.CP2X)
			cy2 := emuToPx(cn.CP2Y)
			return "", fmt.Sprintf(`  <path d="M%d %d C%d %d %d %d %d %d" stroke="%s" stroke-width="%s" fill="none"%s%s/>
`, x1, y1, cx1, cy1, cx2, cy2, x2, y2, color, swAttr, dashAttr, arrow)
		}
		// Fallback: midpoint S-curve heuristic.
		dx := intAbs(x2 - x1)
		dy := intAbs(y2 - y1)
		if dy > dx {
			midY := (y1 + y2) / 2
			return "", fmt.Sprintf(`  <path d="M%d %d C%d %d %d %d %d %d" stroke="%s" stroke-width="%s" fill="none"%s%s/>
`, x1, y1, x1, midY, x2, midY, x2, y2, color, swAttr, dashAttr, arrow)
		}
		midX := (x1 + x2) / 2
		return "", fmt.Sprintf(`  <path d="M%d %d C%d %d %d %d %d %d" stroke="%s" stroke-width="%s" fill="none"%s%s/>
`, x1, y1, midX, y1, midX, y2, x2, y2, color, swAttr, dashAttr, arrow)

	default:
		// Straight connector â€” render as compound path for arrows (gradient or solid)
		if cn.HasArrow {
			if len(cn.GradColors) >= 2 {
				return renderCompoundArrowConnector(cn, x1, y1, x2, y2)
			}
			// Solid arrow connector â€” also render as compound filled path
			// This matches expected SVG output and gives per-connector arrowhead color
			return renderCompoundArrowConnector(cn, x1, y1, x2, y2)
		}
		// No-arrow straight connector (plain line)
		return "", fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="%s"%s/>
`, x1, y1, x2, y2, color, swAttr, dashAttr)
	}
}

// renderCompoundArrowConnector renders a straight arrow connector as a compound filled path
// (shaft rectangle + arrowhead triangle). For gradient connectors, a <linearGradient> def is
// returned. For solid connectors, the fill is the connector's solid color.
// ArrowAtHead controls which end gets the arrowhead (head=start, tail=end).
func renderCompoundArrowConnector(cn RenderConnector, x1, y1, x2, y2 int) (string, string) {
	// Compute arrow dimensions based on w/len attributes
	// OOXML arrowhead sizes: sm=5, med=8, lg=13 (approximate px at 96 DPI)
	arrowHalfBase := 6.5 // default "med" half-base width
	arrowLength := 8.0   // default "med" length
	switch cn.ArrowW {
	case "sm":
		arrowHalfBase = 3.5
	case "lg":
		arrowHalfBase = 6.667
	}
	switch cn.ArrowLen {
	case "sm":
		arrowLength = 5.0
	case "lg":
		arrowLength = 8.0
	}

	// Shaft thickness = line stroke width
	halfW := cn.StrokeWidth / 2.0
	if halfW < 0.5 {
		halfW = 0.667 // minimum visible half-width
	}

	// Determine arrow direction: tail end (x1â†’x2, arrow at x2) or head end (arrow at x1)
	// For head-end arrows, we swap endpoints so the arrow geometry is built at the "end"
	// then the path naturally has the arrowhead at the original start (head) position.
	fx1, fy1, fx2, fy2 := float64(x1), float64(y1), float64(x2), float64(y2)
	if cn.ArrowAtHead {
		fx1, fy1, fx2, fy2 = fx2, fy2, fx1, fy1
	}

	// Direction vector (from shaft start toward arrow tip)
	dx := fx2 - fx1
	dy := fy2 - fy1
	length := math.Sqrt(dx*dx + dy*dy)
	if length < 1 {
		return "", ""
	}

	// Unit direction and perpendicular
	ux, uy := dx/length, dy/length // along arrow direction
	px, py := -uy, ux              // perpendicular (90Â° CCW)

	// Shaft rectangle: from start to (end - arrowLength), offset by halfW perpendicular
	shaftEndX := fx2 - ux*arrowLength
	shaftEndY := fy2 - uy*arrowLength

	// Four corners of the shaft rectangle
	s1x := fx1 + px*halfW
	s1y := fy1 + py*halfW
	s2x := shaftEndX + px*halfW
	s2y := shaftEndY + py*halfW
	s3x := shaftEndX - px*halfW
	s3y := shaftEndY - py*halfW
	s4x := fx1 - px*halfW
	s4y := fy1 - py*halfW

	// Arrowhead triangle: base at shaftEnd, tip at (fx2, fy2)
	a1x := shaftEndX + px*arrowHalfBase
	a1y := shaftEndY + py*arrowHalfBase
	a2x := fx2
	a2y := fy2
	a3x := shaftEndX - px*arrowHalfBase
	a3y := shaftEndY - py*arrowHalfBase

	// Build compound path (shaft rect + arrowhead triangle)
	pathD := fmt.Sprintf("M%.1f %.1f %.1f %.1f %.1f %.1f %.1f %.1fZ"+
		"M%.1f %.1f %.1f %.1f %.1f %.1fZ",
		s1x, s1y, s2x, s2y, s3x, s3y, s4x, s4y,
		a1x, a1y, a2x, a2y, a3x, a3y)

	// Gradient arrow: build linearGradient definition
	if len(cn.GradColors) >= 2 {
		gradID := fmt.Sprintf("grad_arrow_%d", cn.ShapeID)
		var gradDef strings.Builder
		// Gradient runs from original start to original end (before head/tail swap)
		gx1, gy1, gx2, gy2 := float64(x1), float64(y1), float64(x2), float64(y2)
		gradDef.WriteString(fmt.Sprintf("    <linearGradient id=\"%s\" x1=\"%.1f\" y1=\"%.1f\" x2=\"%.1f\" y2=\"%.1f\" gradientUnits=\"userSpaceOnUse\">\n",
			gradID, gx1, gy1, gx2, gy2))
		for i, c := range cn.GradColors {
			offset := 0.0
			if i < len(cn.GradPositions) {
				offset = float64(cn.GradPositions[i]) / 100000.0
			}
			gradDef.WriteString(fmt.Sprintf("      <stop offset=\"%.2f\" stop-color=\"%s\"/>\n", offset, c))
		}
		gradDef.WriteString("    </linearGradient>\n")

		element := fmt.Sprintf("  <path d=\"%s\" fill=\"url(#%s)\"/>\n", pathD, gradID)
		return gradDef.String(), element
	}

	// Solid arrow: fill with connector color
	color := cn.Color
	if color == "" {
		color = "#4472C4"
	}
	element := fmt.Sprintf("  <path d=\"%s\" fill=\"%s\"/>\n", pathD, color)
	return "", element
}

// ======================== Deduplication ========================

func deduplicateRenderShapes(shapes []RenderShape) []RenderShape {
	normalize := func(s string) string {
		return strings.TrimSpace(strings.Join(strings.Fields(s), " "))
	}
	priority := func(t string) int {
		switch t {
		case "title":
			return 5
		case "body":
			return 4
		case "rect":
			return 3
		case "roundRect":
			return 2
		case "ellipse":
			return 2
		case "textbox":
			return 1
		default:
			return 0
		}
	}
	const maxCenterDist = 80
	drop := make(map[int]bool)

	for i := range shapes {
		if drop[i] {
			continue
		}
		a := &shapes[i]
		textA := normalize(a.Text)
		if textA == "" {
			continue
		}
		ax, ay := emuToPx(a.X), emuToPx(a.Y)
		aw, ah := emuToPx(a.W), emuToPx(a.H)
		acx, acy := ax+aw/2, ay+ah/2
		for j := i + 1; j < len(shapes); j++ {
			if drop[j] {
				continue
			}
			b := &shapes[j]
			textB := normalize(b.Text)
			if textA != textB {
				continue
			}
			bx, by := emuToPx(b.X), emuToPx(b.Y)
			bw, bh := emuToPx(b.W), emuToPx(b.H)
			bcx, bcy := bx+bw/2, by+bh/2
			distSq := (acx-bcx)*(acx-bcx) + (acy-bcy)*(acy-bcy)
			overlap := (ax < bx+bw && bx < ax+aw && ay < by+bh && by < ay+ah) || distSq < maxCenterDist*maxCenterDist
			if !overlap {
				continue
			}
			if priority(a.ShapeType) >= priority(b.ShapeType) {
				drop[j] = true
			} else {
				drop[i] = true
				break
			}
		}
	}

	out := make([]RenderShape, 0, len(shapes))
	for i, s := range shapes {
		if !drop[i] {
			out = append(out, s)
		}
	}
	return out
}

// ======================== Helpers ========================

func emuToPx(emu int64) int {
	return int(math.Round(float64(emu) / emuPerPx))
}

// ptToPx converts typographic points to SVG pixels.
// In SVG, user units correspond to CSS px (1/96 of an inch).
// 1pt = 1/72 inch, 1px = 1/96 inch, so pt * 96/72 = pt * 4/3.
func ptToPx(pt int) int {
	return int(math.Round(float64(pt) * 96.0 / 72.0))
}

// ptToPxF converts typographic points (float) to SVG pixels.
func ptToPxF(pt float64) int {
	return int(math.Round(pt * 96.0 / 72.0))
}

func intAbs(x int) int {
	if x < 0 {
		return -x
	}
	return x
}

// splitRunTextForPlacement splits run text on " â€“ " so we emit one <text> per segment (matches expected SVG with separate SOLID FILL, â€“, caption).
func splitRunTextForPlacement(s string) []string {
	if s == "" {
		return nil
	}
	parts := strings.Split(s, " â€“ ")
	if len(parts) <= 1 {
		return []string{s}
	}
	var out []string
	for i, p := range parts {
		out = append(out, strings.TrimSpace(p))
		if i < len(parts)-1 {
			out = append(out, "â€“")
		}
	}
	return out
}

func escapeXMLText(s string) string {
	var buf bytes.Buffer
	xml.EscapeText(&buf, []byte(s))
	return buf.String()
}

// normAutofitScale returns a scale factor (0 < scale <= 1) for "Shrink Text on Overflow".
// Used when s.TextAutofit == "norm" to scale font so text fits in effectiveW x effectiveH.
func normAutofitScale(s RenderShape, effectiveW, effectiveH, fs int, cfg *RenderConfig) float64 {
	if effectiveW <= 0 || effectiveH <= 0 || fs <= 0 || s.Text == "" {
		return 1.0
	}
	maxChars := cfg.maxCharsPerLine(effectiveW, fs)
	lines := wrapText(s.Text, maxChars)
	if len(lines) == 0 {
		return 1.0
	}
	totalHeight := len(lines) * cfg.lineHeight(fs)
	maxLineChars := 0
	for _, line := range lines {
		if len(line) > maxLineChars {
			maxLineChars = len(line)
		}
	}
	charW := cfg.TextLayout.CharWidthRatio * float64(fs)
	if charW < 1 {
		charW = 1
	}
	naturalW := float64(maxLineChars) * charW
	scaleW := 1.0
	if naturalW > 0 {
		scaleW = float64(effectiveW) / naturalW
	}
	scaleH := 1.0
	if totalHeight > 0 {
		scaleH = float64(effectiveH) / float64(totalHeight)
	}
	scale := scaleW
	if scaleH < scale {
		scale = scaleH
	}
	if scale > 1.0 {
		scale = 1.0
	}
	// OOXML fontScale hint: when set, use as ceiling (don't scale above it)
	if s.NormFontScalePct > 0 && scale > s.NormFontScalePct/100 {
		scale = s.NormFontScalePct / 100
	}
	// Interpolation: don't shrink too much â€” use a floor so text stays readable.
	minScale := 0.5
	if s.NormFontScalePct > 0 {
		minScale = s.NormFontScalePct / 100
		if minScale > 1.0 {
			minScale = 1.0
		}
	}
	if scale < minScale {
		scale = minScale
	}
	return scale
}

func wrapText(text string, maxChars int) []string {
	words := strings.Fields(text)
	var lines []string
	var cur string
	for _, w := range words {
		if len(w) > maxChars {
			if cur != "" {
				lines = append(lines, cur)
				cur = ""
			}
			// Split long word into maxChars pieces
			for len(w) > maxChars {
				lines = append(lines, w[:maxChars])
				w = w[maxChars:]
			}
			cur = w
		} else if len(cur)+len(w)+1 <= maxChars {
			if cur != "" {
				cur += " "
			}
			cur += w
		} else {
			if cur != "" {
				lines = append(lines, cur)
			}
			cur = w
		}
	}
	if cur != "" {
		lines = append(lines, cur)
	}
	return lines
}

func wrapTextPx(text string, maxPxWidth float64, family string, fs int, bold, italic bool) []string {
	words := strings.Fields(text)
	var lines []string
	var cur string
	for _, w := range words {
		wPx := GetStringWidth(w, family, fs, bold, italic)

		if wPx > maxPxWidth {
			if cur != "" {
				lines = append(lines, cur)
				cur = ""
			}
			var wChunk string
			for _, r := range w {
				charExt := string(r)
				if GetStringWidth(wChunk+charExt, family, fs, bold, italic) > maxPxWidth && wChunk != "" {
					lines = append(lines, wChunk)
					wChunk = charExt
				} else {
					wChunk += charExt
				}
			}
			cur = wChunk
		} else {
			testLine := cur
			if testLine != "" {
				testLine += " "
			}
			testLine += w
			if GetStringWidth(testLine, family, fs, bold, italic) <= maxPxWidth {
				cur = testLine
			} else {
				if cur != "" {
					lines = append(lines, cur)
				}
				cur = w
			}
		}
	}
	if cur != "" {
		lines = append(lines, cur)
	}
	return lines
}