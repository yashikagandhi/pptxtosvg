package pptxtosvg

import (
	"bytes"
	"archive/zip"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"testing"
)

// ============================================================================
// Configuration
// ============================================================================

func resolveTestDataDir() string {
	if env := os.Getenv("PPTX_TEST_DATA_DIR"); env != "" {
		return env
	}
	candidates := []string{
		"testData",
		"../course_sample/testData",
		"../../course_sample/testData",
		"C:/LMS/course_sample/testData",
	}
	for _, c := range candidates {
		if fi, err := os.Stat(c); err == nil && fi.IsDir() {
			return c
		}
	}
	return ""
}

func resolveFontsDir() string {
	if env := os.Getenv("PPTX_FONTS_DIR"); env != "" {
		return env
	}
	candidates := []string{
		"fonts",
		"../backend/fonts",
		"../../backend/fonts",
		"C:/LMS/backend/fonts",
	}
	for _, c := range candidates {
		if fi, err := os.Stat(c); err == nil && fi.IsDir() {
			return c
		}
	}
	return ""
}

var (
	testDataDir = resolveTestDataDir()
	resultDir   = "../../testData_results"
	fontsDir    = resolveFontsDir()
)

// defaultTestFiles lists all test PPTX files available for regression testing.
var defaultTestFiles = []string{
	"test1", "test2", "test3", "test4", "test5",
	"test6", "test7", "test8", "test9", "test10", "test11",
}

// ============================================================================
// TestLocalConvert Ã¢â‚¬â€ Convert a single slide from a specific test file
// Replaces: pptx_local_test.go (original), extract_slide3.ps1
// ============================================================================

func TestLocalConvert(t *testing.T) {
	if testDataDir == "" {
		t.Skip("testDataDir not found")
	}
	if fontsDir != "" {
		LoadFontsOnce(fontsDir)
	}

	// Configure which test file and slide to convert
	testFile := "test11"
	slideNum := 15

	pptxPath := filepath.Join(testDataDir, testFile+".pptx")
	zr, err := zip.OpenReader(pptxPath)
	if err != nil {
		t.Fatalf("Cannot open %s: %v", pptxPath, err)
	}
	defer zr.Close()

	slideXMLPath := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
	slideXML, err := loadFileFromZip(&zr.Reader, slideXMLPath)
	if err != nil {
		t.Fatalf("Cannot load %s: %v", slideXMLPath, err)
	}

	svg := convertSlideToSVG(slideXML, slideNum, &zr.Reader)
	outPath := fmt.Sprintf("slide%d_out.svg", slideNum)
	if err := os.WriteFile(outPath, []byte(svg), 0644); err != nil {
		t.Fatal(err)
	}
	t.Logf("Wrote %s (%d bytes)", outPath, len(svg))
}

// ============================================================================
// TestConvertAllSlides Ã¢â‚¬â€ Convert ALL slides from a specific test file
// Replaces: test_convert.ps1, test_convert_test1.ps1, test_convert_all.ps1
// ============================================================================

func TestConvertAllSlides(t *testing.T) {
	if testDataDir == "" {
		t.Skip("testDataDir not found")
	}
	if fontsDir != "" {
		LoadFontsOnce(fontsDir)
	}

	// Configure which test files to convert (all by default)
	testFiles := defaultTestFiles

	slideRe := regexp.MustCompile(`slide(\d+)\.xml`)

	for _, testName := range testFiles {
		t.Run(testName, func(t *testing.T) {
			pptxPath := filepath.Join(testDataDir, testName+".pptx")
			zr, err := zip.OpenReader(pptxPath)
			if err != nil {
				t.Skipf("Cannot open %s: %v", pptxPath, err)
				return
			}
			defer zr.Close()

			// Find all slide files
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
				t.Fatalf("No slides found in %s", pptxPath)
			}

			// Create output directory
			outDir := filepath.Join(resultDir, testName)
			os.MkdirAll(outDir, 0755)

			count := 0
			for i := 1; i <= len(slideFiles); i++ {
				sf, ok := slideFiles[i]
				if !ok {
					continue
				}
				rc, err := sf.Open()
				if err != nil {
					t.Errorf("Slide %d: open error: %v", i, err)
					continue
				}
				data, err := io.ReadAll(rc)
				rc.Close()
				if err != nil {
					t.Errorf("Slide %d: read error: %v", i, err)
					continue
				}

				svg := convertSlideToSVG(data, i, &zr.Reader)
				outPath := filepath.Join(outDir, fmt.Sprintf("slide%d.svg", i))
				if err := os.WriteFile(outPath, []byte(svg), 0644); err != nil {
					t.Errorf("Slide %d: write error: %v", i, err)
					continue
				}
				count++
			}
			t.Logf("%s: %d SVGs generated in %s", testName, count, outDir)
		})
	}
}

// ============================================================================
// SVG Parsing types for regression comparison
// ============================================================================

type svgTextElement struct {
	Content    string
	X, Y       float64
	FontSize   float64
	Fill       string
	FontWeight string
	FontStyle  string
}

type svgRectElement struct {
	X, Y, W, H  float64
	Fill        string
	HasFilter   bool
	FillOpacity string
	SourceIndex int
}

type svgEllipseElement struct {
	CX, CY float64
	Fill   string
}

type svgPathElement struct {
	Fill, Stroke   string
	StartX, StartY float64
	HasTransform   bool
	TransformVal   string
	FillOpacity    string
}

type svgImageElement struct {
	X, Y, W, H float64
}

type svgTransformInfo struct {
	Total       int
	Rotations   []float64
	FlipH       int
	FlipV       int
	MatrixCount int
	RotateCount int
	ScaleCount  int
}

type svgParsedResult struct {
	Texts        []svgTextElement
	Rects        []svgRectElement
	Ellipses     []svgEllipseElement
	Paths        []svgPathElement
	ImageList    []svgImageElement
	Images       int
	Gradients    int
	Filters      int
	AllText      string
	UniqueColors map[string]bool
	Transforms   svgTransformInfo
	FlippedTexts int
}

// ============================================================================
// SVG Parser Ã¢â‚¬â€ extracts structural elements from SVG content
// Replaces: Parse-SvgElements function from regression_test.ps1
// ============================================================================

func parseSVGElements(svgContent string) svgParsedResult {
	result := svgParsedResult{
		UniqueColors: make(map[string]bool),
	}

	colorRe := regexp.MustCompile(`^#[0-9A-F]{6}$`)

	// -- Text elements --
	textRe := regexp.MustCompile(`<text\s+([^>]*)>([^<]*)</text>`)
	attrXRe := regexp.MustCompile(`\bx="([^"]*)"`)
	attrYRe := regexp.MustCompile(`\by="([^"]*)"`)
	attrFSRe := regexp.MustCompile(`font-size="([^"]*)"`)
	attrFillRe := regexp.MustCompile(`fill="([^"]*)"`)
	attrFWRe := regexp.MustCompile(`font-weight="([^"]*)"`)
	attrFStRe := regexp.MustCompile(`font-style="([^"]*)"`)
	matrixRe := regexp.MustCompile(`transform="matrix\(([^)]+)\)"`)

	for _, m := range textRe.FindAllStringSubmatch(svgContent, -1) {
		attrs, content := m[1], strings.TrimSpace(m[2])
		if content == "" {
			continue
		}
		t := svgTextElement{Content: content}

		if xm := attrXRe.FindStringSubmatch(attrs); xm != nil {
			t.X, _ = strconv.ParseFloat(xm[1], 64)
			t.X = math.Round(t.X)
		}
		if ym := attrYRe.FindStringSubmatch(attrs); ym != nil {
			t.Y, _ = strconv.ParseFloat(ym[1], 64)
			t.Y = math.Round(t.Y)
		}

		// Override with matrix transform
		if mm := matrixRe.FindStringSubmatch(attrs); mm != nil {
			parts := strings.Fields(mm[1])
			if len(parts) >= 6 {
				tx, _ := strconv.ParseFloat(parts[4], 64)
				ty, _ := strconv.ParseFloat(parts[5], 64)
				t.X = math.Round(tx)
				t.Y = math.Round(ty)
			}
		}

		if fsm := attrFSRe.FindStringSubmatch(attrs); fsm != nil {
			t.FontSize, _ = strconv.ParseFloat(fsm[1], 64)
			t.FontSize = math.Round(t.FontSize)
		}
		if fm := attrFillRe.FindStringSubmatch(attrs); fm != nil {
			t.Fill = strings.ToUpper(fm[1])
		}
		if fwm := attrFWRe.FindStringSubmatch(attrs); fwm != nil {
			t.FontWeight = fwm[1]
		}
		if fstm := attrFStRe.FindStringSubmatch(attrs); fstm != nil {
			t.FontStyle = fstm[1]
		}

		result.Texts = append(result.Texts, t)
		if colorRe.MatchString(t.Fill) {
			result.UniqueColors[t.Fill] = true
		}
	}

	// Concatenated text
	var allTexts []string
	for _, t := range result.Texts {
		allTexts = append(allTexts, t.Content)
	}
	result.AllText = strings.Join(allTexts, " ")

	// -- Rect elements --
	attrWRe := regexp.MustCompile(`width="([^"]*)"`)
	attrHRe := regexp.MustCompile(`height="([^"]*)"`)
	filterRe := regexp.MustCompile(`filter="`)
	fillOpRe := regexp.MustCompile(`fill-opacity="([^"]*)"`)

	rectRe := regexp.MustCompile(`<rect\s+([^>]*)/?>`)
	for _, m := range rectRe.FindAllStringSubmatchIndex(svgContent, -1) {
		attrs := svgContent[m[2]:m[3]]
		r := svgRectElement{FillOpacity: "1", SourceIndex: m[0]}

		if xm := attrXRe.FindStringSubmatch(attrs); xm != nil {
			r.X, _ = strconv.ParseFloat(xm[1], 64)
			r.X = math.Round(r.X)
		}
		if ym := attrYRe.FindStringSubmatch(attrs); ym != nil {
			r.Y, _ = strconv.ParseFloat(ym[1], 64)
			r.Y = math.Round(r.Y)
		}
		if wm := attrWRe.FindStringSubmatch(attrs); wm != nil {
			r.W, _ = strconv.ParseFloat(wm[1], 64)
			r.W = math.Round(r.W)
		}
		if hm := attrHRe.FindStringSubmatch(attrs); hm != nil {
			r.H, _ = strconv.ParseFloat(hm[1], 64)
			r.H = math.Round(r.H)
		}
		if fm := attrFillRe.FindStringSubmatch(attrs); fm != nil {
			r.Fill = strings.ToUpper(fm[1])
		}
		r.HasFilter = filterRe.MatchString(attrs)
		if fom := fillOpRe.FindStringSubmatch(attrs); fom != nil {
			r.FillOpacity = fom[1]
		}
		if colorRe.MatchString(r.Fill) {
			result.UniqueColors[r.Fill] = true
		}
		result.Rects = append(result.Rects, r)
	}

	// -- Ellipse elements --
	ellipseRe := regexp.MustCompile(`<ellipse\s+([^>]*)/?>`)
	cxRe := regexp.MustCompile(`cx="([^"]*)"`)
	cyRe := regexp.MustCompile(`cy="([^"]*)"`)

	for _, m := range ellipseRe.FindAllStringSubmatch(svgContent, -1) {
		attrs := m[1]
		e := svgEllipseElement{}
		if cxm := cxRe.FindStringSubmatch(attrs); cxm != nil {
			e.CX, _ = strconv.ParseFloat(cxm[1], 64)
			e.CX = math.Round(e.CX)
		}
		if cym := cyRe.FindStringSubmatch(attrs); cym != nil {
			e.CY, _ = strconv.ParseFloat(cym[1], 64)
			e.CY = math.Round(e.CY)
		}
		if fm := attrFillRe.FindStringSubmatch(attrs); fm != nil {
			e.Fill = strings.ToUpper(fm[1])
		}
		if colorRe.MatchString(e.Fill) {
			result.UniqueColors[e.Fill] = true
		}
		result.Ellipses = append(result.Ellipses, e)
	}

	// -- Path elements --
	pathRe := regexp.MustCompile(`<path\s+([^>]*)/?>`)
	strokeRe := regexp.MustCompile(`stroke="([^"]*)"`)
	dStartRe := regexp.MustCompile(`d="M\s*([0-9.\-]+)[\s,]+([0-9.\-]+)`)
	transformRe := regexp.MustCompile(`transform="([^"]*)"`)

	for _, m := range pathRe.FindAllStringSubmatch(svgContent, -1) {
		attrs := m[1]
		p := svgPathElement{}
		if fm := attrFillRe.FindStringSubmatch(attrs); fm != nil {
			p.Fill = strings.ToUpper(fm[1])
		}
		if sm := strokeRe.FindStringSubmatch(attrs); sm != nil {
			p.Stroke = strings.ToUpper(sm[1])
		}
		if fom := fillOpRe.FindStringSubmatch(attrs); fom != nil {
			p.FillOpacity = fom[1]
		}
		if dm := dStartRe.FindStringSubmatch(attrs); dm != nil {
			p.StartX, _ = strconv.ParseFloat(dm[1], 64)
			p.StartY, _ = strconv.ParseFloat(dm[2], 64)
			p.StartX = math.Round(p.StartX)
			p.StartY = math.Round(p.StartY)
		}
		if tfm := transformRe.FindStringSubmatch(attrs); tfm != nil {
			p.HasTransform = true
			p.TransformVal = tfm[1]
		}
		if colorRe.MatchString(p.Fill) {
			result.UniqueColors[p.Fill] = true
		}
		if colorRe.MatchString(p.Stroke) {
			result.UniqueColors[p.Stroke] = true
		}
		result.Paths = append(result.Paths, p)
	}

	// -- Image elements --
	imageRe := regexp.MustCompile(`<image\s+([^>]*)>`)
	for _, m := range imageRe.FindAllStringSubmatch(svgContent, -1) {
		attrs := m[1]
		img := svgImageElement{}
		if xm := attrXRe.FindStringSubmatch(attrs); xm != nil {
			img.X, _ = strconv.ParseFloat(xm[1], 64)
			img.X = math.Round(img.X)
		}
		if ym := attrYRe.FindStringSubmatch(attrs); ym != nil {
			img.Y, _ = strconv.ParseFloat(ym[1], 64)
			img.Y = math.Round(img.Y)
		}
		if wm := attrWRe.FindStringSubmatch(attrs); wm != nil {
			img.W, _ = strconv.ParseFloat(wm[1], 64)
			img.W = math.Round(img.W)
		}
		if hm := attrHRe.FindStringSubmatch(attrs); hm != nil {
			img.H, _ = strconv.ParseFloat(hm[1], 64)
			img.H = math.Round(img.H)
		}
		result.ImageList = append(result.ImageList, img)
	}
	result.Images = len(result.ImageList)

	// -- Counts --
	result.Gradients = len(regexp.MustCompile(`(linear|radial)Gradient`).FindAllString(svgContent, -1))
	result.Filters = len(regexp.MustCompile(`<filter\s+`).FindAllString(svgContent, -1))

	// -- Transform analysis --
	allTransRe := regexp.MustCompile(`<(\w+)\s+[^>]*transform="([^"]*)"`)
	rotateValRe := regexp.MustCompile(`rotate\(\s*([0-9.\-]+)`)
	matrixValRe := regexp.MustCompile(`matrix\(([^)]+)\)`)
	scaleValRe := regexp.MustCompile(`scale\(([^)]+)\)`)

	for _, m := range allTransRe.FindAllStringSubmatch(svgContent, -1) {
		tVal := m[2]
		result.Transforms.Total++

		if rm := rotateValRe.FindStringSubmatch(tVal); rm != nil {
			result.Transforms.RotateCount++
			angle, _ := strconv.ParseFloat(rm[1], 64)
			result.Transforms.Rotations = append(result.Transforms.Rotations, math.Round(angle*10)/10)
		}

		if mm := matrixValRe.FindStringSubmatch(tVal); mm != nil {
			result.Transforms.MatrixCount++
			parts := strings.FieldsFunc(mm[1], func(r rune) bool { return r == ' ' || r == ',' })
			if len(parts) >= 4 {
				a, _ := strconv.ParseFloat(parts[0], 64)
				b, _ := strconv.ParseFloat(parts[1], 64)
				// c, _ := strconv.ParseFloat(parts[2], 64)
				// Non-identity rotation
				if math.Abs(b) > 0.01 {
					angleDeg := math.Round(math.Atan2(b, a)*180/math.Pi*10) / 10
					result.Transforms.Rotations = append(result.Transforms.Rotations, angleDeg)
					result.Transforms.RotateCount++
				}
			}
		}

		if sm := scaleValRe.FindStringSubmatch(tVal); sm != nil {
			result.Transforms.ScaleCount++
			sParts := strings.FieldsFunc(sm[1], func(r rune) bool { return r == ' ' || r == ',' })
			if len(sParts) >= 1 {
				sx, _ := strconv.ParseFloat(sParts[0], 64)
				if sx < 0 {
					result.Transforms.FlipH++
				}
			}
			if len(sParts) >= 2 {
				sy, _ := strconv.ParseFloat(sParts[1], 64)
				if sy < 0 {
					result.Transforms.FlipV++
				}
			}
		}
	}

	// -- Detect flipped text --
	flipDepth := 0
	groupDepthStack := []int{}
	currentGroupDepth := 0
	tokenRe := regexp.MustCompile(`<(/?)(g|text)\b([^>]*)>`)
	flipScaleRe := regexp.MustCompile(`scale\(\s*-1`)
	flipScaleVRe := regexp.MustCompile(`scale\(\s*1\s*,\s*-1\)`)

	for _, m := range tokenRe.FindAllStringSubmatch(svgContent, -1) {
		isClose := m[1] == "/"
		tagName := m[2]
		attrs := m[3]

		if tagName == "g" {
			if !isClose {
				currentGroupDepth++
				if flipScaleRe.MatchString(attrs) || flipScaleVRe.MatchString(attrs) {
					flipDepth++
					groupDepthStack = append(groupDepthStack, currentGroupDepth)
				}
			} else {
				if len(groupDepthStack) > 0 && groupDepthStack[len(groupDepthStack)-1] == currentGroupDepth {
					flipDepth--
					groupDepthStack = groupDepthStack[:len(groupDepthStack)-1]
				}
				currentGroupDepth--
				if currentGroupDepth < 0 {
					currentGroupDepth = 0
				}
			}
		} else if tagName == "text" && !isClose {
			if flipDepth > 0 {
				result.FlippedTexts++
			}
		}
	}

	return result
}

// ============================================================================
// Comparison helpers Ã¢â‚¬â€ color distance, text coverage, color coverage
// Replaces: Get-ColorDistance, Get-TextCoverage, Get-ColorCoverage from regression_test.ps1
// ============================================================================

func colorDistance(c1, c2 string) int {
	if c1 == c2 {
		return 0
	}
	if c1 == "" || c2 == "" {
		return 999
	}
	a := strings.TrimPrefix(strings.ToUpper(c1), "#")
	b := strings.TrimPrefix(strings.ToUpper(c2), "#")
	if len(a) != 6 || len(b) != 6 {
		if a == b {
			return 0
		}
		return 999
	}
	r1, _ := strconv.ParseInt(a[0:2], 16, 64)
	g1, _ := strconv.ParseInt(a[2:4], 16, 64)
	b1v, _ := strconv.ParseInt(a[4:6], 16, 64)
	r2, _ := strconv.ParseInt(b[0:2], 16, 64)
	g2, _ := strconv.ParseInt(b[2:4], 16, 64)
	b2v, _ := strconv.ParseInt(b[4:6], 16, 64)

	dr := int(math.Abs(float64(r1 - r2)))
	dg := int(math.Abs(float64(g1 - g2)))
	db := int(math.Abs(float64(b1v - b2v)))
	maxD := dr
	if dg > maxD {
		maxD = dg
	}
	if db > maxD {
		maxD = db
	}
	return maxD
}

func textCoverage(expTexts, genTexts []svgTextElement) float64 {
	expWords := make(map[string]bool)
	for _, t := range expTexts {
		for _, w := range strings.Fields(t.Content) {
			w = strings.TrimSpace(w)
			if len(w) > 1 {
				expWords[w] = true
			}
		}
	}
	if len(expWords) == 0 {
		return 100.0
	}
	var genAll strings.Builder
	for _, t := range genTexts {
		genAll.WriteString(t.Content)
		genAll.WriteString(" ")
	}
	genStr := genAll.String()
	found := 0
	for w := range expWords {
		if strings.Contains(genStr, w) {
			found++
		}
	}
	return math.Round(float64(found)/float64(len(expWords))*1000) / 10
}

func colorCoverage(expColors, genColors map[string]bool, tolerance int) float64 {
	if len(expColors) == 0 {
		return 100.0
	}
	matched := 0
	for ec := range expColors {
		if genColors[ec] {
			matched++
			continue
		}
		for gc := range genColors {
			if colorDistance(ec, gc) <= tolerance {
				matched++
				break
			}
		}
	}
	return math.Round(float64(matched)/float64(len(expColors))*1000) / 10
}

// ============================================================================
// Regression issue tracking
// ============================================================================

type regressionIssue struct {
	Level string // "ERROR", "WARN", "INFO"
	Msg   string
}

// ============================================================================
// compareTextElements Ã¢â‚¬â€ content-based matching with position check
// Replaces: Compare-TextElements from regression_test.ps1
// ============================================================================

func compareTextElements(expTexts, genTexts []svgTextElement, posTol, fsTol, colorTol int) []regressionIssue {
	var issues []regressionIssue
	matched := make(map[int]bool)

	for _, exp := range expTexts {
		bestIdx := -1
		for i, gen := range genTexts {
			if matched[i] {
				continue
			}
			if gen.Content == exp.Content {
				bestIdx = i
				break
			}
		}
		if bestIdx < 0 {
			continue
		}
		matched[bestIdx] = true
		gen := genTexts[bestIdx]

		// Position comparison
		if exp.X > 0 && exp.Y > 0 && gen.X > 0 && gen.Y > 0 {
			xD := math.Abs(exp.X - gen.X)
			yD := math.Abs(exp.Y - gen.Y)
			if xD > float64(posTol) || yD > float64(posTol) {
				issues = append(issues, regressionIssue{
					Level: "WARN",
					Msg:   fmt.Sprintf("TEXT POS '%s': expected (%.0f,%.0f) got (%.0f,%.0f) delta=(%.0f,%.0f)", exp.Content, exp.X, exp.Y, gen.X, gen.Y, xD, yD),
				})
			}
		}

		// Font size comparison
		if exp.FontSize > 0 && gen.FontSize > 0 {
			fsDiff := math.Abs(exp.FontSize - gen.FontSize)
			if fsDiff > float64(fsTol) {
				issues = append(issues, regressionIssue{
					Level: "WARN",
					Msg:   fmt.Sprintf("FONT SIZE '%s': expected %.0f got %.0f", exp.Content, exp.FontSize, gen.FontSize),
				})
			}
		}

		// Color comparison
		if exp.Fill != "" && gen.Fill != "" {
			cd := colorDistance(exp.Fill, gen.Fill)
			if cd > colorTol {
				issues = append(issues, regressionIssue{
					Level: "WARN",
					Msg:   fmt.Sprintf("TEXT COLOR '%s': expected %s got %s (delta=%d)", exp.Content, exp.Fill, gen.Fill, cd),
				})
			}
		}

		// Font style comparison
		if exp.FontStyle != "" && exp.FontStyle != gen.FontStyle {
			issues = append(issues, regressionIssue{
				Level: "WARN",
				Msg:   fmt.Sprintf("FONT STYLE '%s': expected '%s' got '%s'", exp.Content, exp.FontStyle, gen.FontStyle),
			})
		}
	}
	return issues
}

// ============================================================================
// comparePathElements Ã¢â‚¬â€ color-matched position check for paths
// Replaces: Compare-PathElements from regression_test.ps1
// ============================================================================

func comparePathElements(expPaths, genPaths []svgPathElement, posTol, colorTol int) []regressionIssue {
	var issues []regressionIssue
	matched := make(map[int]bool)
	pathPosTol := float64(posTol * 5)

	for _, exp := range expPaths {
		if exp.Fill == "" || strings.ToUpper(exp.Fill) == "NONE" {
			continue
		}
		if exp.StartX == 0 && exp.StartY == 0 {
			continue
		}

		bestIdx := -1
		bestDist := 9999.0
		for i, gen := range genPaths {
			if matched[i] {
				continue
			}
			cd := colorDistance(exp.Fill, gen.Fill)
			if cd <= colorTol {
				posDist := math.Abs(exp.StartX-gen.StartX) + math.Abs(exp.StartY-gen.StartY)
				if posDist < bestDist {
					bestDist = posDist
					bestIdx = i
				}
			}
		}
		if bestIdx < 0 {
			continue
		}
		matched[bestIdx] = true
		gen := genPaths[bestIdx]

		if gen.StartX > 0 && gen.StartY > 0 {
			xD := math.Abs(exp.StartX - gen.StartX)
			yD := math.Abs(exp.StartY - gen.StartY)
			if xD > pathPosTol || yD > pathPosTol {
				issues = append(issues, regressionIssue{
					Level: "WARN",
					Msg:   fmt.Sprintf("PATH POS (fill=%s): expected start (%.0f,%.0f) got (%.0f,%.0f) delta=(%.0f,%.0f) - possible rotation/flip issue", exp.Fill, exp.StartX, exp.StartY, gen.StartX, gen.StartY, xD, yD),
				})
			}
		}
	}
	return issues
}

// ============================================================================
// compareSlide Ã¢â‚¬â€ full slide comparison, returns issues and metrics
// Replaces: Compare-Slide from regression_test.ps1
// ============================================================================

type slideMetrics struct {
	TextCoverage    float64
	ColorCoverage   float64
	StructuralScore float64
	FlippedTexts    int
}

func compareSlide(exp, gen svgParsedResult, posTol, fsTol, colorTol int) ([]regressionIssue, slideMetrics) {
	var issues []regressionIssue
	var m slideMetrics

	// Text coverage
	m.TextCoverage = textCoverage(exp.Texts, gen.Texts)

	// Structural score
	expTotal := len(exp.Texts) + len(exp.Rects) + len(exp.Ellipses) + len(exp.Paths) + exp.Images
	genTotal := len(gen.Texts) + len(gen.Rects) + len(gen.Ellipses) + len(gen.Paths) + gen.Images
	if expTotal > 0 {
		m.StructuralScore = math.Round(float64(genTotal)/float64(expTotal)*1000) / 10
	} else {
		m.StructuralScore = 100
	}

	if m.StructuralScore < 50 {
		lvl := "ERROR"
		if gen.Filters >= exp.Filters && gen.Filters > 0 {
			lvl = "WARN"
		}
		issues = append(issues, regressionIssue{Level: lvl, Msg: fmt.Sprintf("STRUCTURAL SCORE %.1f%% -- generated has far fewer visual elements than expected (%d vs %d)", m.StructuralScore, genTotal, expTotal)})
	} else if m.StructuralScore < 70 {
		issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("STRUCTURAL SCORE %.1f%% -- significant element deficit (%d vs %d)", m.StructuralScore, genTotal, expTotal)})
	}

	// Per-element-type checks
	type countPair struct {
		name string
		genC int
		expC int
	}
	counts := []countPair{
		{"text", len(gen.Texts), len(exp.Texts)},
		{"rect", len(gen.Rects), len(exp.Rects)},
		{"ellipse", len(gen.Ellipses), len(exp.Ellipses)},
		{"path", len(gen.Paths), len(exp.Paths)},
		{"image", gen.Images, exp.Images},
		{"gradient", gen.Gradients, exp.Gradients},
		{"filter", gen.Filters, exp.Filters},
	}

	for _, c := range counts {
		diff := c.genC - c.expC
		absDiff := int(math.Abs(float64(diff)))

		if c.expC >= 3 && c.genC < int(float64(c.expC)*0.6) {
			pct := int(math.Round(float64(c.genC) / float64(c.expC) * 100))
			if c.name == "text" && m.TextCoverage >= 90 {
				issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("%s count: expected %d, got %d (%d%% -- text coverage %.1f%% indicates rendering style difference)", c.name, c.expC, c.genC, pct, m.TextCoverage)})
			} else if c.name == "image" && c.genC == 0 && (len(gen.Rects)+len(gen.Paths)) >= (len(exp.Rects)+len(exp.Paths)) {
				issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("%s count: expected %d, got %d (vector rendering uses rects/paths instead of raster images)", c.name, c.expC, c.genC)})
			} else if (c.name == "rect" || c.name == "path") && gen.Filters >= exp.Filters && gen.Filters > 0 {
				issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("%s count: expected %d, got %d (deficit likely due to shadow rendering)", c.name, c.expC, c.genC)})
			} else {
				issues = append(issues, regressionIssue{Level: "ERROR", Msg: fmt.Sprintf("%s count: expected %d, got %d (only %d%% -- major deficit)", c.name, c.expC, c.genC, pct)})
			}
		} else if c.name == "text" && absDiff > 3 {
			issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("%s count: expected %d, got %d (%+d)", c.name, c.expC, c.genC, diff)})
		} else if c.name != "text" && absDiff > 5 {
			issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("%s count: expected %d, got %d (%+d)", c.name, c.expC, c.genC, diff)})
		}
	}

	// Text coverage check
	if m.TextCoverage < 50 {
		issues = append(issues, regressionIssue{Level: "ERROR", Msg: fmt.Sprintf("Text coverage only %.1f%% (many expected words missing)", m.TextCoverage)})
	} else if m.TextCoverage < 80 {
		issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("Text coverage %.1f%% (some expected words missing)", m.TextCoverage)})
	}

	// Color coverage
	m.ColorCoverage = colorCoverage(exp.UniqueColors, gen.UniqueColors, colorTol)
	if m.ColorCoverage < 30 {
		issues = append(issues, regressionIssue{Level: "ERROR", Msg: fmt.Sprintf("Color coverage only %.1f%% (critical: most expected colors missing)", m.ColorCoverage)})
	} else if m.ColorCoverage < 50 {
		issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("Color coverage only %.1f%% (many expected colors missing)", m.ColorCoverage)})
	}

	// Text element matching
	issues = append(issues, compareTextElements(exp.Texts, gen.Texts, posTol, fsTol, colorTol)...)

	// Path position comparison
	issues = append(issues, comparePathElements(exp.Paths, gen.Paths, posTol, colorTol)...)

	// Rotation check
	expRealRots := filterRotations(exp.Transforms.Rotations, 1.0)
	genRealRots := filterRotations(gen.Transforms.Rotations, 1.0)
	if len(expRealRots) > 0 && len(genRealRots) == 0 {
		issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("ROTATION MISSING: expected %d rotated element(s), got 0", len(expRealRots))})
	} else if len(expRealRots) > len(genRealRots) {
		issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("ROTATION COUNT: expected %d rotated elements, got %d", len(expRealRots), len(genRealRots))})
	}

	// Flip checks
	if exp.Transforms.FlipH > 0 && gen.Transforms.FlipH == 0 {
		issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("FLIP-H MISSING: expected %d horizontally flipped element(s), got 0", exp.Transforms.FlipH)})
	}
	flipHDelta := gen.Transforms.FlipH - exp.Transforms.FlipH
	if flipHDelta > 0 {
		lvl := "WARN"
		if flipHDelta >= 3 {
			lvl = "ERROR"
		}
		issues = append(issues, regressionIssue{Level: lvl, Msg: fmt.Sprintf("FLIP-H EXCESS: expected %d flipH, got %d - %d extra flip(s)", exp.Transforms.FlipH, gen.Transforms.FlipH, flipHDelta)})
	}
	if exp.Transforms.FlipV > 0 && gen.Transforms.FlipV == 0 {
		issues = append(issues, regressionIssue{Level: "WARN", Msg: fmt.Sprintf("FLIP-V MISSING: expected %d vertically flipped element(s), got 0", exp.Transforms.FlipV)})
	}
	flipVDelta := gen.Transforms.FlipV - exp.Transforms.FlipV
	if flipVDelta > 0 {
		lvl := "WARN"
		if flipVDelta >= 3 {
			lvl = "ERROR"
		}
		issues = append(issues, regressionIssue{Level: lvl, Msg: fmt.Sprintf("FLIP-V EXCESS: expected %d flipV, got %d - %d extra flip(s)", exp.Transforms.FlipV, gen.Transforms.FlipV, flipVDelta)})
	}

	// Flipped text (always a bug)
	m.FlippedTexts = gen.FlippedTexts
	if gen.FlippedTexts > 0 {
		issues = append(issues, regressionIssue{Level: "ERROR", Msg: fmt.Sprintf("FLIPPED TEXT: %d text element(s) inside flip transform - text should never be flipped", gen.FlippedTexts)})
	}

	// Content occlusion check
	if len(gen.Rects) > 1 {
		for ri := 1; ri < len(gen.Rects); ri++ {
			rect := gen.Rects[ri]
			if !rect.HasFilter || rect.W < 100 || rect.H < 50 {
				continue
			}
			if strings.ToUpper(rect.Fill) != "WHITE" && rect.Fill != "#FFFFFF" {
				continue
			}
			coveredTexts := 0
			for _, txt := range gen.Texts {
				if txt.X >= rect.X && txt.X <= rect.X+rect.W && txt.Y >= rect.Y && txt.Y <= rect.Y+rect.H {
					coveredTexts++
				}
			}
			lvl := "WARN"
			if coveredTexts >= 2 {
				lvl = "ERROR"
			}
			issues = append(issues, regressionIssue{Level: lvl, Msg: fmt.Sprintf("CONTENT OCCLUSION: Filter rect with opaque white fill (%.0fx%.0f at %.0f,%.0f) overlaps %d text element(s)", rect.W, rect.H, rect.X, rect.Y, coveredTexts)})
		}
	}

	return issues, m
}

func filterRotations(rotations []float64, minAbs float64) []float64 {
	var out []float64
	for _, r := range rotations {
		if math.Abs(r) > minAbs {
			out = append(out, r)
		}
	}
	return out
}

// ============================================================================
// TestRegressionSuite Ã¢â‚¬â€ Full regression test comparing generated SVGs to baselines
// Replaces: regression_test.ps1, run_regression_docker.ps1
// ============================================================================

func TestRegressionSuite(t *testing.T) {
	if testDataDir == "" {
		t.Skip("testDataDir not found")
	}
	if fontsDir != "" {
		LoadFontsOnce(fontsDir)
	}

	const posTolerance = 15
	const fontSizeTolerance = 3
	const colorTolerance = 30

	testNames := defaultTestFiles
	slideRe := regexp.MustCompile(`slide(\d+)\.xml`)

	totalSlides := 0
	passedSlides := 0
	warnSlides := 0
	failedSlides := 0

	for _, testName := range testNames {
		t.Run(testName, func(t *testing.T) {
			pptxPath := filepath.Join(testDataDir, testName+".pptx")
			expDir := filepath.Join(testDataDir, testName)

			// Check if expected SVGs exist
			if _, err := os.Stat(expDir); os.IsNotExist(err) {
				t.Skipf("No expected SVGs directory: %s", expDir)
				return
			}

			// Open PPTX and convert all slides
			zr, err := zip.OpenReader(pptxPath)
			if err != nil {
				t.Skipf("Cannot open %s: %v", pptxPath, err)
				return
			}
			defer zr.Close()

			// Find all slides
			slideFiles := make(map[int]*zip.File)
			for _, f := range zr.File {
				if strings.HasPrefix(f.Name, "ppt/slides/slide") && strings.HasSuffix(f.Name, ".xml") {
					if m := slideRe.FindStringSubmatch(f.Name); len(m) == 2 {
						num, _ := strconv.Atoi(m[1])
						slideFiles[num] = f
					}
				}
			}

			// Find expected SVG files
			expEntries, err := os.ReadDir(expDir)
			if err != nil {
				t.Fatalf("Cannot read expected dir: %v", err)
			}

			expSVGs := make(map[int]string)
			slideNumRe := regexp.MustCompile(`(?i)slide(\d+)\.svg`)
			for _, e := range expEntries {
				if m := slideNumRe.FindStringSubmatch(e.Name()); m != nil {
					num, _ := strconv.Atoi(m[1])
					expSVGs[num] = filepath.Join(expDir, e.Name())
				}
			}

			// Sort slide numbers
			var slideNums []int
			for n := range expSVGs {
				slideNums = append(slideNums, n)
			}
			sort.Ints(slideNums)

			for _, slideNum := range slideNums {
				t.Run(fmt.Sprintf("slide%d", slideNum), func(t *testing.T) {
					totalSlides++

					// Read expected SVG
					expContent, err := os.ReadFile(expSVGs[slideNum])
					if err != nil {
						t.Fatalf("Cannot read expected SVG: %v", err)
					}

					// Generate SVG
					sf, ok := slideFiles[slideNum]
					if !ok {
						failedSlides++
						t.Fatalf("Slide %d not found in PPTX", slideNum)
					}
					rc, err := sf.Open()
					if err != nil {
						failedSlides++
						t.Fatalf("Cannot open slide %d: %v", slideNum, err)
					}
					data, err := io.ReadAll(rc)
					rc.Close()
					if err != nil {
						failedSlides++
						t.Fatalf("Cannot read slide %d: %v", slideNum, err)
					}

					genSVG := convertSlideToSVG(data, slideNum, &zr.Reader)

					// Parse both SVGs
					expParsed := parseSVGElements(string(expContent))
					genParsed := parseSVGElements(genSVG)

					// Compare
					issues, metrics := compareSlide(expParsed, genParsed, posTolerance, fontSizeTolerance, colorTolerance)

					errors := 0
					warns := 0
					for _, issue := range issues {
						switch issue.Level {
						case "ERROR":
							errors++
						case "WARN":
							warns++
						}
					}

					// Build summary line
					summary := fmt.Sprintf("textCov=%.1f%% colorCov=%.1f%% struct=%.1f%% | txt=%d/%d rect=%d/%d ell=%d/%d path=%d/%d img=%d/%d",
						metrics.TextCoverage, metrics.ColorCoverage, metrics.StructuralScore,
						len(genParsed.Texts), len(expParsed.Texts),
						len(genParsed.Rects), len(expParsed.Rects),
						len(genParsed.Ellipses), len(expParsed.Ellipses),
						len(genParsed.Paths), len(expParsed.Paths),
						genParsed.Images, expParsed.Images,
					)

					if errors > 0 {
						failedSlides++
						t.Errorf("FAIL %s", summary)
						for _, issue := range issues {
							if issue.Level == "ERROR" {
								t.Errorf("  [ERROR] %s", issue.Msg)
							}
						}
						for _, issue := range issues {
							if issue.Level == "WARN" {
								t.Logf("  [WARN] %s", issue.Msg)
							}
						}
					} else if warns > 0 {
						warnSlides++
						t.Logf("WARN %s", summary)
						for _, issue := range issues {
							if issue.Level == "WARN" {
								t.Logf("  [WARN] %s", issue.Msg)
							}
						}
					} else {
						passedSlides++
						t.Logf("PASS %s", summary)
					}
				})
			}
		})
	}

	t.Logf("\n=== REGRESSION SUMMARY ===")
	t.Logf("Total: %d  PASSED: %d  WARNINGS: %d  FAILED: %d", totalSlides, passedSlides, warnSlides, failedSlides)
}

// ============================================================================
// TestVerifySpecificContent Ã¢â‚¬â€ Verify specific content in generated SVGs
// Replaces: verify_test2_green_placeholder.ps1
// ============================================================================

func TestVerifyGreenPlaceholder(t *testing.T) {
	if testDataDir == "" {
		t.Skip("testDataDir not found")
	}
	if fontsDir != "" {
		LoadFontsOnce(fontsDir)
	}

	pptxPath := filepath.Join(testDataDir, "test2.pptx")
	zr, err := zip.OpenReader(pptxPath)
	if err != nil {
		t.Skipf("Cannot open %s: %v", pptxPath, err)
		return
	}
	defer zr.Close()

	slideXML, err := loadFileFromZip(&zr.Reader, "ppt/slides/slide1.xml")
	if err != nil {
		t.Fatal(err)
	}
	svg := convertSlideToSVG(slideXML, 1, &zr.Reader)

	// Check for green (#87C389) Placeholder text
	greenPlaceholderRe := regexp.MustCompile(`fill="#87C389"[^>]*>[^<]*Placeholder|<text[^>]*fill="#87C389"[^>]*>Placeholder`)
	placeholderCountRe := regexp.MustCompile(`>Placeholder</text>`)

	count := len(placeholderCountRe.FindAllString(svg, -1))
	hasGreen := greenPlaceholderRe.MatchString(svg)

	t.Logf("Text elements with 'Placeholder': %d", count)
	if count < 3 {
		t.Errorf("Expected at least 3 Placeholder text elements, got %d", count)
	}

	if hasGreen {
		t.Logf("Green (#87C389) Placeholder: FOUND")
	} else {
		t.Errorf("Green (#87C389) Placeholder: NOT FOUND")
	}
}

// ============================================================================
// TestExtractSlideXML Ã¢â‚¬â€ Extract and pretty-print slide XML for debugging
// Replaces: extract_slide3.ps1, temp_s3.ps1, temp_s4.ps1, temp_s4b.ps1,
//           temp_analyze.ps1, extract_table_styles.ps1, extract_tablestyles.ps1
// ============================================================================

func TestExtractSlideXML(t *testing.T) {
	// Configure which test file and slides to extract
	testFile := "test10"
	slideNums := []int{16}

	pptxPath := filepath.Join(testDataDir, testFile+".pptx")
	zr, err := zip.OpenReader(pptxPath)
	if err != nil {
		t.Fatalf("Cannot open %s: %v", pptxPath, err)
	}
	defer zr.Close()

	for _, slideNum := range slideNums {
		slidePath := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
		data, err := loadFileFromZip(&zr.Reader, slidePath)
		if err != nil {
			t.Errorf("Cannot load %s: %v", slidePath, err)
			continue
		}

		// Pretty-print the XML
		formatted := prettyPrintXML(data)
		outPath := fmt.Sprintf("%s_slide%d.xml", testFile, slideNum)
		if err := os.WriteFile(outPath, []byte(formatted), 0644); err != nil {
			t.Errorf("Cannot write %s: %v", outPath, err)
			continue
		}
		t.Logf("Extracted %s (%d bytes) -> %s", slidePath, len(formatted), outPath)
	}

	// Also extract slide relationships
	for _, slideNum := range slideNums {
		relsPath := fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", slideNum)
		data, err := loadFileFromZip(&zr.Reader, relsPath)
		if err != nil {
			continue // rels may not exist
		}
		formatted := prettyPrintXML(data)
		outPath := fmt.Sprintf("%s_slide%d_rels.xml", testFile, slideNum)
		os.WriteFile(outPath, []byte(formatted), 0644)
		t.Logf("Extracted %s -> %s", relsPath, outPath)
	}
}

// TestExtractTableStyles extracts table styles XML for debugging
func TestExtractTableStyles(t *testing.T) {
	testFile := "test5"
	pptxPath := filepath.Join(testDataDir, testFile+".pptx")
	zr, err := zip.OpenReader(pptxPath)
	if err != nil {
		t.Skipf("Cannot open %s: %v", pptxPath, err)
		return
	}
	defer zr.Close()

	// Extract tableStyles.xml
	data, err := loadFileFromZip(&zr.Reader, "ppt/tableStyles.xml")
	if err != nil {
		t.Logf("No tableStyles.xml found")
	} else {
		formatted := prettyPrintXML(data)
		t.Logf("=== tableStyles.xml ===\n%s", formatted)
	}

	// List all entries matching "slide"
	t.Logf("=== Slide entries ===")
	for _, f := range zr.File {
		if strings.Contains(f.Name, "slide") {
			t.Logf("  %s", f.Name)
		}
	}
}

// TestInspectConnectors inspects connector and shape data for debugging
// Replaces: check_connectors.ps1, check_shapes2.ps1
func TestInspectConnectors(t *testing.T) {
	testFile := "test2"
	slideNum := 1

	pptxPath := filepath.Join(testDataDir, testFile+".pptx")
	zr, err := zip.OpenReader(pptxPath)
	if err != nil {
		t.Skipf("Cannot open %s: %v", pptxPath, err)
		return
	}
	defer zr.Close()

	slidePath := fmt.Sprintf("ppt/slides/slide%d.xml", slideNum)
	data, err := loadFileFromZip(&zr.Reader, slidePath)
	if err != nil {
		t.Fatalf("Cannot load %s: %v", slidePath, err)
	}

	var slide OOXMLSlide
	if err := xml.Unmarshal(data, &slide); err != nil {
		t.Fatalf("Cannot parse slide XML: %v", err)
	}

	// Report connectors
	t.Logf("=== Connectors ===")
	for _, cxn := range slide.CSld.SpTree.Connectors {
		name := ""
		if cxn.NvCxnSpPr.CNvPr.Name != "" {
			name = cxn.NvCxnSpPr.CNvPr.Name
		}
		off := cxn.SpPr.Xfrm.Off
		ext := cxn.SpPr.Xfrm.Ext
		t.Logf("  Connector '%s': off=(%d,%d) ext=(%d,%d)",
			name, off.X, off.Y, ext.Cx, ext.Cy)
	}

	// Report shapes without explicit fills
	t.Logf("=== Shapes without explicit fill ===")
	for _, sp := range slide.CSld.SpTree.Shapes {
		hasFill := sp.SpPr.SolidFill != nil || sp.SpPr.GradFill != nil
		hasNoFill := sp.SpPr.NoFill != nil

		if !hasFill && !hasNoFill {
			name := sp.NvSpPr.CNvPr.Name
			off := sp.SpPr.Xfrm.Off
			ext := sp.SpPr.Xfrm.Ext
			geom := ""
			if sp.SpPr.PrstGeom != nil {
				geom = sp.SpPr.PrstGeom.Prst
			}
			if sp.SpPr.CustGeom != nil {
				geom = "custGeom"
			}
			t.Logf("  Shape '%s': off=(%d,%d) ext=(%d,%d) geom=%s",
				name, off.X, off.Y, ext.Cx, ext.Cy, geom)
		}
	}
}

// ============================================================================
// prettyPrintXML Ã¢â‚¬â€ format XML with indentation for debugging
// ============================================================================

func prettyPrintXML(data []byte) string {
	decoder := xml.NewDecoder(strings.NewReader(string(data)))
	var buf strings.Builder
	encoder := xml.NewEncoder(&buf)
	encoder.Indent("", "  ")

	for {
		token, err := decoder.Token()
		if err != nil {
			break
		}
		if err := encoder.EncodeToken(token); err != nil {
			break
		}
	}
	encoder.Flush()
	return buf.String()
}

// ============================================================================
// TestListZipEntries Ã¢â‚¬â€ List all entries in a PPTX file for debugging
// ============================================================================

func TestListZipEntries(t *testing.T) {
	testFile := "test10"
	pptxPath := filepath.Join(testDataDir, testFile+".pptx")
	zr, err := zip.OpenReader(pptxPath)
	if err != nil {
		t.Fatalf("Cannot open %s: %v", pptxPath, err)
	}
	defer zr.Close()

	t.Logf("=== %s entries ===", testFile)
	for _, f := range zr.File {
		t.Logf("  %s (%d bytes)", f.Name, f.UncompressedSize64)
	}
}


func TestConvertPPTXBytes_Invalid(t *testing.T) {
	_, err := ConvertPPTXBytes([]byte("not a zip"))
	if err == nil {
		t.Fatal("expected error for invalid zip bytes, got nil")
	}
}

func TestConvertPPTXBytes_EmptyZip(t *testing.T) {
	buf := new(bytes.Buffer)
	zw := zip.NewWriter(buf)
	zw.Close()
	_, err := ConvertPPTXBytes(buf.Bytes())
	if err == nil {
		t.Fatal("expected error for empty zip (no slides), got nil")
	}
}