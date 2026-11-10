package pptxtosvg

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"path"
	"sort"
	"strconv"
	"strings"
)

// ====================================================================
// PPTX Chart Renderer
//
// Parses chart XML from OOXML and renders clustered bar charts as SVG.
// ====================================================================

// ChartData holds parsed chart information ready for rendering.
type ChartData struct {
	ChartType      string // "bar", "line", "pie", "doughnut"
	Title          string
	Categories     []string
	Series         []SeriesData
	HasLegend      bool
	LegendPos      string // "b","t","l","r"
	HasGrid        bool
	HasGapWidth    bool
	GapWidth       float64 // gap width percentage (e.g., 2.19)
	HasOverlap     bool
	Overlap        float64           // overlap percentage (e.g., -0.27)
	TitleFontSize  int               // title font size in px (from chart XML); 0 = use default
	ThemeColors    map[string]string // theme color scheme for accent color lookups
	LabelFontSize  int               // axis/legend label font size in px; 0 = use default
	ShowPercent    bool              // pie chart: show percentage labels
	FirstSliceAng  float64           // pie chart: starting angle in degrees (default 0)
	Explosion      float64           // pie chart: explosion distance (0-100, % of radius)
	BgColor        string            // chart background colour (empty = default light grey)
	BgIsGradient   bool
	BgGradIsRadial bool
	BgGradStops    []GradientStop
	DPtFills       []DataPointFill // per-data-point fill overrides (indexed by data point)
	HoleSize       float64         // doughnut chart: hole size percentage (0-100, default 50)
	CustomColors   []string        // parsed from chart color styles (colors*.xml)
}

// SeriesData holds a single chart series.
type SeriesData struct {
	Type         string
	Name         string
	Values       []float64
	XValues      []float64
	Color        string // resolved hex color
	HasLine      bool
	LineColor    string
	HasShadow    bool
	Shadow       *RenderShadow
	Grouping     string // e.g. "standard", "stacked", "percentStacked", "clustered"
	IsGradient   bool
	GradIsRadial bool
	GradAngle    float64
	GradStops    []GradientStop

	LineIsGradient   bool
	LineGradIsRadial bool
	LineGradAngle    float64
	LineGradStops    []GradientStop
}

// GradientStop represents a color stop in a gradient fill.
type GradientStop struct {
	Pos   float64 // position from 0.0 to 1.0
	Color string  // hex color
}

// DataPointFill holds the fill information for a single data point (pie slice).
type DataPointFill struct {
	IsSolid    bool   // true = solid fill, false = pattern fill
	SolidColor string // hex color for solid fill
	// Pattern fill fields
	PatternPreset string // OOXML preset name: "ltUpDiag", "ltDnDiag", etc.
	FgColor       string // foreground color of pattern
	BgColor       string // background color of pattern
	// Gradient fill fields
	IsGradient   bool
	GradIsRadial bool
	GradAngle    float64
	GradStops    []GradientStop
	// Raster Image fill fields
	HasBlipFill  bool
	BlipEmbedRID string
	// Effects (Shadow)
	HasShadow bool
	Shadow    *RenderShadow
	// Stroke override from dPt
	StrokeColor string  // stroke color (from dPt <a:ln>)
	StrokeWidth float64 // stroke width in px (from dPt <a:ln w="...">)

	LineIsGradient   bool
	LineGradIsRadial bool
	LineGradAngle    float64
	LineGradStops    []GradientStop
}

// getChartAccentColors returns accent colors from theme, falling back to Office defaults.
func getChartAccentColors(themeColors map[string]string) []string {
	return []string{
		resolveSchemeColor("accent1", themeColors),
		resolveSchemeColor("accent2", themeColors),
		resolveSchemeColor("accent3", themeColors),
		resolveSchemeColor("accent4", themeColors),
		resolveSchemeColor("accent5", themeColors),
		resolveSchemeColor("accent6", themeColors),
	}
}

// parseChartXML parses a chart XML file and returns structured chart data.
func parseChartXML(chartXMLBytes []byte, themeColors map[string]string) (*ChartData, error) {
	var cs ChartSpace
	if err := xml.Unmarshal(chartXMLBytes, &cs); err != nil {
		return nil, fmt.Errorf("chart XML parse error: %w", err)
	}

	data := &ChartData{
		Title:       "Chart Title", // PowerPoint default auto-title
		HasGrid:     true,
		GapWidth:    2.19,  // default 219%
		Overlap:     -0.27, // default -27%
		ThemeColors: themeColors,
	}

	pa := cs.Chart.PlotArea

	// Detect chart type and select series source
	type typedSeries struct {
		Type     string
		Ser      ChartSeries
		Grouping string
		Order    int
	}
	var seriesList []typedSeries

	if len(pa.BarCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "bar"
		}
		bc := pa.BarCharts[0]
		if bc.GapWidth.Val != "" {
			if v, err := strconv.ParseFloat(bc.GapWidth.Val, 64); err == nil {
				data.GapWidth = v / 100.0
				data.HasGapWidth = true
			}
		}
		if bc.Overlap.Val != "" {
			if v, err := strconv.ParseFloat(bc.Overlap.Val, 64); err == nil {
				data.Overlap = v / 100.0
				data.HasOverlap = true
			}
		}
		for _, bcInst := range pa.BarCharts {
			for _, s := range bcInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"bar", s, bcInst.Grouping.Val, order})
			}
		}
	}

	if len(pa.LineCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "line"
		}
		for _, lcInst := range pa.LineCharts {
			for _, s := range lcInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"line", s, lcInst.Grouping.Val, order})
			}
		}
	}

	if len(pa.PieCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "pie"
		}
		pc := pa.PieCharts[0]
		if pc.FirstSliceAng.Val != "" {
			if v, err := strconv.ParseFloat(pc.FirstSliceAng.Val, 64); err == nil {
				data.FirstSliceAng = v
			}
		}
		if pc.DLbls != nil && pc.DLbls.ShowPercent.Val == "1" {
			data.ShowPercent = true
		}
		for _, pcInst := range pa.PieCharts {
			for _, s := range pcInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"pie", s, "", order})
			}
		}
		if len(seriesList) > 0 && seriesList[0].Ser.DLbls != nil {
			if seriesList[0].Ser.DLbls.ShowPercent.Val == "1" {
				data.ShowPercent = true
			}
		}
		if len(seriesList) > 0 && seriesList[0].Ser.Explosion.Val != "" {
			if v, err := strconv.ParseFloat(seriesList[0].Ser.Explosion.Val, 64); err == nil {
				data.Explosion = v
			}
		}
	}

	if len(pa.DoughnutCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "doughnut"
		}
		dc := pa.DoughnutCharts[0]
		data.HoleSize = 50
		if dc.HoleSize.Val != "" {
			if v, err := strconv.ParseFloat(dc.HoleSize.Val, 64); err == nil {
				data.HoleSize = v
			}
		}
		if dc.FirstSliceAng.Val != "" {
			if v, err := strconv.ParseFloat(dc.FirstSliceAng.Val, 64); err == nil {
				data.FirstSliceAng = v
			}
		}
		if dc.DLbls != nil && dc.DLbls.ShowPercent.Val == "1" {
			data.ShowPercent = true
		}
		for _, dcInst := range pa.DoughnutCharts {
			for _, s := range dcInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"doughnut", s, "", order})
			}
		}
		if len(seriesList) > 0 && seriesList[0].Ser.DLbls != nil {
			if seriesList[0].Ser.DLbls.ShowPercent.Val == "1" {
				data.ShowPercent = true
			}
		}
	}

	if len(pa.AreaCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "area"
		}
		for _, cInst := range pa.AreaCharts {
			for _, s := range cInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"area", s, cInst.Grouping.Val, order})
			}
		}
	}

	if len(pa.ScatterCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "scatter"
		}
		for _, cInst := range pa.ScatterCharts {
			for _, s := range cInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"scatter", s, "", order})
			}
		}
	}

	if len(pa.StockCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "stock"
		}
		for _, cInst := range pa.StockCharts {
			for _, s := range cInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"stock", s, "", order})
			}
		}
	}

	if len(pa.SurfaceCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "surface"
		}
		for _, cInst := range pa.SurfaceCharts {
			for _, s := range cInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"surface", s, "", order})
			}
		}
	}

	if len(pa.Surface3DCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "surface3D"
		}
		for _, cInst := range pa.Surface3DCharts {
			for _, s := range cInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"surface3D", s, "", order})
			}
		}
	}

	if len(pa.RadarCharts) > 0 {
		if data.ChartType == "" {
			data.ChartType = "radar"
		}
		for _, cInst := range pa.RadarCharts {
			for _, s := range cInst.Series {
				order, _ := strconv.Atoi(s.Order.Val)
				seriesList = append(seriesList, typedSeries{"radar", s, "", order})
			}
		}
	}

	if len(seriesList) == 0 {
		return nil, fmt.Errorf("unsupported chart type in chart XML")
	}

	// Sort the aggregated series exactly by their PPTX order property so that
	// fallback colors appropriately map to matching dataset keys sequentially.
	sort.SliceStable(seriesList, func(i, j int) bool {
		return seriesList[i].Order < seriesList[j].Order
	})

	// Update chart type to combo if more than one distinct type is present
	distinctTypes := make(map[string]bool)
	for _, sl := range seriesList {
		distinctTypes[sl.Type] = true
	}
	if len(distinctTypes) > 1 {
		data.ChartType = "combo"
	}

	// Parse categories from first series (if exist)
	for _, sl := range seriesList {
		if sl.Ser.Cat != nil && sl.Ser.Cat.StrRef != nil && len(sl.Ser.Cat.StrRef.StrCache.Pts) > 0 {
			for _, pt := range sl.Ser.Cat.StrRef.StrCache.Pts {
				data.Categories = append(data.Categories, pt.V)
			}
			break
		}
	}

	// Parse series
	for i, sl := range seriesList {
		sd := SeriesData{Name: "Series", Type: sl.Type, Grouping: sl.Grouping}
		// Series name
		if len(sl.Ser.Tx.StrRef.StrCache.Pts) > 0 {
			sd.Name = sl.Ser.Tx.StrRef.StrCache.Pts[0].V
		}

		// Series color
		sd.Color = resolveSeriesColor(sl.Ser, i, themeColors)

		// Check for series-level shadows
		if sl.Ser.SpPr != nil && sl.Ser.SpPr.EffectLst != nil {
			sd.Shadow = extractShapeShadow(sl.Ser.SpPr.EffectLst, themeColors)
			if sd.Shadow != nil {
				sd.HasShadow = true
			}
		}

		// Check for series-level gradient fills
		if sl.Ser.SpPr != nil && sl.Ser.SpPr.GradFill != nil {
			sd.IsGradient = true
			if sl.Ser.SpPr.GradFill.Path != nil {
				sd.GradIsRadial = true
			} else if sl.Ser.SpPr.GradFill.Lin != nil {
				sd.GradAngle = float64(sl.Ser.SpPr.GradFill.Lin.Ang) / 60000.0
			}
			for _, gs := range sl.Ser.SpPr.GradFill.GsLst.Gs {
				stop := GradientStop{
					Pos: float64(gs.Pos) / 100000.0,
				}
				if gs.SchemeClr != nil {
					cc := &OOXMLColorChoice{SchemeClr: gs.SchemeClr}
					stop.Color = resolveColorChoice(cc, themeColors)
				} else if gs.SrgbClr != nil {
					cc := &OOXMLColorChoice{SrgbClr: gs.SrgbClr}
					stop.Color = resolveColorChoice(cc, themeColors)
				}
				if stop.Color == "" {
					accentColors := getChartAccentColors(themeColors)
					stop.Color = accentColors[i%len(accentColors)]
				}
				sd.GradStops = append(sd.GradStops, stop)
			}
		}

		// Check for line gradients (borders)
		if sl.Ser.SpPr != nil && sl.Ser.SpPr.Ln != nil && sl.Ser.SpPr.Ln.GradFill != nil {
			sd.LineIsGradient = true
			if sl.Ser.SpPr.Ln.GradFill.Path != nil {
				sd.LineGradIsRadial = true
			} else if sl.Ser.SpPr.Ln.GradFill.Lin != nil {
				sd.LineGradAngle = float64(sl.Ser.SpPr.Ln.GradFill.Lin.Ang) / 60000.0
			}
			for _, gs := range sl.Ser.SpPr.Ln.GradFill.GsLst.Gs {
				stop := GradientStop{
					Pos: float64(gs.Pos) / 100000.0,
				}
				if gs.SchemeClr != nil {
					cc := &OOXMLColorChoice{SchemeClr: gs.SchemeClr}
					stop.Color = resolveColorChoice(cc, themeColors)
				} else if gs.SrgbClr != nil {
					cc := &OOXMLColorChoice{SrgbClr: gs.SrgbClr}
					stop.Color = resolveColorChoice(cc, themeColors)
				}
				if stop.Color == "" {
					accentColors := getChartAccentColors(themeColors)
					stop.Color = accentColors[i%len(accentColors)]
				}
				sd.LineGradStops = append(sd.LineGradStops, stop)
			}
		}

		// Y-Values
		if sl.Ser.Val.NumRef != nil {
			for _, pt := range sl.Ser.Val.NumRef.NumCache.Pts {
				v, _ := strconv.ParseFloat(pt.V, 64)
				sd.Values = append(sd.Values, v)
			}
		} else if sl.Ser.YVal != nil && sl.Ser.YVal.NumRef != nil {
			for _, pt := range sl.Ser.YVal.NumRef.NumCache.Pts {
				v, _ := strconv.ParseFloat(pt.V, 64)
				sd.Values = append(sd.Values, v)
			}
		}

		// X-Values (for scatter)
		if sl.Ser.XVal != nil && sl.Ser.XVal.NumRef != nil {
			for _, pt := range sl.Ser.XVal.NumRef.NumCache.Pts {
				v, _ := strconv.ParseFloat(pt.V, 64)
				sd.XValues = append(sd.XValues, v)
			}
		}

		data.Series = append(data.Series, sd)

		// Parse data point fills (dPt elements) for pie and doughnut charts
		if (sl.Type == "pie" || sl.Type == "doughnut") && len(sl.Ser.DPts) > 0 {
			for _, dpt := range sl.Ser.DPts {
				idx, _ := strconv.Atoi(dpt.Idx.Val)
				// Ensure DPtFills is large enough
				for len(data.DPtFills) <= idx {
					data.DPtFills = append(data.DPtFills, DataPointFill{})
				}
				if dpt.SpPr != nil {
					dpf := DataPointFill{}
					if dpt.SpPr.PattFill != nil {
						// Pattern fill
						pf := dpt.SpPr.PattFill
						dpf.PatternPreset = pf.Prst
						dpf.FgColor = resolveColorChoice(pf.FgClr, themeColors)
						dpf.BgColor = resolveColorChoice(pf.BgClr, themeColors)
						if dpf.FgColor == "" {
							accentColors := getChartAccentColors(themeColors)
							dpf.FgColor = accentColors[idx%len(accentColors)]
						}
						if dpf.BgColor == "" {
							dpf.BgColor = "#FFFFFF"
						}
					} else if dpt.SpPr.SolidFill != nil {
						// Solid fill override
						dpf.IsSolid = true
						dpf.SolidColor = resolveChartFill(dpt.SpPr.SolidFill, themeColors)
					} else if dpt.SpPr.GradFill != nil {
						dpf.IsGradient = true
						if dpt.SpPr.GradFill.Path != nil {
							dpf.GradIsRadial = true
						} else if dpt.SpPr.GradFill.Lin != nil {
							dpf.GradAngle = float64(dpt.SpPr.GradFill.Lin.Ang) / 60000.0
						}
						// Depending on struct shape we assume dpt.SpPr.GradFill.GsLst based on fix logic
						for _, gs := range dpt.SpPr.GradFill.GsLst {
							stop := GradientStop{
								Pos: float64(gs.Pos) / 100000.0,
							}
							if gs.SchemeClr != nil {
								cc := &OOXMLColorChoice{SchemeClr: gs.SchemeClr}
								stop.Color = resolveColorChoice(cc, themeColors)
							} else if gs.SrgbClr != nil {
								cc := &OOXMLColorChoice{SrgbClr: gs.SrgbClr}
								stop.Color = resolveColorChoice(cc, themeColors)
							}
							if stop.Color == "" {
								accentColors := getChartAccentColors(themeColors)
								stop.Color = accentColors[idx%len(accentColors)]
							}
							dpf.GradStops = append(dpf.GradStops, stop)
						}
					} else if dpt.SpPr.BlipFill != nil {
						dpf.HasBlipFill = true
						dpf.BlipEmbedRID = dpt.SpPr.BlipFill.Blip.Embed
					}

					// Effects (Shadow)
					if dpt.SpPr.EffectLst != nil {
						dpf.Shadow = extractShapeShadow(dpt.SpPr.EffectLst, themeColors)
						if dpf.Shadow != nil {
							dpf.HasShadow = true
						}
					}

					if dpt.SpPr.Ln != nil && dpt.SpPr.Ln.GradFill != nil {
						dpf.LineIsGradient = true
						if dpt.SpPr.Ln.GradFill.Path != nil {
							dpf.LineGradIsRadial = true
						} else if dpt.SpPr.Ln.GradFill.Lin != nil {
							dpf.LineGradAngle = float64(dpt.SpPr.Ln.GradFill.Lin.Ang) / 60000.0
						}
						// Depending on struct shape we assume dpt.SpPr.Ln.GradFill.GsLst based on fix logic
						for _, gs := range dpt.SpPr.Ln.GradFill.GsLst {
							stop := GradientStop{
								Pos: float64(gs.Pos) / 100000.0,
							}
							if gs.SchemeClr != nil {
								cc := &OOXMLColorChoice{SchemeClr: gs.SchemeClr}
								stop.Color = resolveColorChoice(cc, themeColors)
							} else if gs.SrgbClr != nil {
								cc := &OOXMLColorChoice{SrgbClr: gs.SrgbClr}
								stop.Color = resolveColorChoice(cc, themeColors)
							}
							if stop.Color == "" {
								accentColors := getChartAccentColors(themeColors)
								stop.Color = accentColors[idx%len(accentColors)]
							}
							dpf.LineGradStops = append(dpf.LineGradStops, stop)
						}
					}

					// Stroke from dPt line
					if dpt.SpPr.Ln != nil {
						if dpt.SpPr.Ln.SolidFill != nil {
							dpf.StrokeColor = resolveChartFill(dpt.SpPr.Ln.SolidFill, themeColors)
						}
						if dpt.SpPr.Ln.W > 0 {
							dpf.StrokeWidth = float64(dpt.SpPr.Ln.W) / 12700.0 // EMU to px
						}
					}
					data.DPtFills[idx] = dpf
				}
			}
		}
	}

	// Resolve chart title
	// 1. Check for explicit title text in <c:title><c:tx><c:rich>
	if cs.Chart.Title != nil && cs.Chart.Title.Tx != nil && cs.Chart.Title.Tx.Rich != nil {
		var parts []string
		capMode := "" // "all", "small", or ""
		for _, p := range cs.Chart.Title.Tx.Rich.Paragraphs {
			// Check for cap attribute on paragraph default run properties
			if p.PPr != nil && p.PPr.DefRPr != nil && p.PPr.DefRPr.Cap != "" {
				capMode = p.PPr.DefRPr.Cap
			}
			for _, r := range p.Runs {
				if r.T != "" {
					parts = append(parts, r.T)
				}
			}
		}
		if len(parts) > 0 {
			title := strings.Join(parts, " ")
			// Apply capitalization
			switch capMode {
			case "all":
				title = strings.ToUpper(title)
			case "small":
				// small caps: visually uppercase but not semantically; just uppercase for SVG
				title = strings.ToUpper(title)
			}
			data.Title = title
		}
	}
	// 2. Auto-title: if still default, use first series name
	if data.Title == "Chart Title" && len(data.Series) > 0 {
		data.Title = data.Series[0].Name
	}

	// Legend
	if cs.Chart.Legend != nil {
		data.HasLegend = true
		data.LegendPos = cs.Chart.Legend.LegendPos.Val
		if data.LegendPos == "" {
			data.LegendPos = "b"
		}
	}

	// Gridlines
	if cs.Chart.PlotArea.ValAx != nil && cs.Chart.PlotArea.ValAx.MajorGridlines != nil {
		data.HasGrid = true
	}

	// Extract font sizes from chart XML (hundredths of a point â†’ pixels)
	// Use direct conversion to avoid integer truncation: hundredths_pt / 100 * 96/72
	if cs.Chart.Title != nil && cs.Chart.Title.TxPr != nil {
		for _, p := range cs.Chart.Title.TxPr.Paragraphs {
			if p.PPr != nil && p.PPr.DefRPr != nil && p.PPr.DefRPr.Sz > 0 {
				data.TitleFontSize = int(math.Round(float64(p.PPr.DefRPr.Sz) / 100.0 * 96.0 / 72.0))
				break
			}
		}
	}
	// Axis label font size (from catAx or valAx txPr)
	if cs.Chart.PlotArea.CatAx != nil && cs.Chart.PlotArea.CatAx.TxPr != nil {
		for _, p := range cs.Chart.PlotArea.CatAx.TxPr.Paragraphs {
			if p.PPr != nil && p.PPr.DefRPr != nil && p.PPr.DefRPr.Sz > 0 {
				data.LabelFontSize = int(math.Round(float64(p.PPr.DefRPr.Sz) / 100.0 * 96.0 / 72.0))
				break
			}
		}
	}

	// Chart background colour from <c:chartSpace><c:spPr>
	if cs.SpPr != nil {
		if cs.SpPr.SolidFill != nil {
			if c := resolveChartFill(cs.SpPr.SolidFill, themeColors); c != "" {
				data.BgColor = c
			}
		} else if cs.SpPr.GradFill != nil {
			data.BgIsGradient = true
			if cs.SpPr.GradFill.Path != nil && (cs.SpPr.GradFill.Path.Path == "circle" || cs.SpPr.GradFill.Path.Path == "shape") {
				data.BgGradIsRadial = true
			}
			for _, gs := range cs.SpPr.GradFill.GsLst {
				stop := GradientStop{
					Pos: float64(gs.Pos) / 100000.0,
				}
				if gs.SchemeClr != nil {
					cc := &OOXMLColorChoice{SchemeClr: gs.SchemeClr}
					stop.Color = resolveColorChoice(cc, themeColors)
				} else if gs.SrgbClr != nil {
					cc := &OOXMLColorChoice{SrgbClr: gs.SrgbClr}
					stop.Color = resolveColorChoice(cc, themeColors)
				}
				if stop.Color == "" {
					stop.Color = "#FFFFFF" // Fallback
				}
				data.BgGradStops = append(data.BgGradStops, stop)
			}
		}
	}

	return data, nil
}

// resolveSeriesColor extracts the fill or line color from a series definition.
func resolveSeriesColor(ser ChartSeries, index int, themeColors map[string]string) string {
	// Check direct solidFill first (bar/pie charts)
	if ser.SpPr != nil && ser.SpPr.SolidFill != nil {
		if c := resolveChartFill(ser.SpPr.SolidFill, themeColors); c != "" {
			return c
		}
	}
	// Check line color (line charts use <a:ln><a:solidFill>)
	if ser.SpPr != nil && ser.SpPr.Ln != nil && ser.SpPr.Ln.SolidFill != nil {
		if c := resolveChartFill(ser.SpPr.Ln.SolidFill, themeColors); c != "" {
			return c
		}
	}
	// Fallback to theme accent colors by index
	accentColors := getChartAccentColors(themeColors)
	if index < len(accentColors) {
		return accentColors[index]
	}
	return resolveSchemeColor("accent1", themeColors)
}

// resolveChartFill resolves a color from an OOXMLSolidFill (srgbClr or schemeClr).
func resolveChartFill(sf *OOXMLSolidFill, themeColors map[string]string) string {
	if sf.SrgbClr != nil {
		return "#" + sf.SrgbClr.Val
	}
	if sf.SchemeClr != nil {
		return resolveSchemeColor(sf.SchemeClr.Val, themeColors)
	}
	return ""
}

func generateSVGLinearGradient(id string, angle float64, stops []GradientStop) string {
	rad := angle * math.Pi / 180.0
	x1 := 50.0 - math.Cos(rad)*50.0
	y1 := 50.0 - math.Sin(rad)*50.0
	x2 := 50.0 + math.Cos(rad)*50.0
	y2 := 50.0 + math.Sin(rad)*50.0

	var out strings.Builder
	out.WriteString(fmt.Sprintf("    <linearGradient id=\"%s\" x1=\"%.2f%%\" y1=\"%.2f%%\" x2=\"%.2f%%\" y2=\"%.2f%%\">\n", id, x1, y1, x2, y2))
	for _, stop := range stops {
		out.WriteString(fmt.Sprintf("      <stop offset=\"%.0f%%\" stop-color=\"%s\" />\n", stop.Pos*100, stop.Color))
	}
	out.WriteString("    </linearGradient>\n")
	return out.String()
}

func generateSVGRadialGradient(id string, stops []GradientStop) string {
	var out strings.Builder
	out.WriteString(fmt.Sprintf("    <radialGradient id=\"%s\" cx=\"50%%\" cy=\"50%%\" r=\"75%%\" fx=\"50%%\" fy=\"50%%\">\n", id))
	for _, stop := range stops {
		out.WriteString(fmt.Sprintf("      <stop offset=\"%.0f%%\" stop-color=\"%s\" />\n", stop.Pos*100, stop.Color))
	}
	out.WriteString("    </radialGradient>\n")
	return out.String()
}

// renderChartSVG dispatches chart rendering based on chart type.
func renderChartSVG(data *ChartData, x, y, w, h int) string {
	switch data.ChartType {
	case "bar":
		return renderBarChartSVG(data, x, y, w, h)
	case "line":
		return renderLineChartSVG(data, x, y, w, h)
	case "pie":
		return renderPieChartSVG(data, x, y, w, h)
	case "doughnut":
		return renderDoughnutChartSVG(data, x, y, w, h)
	case "area":
		return renderAreaChartSVG(data, x, y, w, h)
	case "scatter":
		return renderScatterChartSVG(data, x, y, w, h)
	case "radar":
		return renderRadarChartSVG(data, x, y, w, h)
	case "treemap":
		return renderTreemapChartSVG(data, x, y, w, h)
	case "sunburst":
		return renderSunburstChartSVG(data, x, y, w, h)
	case "waterfall":
		return renderWaterfallChartSVG(data, x, y, w, h)
	case "clusteredColumn", "combo":
		return renderBarChartSVG(data, x, y, w, h)
	default:
		// Fallback for surface/stock
		if data.ChartType == "stock" {
			return renderLineChartSVG(data, x, y, w, h)
		}
		if data.ChartType == "surface" || data.ChartType == "surface3D" {
			return renderAreaChartSVG(data, x, y, w, h)
		}
		return renderBarChartSVG(data, x, y, w, h)
	}
}

// renderBarChartSVG renders a clustered bar chart as an SVG fragment.
func renderBarChartSVG(data *ChartData, x, y, w, h int) string {
	if len(data.Categories) == 0 {
		return ""
	}

	var out strings.Builder

	// Font sizes from chart XML or defaults
	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 25 // ~18.62pt
	}
	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 16 // ~12pt
	}

	// Chart layout margins (calibrated to match PowerPoint default chart layout)
	titleH := titleFS + 25 // height for title area
	legendH := 0
	if data.HasLegend {
		legendH = labelFS + 18 // height for legend area
	}
	leftMargin := 33 // for Y-axis labels
	rightMargin := 15
	bottomMargin := labelFS + 28 // for X-axis labels + gap
	topMargin := 17

	// Compute plot area
	plotX := x + leftMargin
	plotY := y + topMargin + titleH
	plotW := w - leftMargin - rightMargin
	plotH := h - topMargin - titleH - bottomMargin - legendH

	if plotW < 50 || plotH < 50 {
		return "" // too small to render
	}

	hasParetoLine := false
	hasAnyShadow := false
	hasAnyGradient := false
	for _, ser := range data.Series {
		if ser.Type == "paretoLine" {
			hasParetoLine = true
		}
		if ser.HasShadow && ser.Shadow != nil {
			hasAnyShadow = true
		}
		if ser.IsGradient && len(ser.GradStops) > 0 {
			hasAnyGradient = true
		}
	}

	if hasAnyShadow || hasAnyGradient {
		out.WriteString("  <defs>\n")
		// Generate Drop Shadows & Gradients for Series
		for i, series := range data.Series {
			if series.HasShadow && series.Shadow != nil {
				filterID := fmt.Sprintf("chart-shadow-%d", i)
				if series.Shadow.Type == "inner" {
					out.WriteString(generateInnerShadowFilter(series.Shadow, filterID))
				} else {
					out.WriteString(generateShadowFilter(series.Shadow, filterID, false))
				}
			}
			if series.IsGradient {
				hasAnyGradient = true
				if series.GradIsRadial {
					out.WriteString(generateSVGRadialGradient(fmt.Sprintf("chart-series-grad-%d", i), series.GradStops))
				} else {
					out.WriteString(generateSVGLinearGradient(fmt.Sprintf("chart-series-grad-%d", i), series.GradAngle, series.GradStops))
				}
			}
		}
		out.WriteString("  </defs>\n")
	}

	// Find min/max values
	maxVal := 0.0

	// Track stacked sums per grouping/type
	type stackKey struct {
		chartType string
		grouping  string
	}
	stackedSums := make(map[stackKey][]float64)

	for _, s := range data.Series {
		if s.Type == "paretoLine" {
			continue
		}
		isStacked := s.Grouping == "stacked" || s.Grouping == "percentStacked"
		var sums []float64

		if isStacked {
			key := stackKey{s.Type, s.Grouping}
			if stackedSums[key] == nil {
				stackedSums[key] = make([]float64, len(data.Categories))
			}
			sums = stackedSums[key]
		}

		for i, v := range s.Values {
			if isStacked && i < len(sums) {
				sums[i] += v
				if sums[i] > maxVal {
					maxVal = sums[i]
				}
			} else {
				if v > maxVal {
					maxVal = v
				}
			}
		}
	}
	// Compute nice tick interval and axis max
	tickInterval := niceTickInterval(maxVal)
	niceMax := niceAxisMax(maxVal)
	if niceMax == 0 {
		niceMax = 10
	}
	numTicks := int(math.Round(niceMax / tickInterval))

	// --- Background rect for chart area ---
	barBgColor := "#F5F5F5"
	if data.BgColor != "" {
		barBgColor = data.BgColor
	}
	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="%s" rx="0"/>
`, x, y, w, h, barBgColor))

	// --- Chart Title ---
	titleX := x + w/2
	titleY := y + topMargin + titleFS + 2
	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="#595959" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, titleX, titleY, titleFS, escapeXMLText(data.Title)))

	// --- Gridlines ---
	gridColor := "#D9D9D9"
	axisLabelColor := "#595959"

	// Y-axis labels and horizontal gridlines
	for i := 0; i <= numTicks; i++ {
		val := float64(i) * tickInterval
		yPos := plotY + plotH - int(float64(plotH)*(val/niceMax))

		// Gridline (skip bottom line, draw it as axis)
		if i > 0 {
			out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, plotX, yPos, plotX+plotW, yPos, gridColor))
		}

		// Y-axis label
		labelStr := formatAxisValue(val)
		labelX := plotX - 5
		labelY := yPos + labelFS/3
		out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="end">%s</text>
`, labelX, labelY, labelFS, axisLabelColor, labelStr))

		// Secondary Y-axis label (if Pareto)
		if hasParetoLine {
			pctVal := 100.0 * float64(i) / float64(numTicks)
			secLabelStr := fmt.Sprintf("%.0f%%", pctVal)
			secLabelX := plotX + plotW + 5
			out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="start">%s</text>
`, secLabelX, labelY, labelFS, axisLabelColor, secLabelStr))
		}
	}

	// Bottom axis line
	axisY := plotY + plotH
	out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, plotX, axisY, plotX+plotW, axisY, gridColor))

	// Draw primary Y-axis vertical line for clarity on Extended Charts
	out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, plotX, plotY, plotX, axisY, gridColor))

	if hasParetoLine {
		out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, plotX+plotW, plotY, plotX+plotW, axisY, gridColor))
	}

	// --- Area Series (Combo Overlay) ---
	areaBottomY := make(map[string][]float64) // key is grouping
	numCats := len(data.Categories)

	for origSerIdx, ser := range data.Series {
		if ser.Type != "area" {
			continue
		}

		isStacked := ser.Grouping == "stacked" || ser.Grouping == "percentStacked"
		key := ser.Grouping

		if isStacked && areaBottomY[key] == nil {
			initY := make([]float64, numCats)
			for i := range initY {
				initY[i] = float64(plotY + plotH)
			}
			areaBottomY[key] = initY
		}

		var points []string
		var topY []float64

		// Top edge (left to right)
		for catIdx := 0; catIdx < numCats && catIdx < len(ser.Values); catIdx++ {
			px := float64(plotX) + (float64(catIdx)+0.5)*float64(plotW)/float64(numCats)
			var py float64

			if isStacked {
				bottom := areaBottomY[key][catIdx]
				delta := float64(plotH) * (ser.Values[catIdx] / niceMax)
				py = bottom - delta
				topY = append(topY, py)
			} else {
				py = float64(plotY+plotH) - float64(plotH)*(ser.Values[catIdx]/niceMax)
			}
			points = append(points, fmt.Sprintf("%.1f,%.1f", px, py))
		}

		// Bottom edge closing path
		if isStacked {
			for catIdx := len(topY) - 1; catIdx >= 0; catIdx-- {
				px := float64(plotX) + (float64(catIdx)+0.5)*float64(plotW)/float64(numCats)
				py := areaBottomY[key][catIdx]
				points = append(points, fmt.Sprintf("%.1f,%.1f", px, py))
			}
			for i, ty := range topY {
				areaBottomY[key][i] = ty
			}
		} else {
			if len(ser.Values) > 0 {
				lastIdx := len(ser.Values) - 1
				if lastIdx >= numCats {
					lastIdx = numCats - 1
				}
				lastPx := float64(plotX) + (float64(lastIdx)+0.5)*float64(plotW)/float64(numCats)
				points = append(points, fmt.Sprintf("%.1f,%.1f", lastPx, float64(plotY+plotH)))

				firstPx := float64(plotX) + 0.5*float64(plotW)/float64(numCats)
				points = append(points, fmt.Sprintf("%.1f,%.1f", firstPx, float64(plotY+plotH)))
			}
		}

		if len(points) > 0 {
			filterAttr := ""
			if ser.HasShadow && ser.Shadow != nil {
				filterAttr = fmt.Sprintf(` filter="url(#chart-shadow-%d)"`, origSerIdx)
			}
			fillAttr := fmt.Sprintf(`fill="%s"`, ser.Color)
			if ser.IsGradient {
				fillAttr = fmt.Sprintf(`fill="url(#chart-series-grad-%d)"`, origSerIdx)
			}
			out.WriteString(fmt.Sprintf(`  <polygon points="%s" %s stroke="%s" stroke-width="1.5" stroke-linejoin="round"%s/>
`, strings.Join(points, " "), fillAttr, ser.Color, filterAttr))
		}
	}

	// --- Bars ---
	numBarSeries := 0
	for _, s := range data.Series {
		if s.Type != "line" && s.Type != "paretoLine" && s.Type != "scatter" && s.Type != "area" {
			numBarSeries++
		}
	}

	if numCats == 0 || len(data.Series) == 0 {
		return out.String()
	}

	// Compute bar dimensions
	// Category width = plot width / number of categories
	catW := float64(plotW) / float64(numCats)

	// Gap between categories (as fraction of bar width)
	gapFraction := 2.19 // Default 219%
	if data.HasGapWidth {
		gapFraction = data.GapWidth
	}
	overlapFraction := 0.0
	if data.HasOverlap {
		overlapFraction = data.Overlap
	}

	var singleBarW, gapOneSide float64

	if numBarSeries > 0 {
		// Calculate single bar width such that:
		// catW = SingleBarWidth * (GapFraction + NumSeries - (NumSeries-1)*OverlapFraction)
		denominator := gapFraction + float64(numBarSeries) - float64(numBarSeries-1)*overlapFraction
		if denominator <= 0 {
			denominator = float64(numBarSeries) + 1
		}
		singleBarW = catW / denominator
		if singleBarW < 2 {
			singleBarW = 2
		}

		// Total gap size for the category is GapFraction * singleBarW
		gapOneSide = (gapFraction * singleBarW) / 2.0
	}

	for catIdx := 0; catIdx < numCats; catIdx++ {
		catStartX := float64(plotX) + float64(catIdx)*catW

		barSerIdx := 0
		for origSerIdx, ser := range data.Series {
			if ser.Type == "line" || ser.Type == "paretoLine" || ser.Type == "scatter" || ser.Type == "area" {
				continue
			}

			if catIdx >= len(ser.Values) {
				barSerIdx++
				continue
			}
			val := ser.Values[catIdx]
			if val <= 0 {
				barSerIdx++
				continue
			}

			// Apply overlap correctly. Without overlap, consecutive bars are spaced by `singleBarW`.
			// With overlap, the distance between left edges of consecutive bars shrinks.
			offsetMultiplier := 1.0 - overlapFraction
			barX := catStartX + gapOneSide + float64(barSerIdx)*(singleBarW*offsetMultiplier)

			barH := float64(plotH) * (val / niceMax)
			barY := float64(plotY+plotH) - barH

			filterAttr := ""
			if ser.HasShadow && ser.Shadow != nil {
				filterID := fmt.Sprintf("chart-shadow-%d", origSerIdx)
				filterAttr = fmt.Sprintf(` filter="url(#%s)"`, filterID)
			}
			fillAttr := fmt.Sprintf(`fill="%s"`, ser.Color)
			if ser.IsGradient {
				fillAttr = fmt.Sprintf(`fill="url(#chart-series-grad-%d)"`, origSerIdx)
			}

			strokeAttr := ""
			if ser.HasLine && ser.LineColor != "" {
				strokeAttr = fmt.Sprintf(` stroke="%s" stroke-width="1.5"`, ser.LineColor)
			}

			out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%.0f" width="%.0f" height="%.0f" %s%s%s/>
`, barX, barY, singleBarW, barH, fillAttr, filterAttr, strokeAttr))

			// Add data label if there's enough height
			if barH > 15 {
				labelValStr := fmt.Sprintf("%g", val)
				out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%.0f" font-size="16" fill="#FEFEFE" font-family="Calibri, sans-serif" text-anchor="middle">%s</text>
`, barX+singleBarW/2, barY+16, escapeXMLText(labelValStr)))
			}

			barSerIdx++
		}
	}

	// --- X-axis category labels ---
	catLabelY := plotY + plotH + labelFS + 8
	for catIdx, catName := range data.Categories {
		catCenterX := float64(plotX) + (float64(catIdx)+0.5)*catW
		out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, catCenterX, catLabelY, labelFS, axisLabelColor, escapeXMLText(catName)))
	}

	// --- Overlay Lines (Combo / Pareto) ---
	for origSerIdx, ser := range data.Series {
		if ser.Type != "line" && ser.Type != "paretoLine" && ser.Type != "scatter" {
			continue
		}

		var pathBuilder strings.Builder
		for catIdx := 0; catIdx < numCats; catIdx++ {
			if catIdx >= len(ser.Values) {
				continue
			}
			val := ser.Values[catIdx]

			catCenterX := float64(plotX) + (float64(catIdx)+0.5)*catW

			plotVal := val
			lineNiceMax := niceMax
			if ser.Type == "paretoLine" {
				lineNiceMax = 100.0 // Pareto line is 0 to 100%
			}

			centerY := float64(plotY+plotH) - float64(plotH)*(plotVal/lineNiceMax)

			// Drop shadow filter application for dots/lines
			filterAttr := ""
			if ser.HasShadow && ser.Shadow != nil {
				filterAttr = fmt.Sprintf(` filter="url(#chart-shadow-%d)"`, origSerIdx)
			}

			if ser.Type == "line" || ser.Type == "paretoLine" {
				if catIdx == 0 {
					pathBuilder.WriteString(fmt.Sprintf("M%.2f %.2f ", catCenterX, centerY))
				} else {
					pathBuilder.WriteString(fmt.Sprintf("L%.2f %.2f ", catCenterX, centerY))
				}
			}

			// Draw data point marker
			out.WriteString(fmt.Sprintf(`  <circle cx="%.2f" cy="%.2f" r="4" fill="%s"%s/>
`, catCenterX, centerY, ser.Color, filterAttr))
		}

		if pathBuilder.Len() > 0 {
			filterAttr := ""
			if ser.HasShadow && ser.Shadow != nil {
				filterAttr = fmt.Sprintf(` filter="url(#chart-shadow-%d)"`, origSerIdx)
			}
			out.WriteString(fmt.Sprintf(`  <path d="%s" stroke="%s" stroke-width="2" fill="none"%s/>
`, pathBuilder.String(), ser.Color, filterAttr))
		}
	}

	// --- Legend ---
	if data.HasLegend && len(data.Series) > 0 {
		legendY := plotY + plotH + bottomMargin + labelFS + 2
		legendBoxSize := 9

		// Calculate total legend width for centering
		// Use approximate character width based on font size
		charW := float64(labelFS) * 0.55
		totalLegendW := 0.0
		for _, ser := range data.Series {
			totalLegendW += float64(legendBoxSize) + 4 + float64(len(ser.Name))*charW + 16
		}
		legendStartX := float64(x) + (float64(w)-totalLegendW)/2.0

		curX := legendStartX
		for _, ser := range data.Series {
			// Color box
			boxY := legendY - legendBoxSize + 1
			out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="%s"/>
`, curX, boxY, legendBoxSize, legendBoxSize, ser.Color))
			curX += float64(legendBoxSize) + 4

			// Label
			out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif">%s</text>
`, curX, legendY, labelFS, axisLabelColor, escapeXMLText(ser.Name)))
			curX += float64(len(ser.Name))*charW + 16
		}
	}

	return out.String()
}

// ====================================================================
// Line Chart Renderer
// ====================================================================

// renderLineChartSVG renders a line chart as an SVG fragment.
func renderLineChartSVG(data *ChartData, x, y, w, h int) string {
	if len(data.Categories) == 0 {
		return ""
	}

	var out strings.Builder

	// Font sizes
	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 25
	}
	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 16
	}

	// Layout
	titleH := titleFS + 25
	legendH := 0
	if data.HasLegend {
		legendH = labelFS + 18
	}
	leftMargin := 33
	rightMargin := 15
	bottomMargin := labelFS + 28
	topMargin := 17

	plotX := x + leftMargin
	plotY := y + topMargin + titleH
	plotW := w - leftMargin - rightMargin
	plotH := h - topMargin - titleH - bottomMargin - legendH

	if plotW < 50 || plotH < 50 {
		return ""
	}

	// Find max value
	maxVal := 0.0
	for _, s := range data.Series {
		for _, v := range s.Values {
			if v > maxVal {
				maxVal = v
			}
		}
	}
	tickInterval := niceTickInterval(maxVal)
	niceMax := niceAxisMax(maxVal)
	if niceMax == 0 {
		niceMax = 10
	}
	numTicks := int(math.Round(niceMax / tickInterval))

	// Background
	lineBgColor := "#F5F5F5"
	if data.BgColor != "" {
		lineBgColor = data.BgColor
	}
	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="%s" rx="0"/>
`, x, y, w, h, lineBgColor))

	// Title
	titleX := x + w/2
	titleY := y + topMargin + titleFS + 2
	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="#595959" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, titleX, titleY, titleFS, escapeXMLText(data.Title)))

	// Gridlines + Y-axis labels
	gridColor := "#D9D9D9"
	axisLabelColor := "#595959"

	for i := 0; i <= numTicks; i++ {
		val := float64(i) * tickInterval
		yPos := plotY + plotH - int(float64(plotH)*(val/niceMax))

		if i > 0 {
			out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, plotX, yPos, plotX+plotW, yPos, gridColor))
		}

		labelStr := formatAxisValue(val)
		labelX := plotX - 5
		labelY := yPos + labelFS/3
		out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="end">%s</text>
`, labelX, labelY, labelFS, axisLabelColor, labelStr))
	}

	// Bottom axis line
	axisY := plotY + plotH
	out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, plotX, axisY, plotX+plotW, axisY, gridColor))

	// Draw lines for each series
	numCats := len(data.Categories)
	for _, ser := range data.Series {
		if len(ser.Values) == 0 {
			continue
		}
		var points []string
		for catIdx := 0; catIdx < numCats && catIdx < len(ser.Values); catIdx++ {
			// Line charts place data points centered in each category slot
			px := float64(plotX) + (float64(catIdx)+0.5)*float64(plotW)/float64(numCats)
			py := float64(plotY+plotH) - float64(plotH)*(ser.Values[catIdx]/niceMax)
			points = append(points, fmt.Sprintf("%.1f,%.1f", px, py))
		}
		if len(points) > 0 {
			out.WriteString(fmt.Sprintf(`  <polyline points="%s" fill="none" stroke="%s" stroke-width="2.5" stroke-linejoin="round"/>
`, strings.Join(points, " "), ser.Color))
		}
	}

	// X-axis category labels
	catLabelY := plotY + plotH + labelFS + 8
	for catIdx, catName := range data.Categories {
		catCenterX := float64(plotX) + (float64(catIdx)+0.5)*float64(plotW)/float64(numCats)
		out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, catCenterX, catLabelY, labelFS, axisLabelColor, escapeXMLText(catName)))
	}

	// Legend
	if data.HasLegend && len(data.Series) > 0 {
		legendY := plotY + plotH + bottomMargin + labelFS + 2
		charW := float64(labelFS) * 0.55

		// Calculate total legend width for centering
		totalLegendW := 0.0
		for _, ser := range data.Series {
			totalLegendW += 20 + 4 + float64(len(ser.Name))*charW + 16 // line swatch + gap + text + spacing
		}
		legendStartX := float64(x) + (float64(w)-totalLegendW)/2.0

		curX := legendStartX
		for _, ser := range data.Series {
			// Small line swatch instead of box
			swatchY := legendY - 5
			out.WriteString(fmt.Sprintf(`  <line x1="%.0f" y1="%.0f" x2="%.0f" y2="%.0f" stroke="%s" stroke-width="2.5"/>
`, curX, float64(swatchY), curX+20, float64(swatchY), ser.Color))
			curX += 24

			// Label
			out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif">%s</text>
`, curX, legendY, labelFS, axisLabelColor, escapeXMLText(ser.Name)))
			curX += float64(len(ser.Name))*charW + 16
		}
	}

	return out.String()
}

// ====================================================================
// Pie Chart Renderer
// ====================================================================

// generateSVGFillAndEffectDefs produces <defs>...</defs> with SVG pattern, gradient, and filter elements
// for any data points that use them. Returns the defs SVG string and maps from data-point index to pattern/gradient ID and filter ID.
func generateSVGFillAndEffectDefs(data *ChartData, chartID string) (string, map[int]string, map[int]string, map[int]string) {
	patternIDs := make(map[int]string)
	gradientIDs := make(map[int]string)
	filterIDs := make(map[int]string)

	if len(data.DPtFills) == 0 {
		return "", patternIDs, gradientIDs, filterIDs
	}

	var defs strings.Builder
	for i, dpf := range data.DPtFills {
		// Patterns
		if dpf.PatternPreset != "" {
			patID := fmt.Sprintf("chartPat%s_%d", chartID, i)
			patternIDs[i] = patID
			defs.WriteString(generateOOXMLPattern(patID, dpf.PatternPreset, dpf.FgColor, dpf.BgColor))
		}
		// Gradients
		if dpf.IsGradient && len(dpf.GradStops) > 0 {
			gradID := fmt.Sprintf("chartGrad%s_%d", chartID, i)
			gradientIDs[i] = gradID
			if dpf.GradIsRadial {
				defs.WriteString(generateSVGRadialGradient(gradID, dpf.GradStops))
			} else {
				defs.WriteString(generateSVGLinearGradient(gradID, dpf.GradAngle, dpf.GradStops))
			}
		}
		// Shadows
		if dpf.HasShadow && dpf.Shadow != nil {
			filterID := fmt.Sprintf("chartShdw%s_%d", chartID, i)
			filterIDs[i] = filterID

			if dpf.Shadow.Type == "inner" {
				defs.WriteString(generateInnerShadowFilter(dpf.Shadow, filterID))
			} else {
				defs.WriteString(generateShadowFilter(dpf.Shadow, filterID, false))
			}
		}
	}

	if data.BgIsGradient && len(data.BgGradStops) > 0 {
		gradID := fmt.Sprintf("chartBgGrad%s", chartID)
		if data.BgGradIsRadial {
			defs.WriteString(fmt.Sprintf("    <radialGradient id=\"%s\" cx=\"50%%\" cy=\"50%%\" r=\"75%%\" fx=\"50%%\" fy=\"50%%\">\n", gradID))
		} else {
			defs.WriteString(fmt.Sprintf("    <linearGradient id=\"%s\" x1=\"0%%\" y1=\"0%%\" x2=\"0%%\" y2=\"100%%\">\n", gradID))
		}
		for _, stop := range data.BgGradStops {
			defs.WriteString(fmt.Sprintf("      <stop offset=\"%.0f%%\" stop-color=\"%s\" />\n", stop.Pos*100, stop.Color))
		}
		if data.BgGradIsRadial {
			defs.WriteString("    </radialGradient>\n")
		} else {
			defs.WriteString("    </linearGradient>\n")
		}
	}

	if defs.Len() == 0 {
		return "", patternIDs, gradientIDs, filterIDs
	}
	return fmt.Sprintf("  <defs>\n%s  </defs>\n", defs.String()), patternIDs, gradientIDs, filterIDs
}

// generateOOXMLPattern creates an SVG <pattern> element for a given OOXML pattern preset.
func generateOOXMLPattern(id, preset, fgColor, bgColor string) string {
	// Pattern tile size
	sz := 8
	var inner string

	switch preset {
	case "ltUpDiag": // Light Upward Diagonal
		inner = fmt.Sprintf(
			`    <rect width="%d" height="%d" fill="%s"/>
    <path d="M0 %d L%d 0 M-%d %d L%d -%d M%d %d L%d %d" stroke="%s" stroke-width="1"/>`,
			sz, sz, bgColor,
			sz, sz, 1, 1, 1, 1, sz-1, sz+1, sz+1, sz-1, fgColor)
	case "ltDnDiag": // Light Downward Diagonal
		inner = fmt.Sprintf(
			`    <rect width="%d" height="%d" fill="%s"/>
    <path d="M0 0 L%d %d M-%d %d L%d %d M%d -%d L%d %d" stroke="%s" stroke-width="1"/>`,
			sz, sz, bgColor,
			sz, sz, 1, sz-1, 1, sz+1, sz-1, 1, sz+1, 1, fgColor)
	case "dkUpDiag": // Dark Upward Diagonal
		inner = fmt.Sprintf(
			`    <rect width="%d" height="%d" fill="%s"/>
    <path d="M0 %d L%d 0 M-%d %d L%d -%d M%d %d L%d %d" stroke="%s" stroke-width="2"/>`,
			sz, sz, bgColor,
			sz, sz, 1, 1, 1, 1, sz-1, sz+1, sz+1, sz-1, fgColor)
	case "dkDnDiag": // Dark Downward Diagonal
		inner = fmt.Sprintf(
			`    <rect width="%d" height="%d" fill="%s"/>
    <path d="M0 0 L%d %d M-%d %d L%d %d M%d -%d L%d %d" stroke="%s" stroke-width="2"/>`,
			sz, sz, bgColor,
			sz, sz, 1, sz-1, 1, sz+1, sz-1, 1, sz+1, 1, fgColor)
	case "ltHorz": // Light Horizontal
		inner = fmt.Sprintf(
			`    <rect width="%d" height="%d" fill="%s"/>
    <line x1="0" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>`,
			sz, sz, bgColor,
			sz/2, sz, sz/2, fgColor)
	case "ltVert": // Light Vertical
		inner = fmt.Sprintf(
			`    <rect width="%d" height="%d" fill="%s"/>
    <line x1="%d" y1="0" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>`,
			sz, sz, bgColor,
			sz/2, sz/2, sz, fgColor)
	case "smGrid": // Small Grid
		inner = fmt.Sprintf(
			`    <rect width="%d" height="%d" fill="%s"/>
    <line x1="0" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
    <line x1="%d" y1="0" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>`,
			sz, sz, bgColor,
			sz/2, sz, sz/2, fgColor,
			sz/2, sz/2, sz, fgColor)
	case "diagCross": // Diagonal Cross
		inner = fmt.Sprintf(
			`    <rect width="%d" height="%d" fill="%s"/>
    <path d="M0 0 L%d %d M0 %d L%d 0" stroke="%s" stroke-width="1"/>`,
			sz, sz, bgColor,
			sz, sz, sz, sz, fgColor)
	default:
		// Fallback: light upward diagonal
		inner = fmt.Sprintf(
			`    <rect width="%d" height="%d" fill="%s"/>
    <path d="M0 %d L%d 0 M-%d %d L%d -%d M%d %d L%d %d" stroke="%s" stroke-width="1"/>`,
			sz, sz, bgColor,
			sz, sz, 1, 1, 1, 1, sz-1, sz+1, sz+1, sz-1, fgColor)
	}

	return fmt.Sprintf(`    <pattern id="%s" width="%d" height="%d" patternUnits="userSpaceOnUse">
%s
    </pattern>
`, id, sz, sz, inner)
}

// renderPieChartSVG renders a pie chart as an SVG fragment.
func renderPieChartSVG(data *ChartData, x, y, w, h int) string {
	if len(data.Series) == 0 || len(data.Series[0].Values) == 0 {
		return ""
	}

	var out strings.Builder

	// Font sizes
	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 25
	}
	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 14
	}

	// Layout
	legendItemH := 0
	legendBoxSize := 9
	charW := float64(labelFS) * 0.55

	if data.HasLegend {
		legendItemH = labelFS + 10
	}
	topMargin := 8
	bottomMargin := 8
	leftMargin := 8
	rightMargin := 8
	titleH := titleFS + 10

	topUsed := topMargin + titleH
	bottomUsed := bottomMargin
	leftUsed := leftMargin
	rightUsed := rightMargin

	var legendStartX, legendStartY float64
	var totalLegendW, totalLegendH float64

	if data.HasLegend && len(data.Categories) > 0 {
		if data.LegendPos == "r" || data.LegendPos == "l" {
			maxTextW := 0.0
			for _, catName := range data.Categories {
				textW := float64(len(catName)) * charW
				if textW > maxTextW {
					maxTextW = textW
				}
			}
			totalLegendW = float64(legendBoxSize) + 4 + maxTextW + 16
			totalLegendH = float64(len(data.Categories)) * float64(legendItemH)
			if data.LegendPos == "r" {
				rightUsed += int(totalLegendW)
			} else {
				leftUsed += int(totalLegendW)
			}
		} else {
			maxRowW := 0.0
			curRowW := 0.0
			rowsCount := 1
			for _, catName := range data.Categories {
				itemW := float64(legendBoxSize) + 4 + float64(len(catName))*charW + 16
				if curRowW+itemW > float64(w-leftMargin-rightMargin) && curRowW > 0 {
					if curRowW > maxRowW {
						maxRowW = curRowW
					}
					curRowW = itemW
					rowsCount++
				} else {
					curRowW += itemW
				}
			}
			if curRowW > maxRowW {
				maxRowW = curRowW
			}

			totalLegendW = maxRowW
			totalLegendH = float64(rowsCount) * float64(legendItemH)
			if data.LegendPos == "t" {
				topUsed += int(totalLegendH)
			} else {
				bottomUsed += int(totalLegendH)
			}
		}
	}

	pieAreaY := y + topUsed
	pieAreaH := h - topUsed - bottomUsed
	pieAreaX := x + leftUsed
	pieAreaW := w - leftUsed - rightUsed

	if pieAreaW < 40 || pieAreaH < 40 {
		pieAreaW = 40
		pieAreaH = 40
	}

	// Generate unique chart ID for pattern naming (use position)
	chartID := fmt.Sprintf("%d_%d", x, y)

	// Generate SVG fill and effect definitions if needed
	fillDefs, patternIDs, gradientIDs, filterIDs := generateSVGFillAndEffectDefs(data, chartID)
	if fillDefs != "" {
		out.WriteString(fillDefs)
	}

	// Background
	bgColor := "#F5F5F5"
	if data.BgColor != "" {
		bgColor = data.BgColor
	}
	bgFillAttr := fmt.Sprintf(`fill="%s"`, bgColor)
	if data.BgIsGradient {
		bgFillAttr = fmt.Sprintf(`fill="url(#chartBgGrad%s)"`, chartID)
	}
	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" %s rx="0"/>
`, x, y, w, h, bgFillAttr))

	// Detect dark background for text colour
	isDarkBg := isDarkColor(bgColor)
	if data.BgIsGradient && len(data.BgGradStops) > 0 {
		isDarkBg = isDarkColor(data.BgGradStops[0].Color)
	}
	titleColor := "#595959"
	legendTextColor := "#595959"
	if isDarkBg {
		titleColor = "#FFFFFF"
		legendTextColor = "#CCCCCC"
	}

	// Title
	titleX := x + w/2
	titleY := y + topMargin + titleFS
	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, titleX, titleY, titleFS, titleColor, escapeXMLText(data.Title)))

	// Calculate pie center and radius
	cx := float64(pieAreaX) + float64(pieAreaW)/2.0
	cy := float64(pieAreaY) + float64(pieAreaH)/2.0
	// Reduce radius when explosion is used to keep slices within bounds
	radiusScale := 0.85
	if data.Explosion > 0 {
		radiusScale = 0.70 // leave room for exploded slices
	}
	radius := math.Min(float64(pieAreaW), float64(pieAreaH)) / 2.0 * radiusScale

	// Compute total
	values := data.Series[0].Values
	total := 0.0
	for _, v := range values {
		total += v
	}
	if total == 0 {
		return ""
	}

	// Determine colors for each slice (pie charts use category-based colors)
	accentColors := getChartAccentColors(data.ThemeColors)
	sliceColors := make([]string, len(values))
	for i := range values {
		sliceColors[i] = accentColors[i%len(accentColors)]
	}

	// Draw pie slices
	startAngle := data.FirstSliceAng * math.Pi / 180.0 // convert degrees to radians
	// PowerPoint default: 0 degrees = top (12 o'clock), going clockwise
	// In SVG: 0 radians = right (3 o'clock). Offset by -90 degrees
	startAngle -= math.Pi / 2.0

	// Explosion offset: each slice center is moved outward along its mid-angle
	explosionOffset := 0.0
	if data.Explosion > 0 {
		explosionOffset = radius * data.Explosion / 100.0
	}

	for i, val := range values {
		if val <= 0 {
			continue
		}
		sliceAngle := 2 * math.Pi * val / total

		endAngle := startAngle + sliceAngle
		midAngle := startAngle + sliceAngle/2.0
		largeArc := 0
		if sliceAngle > math.Pi {
			largeArc = 1
		}

		// Apply explosion offset
		sliceCx := cx
		sliceCy := cy
		if explosionOffset > 0 {
			sliceCx += explosionOffset * math.Cos(midAngle)
			sliceCy += explosionOffset * math.Sin(midAngle)
		}

		x1 := sliceCx + radius*math.Cos(startAngle)
		y1 := sliceCy + radius*math.Sin(startAngle)
		x2 := sliceCx + radius*math.Cos(endAngle)
		y2 := sliceCy + radius*math.Sin(endAngle)

		pathD := fmt.Sprintf("M%.1f %.1f L%.1f %.1f A%.1f %.1f 0 %d 1 %.1f %.1f Z",
			sliceCx, sliceCy, x1, y1, radius, radius, largeArc, x2, y2)

		// Determine fill: pattern, gradient or solid
		fillAttr := fmt.Sprintf(`fill="%s"`, sliceColors[i])
		strokeAttr := `stroke="#FFFFFF" stroke-width="1.5"`
		filterAttr := ""

		if patID, ok := patternIDs[i]; ok {
			fillAttr = fmt.Sprintf(`fill="url(#%s)"`, patID)
		} else if gradID, ok := gradientIDs[i]; ok {
			fillAttr = fmt.Sprintf(`fill="url(#%s)"`, gradID)
		} else if i < len(data.DPtFills) && data.DPtFills[i].IsSolid && data.DPtFills[i].SolidColor != "" {
			fillAttr = fmt.Sprintf(`fill="%s"`, data.DPtFills[i].SolidColor)
		}

		if fid, ok := filterIDs[i]; ok {
			filterAttr = fmt.Sprintf(`filter="url(#%s)"`, fid)
		}

		// Override stroke if dPt specifies it
		if i < len(data.DPtFills) && data.DPtFills[i].StrokeColor != "" {
			sw := data.DPtFills[i].StrokeWidth
			if sw <= 0 {
				sw = 1.5
			}
			strokeAttr = fmt.Sprintf(`stroke="%s" stroke-width="%.1f"`, data.DPtFills[i].StrokeColor, sw)
		}

		out.WriteString(fmt.Sprintf(`  <path d="%s" %s %s %s/>
`, pathD, fillAttr, strokeAttr, filterAttr))

		// Percentage label
		if data.ShowPercent {
			pct := val / total * 100
			labelR := radius * 0.65 // place labels at 65% of radius
			lx := sliceCx + labelR*math.Cos(midAngle)
			ly := sliceCy + labelR*math.Sin(midAngle)
			out.WriteString(fmt.Sprintf(`  <text x="%.1f" y="%.1f" font-size="%d" fill="#FFFFFF" font-family="Segoe UI, sans-serif" text-anchor="middle" dominant-baseline="central">%.0f%%</text>
`, lx, ly, labelFS, pct))
		}

		startAngle = endAngle
	}

	// Legend
	if data.HasLegend && len(data.Categories) > 0 {
		if data.LegendPos == "r" || data.LegendPos == "l" {
			if data.LegendPos == "r" {
				legendStartX = float64(x + w - rightUsed + rightMargin)
			} else {
				legendStartX = float64(x + leftMargin)
			}
			legendStartY = float64(y+topUsed) + (float64(pieAreaH)-totalLegendH)/2.0

			curY := legendStartY + float64(legendItemH)/2.0
			for i, catName := range data.Categories {
				legendFill := sliceColors[0]
				if i < len(sliceColors) {
					legendFill = sliceColors[i]
				}

				boxY := int(curY) - legendBoxSize/2
				if patID, ok := patternIDs[i]; ok {
					out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="url(#%s)"/>
`, legendStartX, boxY, legendBoxSize, legendBoxSize, patID))
					if i < len(data.DPtFills) && data.DPtFills[i].StrokeColor != "" {
						out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="none" stroke="%s" stroke-width="1"/>
`, legendStartX, boxY, legendBoxSize, legendBoxSize, data.DPtFills[i].StrokeColor))
					}
				} else {
					out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="%s"/>
`, legendStartX, boxY, legendBoxSize, legendBoxSize, legendFill))
				}

				textStartX := legendStartX + float64(legendBoxSize) + 4
				out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%.0f" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" dominant-baseline="central">%s</text>
`, textStartX, curY, labelFS, legendTextColor, escapeXMLText(catName)))
				curY += float64(legendItemH)
			}
		} else {
			if data.LegendPos == "t" {
				legendStartY = float64(y + topMargin + titleH + labelFS)
			} else {
				legendStartY = float64(y + h - bottomMargin - 2)
			}
			legendStartX = float64(pieAreaX) + (float64(pieAreaW)-totalLegendW)/2.0

			curX := legendStartX
			for i, catName := range data.Categories {
				legendFill := sliceColors[0]
				if i < len(sliceColors) {
					legendFill = sliceColors[i]
				}

				boxY := int(legendStartY) - legendBoxSize + 1
				if patID, ok := patternIDs[i]; ok {
					out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="url(#%s)"/>
`, curX, boxY, legendBoxSize, legendBoxSize, patID))
					if i < len(data.DPtFills) && data.DPtFills[i].StrokeColor != "" {
						out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="none" stroke="%s" stroke-width="1"/>
`, curX, boxY, legendBoxSize, legendBoxSize, data.DPtFills[i].StrokeColor))
					}
				} else {
					out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="%s"/>
`, curX, boxY, legendBoxSize, legendBoxSize, legendFill))
				}
				curX += float64(legendBoxSize) + 4

				out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%.0f" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif">%s</text>
`, curX, legendStartY, labelFS, legendTextColor, escapeXMLText(catName)))
				curX += float64(len(catName))*charW + 16
			}
		}
	}

	return out.String()
}

// ====================================================================
// Doughnut Chart Renderer
// ====================================================================

// renderDoughnutChartSVG renders a doughnut chart as an SVG fragment.
// It is structurally identical to renderPieChartSVG but draws annular arcs
// with an inner radius determined by HoleSize (percentage of outer radius).
func renderDoughnutChartSVG(data *ChartData, x, y, w, h int) string {
	if len(data.Series) == 0 || len(data.Series[0].Values) == 0 {
		return ""
	}

	var out strings.Builder

	// Font sizes
	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 25
	}
	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 14
	}

	// Layout
	legendItemH := 0
	legendBoxSize := 9
	charW := float64(labelFS) * 0.55

	if data.HasLegend {
		legendItemH = labelFS + 10
	}
	topMargin := 8
	bottomMargin := 8
	leftMargin := 8
	rightMargin := 8
	titleH := titleFS + 10

	topUsed := topMargin + titleH
	bottomUsed := bottomMargin
	leftUsed := leftMargin
	rightUsed := rightMargin

	var legendStartX, legendStartY float64
	var totalLegendW, totalLegendH float64

	if data.HasLegend && len(data.Categories) > 0 {
		if data.LegendPos == "r" || data.LegendPos == "l" {
			maxTextW := 0.0
			for _, catName := range data.Categories {
				textW := float64(len(catName)) * charW
				if textW > maxTextW {
					maxTextW = textW
				}
			}
			totalLegendW = float64(legendBoxSize) + 4 + maxTextW + 16
			totalLegendH = float64(len(data.Categories)) * float64(legendItemH)
			if data.LegendPos == "r" {
				rightUsed += int(totalLegendW)
			} else {
				leftUsed += int(totalLegendW)
			}
		} else {
			maxRowW := 0.0
			curRowW := 0.0
			rowsCount := 1
			for _, catName := range data.Categories {
				itemW := float64(legendBoxSize) + 4 + float64(len(catName))*charW + 16
				if curRowW+itemW > float64(w-leftMargin-rightMargin) && curRowW > 0 {
					if curRowW > maxRowW {
						maxRowW = curRowW
					}
					curRowW = itemW
					rowsCount++
				} else {
					curRowW += itemW
				}
			}
			if curRowW > maxRowW {
				maxRowW = curRowW
			}

			totalLegendW = maxRowW
			totalLegendH = float64(rowsCount) * float64(legendItemH)
			if data.LegendPos == "t" {
				topUsed += int(totalLegendH)
			} else {
				bottomUsed += int(totalLegendH)
			}
		}
	}

	pieAreaY := y + topUsed
	pieAreaH := h - topUsed - bottomUsed
	pieAreaX := x + leftUsed
	pieAreaW := w - leftUsed - rightUsed

	if pieAreaW < 40 || pieAreaH < 40 {
		pieAreaW = 40
		pieAreaH = 40
	}

	// Generate unique chart ID for pattern naming
	chartID := fmt.Sprintf("%d_%d", x, y)

	// Generate SVG fill and effect definitions if needed
	fillDefs, patternIDs, gradientIDs, filterIDs := generateSVGFillAndEffectDefs(data, chartID)
	if fillDefs != "" {
		out.WriteString(fillDefs)
	}

	// Background
	bgColor := "#F5F5F5"
	if data.BgColor != "" {
		bgColor = data.BgColor
	}
	bgFillAttr := fmt.Sprintf(`fill="%s"`, bgColor)
	if data.BgIsGradient {
		bgFillAttr = fmt.Sprintf(`fill="url(#chartBgGrad%s)"`, chartID)
	}

	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" %s rx="0"/>
`, x, y, w, h, bgFillAttr))

	isDarkBg := isDarkColor(bgColor)
	if data.BgIsGradient && len(data.BgGradStops) > 0 {
		isDarkBg = isDarkColor(data.BgGradStops[0].Color)
	}
	titleColor := "#595959"
	legendTextColor := "#595959"
	if isDarkBg {
		titleColor = "#FFFFFF"
		legendTextColor = "#CCCCCC"
	}

	// Title
	titleX := x + w/2
	titleY := y + topMargin + titleFS
	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, titleX, titleY, titleFS, titleColor, escapeXMLText(data.Title)))

	// Calculate center and radii
	cx := float64(pieAreaX) + float64(pieAreaW)/2.0
	cy := float64(pieAreaY) + float64(pieAreaH)/2.0
	outerRadius := math.Min(float64(pieAreaW), float64(pieAreaH)) / 2.0 * 0.85
	innerRadius := outerRadius * data.HoleSize / 100.0

	// Compute total
	values := data.Series[0].Values
	total := 0.0
	for _, v := range values {
		total += v
	}
	if total == 0 {
		return ""
	}

	// Slice colors
	accentColors := getChartAccentColors(data.ThemeColors)
	sliceColors := make([]string, len(values))
	for i := range values {
		sliceColors[i] = accentColors[i%len(accentColors)]
	}

	// Draw doughnut slices as annular arcs
	startAngle := data.FirstSliceAng*math.Pi/180.0 - math.Pi/2.0

	for i, val := range values {
		if val <= 0 {
			continue
		}
		sliceAngle := 2 * math.Pi * val / total
		endAngle := startAngle + sliceAngle
		midAngle := startAngle + sliceAngle/2.0
		largeArc := 0
		if sliceAngle > math.Pi {
			largeArc = 1
		}

		// Outer arc endpoints
		ox1 := cx + outerRadius*math.Cos(startAngle)
		oy1 := cy + outerRadius*math.Sin(startAngle)
		ox2 := cx + outerRadius*math.Cos(endAngle)
		oy2 := cy + outerRadius*math.Sin(endAngle)

		// Inner arc endpoints (reversed direction)
		ix1 := cx + innerRadius*math.Cos(endAngle)
		iy1 := cy + innerRadius*math.Sin(endAngle)
		ix2 := cx + innerRadius*math.Cos(startAngle)
		iy2 := cy + innerRadius*math.Sin(startAngle)

		// Annular arc path: outer arc CW, line to inner, inner arc CCW, close
		pathD := fmt.Sprintf("M%.1f %.1f A%.1f %.1f 0 %d 1 %.1f %.1f L%.1f %.1f A%.1f %.1f 0 %d 0 %.1f %.1f Z",
			ox1, oy1, outerRadius, outerRadius, largeArc, ox2, oy2,
			ix1, iy1, innerRadius, innerRadius, largeArc, ix2, iy2)

		// Determine fill: pattern, gradient, or solid
		fillAttr := fmt.Sprintf(`fill="%s"`, sliceColors[i])
		strokeAttr := `stroke="#FFFFFF" stroke-width="1.5"`
		filterAttr := ""

		if patID, ok := patternIDs[i]; ok {
			fillAttr = fmt.Sprintf(`fill="url(#%s)"`, patID)
		} else if gradID, ok := gradientIDs[i]; ok {
			fillAttr = fmt.Sprintf(`fill="url(#%s)"`, gradID)
		} else if i < len(data.DPtFills) && data.DPtFills[i].IsSolid && data.DPtFills[i].SolidColor != "" {
			fillAttr = fmt.Sprintf(`fill="%s"`, data.DPtFills[i].SolidColor)
		}

		if fid, ok := filterIDs[i]; ok {
			filterAttr = fmt.Sprintf(`filter="url(#%s)"`, fid)
		}

		// Override stroke if dPt specifies it
		if i < len(data.DPtFills) && data.DPtFills[i].StrokeColor != "" {
			sw := data.DPtFills[i].StrokeWidth
			if sw <= 0 {
				sw = 1.5
			}
			strokeAttr = fmt.Sprintf(`stroke="%s" stroke-width="%.1f"`, data.DPtFills[i].StrokeColor, sw)
		}

		out.WriteString(fmt.Sprintf(`  <path d="%s" %s %s %s/>
`, pathD, fillAttr, strokeAttr, filterAttr))

		// Percentage label (placed at midpoint of annular ring)
		if data.ShowPercent {
			pct := val / total * 100
			labelR := (outerRadius + innerRadius) / 2.0
			lx := cx + labelR*math.Cos(midAngle)
			ly := cy + labelR*math.Sin(midAngle)
			out.WriteString(fmt.Sprintf(`  <text x="%.1f" y="%.1f" font-size="%d" fill="#FFFFFF" font-family="Segoe UI, sans-serif" text-anchor="middle" dominant-baseline="central">%.0f%%</text>
`, lx, ly, labelFS, pct))
		}

		startAngle = endAngle
	}

	// Legend (same as pie chart)
	if data.HasLegend && len(data.Categories) > 0 {
		if data.LegendPos == "r" || data.LegendPos == "l" {
			if data.LegendPos == "r" {
				legendStartX = float64(x + w - rightUsed + rightMargin)
			} else {
				legendStartX = float64(x + leftMargin)
			}
			legendStartY = float64(y+topUsed) + (float64(pieAreaH)-totalLegendH)/2.0

			curY := legendStartY + float64(legendItemH)/2.0
			for i, catName := range data.Categories {
				legendFill := sliceColors[0]
				if i < len(sliceColors) {
					legendFill = sliceColors[i]
				}

				boxY := int(curY) - legendBoxSize/2
				if patID, ok := patternIDs[i]; ok {
					out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="url(#%s)"/>
`, legendStartX, boxY, legendBoxSize, legendBoxSize, patID))
					if i < len(data.DPtFills) && data.DPtFills[i].StrokeColor != "" {
						out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="none" stroke="%s" stroke-width="1"/>
`, legendStartX, boxY, legendBoxSize, legendBoxSize, data.DPtFills[i].StrokeColor))
					}
				} else {
					out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="%s"/>
`, legendStartX, boxY, legendBoxSize, legendBoxSize, legendFill))
				}

				textStartX := legendStartX + float64(legendBoxSize) + 4
				out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%.0f" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" dominant-baseline="central">%s</text>
`, textStartX, curY, labelFS, legendTextColor, escapeXMLText(catName)))
				curY += float64(legendItemH)
			}
		} else {
			if data.LegendPos == "t" {
				legendStartY = float64(y + topMargin + titleH + labelFS)
			} else {
				legendStartY = float64(y+h-bottomMargin) - totalLegendH + float64(labelFS)
			}

			type legendRow struct {
				w     float64
				items []int
			}
			var currentRow legendRow
			var rows []legendRow
			var curRowW float64

			for i, catName := range data.Categories {
				itemW := float64(legendBoxSize) + 4 + float64(len(catName))*charW + 16
				if curRowW+itemW > float64(w-leftMargin-rightMargin) && len(currentRow.items) > 0 {
					currentRow.w = curRowW
					rows = append(rows, currentRow)
					currentRow = legendRow{}
					curRowW = 0
				}
				currentRow.items = append(currentRow.items, i)
				curRowW += itemW
			}
			if len(currentRow.items) > 0 {
				currentRow.w = curRowW
				rows = append(rows, currentRow)
			}

			curY := legendStartY
			for _, r := range rows {
				curX := float64(x) + (float64(w)-r.w)/2.0
				for _, i := range r.items {
					catName := data.Categories[i]
					legendFill := sliceColors[0]
					if i < len(sliceColors) {
						legendFill = sliceColors[i]
					}

					boxY := int(curY) - legendBoxSize + 1
					if patID, ok := patternIDs[i]; ok {
						out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="url(#%s)"/>
`, curX, boxY, legendBoxSize, legendBoxSize, patID))
						if i < len(data.DPtFills) && data.DPtFills[i].StrokeColor != "" {
							out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="none" stroke="%s" stroke-width="1"/>
`, curX, boxY, legendBoxSize, legendBoxSize, data.DPtFills[i].StrokeColor))
						}
					} else {
						out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%d" width="%d" height="%d" fill="%s"/>
`, curX, boxY, legendBoxSize, legendBoxSize, legendFill))
					}
					curX += float64(legendBoxSize) + 4

					out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%.0f" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif">%s</text>
`, curX, curY, labelFS, legendTextColor, escapeXMLText(catName)))
					curX += float64(len(catName))*charW + 16
				}
				curY += float64(legendItemH)
			}
		}
	}

	return out.String()
}

// niceAxisMax computes a nice maximum value for the axis, matching PowerPoint behavior.
// PowerPoint adds headroom above the max data value and rounds to a clean tick interval.
func niceAxisMax(maxVal float64) float64 {
	if maxVal <= 0 {
		return 10
	}

	// PowerPoint's auto-scale: pick a tick interval, then set max = ceil(maxVal/interval)*interval + buffer
	interval := niceTickInterval(maxVal)
	niceMax := math.Ceil(maxVal/interval) * interval

	// Add one tick of headroom if max lands exactly on a data point
	if niceMax <= maxVal {
		niceMax += interval
	}

	return niceMax
}

// niceTickInterval determines a good tick interval for the axis (PowerPoint-style).
func niceTickInterval(maxVal float64) float64 {
	if maxVal <= 0 {
		return 1
	}

	// Approximate: we want ~5-10 ticks
	rawInterval := maxVal / 8.0
	magnitude := math.Pow(10, math.Floor(math.Log10(rawInterval)))

	normalized := rawInterval / magnitude
	var niceInterval float64
	switch {
	case normalized <= 1:
		niceInterval = 1
	case normalized <= 2:
		niceInterval = 2
	case normalized <= 5:
		niceInterval = 5
	default:
		niceInterval = 10
	}

	return niceInterval * magnitude
}

// isDarkColor returns true if a hex colour is considered "dark" (low luminance).
func isDarkColor(hex string) bool {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return false
	}
	r, _ := strconv.ParseInt(hex[0:2], 16, 64)
	g, _ := strconv.ParseInt(hex[2:4], 16, 64)
	b, _ := strconv.ParseInt(hex[4:6], 16, 64)
	// Relative luminance (sRGB)
	lum := 0.299*float64(r) + 0.587*float64(g) + 0.114*float64(b)
	return lum < 128
}

// formatAxisValue formats a numeric axis label.
func formatAxisValue(val float64) string {
	if val == math.Floor(val) {
		return strconv.Itoa(int(val))
	}
	return strconv.FormatFloat(val, 'f', 1, 64)
}

// processGraphicFrame processes a <p:graphicFrame> element, detecting charts and tables.
// Returns multiple RenderShapes and RenderConnectors to support SmartArt diagrams.
func processGraphicFrame(gf OOXMLGraphicFrame, ctx *slideContext) ([]RenderShape, []RenderConnector) {
	// Check for table first
	if gf.Graphic.GraphicData.URI == tableURI && gf.Graphic.GraphicData.Tbl != nil {
		rs := processTableGraphicFrame(gf, ctx)
		if rs != nil {
			return []RenderShape{*rs}, nil
		}
		return nil, nil
	}

	// Check if this is a chart
	chartURI := "http://schemas.openxmlformats.org/drawingml/2006/chart"
	chartExURI := "http://schemas.microsoft.com/office/drawing/2014/chartex"
	if (gf.Graphic.GraphicData.URI == chartURI || gf.Graphic.GraphicData.URI == chartExURI) && gf.Graphic.GraphicData.Chart != nil {
		rID := gf.Graphic.GraphicData.Chart.RID
		if rID != "" {
			return []RenderShape{{
				XmlID:     int(gf.NvGraphicFramePr.CNvPr.ID),
				Name:      gf.NvGraphicFramePr.CNvPr.Name,
				ShapeType: "chart",
				X:         gf.Xfrm.Off.X,
				Y:         gf.Xfrm.Off.Y,
				W:         gf.Xfrm.Ext.Cx,
				H:         gf.Xfrm.Ext.Cy,
				EmbedRID:  rID,
			}}, nil
		}
	}

	// Check if this is a SmartArt diagram
	diagramURI := "http://schemas.openxmlformats.org/drawingml/2006/diagram"
	if gf.Graphic.GraphicData.URI == diagramURI {
		return processSmartArtDiagram(gf, ctx)
	}

	return nil, nil
}

// processSmartArtDiagram parses the corresponding drawingX.xml file for a SmartArt graphic frame.
func processSmartArtDiagram(gf OOXMLGraphicFrame, ctx *slideContext) ([]RenderShape, []RenderConnector) {
	if ctx == nil {
		return nil, nil
	}

	var drawingPath string
	for _, rel := range ctx.slideRels {
		if rel.Type == "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" {
			drawingPath = rel.Target
			break
		}
	}
	if drawingPath == "" {
		return nil, nil
	}

	fullPath := path.Clean(path.Join("ppt/slides", drawingPath))

	// Load relationships for the drawing (e.g. for images inside the SmartArt)
	// The drawing file (ppt/diagrams/drawingX.xml) has its own relationships in ppt/diagrams/_rels/drawingX.xml
	dir := path.Dir(fullPath)
	base := path.Base(fullPath)
	relsPath := path.Join(dir, "_rels", base+".rels")
	drawingRels := loadRelationships(ctx.zipReader, relsPath)

	// Create a new context with the drawing's relationships so images can be resolved
	drawingCtx := *ctx
	if len(drawingRels) > 0 {
		drawingCtx.slideRels = drawingRels
	}

	var drawingData []byte
	for _, f := range ctx.zipReader.File {
		if f.Name == fullPath {
			rc, err := f.Open()
			if err == nil {
				drawingData, _ = io.ReadAll(rc)
				rc.Close()
			}
			break
		}
	}
	if len(drawingData) == 0 {
		return nil, nil
	}

	// Convert dsp: tags to p: tags so OOXMLSpTree can parse it directly
	modData := bytes.ReplaceAll(drawingData, []byte("dsp:"), []byte("p:"))
	modData = bytes.ReplaceAll(modData, []byte("</dsp:"), []byte("</p:"))

	type drawingRoot struct {
		SpTree OOXMLSpTree `xml:"spTree"`
	}
	var root drawingRoot
	if err := xml.Unmarshal(modData, &root); err != nil {
		return nil, nil
	}

	// Add offsets directly to the parsed nodes so that their internal geometry
	// paths (CustPathD) are also generated with the correct absolute coordinates.
	offsetX := gf.Xfrm.Off.X
	offsetY := gf.Xfrm.Off.Y
	if offsetX != 0 || offsetY != 0 {
		for i := range root.SpTree.Shapes {
			root.SpTree.Shapes[i].SpPr.Xfrm.Off.X += offsetX
			root.SpTree.Shapes[i].SpPr.Xfrm.Off.Y += offsetY
			if root.SpTree.Shapes[i].TxXfrm != nil {
				root.SpTree.Shapes[i].TxXfrm.Off.X += offsetX
				root.SpTree.Shapes[i].TxXfrm.Off.Y += offsetY
			}
		}
		for i := range root.SpTree.Groups {
			root.SpTree.Groups[i].GrpSpPr.Xfrm.Off.X += offsetX
			root.SpTree.Groups[i].GrpSpPr.Xfrm.Off.Y += offsetY
		}
		for i := range root.SpTree.Pictures {
			root.SpTree.Pictures[i].SpPr.Xfrm.Off.X += offsetX
			root.SpTree.Pictures[i].SpPr.Xfrm.Off.Y += offsetY
		}
		for i := range root.SpTree.Connectors {
			root.SpTree.Connectors[i].SpPr.Xfrm.Off.X += offsetX
			root.SpTree.Connectors[i].SpPr.Xfrm.Off.Y += offsetY
		}
		for i := range root.SpTree.GraphicFrames {
			root.SpTree.GraphicFrames[i].Xfrm.Off.X += offsetX
			root.SpTree.GraphicFrames[i].Xfrm.Off.Y += offsetY
		}
	}

	shapes, conns := convertSpTree(&root.SpTree, &drawingCtx)

	return shapes, conns
}

// loadAndRenderChart loads chart XML from the PPTX zip and renders it as SVG.
func loadAndRenderChart(s RenderShape, ctx *slideContext) string {
	if s.EmbedRID == "" || ctx == nil {
		return ""
	}

	// Resolve chart path from slide relationships
	chartPath := ""
	for _, rel := range ctx.slideRels {
		if rel.ID == s.EmbedRID {
			chartPath = rel.Target
			break
		}
	}
	if chartPath == "" {
		return ""
	}

	// Resolve relative path: ../charts/chart1.xml â†’ ppt/charts/chart1.xml
	// slideRels target is relative to the slide file (ppt/slides/slideN.xml)
	fullChartPath := path.Join("ppt/slides", chartPath)
	fullChartPath = path.Clean(fullChartPath)

	// Read chart XML from zip
	var chartXML []byte
	for _, f := range ctx.zipReader.File {
		if f.Name == fullChartPath {
			rc, err := f.Open()
			if err != nil {
				return ""
			}
			data, err := io.ReadAll(rc)
			rc.Close()
			if err != nil {
				return ""
			}
			chartXML = data
			break
		}
	}

	if len(chartXML) == 0 {
		return ""
	}

	var themeColors map[string]string
	if ctx != nil {
		themeColors = ctx.themeColors
	}

	var chartData *ChartData
	var err error
	if bytes.Contains(chartXML, []byte("cx:chartSpace")) {
		var overrideColors []string
		relsPath := path.Dir(fullChartPath) + "/_rels/" + path.Base(fullChartPath) + ".rels"
		for _, f := range ctx.zipReader.File {
			if f.Name == relsPath {
				rc, err := f.Open()
				if err == nil {
					relsData, _ := io.ReadAll(rc)
					rc.Close()
					var rels OOXMLRelationships
					if err := xml.Unmarshal(relsData, &rels); err == nil {
						for _, r := range rels.Rels {
							if r.Type == "http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" {
								colorPath := path.Join(path.Dir(fullChartPath), r.Target)
								colorPath = path.Clean(colorPath)
								for _, cf := range ctx.zipReader.File {
									if cf.Name == colorPath {
										crc, cerr := cf.Open()
										if cerr == nil {
											cData, _ := io.ReadAll(crc)
											crc.Close()
											type colorStyle struct {
												SchemeClrs []struct {
													Val string `xml:"val,attr"`
												} `xml:"schemeClr"`
												SrgbClrs []struct {
													Val string `xml:"val,attr"`
												} `xml:"srgbClr"`
											}
											var cs colorStyle
											if err := xml.Unmarshal(cData, &cs); err == nil {
												for _, sc := range cs.SchemeClrs {
													if hex, ok := themeColors[sc.Val]; ok {
														overrideColors = append(overrideColors, hex)
													} else {
														if sc.Val == "dk1" || sc.Val == "lt1" {
															if hex, ok := themeColors[sc.Val]; ok {
																overrideColors = append(overrideColors, hex)
															} else {
																overrideColors = append(overrideColors, "#000000")
															}
														}
													}
												}
												for _, sr := range cs.SrgbClrs {
													overrideColors = append(overrideColors, "#"+sr.Val)
												}
											}
										}
										break
									}
								}
								break
							}
						}
					}
				}
				break
			}
		}

		chartData, err = parseExtendedChartXML(chartXML, themeColors)
		if chartData != nil {
			if len(overrideColors) > 0 {
				chartData.CustomColors = overrideColors
			}
			fmt.Printf("Parsed cx:chartSpace type: %s, categories: %d, series: %d\n", chartData.ChartType, len(chartData.Categories), len(chartData.Series))
		}
	} else {
		fmt.Println("--- CHART XML ---")
		fmt.Println(string(chartXML))
		fmt.Println("-----------------")
		chartData, err = parseChartXML(chartXML, themeColors)
	}
	if err != nil {
		fmt.Printf("chart parse err: %v\n", err)
		return ""
	}

	x := emuToPx(s.X)
	y := emuToPx(s.Y)
	w := emuToPx(s.W)
	h := emuToPx(s.H)

	svg := renderChartSVG(chartData, x, y, w, h)
	if bytes.Contains(chartXML, []byte("cx:chartSpace")) {
		fmt.Printf("Rendered SVG size: %d\n", len(svg))
	}
	return svg
}

// renderAreaChartSVG renders an area chart as an SVG fragment.
func renderAreaChartSVG(data *ChartData, x, y, w, h int) string {
	if len(data.Categories) == 0 {
		return ""
	}

	var out strings.Builder

	// Font sizes
	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 25
	}
	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 16
	}

	// Layout
	titleH := titleFS + 25
	legendH := 0
	if data.HasLegend {
		legendH = labelFS + 18
	}
	leftMargin := 33
	rightMargin := 15
	bottomMargin := labelFS + 28
	topMargin := 17

	plotX := x + leftMargin
	plotY := y + topMargin + titleH
	plotW := w - leftMargin - rightMargin
	plotH := h - topMargin - titleH - bottomMargin - legendH

	if plotW < 50 || plotH < 50 {
		return ""
	}

	// Find max value
	maxVal := 0.0
	for _, s := range data.Series {
		for _, v := range s.Values {
			if v > maxVal {
				maxVal = v
			}
		}
	}
	tickInterval := niceTickInterval(maxVal)
	niceMax := niceAxisMax(maxVal)
	if niceMax == 0 {
		niceMax = 10
	}
	numTicks := int(math.Round(niceMax / tickInterval))

	// Background
	lineBgColor := "#F5F5F5"
	if data.BgColor != "" {
		lineBgColor = data.BgColor
	}
	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="%s" rx="0"/>
`, x, y, w, h, lineBgColor))

	// Title
	titleX := x + w/2
	titleY := y + topMargin + titleFS + 2
	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="#595959" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, titleX, titleY, titleFS, escapeXMLText(data.Title)))

	// Gridlines + Y-axis labels
	gridColor := "#D9D9D9"
	axisLabelColor := "#595959"

	for i := 0; i <= numTicks; i++ {
		val := float64(i) * tickInterval
		yPos := plotY + plotH - int(float64(plotH)*(val/niceMax))

		if i > 0 {
			out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, plotX, yPos, plotX+plotW, yPos, gridColor))
		}

		labelStr := formatAxisValue(val)
		labelX := plotX - 5
		labelY := yPos + labelFS/3
		out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="end">%s</text>
`, labelX, labelY, labelFS, axisLabelColor, labelStr))
	}

	// Bottom axis line
	axisY := plotY + plotH
	out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, plotX, axisY, plotX+plotW, axisY, gridColor))

	// Draw areas for each series
	numCats := len(data.Categories)
	for _, ser := range data.Series {
		if len(ser.Values) == 0 {
			continue
		}
		var points []string

		// First point to bottom corner
		firstPx := float64(plotX) + 0.5*float64(plotW)/float64(numCats)
		points = append(points, fmt.Sprintf("%.1f,%.1f", firstPx, float64(plotY+plotH)))

		for catIdx := 0; catIdx < numCats && catIdx < len(ser.Values); catIdx++ {
			px := float64(plotX) + (float64(catIdx)+0.5)*float64(plotW)/float64(numCats)
			py := float64(plotY+plotH) - float64(plotH)*(ser.Values[catIdx]/niceMax)
			points = append(points, fmt.Sprintf("%.1f,%.1f", px, py))
		}

		// Last point back to bottom corner
		lastPx := float64(plotX) + (float64(len(ser.Values)-1)+0.5)*float64(plotW)/float64(numCats)
		points = append(points, fmt.Sprintf("%.1f,%.1f", lastPx, float64(plotY+plotH)))

		if len(points) > 0 {
			out.WriteString(fmt.Sprintf(`  <polyline points="%s" fill="%s" stroke="%s" stroke-width="2.5" stroke-linejoin="round"/>
`, strings.Join(points, " "), ser.Color, ser.Color))
		}
	}

	// X-axis category labels
	catLabelY := plotY + plotH + labelFS + 8
	for catIdx, catName := range data.Categories {
		catCenterX := float64(plotX) + (float64(catIdx)+0.5)*float64(plotW)/float64(numCats)
		out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, catCenterX, catLabelY, labelFS, axisLabelColor, escapeXMLText(catName)))
	}

	// Legend
	if data.HasLegend && len(data.Series) > 0 {
		legendY := plotY + plotH + bottomMargin + labelFS + 2
		charW := float64(labelFS) * 0.55

		totalLegendW := 0.0
		for _, ser := range data.Series {
			totalLegendW += 20 + 4 + float64(len(ser.Name))*charW + 16
		}
		legendStartX := float64(x) + (float64(w)-totalLegendW)/2.0

		curX := legendStartX
		for _, ser := range data.Series {
			// Small square swatch for area charts
			swatchY := legendY - 10
			out.WriteString(fmt.Sprintf(`  <rect x="%.0f" y="%.0f" width="12" height="12" fill="%s" fill-opacity="0.6" stroke="%s"/>
`, curX, float64(swatchY), ser.Color, ser.Color))
			curX += 16

			out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif">%s</text>
`, curX, legendY, labelFS, axisLabelColor, escapeXMLText(ser.Name)))
			curX += float64(len(ser.Name))*charW + 16
		}
	}

	return out.String()
}

// renderScatterChartSVG renders a scatter chart as an SVG fragment.
func renderScatterChartSVG(data *ChartData, x, y, w, h int) string {
	var out strings.Builder

	// Font sizes
	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 25
	}
	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 16
	}

	titleH := titleFS + 25
	legendH := 0
	if data.HasLegend {
		legendH = labelFS + 18
	}
	leftMargin := 45
	rightMargin := 15
	bottomMargin := labelFS + 28
	topMargin := 17

	plotX := x + leftMargin
	plotY := y + topMargin + titleH
	plotW := w - leftMargin - rightMargin
	plotH := h - topMargin - titleH - bottomMargin - legendH

	if plotW < 50 || plotH < 50 {
		return ""
	}

	// Background and Title
	lineBgColor := "#F5F5F5"
	if data.BgColor != "" {
		lineBgColor = data.BgColor
	}
	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="%s" rx="0"/>
`, x, y, w, h, lineBgColor))

	titleX := x + w/2
	titleY := y + topMargin + titleFS + 2
	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="#595959" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, titleX, titleY, titleFS, escapeXMLText(data.Title)))

	// Find domains
	minY, maxY, minX, maxX := math.MaxFloat64, -math.MaxFloat64, math.MaxFloat64, -math.MaxFloat64
	for _, s := range data.Series {
		for _, v := range s.Values {
			if v > maxY {
				maxY = v
			}
			if v < minY {
				minY = v
			}
		}
		for _, xv := range s.XValues {
			if xv > maxX {
				maxX = xv
			}
			if xv < minX {
				minX = xv
			}
		}
	}
	if minY == math.MaxFloat64 {
		minY = 0
	}
	if minX == math.MaxFloat64 {
		minX = 0
	}

	if maxY <= 0 {
		maxY = 10
	}
	if maxX <= 0 {
		maxX = 10
	}
	if minY >= 0 {
		minY = 0
	}
	if minX >= 0 {
		minX = 0
	}

	tickIntervalY := niceTickInterval(maxY)
	numTicksY := int(math.Round(maxY / tickIntervalY))
	niceMaxY := float64(numTicksY) * tickIntervalY
	if niceMaxY < maxY {
		niceMaxY += tickIntervalY
		numTicksY++
	}
	if niceMaxY == 0 {
		niceMaxY = 10
		numTicksY = 5
		tickIntervalY = 2
	}

	tickIntervalX := niceTickInterval(maxX)
	numTicksX := int(math.Round(maxX / tickIntervalX))
	niceMaxX := float64(numTicksX) * tickIntervalX
	if niceMaxX < maxX {
		niceMaxX += tickIntervalX
		numTicksX++
	}
	if niceMaxX == 0 {
		niceMaxX = 10
		numTicksX = 5
		tickIntervalX = 2
	}

	gridColor := "#D9D9D9"
	axisLabelColor := "#595959"

	// Y-axis grid & labels
	for i := 0; i <= numTicksY; i++ {
		val := float64(i) * tickIntervalY
		yPos := plotY + plotH - int(float64(plotH)*(val/niceMaxY))
		if i > 0 {
			out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, plotX, yPos, plotX+plotW, yPos, gridColor))
		}
		labelStr := formatAxisValue(val)
		labelX := plotX - 5
		labelY := yPos + labelFS/3
		out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="end">%s</text>
`, labelX, labelY, labelFS, axisLabelColor, labelStr))
	}

	// X-axis grid & labels
	for i := 0; i <= numTicksX; i++ {
		val := float64(i) * tickIntervalX
		xPos := plotX + int(float64(plotW)*(val/niceMaxX))
		if i > 0 {
			out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="1"/>
`, xPos, plotY, xPos, plotY+plotH, gridColor))
		}
		labelStr := formatAxisValue(val)
		labelY := plotY + plotH + labelFS + 5
		out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, xPos, labelY, labelFS, axisLabelColor, labelStr))
	}

	// Bottom/Left axis line
	axisY := plotY + plotH
	out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="2"/>
`, plotX, axisY, plotX+plotW, axisY, gridColor))
	out.WriteString(fmt.Sprintf(`  <line x1="%d" y1="%d" x2="%d" y2="%d" stroke="%s" stroke-width="2"/>
`, plotX, plotY, plotX, plotY+plotH, gridColor))

	// Draw points
	for _, ser := range data.Series {
		for pointIdx := 0; pointIdx < len(ser.Values) && pointIdx < len(ser.XValues); pointIdx++ {
			yVal := ser.Values[pointIdx]
			xVal := ser.XValues[pointIdx]

			px := float64(plotX) + float64(plotW)*(xVal/niceMaxX)
			py := float64(plotY+plotH) - float64(plotH)*(yVal/niceMaxY)

			out.WriteString(fmt.Sprintf(`  <circle cx="%.1f" cy="%.1f" r="4" fill="%s" />
`, px, py, ser.Color))
		}
	}

	// Legend
	if data.HasLegend && len(data.Series) > 0 {
		legendY := plotY + plotH + bottomMargin + labelFS + 2
		charW := float64(labelFS) * 0.55

		totalLegendW := 0.0
		for _, ser := range data.Series {
			totalLegendW += 16 + float64(len(ser.Name))*charW + 10
		}
		legendStartX := float64(x) + (float64(w)-totalLegendW)/2.0

		curX := legendStartX
		for _, ser := range data.Series {
			swatchY := legendY - 5
			out.WriteString(fmt.Sprintf(`  <circle cx="%.0f" cy="%.0f" r="4" fill="%s" />
`, curX+4, float64(swatchY), ser.Color))
			curX += 12

			out.WriteString(fmt.Sprintf(`  <text x="%.0f" y="%d" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif">%s</text>
`, curX, legendY, labelFS, axisLabelColor, escapeXMLText(ser.Name)))
			curX += float64(len(ser.Name))*charW + 10
		}
	}

	return out.String()
}

// renderRadarChartSVG renders a basic radar chart as an SVG fragment.
func renderRadarChartSVG(data *ChartData, x, y, w, h int) string {
	if len(data.Categories) == 0 {
		return ""
	}

	var out strings.Builder

	// Font sizes
	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 25
	}
	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 14
	}

	titleH := titleFS + 25
	plotY := y + titleH
	plotH := h - titleH - 30
	centerX := float64(x + w/2)
	centerY := float64(plotY + plotH/2)

	radius := float64(w) / 2.5
	if float64(plotH)/2.2 < radius {
		radius = float64(plotH) / 2.2
	}

	if radius < 20 {
		return ""
	}

	// Find max value
	maxVal := 0.0
	for _, s := range data.Series {
		for _, v := range s.Values {
			if v > maxVal {
				maxVal = v
			}
		}
	}
	if maxVal == 0 {
		maxVal = 10
	}
	tickInterval := niceTickInterval(maxVal)
	numTicks := int(math.Round(maxVal/tickInterval)) + 1
	niceMax := float64(numTicks) * tickInterval

	// Background and Title
	lineBgColor := "#F5F5F5"
	if data.BgColor != "" {
		lineBgColor = data.BgColor
	}
	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="%s" rx="0"/>
`, x, y, w, h, lineBgColor))

	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="#595959" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, x+w/2, y+titleFS+8, titleFS, escapeXMLText(data.Title)))

	gridColor := "#D9D9D9"
	axisLabelColor := "#595959"
	numCats := len(data.Categories)

	// Draw Grid Concentric Polygons
	for i := 1; i <= numTicks; i++ {
		r := radius * (float64(i) / float64(numTicks))
		var pts []string
		for catIdx := 0; catIdx < numCats; catIdx++ {
			angle := float64(catIdx)*(2.0*math.Pi/float64(numCats)) - math.Pi/2.0
			px := centerX + r*math.Cos(angle)
			py := centerY + r*math.Sin(angle)
			pts = append(pts, fmt.Sprintf("%.2f,%.2f", px, py))
		}
		out.WriteString(fmt.Sprintf(`  <polygon points="%s" fill="none" stroke="%s" stroke-width="1"/>
`, strings.Join(pts, " "), gridColor))
	}

	// Draw Axes & Labels
	for catIdx := 0; catIdx < numCats; catIdx++ {
		angle := float64(catIdx)*(2.0*math.Pi/float64(numCats)) - math.Pi/2.0
		axX := centerX + radius*math.Cos(angle)
		axY := centerY + radius*math.Sin(angle)

		out.WriteString(fmt.Sprintf(`  <line x1="%.2f" y1="%.2f" x2="%.2f" y2="%.2f" stroke="%s" stroke-width="1"/>
`, centerX, centerY, axX, axY, gridColor))

		// Label
		lblPad := 15.0
		lblX := centerX + (radius+lblPad)*math.Cos(angle)
		lblY := centerY + (radius+lblPad)*math.Sin(angle) + float64(labelFS)/3.0
		anchor := "middle"
		if math.Abs(math.Cos(angle)) > 0.1 {
			if math.Cos(angle) > 0 {
				anchor = "start"
			} else {
				anchor = "end"
			}
		}
		out.WriteString(fmt.Sprintf(`  <text x="%.2f" y="%.2f" font-size="%d" fill="%s" font-family="Segoe UI, sans-serif" text-anchor="%s">%s</text>
`, lblX, lblY, labelFS, axisLabelColor, anchor, escapeXMLText(data.Categories[catIdx])))
	}

	// Draw Series Polygons
	for _, ser := range data.Series {
		var pts []string
		for catIdx := 0; catIdx < numCats && catIdx < len(ser.Values); catIdx++ {
			val := ser.Values[catIdx]
			if val < 0 {
				val = 0
			}
			if val > niceMax {
				val = niceMax
			}

			r := radius * (val / niceMax)
			angle := float64(catIdx)*(2.0*math.Pi/float64(numCats)) - math.Pi/2.0
			px := centerX + r*math.Cos(angle)
			py := centerY + r*math.Sin(angle)
			pts = append(pts, fmt.Sprintf("%.2f,%.2f", px, py))
		}
		out.WriteString(fmt.Sprintf(`  <polygon points="%s" fill="%s" fill-opacity="0.3" stroke="%s" stroke-width="2"/>
`, strings.Join(pts, " "), ser.Color, ser.Color))
	}

	return out.String()
}

// parseExtendedChartXML parses a new cx:chartSpace extended chart XML file.
func parseExtendedChartXML(chartXMLBytes []byte, themeColors map[string]string) (*ChartData, error) {
	var cs ExtendedChartSpace
	if err := xml.Unmarshal(chartXMLBytes, &cs); err != nil {
		return nil, fmt.Errorf("extended chart XML parse error: %w", err)
	}

	data := &ChartData{
		Title:       "Chart Title", // default
		HasGrid:     true,
		ThemeColors: themeColors,
	}

	if cs.SpPr != nil {
		if cs.SpPr.SolidFill != nil {
			c := ""
			if cs.SpPr.SolidFill.SchemeClr != nil {
				c = cs.SpPr.SolidFill.SchemeClr.Val
			} else if cs.SpPr.SolidFill.SrgbClr != nil {
				c = cs.SpPr.SolidFill.SrgbClr.Val
			}
			if c != "" {
				if strings.HasPrefix(c, "accent") || strings.HasPrefix(c, "bg") || strings.HasPrefix(c, "dk") || strings.HasPrefix(c, "lt") {
					if hex, ok := themeColors[c]; ok {
						data.BgColor = hex
					}
				} else {
					data.BgColor = "#" + c
				}
			}
		} else if cs.SpPr.GradFill != nil {
			// Basic gradient conversion
			data.BgIsGradient = true
			var gradStops []GradientStop
			if cs.SpPr.GradFill.Path != nil && (cs.SpPr.GradFill.Path.Path == "circle" || cs.SpPr.GradFill.Path.Path == "shape") {
				data.BgGradIsRadial = true
			}
			for _, gs := range cs.SpPr.GradFill.GsLst {
				c := ""
				if gs.SchemeClr != nil {
					c = gs.SchemeClr.Val
				} else if gs.SrgbClr != nil {
					c = gs.SrgbClr.Val
				}
				if c != "" {
					if strings.HasPrefix(c, "accent") || strings.HasPrefix(c, "bg") || strings.HasPrefix(c, "dk") || strings.HasPrefix(c, "lt") {
						if hex, ok := themeColors[c]; ok {
							c = hex
						}
					} else {
						if len(c) == 6 {
							c = "#" + c
						}
					}
				}
				p := float64(gs.Pos) / 100000.0
				gradStops = append(gradStops, GradientStop{
					Color: c,
					Pos:   p,
				})
			}
			data.BgGradStops = gradStops
		}
	}

	if cs.Chart.Title != nil && cs.Chart.Title.Tx != nil {
		if cs.Chart.Title.Tx.TxData != nil {
			data.Title = cs.Chart.Title.Tx.TxData.V
		} else if cs.Chart.Title.Tx.Rich != nil {
			var parts []string
			for _, p := range cs.Chart.Title.Tx.Rich.Paragraphs {
				for _, r := range p.Runs {
					if r.T != "" {
						parts = append(parts, r.T)
					}
				}
			}
			if len(parts) > 0 {
				data.Title = strings.Join(parts, " ")
			}
		}
	}

	// Extract CatScaling gapWidth (used in Histograms/Pareto)
	for _, ax := range cs.Chart.PlotArea.Axes {
		if ax.CatScaling != nil && ax.CatScaling.GapWidth != "" {
			if v, err := strconv.ParseFloat(ax.CatScaling.GapWidth, 64); err == nil {
				data.GapWidth = v // e.g. 0.33, used for bar layout calculation if needed
				data.HasGapWidth = true
			}
		}
	}

	// Extract data points
	var globalCats []string
	var globalVals []float64

	if len(cs.ChartData.Data) > 0 {
		cxData := cs.ChartData.Data[0]
		// Determine the max ptCount across all levels to size our arrays
		maxPtCount := 0
		if cxData.StrDim != nil {
			for _, lvl := range cxData.StrDim.Lvls {
				for _, pt := range lvl.Pts {
					if pt.Idx >= maxPtCount {
						maxPtCount = pt.Idx + 1
					}
				}
			}
		}
		if cxData.NumDim != nil {
			for _, lvl := range cxData.NumDim.Lvls {
				for _, pt := range lvl.Pts {
					if pt.Idx >= maxPtCount {
						maxPtCount = pt.Idx + 1
					}
				}
			}
		}

		if maxPtCount > 0 {
			// Parse Categories with explicit sparse idx mapping
			if cxData.StrDim != nil && len(cxData.StrDim.Lvls) > 0 {
				numLvls := len(cxData.StrDim.Lvls)
				for i := 0; i < maxPtCount; i++ {
					var path []string
					for lvl := numLvls - 1; lvl >= 0; lvl-- {
						// Find if this level has a point mapped at 'i'
						val := ""
						for _, pt := range cxData.StrDim.Lvls[lvl].Pts {
							if pt.Idx == i {
								val = pt.V
								break
							}
						}
						if val != "" {
							path = append(path, val)
						}
					}
					globalCats = append(globalCats, strings.Join(path, "|"))
				}
			}

			// Parse Values with explicit sparse idx mapping
			if cxData.NumDim != nil && len(cxData.NumDim.Lvls) > 0 {
				globalVals = make([]float64, maxPtCount)
				for _, pt := range cxData.NumDim.Lvls[0].Pts {
					v, _ := strconv.ParseFloat(pt.V, 64)
					if pt.Idx < maxPtCount {
						globalVals[pt.Idx] = v
					}
				}
			}
		}
	}

	data.Categories = globalCats

	for _, curSer := range cs.Chart.PlotArea.PlotAreaRegion.Series {
		data.ChartType = curSer.LayoutId // Default case: use globalCats and globalVals directly
		sd := SeriesData{Name: "Series", Type: curSer.LayoutId}
		if curSer.Tx != nil && curSer.Tx.TxData != nil {
			sd.Name = curSer.Tx.TxData.V
		}

		sd.LineColor = "#FFFFFF" // Automatic border fallback
		sd.HasLine = true        // By default, Histograms have white outlines unless otherwise disabled

		if curSer.SpPr != nil {
			if curSer.SpPr.Ln != nil {
				if curSer.SpPr.Ln.SolidFill != nil {
					c := ""
					if curSer.SpPr.Ln.SolidFill.SrgbClr != nil {
						c = curSer.SpPr.Ln.SolidFill.SrgbClr.Val
					} else if curSer.SpPr.Ln.SolidFill.SchemeClr != nil {
						c = curSer.SpPr.Ln.SolidFill.SchemeClr.Val
					}
					if c != "" {
						if strings.HasPrefix(c, "accent") || strings.HasPrefix(c, "bg") || strings.HasPrefix(c, "dk") || strings.HasPrefix(c, "lt") {
							if hex, ok := themeColors[c]; ok {
								sd.LineColor = hex
							}
						} else {
							sd.LineColor = "#" + c
						}
					}
				}
			}
			if curSer.SpPr.EffectLst != nil && curSer.SpPr.EffectLst.OuterShdw != nil {
				sd.HasShadow = true
			}
		}

		catsToAttach := make([]string, len(globalCats))
		copy(catsToAttach, globalCats)
		valsToAttach := make([]float64, len(globalVals))
		copy(valsToAttach, globalVals)

		// Handle Histogram Binning
		if curSer.LayoutPr != nil && curSer.LayoutPr.Binning != nil {
			// Basic binning algorithm (like PPTX default):
			// count bins = round(sqrt(N)) or handle manually
			n := len(globalVals)
			if n > 0 {
				bins := int(math.Round(math.Sqrt(float64(n))))
				if bins < 2 {
					bins = 2
				}
				minVal, maxVal := globalVals[0], globalVals[0]
				for _, v := range globalVals {
					if v < minVal {
						minVal = v
					}
					if v > maxVal {
						maxVal = v
					}
				}
				binSize := (maxVal - minVal) / float64(bins)
				if curSer.LayoutPr.Binning.BinSize > 0 {
					binSize = curSer.LayoutPr.Binning.BinSize
				}

				counts := make([]float64, bins)
				binLabels := make([]string, bins)

				for b := 0; b < bins; b++ {
					start := minVal + float64(b)*binSize
					end := start + binSize
					if curSer.LayoutPr.Binning.IntervalClosed == "l" {
						binLabels[b] = fmt.Sprintf("[%g, %g)", start, end)
					} else {
						// default "r" closed
						binLabels[b] = fmt.Sprintf("(%g, %g]", start, end)
						if b == 0 {
							// first bin usually closed on left too
							binLabels[b] = fmt.Sprintf("[%g, %g]", start, end)
						}
					}
				}

				for _, v := range globalVals {
					// find bin
					idx := -1
					for b := 0; b < bins; b++ {
						start := minVal + float64(b)*binSize
						end := start + binSize
						if curSer.LayoutPr.Binning.IntervalClosed == "l" {
							if v >= start && (v < end || (b == bins-1 && math.Abs(v-end) < 1e-6)) {
								idx = b
								break
							}
						} else {
							if (v > start || (b == 0 && math.Abs(v-start) < 1e-6)) && v <= end {
								idx = b
								break
							}
						}
					}
					if idx >= 0 && idx < bins {
						counts[idx]++
					} else {
						// fallback safe bound
						counts[bins-1]++
					}
				}

				catsToAttach = binLabels
				valsToAttach = counts
			}
		}

		// Handle Pareto Aggregation / Line
		if curSer.LayoutPr != nil && curSer.LayoutId == "clusteredColumn" {
			// check if there's a paretoLine in this chart space
			hasParetoLine := false
			for _, checkSer := range cs.Chart.PlotArea.PlotAreaRegion.Series {
				if checkSer.LayoutId == "paretoLine" {
					hasParetoLine = true
				}
			}
			if hasParetoLine {
				// Aggregate by top-level category
				aggMap := make(map[string]float64)
				var aggOrder []string
				for i, cat := range catsToAttach {
					parts := strings.Split(cat, "|")
					top := parts[0]
					if _, exists := aggMap[top]; !exists {
						aggOrder = append(aggOrder, top)
					}
					if i < len(valsToAttach) {
						aggMap[top] += valsToAttach[i]
					}
				}

				var newCats []string
				var newVals []float64
				for _, top := range aggOrder {
					newCats = append(newCats, top)
					newVals = append(newVals, aggMap[top])
				}
				catsToAttach = newCats
				valsToAttach = newVals

				// Sort categories by descending value for Pareto
				// But we just do a simple paired sort
				type pair struct {
					cat string
					val float64
				}
				pairs := make([]pair, len(valsToAttach))
				for i := range valsToAttach {
					cat := ""
					if i < len(catsToAttach) {
						cat = catsToAttach[i]
					}
					pairs[i] = pair{cat, valsToAttach[i]}
				}
				// Sort descending
				for i := 0; i < len(pairs)-1; i++ {
					for j := i + 1; j < len(pairs); j++ {
						if pairs[j].val > pairs[i].val {
							pairs[i], pairs[j] = pairs[j], pairs[i]
						}
					}
				}
				for i, p := range pairs {
					if i < len(catsToAttach) {
						catsToAttach[i] = p.cat
					}
					valsToAttach[i] = p.val
				}

				// Update global elements so paretoLine series inherits correct dimensions
				globalCats = catsToAttach
				globalVals = valsToAttach
			}
		}

		if curSer.LayoutId == "paretoLine" {
			// recalculate line based on aggregated and sorted categories from the main clusteredColumn
			// The values are cumulative percentages
			var total float64
			for _, v := range globalVals {
				total += v
			}

			cumulative := make([]float64, len(globalVals))
			var sum float64
			for i, v := range globalVals {
				sum += v
				cumulative[i] = (sum / total) * 100.0
			}
			valsToAttach = cumulative
			catsToAttach = globalCats
		}

		// Update globalCats if this is the primary generating series
		if len(catsToAttach) > 0 && curSer.LayoutId != "paretoLine" {
			data.Categories = catsToAttach
		}

		for _, v := range valsToAttach {
			sd.Values = append(sd.Values, v)
		}

		curIdx := len(data.Series)

		// Set a generic color if none available
		accentColors := getChartAccentColors(themeColors)
		if len(accentColors) > 0 {
			sd.Color = accentColors[curIdx%len(accentColors)]
		} else {
			colors := []string{"#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#4472C4", "#70AD47"}
			sd.Color = colors[curIdx%len(colors)]
		}

		data.Series = append(data.Series, sd)
	}

	if len(cs.Chart.PlotArea.PlotAreaRegion.Series) > 1 {
		data.ChartType = "combo"
	}

	return data, nil
}

// renderTreemapChartSVG renders a basic 1-level squarified treemap chart.
func renderTreemapChartSVG(data *ChartData, x, y, w, h int) string {
	if len(data.Categories) == 0 || len(data.Series) == 0 {
		return ""
	}

	var out strings.Builder

	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 20
	}
	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 12
	}

	titleH := titleFS + 20
	plotX := float64(x + 10)
	plotY := float64(y + titleH)
	plotW := float64(w - 20)
	plotH := float64(h - titleH - 10)

	if plotW < 50 || plotH < 50 {
		return ""
	}

	// Background
	bgColor := "#F5F5F5"
	if data.BgColor != "" {
		bgColor = data.BgColor
	}
	bgFillAttr := fmt.Sprintf(`fill="%s"`, bgColor)

	fillDefs, _, _, _ := generateSVGFillAndEffectDefs(data, "treemap_bg")
	if fillDefs != "" {
		out.WriteString(fillDefs)
	}
	if data.BgIsGradient {
		bgFillAttr = `fill="url(#chartBgGradtreemap_bg)"`
	}
	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" %s rx="0"/>
`, x, y, w, h, bgFillAttr))

	// Title
	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="#595959" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, x+w/2, y+titleFS+8, titleFS, escapeXMLText(data.Title)))

	ser := data.Series[0]
	total := 0.0
	for _, v := range ser.Values {
		if v > 0 {
			total += v
		}
	}

	if total <= 0 {
		return out.String()
	}

	colors := []string{"#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#4472C4", "#70AD47"}
	if len(data.ThemeColors) > 0 {
		accent := getChartAccentColors(data.ThemeColors)
		if len(accent) > 0 {
			colors = accent
		}
	}
	if len(data.CustomColors) > 0 {
		colors = data.CustomColors
	}

	// Map colors by top-level branch rank (largest sum gets colors[0], 2nd gets colors[1], etc)
	maxDepth := 1
	var paths [][]string
	topLevelSums := make(map[string]float64)

	for i, c := range data.Categories {
		parts := strings.Split(c, "|")
		if len(parts) > maxDepth {
			maxDepth = len(parts)
		}
		paths = append(paths, parts)

		if len(parts) > 0 && i < len(ser.Values) && ser.Values[i] > 0 {
			topLevelSums[parts[0]] += ser.Values[i]
		}
	}

	type topRank struct {
		name string
		val  float64
	}
	var ranks []topRank
	for name, val := range topLevelSums {
		ranks = append(ranks, topRank{name, val})
	}
	// Sort descending
	for i := 0; i < len(ranks)-1; i++ {
		for j := i + 1; j < len(ranks); j++ {
			if ranks[j].val > ranks[i].val {
				ranks[i], ranks[j] = ranks[j], ranks[i]
			}
		}
	}

	topColors := make(map[string]string)
	for i, r := range ranks {
		topColors[r.name] = colors[i%len(colors)]
	}

	// Pre-calculate colors per data point
	ptColor := make([]string, len(ser.Values))
	for i := 0; i < len(ser.Values); i++ {
		if i >= len(paths) || len(paths[i]) == 0 {
			ptColor[i] = colors[0]
			continue
		}
		if c, ok := topColors[paths[i][0]]; ok {
			ptColor[i] = c
		} else {
			ptColor[i] = colors[0]
		}
	}

	type tmNode struct {
		val      float64
		cat      string
		color    string
		isLeaf   bool
		children []*tmNode
	}

	var root tmNode
	root.cat = "root"

	for i, v := range ser.Values {
		if v <= 0 || i >= len(data.Categories) || i >= len(paths) {
			continue
		}

		curr := &root
		for level := 0; level < len(paths[i]); level++ {
			nodeName := paths[i][level]
			var child *tmNode
			for _, c := range curr.children {
				if c.cat == nodeName {
					child = c
					break
				}
			}
			if child == nil {
				child = &tmNode{
					cat:    nodeName,
					isLeaf: level == len(paths[i])-1,
					color:  ptColor[i],
				}
				curr.children = append(curr.children, child)
			}
			child.val += v
			curr = child
		}
	}

	var layout func(x, y, w, h float64, node *tmNode)
	layout = func(x, y, w, h float64, node *tmNode) {
		if w < 1 || h < 1 {
			return
		}
		if node.isLeaf {
			out.WriteString(fmt.Sprintf(`  <rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s" stroke="#FFFFFF" stroke-width="1.5"/>
`, x, y, w, h, node.color))
			if w > 30 && h > 20 {
				out.WriteString(fmt.Sprintf(`  <text x="%.1f" y="%.1f" font-size="%d" fill="#FFFFFF" font-family="Segoe UI, sans-serif" text-anchor="start">%s</text>
`, x+5, y+float64(labelFS)+5, labelFS, escapeXMLText(node.cat)))
			}
			return
		}

		// Group Node (Branch/Stem)
		bannerH := 0.0
		if node.cat != "root" {
			bannerH = float64(labelFS) + 8
			if bannerH > h {
				bannerH = h
			}
			// Bounding box for group (darkened or base color)
			out.WriteString(fmt.Sprintf(`  <rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s" stroke="#FFFFFF" stroke-width="1.5"/>
`, x, y, w, h, node.color))
			if w > 30 && bannerH >= float64(labelFS) {
				out.WriteString(fmt.Sprintf(`  <text x="%.1f" y="%.1f" font-size="%d" fill="#FFFFFF" font-family="Segoe UI, sans-serif" text-anchor="start">%s</text>
`, x+5, y+float64(labelFS)+2, labelFS, escapeXMLText(node.cat)))
			}
			y += bannerH
			h -= bannerH
		}

		if h <= 0 || len(node.children) == 0 {
			return
		}

		var sliceLayout func(cx, cy, cw, ch float64, kids []*tmNode)
		sliceLayout = func(cx, cy, cw, ch float64, kids []*tmNode) {
			if len(kids) == 0 {
				return
			}
			if len(kids) == 1 {
				layout(cx, cy, cw, ch, kids[0])
				return
			}
			sum := 0.0
			for _, k := range kids {
				sum += k.val
			}
			if sum <= 0 {
				return
			}

			half := sum / 2
			running := 0.0
			splitIdx := 1
			for i := 0; i < len(kids)-1; i++ {
				running += kids[i].val
				if running >= half {
					splitIdx = i + 1
					if i > 0 && (running-half) > (half-(running-kids[i].val)) {
						splitIdx = i
					}
					break
				}
			}

			sum1 := 0.0
			for i := 0; i < splitIdx; i++ {
				sum1 += kids[i].val
			}
			ratio := sum1 / sum

			if cw > ch {
				w1 := cw * ratio
				sliceLayout(cx, cy, w1, ch, kids[:splitIdx])
				sliceLayout(cx+w1, cy, cw-w1, ch, kids[splitIdx:])
			} else {
				h1 := ch * ratio
				sliceLayout(cx, cy, cw, h1, kids[:splitIdx])
				sliceLayout(cx, cy+h1, cw, ch-h1, kids[splitIdx:])
			}
		}

		sliceLayout(x, y, w, h, node.children)
	}

	layout(float64(plotX), plotY, plotW, plotH, &root)

	return out.String()
}

// renderSunburstChartSVG renders a basic 1-level sunburst chart (similar to doughnut).
func renderSunburstChartSVG(data *ChartData, x, y, w, h int) string {
	if len(data.Categories) == 0 || len(data.Series) == 0 {
		return ""
	}

	var out strings.Builder

	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 25
	}

	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 16
	}

	titleH := titleFS + 20
	plotY := float64(y + titleH)
	plotH := float64(h - titleH - 10)

	// Background
	bgColor := "#F5F5F5"
	if data.BgColor != "" {
		bgColor = data.BgColor
	}
	bgFillAttr := fmt.Sprintf(`fill="%s"`, bgColor)

	fillDefs, _, _, _ := generateSVGFillAndEffectDefs(data, "sunburst_bg")
	if fillDefs != "" {
		out.WriteString(fillDefs)
	}
	if data.BgIsGradient {
		bgFillAttr = `fill="url(#chartBgGradsunburst_bg)"`
	}
	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" %s rx="0"/>
`, x, y, w, h, bgFillAttr))

	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="#595959" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, x+w/2, y+titleFS+8, titleFS, escapeXMLText(data.Title)))

	ser := data.Series[0]
	total := 0.0
	for _, v := range ser.Values {
		if v > 0 {
			total += v
		}
	}

	if total <= 0 {
		return out.String()
	}

	colors := []string{"#5B9BD5", "#ED7D31", "#A5A5A5", "#FFC000", "#4472C4", "#70AD47"}
	if len(data.ThemeColors) > 0 {
		accent := getChartAccentColors(data.ThemeColors)
		if len(accent) > 0 {
			colors = accent
		}
	}
	if len(data.CustomColors) > 0 {
		colors = data.CustomColors
	}

	centerX := float64(x + w/2)
	centerY := plotY + plotH/2
	outerRadius := math.Min(float64(w), plotH) / 2.2

	// Map colors by top-level branch rank (largest sum gets colors[0], 2nd gets colors[1], etc)
	maxDepth := 1
	var paths [][]string
	topLevelSums := make(map[string]float64)

	for i, c := range data.Categories {
		parts := strings.Split(c, "|")
		if len(parts) > maxDepth {
			maxDepth = len(parts)
		}
		paths = append(paths, parts)

		if len(parts) > 0 && i < len(ser.Values) && ser.Values[i] > 0 {
			topLevelSums[parts[0]] += ser.Values[i]
		}
	}

	type topRank struct {
		name string
		val  float64
	}
	var ranks []topRank
	for name, val := range topLevelSums {
		ranks = append(ranks, topRank{name, val})
	}
	// Sort descending
	for i := 0; i < len(ranks)-1; i++ {
		for j := i + 1; j < len(ranks); j++ {
			if ranks[j].val > ranks[i].val {
				ranks[i], ranks[j] = ranks[j], ranks[i]
			}
		}
	}

	topColors := make(map[string]string)
	for i, r := range ranks {
		topColors[r.name] = colors[i%len(colors)]
	}

	// Pre-calculate colors per data point
	ptColor := make([]string, len(ser.Values))
	for i := 0; i < len(ser.Values); i++ {
		if i >= len(paths) || len(paths[i]) == 0 {
			ptColor[i] = colors[0]
			continue
		}
		if c, ok := topColors[paths[i][0]]; ok {
			ptColor[i] = c
		} else {
			ptColor[i] = colors[0]
		}
	}

	ringWidth := outerRadius / float64(maxDepth+1)
	startAngleBase := -math.Pi / 2

	for level := 0; level < maxDepth; level++ {
		startAngle := startAngleBase
		innerR := ringWidth * float64(level+1)
		outerR := ringWidth * float64(level+2)

		// Group items at this level
		// If multiple consecutive points share the same path up to this level, merge them
		i := 0
		for i < len(ser.Values) {
			if i >= len(paths) {
				break
			}
			if len(paths[i]) <= level {
				// leaf node stopped early
				v := ser.Values[i]
				ratio := v / total
				sweepAngle := ratio * 2 * math.Pi
				startAngle += sweepAngle
				i++
				// Optionally extend leaf node to outer radius? In standard sunburst, leaves stop or extend.
				// We'll let it be blank for this level if it ended early.
				continue
			}

			groupCat := paths[i][level]
			groupVal := ser.Values[i]
			groupStart := startAngle
			groupColor := ptColor[i]

			j := i + 1
			for j < len(ser.Values) && j < len(paths) && len(paths[j]) > level && paths[j][level] == groupCat {
				// For higher levels (stem/branch), they must match all previous levels too
				match := true
				for k := 0; k <= level; k++ {
					if paths[j][k] != paths[i][k] {
						match = false
						break
					}
				}
				if !match {
					break
				}
				groupVal += ser.Values[j]
				j++
			}

			if groupVal > 0 && groupCat != "" {
				ratio := groupVal / total
				sweepAngle := ratio * 2 * math.Pi
				endAngle := groupStart + sweepAngle
				// color := colors[colorIdx%len(colors)]

				x1Out := centerX + outerR*math.Cos(groupStart)
				y1Out := centerY + outerR*math.Sin(groupStart)
				x2Out := centerX + outerR*math.Cos(endAngle)
				y2Out := centerY + outerR*math.Sin(endAngle)

				x1In := centerX + innerR*math.Cos(groupStart)
				y1In := centerY + innerR*math.Sin(groupStart)
				x2In := centerX + innerR*math.Cos(endAngle)
				y2In := centerY + innerR*math.Sin(endAngle)

				largeArcFlag := 0
				if sweepAngle > math.Pi {
					largeArcFlag = 1
				}

				path := fmt.Sprintf("M %.2f %.2f A %.2f %.2f 0 %d 1 %.2f %.2f L %.2f %.2f A %.2f %.2f 0 %d 0 %.2f %.2f Z",
					x1Out, y1Out, outerR, outerR, largeArcFlag, x2Out, y2Out,
					x2In, y2In, innerR, innerR, largeArcFlag, x1In, y1In)

				out.WriteString(fmt.Sprintf(`  <path d="%s" fill="%s" stroke="#FFFFFF" stroke-width="1.5"/>
`, path, groupColor))

				clipId := fmt.Sprintf("clip_sunburst_%d_%d", level, i)
				out.WriteString(fmt.Sprintf(`  <clipPath id="%s"><path d="%s"/></clipPath>
`, clipId, path))

				// Text
				midAngle := groupStart + sweepAngle/2
				textRadius := innerR + (outerR-innerR)/2
				txtX := centerX + textRadius*math.Cos(midAngle)
				txtY := centerY + textRadius*math.Sin(midAngle)

				// only render text if arc is wide enough
				if sweepAngle > 0.05 {
					deg := midAngle * 180 / math.Pi
					if math.Cos(midAngle) < 0 {
						deg += 180
					}
					out.WriteString(fmt.Sprintf(`  <text x="%.2f" y="%.2f" font-size="%d" fill="#FFFFFF" font-family="Segoe UI, sans-serif" text-anchor="middle" transform="rotate(%.2f, %.2f, %.2f)">%s</text>
`, txtX, txtY+float64(labelFS)/3, labelFS, deg, txtX, txtY, escapeXMLText(groupCat)))
				}
			}

			// Advance by the slice we processed
			for k := i; k < j; k++ {
				startAngle += (ser.Values[k] / total) * 2 * math.Pi
			}

			i = j
		}
	}

	return out.String()
}

// renderWaterfallChartSVG renders a basic waterfall chart.
func renderWaterfallChartSVG(data *ChartData, x, y, w, h int) string {
	if len(data.Categories) == 0 || len(data.Series) == 0 {
		return ""
	}

	var out strings.Builder

	titleFS := data.TitleFontSize
	if titleFS == 0 {
		titleFS = 25
	}
	labelFS := data.LabelFontSize
	if labelFS == 0 {
		labelFS = 16
	}

	titleH := titleFS + 20
	plotX := float64(x + 50)
	plotY := float64(y + titleH)
	plotW := float64(w - 60)
	plotH := float64(h - titleH - 30)

	if plotW < 50 || plotH < 50 {
		return ""
	}

	lineBgColor := "#F5F5F5"
	if data.BgColor != "" {
		lineBgColor = data.BgColor
	}
	out.WriteString(fmt.Sprintf(`  <rect x="%d" y="%d" width="%d" height="%d" fill="%s" rx="0"/>
`, x, y, w, h, lineBgColor))

	// Title
	out.WriteString(fmt.Sprintf(`  <text x="%d" y="%d" font-size="%d" fill="#595959" font-family="Segoe UI, sans-serif" text-anchor="middle">%s</text>
`, x+w/2, y+titleFS+8, titleFS, escapeXMLText(data.Title)))

	ser := data.Series[0]
	// Calculate max cumulative value
	running := 0.0
	maxVal := 0.0
	minVal := 0.0
	for _, v := range ser.Values {
		running += v
		if running > maxVal {
			maxVal = running
		}
		if running < minVal {
			minVal = running
		}
	}
	if maxVal == 0 {
		maxVal = 10
	}
	rangeVal := maxVal - minVal
	if rangeVal == 0 {
		rangeVal = 10
	}

	niceMax := maxVal + rangeVal*0.1
	niceMin := minVal
	if niceMin > 0 {
		niceMin = 0
	}
	if niceMin < 0 {
		niceMin -= rangeVal * 0.1
	}
	totalRange := niceMax - niceMin

	barW := plotW / float64(len(data.Categories)) * 0.8
	gapW := plotW / float64(len(data.Categories)) * 0.2

	runningY := 0.0
	curX := plotX + gapW/2

	// Map generic colors
	colors := []string{"#5B9BD5", "#ED7D31", "#A5A5A5"} // 0=increase, 1=decrease, 2=total
	if len(data.ThemeColors) > 0 {
		accent := getChartAccentColors(data.ThemeColors)
		if len(accent) >= 3 {
			colors = []string{accent[0], accent[1], accent[2]}
		}
	}

	for i, v := range ser.Values {
		if i >= len(data.Categories) {
			break
		}

		var rectH, rectY float64
		color := colors[0] // increase

		if v >= 0 {
			rectH = (v / totalRange) * plotH
			// Bottom is old running, top is new running
			newRunning := runningY + v
			rectY = plotY + plotH - ((newRunning - niceMin) / totalRange * plotH)
			runningY = newRunning
		} else {
			color = colors[1] // decrease
			rectH = (-v / totalRange) * plotH
			// Top is old running, bottom is new running
			rectY = plotY + plotH - ((runningY - niceMin) / totalRange * plotH)
			runningY += v
		}

		out.WriteString(fmt.Sprintf(`  <rect x="%.1f" y="%.1f" width="%.1f" height="%.1f" fill="%s"/>
`, curX, rectY, barW, rectH, color))

		catName := data.Categories[i]
		parts := strings.Split(catName, "|")

		// Data label inside the bar
		if rectH > 15 {
			labelY := rectY + 16
			for pIdx := 0; pIdx < len(parts); pIdx++ {
				out.WriteString(fmt.Sprintf(`  <text x="%.1f" y="%.1f" font-size="16" fill="#FFFFFF" font-family="Calibri, sans-serif" text-anchor="middle">%s</text>
`, curX+barW/2, labelY, escapeXMLText(parts[pIdx])))
				labelY += 16
			}
		}

		// Rotate X-axis label by -90 degrees
		lblX := curX + barW/2
		lblYStart := plotY + plotH + float64(labelFS) + 15
		// Draw parts from Leaf (last) outwards to Branch (first) placing them further down
		for pIdx := len(parts) - 1; pIdx >= 0; pIdx-- {
			out.WriteString(fmt.Sprintf(`  <text font-size="%d" fill="#595959" font-family="Segoe UI, sans-serif" transform="matrix(0 -1 1 0 %.1f %.1f)">%s</text>
`, labelFS, lblX, lblYStart, escapeXMLText(parts[pIdx])))
			lblYStart += float64(labelFS) + 4
		}

		curX += barW + gapW
	}

	return out.String()
}
