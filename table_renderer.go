package pptxtosvg

import (
	"fmt"
	"math"
	"strconv"
	"strings"
)

// ====================================================================
// PPTX Table Renderer
//
// Parses table data from OOXML <a:tbl> elements inside <p:graphicFrame>
// and renders them as SVG fragments.
//
// Style resolution is data-driven: reads ppt/tableStyles.xml definitions
// rather than hardcoding GUID-to-style mappings.
// ====================================================================

const tableURI = "http://schemas.openxmlformats.org/drawingml/2006/table"

// processTableGraphicFrame processes a <p:graphicFrame> containing a table.
func processTableGraphicFrame(gf OOXMLGraphicFrame, ctx *slideContext) *RenderShape {
	if gf.Graphic.GraphicData.Tbl == nil {
		return nil
	}
	if gf.Graphic.GraphicData.URI != tableURI {
		return nil
	}

	tbl := gf.Graphic.GraphicData.Tbl

	frameX := gf.Xfrm.Off.X
	frameY := gf.Xfrm.Off.Y
	frameCx := gf.Xfrm.Ext.Cx
	frameCy := gf.Xfrm.Ext.Cy

	svg := renderTableSVG(tbl, frameX, frameY, frameCx, frameCy, ctx)

	return &RenderShape{
		XmlID:          int(gf.NvGraphicFramePr.CNvPr.ID),
		Name:           gf.NvGraphicFramePr.CNvPr.Name,
		ShapeType:      "table",
		X:              frameX,
		Y:              frameY,
		W:              frameCx,
		H:              frameCy,
		PreRenderedSVG: svg,
	}
}

// resolveStyleSchemeColor resolves a scheme color reference with tint/shade modifiers.
func resolveStyleSchemeColor(sc *OOXMLSchemeClr, themeColors map[string]string) string {
	if sc == nil {
		return ""
	}
	c := resolveSchemeColor(sc.Val, themeColors)
	if sc.Tint != nil && sc.Tint.Val > 0 {
		c = applyTint(c, sc.Tint.Val)
	}
	if sc.Shade != nil && sc.Shade.Val > 0 {
		c = applyShade(c, sc.Shade.Val)
	}
	if sc.LumMod != nil && sc.LumOff != nil {
		c = applyLumModOff(c, sc.LumMod.Val, sc.LumOff.Val)
	} else if sc.LumMod != nil && sc.LumMod.Val > 0 {
		c = applyLumMod(c, sc.LumMod.Val)
	}
	return c
}

// resolveStylePartFill resolves the fill color and opacity from a table style part's fill definition.
func resolveStylePartFill(part *OOXMLTableStylePart, themeColors map[string]string) (string, float64) {
	if part == nil || part.TcStyle == nil || part.TcStyle.Fill == nil {
		return "", 1.0
	}
	fill := part.TcStyle.Fill
	if fill.NoFill != nil {
		return "none", 1.0
	}
	if fill.SolidFill != nil {
		c := resolveSolidFill(fill.SolidFill, themeColors)
		opacity := 1.0
		if fill.SolidFill.SchemeClr != nil && fill.SolidFill.SchemeClr.Alpha != nil && fill.SolidFill.SchemeClr.Alpha.Val > 0 {
			opacity = float64(fill.SolidFill.SchemeClr.Alpha.Val) / 100000.0
		}
		return c, opacity
	}
	return "", 1.0
}

// resolveStyleBorder resolves a single border from a table style border definition.
// Returns color, width in px, and whether the border should be drawn.
func resolveStyleBorder(border *OOXMLTableStyleBorder, themeColors map[string]string) (string, float64, bool) {
	if border == nil {
		return "", 0, false
	}
	if border.Ln != nil {
		if border.Ln.NoFill != nil {
			return "", 0, false
		}
		w := float64(border.Ln.W) / float64(emuPerPx)
		if w <= 0 {
			w = 1.33
		}
		if border.Ln.SolidFill != nil {
			c := resolveSolidFill(border.Ln.SolidFill, themeColors)
			if c != "" {
				return c, w, true
			}
		}
		return "", 0, false
	}
	if border.LnRef != nil {
		// lnRef uses theme line definitions; approximate with scheme color and a default width
		if border.LnRef.SchemeClr != nil {
			c := resolveStyleSchemeColor(border.LnRef.SchemeClr, themeColors)
			w := 1.33
			return c, w, c != ""
		}
	}
	return "", 0, false
}

// resolveStyleTextProps extracts text properties from a table style part's tcTxStyle.
func resolveStyleTextProps(part *OOXMLTableStylePart, themeColors map[string]string) (bold bool, italic bool, color string) {
	if part == nil || part.TcTxStyle == nil {
		return false, false, ""
	}
	ts := part.TcTxStyle
	bold = ts.B == "on"
	italic = ts.I == "on"
	if ts.SchemeClr != nil {
		color = resolveStyleSchemeColor(ts.SchemeClr, themeColors)
	}
	return
}

// renderTableSVG renders an OOXML table as an SVG fragment.
func renderTableSVG(tbl *OOXMLTable, frameX, frameY, frameCx, frameCy int64, ctx *slideContext) string {
	if tbl == nil || len(tbl.Rows) == 0 || len(tbl.TblGrid.GridCols) == 0 {
		return ""
	}

	var out strings.Builder

	var themeColors map[string]string
	var tblStyleDef *OOXMLTableStyle
	if ctx != nil {
		themeColors = ctx.themeColors
		styleID := strings.TrimSpace(tbl.TblPr.TableStyleId)
		if ctx.tableStyles != nil && styleID != "" {
			tblStyleDef = ctx.tableStyles[styleID]
		}
	}

	defaultFontSizePx := 24
	if ctx != nil && ctx.cfg != nil && ctx.cfg.DefaultFontSizePx > 0 {
		defaultFontSizePx = ctx.cfg.DefaultFontSizePx
	}

	numCols := len(tbl.TblGrid.GridCols)

	// Column X positions in EMU (relative to frame)
	colXEmu := make([]int64, numCols+1)
	colXEmu[0] = 0
	for i, gc := range tbl.TblGrid.GridCols {
		colXEmu[i+1] = colXEmu[i] + gc.W
	}

	// Row heights, handling h=0 (auto-fit) rows
	numRows := len(tbl.Rows)
	rowHeights := make([]int64, numRows)
	var nonZeroSum int64
	nonZeroCount := 0
	for i, row := range tbl.Rows {
		rowHeights[i] = row.H
		if row.H > 0 {
			nonZeroSum += row.H
			nonZeroCount++
		}
	}
	if nonZeroCount > 0 && nonZeroCount < numRows {
		avgH := nonZeroSum / int64(nonZeroCount)
		for i := range rowHeights {
			if rowHeights[i] == 0 {
				rowHeights[i] = avgH
			}
		}
	} else if nonZeroCount == 0 && numRows > 0 {
		eachH := frameCy / int64(numRows)
		for i := range rowHeights {
			rowHeights[i] = eachH
		}
	}

	rowYEmu := make([]int64, numRows+1)
	rowYEmu[0] = 0
	for i := range rowHeights {
		rowYEmu[i+1] = rowYEmu[i] + rowHeights[i]
	}

	// Scale factors
	gridTotalW := colXEmu[numCols]
	gridTotalH := rowYEmu[numRows]
	scaleX := 1.0
	scaleY := 1.0
	if gridTotalW > 0 && frameCx > 0 {
		scaleX = float64(frameCx) / float64(gridTotalW)
	}
	if gridTotalH > 0 && frameCy > 0 {
		scaleY = float64(frameCy) / float64(gridTotalH)
	}

	toPxX := func(emu int64) float64 {
		return float64(emu)*scaleX/float64(emuPerPx) + float64(frameX)/float64(emuPerPx)
	}
	toPxY := func(emu int64) float64 {
		return float64(emu)*scaleY/float64(emuPerPx) + float64(frameY)/float64(emuPerPx)
	}
	isFirstRow := tbl.TblPr.FirstRow == "1"
	isBandRow := tbl.TblPr.BandRow == "1"

	// === Render outer shadow (from tblPr.effectLst) ===
	if tbl.TblPr.EffectLst != nil && tbl.TblPr.EffectLst.OuterShdw != nil {
		shdw := tbl.TblPr.EffectLst.OuterShdw
		if shdw.Sy != 0 {
			renderTablePerspectiveShadow(&out, tbl, shdw,
				frameX, frameY, frameCx, frameCy,
				scaleX, scaleY, colXEmu, rowYEmu,
				numCols, numRows, isFirstRow, isBandRow,
				tblStyleDef, themeColors, ctx)
		} else {
			renderTableBgShadow(&out, tbl, shdw,
				frameX, frameY, frameCx, frameCy,
				scaleX, scaleY, colXEmu, rowYEmu,
				numCols, numRows, isFirstRow, isBandRow,
				tblStyleDef, themeColors, ctx)
		}
	}

	// === Render tblBg effectRef shadow (from table style, only if tblPr has no explicit effectLst) ===
	hasTblBg := false
	tblBgFill := ""
	hasTblPrEffect := tbl.TblPr.EffectLst != nil && (tbl.TblPr.EffectLst.OuterShdw != nil || tbl.TblPr.EffectLst.InnerShdw != nil)
	if tblStyleDef != nil && tblStyleDef.TblBg != nil {
		if !hasTblPrEffect && tblStyleDef.TblBg.EffectRef != nil && ctx.themeFmtScheme != nil {
			shdw := resolveThemeEffectRef(tblStyleDef.TblBg.EffectRef, ctx.themeFmtScheme)
			if shdw != nil {
				renderTableBgShadow(&out, tbl, shdw,
					frameX, frameY, frameCx, frameCy,
					scaleX, scaleY, colXEmu, rowYEmu,
					numCols, numRows, isFirstRow, isBandRow,
					tblStyleDef, themeColors, ctx)
			}
		}
		if tblStyleDef.TblBg.FillRef != nil && ctx.themeFmtScheme != nil {
			bgFill, _ := resolveThemeFillRef(tblStyleDef.TblBg.FillRef, ctx.themeFmtScheme, themeColors)
			if bgFill != "" {
				hasTblBg = true
				tblBgFill = bgFill
				tableXpx := float64(frameX) / emuPerPx
				tableYpx := float64(frameY) / emuPerPx
				tableWpx := float64(frameCx) / emuPerPx
				tableHpx := float64(frameCy) / emuPerPx
				out.WriteString(fmt.Sprintf(`  <rect x="%.3f" y="%.3f" width="%.3f" height="%.3f" fill="%s"/>
`, tableXpx, tableYpx, tableWpx, tableHpx, tblBgFill))
			}
		}
	}
	_ = hasTblBg
	_ = tblBgFill

	// === Render cell backgrounds ===
	for rowIdx, row := range tbl.Rows {
		colOffset := 0
		for _, cell := range row.Cells {
			actualCol := colOffset
			if actualCol >= numCols {
				break
			}
			if cell.HMerge == "1" || cell.VMerge == "1" {
				colOffset++
				continue
			}

			gridSpan := 1
			if cell.GridSpan > 1 {
				gridSpan = cell.GridSpan
			}
			rowSpan := 1
			if cell.RowSpan > 1 {
				rowSpan = cell.RowSpan
			}

			endCol := actualCol + gridSpan
			if endCol > numCols {
				endCol = numCols
			}
			endRow := rowIdx + rowSpan
			if endRow > numRows {
				endRow = numRows
			}

			x1 := toPxX(colXEmu[actualCol])
			y1 := toPxY(rowYEmu[rowIdx])
			x2 := toPxX(colXEmu[endCol])
			y2 := toPxY(rowYEmu[endRow])

			cellFill, cellOpacity := resolveTableCellFill(cell, rowIdx, isFirstRow, isBandRow, tblStyleDef, themeColors, ctx.themeFmtScheme)

			if cellFill != "none" && cellFill != "" {
				opacityAttr := ""
				if cellOpacity < 1.0 && cellOpacity > 0 {
					opacityAttr = fmt.Sprintf(` fill-opacity="%.1f"`, cellOpacity)
				}
				out.WriteString(fmt.Sprintf(`  <path d="M%.3f %.3f L%.3f %.3f L%.3f %.3f L%.3f %.3fZ" fill="%s"%s fill-rule="evenodd"/>
`, x1, y1, x2, y1, x2, y2, x1, y2, cellFill, opacityAttr))
			}

			colOffset += gridSpan
		}
	}

	// === Render borders ===
	cellBorderLn := func(tcPr *OOXMLTableCellPr, edge string) (string, float64, bool) {
		if tcPr == nil {
			return "", 0, false
		}
		var ln *OOXMLTableBorderLn
		switch edge {
		case "L":
			ln = tcPr.LnL
		case "R":
			ln = tcPr.LnR
		case "T":
			ln = tcPr.LnT
		case "B":
			ln = tcPr.LnB
		}
		if ln == nil || ln.NoFill != nil {
			return "", 0, false
		}
		if ln.SolidFill != nil {
			c := resolveSolidFill(ln.SolidFill, themeColors)
			if c != "" {
				w := float64(ln.W) / float64(emuPerPx)
				return c, w, true
			}
		}
		return "", 0, false
	}

	rowExplicitBorder := func(rowIdx int, edge string) (string, float64, bool) {
		if rowIdx < 0 || rowIdx >= numRows {
			return "", 0, false
		}
		for _, cell := range tbl.Rows[rowIdx].Cells {
			if c, w, ok := cellBorderLn(cell.TcPr, edge); ok {
				return c, w, true
			}
		}
		return "", 0, false
	}

	colExplicitBorder := func(colIdx int, edge string) (string, float64, bool) {
		for _, row := range tbl.Rows {
			if colIdx >= 0 && colIdx < len(row.Cells) {
				if c, w, ok := cellBorderLn(row.Cells[colIdx].TcPr, edge); ok {
					return c, w, true
				}
			}
		}
		return "", 0, false
	}

	writeBorderPath := func(x1, y1, x2, y2 float64, color string, width float64) {
		out.WriteString(fmt.Sprintf(`  <path d="M%.3f %.3f L%.3f %.3f" stroke="%s" stroke-width="%.5f" stroke-linejoin="round" stroke-miterlimit="10" fill="none" fill-rule="evenodd"/>
`, x1, y1, x2, y2, color, width))
	}

	// Resolve style borders from table style definition
	type borderSpec struct {
		color string
		width float64
		draw  bool
	}

	resolveWholeTblBorder := func(edge string) borderSpec {
		if tblStyleDef == nil || tblStyleDef.WholeTbl == nil || tblStyleDef.WholeTbl.TcStyle == nil || tblStyleDef.WholeTbl.TcStyle.TcBdr == nil {
			return borderSpec{"#FFFFFF", 1.33, true}
		}
		bdr := tblStyleDef.WholeTbl.TcStyle.TcBdr
		var sb *OOXMLTableStyleBorder
		switch edge {
		case "left":
			sb = bdr.Left
		case "right":
			sb = bdr.Right
		case "top":
			sb = bdr.Top
		case "bottom":
			sb = bdr.Bottom
		case "insideH":
			sb = bdr.InsideH
		case "insideV":
			sb = bdr.InsideV
		}
		c, w, ok := resolveStyleBorder(sb, themeColors)
		return borderSpec{c, w, ok}
	}

	// Check if the firstRow style defines a special bottom border
	resolveFirstRowBottomBorder := func() borderSpec {
		if !isFirstRow || tblStyleDef == nil || tblStyleDef.FirstRow == nil || tblStyleDef.FirstRow.TcStyle == nil || tblStyleDef.FirstRow.TcStyle.TcBdr == nil {
			return borderSpec{}
		}
		bdr := tblStyleDef.FirstRow.TcStyle.TcBdr
		if bdr.Bottom != nil {
			c, w, ok := resolveStyleBorder(bdr.Bottom, themeColors)
			return borderSpec{c, w, ok}
		}
		return borderSpec{}
	}

	// 1. Interior vertical column separators
	for c := 1; c < numCols; c++ {
		x := toPxX(colXEmu[c])
		y1 := toPxY(rowYEmu[0])
		y2 := toPxY(rowYEmu[numRows])
		if col, w, ok := colExplicitBorder(c-1, "R"); ok {
			writeBorderPath(x, y1, x, y2, col, w)
		} else if col, w, ok := colExplicitBorder(c, "L"); ok {
			writeBorderPath(x, y1, x, y2, col, w)
		} else {
			bs := resolveWholeTblBorder("insideV")
			if bs.draw {
				writeBorderPath(x, y1, x, y2, bs.color, bs.width)
			}
		}
	}

	// 2. Interior horizontal separators
	for r := 1; r < numRows; r++ {
		y := toPxY(rowYEmu[r])
		x1 := toPxX(colXEmu[0])
		x2 := toPxX(colXEmu[numCols])

		if col, w, ok := rowExplicitBorder(r-1, "B"); ok {
			writeBorderPath(x1, y, x2, y, col, w)
		} else if col, w, ok := rowExplicitBorder(r, "T"); ok {
			writeBorderPath(x1, y, x2, y, col, w)
		} else if isFirstRow && r == 1 {
			bs := resolveFirstRowBottomBorder()
			if bs.draw {
				writeBorderPath(x1, y, x2, y, bs.color, bs.width)
			} else {
				bs = resolveWholeTblBorder("insideH")
				if bs.draw {
					writeBorderPath(x1, y, x2, y, bs.color, bs.width)
				}
			}
		} else {
			bs := resolveWholeTblBorder("insideH")
			if bs.draw {
				writeBorderPath(x1, y, x2, y, bs.color, bs.width)
			}
		}
	}

	// 3. Edge verticals
	{
		lx := toPxX(colXEmu[0])
		rx := toPxX(colXEmu[numCols])
		vy1 := toPxY(rowYEmu[0])
		vy2 := toPxY(rowYEmu[numRows])
		if col, w, ok := colExplicitBorder(0, "L"); ok {
			writeBorderPath(lx, vy1, lx, vy2, col, w)
		} else {
			bs := resolveWholeTblBorder("left")
			if bs.draw {
				writeBorderPath(lx, vy1, lx, vy2, bs.color, bs.width)
			}
		}
		if col, w, ok := colExplicitBorder(numCols-1, "R"); ok {
			writeBorderPath(rx, vy1, rx, vy2, col, w)
		} else {
			bs := resolveWholeTblBorder("right")
			if bs.draw {
				writeBorderPath(rx, vy1, rx, vy2, bs.color, bs.width)
			}
		}
	}

	// 4. Edge horizontals
	{
		hx1 := toPxX(colXEmu[0])
		hx2 := toPxX(colXEmu[numCols])
		ty := toPxY(rowYEmu[0])
		by := toPxY(rowYEmu[numRows])
		if col, w, ok := rowExplicitBorder(0, "T"); ok {
			writeBorderPath(hx1, ty, hx2, ty, col, w)
		} else {
			bs := resolveWholeTblBorder("top")
			if bs.draw {
				writeBorderPath(hx1, ty, hx2, ty, bs.color, bs.width)
			}
		}
		if col, w, ok := rowExplicitBorder(numRows-1, "B"); ok {
			writeBorderPath(hx1, by, hx2, by, col, w)
		} else {
			bs := resolveWholeTblBorder("bottom")
			if bs.draw {
				writeBorderPath(hx1, by, hx2, by, bs.color, bs.width)
			}
		}
	}

	// === Render cell text ===
	for rowIdx, row := range tbl.Rows {
		colOffset := 0
		for _, cell := range row.Cells {
			actualCol := colOffset
			if actualCol >= numCols {
				break
			}
			if cell.HMerge == "1" || cell.VMerge == "1" {
				colOffset++
				continue
			}

			gridSpan := 1
			if cell.GridSpan > 1 {
				gridSpan = cell.GridSpan
			}
			rowSpan := 1
			if cell.RowSpan > 1 {
				rowSpan = cell.RowSpan
			}

			endCol := actualCol + gridSpan
			if endCol > numCols {
				endCol = numCols
			}
			endRow := rowIdx + rowSpan
			if endRow > numRows {
				endRow = numRows
			}

			cx1 := toPxX(colXEmu[actualCol])
			cy1 := toPxY(rowYEmu[rowIdx])
			cx2 := toPxX(colXEmu[endCol])
			cy2 := toPxY(rowYEmu[endRow])
			cellW := cx2 - cx1
			cellH := cy2 - cy1

			renderTableCellText(&out, cell, cx1, cy1, cellW, cellH,
				rowIdx, isFirstRow, defaultFontSizePx, tblStyleDef, themeColors, ctx)

			colOffset += gridSpan
		}
	}

	// === Render inner shadow overlay (from tblPr.effectLst.innerShdw) ===
	if tbl.TblPr.EffectLst != nil && tbl.TblPr.EffectLst.InnerShdw != nil {
		renderTableInnerShadow(&out, tbl.TblPr.EffectLst.InnerShdw,
			frameX, frameY, frameCx, frameCy, themeColors)
	}

	return out.String()
}

// resolveShadowColorHex resolves the shadow color from SrgbClr, SchemeClr, or PrstClr to a hex string like "#RRGGBB".
func resolveShadowColorHex(srgb *OOXMLSrgbClr, scheme *OOXMLSchemeClr, prst *OOXMLPrstClr, themeColors map[string]string) string {
	if srgb != nil && srgb.Val != "" {
		return "#" + srgb.Val
	}
	if scheme != nil && scheme.Val != "" {
		return resolveSchemeColor(scheme.Val, themeColors)
	}
	if prst != nil && prst.Val != "" {
		return resolvePresetColor(prst.Val)
	}
	return "#000000"
}

// hexToRGBFractions converts "#RRGGBB" to three float64 values in [0,1].
func hexToRGBFractions(hex string) (float64, float64, float64) {
	hex = strings.TrimPrefix(hex, "#")
	if len(hex) != 6 {
		return 0, 0, 0
	}
	var r, g, b int
	fmt.Sscanf(hex[0:2], "%x", &r)
	fmt.Sscanf(hex[2:4], "%x", &g)
	fmt.Sscanf(hex[4:6], "%x", &b)
	return float64(r) / 255.0, float64(g) / 255.0, float64(b) / 255.0
}

// renderTablePerspectiveShadow renders a perspective shadow for a table using SVG filters and transforms.
// The shadow is a reflected, scaled, optionally skewed copy of the table rendered through a blur+alpha filter.
func renderTablePerspectiveShadow(out *strings.Builder, tbl *OOXMLTable, shdw *OOXMLOuterShadow,
	frameX, frameY, frameCx, frameCy int64,
	scaleX, scaleY float64,
	colXEmu []int64, rowYEmu []int64,
	numCols, numRows int,
	isFirstRow, isBandRow bool,
	tblStyleDef *OOXMLTableStyle, themeColors map[string]string,
	ctx *slideContext) {

	if shdw.Sy == 0 {
		return // no vertical scale means no perspective shadow
	}

	sxFactor := 1.0
	if shdw.Sx != 0 {
		sxFactor = float64(shdw.Sx) / 100000.0
	}
	syFactor := float64(shdw.Sy) / 100000.0 // negative = lower/reflected, positive = upper

	// Shadow color alpha (OOXML default is 100000 = 100% opaque when not explicitly specified)
	alpha := 1.0
	if shdw.PrstClr != nil && shdw.PrstClr.Alpha != nil {
		alpha = float64(shdw.PrstClr.Alpha.Val) / 100000.0
	} else if shdw.SrgbClr != nil && shdw.SrgbClr.Alpha != nil {
		alpha = float64(shdw.SrgbClr.Alpha.Val) / 100000.0
	} else if shdw.SchemeClr != nil && shdw.SchemeClr.Alpha != nil {
		alpha = float64(shdw.SchemeClr.Alpha.Val) / 100000.0
	}

	// Blur standard deviation: blurRad is the blur radius in EMU, stdDev â‰ˆ radius/3
	blurRad, _ := strconv.ParseFloat(shdw.BlurRad, 64)
	blurStdDev := blurRad / emuPerPx / 3.0

	// Distance and direction
	dirEmu, _ := strconv.ParseFloat(shdw.Dir, 64)
	dirRad := dirEmu * math.Pi / 10800000.0

	distEmu, _ := strconv.ParseFloat(shdw.Dist, 64)
	distPx := distEmu / emuPerPx
	distXpx := distPx * math.Cos(dirRad)
	distYpx := distPx * math.Sin(dirRad)

	// Horizontal skew
	kxRad := float64(shdw.Kx) * math.Pi / 10800000.0
	skewC := math.Tan(kxRad)

	// Table pixel dimensions
	tableXpx := float64(frameX) / emuPerPx
	tableYpx := float64(frameY) / emuPerPx
	tableWpx := float64(frameCx) / emuPerPx
	tableHpx := float64(frameCy) / emuPerPx
	tableB := tableYpx + tableHpx

	// Anchor X depends on alignment
	var anchorLocalX float64
	switch shdw.Algn {
	case "br", "r", "tr":
		anchorLocalX = tableWpx
	case "bl", "l", "tl":
		anchorLocalX = 0
	default: // "b", "t", "ctr", or empty â†’ center
		anchorLocalX = tableWpx / 2.0
	}
	anchorScreenX := tableXpx + anchorLocalX

	// Matrix coefficients: matrix(a, 0, c, d, e, f)
	// Content is rendered at local coordinates (0,0) to (tableWpx, tableHpx).
	// For lower shadows (sy<0): d is negative (reflection), shadow below table.
	// For upper shadows (sy>0): d is positive (no reflection), shadow above/within table.
	a := sxFactor
	c := skewC
	d := syFactor
	e := anchorScreenX + distXpx - a*anchorLocalX - c*tableHpx
	f := tableB + distYpx - syFactor*tableHpx

	// Generate a unique filter ID
	filterID := fmt.Sprintf("tblShdw%d%d", frameX/1000, frameY/1000)

	// Shadow color resolved from XML (SrgbClr, SchemeClr, or PrstClr)
	shdwHex := resolveShadowColorHex(shdw.SrgbClr, shdw.SchemeClr, shdw.PrstClr, themeColors)
	shdwR, shdwG, shdwB := hexToRGBFractions(shdwHex)

	// Emit SVG filter
	rVal := fmt.Sprintf("%.4g %.4g", shdwR, shdwR)
	gVal := fmt.Sprintf("%.4g %.4g", shdwG, shdwG)
	bVal := fmt.Sprintf("%.4g %.4g", shdwB, shdwB)

	out.WriteString(fmt.Sprintf(`<defs><filter id="%s" x="-10%%" y="-10%%" width="120%%" height="120%%" filterUnits="userSpaceOnUse" primitiveUnits="userSpaceOnUse"><feComponentTransfer color-interpolation-filters="sRGB"><feFuncR type="discrete" tableValues="%s"/><feFuncG type="discrete" tableValues="%s"/><feFuncB type="discrete" tableValues="%s"/><feFuncA type="linear" slope="%.5g" intercept="0"/></feComponentTransfer><feGaussianBlur stdDeviation="%.5g %.5g"/></filter></defs>`,
		filterID, rVal, gVal, bVal, alpha, blurStdDev, blurStdDev))

	// Local coordinate functions: table content at (0, 0) origin
	toLocalX := func(emu int64) float64 {
		return float64(emu) * scaleX / emuPerPx
	}
	toLocalY := func(emu int64) float64 {
		return float64(emu) * scaleY / emuPerPx
	}

	// Open shadow group with filter and perspective transform
	out.WriteString(fmt.Sprintf(`<g filter="url(#%s)"><g transform="matrix(%.5g 0 %.5g %.5g %.5g %.5g)">`,
		filterID, a, c, d, e, f))

	// Render tblBg fill rect inside the perspective shadow (if present)
	// Use tableWpx/tableHpx (frameCx/Cy / emuPerPx) not toLocalX/Y(frameCx/Cy),
	// because toLocal applies scaleX/Y which inflates dimensions when totalRowHeight < frameCy.
	// Cell content maps to [0, tableHpx] in local coords; the rect must match.
	if tblStyleDef != nil && tblStyleDef.TblBg != nil && tblStyleDef.TblBg.FillRef != nil && ctx.themeFmtScheme != nil {
		bgFill, _ := resolveThemeFillRef(tblStyleDef.TblBg.FillRef, ctx.themeFmtScheme, themeColors)
		if bgFill != "" {
			out.WriteString(fmt.Sprintf(`<rect x="0" y="0" width="%.3f" height="%.3f" fill="%s"/>`, tableWpx, tableHpx, bgFill))
		}
	}

	// Render shadow cell backgrounds (same geometry as real table, at local coords)
	for rowIdx, row := range tbl.Rows {
		colOffset := 0
		for _, cell := range row.Cells {
			actualCol := colOffset
			if actualCol >= numCols {
				break
			}
			if cell.HMerge == "1" || cell.VMerge == "1" {
				colOffset++
				continue
			}

			gridSpan := 1
			if cell.GridSpan > 1 {
				gridSpan = cell.GridSpan
			}
			rowSpan := 1
			if cell.RowSpan > 1 {
				rowSpan = cell.RowSpan
			}

			endCol := actualCol + gridSpan
			if endCol > numCols {
				endCol = numCols
			}
			endRow := rowIdx + rowSpan
			if endRow > numRows {
				endRow = numRows
			}

			x1 := toLocalX(colXEmu[actualCol])
			y1 := toLocalY(rowYEmu[rowIdx])
			x2 := toLocalX(colXEmu[endCol])
			y2 := toLocalY(rowYEmu[endRow])

			cellFill, cellOpacity := resolveTableCellFill(cell, rowIdx, isFirstRow, isBandRow, tblStyleDef, themeColors, ctx.themeFmtScheme)
			if cellFill == "none" || cellFill == "" {
				colOffset += gridSpan
				continue
			}
			opacityAttr := ""
			if cellOpacity < 1.0 && cellOpacity > 0 {
				opacityAttr = fmt.Sprintf(` fill-opacity="%.1f"`, cellOpacity)
			}
			out.WriteString(fmt.Sprintf(`<path d="M%.3f %.3f L%.3f %.3f L%.3f %.3f L%.3f %.3fZ" fill="%s"%s fill-rule="evenodd"/>`,
				x1, y1, x2, y1, x2, y2, x1, y2, cellFill, opacityAttr))

			colOffset += gridSpan
		}
	}

	// Render shadow borders
	writeShadowBorder := func(x1, y1, x2, y2 float64, color string, width float64) {
		out.WriteString(fmt.Sprintf(`<path d="M%.3f %.3f L%.3f %.3f" stroke="%s" stroke-width="%.5f" stroke-linejoin="round" stroke-miterlimit="10" fill="none" fill-rule="evenodd"/>`,
			x1, y1, x2, y2, color, width))
	}

	// Interior vertical separators
	if tblStyleDef != nil && tblStyleDef.WholeTbl != nil && tblStyleDef.WholeTbl.TcStyle != nil &&
		tblStyleDef.WholeTbl.TcStyle.TcBdr != nil && tblStyleDef.WholeTbl.TcStyle.TcBdr.InsideV != nil {
		bdrColor, bdrWidth, ok := resolveStyleBorder(tblStyleDef.WholeTbl.TcStyle.TcBdr.InsideV, themeColors)
		if ok {
			for col := 1; col < numCols; col++ {
				x := toLocalX(colXEmu[col])
				y1 := toLocalY(rowYEmu[0])
				y2 := toLocalY(rowYEmu[numRows])
				writeShadowBorder(x, y1, x, y2, bdrColor, bdrWidth)
			}
		}
	}

	// Interior horizontal separators (including first-row bottom border)
	for r := 1; r < numRows; r++ {
		y := toLocalY(rowYEmu[r])
		x1 := toLocalX(colXEmu[0])
		x2 := toLocalX(colXEmu[numCols])
		drawn := false
		// First-row bottom border override
		if isFirstRow && r == 1 && tblStyleDef != nil && tblStyleDef.FirstRow != nil &&
			tblStyleDef.FirstRow.TcStyle != nil && tblStyleDef.FirstRow.TcStyle.TcBdr != nil &&
			tblStyleDef.FirstRow.TcStyle.TcBdr.Bottom != nil {
			if col, w, ok := resolveStyleBorder(tblStyleDef.FirstRow.TcStyle.TcBdr.Bottom, themeColors); ok {
				writeShadowBorder(x1, y, x2, y, col, w)
				drawn = true
			}
		}
		if !drawn && tblStyleDef != nil && tblStyleDef.WholeTbl != nil && tblStyleDef.WholeTbl.TcStyle != nil &&
			tblStyleDef.WholeTbl.TcStyle.TcBdr != nil && tblStyleDef.WholeTbl.TcStyle.TcBdr.InsideH != nil {
			if col, w, ok := resolveStyleBorder(tblStyleDef.WholeTbl.TcStyle.TcBdr.InsideH, themeColors); ok {
				writeShadowBorder(x1, y, x2, y, col, w)
			}
		}
	}

	// Edge borders
	edgeBorders := []struct {
		edge               string
		x1f, y1f, x2f, y2f func() float64
	}{
		{"left", func() float64 { return toLocalX(colXEmu[0]) }, func() float64 { return toLocalY(rowYEmu[0]) },
			func() float64 { return toLocalX(colXEmu[0]) }, func() float64 { return toLocalY(rowYEmu[numRows]) }},
		{"right", func() float64 { return toLocalX(colXEmu[numCols]) }, func() float64 { return toLocalY(rowYEmu[0]) },
			func() float64 { return toLocalX(colXEmu[numCols]) }, func() float64 { return toLocalY(rowYEmu[numRows]) }},
		{"top", func() float64 { return toLocalX(colXEmu[0]) }, func() float64 { return toLocalY(rowYEmu[0]) },
			func() float64 { return toLocalX(colXEmu[numCols]) }, func() float64 { return toLocalY(rowYEmu[0]) }},
		{"bottom", func() float64 { return toLocalX(colXEmu[0]) }, func() float64 { return toLocalY(rowYEmu[numRows]) },
			func() float64 { return toLocalX(colXEmu[numCols]) }, func() float64 { return toLocalY(rowYEmu[numRows]) }},
	}
	for _, eb := range edgeBorders {
		if tblStyleDef != nil && tblStyleDef.WholeTbl != nil && tblStyleDef.WholeTbl.TcStyle != nil &&
			tblStyleDef.WholeTbl.TcStyle.TcBdr != nil {
			bdr := tblStyleDef.WholeTbl.TcStyle.TcBdr
			var sb *OOXMLTableStyleBorder
			switch eb.edge {
			case "left":
				sb = bdr.Left
			case "right":
				sb = bdr.Right
			case "top":
				sb = bdr.Top
			case "bottom":
				sb = bdr.Bottom
			}
			if col, w, ok := resolveStyleBorder(sb, themeColors); ok {
				writeShadowBorder(eb.x1f(), eb.y1f(), eb.x2f(), eb.y2f(), col, w)
			}
		}
	}

	// Close shadow groups
	out.WriteString("</g></g>")
}

// renderTableInnerShadow renders an inner (inset) shadow on the table using SVG filters.
// The shadow appears inside the table boundary, creating a recessed appearance.
func renderTableInnerShadow(out *strings.Builder, shdw *OOXMLInnerShadow,
	frameX, frameY, frameCx, frameCy int64, themeColors map[string]string) {

	// OOXML default alpha is 100% when not explicitly specified
	alpha := 1.0
	if shdw.PrstClr != nil && shdw.PrstClr.Alpha != nil {
		alpha = float64(shdw.PrstClr.Alpha.Val) / 100000.0
	} else if shdw.SrgbClr != nil && shdw.SrgbClr.Alpha != nil {
		alpha = float64(shdw.SrgbClr.Alpha.Val) / 100000.0
	} else if shdw.SchemeClr != nil && shdw.SchemeClr.Alpha != nil {
		alpha = float64(shdw.SchemeClr.Alpha.Val) / 100000.0
	}

	blurRad, _ := strconv.ParseFloat(shdw.BlurRad, 64)
	blurStdDev := blurRad / emuPerPx / 3.0

	dirEmu, _ := strconv.ParseFloat(shdw.Dir, 64)
	dirRad := dirEmu * math.Pi / 10800000.0

	distEmu, _ := strconv.ParseFloat(shdw.Dist, 64)
	distPx := distEmu / emuPerPx

	// Negate the offset: the feOffset shifts the inverted alpha mask, so the
	// shadow appears on the OPPOSITE side of the offset. To place shadow at
	// the direction indicated by dir, we shift the mask the other way.
	dx := -distPx * math.Cos(dirRad)
	dy := -distPx * math.Sin(dirRad)

	tableXpx := float64(frameX) / emuPerPx
	tableYpx := float64(frameY) / emuPerPx
	tableWpx := float64(frameCx) / emuPerPx
	tableHpx := float64(frameCy) / emuPerPx

	shadowColor := resolveShadowColorHex(shdw.SrgbClr, shdw.SchemeClr, shdw.PrstClr, themeColors)

	filterID := fmt.Sprintf("tblInnerShdw%d%d", frameX/1000, frameY/1000)

	out.WriteString(fmt.Sprintf(`<defs><filter id="%s" x="-50%%" y="-50%%" width="200%%" height="200%%" filterUnits="objectBoundingBox" primitiveUnits="userSpaceOnUse">`, filterID))
	out.WriteString(`<feComponentTransfer in="SourceAlpha"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer>`)
	out.WriteString(fmt.Sprintf(`<feOffset dx="%.4f" dy="%.4f" result="offsetInvAlpha"/>`, dx, dy))
	out.WriteString(fmt.Sprintf(`<feGaussianBlur stdDeviation="%.4f" result="blurredInvAlpha"/>`, blurStdDev))
	out.WriteString(`<feComposite operator="in" in="blurredInvAlpha" in2="SourceAlpha" result="innerShadowAlpha"/>`)
	out.WriteString(fmt.Sprintf(`<feFlood flood-color="%s" flood-opacity="%.4f" result="color"/>`, shadowColor, alpha))
	out.WriteString(`<feComposite operator="in" in="color" in2="innerShadowAlpha"/>`)
	out.WriteString(`</filter></defs>`)

	// Use fill="black" so SourceAlpha is fully opaque within the rect bounds.
	// The filter replaces visual rendering entirely (no feMerge with SourceGraphic),
	// so only the semi-transparent shadow overlay is rendered â€” table content below
	// remains visible.
	out.WriteString(fmt.Sprintf(`<rect x="%.3f" y="%.3f" width="%.3f" height="%.3f" fill="black" fill-opacity="1" filter="url(#%s)"/>`,
		tableXpx, tableYpx, tableWpx, tableHpx, filterID))
}

// renderTableBgShadow renders a simple offset shadow from the table style's tblBg effectRef.
// Unlike perspective shadows, this is a non-reflected copy offset by the shadow distance.
func renderTableBgShadow(out *strings.Builder, tbl *OOXMLTable, shdw *OOXMLOuterShadow,
	frameX, frameY, frameCx, frameCy int64,
	scaleX, scaleY float64,
	colXEmu []int64, rowYEmu []int64,
	numCols, numRows int,
	isFirstRow, isBandRow bool,
	tblStyleDef *OOXMLTableStyle, themeColors map[string]string,
	ctx *slideContext) {

	alpha := 1.0 // OOXML default: 100% opaque when not explicitly specified
	if shdw.PrstClr != nil && shdw.PrstClr.Alpha != nil {
		alpha = float64(shdw.PrstClr.Alpha.Val) / 100000.0
	} else if shdw.SrgbClr != nil && shdw.SrgbClr.Alpha != nil {
		alpha = float64(shdw.SrgbClr.Alpha.Val) / 100000.0
	} else if shdw.SchemeClr != nil && shdw.SchemeClr.Alpha != nil {
		alpha = float64(shdw.SchemeClr.Alpha.Val) / 100000.0
	}

	blurRad, _ := strconv.ParseFloat(shdw.BlurRad, 64)
	blurStdDev := blurRad / emuPerPx / 3.0

	dirEmu, _ := strconv.ParseFloat(shdw.Dir, 64)
	dirRad := dirEmu * math.Pi / 10800000.0

	distEmu, _ := strconv.ParseFloat(shdw.Dist, 64)
	distPx := distEmu / emuPerPx
	distXpx := distPx * math.Cos(dirRad)
	distYpx := distPx * math.Sin(dirRad)

	tableXpx := float64(frameX) / emuPerPx
	tableYpx := float64(frameY) / emuPerPx
	tableWpx := float64(frameCx) / emuPerPx
	tableHpx := float64(frameCy) / emuPerPx

	filterID := fmt.Sprintf("tblBgShdw%d%d", frameX/1000, frameY/1000)

	shdwHex := resolveShadowColorHex(shdw.SrgbClr, shdw.SchemeClr, shdw.PrstClr, themeColors)
	bgR, bgG, bgB := hexToRGBFractions(shdwHex)
	rVal := fmt.Sprintf("%.4g %.4g", bgR, bgR)
	gVal := fmt.Sprintf("%.4g %.4g", bgG, bgG)
	bVal := fmt.Sprintf("%.4g %.4g", bgB, bgB)

	out.WriteString(fmt.Sprintf(`<defs><filter id="%s" x="-10%%" y="-10%%" width="120%%" height="120%%" filterUnits="userSpaceOnUse" primitiveUnits="userSpaceOnUse"><feComponentTransfer color-interpolation-filters="sRGB"><feFuncR type="discrete" tableValues="%s"/><feFuncG type="discrete" tableValues="%s"/><feFuncB type="discrete" tableValues="%s"/><feFuncA type="linear" slope="%.5g" intercept="0"/></feComponentTransfer><feGaussianBlur stdDeviation="%.5g %.5g"/></filter></defs>`,
		filterID, rVal, gVal, bVal, alpha, blurStdDev, blurStdDev))

	shdwX := tableXpx + distXpx
	shdwY := tableYpx + distYpx

	out.WriteString(fmt.Sprintf(`<g filter="url(#%s)"><g transform="translate(%.3f, %.3f)">`, filterID, shdwX, shdwY))

	// Background rect covering the entire table area
	out.WriteString(fmt.Sprintf(`<rect x="0" y="0" width="%.3f" height="%.3f" fill="#FF0000"/>`, tableWpx, tableHpx))

	// Cell fill paths (colXEmu/rowYEmu are already 0-based table-local)
	toLocalX := func(emu int64) float64 { return float64(emu) / emuPerPx }
	toLocalY := func(emu int64) float64 { return float64(emu) / emuPerPx }

	for rowIdx, row := range tbl.Rows {
		colOffset := 0
		for _, cell := range row.Cells {
			actualCol := colOffset
			if actualCol >= numCols {
				break
			}
			if cell.HMerge == "1" || cell.VMerge == "1" {
				colOffset++
				continue
			}
			gridSpan := 1
			if cell.GridSpan > 1 {
				gridSpan = cell.GridSpan
			}
			rowSpan := 1
			if cell.RowSpan > 1 {
				rowSpan = cell.RowSpan
			}
			endCol := actualCol + gridSpan
			if endCol > numCols {
				endCol = numCols
			}
			endRow := rowIdx + rowSpan
			if endRow > numRows {
				endRow = numRows
			}

			x1 := toLocalX(colXEmu[actualCol])
			y1 := toLocalY(rowYEmu[rowIdx])
			x2 := toLocalX(colXEmu[endCol])
			y2 := toLocalY(rowYEmu[endRow])

			cellFill, cellOpacity := resolveTableCellFill(cell, rowIdx, isFirstRow, isBandRow, tblStyleDef, themeColors, ctx.themeFmtScheme)
			if cellFill == "none" || cellFill == "" {
				colOffset += gridSpan
				continue
			}
			opacityAttr := ""
			if cellOpacity < 1.0 && cellOpacity > 0 {
				opacityAttr = fmt.Sprintf(` fill-opacity="%.1f"`, cellOpacity)
			}
			out.WriteString(fmt.Sprintf(`<path d="M%.3f %.3f L%.3f %.3f L%.3f %.3f L%.3f %.3fZ" fill="%s"%s fill-rule="evenodd"/>`,
				x1, y1, x2, y1, x2, y2, x1, y2, cellFill, opacityAttr))
			colOffset += gridSpan
		}
	}

	// Shadow borders
	writeShadowBorder := func(x1, y1, x2, y2 float64, color string, width float64) {
		out.WriteString(fmt.Sprintf(`<path d="M%.3f %.3f L%.3f %.3f" stroke="%s" stroke-width="%.5f" stroke-linejoin="round" stroke-miterlimit="10" fill="none" fill-rule="evenodd"/>`,
			x1, y1, x2, y2, color, width))
	}

	// Interior vertical separators
	if tblStyleDef != nil && tblStyleDef.WholeTbl != nil && tblStyleDef.WholeTbl.TcStyle != nil &&
		tblStyleDef.WholeTbl.TcStyle.TcBdr != nil && tblStyleDef.WholeTbl.TcStyle.TcBdr.InsideV != nil {
		bdrColor, bdrWidth, ok := resolveStyleBorder(tblStyleDef.WholeTbl.TcStyle.TcBdr.InsideV, themeColors)
		if ok {
			for col := 1; col < numCols; col++ {
				x := toLocalX(colXEmu[col])
				y1 := toLocalY(rowYEmu[0])
				y2 := toLocalY(rowYEmu[numRows])
				writeShadowBorder(x, y1, x, y2, bdrColor, bdrWidth)
			}
		}
	}

	// Interior horizontal separators (including first-row bottom border)
	for r := 1; r < numRows; r++ {
		y := toLocalY(rowYEmu[r])
		x1 := toLocalX(colXEmu[0])
		x2 := toLocalX(colXEmu[numCols])
		drawn := false
		if isFirstRow && r == 1 && tblStyleDef != nil && tblStyleDef.FirstRow != nil &&
			tblStyleDef.FirstRow.TcStyle != nil && tblStyleDef.FirstRow.TcStyle.TcBdr != nil &&
			tblStyleDef.FirstRow.TcStyle.TcBdr.Bottom != nil {
			if col, w, ok := resolveStyleBorder(tblStyleDef.FirstRow.TcStyle.TcBdr.Bottom, themeColors); ok {
				writeShadowBorder(x1, y, x2, y, col, w)
				drawn = true
			}
		}
		if !drawn && tblStyleDef != nil && tblStyleDef.WholeTbl != nil && tblStyleDef.WholeTbl.TcStyle != nil &&
			tblStyleDef.WholeTbl.TcStyle.TcBdr != nil && tblStyleDef.WholeTbl.TcStyle.TcBdr.InsideH != nil {
			if col, w, ok := resolveStyleBorder(tblStyleDef.WholeTbl.TcStyle.TcBdr.InsideH, themeColors); ok {
				writeShadowBorder(x1, y, x2, y, col, w)
			}
		}
	}

	// Edge borders
	edgeBorders := []struct {
		edge               string
		x1f, y1f, x2f, y2f func() float64
	}{
		{"left", func() float64 { return toLocalX(colXEmu[0]) }, func() float64 { return toLocalY(rowYEmu[0]) },
			func() float64 { return toLocalX(colXEmu[0]) }, func() float64 { return toLocalY(rowYEmu[numRows]) }},
		{"right", func() float64 { return toLocalX(colXEmu[numCols]) }, func() float64 { return toLocalY(rowYEmu[0]) },
			func() float64 { return toLocalX(colXEmu[numCols]) }, func() float64 { return toLocalY(rowYEmu[numRows]) }},
		{"top", func() float64 { return toLocalX(colXEmu[0]) }, func() float64 { return toLocalY(rowYEmu[0]) },
			func() float64 { return toLocalX(colXEmu[numCols]) }, func() float64 { return toLocalY(rowYEmu[0]) }},
		{"bottom", func() float64 { return toLocalX(colXEmu[0]) }, func() float64 { return toLocalY(rowYEmu[numRows]) },
			func() float64 { return toLocalX(colXEmu[numCols]) }, func() float64 { return toLocalY(rowYEmu[numRows]) }},
	}
	for _, eb := range edgeBorders {
		if tblStyleDef != nil && tblStyleDef.WholeTbl != nil && tblStyleDef.WholeTbl.TcStyle != nil &&
			tblStyleDef.WholeTbl.TcStyle.TcBdr != nil {
			bdr := tblStyleDef.WholeTbl.TcStyle.TcBdr
			var sb *OOXMLTableStyleBorder
			switch eb.edge {
			case "left":
				sb = bdr.Left
			case "right":
				sb = bdr.Right
			case "top":
				sb = bdr.Top
			case "bottom":
				sb = bdr.Bottom
			}
			if sb != nil {
				if col, w, ok := resolveStyleBorder(sb, themeColors); ok {
					writeShadowBorder(eb.x1f(), eb.y1f(), eb.x2f(), eb.y2f(), col, w)
				}
			}
		}
	}

	out.WriteString("</g></g>")
}

// resolveThemeFillRef resolves a fillRef against the theme's fillStyleLst.
// It returns the resolved fill color and opacity. For gradient fills, it
// approximates with the midpoint gradient stop color.
func resolveThemeFillRef(fillRef *OOXMLStyleRef, fmtScheme *OOXMLFmtScheme, themeColors map[string]string) (string, float64) {
	if fillRef == nil || fmtScheme == nil || fillRef.SchemeClr == nil {
		return "", 1.0
	}
	idx, _ := strconv.Atoi(fillRef.Idx)
	if idx <= 0 || idx > len(fmtScheme.FillStyleLst.Entries) {
		return "", 1.0
	}
	baseColor := resolveSchemeColor(fillRef.SchemeClr.Val, themeColors)
	if baseColor == "" {
		return "", 1.0
	}

	entry := fmtScheme.FillStyleLst.Entries[idx-1]

	if entry.SolidFill != nil {
		return baseColor, 1.0
	}

	if entry.GradFill != nil && len(entry.GradFill.GsLst) > 0 {
		midStop := entry.GradFill.GsLst[len(entry.GradFill.GsLst)/2]
		color := baseColor
		if midStop.SchemeClr != nil {
			if midStop.SchemeClr.Tint != nil {
				color = applyTint(color, midStop.SchemeClr.Tint.Val)
			}
			if midStop.SchemeClr.Shade != nil {
				color = applyShade(color, midStop.SchemeClr.Shade.Val)
			}
			if midStop.SchemeClr.SatMod != nil {
				color = applySatMod(color, midStop.SchemeClr.SatMod.Val)
			}
		}
		return color, 1.0
	}

	return baseColor, 1.0
}

// resolveThemeEffectRef resolves an effectRef against the theme's effectStyleLst.
// Returns the outerShdw definition if present, nil otherwise.
func resolveThemeEffectRef(effectRef *OOXMLStyleRef, fmtScheme *OOXMLFmtScheme) *OOXMLOuterShadow {
	if effectRef == nil || fmtScheme == nil {
		return nil
	}
	idx, _ := strconv.Atoi(effectRef.Idx)
	if idx <= 0 || idx > len(fmtScheme.EffectStyleLst.EffectStyles) {
		return nil
	}
	style := fmtScheme.EffectStyleLst.EffectStyles[idx-1]
	if style.EffectLst != nil && style.EffectLst.OuterShdw != nil {
		return style.EffectLst.OuterShdw
	}
	return nil
}

// resolveTableCellFill determines the fill color and opacity for a table cell using data-driven style definitions.
func resolveTableCellFill(cell OOXMLTableCell, rowIdx int, isFirstRow, isBandRow bool, styleDef *OOXMLTableStyle, themeColors map[string]string, fmtScheme *OOXMLFmtScheme) (string, float64) {
	// Explicit cell fill overrides table style
	if cell.TcPr != nil {
		if cell.TcPr.NoFill != nil {
			return "none", 1.0
		}
		if cell.TcPr.SolidFill != nil {
			c := resolveSolidFill(cell.TcPr.SolidFill, themeColors)
			if c != "" {
				return c, 1.0
			}
		}
	}

	if styleDef == nil {
		if isFirstRow && rowIdx == 0 {
			return resolveSchemeColor("accent1", themeColors), 1.0
		}
		return "#FFFFFF", 1.0
	}

	// Header row: use firstRow fill from style definition
	if isFirstRow && rowIdx == 0 {
		c, op := resolveStylePartFill(styleDef.FirstRow, themeColors)
		if c != "" && c != "none" {
			return c, op
		}
		c, op = resolveStylePartFill(styleDef.WholeTbl, themeColors)
		if c != "" {
			return c, op
		}
		return resolveSchemeColor("accent1", themeColors), 1.0
	}

	// Determine data row index (offset by 1 if first row is a header)
	dataRowIdx := rowIdx
	if isFirstRow {
		dataRowIdx = rowIdx - 1
	}

	// Banding: alternate between band1H and band2H (or wholeTbl fill)
	if isBandRow && dataRowIdx >= 0 {
		if dataRowIdx%2 == 0 {
			c, op := resolveStylePartFill(styleDef.Band1H, themeColors)
			if c != "" && c != "none" {
				return c, op
			}
		} else {
			c, op := resolveStylePartFill(styleDef.Band2H, themeColors)
			if c != "" && c != "none" {
				return c, op
			}
		}
	}

	// Default: wholeTbl fill
	c, op := resolveStylePartFill(styleDef.WholeTbl, themeColors)
	if c != "" && c != "none" {
		return c, op
	}

	// Fall back to table background (tblBg) fillRef resolved via theme
	if styleDef.TblBg != nil && styleDef.TblBg.FillRef != nil {
		bgColor, bgOp := resolveThemeFillRef(styleDef.TblBg.FillRef, fmtScheme, themeColors)
		if bgColor != "" {
			return bgColor, bgOp
		}
	}

	if c == "none" {
		return "none", 1.0
	}
	return "#FFFFFF", 1.0
}

// renderTableCellText renders the text content of a table cell as SVG text elements.
func renderTableCellText(out *strings.Builder, cell OOXMLTableCell,
	cellX, cellY, cellW, cellH float64,
	rowIdx int, isFirstRow bool, defaultFontSizePx int,
	styleDef *OOXMLTableStyle, themeColors map[string]string, ctx *slideContext) {

	if cell.TxBody == nil || len(cell.TxBody.Paragraphs) == 0 {
		return
	}

	// Cell margins in EMU (OOXML defaults)
	marginL := int64(91440)
	marginR := int64(91440)
	marginT := int64(45720)
	marginB := int64(45720)

	if cell.TcPr != nil {
		if cell.TcPr.MarL != nil {
			marginL = *cell.TcPr.MarL
		}
		if cell.TcPr.MarR != nil {
			marginR = *cell.TcPr.MarR
		}
		if cell.TcPr.MarT != nil {
			marginT = *cell.TcPr.MarT
		}
		if cell.TcPr.MarB != nil {
			marginB = *cell.TcPr.MarB
		}
	}

	lPx := float64(emuToPx(marginL))
	rPx := float64(emuToPx(marginR))
	tPx := float64(emuToPx(marginT))
	_ = float64(emuToPx(marginB))

	// Resolve default text properties from the table style definition
	isHeader := isFirstRow && rowIdx == 0
	defaultColor := "#000000"
	defaultBold := false
	defaultItalic := false
	defaultFontWeight := "400"

	// Apply style-based text defaults
	if styleDef != nil {
		// wholeTbl text style provides defaults for all cells
		if bold, italic, color := resolveStyleTextProps(styleDef.WholeTbl, themeColors); color != "" {
			defaultColor = color
			defaultBold = bold
			defaultItalic = italic
		}

		// firstRow overrides for header
		if isHeader {
			if bold, italic, color := resolveStyleTextProps(styleDef.FirstRow, themeColors); color != "" {
				defaultColor = color
				defaultBold = bold
				defaultItalic = italic
			} else if styleDef.FirstRow != nil && styleDef.FirstRow.TcTxStyle != nil {
				if styleDef.FirstRow.TcTxStyle.B == "on" {
					defaultBold = true
				}
			}
		}
	} else if isHeader {
		defaultBold = true
		defaultColor = "#FFFFFF"
	}

	if defaultBold {
		defaultFontWeight = "700"
	}

	// Resolve minor font from theme (table text uses minor font by default)
	defaultFontFamily := "Segoe UI"
	if ctx != nil && ctx.minorFont != "" {
		defaultFontFamily = ctx.minorFont
	}

	type textLine struct {
		text        string
		fontSize    int
		color       string
		fontWeight  string
		fontFamily  string
		fontStretch string
		fontStyle   string
		alignment   string
		lineH       float64
	}
	var lines []textLine

	for _, para := range cell.TxBody.Paragraphs {
		paraAlign := ""
		if para.PPr != nil && para.PPr.Algn != "" {
			paraAlign = para.PPr.Algn
		}

		if len(para.Runs) == 0 {
			continue
		}

		var textParts []string
		fs := defaultFontSizePx
		color := defaultColor
		bold := defaultBold
		fontWeight := defaultFontWeight
		fontFamily := defaultFontFamily
		fontStretch := ""
		fontStyle := ""
		if defaultItalic {
			fontStyle = "italic"
		}

		for _, run := range para.Runs {
			t := run.T
			if t == "" {
				continue
			}
			textParts = append(textParts, t)

			if run.RPr.Sz > 0 {
				fs = int(math.Round(float64(run.RPr.Sz) / 100.0 * 96.0 / 72.0))
			}
			if run.RPr.SolidFill != nil {
				c := resolveSolidFill(run.RPr.SolidFill, themeColors)
				if c != "" {
					color = c
				}
			}
			if run.RPr.B == "1" {
				bold = true
			}
			if bold {
				fontWeight = "700"
			}
			if run.RPr.I == "1" {
				fontStyle = "italic"
			}
			if run.RPr.Latin != nil && run.RPr.Latin.Typeface != "" && !strings.HasPrefix(run.RPr.Latin.Typeface, "+") {
				fontFamily = run.RPr.Latin.Typeface
				nameLower := strings.ToLower(fontFamily)
				if strings.Contains(nameLower, "narrow") || strings.Contains(nameLower, "condensed") {
					fontStretch = "condensed"
				}
			}
		}

		paraText := strings.Join(textParts, "")
		if strings.TrimSpace(paraText) == "" {
			continue
		}

		// Word wrap
		textAreaW := cellW - lPx - rPx
		if textAreaW < 20 {
			textAreaW = 20
		}
		charWidth := float64(fs) * 0.55
		if charWidth < 4 {
			charWidth = 4
		}
		maxChars := int(textAreaW / charWidth)
		if maxChars < 3 {
			maxChars = 3
		}

		wrapped := wrapText(paraText, maxChars)
		lineH := float64(fs) * 1.3

		for _, wl := range wrapped {
			lines = append(lines, textLine{
				text:        wl,
				fontSize:    fs,
				color:       color,
				fontWeight:  fontWeight,
				fontFamily:  fontFamily,
				fontStretch: fontStretch,
				fontStyle:   fontStyle,
				alignment:   paraAlign,
				lineH:       lineH,
			})
		}
	}

	if len(lines) == 0 {
		return
	}

	totalH := 0.0
	for _, l := range lines {
		totalH += l.lineH
	}

	vAnchor := "t"
	if cell.TcPr != nil && cell.TcPr.Anchor != "" {
		vAnchor = cell.TcPr.Anchor
	}

	var curY float64
	switch vAnchor {
	case "ctr":
		curY = cellY + (cellH-totalH)/2
	case "b":
		curY = cellY + cellH - totalH - tPx
	default:
		curY = cellY + tPx
	}

	for _, l := range lines {
		baselineY := curY + float64(l.fontSize)*0.85
		if l.text == "" {
			curY += l.lineH
			continue
		}

		textX := cellX + lPx
		anchorAttr := ""
		switch l.alignment {
		case "ctr":
			textX = cellX + cellW/2
			anchorAttr = ` text-anchor="middle"`
		case "r":
			textX = cellX + cellW - rPx
			anchorAttr = ` text-anchor="end"`
		}

		fontFamAttr := ""
		if l.fontFamily != "" {
			fontFamAttr = fmt.Sprintf(` font-family="%s"`, escapeXMLText(l.fontFamily))
		}
		fontStyleAttr := ""
		if l.fontStyle != "" {
			fontStyleAttr = fmt.Sprintf(` font-style="%s"`, l.fontStyle)
		}
		fontStretchAttr := ""
		if l.fontStretch != "" {
			fontStretchAttr = fmt.Sprintf(` font-stretch="%s"`, l.fontStretch)
		}

		out.WriteString(fmt.Sprintf(`  <text fill="%s"%s%s font-weight="%s"%s font-size="%d"%s transform="matrix(1 0 0 1 %.3f %.0f)">%s</text>
`, l.color, fontFamAttr, fontStyleAttr, l.fontWeight, fontStretchAttr, l.fontSize, anchorAttr, textX, baselineY, escapeXMLText(l.text)))

		curY += l.lineH
	}
}
