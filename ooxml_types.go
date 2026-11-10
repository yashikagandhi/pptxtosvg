package pptxtosvg

import (
	"encoding/xml"
	"strings"
)

// ====================================================================
// OOXML Type Definitions for PowerPoint Slide Parsing (ECMA-376)
//
// Namespace prefixes used in PPTX:
//   p:  http://schemas.openxmlformats.org/presentationml/2006/main
//   a:  http://schemas.openxmlformats.org/drawingml/2006/main
//   r:  http://schemas.openxmlformats.org/officeDocument/2006/relationships
//
// Go struct tags use local names only (no namespace prefix) since
// there are no local-name collisions across namespaces in our usage.
// ====================================================================

// ======================== Presentation Root ========================

// OOXMLPresentation represents the root <p:presentation> element from presentation.xml.
// We only parse defaultTextStyle which provides fallback font sizes for shapes without
// explicit font size or placeholder inheritance.
type OOXMLPresentation struct {
	XMLName          xml.Name            `xml:"presentation"`
	DefaultTextStyle *OOXMLTextListStyle `xml:"defaultTextStyle"`
}

// ======================== Slide Root ========================

// OOXMLSlide represents the root <p:sld> element.
type OOXMLSlide struct {
	XMLName xml.Name  `xml:"sld"`
	CSld    OOXMLCSld `xml:"cSld"`
}

// OOXMLSlideLayout represents <p:sldLayout>.
type OOXMLSlideLayout struct {
	XMLName xml.Name  `xml:"sldLayout"`
	CSld    OOXMLCSld `xml:"cSld"`
}

// OOXMLSlideMaster represents <p:sldMaster>.
type OOXMLSlideMaster struct {
	XMLName  xml.Name       `xml:"sldMaster"`
	CSld     OOXMLCSld      `xml:"cSld"`
	ClrMap   *OOXMLClrMap   `xml:"clrMap"` // maps logical names (tx1, bg1, ...) to theme slot (dk1, lt1, ...)
	TxStyles *OOXMLTxStyles `xml:"txStyles"`
}

// OOXMLClrMap represents <p:clrMap> on master/slide: maps logical color roles to theme scheme index.
// e.g. tx1="dk1" means "text 1" uses theme's dk1; bg1="lt1" means "background 1" uses theme's lt1.
type OOXMLClrMap struct {
	Bg1      string `xml:"bg1,attr,omitempty"`
	Tx1      string `xml:"tx1,attr,omitempty"`
	Bg2      string `xml:"bg2,attr,omitempty"`
	Tx2      string `xml:"tx2,attr,omitempty"`
	Dk1      string `xml:"dk1,attr,omitempty"`
	Lt1      string `xml:"lt1,attr,omitempty"`
	Dk2      string `xml:"dk2,attr,omitempty"`
	Lt2      string `xml:"lt2,attr,omitempty"`
	Accent1  string `xml:"accent1,attr,omitempty"`
	Accent2  string `xml:"accent2,attr,omitempty"`
	Accent3  string `xml:"accent3,attr,omitempty"`
	Accent4  string `xml:"accent4,attr,omitempty"`
	Accent5  string `xml:"accent5,attr,omitempty"`
	Accent6  string `xml:"accent6,attr,omitempty"`
	Hlink    string `xml:"hlink,attr,omitempty"`
	FolHlink string `xml:"folHlink,attr,omitempty"`
}

// OOXMLTxStyles holds master-level title/body/other text styles.
type OOXMLTxStyles struct {
	TitleStyle *OOXMLTextListStyle `xml:"titleStyle"`
	BodyStyle  *OOXMLTextListStyle `xml:"bodyStyle"`
}

// OOXMLTextListStyle represents a list style with per-level paragraph defaults.
type OOXMLTextListStyle struct {
	Lvl1pPr *OOXMLLvlPPr `xml:"lvl1pPr"`
	Lvl2pPr *OOXMLLvlPPr `xml:"lvl2pPr"`
}

// OOXMLLvlPPr represents a level-specific paragraph property default.
type OOXMLLvlPPr struct {
	MarL      int64           `xml:"marL,attr,omitempty"`
	Indent    int64           `xml:"indent,attr,omitempty"`
	Algn      string          `xml:"algn,attr,omitempty"`
	SpcBef    *OOXMLSpacing   `xml:"spcBef"`
	SpcAft    *OOXMLSpacing   `xml:"spcAft"`
	LnSpc     *OOXMLSpacing   `xml:"lnSpc"`
	DefRPr    *OOXMLDefRPr    `xml:"defRPr"`
	BuFont    *OOXMLFont      `xml:"buFont"`
	BuChar    *OOXMLBuChar    `xml:"buChar"`
	BuAutoNum *OOXMLBuAutoNum `xml:"buAutoNum"`
	BuNone    *struct{}       `xml:"buNone"`
}

// OOXMLDefRPr holds default run properties (font size, color, etc.) in txStyles/lstStyle.
type OOXMLDefRPr struct {
	Sz        int             `xml:"sz,attr,omitempty"` // hundredths of a point
	SolidFill *OOXMLSolidFill `xml:"solidFill"`
}

// OOXMLCSld represents <p:cSld> (common slide data).
type OOXMLCSld struct {
	Bg     *OOXMLBg    `xml:"bg"`
	SpTree OOXMLSpTree `xml:"spTree"`
}

// OOXMLBg represents <p:bg> (slide background).
type OOXMLBg struct {
	BgPr *OOXMLBgPr `xml:"bgPr"`
}

// OOXMLBgPr represents <p:bgPr> (background properties).
type OOXMLBgPr struct {
	SolidFill *OOXMLSolidFill `xml:"solidFill"`
	GradFill  *OOXMLGradFill  `xml:"gradFill"`
}

// SpTreeChildRef records one child of a shape tree in document order (ECMA-376 19.3.1.45:
// z-order is lexical order; first in XML = back, last = front).
type SpTreeChildRef struct {
	Kind  string // "sp", "grpSp", "cxnSp", "pic", "graphicFrame"
	Index int    // index into the corresponding slice
}

type OOXMLAlternateContent struct {
	Choice   *OOXMLChoice   `xml:"Choice"`
	Fallback *OOXMLFallback `xml:"Fallback"`
}

type OOXMLChoice struct {
	GraphicFrame *OOXMLGraphicFrame `xml:"graphicFrame"`
	Shape        *OOXMLShape        `xml:"sp"`
	Picture      *OOXMLPictureShape `xml:"pic"`
}

type OOXMLFallback struct {
	GraphicFrame *OOXMLGraphicFrame `xml:"graphicFrame"`
	Shape        *OOXMLShape        `xml:"sp"`
	Picture      *OOXMLPictureShape `xml:"pic"`
}

// OOXMLSpTree represents <p:spTree>, the shape tree containing all slide content.
// Children (sp, grpSp, cxnSp, pic, graphicFrame) are collected into typed slices.
// ChildOrder, when non-nil, gives the document order of those children for correct z-order.
type OOXMLSpTree struct {
	NvGrpSpPr     OOXMLNvGrpSpPr        `xml:"nvGrpSpPr"`
	GrpSpPr       OOXMLGrpSpPr          `xml:"grpSpPr"`
	Shapes        []OOXMLShape          `xml:"sp"`
	Groups        []OOXMLGroupShape     `xml:"grpSp"`
	Connectors    []OOXMLConnectorShape `xml:"cxnSp"`
	Pictures      []OOXMLPictureShape   `xml:"pic"`
	GraphicFrames []OOXMLGraphicFrame   `xml:"graphicFrame"`
	ChildOrder    []SpTreeChildRef      `xml:"-"` // filled by UnmarshalXML for z-order
}

// UnmarshalXML preserves document order of shape-tree children for correct z-order (ECMA-376 19.3.1.45).
func (t *OOXMLSpTree) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch se := tok.(type) {
		case xml.StartElement:
			switch se.Name.Local {
			case "nvGrpSpPr":
				if err := d.DecodeElement(&t.NvGrpSpPr, &se); err != nil {
					return err
				}
			case "grpSpPr":
				if err := d.DecodeElement(&t.GrpSpPr, &se); err != nil {
					return err
				}
			case "sp":
				var s OOXMLShape
				if err := d.DecodeElement(&s, &se); err != nil {
					return err
				}
				t.Shapes = append(t.Shapes, s)
				t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"sp", len(t.Shapes) - 1})
			case "grpSp":
				var g OOXMLGroupShape
				if err := d.DecodeElement(&g, &se); err != nil {
					return err
				}
				t.Groups = append(t.Groups, g)
				t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"grpSp", len(t.Groups) - 1})
			case "cxnSp":
				var c OOXMLConnectorShape
				if err := d.DecodeElement(&c, &se); err != nil {
					return err
				}
				t.Connectors = append(t.Connectors, c)
				t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"cxnSp", len(t.Connectors) - 1})
			case "pic":
				var p OOXMLPictureShape
				if err := d.DecodeElement(&p, &se); err != nil {
					return err
				}
				t.Pictures = append(t.Pictures, p)
				t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"pic", len(t.Pictures) - 1})
			case "graphicFrame":
				var gf OOXMLGraphicFrame
				if err := d.DecodeElement(&gf, &se); err != nil {
					return err
				}
				t.GraphicFrames = append(t.GraphicFrames, gf)
				t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"graphicFrame", len(t.GraphicFrames) - 1})
			case "AlternateContent":
				var ac OOXMLAlternateContent
				if err := d.DecodeElement(&ac, &se); err != nil {
					return err
				}
				if ac.Choice != nil {
					if ac.Choice.GraphicFrame != nil {
						t.GraphicFrames = append(t.GraphicFrames, *ac.Choice.GraphicFrame)
						t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"graphicFrame", len(t.GraphicFrames) - 1})
					} else if ac.Choice.Shape != nil {
						t.Shapes = append(t.Shapes, *ac.Choice.Shape)
						t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"sp", len(t.Shapes) - 1})
					} else if ac.Choice.Picture != nil {
						t.Pictures = append(t.Pictures, *ac.Choice.Picture)
						t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"pic", len(t.Pictures) - 1})
					}
				} else if ac.Fallback != nil {
					if ac.Fallback.GraphicFrame != nil {
						t.GraphicFrames = append(t.GraphicFrames, *ac.Fallback.GraphicFrame)
						t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"graphicFrame", len(t.GraphicFrames) - 1})
					} else if ac.Fallback.Shape != nil {
						t.Shapes = append(t.Shapes, *ac.Fallback.Shape)
						t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"sp", len(t.Shapes) - 1})
					} else if ac.Fallback.Picture != nil {
						t.Pictures = append(t.Pictures, *ac.Fallback.Picture)
						t.ChildOrder = append(t.ChildOrder, SpTreeChildRef{"pic", len(t.Pictures) - 1})
					}
				}
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// ======================== Relationships (.rels) ========================

// OOXMLRelationships represents the root <Relationships> element in .rels files.
type OOXMLRelationships struct {
	XMLName xml.Name            `xml:"Relationships"`
	Rels    []OOXMLRelationship `xml:"Relationship"`
}

// OOXMLRelationship represents a single <Relationship> element.
type OOXMLRelationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

// ======================== Notes Slide (p:notes) ========================

// OOXMLNotesSlide represents the root <p:notes> element in notesSlide*.xml.
type OOXMLNotesSlide struct {
	XMLName xml.Name       `xml:"notes"`
	CSld    OOXMLNotesCSld `xml:"cSld"`
}

// OOXMLNotesCSld represents the <p:cSld> inside a notes slide.
type OOXMLNotesCSld struct {
	SpTree OOXMLNotesSpTree `xml:"spTree"`
}

// OOXMLNotesSpTree represents the <p:spTree> inside a notes slide.
type OOXMLNotesSpTree struct {
	Shapes []OOXMLNotesShape `xml:"sp"`
}

// OOXMLNotesShape represents a <p:sp> in a notes slide.
type OOXMLNotesShape struct {
	NvSpPr OOXMLNvSpPr       `xml:"nvSpPr"`
	TxBody *OOXMLNotesTxBody `xml:"txBody"`
}

// OOXMLNotesTxBody represents the <p:txBody> of a notes shape.
type OOXMLNotesTxBody struct {
	Paragraphs []OOXMLNotesParagraph `xml:"p"`
}

// OOXMLNotesParagraph represents a single <a:p> in notes.
type OOXMLNotesParagraph struct {
	PPr  *OOXMLNotesPPr  `xml:"pPr"`
	Runs []OOXMLNotesRun `xml:"r"`
}

// OOXMLNotesPPr represents <a:pPr> in notes paragraphs.
type OOXMLNotesPPr struct {
	BuChar    *OOXMLBuChar    `xml:"buChar"`
	BuAutoNum *OOXMLBuAutoNum `xml:"buAutoNum"`
	BuNone    *struct{}       `xml:"buNone"`
}

// OOXMLBuChar represents <a:buChar> (character bullet).
type OOXMLBuChar struct {
	Char string `xml:"char,attr"`
}

// OOXMLBuAutoNum represents <a:buAutoNum> (auto-numbered bullet).
type OOXMLBuAutoNum struct {
	Type string `xml:"type,attr"`
}

// OOXMLNotesRun represents a single <a:r> text run in notes.
type OOXMLNotesRun struct {
	RPr  *OOXMLNotesRPr `xml:"rPr"`
	Text string         `xml:"t"`
}

// OOXMLNotesRPr represents <a:rPr> (run properties) in notes.
type OOXMLNotesRPr struct {
	Bold       string           `xml:"b,attr,omitempty"`
	Italic     string           `xml:"i,attr,omitempty"`
	Underline  string           `xml:"u,attr,omitempty"`
	HlinkClick *OOXMLHlinkClick `xml:"hlinkClick"`
}

// OOXMLHlinkClick represents <a:hlinkClick r:id="rId5"/> â€” a hyperlink on a shape or text run.
// RID uses the long namespace for slide shapes; RIDNS uses the short one for notes compatibility.
type OOXMLHlinkClick struct {
	RID    string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr,omitempty"`
	Action string `xml:"action,attr,omitempty"` // e.g. "ppaction://hlinksldjump"
}

// ======================== Shape (p:sp) ========================

// OOXMLShape represents a single shape <p:sp>.
type OOXMLShape struct {
	NvSpPr OOXMLNvSpPr `xml:"nvSpPr"`
	SpPr   OOXMLSpPr   `xml:"spPr"`
	Style  OOXMLStyle  `xml:"style"`
	TxBody OOXMLTxBody `xml:"txBody"`
	TxXfrm *OOXMLXfrm  `xml:"txXfrm"` // Optional text frame transform (used in SmartArt dsp:sp)
}

// OOXMLNvSpPr represents <p:nvSpPr> (non-visual shape properties).
type OOXMLNvSpPr struct {
	CNvPr   OOXMLCNvPr   `xml:"cNvPr"`
	CNvSpPr OOXMLCNvSpPr `xml:"cNvSpPr"`
	NvPr    OOXMLNvPr    `xml:"nvPr"`
}

// OOXMLCNvPr represents <p:cNvPr> with id and name attributes.
type OOXMLCNvPr struct {
	ID         int              `xml:"id,attr"`
	Name       string           `xml:"name,attr"`
	Descr      string           `xml:"descr,attr,omitempty"`
	HlinkClick *OOXMLHlinkClick `xml:"hlinkClick"`
}

// OOXMLCNvSpPr represents <p:cNvSpPr>.
type OOXMLCNvSpPr struct {
	TxBox string `xml:"txBox,attr,omitempty"`
}

// OOXMLNvPr represents <p:nvPr>, containing placeholder info.
type OOXMLNvPr struct {
	Ph *OOXMLPlaceholder `xml:"ph"`
}

// OOXMLPlaceholder represents <p:ph> with type and idx attributes.
type OOXMLPlaceholder struct {
	Type string `xml:"type,attr,omitempty"` // "title", "body", "subTitle", "ctrTitle"
	Idx  string `xml:"idx,attr,omitempty"`  // content placeholder index
}

// ======================== Shape Properties (p:spPr / a:spPr) ========================

// OOXMLSpPr represents <p:spPr> / <a:spPr> (visual shape properties).
type OOXMLSpPr struct {
	BWMode    string          `xml:"bwMode,attr,omitempty"`
	Xfrm      OOXMLXfrm       `xml:"xfrm"`
	PrstGeom  *OOXMLPrstGeom  `xml:"prstGeom"`
	CustGeom  *OOXMLCustGeom  `xml:"custGeom"`
	SolidFill *OOXMLSolidFill `xml:"solidFill"`
	GradFill  *OOXMLGradFill  `xml:"gradFill"`
	PattFill  *OOXMLPattFill  `xml:"pattFill"`
	BlipFill  *OOXMLBlipFill  `xml:"blipFill"`
	NoFill    *struct{}       `xml:"noFill"`
	Ln        *OOXMLLn        `xml:"ln"`
	EffectLst *OOXMLEffectLst `xml:"effectLst"`
}

// OOXMLEffectLst represents <a:effectLst>.
type OOXMLEffectLst struct {
	SoftEdge   *OOXMLSoftEdge    `xml:"softEdge"`
	OuterShdw  *OOXMLOuterShadow `xml:"outerShdw"`
	InnerShdw  *OOXMLInnerShadow `xml:"innerShdw"`
	Reflection *OOXMLReflection  `xml:"reflection"`
}

// OOXMLReflection represents <a:reflection> text/shape effect.
type OOXMLReflection struct {
	BlurRad  string `xml:"blurRad,attr,omitempty"`
	StA      string `xml:"stA,attr,omitempty"`  // start alpha (1/1000 of percent)
	EndA     string `xml:"endA,attr,omitempty"`  // end alpha
	EndPos   string `xml:"endPos,attr,omitempty"`
	Dir      string `xml:"dir,attr,omitempty"`
	Sy       string `xml:"sy,attr,omitempty"` // vertical scale (negative = flip)
	Algn     string `xml:"algn,attr,omitempty"`
	RotWithShape string `xml:"rotWithShape,attr,omitempty"`
}

// OOXMLInnerShadow represents <a:innerShdw> (inner/inset shadow).
type OOXMLInnerShadow struct {
	BlurRad   string          `xml:"blurRad,attr,omitempty"` // blur radius in EMU
	Dist      string          `xml:"dist,attr,omitempty"`    // distance in EMU
	Dir       string          `xml:"dir,attr,omitempty"`     // direction in 60000ths of a degree
	SrgbClr   *OOXMLSrgbClr   `xml:"srgbClr"`
	SchemeClr *OOXMLSchemeClr `xml:"schemeClr"`
	PrstClr   *OOXMLPrstClr   `xml:"prstClr"`
}

// OOXMLSoftEdge represents <a:softEdge rad="..."/>.
type OOXMLSoftEdge struct {
	Rad int `xml:"rad,attr"` // soft edge radius in EMU
}

// OOXMLOuterShadow represents <a:outerShdw> (outer shadow / perspective shadow).
type OOXMLOuterShadow struct {
	BlurRad      string          `xml:"blurRad,attr,omitempty"`      // blur radius in EMU
	Dist         string          `xml:"dist,attr,omitempty"`         // distance in EMU
	Dir          string          `xml:"dir,attr,omitempty"`          // direction in 60000ths of a degree
	Sx           int64           `xml:"sx,attr,omitempty"`           // horizontal scale (1000ths of percent, 100000 = 100%)
	Sy           int64           `xml:"sy,attr,omitempty"`           // vertical scale (negative = reflection)
	Kx           int64           `xml:"kx,attr,omitempty"`           // horizontal skew in 60000ths of degree
	Ky           int64           `xml:"ky,attr,omitempty"`           // vertical skew in 60000ths of degree
	Algn         string          `xml:"algn,attr,omitempty"`         // alignment anchor: tl, t, tr, l, ctr, r, bl, b, br
	RotWithShape int             `xml:"rotWithShape,attr,omitempty"` // 0 = don't rotate with shape
	SrgbClr      *OOXMLSrgbClr   `xml:"srgbClr"`
	SchemeClr    *OOXMLSchemeClr `xml:"schemeClr"`
	PrstClr      *OOXMLPrstClr   `xml:"prstClr"`
}

// ======================== Transform (a:xfrm) ========================

// OOXMLXfrm represents <a:xfrm> with position, size, and optional flip/rotation.
type OOXMLXfrm struct {
	Rot   int       `xml:"rot,attr,omitempty"`
	FlipH string    `xml:"flipH,attr,omitempty"` // "1" = horizontally flipped
	FlipV string    `xml:"flipV,attr,omitempty"` // "1" = vertically flipped
	Off   OOXMLOff  `xml:"off"`
	Ext   OOXMLExt  `xml:"ext"`
	ChOff *OOXMLOff `xml:"chOff"` // child coordinate offset (groups only)
	ChExt *OOXMLExt `xml:"chExt"` // child coordinate extent (groups only)
}

// OOXMLOff represents <a:off x="..." y="..."/> (offset in EMU).
type OOXMLOff struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// OOXMLExt represents <a:ext cx="..." cy="..."/> (extent in EMU).
type OOXMLExt struct {
	Cx int64 `xml:"cx,attr"`
	Cy int64 `xml:"cy,attr"`
}

// ======================== Geometry ========================

// OOXMLPrstGeom represents <a:prstGeom prst="..."> (preset geometry).
type OOXMLPrstGeom struct {
	Prst  string     `xml:"prst,attr"`
	AvLst OOXMLAvLst `xml:"avLst"`
}

// OOXMLAvLst represents <a:avLst> (adjustment value list).
type OOXMLAvLst struct {
	Gd []OOXMLGd `xml:"gd"`
}

// OOXMLGd represents <a:gd name="..." fmla="..."/> (guide/adjustment value).
type OOXMLGd struct {
	Name string `xml:"name,attr"`
	Fmla string `xml:"fmla,attr"`
}

// OOXMLCustGeom represents <a:custGeom> (custom/freeform geometry).
type OOXMLCustGeom struct {
	PathLst OOXMLPathLst `xml:"pathLst"`
}

// OOXMLPathLst represents <a:pathLst>.
type OOXMLPathLst struct {
	Paths []OOXMLPathRaw `xml:"path"`
}

// OOXMLPathRaw represents <a:path w="..." h="..."> with raw inner XML.
// We use innerxml to preserve the order of moveTo/lineTo/cubicBezTo/close commands.
type OOXMLPathRaw struct {
	W     int64  `xml:"w,attr,omitempty"`
	H     int64  `xml:"h,attr,omitempty"`
	Inner string `xml:",innerxml"`
}

// OOXMLPt represents <a:pt x="..." y="..."/>.
type OOXMLPt struct {
	X int64 `xml:"x,attr"`
	Y int64 `xml:"y,attr"`
}

// ======================== Fills ========================

// OOXMLSolidFill represents <a:solidFill>.
type OOXMLSolidFill struct {
	SrgbClr   *OOXMLSrgbClr   `xml:"srgbClr"`
	SchemeClr *OOXMLSchemeClr `xml:"schemeClr"`
	PrstClr   *OOXMLPrstClr   `xml:"prstClr"`
}

// OOXMLSrgbClr represents <a:srgbClr val="RRGGBB">.
type OOXMLSrgbClr struct {
	Val    string       `xml:"val,attr"`
	Shade  *OOXMLIntVal `xml:"shade"`
	Tint   *OOXMLIntVal `xml:"tint"`
	HueMod *OOXMLIntVal `xml:"hueMod"`
	HueOff *OOXMLIntVal `xml:"hueOff"`
	SatMod *OOXMLIntVal `xml:"satMod"`
	SatOff *OOXMLIntVal `xml:"satOff"`
	LumMod *OOXMLIntVal `xml:"lumMod"`
	LumOff *OOXMLIntVal `xml:"lumOff"`
	Alpha  *OOXMLIntVal `xml:"alpha"`
}

// OOXMLSchemeClr represents <a:schemeClr val="accent1"> with optional modifiers.
type OOXMLSchemeClr struct {
	Val    string       `xml:"val,attr"` // "accent1"â€“"accent6", "tx1", "lt1", "dk1", "bg1"
	Shade  *OOXMLIntVal `xml:"shade"`
	Tint   *OOXMLIntVal `xml:"tint"`
	HueMod *OOXMLIntVal `xml:"hueMod"`
	HueOff *OOXMLIntVal `xml:"hueOff"`
	SatMod *OOXMLIntVal `xml:"satMod"`
	SatOff *OOXMLIntVal `xml:"satOff"`
	LumMod *OOXMLIntVal `xml:"lumMod"`
	LumOff *OOXMLIntVal `xml:"lumOff"`
	Alpha  *OOXMLIntVal `xml:"alpha"`
}

// OOXMLPrstClr represents <a:prstClr val="black">.
type OOXMLPrstClr struct {
	Val    string       `xml:"val,attr"`
	LumMod *OOXMLIntVal `xml:"lumMod"`
	LumOff *OOXMLIntVal `xml:"lumOff"`
	Alpha  *OOXMLIntVal `xml:"alpha"`
}

// OOXMLIntVal represents a child element with a single int val attribute (shade, lumMod, etc.).
type OOXMLIntVal struct {
	Val int `xml:"val,attr"`
}

// OOXMLGradFill represents <a:gradFill>.
type OOXMLGradFill struct {
	GsLst OOXMLGsLst     `xml:"gsLst"`
	Lin   *OOXMLLin      `xml:"lin"`
	Path  *OOXMLGradPath `xml:"path"` // radial gradient (path="circle", "rect", etc.)
}

// OOXMLGradPath represents <a:path path="circle"> for radial gradients.
type OOXMLGradPath struct {
	PathType   string           `xml:"path,attr"` // "circle", "rect", "shape"
	FillToRect *OOXMLFillToRect `xml:"fillToRect"`
}

// OOXMLFillToRect represents <a:fillToRect l="..." t="..." r="..." b="..."/>.
type OOXMLFillToRect struct {
	L int `xml:"l,attr,omitempty"` // left offset (in 1/100000)
	T int `xml:"t,attr,omitempty"` // top offset
	R int `xml:"r,attr,omitempty"` // right offset
	B int `xml:"b,attr,omitempty"` // bottom offset
}

// OOXMLGsLst represents <a:gsLst> (gradient stop list).
type OOXMLGsLst struct {
	Gs []OOXMLGs `xml:"gs"`
}

// OOXMLGs represents a single gradient stop <a:gs pos="...">.
type OOXMLGs struct {
	Pos       int             `xml:"pos,attr"`
	SrgbClr   *OOXMLSrgbClr   `xml:"srgbClr"`
	SchemeClr *OOXMLSchemeClr `xml:"schemeClr"`
}

// OOXMLLin represents <a:lin ang="..." scaled="..."/>.
type OOXMLLin struct {
	Ang    int    `xml:"ang,attr"`
	Scaled string `xml:"scaled,attr,omitempty"`
}

// ======================== Line Properties (a:ln) ========================

// OOXMLLn represents <a:ln w="...">.
type OOXMLLn struct {
	W         int             `xml:"w,attr,omitempty"`    // line width in EMU
	Cap       string          `xml:"cap,attr,omitempty"`  // line cap: "rnd","sq","flat" or empty
	Cmpd      string          `xml:"cmpd,attr,omitempty"` // compound type: "sng","dbl","thickThin","thinThick","tri"
	SolidFill *OOXMLSolidFill `xml:"solidFill"`
	GradFill  *OOXMLGradFill  `xml:"gradFill"` // gradient fill on line stroke
	NoFill    *struct{}       `xml:"noFill"`   // <a:noFill/> â€” suppresses stroke
	PrstDash  *OOXMLPrstDash  `xml:"prstDash"`
	TailEnd   *OOXMLLineEnd   `xml:"tailEnd"`
	HeadEnd   *OOXMLLineEnd   `xml:"headEnd"`
}

// OOXMLPrstDash represents <a:prstDash val="dash"/>.
type OOXMLPrstDash struct {
	Val string `xml:"val,attr"`
}

// OOXMLLineEnd represents <a:tailEnd type="triangle"/> or <a:headEnd ...>.
type OOXMLLineEnd struct {
	Type string `xml:"type,attr"`
	W    string `xml:"w,attr,omitempty"`
	Len  string `xml:"len,attr,omitempty"`
}

// ======================== Shape Style (p:style) ========================

// OOXMLStyle represents <p:style> with theme references.
type OOXMLStyle struct {
	LnRef     OOXMLStyleRef `xml:"lnRef"`
	FillRef   OOXMLStyleRef `xml:"fillRef"`
	EffectRef OOXMLStyleRef `xml:"effectRef"`
	FontRef   OOXMLStyleRef `xml:"fontRef"`
}

// OOXMLStyleRef represents a style reference (lnRef/fillRef/effectRef/fontRef).
// Idx is string because fontRef uses "minor"/"major" while others use numeric values.
type OOXMLStyleRef struct {
	Idx       string          `xml:"idx,attr"`
	SchemeClr *OOXMLSchemeClr `xml:"schemeClr"`
}

// ======================== Text Body (p:txBody) ========================

// OOXMLTxBody represents <p:txBody>.
type OOXMLTxBody struct {
	BodyPr     OOXMLBodyPr         `xml:"bodyPr"`
	LstStyle   *OOXMLTextListStyle `xml:"lstStyle"`
	Paragraphs []OOXMLParagraph    `xml:"p"`
}

// OOXMLBodyPr represents <a:bodyPr>.
type OOXMLBodyPr struct {
	Anchor      string            `xml:"anchor,attr,omitempty"`    // "t", "ctr", "b"
	AnchorCtr   string            `xml:"anchorCtr,attr,omitempty"` // "1" or "0"
	Vert        string            `xml:"vert,attr,omitempty"`      // "horz", "eaVert", "vert", "vert270"
	Wrap        string            `xml:"wrap,attr,omitempty"`
	RtlCol      string            `xml:"rtlCol,attr,omitempty"`
	LIns        *int64            `xml:"lIns,attr,omitempty"` // left inset in EMU (default 91440)
	RIns        *int64            `xml:"rIns,attr,omitempty"` // right inset in EMU (default 91440)
	TIns        *int64            `xml:"tIns,attr,omitempty"` // top inset in EMU (default 45720)
	BIns        *int64            `xml:"bIns,attr,omitempty"` // bottom inset in EMU (default 45720)
	NoAutofit   *struct{}         `xml:"noAutofit"`           // Do not Auto Fit
	NormAutofit *OOXMLNormAutofit `xml:"normAutofit"`         // Shrink Text on Overflow
	SpAutoFit   *struct{}         `xml:"spAutoFit"`           // Resize Shape to Fit Text
}

// OOXMLNormAutofit represents <a:normAutofit> (Shrink Text on Overflow).
// Attributes: fontScale and lnSpcReduction as percent strings (e.g. "92.000%").
type OOXMLNormAutofit struct {
	FontScale      string `xml:"fontScale,attr,omitempty"`
	LnSpcReduction string `xml:"lnSpcReduction,attr,omitempty"`
}

// OOXMLParaItem represents one item in paragraph content order: either a run or a line break.
// Used when custom-unmarshaling <a:p> to preserve <a:r> and <a:br/> order for multi-line titles.
type OOXMLParaItem struct {
	Run *OOXMLRun
	Br  bool
}

// OOXMLParagraph represents <a:p>.
// When unmarshaled with custom UnmarshalXML, Content preserves <a:r> and <a:br/> order; Runs is then filled from Content.
type OOXMLParagraph struct {
	PPr        *OOXMLPPr       `xml:"pPr"`
	Runs       []OOXMLRun      `xml:"r"`
	Content    []OOXMLParaItem `xml:"-"` // filled by UnmarshalXML when <a:br/> present
	EndParaRPr *OOXMLRPr       `xml:"endParaRPr"`
}

// UnmarshalXML for OOXMLParagraph preserves order of <a:r> and <a:br/> so line breaks produce separate lines.
func (p *OOXMLParagraph) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	var runs []OOXMLRun
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "pPr":
				p.PPr = &OOXMLPPr{}
				if err := d.DecodeElement(p.PPr, &t); err != nil {
					return err
				}
			case "r":
				var r OOXMLRun
				if err := d.DecodeElement(&r, &t); err != nil {
					return err
				}
				runs = append(runs, r)
				p.Content = append(p.Content, OOXMLParaItem{Run: &runs[len(runs)-1]})
			case "br":
				if err := d.Skip(); err != nil {
					return err
				}
				p.Content = append(p.Content, OOXMLParaItem{Br: true})
			case "endParaRPr":
				p.EndParaRPr = &OOXMLRPr{}
				if err := d.DecodeElement(p.EndParaRPr, &t); err != nil {
					return err
				}
			case "fld":
				var fld struct {
					Runs []OOXMLRun `xml:"r"`
				}
				if err := d.DecodeElement(&fld, &t); err != nil {
					return err
				}
				for i := range fld.Runs {
					runs = append(runs, fld.Runs[i])
					p.Content = append(p.Content, OOXMLParaItem{Run: &runs[len(runs)-1]})
				}
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			p.Runs = runs
			return nil
		}
	}
}

// OOXMLPPr represents <a:pPr> (paragraph properties).
type OOXMLPPr struct {
	Algn      string          `xml:"algn,attr,omitempty"` // "l", "ctr", "r"
	Lvl       int             `xml:"lvl,attr,omitempty"`
	MarL      int64           `xml:"marL,attr,omitempty"`   // left margin in EMU
	Indent    int64           `xml:"indent,attr,omitempty"` // first line indent in EMU
	DefRPr    *OOXMLRPr       `xml:"defRPr"`                // default run properties for paragraph
	LnSpc     *OOXMLSpacing   `xml:"lnSpc"`                 // line spacing
	SpcBef    *OOXMLSpacing   `xml:"spcBef"`                // space before paragraph
	SpcAft    *OOXMLSpacing   `xml:"spcAft"`                // space after paragraph
	BuFont    *OOXMLFont      `xml:"buFont"`                // bullet font typeface
	BuChar    *OOXMLBuChar    `xml:"buChar"`
	BuAutoNum *OOXMLBuAutoNum `xml:"buAutoNum"`
	BuNone    *struct{}       `xml:"buNone"` // suppress bullets
}

// OOXMLSpacing represents spacing elements (lnSpc, spcBef, spcAft).
type OOXMLSpacing struct {
	SpcPct *OOXMLIntVal `xml:"spcPct"` // percentage (e.g., 150000 = 150%)
	SpcPts *OOXMLIntVal `xml:"spcPts"` // hundredths of a point (e.g., 1200 = 12pt)
}

// OOXMLRun represents <a:r> (text run).
type OOXMLRun struct {
	RPr OOXMLRPr `xml:"rPr"`
	T   string   `xml:"t"`
}

// OOXMLRPr represents <a:rPr> (run properties).
type OOXMLRPr struct {
	Lang       string           `xml:"lang,attr,omitempty"`
	Sz         int              `xml:"sz,attr,omitempty"`     // hundredths of a point (1400 = 14pt)
	B          string           `xml:"b,attr,omitempty"`      // "1" = bold
	I          string           `xml:"i,attr,omitempty"`      // "1" = italic
	U          string           `xml:"u,attr,omitempty"`      // "sng","dbl","heavy","dotted","dottedHeavy","dash","dashHeavy","dashLong","dashLongHeavy","dotDash","dotDashHeavy","dotDotDash","dotDotDashHeavy","wavy","wavyHeavy","wavyDbl","none"
	Strike     string           `xml:"strike,attr,omitempty"` // "sngStrike","dblStrike","noStrike"
	Baseline   int              `xml:"baseline,attr,omitempty"`
	Cap        string           `xml:"cap,attr,omitempty"`    // "all","small","none"
	Spc        int              `xml:"spc,attr,omitempty"`    // letter-spacing in hundredths of a point
	SolidFill  *OOXMLSolidFill  `xml:"solidFill"`
	GradFill   *OOXMLGradFill   `xml:"gradFill"`
	PattFill   *OOXMLPattFill   `xml:"pattFill"`
	Ln         *OOXMLLn         `xml:"ln"`
	EffectLst  *OOXMLEffectLst  `xml:"effectLst"`
	Latin      *OOXMLFont       `xml:"latin"`
	Cs         *OOXMLFont       `xml:"cs"`
	HlinkClick *OOXMLHlinkClick `xml:"hlinkClick"`
}

// OOXMLFont represents <a:latin typeface="..."/> or <a:cs typeface="..."/>.
type OOXMLFont struct {
	Typeface string `xml:"typeface,attr"`
}

// ======================== Group Shape (p:grpSp) ========================

// OOXMLGroupShape represents <p:grpSp>.
// ChildOrder, when non-nil, gives document order of children for correct z-order.
type OOXMLGroupShape struct {
	NvGrpSpPr  OOXMLNvGrpSpPr        `xml:"nvGrpSpPr"`
	GrpSpPr    OOXMLGrpSpPr          `xml:"grpSpPr"`
	Shapes     []OOXMLShape          `xml:"sp"`
	Groups     []OOXMLGroupShape     `xml:"grpSp"` // nested groups
	Connectors []OOXMLConnectorShape `xml:"cxnSp"`
	Pictures   []OOXMLPictureShape   `xml:"pic"`
	ChildOrder []SpTreeChildRef      `xml:"-"` // filled by UnmarshalXML for z-order
}

// UnmarshalXML preserves document order of group children for correct z-order (ECMA-376 19.3.1.45).
func (g *OOXMLGroupShape) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch se := tok.(type) {
		case xml.StartElement:
			switch se.Name.Local {
			case "nvGrpSpPr":
				if err := d.DecodeElement(&g.NvGrpSpPr, &se); err != nil {
					return err
				}
			case "grpSpPr":
				if err := d.DecodeElement(&g.GrpSpPr, &se); err != nil {
					return err
				}
			case "sp":
				var s OOXMLShape
				if err := d.DecodeElement(&s, &se); err != nil {
					return err
				}
				g.Shapes = append(g.Shapes, s)
				g.ChildOrder = append(g.ChildOrder, SpTreeChildRef{"sp", len(g.Shapes) - 1})
			case "grpSp":
				var nested OOXMLGroupShape
				if err := d.DecodeElement(&nested, &se); err != nil {
					return err
				}
				g.Groups = append(g.Groups, nested)
				g.ChildOrder = append(g.ChildOrder, SpTreeChildRef{"grpSp", len(g.Groups) - 1})
			case "cxnSp":
				var c OOXMLConnectorShape
				if err := d.DecodeElement(&c, &se); err != nil {
					return err
				}
				g.Connectors = append(g.Connectors, c)
				g.ChildOrder = append(g.ChildOrder, SpTreeChildRef{"cxnSp", len(g.Connectors) - 1})
			case "pic":
				var p OOXMLPictureShape
				if err := d.DecodeElement(&p, &se); err != nil {
					return err
				}
				g.Pictures = append(g.Pictures, p)
				g.ChildOrder = append(g.ChildOrder, SpTreeChildRef{"pic", len(g.Pictures) - 1})
			default:
				if err := d.Skip(); err != nil {
					return err
				}
			}
		case xml.EndElement:
			return nil
		}
	}
}

// OOXMLNvGrpSpPr represents <p:nvGrpSpPr>.
type OOXMLNvGrpSpPr struct {
	CNvPr OOXMLCNvPr `xml:"cNvPr"`
}

// OOXMLGrpSpPr represents <p:grpSpPr>.
type OOXMLGrpSpPr struct {
	Xfrm OOXMLXfrm `xml:"xfrm"`
}

// ======================== Connector Shape (p:cxnSp) ========================

// OOXMLConnectorShape represents <p:cxnSp>.
type OOXMLConnectorShape struct {
	NvCxnSpPr OOXMLNvCxnSpPr `xml:"nvCxnSpPr"`
	SpPr      OOXMLSpPr      `xml:"spPr"`
	Style     OOXMLStyle     `xml:"style"`
}

// OOXMLNvCxnSpPr represents <p:nvCxnSpPr>.
type OOXMLNvCxnSpPr struct {
	CNvPr      OOXMLCNvPr      `xml:"cNvPr"`
	CNvCxnSpPr OOXMLCNvCxnSpPr `xml:"cNvCxnSpPr"`
}

// OOXMLCNvCxnSpPr represents <p:cNvCxnSpPr> with start/end connection references.
type OOXMLCNvCxnSpPr struct {
	StCxn  *OOXMLCxnRef `xml:"stCxn"`
	EndCxn *OOXMLCxnRef `xml:"endCxn"`
}

// OOXMLCxnRef represents <a:stCxn id="..." idx="..."/> or <a:endCxn ...>.
type OOXMLCxnRef struct {
	ID  int `xml:"id,attr"`
	Idx int `xml:"idx,attr"`
}

// ======================== Picture Shape (p:pic) ========================

// OOXMLPictureShape represents <p:pic>.
type OOXMLPictureShape struct {
	NvPicPr  OOXMLNvPicPr  `xml:"nvPicPr"`
	BlipFill OOXMLBlipFill `xml:"blipFill"`
	SpPr     OOXMLSpPr     `xml:"spPr"`
	Style    OOXMLStyle    `xml:"style"`
}

// OOXMLNvPicPr represents <p:nvPicPr>.
type OOXMLNvPicPr struct {
	CNvPr OOXMLCNvPr `xml:"cNvPr"`
}

// OOXMLBlipFill represents <p:blipFill> or <a:blipFill>.
type OOXMLBlipFill struct {
	Blip    OOXMLBlip         `xml:"blip"`
	Tile    *OOXMLTileAttr    `xml:"tile"`
	Stretch *OOXMLStretchAttr `xml:"stretch"`
}

type OOXMLTileAttr struct {
	Tx   string `xml:"tx,attr,omitempty"` // X offset
	Ty   string `xml:"ty,attr,omitempty"` // Y offset
	Sx   string `xml:"sx,attr,omitempty"` // Scale X (percentage per 100,000)
	Sy   string `xml:"sy,attr,omitempty"` // Scale Y
	Flip string `xml:"flip,attr,omitempty"`
	Algn string `xml:"algn,attr,omitempty"`
}

type OOXMLStretchAttr struct {
	FillRect *OOXMLFillRect `xml:"fillRect"`
}

type OOXMLFillRect struct {
	L string `xml:"l,attr,omitempty"`
	T string `xml:"t,attr,omitempty"`
	R string `xml:"r,attr,omitempty"`
	B string `xml:"b,attr,omitempty"`
}

// OOXMLBlip represents <a:blip r:embed="rId...">.
type OOXMLBlip struct {
	Embed  string           `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr,omitempty"`
	ExtLst *OOXMLBlipExtLst `xml:"extLst"`
}

// OOXMLBlipExtLst represents <a:extLst> inside <a:blip>.
type OOXMLBlipExtLst struct {
	Exts []OOXMLBlipExt `xml:"ext"`
}

// OOXMLBlipExt represents <a:ext uri="..."> inside blip extLst.
type OOXMLBlipExt struct {
	URI     string        `xml:"uri,attr"`
	SvgBlip *OOXMLSvgBlip `xml:"svgBlip"`
}

// OOXMLSvgBlip represents <asvg:svgBlip r:embed="rId...">.
type OOXMLSvgBlip struct {
	Embed string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships embed,attr,omitempty"`
}

// ======================== Theme (a:theme) ========================

// OOXMLTheme represents <a:theme> for parsing font and color schemes.
type OOXMLTheme struct {
	XMLName       xml.Name           `xml:"theme"`
	ThemeElements OOXMLThemeElements `xml:"themeElements"`
}

// OOXMLThemeElements wraps the main theme element sections.
type OOXMLThemeElements struct {
	ClrScheme  OOXMLClrScheme  `xml:"clrScheme"`
	FontScheme OOXMLFontScheme `xml:"fontScheme"`
	FmtScheme  OOXMLFmtScheme  `xml:"fmtScheme"`
}

// OOXMLFmtScheme represents <a:fmtScheme> containing fill, line, and effect style lists.
type OOXMLFmtScheme struct {
	Name           string                   `xml:"name,attr"`
	FillStyleLst   OOXMLThemeFillStyleLst   `xml:"fillStyleLst"`
	EffectStyleLst OOXMLThemeEffectStyleLst `xml:"effectStyleLst"`
}

// OOXMLThemeFillStyleLst is an ordered list of fill style entries (solidFill/gradFill).
type OOXMLThemeFillStyleLst struct {
	Entries []OOXMLThemeFillEntry
}

// UnmarshalXML preserves the ordering of solidFill/gradFill children in fillStyleLst.
func (f *OOXMLThemeFillStyleLst) UnmarshalXML(d *xml.Decoder, start xml.StartElement) error {
	for {
		tok, err := d.Token()
		if err != nil {
			return err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			entry := OOXMLThemeFillEntry{}
			switch t.Name.Local {
			case "solidFill":
				var sf OOXMLThemeSolidFill
				if err := d.DecodeElement(&sf, &t); err != nil {
					return err
				}
				entry.SolidFill = &sf
			case "gradFill":
				var gf OOXMLThemeGradFill
				if err := d.DecodeElement(&gf, &t); err != nil {
					return err
				}
				entry.GradFill = &gf
			default:
				if err := d.Skip(); err != nil {
					return err
				}
				continue
			}
			f.Entries = append(f.Entries, entry)
		case xml.EndElement:
			return nil
		}
	}
}

// OOXMLThemeFillEntry represents one fill entry: either solidFill or gradFill.
type OOXMLThemeFillEntry struct {
	SolidFill *OOXMLThemeSolidFill
	GradFill  *OOXMLThemeGradFill
}

// OOXMLThemeSolidFill represents <a:solidFill> in the theme fillStyleLst.
type OOXMLThemeSolidFill struct {
	SchemeClr *OOXMLSchemeClr `xml:"schemeClr"`
	SrgbClr   *OOXMLSrgbClr   `xml:"srgbClr"`
}

// OOXMLThemeGradFill represents <a:gradFill> in the theme fillStyleLst.
type OOXMLThemeGradFill struct {
	RotWithShape string               `xml:"rotWithShape,attr,omitempty"`
	GsLst        []OOXMLThemeGradStop `xml:"gsLst>gs"`
	Lin          *OOXMLLin            `xml:"lin"`
	Path         *OOXMLThemeGradPath  `xml:"path"`
}

// OOXMLThemeGradPath represents <a:path> for a gradient fill shape.
type OOXMLThemeGradPath struct {
	Path string `xml:"path,attr,omitempty"`
}

// OOXMLThemeGradStop represents <a:gs pos="35000"> gradient stop.
type OOXMLThemeGradStop struct {
	Pos       int             `xml:"pos,attr"`
	SchemeClr *OOXMLSchemeClr `xml:"schemeClr"`
	SrgbClr   *OOXMLSrgbClr   `xml:"srgbClr"`
}

// OOXMLThemeEffectStyleLst represents <a:effectStyleLst>.
type OOXMLThemeEffectStyleLst struct {
	EffectStyles []OOXMLThemeEffectStyle `xml:"effectStyle"`
}

// OOXMLThemeEffectStyle represents <a:effectStyle>.
type OOXMLThemeEffectStyle struct {
	EffectLst *OOXMLEffectLst `xml:"effectLst"`
}

// OOXMLClrScheme represents <a:clrScheme> defining the presentation's color theme.
type OOXMLClrScheme struct {
	Name     string          `xml:"name,attr"`
	Dk1      OOXMLThemeColor `xml:"dk1"`
	Lt1      OOXMLThemeColor `xml:"lt1"`
	Dk2      OOXMLThemeColor `xml:"dk2"`
	Lt2      OOXMLThemeColor `xml:"lt2"`
	Accent1  OOXMLThemeColor `xml:"accent1"`
	Accent2  OOXMLThemeColor `xml:"accent2"`
	Accent3  OOXMLThemeColor `xml:"accent3"`
	Accent4  OOXMLThemeColor `xml:"accent4"`
	Accent5  OOXMLThemeColor `xml:"accent5"`
	Accent6  OOXMLThemeColor `xml:"accent6"`
	Hlink    OOXMLThemeColor `xml:"hlink"`
	FolHlink OOXMLThemeColor `xml:"folHlink"`
}

// OOXMLThemeColor represents a single color entry in the theme (may be srgbClr or sysClr).
type OOXMLThemeColor struct {
	SrgbClr *OOXMLThemeSrgbClr `xml:"srgbClr"`
	SysClr  *OOXMLThemeSysClr  `xml:"sysClr"`
}

// Hex returns the resolved hex color (e.g. "#2781CB") for a theme color entry.
func (tc OOXMLThemeColor) Hex() string {
	if tc.SrgbClr != nil && tc.SrgbClr.Val != "" {
		return "#" + strings.ToUpper(tc.SrgbClr.Val)
	}
	if tc.SysClr != nil && tc.SysClr.LastClr != "" {
		return "#" + strings.ToUpper(tc.SysClr.LastClr)
	}
	return ""
}

// OOXMLThemeSrgbClr represents <a:srgbClr val="2781CB"/> in theme color definitions.
type OOXMLThemeSrgbClr struct {
	Val string `xml:"val,attr"`
}

// OOXMLThemeSysClr represents <a:sysClr val="windowText" lastClr="000000"/> in theme color definitions.
type OOXMLThemeSysClr struct {
	Val     string `xml:"val,attr"`
	LastClr string `xml:"lastClr,attr"`
}

// OOXMLFontScheme represents <a:fontScheme>.
type OOXMLFontScheme struct {
	Name      string              `xml:"name,attr"`
	MajorFont OOXMLFontCollection `xml:"majorFont"`
	MinorFont OOXMLFontCollection `xml:"minorFont"`
}

// OOXMLFontCollection represents <a:majorFont> or <a:minorFont>.
type OOXMLFontCollection struct {
	Latin OOXMLFont `xml:"latin"`
}

// ======================== Graphic Frame (p:graphicFrame) ========================

// OOXMLGraphicFrame represents <p:graphicFrame>, used for charts and tables.
type OOXMLGraphicFrame struct {
	NvGraphicFramePr OOXMLNvGraphicFramePr `xml:"nvGraphicFramePr"`
	Xfrm             OOXMLXfrm             `xml:"xfrm"`
	Graphic          OOXMLGraphic          `xml:"graphic"`
}

// OOXMLNvGraphicFramePr represents <p:nvGraphicFramePr>.
type OOXMLNvGraphicFramePr struct {
	CNvPr OOXMLCNvPr `xml:"cNvPr"`
}

// OOXMLGraphic represents <a:graphic>.
type OOXMLGraphic struct {
	GraphicData OOXMLGraphicData `xml:"graphicData"`
}

// OOXMLGraphicData represents <a:graphicData uri="...">.
type OOXMLGraphicData struct {
	URI   string         `xml:"uri,attr"`
	Chart *OOXMLChartRef `xml:"chart"` // <c:chart r:id="rId2"/>
	Tbl   *OOXMLTable    `xml:"tbl"`   // <a:tbl> (inline table)
}

// OOXMLChartRef represents <c:chart r:id="...">.
type OOXMLChartRef struct {
	RID string `xml:"id,attr"` // relationship ID
}

// ======================== Chart XML (c:chartSpace) ========================

// ChartSpace is the root element of a chart XML file.
type ChartSpace struct {
	XMLName xml.Name        `xml:"chartSpace"`
	Chart   ChartRoot       `xml:"chart"`
	SpPr    *ChartSpaceSpPr `xml:"spPr"` // chart-level formatting (background)
}

// ChartSpaceSpPr represents <c:spPr> on the chartSpace (chart background).
type ChartSpaceSpPr struct {
	SolidFill *OOXMLSolidFill     `xml:"solidFill"`
	GradFill  *OOXMLThemeGradFill `xml:"gradFill"`
	NoFill    *struct{}           `xml:"noFill"`
}

// ChartRoot represents <c:chart>.
type ChartRoot struct {
	Title    *ChartTitle   `xml:"title"`
	PlotArea ChartPlotArea `xml:"plotArea"`
	Legend   *ChartLegend  `xml:"legend"`
}

// ChartTitle represents <c:title>.
type ChartTitle struct {
	Tx   *ChartTitleTx `xml:"tx"`
	TxPr *ChartTxPr    `xml:"txPr"`
}

// ChartTitleTx represents <c:tx> inside chart title (explicit title text).
type ChartTitleTx struct {
	Rich *ChartTitleRich `xml:"rich"`
}

// ChartTitleRich represents <a:rich> body with paragraphs for chart title.
type ChartTitleRich struct {
	Paragraphs []ChartTitleParagraph `xml:"p"`
}

// ChartTitleParagraph is a paragraph in a rich chart title.
type ChartTitleParagraph struct {
	PPr  *ChartTitlePPr  `xml:"pPr"`
	Runs []ChartTitleRun `xml:"r"`
}

// ChartTitlePPr represents <a:pPr> inside a chart title paragraph.
type ChartTitlePPr struct {
	DefRPr *ChartTitleDefRPr `xml:"defRPr"`
}

// ChartTitleDefRPr represents <a:defRPr> with cap attribute for chart titles.
type ChartTitleDefRPr struct {
	Cap string `xml:"cap,attr"` // "all", "small", "none", or empty
}

// ChartTitleRun is a text run in a chart title paragraph.
type ChartTitleRun struct {
	T string `xml:"t"`
}

// ChartTxPr represents <c:txPr> for chart text properties.
type ChartTxPr struct {
	Paragraphs []ChartTxPrParagraph `xml:"p"`
}

// ChartTxPrParagraph represents <a:p> inside chart txPr.
type ChartTxPrParagraph struct {
	PPr *ChartTxPrPPr `xml:"pPr"`
}

// ChartTxPrPPr represents <a:pPr> with defRPr for chart text.
type ChartTxPrPPr struct {
	DefRPr *ChartDefRPr `xml:"defRPr"`
}

// ChartDefRPr represents <a:defRPr sz="1862">.
type ChartDefRPr struct {
	Sz int `xml:"sz,attr"` // font size in hundredths of a point
}

// ChartPlotArea represents <c:plotArea>.
type ChartPlotArea struct {
	BarCharts       []ChartBarChart      `xml:"barChart"`
	LineCharts      []ChartLineChart     `xml:"lineChart"`
	PieCharts       []ChartPieChart      `xml:"pieChart"`
	DoughnutCharts  []ChartDoughnutChart `xml:"doughnutChart"`
	AreaCharts      []ChartAreaChart     `xml:"areaChart"`
	ScatterCharts   []ChartScatterChart  `xml:"scatterChart"`
	StockCharts     []ChartStockChart    `xml:"stockChart"`
	SurfaceCharts   []ChartSurfaceChart  `xml:"surfaceChart"`
	Surface3DCharts []ChartSurfaceChart  `xml:"surface3DChart"`
	RadarCharts     []ChartRadarChart    `xml:"radarChart"`
	CatAx           *ChartCatAx          `xml:"catAx"`
	ValAx           *ChartValAx          `xml:"valAx"`
}

// ChartLineChart represents <c:lineChart>.
type ChartLineChart struct {
	Grouping ChartValAttr  `xml:"grouping"` // "standard", "stacked", "percentStacked"
	Series   []ChartSeries `xml:"ser"`
}

// ChartPieChart represents <c:pieChart>.
type ChartPieChart struct {
	VaryColors    ChartValAttr  `xml:"varyColors"`
	Series        []ChartSeries `xml:"ser"`
	DLbls         *ChartDLbls   `xml:"dLbls"` // chart-level data labels
	FirstSliceAng ChartValAttr  `xml:"firstSliceAng"`
}

// ChartDoughnutChart represents <c:doughnutChart>.
type ChartDoughnutChart struct {
	VaryColors    ChartValAttr  `xml:"varyColors"`
	Series        []ChartSeries `xml:"ser"`
	DLbls         *ChartDLbls   `xml:"dLbls"`         // chart-level data labels
	FirstSliceAng ChartValAttr  `xml:"firstSliceAng"` // starting angle in degrees
	HoleSize      ChartValAttr  `xml:"holeSize"`      // hole size percentage (e.g., "50")
}

// ChartBarChart represents <c:barChart>.
type ChartBarChart struct {
	BarDir   ChartValAttr  `xml:"barDir"`
	Grouping ChartValAttr  `xml:"grouping"`
	Series   []ChartSeries `xml:"ser"`
	GapWidth ChartValAttr  `xml:"gapWidth"`
	Overlap  ChartValAttr  `xml:"overlap"`
}

// ChartAreaChart represents <c:areaChart>.
type ChartAreaChart struct {
	Grouping ChartValAttr  `xml:"grouping"`
	Series   []ChartSeries `xml:"ser"`
}

// ChartScatterChart represents <c:scatterChart>.
type ChartScatterChart struct {
	ScatterStyle ChartValAttr  `xml:"scatterStyle"` // "none", "line", "lineMarker", "marker", "smooth", "smoothMarker"
	Series       []ChartSeries `xml:"ser"`
}

// ChartRadarChart represents <c:radarChart>.
type ChartRadarChart struct {
	RadarStyle ChartValAttr  `xml:"radarStyle"` // "standard", "marker", "filled"
	Series     []ChartSeries `xml:"ser"`
}

// ChartStockChart represents <c:stockChart>.
type ChartStockChart struct {
	Series []ChartSeries `xml:"ser"`
}

// ChartSurfaceChart represents <c:surfaceChart>.
type ChartSurfaceChart struct {
	Wireframe ChartValAttr  `xml:"wireframe"`
	Series    []ChartSeries `xml:"ser"`
}

// ChartValAttr is a simple element with a val attribute.
type ChartValAttr struct {
	Val string `xml:"val,attr"`
}

// ChartSeries represents <c:ser>.
type ChartSeries struct {
	Idx       ChartValAttr     `xml:"idx"`
	Order     ChartValAttr     `xml:"order"`
	Tx        ChartSeriesTx    `xml:"tx"`
	SpPr      *OOXMLSpPr       `xml:"spPr"`
	DPts      []ChartDataPoint `xml:"dPt"` // per-data-point formatting overrides
	Cat       *ChartCatData    `xml:"cat"`
	XVal      *ChartCatData    `xml:"xVal"` // used in scatter chart instead of cat
	Val       ChartValData     `xml:"val"`
	YVal      *ChartValData    `xml:"yVal"`      // used in scatter chart instead of val
	DLbls     *ChartDLbls      `xml:"dLbls"`     // data labels (used for pie charts)
	Explosion ChartValAttr     `xml:"explosion"` // pie explosion distance (% of radius)
}

// ChartDataPoint represents <c:dPt> - individual data point formatting.
type ChartDataPoint struct {
	Idx  ChartValAttr  `xml:"idx"`
	SpPr *ChartDPtSpPr `xml:"spPr"`
}

// ChartDPtSpPr represents <c:spPr> on a data point.
type ChartDPtSpPr struct {
	SolidFill *OOXMLSolidFill     `xml:"solidFill"`
	PattFill  *OOXMLPattFill      `xml:"pattFill"`
	GradFill  *OOXMLThemeGradFill `xml:"gradFill"`
	BlipFill  *OOXMLBlipFill      `xml:"blipFill"`
	EffectLst *OOXMLEffectLst     `xml:"effectLst"`
	Ln        *ChartDPtLn         `xml:"ln"`
}

// ChartDPtLn represents <a:ln> on a chart data point.
type ChartDPtLn struct {
	W         int                 `xml:"w,attr,omitempty"`
	SolidFill *OOXMLSolidFill     `xml:"solidFill"`
	GradFill  *OOXMLThemeGradFill `xml:"gradFill"`
}

// OOXMLPattFill represents <a:pattFill prst="...">.
type OOXMLPattFill struct {
	Prst  string            `xml:"prst,attr"` // pattern preset: "ltUpDiag", "ltDnDiag", etc.
	FgClr *OOXMLColorChoice `xml:"fgClr"`
	BgClr *OOXMLColorChoice `xml:"bgClr"`
}

// OOXMLColorChoice represents a color container (<a:fgClr> or <a:bgClr>) that can hold
// either an sRGB color or a scheme color reference.
type OOXMLColorChoice struct {
	SrgbClr   *OOXMLSrgbClr   `xml:"srgbClr"`
	SchemeClr *OOXMLSchemeClr `xml:"schemeClr"`
	PrstClr   *OOXMLPrstClr   `xml:"prstClr"`
}

// ChartDLbls represents <c:dLbls> - data labels config.
type ChartDLbls struct {
	ShowVal     ChartValAttr `xml:"showVal"`
	ShowCatName ChartValAttr `xml:"showCatName"`
	ShowPercent ChartValAttr `xml:"showPercent"`
	DLblPos     ChartValAttr `xml:"dLblPos"`
}

// ChartSeriesTx represents <c:tx> containing the series name.
type ChartSeriesTx struct {
	StrRef ChartStrRef `xml:"strRef"`
}

// ChartStrRef represents <c:strRef>.
type ChartStrRef struct {
	StrCache ChartStrCache `xml:"strCache"`
}

// ChartStrCache represents <c:strCache>.
type ChartStrCache struct {
	PtCount ChartValAttr `xml:"ptCount"`
	Pts     []ChartStrPt `xml:"pt"`
}

// ChartStrPt represents <c:pt idx="0"><c:v>...</c:v></c:pt>.
type ChartStrPt struct {
	Idx int    `xml:"idx,attr"`
	V   string `xml:"v"`
}

// (Removed ChartSeriesSpPr and ChartSeriesLn in favor of OOXMLSpPr)

// ChartCatData represents <c:cat> or <c:xVal> containing category or x-labels.
type ChartCatData struct {
	StrRef *ChartStrRef `xml:"strRef"`
	NumRef *ChartNumRef `xml:"numRef"`
}

// ChartValData represents <c:val> or <c:yVal> containing numeric data.
type ChartValData struct {
	NumRef *ChartNumRef `xml:"numRef"`
}

// ChartNumRef represents <c:numRef>.
type ChartNumRef struct {
	NumCache ChartNumCache `xml:"numCache"`
}

// ChartNumCache represents <c:numCache>.
type ChartNumCache struct {
	PtCount ChartValAttr `xml:"ptCount"`
	Pts     []ChartNumPt `xml:"pt"`
}

// ChartNumPt represents <c:pt idx="0"><c:v>6</c:v></c:pt>.
type ChartNumPt struct {
	Idx int    `xml:"idx,attr"`
	V   string `xml:"v"`
}

// ChartCatAx represents <c:catAx> (category axis).
type ChartCatAx struct {
	AxPos ChartValAttr `xml:"axPos"` // "b" for bottom
	TxPr  *ChartTxPr   `xml:"txPr"`
}

// ChartValAx represents <c:valAx> (value axis).
type ChartValAx struct {
	AxPos          ChartValAttr `xml:"axPos"` // "l" for left
	MajorGridlines *struct{}    `xml:"majorGridlines"`
}

// ChartLegend represents <c:legend>.
type ChartLegend struct {
	LegendPos ChartValAttr `xml:"legendPos"` // "b", "t", "l", "r"
}

// ======================== Table Style Definitions (a:tblStyleLst) ========================

// OOXMLTableStyleList represents <a:tblStyleLst> from ppt/tableStyles.xml.
type OOXMLTableStyleList struct {
	XMLName xml.Name          `xml:"tblStyleLst"`
	Default string            `xml:"def,attr"`
	Styles  []OOXMLTableStyle `xml:"tblStyle"`
}

// OOXMLTableStyle represents <a:tblStyle>.
type OOXMLTableStyle struct {
	StyleId   string               `xml:"styleId,attr"`
	StyleName string               `xml:"styleName,attr"`
	TblBg     *OOXMLTableStyleBg   `xml:"tblBg"`
	WholeTbl  *OOXMLTableStylePart `xml:"wholeTbl"`
	Band1H    *OOXMLTableStylePart `xml:"band1H"`
	Band2H    *OOXMLTableStylePart `xml:"band2H"`
	Band1V    *OOXMLTableStylePart `xml:"band1V"`
	Band2V    *OOXMLTableStylePart `xml:"band2V"`
	FirstRow  *OOXMLTableStylePart `xml:"firstRow"`
	LastRow   *OOXMLTableStylePart `xml:"lastRow"`
	FirstCol  *OOXMLTableStylePart `xml:"firstCol"`
	LastCol   *OOXMLTableStylePart `xml:"lastCol"`
	SeCell    *OOXMLTableStylePart `xml:"seCell"`
	SwCell    *OOXMLTableStylePart `xml:"swCell"`
	NeCell    *OOXMLTableStylePart `xml:"neCell"`
	NwCell    *OOXMLTableStylePart `xml:"nwCell"`
}

// OOXMLTableStyleBg represents <a:tblBg> (table background).
type OOXMLTableStyleBg struct {
	FillRef   *OOXMLStyleRef `xml:"fillRef"`
	EffectRef *OOXMLStyleRef `xml:"effectRef"`
}

// OOXMLTableStylePart represents a style override section (wholeTbl, band1H, firstRow, etc.).
type OOXMLTableStylePart struct {
	TcTxStyle *OOXMLTcTxStyle      `xml:"tcTxStyle"`
	TcStyle   *OOXMLTableCellStyle `xml:"tcStyle"`
}

// OOXMLTcTxStyle represents <a:tcTxStyle> (table cell text style).
type OOXMLTcTxStyle struct {
	B         string          `xml:"b,attr,omitempty"`
	I         string          `xml:"i,attr,omitempty"`
	FontRef   *OOXMLFontRef   `xml:"fontRef"`
	SchemeClr *OOXMLSchemeClr `xml:"schemeClr"`
}

// OOXMLFontRef represents <a:fontRef idx="minor"> with an optional color child.
type OOXMLFontRef struct {
	Idx       string          `xml:"idx,attr"`
	PrstClr   *OOXMLPrstClr   `xml:"prstClr"`
	SchemeClr *OOXMLSchemeClr `xml:"schemeClr"`
	ScrgbClr  *OOXMLScrgbClr  `xml:"scrgbClr"`
}

// OOXMLScrgbClr represents <a:scrgbClr r="0" g="0" b="0"/>.
type OOXMLScrgbClr struct {
	R int `xml:"r,attr"`
	G int `xml:"g,attr"`
	B int `xml:"b,attr"`
}

// OOXMLTableCellStyle represents <a:tcStyle>.
type OOXMLTableCellStyle struct {
	TcBdr *OOXMLTableCellBorders `xml:"tcBdr"`
	Fill  *OOXMLTableStyleFill   `xml:"fill"`
}

// OOXMLTableCellBorders represents <a:tcBdr>.
type OOXMLTableCellBorders struct {
	Left    *OOXMLTableStyleBorder `xml:"left"`
	Right   *OOXMLTableStyleBorder `xml:"right"`
	Top     *OOXMLTableStyleBorder `xml:"top"`
	Bottom  *OOXMLTableStyleBorder `xml:"bottom"`
	InsideH *OOXMLTableStyleBorder `xml:"insideH"`
	InsideV *OOXMLTableStyleBorder `xml:"insideV"`
}

// OOXMLTableStyleBorder represents a border entry in tcBdr (left, right, top, bottom, insideH, insideV).
// Contains either a direct <a:ln> or a theme <a:lnRef>.
type OOXMLTableStyleBorder struct {
	Ln    *OOXMLLn       `xml:"ln"`
	LnRef *OOXMLStyleRef `xml:"lnRef"`
}

// OOXMLTableStyleFill represents <a:fill> inside a table style part.
type OOXMLTableStyleFill struct {
	SolidFill *OOXMLSolidFill `xml:"solidFill"`
	NoFill    *struct{}       `xml:"noFill"`
}

// ======================== Table (a:tbl) ========================

// OOXMLTable represents <a:tbl> (a DrawingML table embedded in a graphicFrame).
type OOXMLTable struct {
	TblPr   OOXMLTablePr    `xml:"tblPr"`
	TblGrid OOXMLTableGrid  `xml:"tblGrid"`
	Rows    []OOXMLTableRow `xml:"tr"`
}

// OOXMLTablePr represents <a:tblPr> (table properties).
type OOXMLTablePr struct {
	FirstRow     string          `xml:"firstRow,attr,omitempty"` // "1" = first row styling
	LastRow      string          `xml:"lastRow,attr,omitempty"`  // "1" = last row styling
	BandRow      string          `xml:"bandRow,attr,omitempty"`  // "1" = banded rows
	FirstCol     string          `xml:"firstCol,attr,omitempty"` // "1" = first column styling
	LastCol      string          `xml:"lastCol,attr,omitempty"`  // "1" = last column styling
	BandCol      string          `xml:"bandCol,attr,omitempty"`  // "1" = banded columns
	RtlTable     string          `xml:"rtl,attr,omitempty"`      // "1" = right-to-left
	EffectLst    *OOXMLEffectLst `xml:"effectLst"`
	TableStyleId string          `xml:"tableStyleId"` // GUID for predefined table style
}

// OOXMLTableGrid represents <a:tblGrid> containing column definitions.
type OOXMLTableGrid struct {
	GridCols []OOXMLGridCol `xml:"gridCol"`
}

// OOXMLGridCol represents <a:gridCol w="..."/>.
type OOXMLGridCol struct {
	W int64 `xml:"w,attr"` // column width in EMU
}

// OOXMLTableRow represents <a:tr h="...">.
type OOXMLTableRow struct {
	H     int64            `xml:"h,attr"` // row height in EMU
	Cells []OOXMLTableCell `xml:"tc"`
}

// OOXMLTableCell represents <a:tc>.
type OOXMLTableCell struct {
	GridSpan int               `xml:"gridSpan,attr,omitempty"` // number of columns this cell spans
	RowSpan  int               `xml:"rowSpan,attr,omitempty"`  // number of rows this cell spans
	HMerge   string            `xml:"hMerge,attr,omitempty"`   // "1" = horizontally merged (skip rendering)
	VMerge   string            `xml:"vMerge,attr,omitempty"`   // "1" = vertically merged (skip rendering)
	TxBody   *OOXMLTxBody      `xml:"txBody"`
	TcPr     *OOXMLTableCellPr `xml:"tcPr"`
}

// OOXMLTableCellPr represents <a:tcPr> (table cell properties).
type OOXMLTableCellPr struct {
	Anchor    string              `xml:"anchor,attr,omitempty"` // "t", "ctr", "b"
	MarL      *int64              `xml:"marL,attr,omitempty"`   // left margin in EMU
	MarR      *int64              `xml:"marR,attr,omitempty"`   // right margin in EMU
	MarT      *int64              `xml:"marT,attr,omitempty"`   // top margin in EMU
	MarB      *int64              `xml:"marB,attr,omitempty"`   // bottom margin in EMU
	SolidFill *OOXMLSolidFill     `xml:"solidFill"`
	LnL       *OOXMLTableBorderLn `xml:"lnL"` // left border
	LnR       *OOXMLTableBorderLn `xml:"lnR"` // right border
	LnT       *OOXMLTableBorderLn `xml:"lnT"` // top border
	LnB       *OOXMLTableBorderLn `xml:"lnB"` // bottom border
	NoFill    *struct{}           `xml:"noFill"`
}

// OOXMLTableBorderLn represents a table cell border line (e.g., <a:lnL w="12700">).
type OOXMLTableBorderLn struct {
	W         int             `xml:"w,attr,omitempty"` // line width in EMU
	SolidFill *OOXMLSolidFill `xml:"solidFill"`
	NoFill    *struct{}       `xml:"noFill"`
}

// ======================== Extended Chart XML (cx:chartSpace) ========================

// ExtendedChartSpace represents <cx:chartSpace> (Office 2016+ charts).
type ExtendedChartSpace struct {
	XMLName   xml.Name          `xml:"chartSpace"`
	ChartData ExtendedChartData `xml:"chartData"`
	Chart     ExtendedChart     `xml:"chart"`
	SpPr      *ChartSpaceSpPr   `xml:"spPr"` // chart-level formatting (background)
}

type ExtendedChartData struct {
	Data []ExtendedData `xml:"data"`
}

type ExtendedData struct {
	Id     string          `xml:"id,attr"`
	StrDim *ExtendedStrDim `xml:"strDim"`
	NumDim *ExtendedNumDim `xml:"numDim"`
}

type ExtendedStrDim struct {
	Type string        `xml:"type,attr"`
	Lvls []ExtendedLvl `xml:"lvl"`
}

type ExtendedNumDim struct {
	Type string        `xml:"type,attr"`
	Lvls []ExtendedLvl `xml:"lvl"`
}

type ExtendedLvl struct {
	PtCount string              `xml:"ptCount,attr"`
	Pts     []ExtendedDataPoint `xml:"pt"`
}

type ExtendedDataPoint struct {
	Idx int    `xml:"idx,attr"`
	V   string `xml:",chardata"`
}

type ExtendedChart struct {
	Title    *ExtendedTitle   `xml:"title"`
	PlotArea ExtendedPlotArea `xml:"plotArea"`
}

type ExtendedPlotArea struct {
	PlotAreaRegion ExtendedPlotAreaRegion `xml:"plotAreaRegion"`
	Axes           []ExtendedAxis         `xml:"axis"`
}

type ExtendedTitle struct {
	Tx *ExtendedTx `xml:"tx"`
}

type ExtendedPlotAreaRegion struct {
	Series []ExtendedSeries `xml:"series"`
}

type ExtendedAxis struct {
	Id         string              `xml:"id,attr"`
	CatScaling *ExtendedCatScaling `xml:"catScaling"`
}

type ExtendedCatScaling struct {
	GapWidth string `xml:"gapWidth,attr"`
}

type ExtendedSeries struct {
	LayoutId string            `xml:"layoutId,attr"`
	Tx       *ExtendedTx       `xml:"tx"`
	LayoutPr *ExtendedLayoutPr `xml:"layoutPr"`
	SpPr     *OOXMLSpPr        `xml:"spPr"`
}

type ExtendedLayoutPr struct {
	Binning *ExtendedBinning `xml:"binning"`
}

type ExtendedBinning struct {
	IntervalClosed string  `xml:"intervalClosed,attr"`
	BinCount       int     `xml:"binCount,attr"`
	BinSize        float64 `xml:"binSize,attr"`
}

type ExtendedTx struct {
	TxData *ExtendedTxData `xml:"txData"` // some use txData
	Rich   *ChartTitleRich `xml:"rich"`   // some use rich text (reusing standard)
}

type ExtendedTxData struct {
	V string `xml:"v"`
}
