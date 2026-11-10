package pptxtosvg

import (
	"encoding/xml"
	"strings"
	"testing"
)

// Test bodyPr unmarshal: noAutofit, normAutofit, spAutoFit must be typed (no innerXML).
// Verifies that autofit options are well-defined in the OOXML model.

func TestOOXMLBodyPr_Unmarshal_NoAutofit(t *testing.T) {
	// DrawingML snippet: Do not Auto Fit
	snippet := `<bodyPr wrap="none" rtlCol="0"><noAutofit/></bodyPr>`
	var body OOXMLBodyPr
	err := xml.Unmarshal([]byte(snippet), &body)
	if err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if body.NoAutofit == nil {
		t.Error("expected NoAutofit to be set (typed), got nil")
	}
	if body.NormAutofit != nil || body.SpAutoFit != nil {
		t.Error("expected only NoAutofit present")
	}
	if body.Wrap != "none" {
		t.Errorf("expected wrap=none, got %q", body.Wrap)
	}
}

func TestOOXMLBodyPr_Unmarshal_NormAutofit(t *testing.T) {
	// Shrink Text on Overflow with attributes
	snippet := `<bodyPr rtlCol="0" anchor="ctr"><normAutofit fontScale="92.000%" lnSpcReduction="20.000%"/></bodyPr>`
	var body OOXMLBodyPr
	err := xml.Unmarshal([]byte(snippet), &body)
	if err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if body.NormAutofit == nil {
		t.Fatal("expected NormAutofit to be set (typed), got nil")
	}
	if body.NormAutofit.FontScale != "92.000%" {
		t.Errorf("expected fontScale=92.000%%, got %q", body.NormAutofit.FontScale)
	}
	if body.NormAutofit.LnSpcReduction != "20.000%" {
		t.Errorf("expected lnSpcReduction=20.000%%, got %q", body.NormAutofit.LnSpcReduction)
	}
	if body.Anchor != "ctr" {
		t.Errorf("expected anchor=ctr, got %q", body.Anchor)
	}
	if body.NoAutofit != nil || body.SpAutoFit != nil {
		t.Error("expected only NormAutofit present")
	}
}

func TestOOXMLBodyPr_Unmarshal_SpAutoFit(t *testing.T) {
	// Resize Shape to Fit Text
	snippet := `<bodyPr rtlCol="0" anchor="ctr"><spAutoFit/></bodyPr>`
	var body OOXMLBodyPr
	err := xml.Unmarshal([]byte(snippet), &body)
	if err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if body.SpAutoFit == nil {
		t.Error("expected SpAutoFit to be set (typed), got nil")
	}
	if body.NoAutofit != nil || body.NormAutofit != nil {
		t.Error("expected only SpAutoFit present")
	}
}

func TestOOXMLBodyPr_Unmarshal_AllAttributesAndNoAutofit(t *testing.T) {
	// bodyPr with insets and noAutofit (typical "Do not Auto Fit" text box)
	snippet := `<bodyPr wrap="square" lIns="91440" rIns="91440" tIns="45720" bIns="45720"><noAutofit/></bodyPr>`
	var body OOXMLBodyPr
	err := xml.Unmarshal([]byte(snippet), &body)
	if err != nil {
		t.Fatalf("Unmarshal: %v", err)
	}
	if body.NoAutofit == nil {
		t.Error("expected NoAutofit to be set")
	}
	if body.Wrap != "square" {
		t.Errorf("wrap: got %q", body.Wrap)
	}
	if body.LIns == nil || *body.LIns != 91440 {
		t.Errorf("lIns: got %v", body.LIns)
	}
}

// Round-trip: marshal back to XML and assert autofit elements are present.
func TestOOXMLBodyPr_MarshalRoundTrip(t *testing.T) {
	body := OOXMLBodyPr{
		Anchor:      "t",
		Wrap:        "square",
		NoAutofit:   &struct{}{},
		NormAutofit: nil,
		SpAutoFit:   nil,
	}
	out, err := xml.Marshal(body)
	if err != nil {
		t.Fatalf("Marshal: %v", err)
	}
	str := string(out)
	if !strings.Contains(str, "noAutofit") {
		t.Errorf("marshal output should contain noAutofit element: %s", str)
	}

	body2 := OOXMLBodyPr{
		Anchor: "ctr",
		NormAutofit: &OOXMLNormAutofit{
			FontScale:      "92.000%",
			LnSpcReduction: "20.000%",
		},
	}
	out2, err := xml.Marshal(body2)
	if err != nil {
		t.Fatalf("Marshal normAutofit: %v", err)
	}
	str2 := string(out2)
	if !strings.Contains(str2, "normAutofit") {
		t.Errorf("marshal output should contain normAutofit element: %s", str2)
	}
	if !strings.Contains(str2, "92.000%") {
		t.Errorf("marshal output should contain fontScale: %s", str2)
	}

	body3 := OOXMLBodyPr{
		Anchor:    "b",
		SpAutoFit: &struct{}{},
	}
	out3, err := xml.Marshal(body3)
	if err != nil {
		t.Fatalf("Marshal spAutoFit: %v", err)
	}
	str3 := string(out3)
	if !strings.Contains(str3, "spAutoFit") {
		t.Errorf("marshal output should contain spAutoFit element: %s", str3)
	}
}
