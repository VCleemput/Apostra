#Requires AutoHotkey v2.0
#Include WinClipAPI.ahk ;includes clipboard functionality
#Include WinClip.ahk
#SingleInstance

; define a global variable to drop the generated text to, in case something goes wrong
global backup
backup := ""

SendHTML(html, aw := WinExist("A")) 
{
	;function to send the html to the clipboard and paste it
    global backup
	wc := Winclip()
	html := '<div style="font-size:11pt">' html  '</div>'
	;OldClipboard := WinClip.Snap() ;om de een of andere reden werkt 'm zonder dit niet in WORD
	OldClipboard := A_Clipboard
	A_Clipboard :=""
	wc.Clear()
	Sleep(100)
	wc.SetHTML(html)
	ClipWait(,1)
	wc.SetText(html)
	Sleep(100)
	ClipWait(,1)
	backup := html
	if WinExist(aw)
		{
			WinActivate(aw)
			WinWaitActive(aw)
		}
	wc.Paste()
	Sleep(100)
	ClipWait(,1)
	A_Clipboard := OldClipboard
}

StrJoin(obj, delimiter:="", OmitChars:="")
{
    ;joins an array of strings
    ;copied from: https://www.autohotkey.com/boards/viewtopic.php?style=19&t=25918
    string := obj[1]
    loop(obj.Length-1)
        string .= delimiter . Trim(obj[A_Index+1], OmitChars)
    return string
}

#b:: 
{ ;resends the previously sent text, in case something goes wrong
SendHTML(backup)
return
}

::*qs::
{	;ER, PR, HER2 and ki67
	aw := WinExist("A") ; captures the current window, so at the end the text is pasted in the same window
	global synopsis ; a value to store the synopsis, to be passed to the next function
	MyGui := Gui(, "Farmacodiagnostiek borst") ; create GUI
	ERCheck := MyGui.AddCheckbox("xm vReceptorER Checked", "ER")
	ERCheck.OnEvent("Click", ER)
	ERpctTekst := MyGui.AddText("xm section vtekstER", "Percentage gekleurde kernen")
	ERpct := MyGui.AddEdit("ys wp", "")
	ERintTekst := MyGui.AddText("xm section vtekstIntER", "Intensiteit aankleuring")
	ERint := MyGui.AddDDL("ys wp Choose4",["afwezig", "zwak", "matig", "sterk"] )
	PRCheck := MyGui.AddCheckbox("xm vReceptorPR Checked", "PR")
	PRCheck.OnEvent("Click", PR)
	PRpctTekst := MyGui.AddText("xm section vtekstPR", "Percentage gekleurde kernen")
	PRpct := MyGui.AddEdit("ys wp", "")
	PRintTekst := MyGui.AddText("xm section vtekstIntPR", "Intensiteit aankleuring")
	PRint := MyGui.AddDDL("ys wp Choose4",["afwezig", "zwak", "matig", "sterk"] )
	HER2Check := MyGui.AddCheckbox("xm section vHER2R Checked", "HER2")
	HER2Check.OnEvent("Click", HER2)
	HER2ind := MyGui.AddDDL("ys wp Choose1", ["borst", "maag", "andere"])
	HER2ihcTekst := MyGui.AddText("xm section vtekstHER2", "IHC score (zonder +)")
	HER2ihc := MyGui.AddEdit("ys wp", "")
	kiCheck := MyGui.AddCheckbox("xm vki Checked", "Ki67")
	kiCheck.OnEvent("Click", ki)
	kiscoreTekst := MyGui.AddText("xm section vtekstki", "Ki67-index (zonder %)")
	kiscore := MyGui.AddEdit("ys wp", "")
	MyGui.AddText("section xm", "Nota: na de sneltekst kan een synopsis geplakt worden met 'Win+s'")
	MyGui.AddButton("xm default", "OK").OnEvent("Click", _qsButtonOK)
	MyGui.Show()

ER(*) ; Do this when ER is checked
{
	if ERCheck.value = 1
		{
			ERint.Enabled := 1
			ERintTekst.Enabled := 1
			ERpct.Enabled := 1
			ERpctTekst.Enabled := 1
		}
	if ERCheck.value = 0
		{
			ERint.Enabled := 0
			ERintTekst.Enabled := 0
			ERpct.Enabled := 0
			ERpctTekst.Enabled := 0
		}
}
	
PR(*)
{
	if PRCheck.value = 1
		{
			PRint.Enabled := 1
			PRintTekst.Enabled := 1
			PRpct.Enabled := 1
			PRpctTekst.Enabled := 1
		}
	if PRCheck.value = 0
		{
			PRint.Enabled := 0
			PRintTekst.Enabled := 0
			PRpct.Enabled := 0
			PRpctTekst.Enabled := 0
		}
	}

HER2(*)
{
	if HER2Check.value = 1
		{
			HER2ihc.Enabled := 1
			HER2ind.Enabled := 1
			HER2ihcTekst.Enabled := 1
		}
	if HER2Check.value = 0
		{
			HER2ihc.Enabled := 0
			HER2ind.Enabled := 0
			HER2ihcTekst.Enabled := 0
		}
}

ki(*)
{
	if kiCheck.value = 1
		{
			kiscoreTekst.Enabled := 1
			kiscore.Enabled := 1
		}
	if kiCheck.value = 0
		{
			kiscoreTekst.Enabled := 0
			kiscore.Enabled := 0
		}
}
	
_qsButtonOK(*)
{
	MyGui.Hide()
	if ERpct.text = "<1"
		ERpct.text := "< 1"
	if PRpct.text = "<1"
		PRpct.text := "< 1"
	QuickScore(Pct,Int) {
		I := Map("afwezig", 0, "zwak", 1, "matig", 2, "sterk", 3)

		if (Pct = 0)
			Pe := 0
		else if ((Pct = "< 1"))
			Pe := 1
		else if (Pct < 1) 
			Pe := 1
		else if (Pct <= 10)
			Pe := 2
		else if (Pct<= 33)
			Pe := 3
		else if (Pct <= 66)
			Pe := 4
		else if (Pct <= 100)
			Pe := 5
		else
			return
		value := I[Int]
		qs := value + Pe
		if ((Pct = "< 1"))
			waarde := "Negatief"
		else if Pct < 1
			waarde := "Negatief"
		else if Pct <=10
			waarde := "Zwak positief"
		else
			waarde := "Positief"
		if (Pe = 0 or value = 0)
			tekst := "Negatief. Allred score 0/8. Geen aankleuring in de laesionele cellen."
		else
			tekst := waarde ". " Pct " pct van de kernen kleurt " Int  " aan. Allred score " Pe " + " value " = " qs "/8."
		return tekst
	}

	ReceptortekstER :=""
	ReceptortekstPR :=""
	ABtekstER :=""
	ABtekstPR :=""
	ABtekst :=""
	HER2tekst :=""
	kitekst :=""
	synopsis :=""

    stainer := IniRead("lab-variables.ini", "breast fd","stainer")
    ab_ER := IniRead("lab-variables.ini", "breast fd","ab_ER")
    ab_PR := IniRead("lab-variables.ini", "breast fd","ab_PR")
    ab_HER2 := IniRead("lab-variables.ini", "breast fd","ab_HER2")
    ASCO_jaar_ER := IniRead("lab-variables.ini", "breast fd","ASCO_jaar_ER")
    ASCO_jaar_HER2_borst := IniRead("lab-variables.ini", "breast fd","ASCO_jaar_HER2_borst")
    ASCO_jaar_rest := IniRead("lab-variables.ini", "breast fd","ASCO_jaar_rest")

	if ERCheck.value = 1
	{
		w := QuickScore(ERpct.text, ERint.text)
		ReceptortekstER := "Oestrogeenreceptor (ER): " w "<br>"
		ABtekstER := ab_ER . ", "
		RegExMatch(w, "\d\/8", &z)
		synopsis := synopsis . "ER " . z[] . "; "
	}
	if PRCheck.value = 1
	{
		w := QuickScore(PRpct.text, PRint.text)
		ReceptortekstPR := "Progesteronreceptor (PR): " w "<br>"
        ABtekstPR := ab_PR . ", "
		RegExMatch(w, "\d\/8", &z)
		synopsis := synopsis . "PR " . z[] . "; "
	}

	if ((ERCheck.value = 1) or (PRCheck.value = 1))
		{
			ABtekst := "<small>Kloon " AbtekstER AbtekstPR "op " . stainer . ". Interpretatie volgens de ASCO/CAP guidelines " . ASCO_jaar_ER . ".</small><br><br>"
		}
	if HER2Check.value = 1
		{
			if (HER2ind.text = "borst") {
				ASCOjaar := ASCO_jaar_HER2_borst
			}
			else {
				ASCOjaar := ASCO_jaar_rest
			}
			SishVraagTekst := ""
			if (((HER2ind.text = "borst") or (HER2ind.text = "maag")) and ((HER2ihc.text = 2) or (HER2ihc.text = 3))) {
				SishVraagTekst := " ISH volgt."
			}
			HER2tekst := 
			(
				"HER2 IHC score: " HER2ihc.text  "+. " SishVraagTekst "<br>"
				"<small>" . ab_HER2 . " op " . stainer . ". Interpretatie volgens de ASCO/CAP guidelines " ASCOjaar ".</small><br><br>"
			)
			synopsis := synopsis . "HER2 IHC score: " . HER2ihc.text . "+; "
		}
	if kiCheck.value = 1
		{
			kitekst := "Ki67-index: " kiscore.text "%.<br>"
			synopsis := synopsis . "ki67: " . kiscore.text . "%"
		}
	tekst :=
(
	ReceptortekstER
	ReceptortekstPR
	ABtekst
	HER2tekst
	kitekst
)
	SendHTML(tekst, aw)
	MyGui.Destroy()
	
}
}

#s::
{
SendHTML(synopsis)
}

::*88::
{	;Melanoom
	aw := WinExist("A")
	MyGui := Gui(, "Melanoom")
	MyGui.AddText("xm section w100", "Locatie")
	locatie := MyGui.AddEdit("ys w400", "")
	MyGui.AddText("xm section w100", "Type")
	typeMM := MyGui.AddComboBox("ys w400 Choose1", ["superficieel spreidend maligne melanoom","lentigo maligne melanoma","acrolentigineus maligne melanoom"])
	MyGui.AddText("xm section w100", "groeifase")
	groeifase := MyGui.AddDDL("ys w400 choose1", ["zuivere radiale","invasieve radiale","invasieve verticale","invasieve horizontale"])
	MyGui.AddText("xm section w100", "Clark Level")
	clarkLevel := MyGui.AddDDL("ys w400 choose1", ["I = melanoma in situ", "II = melanoma aanwezig in papillaire dermis, maar geen opvullen of expansie van papillaire dermis", " III = melanoma aanwezig in papillaire dermis met opvullen en expansie van papillaire dermis", "IV = melanoma invadeert reticulaire dermis", "V = melanoma invadeert subcutis"])
	MyGui.AddText("xm section w100", "Breslow (mm)")
	breslow := MyGui.AddEdit("ys w400", "")
	MyGui.AddText("xm section w100", "LVI")
	lvi := MyGui.AddCheckbox("ys w400", "")
	MyGui.AddText("xm section w100", "PNI")
	pni := MyGui.AddCheckbox("ys w400", "")
	MyGui.AddText("xm section w100", "Ulceratie")
	ulceratie := MyGui.AddCheckbox("ys w400", "")
	MyGui.AddText("xm section w100", "mitosen per mm²")
	mitosenPerMm2 := MyGui.AddEdit("ys w400", "")
	MyGui.AddText("xm section w100", "Stromale afweer")
	stromaleAfweer := MyGui.AddComboBox("ys w400 Choose1", ["niet aantoonbaar","aanwezig, non-brisk","aanwezig, brisk"])
	MyGui.AddText("xm section w100", "Regressie")
	regressie := MyGui.AddCheckbox("ys w400", "")
	MyGui.AddText("xm section w100", "Satellietletsels")
	satellietletsel := MyGui.AddCheckbox("ys w400", "")
	MyGui.AddText("xm section w100", "Voorafbestaande naevus")
	voorafbestaandeNaevus := MyGui.AddCheckbox("ys w400", "")
	MyGui.AddText("xm section w100", "Snederanden")
	snederanden := MyGui.AddEdit("ys w400", "In toto verwijderd")
	MyGui.AddText("xm section w100", "Andere letsels")
	andereLetsels := MyGui.AddEdit("ys w400", "Geen")
	MyGui.AddText("xm section w100", "TNM")
	tnm := MyGui.AddEdit("ys w400", "pT1a")
	MyGui.AddButton("xm w50 h20 default", "OK").OnEvent("click", _88ButtonOK)
	MyGui.Show()


_88ButtonOK(*)
{
	RegExMatch(clarkLevel.text, "[IV]*", &clark)
	clark := clark[]
	bin := Map(0, "niet aantoonbaar", 1, "aanwezig")
	lvi := bin[lvi.value]
	pni := bin[pni.value]
	ulceratie := bin[ulceratie.value]
	satellietletsel := bin[satellietletsel.value]
	voorafbestaandeNaevus := bin[voorafbestaandeNaevus.value]
	regressie := bin[regressie.value]

	html := 
	(
	"<b>Microscopie:</b><br>"
	"- Histologisch tumortype (WHO) : " typeMM.text "<br>"
	"- Groeifase : " groeifase.text " groeifase<br>"
	"- Clark level : " clarkLevel.text "<br>"
	"- Breslow-dikte : " breslow.text "mm<br>"
	"- Lymfovasculaire invasie : " lvi "<br>"
	"- Perineurale invasie : " pni "<br>"
	"- Ulceratie : " ulceratie "<br>"
	"- Mitosen : " mitosenPerMm2.text " mitosen per mm²<br>"
	"- Stromale afweerreactie : " stromaleAfweer.text "<br>"
	"- Regressie : " regressie "<br>"
	"- Satellietletsels : " satellietletsel "<br>"
	"- Voorafbestaande naevus : " voorafbestaandeNaevus "<br>"
	"- Snijranden : " snederanden.text "<br>"
	"- Andere letsels : " andereLetsels.text "<br>"
	"<br>"
	"<b>Besluit:</b><br>"
	"Huidexcisie " locatie.text "<br>"
	"- " typeMM.text " in " groeifase.text " groeifase <br>"
	"- Clark level " clark "<br>"
	"- Breslow-dikte " breslow.text " mm<br>"
	"- Ulceratie: " ulceratie "<br>"
	"- " snederanden.text "<br><br>"
	"- Voorstel stadiëring (8e ed. TNM): " tnm.text "<br>"
	)
	SendHTML(html, aw)
	MyGui.Destroy()
}
}

IniListRead(path, section, key)
{
	return StrSplit(IniRead(path, section, key), ";")
}
::*pdl1::
{	aw := WinExist("A")
	MyGui := Gui(, "pdl1")
	Patholoog_tekst := MyGui.AddText("xm section w200", "Patholoog")
	Patholoog := MyGui.AddDropDownList("ys w500 Choose1", [A_UserName, "AD", "AH", "FC", "DC", "KV", "LF", "JVD", "SV"])
	Orgaan_tekst := MyGui.AddText("xm section w200", "Orgaan")
	Matrix := MyGui.AddDropDownList("ys w500 Choose1", ["Blaas", "Borst", "Cervix", "Hoofd hals", "Long", "Slokdarm-maag Adenocarcinoom", "Slokdarm plaveiselcelcarcinoom"])
	ScoreText := MyGui.AddText("xm section w200", "Score Type:")
	ScoreType := MyGui.AddText("ys section w200")
	
; Create CPS and TPS score inputs
ScoreTextCPS := MyGui.AddText("xm section w200", "CPS-score:")
ScoreEditCPS := MyGui.AddEdit("ys w200")
ScoreTextTPS := MyGui.AddText("xm section w200", "TPS-score:")
ScoreEditTPS := MyGui.AddEdit("ys w200")

ExternalControlsText := MyGui.AddText("xm section w200", "Externe/interne controles:")
ExternalControlsCheckbox := MyGui.AddCheckbox("xm section w200 Checked", "Controles OK")

MyGui.AddButton("xm w50 h20 default", "OK").OnEvent("click", _PDL1ButtonOK)
MyGui.Show()
Matrix.OnEvent("Change", _OrganSelectionChanged)

_OrganSelectionChanged(*)
{
    selectedOrgan := Matrix.Text

    ; Show or hide CPS and TPS score input fields based on organ selection
    ScoreTextCPS.Enabled := (selectedOrgan = "Blaas" || selectedOrgan = "Slokdarm plaveiselcelcarcinoom" || selectedOrgan = "Borst" || selectedOrgan = "Cervix" || selectedOrgan = "Hoofd Hals" || selectedOrgan = "Slokdarm-maag Adenocarcinoom")
    ScoreEditCPS.Enabled := ScoreTextCPS.Enabled
    ScoreTextTPS.Enabled := ( selectedOrgan = "Blaas" || selectedOrgan = "Slokdarm plaveiselcelcarcinoom" || selectedOrgan = "Long")
    ScoreEditTPS.Enabled := ScoreTextTPS.Enabled
}

_PDL1ButtonOK(*)
{
	currentDate := FormatTime(,"dd/MM/yyyy")

    ; Get the  selected organ, scoring method, and entered score
    selectedOrgan := Matrix.Text
    scoringMethod := ScoreType.Text
    CPS_Score := ScoreEditCPS.Text
    TPS_Score := ScoreEditTPS.Text
    externalControlsOK := ExternalControlsCheckbox.Value
	Patholoogname := Patholoog.text

    ; Generate HTML based on the selected organ and scoring method
    result := SetScoresAndCheckPositivity(selectedOrgan, scoringMethod, CPS_Score, TPS_Score, externalControlsOK, currentDate, Patholoogname)
    SendHTML(result, aw)
    MyGui.Destroy()
}

SetScoresAndCheckPositivity(organ, scoreType, CPS_Score, TPS_Score, externalControlsOK, currentDate, Patholoog)
{
    ; Initialize variables for CPS and TPS thresholds for each organ
    cpsThreshold := 10  ; Default CPS threshold
    tpsThreshold := 1   ; Default TPS threshold
	cpsInterpretatie := ""
	tpsInterpretatie := ""
	explanation := ""

	  ; Check if the entered score is a number
	  explanation := ""
	  fout := 0
	  if ScoreEditCPS.Enabled and not IsNumber(CPS_Score)
		  fout := 1
	  if ScoreEditTPS.Enabled and not IsNumber(TPS_Score)
		  fout := 1
	  if fout = 0 {
	
    ; Set thresholds based on the selected organ
    switch (organ) {
        case "Blaas":
            cpsThreshold := 10  ; Threshold for CPS for Blaas
            tpsThreshold := 1  ; Threshold for TPS for Blaas
            cpsInterpretatie := "Combined positivity score (CPS): het percentage PD-L1 aankleurende cellen (tumorcellen en immuuncellen) gedeeld door het totaal aantal viabele tumorcellen x 100 (= score). Voor behandeling met Pembrolizumab bedraagt de cut-off waarde > of = 10."
            tpsInterpretatie := "Tumor Proportion Score (TPS): het aantal PD-L1 aankleurende tumorcellen gedeeld door het totaal aantal viabele tumorcellen (= percentage). Voor behandeling met Nivolumab bedraagt de cut-off waarde > of = 1%."
        case "Borst":
            cpsThreshold := 10  ; Threshold for CPS for Borst
            cpsInterpretatie := "Combined positivity score (CPS): het percentage PD-L1 aankleurende cellen (tumorcellen en immuuncellen) gedeeld door het totaal aantal viabele tumorcellen x 100 (= score). Voor behandeling met Pembrolizumab bedraagt de cut-off waarde > of = 10."	
        case "Cervix":
            cpsThreshold := 1  ; Threshold for CPS for Cervix
            cpsInterpretatie := "Combined positivity score (CPS): het percentage PD-L1 aankleurende cellen (tumorcellen en immuuncellen) gedeeld door het totaal aantal viabele tumorcellen x 100 (= score). Voor behandeling met Pembrolizumab bedraagt de cut-off waarde > of = 1."
	case "Hoofd hals":
            cpsThreshold := 1  ; Threshold for CPS for Hoofd Hals
            cpsInterpretatie := "Combined positivity score (CPS): het percentage PD-L1 aankleurende cellen (tumorcellen en immuuncellen) gedeeld door het totaal aantal viabele tumorcellen x 100 (= score). Voor behandeling met Pembrolizumab bedraagt de cut-off waarde > of = 1."
	case "Slokdarm plaveiselcelcarcinoom":
	     cpsThreshold := 10  ; Threshold for CPS for Slokdarm plaveiselcelcarcinoom
	     tpsThreshold := 1  ; Threshold for TPS for Slokdarm plaveiselcelcarcinoom
	     cpsInterpretatie := "Combined positivity score (CPS): het percentage PD-L1 aankleurende cellen (tumorcellen en immuuncellen) gedeeld door het totaal aantal viabele tumorcellen x 100 (= score). Voor behandeling met Pembrolizumab bedraagt de cut-off waarde > of = 10."
	     tpsInterpretatie := "Tumor Proportion Score (TPS): het aantal PD-L1 aankleurende tumorcellen gedeeld door het totaal aantal viabele tumorcellen (= percentage). Voor behandeling met Nivolumab bedraagt de cut-off waarde > of = 1%."
	case "Slokdarm-maag Adenocarcinoom":
            cpsThreshold := 5  ; Threshold for CPS for Slokdarm-maag Adenocarcinoom
            cpsInterpretatie := "Combined positivity score (CPS): het percentage PD-L1 aankleurende cellen (tumorcellen en immuuncellen) gedeeld door het totaal aantal viabele tumorcellen x 100 (= score). Voor behandeling met Nivolumab bedraagt de cut-off waarde > of = 5."
	case "Long":
    	tpsThreshold := 1  ; Threshold for TPS for long
    	tpsInterpretatie := "Tumor Proportion Score (TPS): het aantal PD-L1 aankleurende tumorcellen gedeeld door het totaal aantal viabele tumorcellen (= percentage). Voor eerstelijnsbehandeling met Pembrolizumab/ Cemiplimab bedraagt de cut-off waarde > of = 50%. Voor eerstelijnsbehandeling met Durvalumab bedraagt de cut-off waarde > of = 1%. Voor tweedelijnsbehandeling met Pembrolizumab bedraagt de cut-off waarde > of = 1 %."
		}

	 ; Check if CPS is positive
if (ScoreEditCPS.Enabled and CPS_Score != "") {
    if (IsNumber(CPS_Score)) {
        CPS_Score := Round(CPS_Score) ; Round the entered score to an integer
        if (CPS_Score >= cpsThreshold) {
            resultaatCPS := "positief"
        } else {
            resultaatCPS := "negatief"
        }
        explanation .= cpsInterpretatie
    } else {
        return "Invalid CPS score. Please enter a number."
    }
}

; Check if TPS is positive
if (ScoreEditTPS.Enabled and TPS_Score != "") {
    if (IsNumber(TPS_Score)) {
        TPS_Score := Round(TPS_Score) ; Round the entered score to an integer
        if (TPS_Score >= tpsThreshold) {
            resultaatTPS := "positief"
        } else {
            resultaatTPS := "negatief"
        }
        if explanation != ""
            explanation .= "<br>"
        explanation := explanation . tpsInterpretatie
    } else {
        return "Invalid TPS score. Please enter a number."
    }
	}
        ; Include the external controls status in the HTML
        if (externalControlsOK) {
            externalControlsStatus := "Externe/interne controles zijn conform de vooropgestelde criteria."
        } else {
            externalControlsStatus := "Externe/interne controles zijn niet conform de vooropgestelde criteria."
        }

        ; Generate the HTML output 
		if (organ = "Long") {
			html := "<b>Immuunhistochemie voor PD-L1 (" . FormatTime(,"dd/MM/yyyy") . ";" . Patholoog . ")</b><br>"
			html .= "<b>Locatie:</b> " organ "<br>"
			html .= "<b>Techniek:</b> uitgevoerd met 22C3 antilichaam (Agilent) op het Benchmark Ultra toestel (Roche).<br>"
			html .= "<b>Interpretatie:</b> " explanation "<br>"
			html .= "<b>Externe/interne controle:</b> " externalControlsStatus "<br>"
			html .= "<b>PD-L1 Score (22C3, Agilent):</b><br>"
			html .= "- TPS = " . TPS_Score . "%.<br>"
		} else {
			html := "<b>Immuunhistochemie voor PD-L1 (" FormatTime(,"dd/MM/yyyy") ";" Patholoog ")" "</b><br>"
			html .= "<b>Locatie:</b> " organ "<br>"
			html .= "<b>Techniek:</b> uitgevoerd met 22C3 antilichaam (Agilent) op het Benchmark Ultra toestel (Roche).<br>"
			html .= "<b>Interpretatie:</b> " explanation "<br>"
			html .= "<b>Externe/interne controle:</b> " externalControlsStatus "<br>"
			html .= "<b>PD-L1 Score (22C3, Agilent):</b> <br>" 
			if ScoreEditCPS.Enabled
				html .= "- CPS = " CPS_Score ". Dit komt overeen met een " resultaatCPS " resultaat. <br>"
			if ScoreEditTPS.Enabled
				html .= "- TPS = " . TPS_Score . "%. Dit komt overeen met een " . resultaatTPS . " resultaat. <br>"
	}
		return html
    
}
}
}
::*41::
{	; Colonresectie
	aw := WinExist("A")
	MyGui := Gui(, "Colonresectie")
	MyGui.AddText("section w200", "Type")
	typeCarcinoom := MyGui.AddComboBox("ys w200 Choose1", ["adenocarcinoom NST","mucineus adenocarcinoom","zegelringcelcarcinoom","neuroendocrien carcinoom","plaveiselcelcarcinoom"])
	MyGui.AddText("xs section w200", "Gradering")
	gradering := MyGui.AddDDL("ys w200 Choose1", ["1","2","3","niet van toepassing"])
	MyGui.AddText("xs section w200", "Invasiediepte")
	invasiediepte := MyGui.AddComboBox("ys w200 Choose1",["intramucosaal adenocarcinoom", "submucosa", "muscularis propria", "omliggend vetweefsel", "serosa, met serosale doorbraak"])
	MyGui.AddText("xs section w200", "LVI")
	lvi := MyGui.AddComboBox("ys w200 Choose1", ["afwezig","aanwezig"])
	MyGui.AddText("xs section w200", "PNI")
	pni := MyGui.AddComboBox("ys w200 Choose1", ["afwezig","aanwezig"])
	MyGui.AddText("xs section w200", "SV")
	sv := MyGui.AddComboBox("ys w200 Choose1", ["tumorvrij","positief"])
	MyGui.AddText("xs section w200", "Budding")
	budding := MyGui.AddComboBox("ys w200 Choose1", ["afwezig","aanwezig: beperkt","aanwezig: matig","aanwezig:uitgesproken"])
	MyGui.AddText("xs section w200", "Extramurale deposits")
	extramuraleDeposits := MyGui.AddComboBox("ys w200 Choose1", ["afwezig","aanwezig"])
	MyGui.AddText("xs section w200", "Totaal # LK")
	totaalLk := MyGui.AddEdit("ys w200", "")
	MyGui.AddText("xs section w200", "# Positieve LK")
	positiefLk := MyGui.AddEdit("ys w200", "")
	MyGui.AddText("xs section w200", "Precursorletsel")
	precursorLetsel := MyGui.AddEdit("ys w200", "Niet aantoonbaar")
	MyGui.AddText("xs section w200", "Specimen")
	specimen := MyGui.AddComboBox("ys w200 Choose1", ["distaal colon", "colon transversum", "colon ascendens", "sigmoid", "rectum", "locatie niet gegeven"])
	specimen.OnEvent("change", _togglerectum)
	MyGui.AddText("xs section w200", "Diameter tumor(mm)")
	diameter := MyGui.AddEdit("ys w200", "")
	MyGui.AddText("xs section w200", "Afstand snedevlak(cm)")
	afstandSnedevlak := MyGui.AddEdit("ys w200", "")
	crmTekst := MyGui.AddText("xs section w200", "CRM(mm)")
	crm := MyGui.AddEdit("ys w200", "")
	MyGui.AddText("xs section w200", "TNM")
	tnm := MyGui.AddEdit("ys w200", "")
	MyGui.AddText("xs section w200", "NACT?")
	chemoCheckBox := MyGui.AddCheckbox("ys  w200", "")
	chemoCheckBox.OnEvent("click", _toggleNACT)
	MyGui.AddText("xs section w100", "Dvorak Regression")
	dvorak := MyGui.AddDDL("ys w300 R4 Choose1", ["GR 4: geen tumorcellen, louter fibrose (= complete respons)", "GR 3: klein aantal tumor cellen (moeilijk te vinden met de microscoop op kleine vergroting) in een dominant fibrotisch stroma", "GR 2: klein aantal tumor cellen (makkelijk te vinden met de microscoop op kleine vergroting) in een dominant fibrotisch stroma", "GR 1: de tumor domineert op de fibrotische massa ", "GR 0: geen regressie"])
	MyGui.AddButton("xm w50 h20 default", "OK").OnEvent("click", _41ButtonOK)
	_togglerectum()
	_toggleNACT()
	MyGui.Show()
	


_togglerectum(*)
{
	crmTekst.Enabled := (specimen.text == "rectum")
	crm.Enabled := (specimen.text == "rectum")
}

_toggleNACT(*)
{
	dvorak.Enabled := (chemoCheckBox.value)
}

_41ButtonOK(*)
{
	if InStr("submucosa , muscularis propria , omliggend vetweefsel , serosa, met serosale doorbraak", invasiediepte.text)
		invasiediepte.text := "Invasief tot in de " . invasiediepte.text . "."
	if (gradering.text == "niet van toepassing")
		graderingbesluit := ", gradering niet van toepassing."
	Else
		graderingbesluit := ", graad " . gradering.text . "."  

	if (crm.text != "") && (specimen.text == "rectum")
		crm.text := "-Circumferentiele resectiemarge: " . crm.text " mm.<br>"
	dvorakbes := ""
	dvorakmic := ""
	if chemoCheckBox.value = 1
		{
			dvorakmic := "-Rectal cancer regression grade (RCRG) na therapie (volgens Dworak et al, 1997):" . dvorak.text . "<br>"
			dvorakbes := "-Dvorak cancer regression grade (DCRG): " SubStr(dvorak.text, 1, 4) ".<br>"
		}
	
	html := 
	(
		"<b>Microscopie:</b><br>"
		"Invasief carcinoom: <br>"
		"-Histologisch type (op basis van WHO classificatie): " . typeCarcinoom.text . "<br>"
		"-Histologische gradering: graad " gradering.text "<br>"
		"-Invasiediepte: " invasiediepte.text "<br>"
		"-Lymfovasculaire invasie: " . lvi.text . "<br>"
		"-Perineurale invasie: " . pni.text . "<br>"
		dvorakmic
		"-Chirurgische snedevlakken (proximaal & distaal): " . sv.text . "<br>"
		crm.text
		"-Tumor budding: " . budding.text . "<br>"
		"-Extramurale tumordeposits: " . extramuraleDeposits.text . "<br>"
		"-Lymfeklieren: " . totaalLk.text . " lymfeklieren gepreleveerd, waarvan " . positiefLk.text . " positief.<br>"
		"-Precursorletsel: " . precursorLetsel.text . "<br>"
		"<br>"
		"<b>Besluit:</b><br>"
		"Partiele colectomie (" . specimen.text . "):<br>"
		"-Op " . afstandSnedevlak.text . " cm van het dichtst bijgelegen snedevlak: tumoraal letsel met maximale diameter van " . diameter.text . " mm.<br>"
		"-Microscopisch overeenstemmend met een " . typeCarcinoom.text . graderingbesluit . "<br>"
		"-Invasiediepte: " . invasiediepte.text . "<br>"
		"-Lymfovasculaire invasie: " . lvi.text . "<br>"
		dvorakbes
		"-Snedevlakken (proximaal & distaal): " . sv.text . "<br>"
		crm.text
		"-Lymfklieren: " . totaalLk.text . " lymfeklieren gepreleveerd, waarvan " . positiefLk.text . " positief.<br>"
		"-Voorstel tot stadiëring (volgens TNM8): " . tnm.text . "<br>"
	)
	SendHTML(html,aw)
	MyGui.Destroy()
}
}

