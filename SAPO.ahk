#Requires AutoHotkey v2.0
#Include WinClipAPI.ahk ;includes clipboard functionality
#Include WinClip.ahk

; define a global variable to drop the generated text to, in case something goes wrong
global backup
backup := ""

SendHTML(html, aw := WinExist("A")) 
{
	;function to send the html to the clipboard and paste it
    global backup
	wc := Winclip()
	html := '<div style="font-size:10pt">' html  '</div>'
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

::*69pb:: 
{ 	;CNB borst
	aw := WinExist("A")
	MyGUI := Gui(,"CNB borst")
	MyGUI.AddText("xm w200 section", "Lateraliteit")
	lateraliteit := MyGUI.AddComboBox("ys w200", ["links","rechts","lateraliteit niet gegeven"])
	MyGUI.AddText("xm w200 section", "Type carcinoom")
	typeCarcinoom := MyGUI.AddComboBox("ys w200 Choose1", ["Invasief carcinoom NST (ductaal)","Invasief lobulair carcinoom","Mucineus carcinoom","Tubulair carcinoom","Metaplastisch carcinoom"])
	MyGUI.AddText("xm w200 section", "Glandulair")
	glandulair := MyGUI.AddDDL("ys w200 Choose2 AltSubmit", ["Score 1, >75% klierbuisformatie","Score 2, 10-75% klierbuisformatie","Score 3, <10% klierbuisformatie"])
	MyGUI.AddText("xm w200 section", "Kern")
	kern := MyGUI.AddDDL("ys w200 Choose2 AltSubmit", ["Score 1, kleine, uniforme kernen","Score 2, matige kernvariabiliteit","Score 3, grote, sterk variabele kernen"])
	MyGUI.AddText("xm w200 section", "Mitose score")
	mitose := MyGUI.AddDDL("ys w200 Choose1 Altsubmit", ["Score 1","Score 2","Score 3"])
	MyGUI.AddText("xm w200 section", "CIS")
	CIS := MyGUI.AddCheckbox("ys", "CIS?")
	CIS.OnEvent("Click", _CISButton)
	typeCistekst := MyGUI.AddText("xm+30 section w200 Disabled", "Type CIS")
	typeCis := MyGUI.AddComboBox("ys w200 Disabled Choose1", ["ductaal","lobulair"])
	graderingtekst := MyGUI.AddText("xm+30 section w200 Disabled", "Gradering")
	graderingCis := MyGUI.AddComboBox("ys w200 Disabled Choose2", ["graad 1","graad 2","graad 3"])
	groeipatroontekst := MyGUI.AddText("xm+30 section w200 Disabled", "Groeipatroon")
	groeipatroon := MyGUI.AddListBox("ys w200 r7 Multi Disabled", ["cribriform","solied","papillair","comedo","clinging","pagetoid","uitbreidend in adenosis"])
	MyGUI.AddText("xm section w200 h20", "Tumorload")
	tumorload := MyGUI.AddEdit("ys w200", "")
	MyGUI.AddText("xm section w200 h20", "LVI")
	lvi := MyGUI.AddCheckbox("ys", "")
	MyGUI.AddText("xm section w200 h20", "PNI")
	pni := MyGUI.AddCheckbox("ys", "")
	MyGUI.AddText("xm section w200 h20", "Necrose")
	necrose := MyGUI.AddCheckbox("ys", "")
	MyGUI.AddText("xm section w200 h20", "microcalcificaties")
	microcalcificaties := MyGUI.AddCheckbox("ys", "")
	MyGUI.AddText("xm section w200 h20", "TILs")
	tils := MyGUI.AddEdit("ys w200", "")
	MyGUI.AddButton("xm section w50 h20 Default", "OK").OnEvent("Click", _69pbButtonOK)
	MyGUI.Show()
	
_CISButton(*){
	if CIS.value = 1
	{
		typeCistekst.Enabled := 1
		typeCis.Enabled := 1
		graderingtekst.Enabled := 1
		graderingCis.Enabled := 1
		groeipatroontekst.Enabled := 1
		groeipatroon.Enabled := 1
	}
	if CIS.value = 0
	{
		typeCistekst.Enabled := 0
		typeCis.Enabled := 0
		graderingtekst.Enabled := 0
		graderingCis.Enabled := 0
		groeipatroontekst.Enabled := 0
		groeipatroon.Enabled := 0
	}
	return
}

_69pbButtonOK(*){
	MyGUI.Hide()
	score := kern.value + mitose.value + glandulair.value
	if score <= 5
		gradering := "graad 1"
	else if score <=7
		gradering := "graad 2"
	else if score <= 9
		gradering := "graad 3"	
	bin := Map(0,"afwezig",1,"aanwezig")
	lvi :=bin[lvi.value]
	pni := bin[pni.value]
	necrose := bin[necrose.value]
	microcalcificaties := bin[microcalcificaties.value]
	if CIS.value = 1
	{
		textCISbesluit := "Eveneens " graderingCis.text " " typeCIS.text " carcinoma in situ."
		cis := "aanwezig: "
		komma := "; "
		groeipatroon := StrJoin(groeipatroon.text, ", ")
		typeCismicro := typeCis.text " carcinoma in situ"
	}
	else if CIS.value = 0
	{
		textCIS :=""
		cis := "afwezig."
		typeCIS := ""
		graderingCis.text := ""
		typeCismicro := ""
		groeipatroon := ""
		komma := ""
		textCISbesluit := ""
	}

	html :=
	(
		"<b>Maligne veranderingen:</b><br>"
		typeCarcinoom.text " - " gradering "<br>"
		"-Glandulair: score " glandulair.value "<br>"
		"-Kernpleiomorfie: score " kern.value "<br>"
		"-Mitosetelling: score " mitose.value "<br><br>"
		"In situ carcinoom: " cis " " graderingCis.text " " typeCismicro komma groeipatroon "<br><br>"
		"Tumorload: " tumorload.text "% tumorcel oppervlak op totale weefseloppervlak.<br>"
		"Lymfovasculaire invasie: " lvi ".<br>"
		"Perineurale invasie: " pni ".<br>"
		"Necrose: " necrose ".<br>"
		"Microcalcificaties: " microcalcificaties ".<br>"
		"TIL's: " tils.text "%.<br><br>"
		"<b>Besluit:</b><br>"
		"CNB borst " lateraliteit.text ": B5: Maligne: " typeCarcinoom.text ", " gradering ". " textCISbesluit " Predictieve merkers volgen."
	) 

	SendHTML(html, aw)
	MyGUI.Destroy()
	return
}
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
::*pd::
{	;PD-L1
	aw := WinActive("A")
	MyGui := Gui(,"PD-L1")
	techniek_tekst := MyGui.AddText("xm section w200", "Techniek")
	techniek := MyGui.AddCombobox("ys w200 Choose1", IniListRead("lab-variables.ini", "PD-L1", "techniek"))
	toestel_tekst := MyGui.AddText("xm section w200", "Toestel")
	toestel := MyGui.AddCombobox("ys w200 Choose1", IniListRead("lab-variables.ini", "PD-L1", "toestel"))
	;drug_tekst := MyGui.AddText("xm section w200", "drug")
	;drug := MyGui.AddCombobox("ys w200", IniListRead("lab-variables.ini", "PD-L1", "drug"))
	interpretatie_tekst := MyGui.AddText("xm section w200", "Interpretatie")
	interpretatie := MyGui.AddDDL("ys w200 Choose1", ["CPS", "TPS", "IC"])
	;cutoff_tekst := MyGui.AddText("xm section w200", "cutoff")
	;cutoff := MyGui.AddEdit("ys w200", "")
	;controles_tekst := MyGui.AddText("xm section w200", "controles geslaagd")
	;controles := MyGui.AddCheckbox("ys w200 Checked", "")
	;geschikt_tekst := MyGui.AddText("xm section w200", "geschikt voor evaluatie")
	;geschikt := MyGui.AddCheckbox("ys w200 Checked", "")
	;MAT_tekst := MyGui.AddText("xm section w200", "membranaire aankleuring tumorcellen")
	;MAT := MyGui.AddDDL("ys w200 Choose2", ["geen", "geringe", "uitgebreide"])
	;MAI_tekst := MyGui.AddText("xm section w200", "aankleuring immuuncellen")
	;MAI := MyGui.AddDDL("ys w200 Choose2", ["geen", "geringe", "uitgebreide"])
	;TAI_tekst := MyGui.AddText("xm section w200", "type aankleuring immuuncellen")
	;TAI := MyGui.AddDDL("ys w200 Choose1", ["membranaire en cytoplasmatische", "membranaire", "cytoplasmatisch"])
	score_tekst := MyGui.AddText("xm section w200", interpretatie.text . "-score")
	score := MyGui.AddEdit("ys w200", )
	interpretatie.OnEvent("change", (*) => (score_tekst.text := interpretatie.text . "-score"))
	;resultaat_tekst := MyGui.AddText("xm section w200", "Resultaat")
	;resultaat := MyGui.AddDDL("ys w200", ["positief", "negatief"])
	MyGui.AddButton("xm w50 h20 default", "OK").OnEvent("click", _PDL1ButtonOK)
	MyGui.Show("AutoSize")

_PDL1ButtonOK(*)
{
	html := "<b><u>Resultaat PD-L1 IHC analyse</u></b><br>"
	html .=	"Techniek: uitgevoerd met " . techniek.text . " op " . toestel.text . ".<br>"
	html .= "<b>Interpretatie:</b><br>"
	if interpretatie.text = "CPS"
		html .= "Combined positivity score (CPS): het aantal PD-L1 aankleurende cellen (tumorcellen en immuuncellen) gedeeld door het totaal aantal viabele tumorcellen x 100 (= score).<br>"
	if interpretatie.text = "TPS"
		html .= "Tumour Proportion Score (TPS): het aantal PD-L1 aankleurende tumorcellen gedeeld door het totaal aantal viabele tumorcellen (= percentage).<br>"
	if interpretatie.text = "IC"
		html .= "Immune cell area (IC): gebied Ingenomen door het PD-L1 aankleurende immuuncellen gedeeld door het totale tumorgebied (= percentage).<br>"
	html.= "<b>Besluit:</b><br>"
	html .= "PD-L1 (" . techniek.text . "): " . interpretatie.text . " = " . score.text
	if interpretatie.text != "CPS"
		html .= "%"
	html.= ".<br>"
	SendHTML(html, aw)
	MyGui.Destroy()
}
}
