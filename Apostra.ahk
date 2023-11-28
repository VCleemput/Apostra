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
	html := '<div style="font-size:10pt">' html  '</div>'
	;OldClipboard := WinClip.Snap() ;om de een of andere reden werkt 'm zonder dit niet in WORD
	OldClipboard := ClipboardAll()
	A_Clipboard :=""
	wc.Clear()
	Sleep(100)
	wc.SetText(html)
	wc.SetHTML(html)
	if A_Clipboard = ""
		{
		wc.SetText(html)
		wc.SetHTML(html)
		}
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
	StaalTekst := MyGui.AddText("xm section vtekstStaal", "Staaltype")
	Staal := MyGui.AddDDL("ys wp Choose1",["CNB", "Operatiespecimen", "Andere:"] )
	ERCheck := MyGui.AddCheckbox("xm section vReceptorER Checked", "ER")
	ERCheck.OnEvent("Click", _toggleqs)
	ERpctTekst := MyGui.AddText("xm section vtekstER", "Percentage gekleurde kernen")
	ERpct := MyGui.AddEdit("ys wp", "")
	ERintTekst := MyGui.AddText("xm section vtekstIntER", "Intensiteit aankleuring")
	ERint := MyGui.AddDDL("ys wp Choose4",["afwezig", "zwak", "matig", "sterk"] )
	PRCheck := MyGui.AddCheckbox("xm vReceptorPR Checked", "PR")
	PRCheck.OnEvent("Click", _toggleqs)
	PRpctTekst := MyGui.AddText("xm section vtekstPR", "Percentage gekleurde kernen")
	PRpct := MyGui.AddEdit("ys wp", "")
	PRintTekst := MyGui.AddText("xm section vtekstIntPR", "Intensiteit aankleuring")
	PRint := MyGui.AddDDL("ys wp Choose4",["afwezig", "zwak", "matig", "sterk"] )
	HER2Check := MyGui.AddCheckbox("xm section vHER2R Checked", "HER2")
	HER2Check.OnEvent("Click", _toggleqs)
	HER2ind := MyGui.AddDDL("ys wp Choose1", ["borst", "maag", "andere"])
	HER2ihcTekst := MyGui.AddText("xm section vtekstHER2", "IHC score")
	HER2ihc := MyGui.AddDDL("ys w200 Choose1", ["HER2 negatief (score 0)", "HER2 negatief (score 1+)", "HER2 equivocal (score 2+)", "HER2 positief (score 3+)"])
	kiCheck := MyGui.AddCheckbox("xm vki Checked", "Ki67")
	kiCheck.OnEvent("Click", _toggleqs)
	kiscoreTekst := MyGui.AddText("xm section vtekstki", "Ki67-index (zonder %)")
	kiscore := MyGui.AddEdit("ys wp", "")
	MyGui.AddText("section xm", "Nota: na de sneltekst kan een synopsis geplakt worden met 'Win+s'")
	MyGui.AddButton("xm default", "OK").OnEvent("Click", _qsButtonOK)
	MyGui.Show()

_toggleqs(*)
{
	checks := [ERCheck, PRCheck, HER2Check, kiCheck]
	params := [[ERint, ErintTekst, ERpct, ERpctTekst], [PRint, PRintTekst, PRpct, PRpctTekst], [HER2ihc, HER2ind, HER2ihcTekst], [kiscore, kiscoreTekst]]
	for i, check in checks
		{
			for j in params[i]
				{
					j.Enabled := check.value
				}
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
		allredTekst := "Allred score " Pe " + " value " = " qs "/8."
		if ((Pct = "< 1"))
			waarde := "Negatief (minder dan 1% nucleaire positiviteit)."
		else if Pct < 1
			waarde := "Negatief (minder dan 1% nucleaire positiviteit)."
		else if Pct <=10
			waarde := "Gering positief (1-10%)."
		else
			{
				waarde := "Positief (>10%)."
				PctTekst := ""
				if Pct >= 11 and Pct <= 20
					PctTekst := "11-20%"
				else if Pct >= 21 and Pct <= 30
					PctTekst := "21-30%"
				else if Pct >= 31 and Pct <= 40
					PctTekst := "31-40%"
				else if Pct >= 41 and Pct <= 50
					PctTekst := "41-50%"
				else if Pct >= 51 and Pct <= 60
					PctTekst := "51-60%"
				else if Pct >= 61 and Pct <= 70
					PctTekst := "61-70%"
				else if Pct >= 71 and Pct <= 80
					PctTekst := "71-80%"
				else if Pct >= 81 and Pct <= 90
					PctTekst := "81-90%"
				else if Pct >= 91 and Pct <= 100
					PctTekst := "91-100%"
				waarde .= " " . PctTekst . " kleurt " . Int . " aan."
	
			}
		if (Pe = 0 or value = 0)
			{
				tekst := "Negatief. Geen aankleuring in de laesionele cellen."
				allredTekst := "Allred score 0/8."
			}
		tekst := waarde . " " . allredTekst
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
		ReceptortekstER := "Resultaat ER (" . Staal.text . "): " w "<br>"
		ABtekstER := ab_ER . ", "
		RegExMatch(w, "\d\/8", &z)
		synopsis := synopsis . "ER " . z[] . "; "
	}
	if PRCheck.value = 1
	{
		w := QuickScore(PRpct.text, PRint.text)
		ReceptortekstPR := "Resultaat PR (" . Staal.text . "): " w "<br>"
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
			if (((HER2ind.text = "borst") or (HER2ind.text = "maag")) and ((HER2ihc.value = 3) or (HER2ihc.value = 4))) {
				SishVraagTekst := " ISH volgt."
			}
			HER2tekst := 
			(
				"Resultaat HER2-IHC (" Staal.text "): " HER2ihc.text  ". " SishVraagTekst "<br>"
				"<small>" . ab_HER2 . " op " . stainer . ". Interpretatie volgens de ASCO/CAP guidelines " ASCOjaar ".</small><br><br>"
			)
			RegExMatch(HER2ihc.text, "\d\+", &h)
			synopsis := synopsis . "HER2 IHC score: " . h[] . "; "
		}
	if kiCheck.value = 1
		{
			kitekst := "Ki67-index: " kiscore.text "%.<br>"
			synopsis := synopsis . "ki67: " . kiscore.text . "%"
		}
	tekst :=
(
	"<b>Farmacodiagnostiek:</b><br>"
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
	groeipatroon := MyGUI.AddListBox("ys w200 r7 Multi Disabled Choose1", ["cribriform","solied","papillair","comedo","clinging","pagetoid","uitbreidend in adenosis"])
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
	interpretatie_tekst := MyGui.AddText("xm section w200", "Interpretatie")
	interpretatie := MyGui.AddDDL("ys w200 Choose1", ["CPS", "TPS", "IC"])
	score_tekst := MyGui.AddText("xm section w200", interpretatie.text . "-score")
	score := MyGui.AddEdit("ys w200", )
	interpretatie.OnEvent("change", (*) => (score_tekst.text := interpretatie.text . "-score"))
	techniek_tekst := MyGui.AddText("xm section w200", "Techniek")
	techniek := MyGui.AddCombobox("ys w200 Choose1", IniListRead("lab-variables.ini", "PD-L1", "techniek"))
	toestel_tekst := MyGui.AddText("xm section w200", "Toestel")
	toestel := MyGui.AddCombobox("ys w200 Choose1", IniListRead("lab-variables.ini", "PD-L1", "toestel"))
	MyGui.AddButton("xm w50 h20 default", "OK").OnEvent("click", _PDL1ButtonOK)
	MyGui.Show("AutoSize")

_PDL1ButtonOK(*)
{
	html := "<b><u>Resultaat PD-L1 IHC analyse</u></b><br>"
	html .=	"<b>Techniek: </b>uitgevoerd met " . techniek.text . " op " . toestel.text . ".<br>"
	html .= "<b>Interpretatie: </b>"
	if interpretatie.text = "CPS"
		html .= "Combined positivity score (CPS): het aantal PD-L1 aankleurende cellen (tumorcellen en immuuncellen) gedeeld door het totaal aantal viabele tumorcellen x 100 (= score).<br>"
	if interpretatie.text = "TPS"
		html .= "Tumour Proportion Score (TPS): het aantal PD-L1 aankleurende tumorcellen gedeeld door het totaal aantal viabele tumorcellen (= percentage).<br>"
	if interpretatie.text = "IC"
		html .= "Immune cell area (IC): gebied Ingenomen door het PD-L1 aankleurende immuuncellen gedeeld door het totale tumorgebied (= percentage).<br>"
	html.= "<b>Besluit: </b>"
	html .= "PD-L1 (" . techniek.text . "): " . interpretatie.text . " = " . score.text
	if interpretatie.text != "CPS"
		html .= "%"
	html.= ".<br>"
	accreditatie := IniRead("lab-variables.ini", "PD-L1", "accreditatie") 
	if accreditatie != ""
		html .= "<small>" . accreditatie . "</small><br>"
	SendHTML(html, aw)
	MyGui.Destroy()
}
}

::*41::
{	; Colonresectie
	aw := WinExist("A")
	MyGui := Gui(, "Colonresectie")
	MyGui.AddText("section w200", "Specimen")
	specimen := MyGui.AddComboBox("ys w200 Choose1", ["Partiële colectomie", "Lokale excisie", "TME", "APRA"])
	specimen.OnEvent("change", _toggle)
	MyGui.AddText("xs section w200", "Locatie")
	locatie := MyGui.AddComboBox("ys w200 Choose1", ["distaal colon", "colon transversum", "colon ascendens", "caecum", "sigmoid", "rectum", "locatie niet gegeven"])
	locatie.OnEvent("change", _toggle)
	MyGui.AddText("xs section w200", "Type")
	typeCarcinoom := MyGui.AddComboBox("ys w200 Choose1", ["adenocarcinoom NST","mucineus adenocarcinoom","zegelringcelcarcinoom","neuroendocrien carcinoom","plaveiselcelcarcinoom"])
	MyGui.AddText("xs section w200", "Gradering")
	gradering := MyGui.AddDDL("ys w200 Choose1", ["1","2","3","niet van toepassing"])
	MyGui.AddText("xs section w200", "Invasiediepte")
	invasiediepte := MyGui.AddCombobox("ys w200 Choose1",["intramucosaal adenocarcinoom", "submucosa", "muscularis propria", "subserosa", "adventitia", "serosa, met serosale doorbraak"])
	invasiediepte.OnEvent("change", _toggle)
	smTekst := MyGui.AddText("xs section w200", "Kikuchi")
	sm := MyGui.AddDDL("ys w200 Choose1", ["sm1","sm2", "sm3"])
	MyGui.AddText("xs section w200", "LVI")
	lvi := MyGui.AddComboBox("ys w200 Choose1", ["afwezig","aanwezig"])
	MyGui.AddText("xs section w200", "PNI")
	pni := MyGui.AddComboBox("ys w200 Choose1", ["afwezig","aanwezig"])
	MyGui.AddText("xs section w200", "SV")
	sv := MyGui.AddComboBox("ys w200 Choose1", ["tumorvrij","positief"])
	MyGui.AddText("xs section w200", "Budding")
	budding := MyGui.AddComboBox("ys w200 Choose1", ["afwezig","aanwezig: beperkt","aanwezig: matig","aanwezig:uitgesproken"])
	extramuraleDepositsTekst := MyGui.AddText("xs section w200", "Extramurale deposits")
	extramuraleDeposits := MyGui.AddComboBox("ys w200 Choose1", ["afwezig","aanwezig"])
	totaalLkTekst := MyGui.AddText("xs section w200", "Totaal # LK")
	totaalLk := MyGui.AddEdit("ys w200", "")
	positiefLkTekst := MyGui.AddText("xs section w200", "# Positieve LK")
	positiefLk := MyGui.AddEdit("ys w200", "")
	MyGui.AddText("xs section w200", "Precursorletsel")
	precursorLetsel := MyGui.AddEdit("ys w200", "Niet aantoonbaar")
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
	chemoCheckBox.OnEvent("click", _toggle)
	MyGui.AddText("xs section w100", "Dvorak Regression")
	dvorak := MyGui.AddDDL("ys w300 R4 Choose1", ["GR 4: geen tumorcellen, louter fibrose (= complete respons)", "GR 3: klein aantal tumor cellen (moeilijk te vinden met de microscoop op kleine vergroting) in een dominant fibrotisch stroma", "GR 2: klein aantal tumor cellen (makkelijk te vinden met de microscoop op kleine vergroting) in een dominant fibrotisch stroma", "GR 1: de tumor domineert op de fibrotische massa ", "GR 0: geen regressie"])
	MyGui.AddButton("xm w50 h20 default", "OK").OnEvent("click", _41ButtonOK)
	_toggle()
	MyGui.Show()
	


_toggle(*)
{
	if specimen.text == "APRA" or specimen.text == "TME"
		locatie.text := "rectum"
	for i in [extramuraleDeposits, extramuraleDepositsTekst, totaalLk, totaalLkTekst, positiefLk, positiefLkTekst]
		i.Enabled := (specimen.text !== "Lokale excisie")
	sm.Enabled := ((invasiediepte.text == "submucosa") and (specimen.text == "Lokale excisie") and (locatie.text == "rectum"))
	smTekst.Enabled := sm.Enabled
	crmTekst.Enabled := ((locatie.text == "rectum") and (specimen.text != "Lokale excisie"))
	crm.Enabled := crmTekst.Enabled
	dvorak.Enabled := (chemoCheckBox.value)
}

_41ButtonOK(*)
{
	if InStr("submucosa , muscularis propria , subserosa, adventitia, serosa, met serosale doorbraak", invasiediepte.text)
		invasiediepte.text := "Invasief tot in de " . invasiediepte.text 
	if (gradering.text == "niet van toepassing")
		graderingbesluit := ", gradering niet van toepassing."
	Else
		graderingbesluit := ", graad " . gradering.text . "."  

	if (crm.text != "") && (locatie.text == "rectum")
		crm.text := "-Circumferentiele resectiemarge: " . crm.text " mm.<br>"
	dvorakbes := ""
	dvorakmic := ""
	if chemoCheckBox.value = 1
		{
			dvorakmic := "-Rectal cancer regression grade (RCRG) na therapie (volgens Dworak et al, 1997):" . dvorak.text . "<br>"
			dvorakbes := "-Dvorak cancer regression grade (DCRG): " SubStr(dvorak.text, 1, 4) ".<br>"
		}
	resectieTekst := ""	
	lkTekst := ""
	if extramuraleDeposits.Enabled
		{
			resectieTekst := "-Extramurale deposits: " . extramuraleDeposits.text . "<br>"
			lkTekst := "-Lymfeklieren: " . totaalLk.text . " lymfeklieren gepreleveerd, waarvan " . positiefLk.text . " positief.<br>"
		}
	smbesluit := ""
	if sm.Enabled
		smbesluit := ", Kikuchi level " . sm.text . "."
	html := 
	(
		"<b>Microscopie:</b><br>"
		"-Histologisch type (op basis van WHO classificatie): " . typeCarcinoom.text . "<br>"
		"-Histologische gradering: graad " gradering.text "<br>"
		"-Invasiediepte: " . invasiediepte.text . smbesluit . "<br>"
		"-Lymfovasculaire invasie: " . lvi.text . "<br>"
		"-Perineurale invasie: " . pni.text . "<br>"
		dvorakmic
		"-Chirurgische snedevlakken (proximaal & distaal): " . sv.text . "<br>"
		crm.text
		"-Tumor budding: " . budding.text . "<br>"
		resectieTekst
		lkTekst
		"-Precursorletsel: " . precursorLetsel.text . "<br>"
		"<br>"
		"<b>Besluit:</b><br>"
		specimen.text . " (" . locatie.text . "):<br>"
		"-Op " . afstandSnedevlak.text . " cm van het dichtst bijgelegen snedevlak: tumoraal letsel met maximale diameter van " . diameter.text . " mm.<br>"
		"-Microscopisch overeenstemmend met een " . typeCarcinoom.text . graderingbesluit . "<br>"
		"-Invasiediepte: " . invasiediepte.text . "<br>"
		"-Lymfovasculaire invasie: " . lvi.text . "<br>"
		dvorakbes
		"-Snedevlakken (proximaal & distaal): " . sv.text . "<br>"
		crm.text
		lkTekst
		"-Voorstel tot stadiëring (volgens TNM8): " . tnm.text . "<br>"
	)
	SendHTML(html,aw)
	MyGui.Destroy()
}
}

