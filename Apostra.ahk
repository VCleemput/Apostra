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
::*pdl1::
{
    aw := WinExist("A")
    MyGui := Gui(, "PDL1")
    Orgaan_tekst := MyGui.AddText("xm section w200", "Orgaan")
    Matrix := MyGui.AddDropDownList("ys w500 Choose1", ["Blaas", "Cervix"])
    ScoreText := MyGui.AddText("xm section w200", "Score:")
    ScoreEdit := MyGui.AddEdit("ys w200")
    ScoreType := MyGui.AddText("xm section w200", "Score Type:")
    ScoreTypeDropdown := MyGui.AddDropDownList("ys w200", ["TPS", "CPS"])
    ExternalControlsText := MyGui.AddText("xm section w200", "Externe/interne controles:")
    ExternalControlsCheckbox := MyGui.AddCheckbox("xm section w200", "Controles OK")

    MyGui.AddButton("xm w50 h20 default", "OK").OnEvent("click", _PDL1ButtonOK)
    MyGui.Show()
    Return


_PDL1ButtonOK(*)
{
   	result := SetScoresAndCheckPositivity(Matrix.text, ScoreTypeDropdown.text, ScoreEdit.text, ExternalControlsCheckbox.value)
    SendHTML(result, aw)
    MyGui.Destroy()
}

SetScoresAndCheckPositivity(organ, scoreType, enteredScore, externalControlsOK)
{
    ; Initialize variables for CPS and TPS thresholds for each organ
    cpsThreshold := 10  ; Default CPS threshold
    tpsThreshold := 1   ; Default TPS threshold

    ; Set thresholds based on the selected organ
    switch (organ) {
        case "Blaas":
            cpsThreshold := 10  ; Threshold for CPS for Blaas
            tpsThreshold := 1  ; Threshold for TPS for Blaas
            cpsInterpretatie := "Combined positivity score (CPS): The number of PD-L1 staining cells (tumor cells and immune cells) divided by the total number of viable tumor cells multiplied by 100 (= score). For treatment with Pembrolizumab, the cutoff value is > or = 10."
            tpsInterpretatie := "Tumor Proportion Score (TPS): The number of PD-L1 staining tumor cells divided by the total number of viable tumor cells (= percentage). For treatment with Nivolumab, the cutoff value is > or = 1%."
        case "Cervix":
            cpsThreshold := 1  ; Threshold for CPS for Cervix
            cpsInterpretatie := "Combined positivity score (CPS): The number of PD-L1 staining cells (tumor cells and immune cells) divided by the total number of viable tumor cells multiplied by 100 (= score). For treatment with Pembrolizumab, the cutoff value is > or = 1."
    }

    ; Check if the entered score is a number
    if IsNumber(enteredScore) {
        enteredScore := Round(enteredScore) ; Round the entered score to an integer

        ; Check if CPS is positive
        if (scoreType = "CPS") {
            if (enteredScore >= cpsThreshold) {
                result := "CPS is positive"
            } else {
                result := "CPS is negative"
            }
            explanation := cpsInterpretatie
        }
        ; Check if TPS is positive
        else if (scoreType = "TPS") {
            if (enteredScore >= tpsThreshold) {
                result := "TPS is positive"
            } else {
                result := "TPS is negative"
            }
            explanation := tpsInterpretatie
        }

        ; Include the external controls status in the HTML
        if (externalControlsOK = "Yes") {
            externalControlsStatus := "Externe/interne controles zijn opgegaan"
        } else {
            externalControlsStatus := "Externe/interne controles zijn niet opgegaan"
        }

        ; Generate the HTML output 
        html := "<b><u>Aanvullende immuunhistochemie voor PD-L1 (<@DATE@>/<@USERLONG@>):</u></b><br>"
        html .= "<b>Orgaan:</b> " organ "<br>"
        html .= "<b>Techniek:</b> uitgevoerd met 22C3 antilichaam (Agilent) op het Benchmark Ultra toestel (Roche)<br>"
        html .= "<b>Score Type:</b> " scoreType "<br>"
        html .= "<b>Interpretatie:</b> " explanation "<br>"
        html .= "<b>Externe/interne controle:</b> " externalControlsStatus "<br>"
        html .= "<b>Score:</b> " enteredScore "<br>"
        html .= "<b>Result:</b> " result "<br>"
		return html
    } else {
        return "Invalid score. Please enter a number."
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

