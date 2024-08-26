/* 
AutoHotKey v2.0 script for determining the Fleischner Society 2017 recommendation for pulmonary nodule 
follow-up based on natural language descriptors in a radiology report. To use, highlight a sentence
in a radiology report and press WinKey+N. It will display a message box with nodule descriptors and
copy the recommendation into the report at the end of the highlighted sentence.

This uses a heuristic approach to analyzing the language, so it may not work for all cases. Some important caveats:
-It accepts 1-3 measurements separated by "x" with optional spaces. Must have "mm" or "cm" at the end
    e.g. 1.0 x 2.0 x 3.0 cm, 1.0x2.0 mm, or 1 mm
-It will always use the single number or average of the two largest measurements for Fleischner category. 
-It will always categoirize a calcified nodule as benign and not recommend follow-up
-If the sentence contains "solid" and "ground glass" descriptors, it will consider it a part-solid nodule

Examples of well-formed sentences:

"Incidental right upper lobe solid noncalcified pulmonary nodule measuring 7 x 8 mm (series 1, image 30)."

"Multiple bilateral groundglass pulmonary nodules, largest measuring 1.5 x 1.5 cm in the left lower lobe."

"Part solid nodule in the right middle lobe with groundglass and solid components. 
    The solid component measures 6 x 4 x 8 mm."
*/

class Nodule extends Object
{
    ;default properties
	Num := "single"
	Composition := ""
	Calcified := false
    mString := "" ;this is the string containing extracted measurements
	Units := ""
	Measurements := Array() ;numerical (float) measurements used for calculation

	__New(nString)
	{
		if !(InStr(nString, "nodule") or InStr(nString, "nodules"))
			throw(-1) ;throw an error if there is no reference to a nodule in the string
				
		nString := Trim(nString, ".") ;trim period from end of sentence
		nString := StrReplace(nString, ",") ;remove commas
		parts := StrSplit(nString, A_Space) ;split string into individual components separated by spaces
	
		Loop parts.Length
		{
			;evaluate nodule number (single vs. multiple)
			if parts[A_Index] = "nodules"
				this.Num := "multiple"
	
			;evaluate nodule composition
			If parts[A_Index] = "solid"
			{
				if this.Composition = "" 
					this.Composition := "solid"
		
				if this.Composition = "ground glass" 
					this.Composition := "part solid"
			}
	
			If ((parts[A_Index] = "ground" and parts[A_Index+1] = "glass") or parts[A_Index] = "groundglass")
			{
				if this.Composition = "" 
					this.Composition := "ground glass"
				if this.Composition = "solid" 
					this.Composition := "part solid"
			}
	
			If ((parts[A_Index] = "part" and parts[A_Index+1] = "solid") or parts[A_Index] = 'part-solid')
			{
				this.Composition := "part solid"
				A_Index++
			}
	
			;Evaluate calcification
			If (parts[A_Index] = "calcified" or parts[A_Index] = "calcification" or parts[A_Index] = "calcifications")
				this.Calcified := true
	
			If (parts[A_Index] = 'noncalcified' or parts[A_Index] = 'non-calcified')
				this.Calcified := false ;any non-calcified
		}

        ;Search for measurements and units. Accepts 1-3 measurements separated by "x" with optional spaces. 
        ;Must have "mm" or "cm" at the end, e.g. 1.0 x 2.0 x 3.0 cm, 1.0x2.0 mm, or 1 mm

        needle := "i)(\d+(?:\.\d+)?)(?:\s*x\s*(\d+(?:\.\d+)?))?(?:\s*x\s*(\d+(?:\.\d+)?))?\s?([cm]m)"

        if (RegExMatch(nString, needle, &match))
        {
            this.mString := match[]
            this.Units := match[match.Count]
            Loop (match.Count-1)
                if IsNumber(match[A_Index])
                    this.Measurements.Push(Float(match[A_Index]))
        }
        else
        {
            throw(-2) ;Error if no appropriate measurements found
        }
	}

	Size() ;if there are at least two measurements, this returns the average of the two largest
	{
		if this.Measurements.Length = 1
			s := this.Measurements[1]

        if this.Measurements.Length = 2
            s := (this.Measurements[1] + this.Measurements[2]) / 2

		if this.Measurements.Length = 3
        {
            s1 := this.Measurements[1]
            s2 := this.Measurements[2]
            if this.Measurements[3] > s1
                s1 := this.Measurements[3]
            else
                if this.Measurements[3] > s2
                    s2 := this.Measurements[3]

            s := (s1 + s2) / 2
        }

		;convert to mm
		if this.Units = "cm" 
			s := s * 10

		return s
	}

	Category() ;determine fleichner 2017 category based on multiplicity, composition, and size
	{
		s := this.Size()

		if (this.Composition = "solid" or this.Composition = "")
		{
			if this.Num = "single"
			{
				if s < 6
					c := 1
				if s >= 6 and s <= 8
					c := 2
				if s > 8
					c := 3
			} else {
				if s < 6
					c := 1
				if s >= 6
					c := 6
			}
		} 
        else 
        {
			if this.Num = "multiple"
			{
				if s < 6
					c := 7
				if s >= 6
					c := 8
			} else {
				if this.Composition = "ground glass"
				{
					if s < 6
						c := 0
					if s >= 6
						c := 4
				}

				if this.Composition = "part solid"
				{
					if s < 6
						c := 0
					if s >= 6
						c := 5
				}
			}
		}

		if this.Calcified
			c := 0

		return c
	}

	Recommendation() ;Returns Fleschner recommendation category
	{
		r := ""
		Switch this.Category()
		{
		case 0: r := "No routine follow-up is indicated."
		case 1: r := "If the patient carries a high risk for lung cancer, consider follow-up CT in 12 months."
		case 2: r := "CT in 6-12 months. If the patient carries a high risk for lung cancer, recommend additional CT at 18-24 months."
		case 3: r := "CT at 3 months, PET/CT, or tissue sampling."
		case 4: r := "CT in 6-12 months to confirm persistence, then CT every 2 years until 5 years."
		case 5: r := "CT in 3-6 months to confirm persistence. If unchanged and solid component remains <6mm, annual CT should be performed for 5 years."
		case 6: r := "CT in 3-6 months. If the patient carries a high risk for lung cancer, recommend additional CT at 18-24 months."
		case 7: r := "CT in 3-6 months. If stable, consider CT at 2 and 4 years."
		case 8: r := "CT in 3-6 months. Subsequent management based on the most suspicious nodule."
		}
		return r
	}
}

#HotIf WinActive("PowerScribe 360 | Reporting")
#n:: ;WinKey + N
{
    ;get hilighted text and store in nString. Preserve current clipboard contents
	temp := A_Clipboard
    A_Clipboard := ""
	Send "^c"
	ClipWait
    nString := A_Clipboard
    A_Clipboard := temp

	try ;attempt to extract nodule data
	{
		n := Nodule(nString)
	} 
	catch Number as e
	{
		if e = -1 
            MsgBox("The highlighted text does not reference a nodule.")

        if e = -2
            MsgBox("The highlighted text does not contain measurements.")
	}
	else
	{
		MsgBox( 
			"Multiplicity: " n.Num
			"`nComposition: " n.Composition
			"`nCalcification: " (n.Calcified ? "calcified" : "noncalcified")
			"`nMeasurements (axial): " n.mString
			"`nSize (Fleischner): " Format("{:.1f} ", n.Size()) "mm"
			"`nRecommendation: " n.Recommendation()
			)
		
		SendInput "{end} Recommended management per 2017 Fleischner Society Guidelines: " n.Recommendation()
	}
}
