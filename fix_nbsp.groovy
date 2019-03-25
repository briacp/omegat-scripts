/* :name=Fix Non-Breaking Spaces :description=Add non-breaking spaces in figures, before question marks, etc... */


/*
 * Source :
 * http://www.druide.com/enquetes/pour-des-espaces-ins%C3%A9cables-impeccables
 * http://bdl.oqlf.gouv.qc.ca/bdl/gabarit_bdl.asp?id=4566
 * http://bdl.oqlf.gouv.qc.ca/bdl/gabarit_bdl.asp?id=2039
 * 
 * 
 */

def ESPACE;

// Espace fine insécable : \u202f
ESPACE = '\u202f';

// Espace insécable      : \u00a0
ESPACE = '\u00a0';

def replaces = [
        // Before combined punctuation signs
        '\\s+([;?!:\u00b0\u00bb])':               ESPACE + '$1',
        // After opening quote
        '([\u00ab])\\s+':                         '$1' + ESPACE,
        // General rule
        '(\\d+)\\s+(\\w\\+)':                     '$1' + ESPACE + '$2',
        // 25e anniversaire Ve republique
        '\\b(\\d+|[XVMCLI]+)e\\s+(\\w\\+)':       '$1' + ESPACE + '$2',
        // After abbrev.
        '\\b(p\\.|page|num.ro|n°|livre|chapitre|chap\\.|article|art\\.|an)\\s+(\\d+)': '$1' + ESPACE + '$2',
        // 21 janvier 2014, 16 h 30, 15 heures 20
        '(\\d{2})\\s+(\\w+)\\s+(\\d{4})':         '$1' + ESPACE + '$2' + ESPACE + '$3',
        // 21 janvier 2014
        '(\\d)\\s+([+-/*=<>])\\s+(\\d)':          '$1' + ESPACE + '$2' + ESPACE + '$3',
        // Symboles SI
        '(\\d+)\\s*(h|%|lb|po|pouces|g|kg|mg|s|A|I|H|cm|m|mm|l|ml|cl)\\b': '$1' + ESPACE + '$2',
        // Money
        '\\s+(\\$|M\\$|\\$\\s*CA|\\$\\s*US|\\u20ac)': ESPACE + '$1',
        '\\$\\s*(CA|US)': '\\$' + ESPACE + '$1',
        '(Mme|Mlle|M\\.|Mgr)\\s+':          '$1' + ESPACE + '$2' + ESPACE + '$3',
        ',\\s*etc\\.?': ',' + ESPACE + 'etc.'
];

def segment_count = 0

console.println("Fix NBSP");

def allEntries = project.allEntries;
for (def i = 0; i < allEntries.size(); i++) {
    if (java.lang.Thread.interrupted()) {
      break;
    }
    def ste = allEntries.get(i);
	def source = ste.getSrcText();
	def target = project.getTranslationInfo(ste) ? project.getTranslationInfo(ste).translation : null;
	
	def initial_target = target

	// Skip untranslated segments
	if (target == null) return

	// The search_string is replaced by the replace_string in the translated text.
    replaces.each   {
        target = target.replaceAll(it.key, it.value);
	}

	// The old translation is checked against the replaced text, if it is different,
	// we jump to the segment number and replace the old text by the new one.
	// "editor" is the OmegaT object used to manipulate the main OmegaT user interface.
	if (initial_target != target) {
		segment_count++
		// Jump to the segment number
		edt {
			editor.gotoEntry(ste.entryNum())
			console.println("MOD\t" + ste.entryNum() + "\t" + ste.srcText + "\t" + target )
			// Replace the translation
			editor.replaceEditText(target)
		}
	}
}

console.println("modified_segments" + ": " + segment_count);

def edt(Closure c) {
    if (javax.swing.SwingUtilities.isEventDispatchThread()) {
        c.call(this)
    } else {
       javax.swing.SwingUtilities.invokeAndWait(c)
    }
}


