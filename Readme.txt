---------------------------------------------------------------------------
----------------------  ScintillaVB MANUAL   ------------------------------
----------------------     Version 0.1       ------------------------------
---------------------------------------------------------------------------


ScintillaVB is a Visual Basic wrapper for the Scintilla editing control. 

And, what is Scintilla?
Here is an excerpt from its documentation: "Scintilla is a free source code
editing component. As well  as  features  found  in  standard  text editing
components, Scintilla includes features  especially useful when editing and
debugging  source code. These  include support  for syntax  styling,  error
indicators,  code  completion and call tips.  The selection margin can con-
tain markers like those  used in  debuggers to indicate breakpoints and the
current line.  Styling  choices  are  more  open  than  with  many editors,
allowing the use  of  proportional fonts,  bold and italics, multiple fore-
ground and background colours and multiple fonts."

ScintillaVB it's a  Windows  ActiveX control (ocx). Its only function is to 
serve as interface between any ActiveX capable  program and  the  Scintilla
control  (thus, the word "wrapper").It's  still  alpha  code,  so there are 
a lot of things to do,  and  possibly bugs!;).  Feel  free to send feedback 
to JM Rodriguez (jmrodriguez@sourceforge.net).


Requirements. 
---------------------------------------------------------------------------
ScintillaVB needs to proper operation:
	-SciLexer.dll, v 1.44 or higher. It must be on an  accesible  place
	 (in your path).
	-The file lexer.xml (or yor own XML conf file; it must  agree  with
	 the embebbed DTD in lexer.xml).
	-Microsoft Windows Common Controls.
	-Microsoft Windows Common Dialog Control.
	-Microsoft Scripting Runtime.
	-Microsoft XML, version 2.
	
The code works with Visual Basic 5 CCE  or  Visual Basic 6. Due to the sub-
classing calls needed to handle Sicintilla's notification, Visual Basic IDE
can crash, so save yor work often (or turn off subclassig while debugging).


Properties and methods related to call tips
---------------------------------------------------------------------------
Properties:
	-CallTipActive: returns true if there is a call tip visible,  false
			if not. Read-only. Boolean.

Methods:
	-CallTipCancel: hides a call tip if one was visible.
			Arguments: none.
	
	-ShowCallTip (tip [,hltstart, hltend]):	show   a  call  tip, at the
			current caret position.
			Arguments: 
			-tip: the string shown in the call tip (string).
			-hltstart: begin of the highlighted portion of  the
			tip (number-long). Optional.
			-hltend: end of the highlight (number-long). Optio-
			nal.


Properties and methods related to text search
---------------------------------------------------------------------------
Properties:			
	-SearchMatchCase: current value  of the match  case  search  flag. 
			Read-write. Boolean.
	
	-SearchWholeWord: current value  of the whole  word  search  flag.
			Read-write. Boolean.
			   
	-SearchWordStart: current value  of the word  start  search  flag.
			Read-write. Boolean.
			   
	-SearchRegExp:	current value  of the regular expression  search
			flag. Read-write. Boolean.

Methods:
	-find (txttofind [, findinrng]): search for the  string 'txttofind'
			within the text of the control, or, if  'findinrng'
			is true, within  the range of text currently selec-
			ted. Returns  the length  of the range  found or -1
			if not found.


Properties and methods related to appearance
---------------------------------------------------------------------------
Properties:
	-WhiteSpaceVisible: there are three  possibilities:  
			SWCS_INVISIBLE (white spaces  are  invisible)
			SWCS_VISIBLEALWAYS (white spaces are visible) 
			SWCS_VISIBLEAFTERINDENT (white  spaces are  visible
				but indentation arrows are not).
			Read-write. Enum (SWCS). 
			Example: 
			Scintilla1.WhiteSpaceVisible = SWCS_VISIBLEALWAYS
	
	-EOLVisible:	if the EndOfLine mark is visible or not. Read-write
			Boolean.
			
			
	-BackColor:	the  color of the  back of the control. Defaults to
			white.
	
	-ForeColor:	the fore color of the control. Defaults to black.
		
	-CallTipBackColor: back color of the call tipo square. Defaults  to
			white.
			
	-HScrollBarVisible: horizontal  scroll  bar visibility. Read-write.
			Boolean.
	
	-IndGuidesVisible: indentation guides visibility. Read-write. 
			Boolean.
			
	-LineNumbers:	with this  property  set to true, Scintilla show  a
			margin on the left than contains the number of each
			line. Read-write. Boolean.


Properties and methods related to text 
---------------------------------------------------------------------------
Properties:
	-Text:		the text of the control. Read-write. String.
	
	-TotalLines:	returns the number of lines of text in the control.
			Read only. Long.
			
	-EndOfLine:	end of line mode: SC_EOL_CRLF  (default), SC_EOL_CR
			or SC_EOL_LF. Read-write. Enum (EOL).
			
	-CanPaste:	will a paste succeed?. Read only. Boolean.
	
	-CanUndo:	are there any undoable actions in the undo history?
			Read only. Boolean.
	
Methods:
	-paste:		paste the  contents of the clipboard into the docu-
			ment replacing the selection.
	
	-copy:		copy the selection to the clipboard.
	
	-cut:		cut the selection to the clipboard.
	
	-undo:		undo one action in the undo history.
	
	-redo:		redoes the next action on the undo history.
	
	-CharAtPos (position): returns the ASCII value of the character  in
			'position' (long).
			
	-LoadFile (file): replaces the text of the control with the content
			of the file 'file' (this argument is a string, full
			path name of a file).
	-ConvertEOLs (newEOL): convert all line endings in the document  to
			'newEOL' mode.


Properties and methods related to position
---------------------------------------------------------------------------
Properties:
	-CurLine: 	get or set the current line. Read-write. Long.
	
	-CurPosition:	get or set the current caret position.  Read-write.
			Long.
	
	-Column:	get the current column. Read only. Long.


Properties and methods related to auto completion lists
---------------------------------------------------------------------------
The auto completion list support it's not complete. Scintilla has more pro-
perties related to this feature. So parts of this section goes to TODO bag.

Properties:
	-AutoCHideIfNoMatch: if the list must be  cancelled  if  there's no
			match in user's typing. Read-write. Boolean.
	
	-AutoCSeparatorChar: get or set the character used as separator  in
			the string list that will  show de  auto-compeltion
			list. Defaults to space. Read-write. String. 

Methods:
	-ShowAutoCList(prevchars, autolist):  display   a   auto-completion
			list. The 'prevchars' parameter  indicates how many
			characters before the caret  should be used to pro-
			vide context. 'autolist' must contain a string list
			of elements (separated by  spaces  by default) that
			will compose the auto-completion list.
	
	-HideAutoCList: hide an auto-completion list.


Other properties
---------------------------------------------------------------------------
	-MatchBraces: 	get the state of or activates  the  highlight  mat-
			ching braces feature.  Read-write.  Boolean. If the
			indentation guides are visible, are  also highligh-
			ted.
			

Styling text
---------------------------------------------------------------------------
One of the most interesting  and powerful features in Scintilla is the syn-
tax stiling. Each language supported has his own  styling. SciLexer DLL has
styling support for a lot of languages, including C and C++, Pascal, Visual
Basic, Java, etc. Each  language  has a set of styles, which  indicates how
Scintilla styles the document; so a language could have a normal or default
style, a special  style for keywords (words with special meaning),  another 
for operators, one more  for strings...  There are  languages that  use few 
styles, and others more than 30, like HTML, due to the support for embebbed 
languages. 

The current text inside Scintilla can be styled by choosing a language. You
can choose a language with ScintillaVB by modifing its "language"  property,
for example: Scintilla1.language = "HTLM". If the language it's not present
in the conf file, an error will be raised.

Each style has several properties: 
	-font (including face, size and bold | italic | underline attrs.)	
	-forecolor
	-backcolor
	-eolfilling (if the backcolor extends to the end of line)
	-visibility-
	-case (upper | lower)
	-charset (not currently supported by ScintillaVB)

You  can  change  (currently only in design mode)  any of  this properties,
by  choosing the "Custom" tab for the ScintillaVB control (in  Visual Basic
is placed into the properties panel). All the language and styling informa-
tion is stored into an XML file, called by default lexer.xml (this  can  be
changed by using the "ConfFile" property).


Version 0.1. April 2000.
Juan Manuel Rodriguez.
jmrodriguez@sourceforge.net
		