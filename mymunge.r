REBOL [
	Title:		"Munge function"
	Owner:		"Ashley G Truter"
	Version:	1.0.8
	Date:		7-Feb-2015
	Purpose:	"Extract and manipulate tabular values in blocks, delimited files and SQL Server tables."
	Tested: {
		2.7.8.3.1	R2			rebol.com
		2.101.0.3.1	R3 Alpha	rebolsource.net
		3.0.91.3.3	R3 64-bit	atronixengineering.com
	}
	Usage: {
		cols?		Number of columns in a delimited file
		rows?		Number of rows in a delimited file
		sheets?		Number of sheets in an XLS file
		execute		Execute a SQL statement (SQL Server, Access or SQLite)
		load-dsv	Loads delimiter-separated values from a file
		munge		Load and / or manipulate a block of tabular (column and row) values
		read-pdf	Reads from a PDF file
		read-xls	Reads from an XLS file
		sqlcmd		Execute a SQL Server statement
		unzip		Uncompress a file into a folder of the same name
		worksheet	Add a worksheet to the current workbook
	}
	Licence:	"MIT. Free for both commercial and non-commercial use."
	History: {
		1.0.0	Initial release
		1.0.1	Renamed load-block to load-dsv
				Added new sqlcmd func
				Merged Excel funcs into new worksheet func
		1.0.2	Added /count
		1.0.3	Fixed /unqiue and /count
				Added /sum
		1.0.4	Added clean-path to load-dsv and worksheet
				Added /where integer! support (i.e. RowID)
				Added lookup, index funcs
				Added /merge
		1.0.5	Minor speed improvements
				Fixed minor merge bug
				Fixed /part make block! bug
				Removed string! as a data and save option
				Added /order
				Added /headings to sqlcmd
		1.0.6	SQL Server 2012 fix
				Refactored load-dsv based on csv-tools.r
				Added /max and /min
				Added /having
				Fixed /save to handle empty? buffer
				load-dsv now handles xls variants (e.g. xlsx, xlsm, xml, ods)
				Fixed bug with part/where/unique
				Added /compact
				Added console null print protection prior to all calls
				Added read-pdf
				Added read-xls
		1.0.7	Compatibility patches
					to-error		does not work in R3
					remove-each		R3 returns integer
					select			R2 /skip returns block
					unique			/skip broken
				Minor changes to work with R3
					read (R3 returns a binary)
					delete/any (not supported in R3)
					find/any (not working in R3)
					read/lines/part (not working in R3)
					call/show (not required or supported in R3)
					call/shell (required in R3 for *.vbs)
					call %file (call form %file works in R3)
				Removed /unique
				Added column name support
				Added /headings
				Added /save none target to return lines
				Merged /having into /group
				worksheet changes
					Removed columns argument
					Removed /widths and /footer refinements
					Added spec argument
					Added support for date and auto cell types
		1.0.8	Replaced to-error with cause-error
				Replaced func with funct
				Added execute function
				Added MS Access support to execute
				Added SQLite support to execute
				Added /only
				Added spec none! support
				Added /save none! support
				Fixed /merge bug
				Fixed sqlcmd /headings/key bug
				Added cols? function
				Added rows? function
				Added sheets? function
				Fixed to work with R3 Alpha (rebolsource.net)
				Added load-dsv /blocks
				Fixed delete/where (missing implied all)
				Added unzip
		1.0.9	Tom added /group 'avg and 'collect
				Tom added to /group action blocks
				Tom removed compose/deep from group
				Tom removed "flip" from group actions
		1.0.10	Tom added rowid comparison to /where (e.g. [rowid < 10] or [find [1 10 20 30] rowid])
		1.0.11	Tom Changed /update
				Added references to columns (eg c1...) and rowid. Functions may be used (eg [ajoin [rowid ") " c1]])
	}
]

;
;	Compatibility patches
;

if integer? remove-each i [][][		; R3 fix
	*remove-each: :remove-each
	remove-each: func ['word data body][also data *remove-each :word data body]
]

if block? select/skip [0 0] 0 2 [	; R2 fix
	*select: :select
	select: func [series value /skip size][
		either skip [
			all [value: *select/skip series value size first value]
		][*select series value]
	]
]

if "1a" = unique/skip "1a1b" 2 [	; R2/R3 fix
	*unique: :unique
	unique: funct [set /skip size][
		either skip [
			row: make block! size
			repeat i size [append row to word! append form 'c i]
			skip: unset!
			do compose/deep [
				remove-each [(row)] sort/skip/all copy set size [
					either skip = reduce [(row)][true][skip: reduce [(row)] false]
				]
			]
		][*unique set]
	]
]

context [

	;
	;	Private
	;

	R3A?: 100 < second system/version
	R64?: 3 = first system/version
	R3?: any [R3A? R64?]

	XLS?: func [
		file [file!]
	][
		find [%.xls %.xml %.ods] copy/part suffix? file 4
	]

	base-path: join to-local-file system/script/path "\"

	excel-columns: none

	call-excel-vbs: func [
		file [file!]
		sheet [integer!]
		cmd [string!]
		/rows
		/sheets
	][
		any [exists? file cause-error 'access 'cannot-open file]
		write %$tmp.vbs either sheets [
			ajoin [{set X=CreateObject("Excel.Application"):X.DisplayAlerts=False:set W=X.Workbooks.Open("} to-local-file clean-path file {"):n=X.Worksheets.Count:X.Workbooks.Close:WScript.Quit n}]
		][
			ajoin [{set X=CreateObject("Excel.Application"):X.DisplayAlerts=False:set W=X.Workbooks.Open("} to-local-file clean-path file {"):} cmd {n=X.ActiveWorkbook.Worksheets(} sheet {).UsedRange.} either rows ['Rows]['Columns] {.Count:X.Workbooks.Close:WScript.Quit n}]
		]
		also excel-columns: call/wait "cmd /C $tmp.vbs" delete %$tmp.vbs
	]

	to-xml-string: func [
		string [string!]
	][
		foreach [char code][
			"<"	"&lt;"
			">"	"&gt;"
			{"}	"&quot;"
		][replace/all string char code]
	]

	;
	;	Public
	;

	set 'cols? funct [
		"Number of columns in a delimited file."
		file [file!]
		/sheet "Excel worksheet (default is 1)"
			number [integer!]
	][
		either XLS? file [
			call-excel-vbs file any [number 1] ""
		][
			;	/lines/part returns n chars in R3
			length? load-dsv first either R3? [read/lines file][read/direct/lines/part file 1] 
		]
	]

	set 'rows? funct [
		"Number of rows in a delimited file."
		file [file!]
		/sheet "Excel worksheet (default is 1)"
			number [integer!]
	][
		either XLS? file [
			call-excel-vbs/rows file any [number 1] ""
		][
			length? read/lines file
		]
	]

	set 'sheets? funct [
		"Number of sheets in an XLS file."
		file [file!]
	][
		call-excel-vbs/sheets file 0 ""
	]

	set 'execute funct [
		"Execute a SQL statement (SQL Server, Access or SQLite)."
		database [string! file!]
		statement [string!]
		/key "Columns to convert to integer"
			columns [integer! block!]
		/headings "Keep column headings"
		/raw "Do not process return buffer"
	][
		call?: "x"
		unless any [R3? empty? call?][call/show clear call?]
		stderr: make string! 256
		case [
			string? database [			; SQL Server
				call/output/error reform ["sqlcmd -S" first parse database "/" "-d" second parse database "/" "-I -Q" ajoin [{"} trim/lines statement {"}] {-W -w 65535 -s"^-"}] buffer: make string! 8192 stderr
				all [empty? stderr "Msg" = copy/part buffer 3 stderr: trim/lines buffer]
				all [R3? trim/with buffer "^M"]
				replace/all buffer "^/^/" "^/" ; sqlcmd 2012 inserts blank lines every 4k or so rows
				replace/all buffer "^/^/" "^/"
			]
			%.accdb = suffix? database [; MS Access
				any [exists? database cause-error 'access 'cannot-open database]
				write %$tmp.vbs ajoin [{set A=CreateObject("ADODB.Connection"):A.Open "Provider=Microsoft.Ace.OLEDB.12.0;Data Source=} to-local-file clean-path database {":A.Execute "} copy/part statement index? find statement " from " {into [text;database=} to-local-file path: first split-path clean-path database {].[$tmp.txt]} find statement " from " {"}]
				write schema: join path %Schema.ini ajoin ["[$tmp.txt]^/ColNameHeader=" either headings ["True"]["False"] "^/Format=TabDelimited"]
				also call/wait/shell either R64? ["$tmp.vbs"]["%windir%\sysnative\wscript $tmp.vbs"] stderr delete %$tmp.vbs
				delete schema
				either exists? file: join path %$tmp.txt [
					buffer: read/string file
					delete file
				][cause-error 'access 'cannot-open file]
			]
			true [						; SQLite
				call/output/error reform compose [{sqlite3 -separator "^-"} (either headings ["-header"][]) to-local-file database ajoin [{"} trim/lines statement {"}]] buffer: make string! 8192 stderr
				all [R3? trim/with buffer "^M"]
			]
		]
		any [empty? stderr cause-error 'user 'message stderr]
		all [raw return buffer]
		either "select" = copy/part statement 6 [
			all ["^/" = buffer return make block! 0] ; SQLite empty result set
			row: copy/part buffer find buffer "^/"
			all [#"^-" = last row append row "^-"]
			size: length? parse/all row "^-"
			buffer: parse/all buffer "^-^/"
			;	clean SQLCMD buffer
			if string? database [
				either headings [remove/part skip buffer size size][remove/part buffer size * 2]
				remove back tail buffer
				foreach val buffer [all [val == "NULL" clear val]]
			]
			;	remove leading and trailing spaces
			foreach val buffer [
				trim val
			]
			;	replace string! with integer!
			if key [
				rows: divide length? buffer size
				all [headings -- rows]
				foreach offset to-block columns [
					all [headings offset: offset + size]
					loop rows [
						poke buffer offset to integer! pick buffer offset
						offset: offset + size
					]
				]
			]
			also buffer buffer: none
		][trim/with buffer "()^/"]
	]

	set 'load-dsv funct [
		"Loads delimiter-separated values from a file."
		file [file! string!]
		/delimit {Alternate delimiter (default is tab then comma)}
			delimiter [char!]
		/sheet "Excel worksheet (default is 1)"
			number [integer!]
		/blocks "Rows as blocks"
	][
		if file? file [
			all [zero? size? file return make block! 0]
			file: either XLS? file [delimiter: #"," read-xls/sheet file any [number 1]][read/string file]
		]
		all [empty? file return make block! 0]
		delimiter: form any [
			delimiter
			either find/part file "^-" any [find file "^/" file] [#"^-"][#","]
		]
		; Parse rules
		valchars: remove/part charset [#"^(00)" - #"^(FF)"] crlf
		valchars: compose [any (remove/part valchars delimiter)]
		value: [
			; Value in quotes, with Excel-compatible handling of bad syntax
			{"} (clear val) x: [to {"} | to end] y: (insert/part tail val x y)
			any [{"} x: {"} [to {"} | to end] y: (insert/part tail val x y)]
			[{"} x: valchars y: (insert/part tail val x y) | end]
			(append blk trim copy val) |
			; Raw value
			x: valchars y: (append blk trim copy/part x y)
		]
		val: make string! 1000
		either blocks [
			output: make block! 1
			parse/all file [z: any [end break | (blk: copy []) value any [delimiter value][crlf | cr | lf | end] (output: insert/only output blk)] z:]
			blk: head output
		][
			blk: make block! 100000
			parse/all file [z: any [end break | value any [delimiter value][crlf | cr | lf | end]] z:]
		]
		also blk (file: output: blk: value: val: x: y: z: valchars: none)
	]

	set 'munge funct [
		"Load and/or manipulate a block of tabular (column and row) values."
		data [block! file!] "REBOL block, CSV or Excel file"
		spec [integer! block! none!] "Size of each record or block of heading words (none! gets cols? file)"
		/update "Offset/value pairs (returns original block). Value can be expression referencing columns (eg c1...) and rowid (eg [ajoin [rowid ") " c1]])"
			action [block!]
		/delete "Delete matching rows (returns original block)"
		/part "Offset position(s) to retrieve"
			columns [block! integer! word!]
		/where "Rowid or expression that can reference columns as c1 or A, c2 or B, specific rows by rowid (eg [rowid < 10]) etc" ; Tom added reference to rowid
			condition [block! integer!]
		/headings "Returns heading words as first row (unless condition is integer)"
		/compact "Remove blank rows"
		/only "Remove duplicate rows"
		/merge "Join outer block (data) to inner block on keys"
			inner-block [block!] "Block to lookup values in"
			inner-size [integer!] "Size of each record"
			cols [block!] "Offset position(s) to retrieve in merged block"
			keys [block!] "Outer/inner join column pairs"
		/group "One of count, flip, max, min, sum, avg, collect" ; Tom added avg, collect (flip doesn't work)
			having [word! block!] "Word or expression that can reference the initial result set column as count, flip, max, etc"
		/order "Sort result set"
		/save "Write result to a delimited or Excel file"
			file [file! none!] "csv, xml, xlsx or tab delimited"
		/list "Return new-line records"
	][
		if file? data [
			any [spec XLS? data spec: cols? data]
			data: load-dsv data
			any [spec spec: excel-columns]
		]

		all [
			empty? data
			either save [write file "" exit][return data]
		]

		either integer? spec [
			spec: make block! size: spec
			repeat i size [append spec to word! append form 'c i]
		][size: length? spec]

		row: spec

		any [
			integer? rows: divide length? data size
			cause-error 'user 'message "size not a multiple of length"
		]

		all [
			not integer? columns
			repeat i length? columns: to block! columns [
				all [
					word? val: pick columns i
					3 > length? val: form val
					poke columns i either 1 = length? uppercase val [subtract to integer! first val 64][(26 * subtract to integer! first val 64) + subtract to integer! second val 64]
				]
			]
		]

		blk: copy []

		all [
			headings
			not update
			either part [foreach i to block! columns [append blk pick spec i]][repeat i size [append blk pick spec i]]
		]
		
		rowid: 0

		case [
			integer? condition [
				i: condition * size - size
				return case [
					update [
						foreach [col val] reduce action [
							poke data i + col val
						]
						data
					]
					delete [head remove/part skip data i size]
					part [
						either integer? columns [pick data i + columns][
							blk: make block! length? columns
							foreach col columns [
								append blk pick data i + col
							]
							blk
						]
					]
					true [copy/part skip data i size]
				]
			]
			update [
				foreach [col val] reduce action [
					append blk compose [
						poke data i + (col) (either datatype? val [compose [to (val) pick data i + (col)]][val])
					]
				]
				i: 0
				either where [bind condition 'rowid bind blk 'rowid][bind blk 'rowid] ; Tom 19.04.15
				do compose/deep [
					foreach [(row)] data [
						++ rowid ; Tom
						either where [all [(condition) (blk)]][(blk)] ; Tom 19.04.15
						i: i + (size)
					]
				]
				return data
			]
			delete [
				either where [
					bind condition 'rowid ; Tom
					do compose/deep [
						remove-each [(row)] data [++ rowid all [(condition)]] ; Tom
					]
				][clear data]
				return data
			]
			part [
				either where [
					bind condition 'rowid ; Tom
					either block? columns [
						part: reduce ['reduce copy []]
						foreach col columns [
							append last part pick row col
						]
					][part: pick row columns]
					do compose/deep [
						foreach [(row)] data [
							++ rowid ; Tom
							all [
								(condition)
								append blk (part)
							]
						]
					]
				][
					either block? columns [
						part: reduce ['reduce copy []]
						foreach col columns [
							append last part either col = 'rowid [[rowid]][compose [pick data (col)]]
						]
					][part: compose [pick data (columns)]]
					repeat rowid rows compose [
						append blk (part)
						data: skip data (size)
					]
					data: head data
				]
				row: make block! size: either integer? columns [1][length? columns]
				foreach col to block! columns [append row either integer? col [pick spec col][col]]
				data: blk
			]
			where [
				bind condition 'rowid ; Tom
				do compose/deep [
					foreach [(row)] data [
						++ rowid ; Tom
						all [(condition) append blk reduce [(row)]]
					]
				]
				data: blk
			]
			headings [
				append blk data
				data: blk
			]
		]

		all [
			compact
			do compose/deep [remove-each [(row)] data [empty? form reduce [(row)]]]
		]

		all [
			empty? data
			either save [write file "" exit][return data]
		]

		if only [
			only: unset!
			either size > 1 [
				do compose/deep [
					remove-each [(row)] sort/skip/all data size [
						either only = reduce [(row)][true][only: reduce [(row)] false]
					]
				]
			][
				remove-each val sort data [
					either only = val [true][only: val false]
				]
			]
		]

		if merge [
			rowids: munge/part inner-block inner-size append munge/part keys 2 2 'rowid
			either 2 = length? keys [key: pick row first keys][
				key: make block! divide length? keys 2
				foreach col munge/part keys 2 1 [
					append key pick row col
				]
			]

			part: copy []
			foreach col cols [
				append part either col <= size [pick row col][
					compose [pick inner-block (col - size)]
				]
			]
			size: length? cols

			blk: copy []
			do compose/deep [
				foreach [(row)] data [
					all [
						rowid: select/skip rowids (either word? key [key][compose/deep [reduce [(key)]]]) (1 + divide length? keys 2)
						inner-block: skip head inner-block rowid - 1 * (inner-size)
						append blk reduce [(part)]
					]
				]
			]

			inner-block: head inner-block

			data: blk
		]
		if group [
			i: s: c: 0 res: copy [] blk: copy []
			sum: funct [blk [block!]][i: 0 foreach n blk [i: i + n]]
			avg: funct [blk [block!]][divide sum blk length? blk]
			get-res: [
				foreach operation having [
					append res switch operation [
						max [first maximum-of val]
						min [first minimum-of val]
						sum [sum val]
						avg [avg val]
						collect [reduce [val]]
						count [length? val]
					]
				] 
				insert insert tail blk group either (length? res) = 1 [res][reduce [res]]
			]
			unless (length? having: to block! having) = (length? intersect having [count flip max min sum avg collect])
				[cause-error 'user 'message "Invalid group operation"]
			either size = 1 [
				foreach operation having [
					append res switch operation [
						count [
							data1: copy data
							sort data1 
							group: copy/part data1 size
							loop rows [
								either group = copy/part data1 size [++ i][
									insert insert tail blk group i
									group: copy/part data1 size
									i: 1
								]
								data1: skip data1 size
							]
							insert insert tail blk group i
							++ size
							append row operation
							reduce [blk]
						]
						max [copy/part maximum-of data 1]
						min [copy/part minimum-of data 1]
						sum [sum data]
						avg [avg data]
						collect [reduce [data]]
					]
					i: s: c: 0
				]
				data: either (length? res) = 1 [first res][res]
				1
			][
				val: copy []
				sort/skip/all data size
				n: 1 ;length? having
				group: copy/part data size - n
				loop rows [
					either group = copy/part data (size - n) [
						append val pick data size
					] [
						do get-res
						group: copy/part data (size - n)
						val: to block! pick data size 
						res: copy []
					]
					data: skip data (size)
				] 
				do get-res
				data: blk
				;poke row size operation ; Tom commented out 14.04.15
			]
			blk: val: res: none ;Tom muutis 13.04.15
			;all [block? having return munge/where data row having] ;Tom commented out 12.04.15
		]

		all [order sort/skip/all skip data either headings [size][0] size]

		also case [
			save [
				any [file file: %.]
				either find [%.xlsx %.xml] suffix? file [
					either headings [worksheet/new/save data size file][worksheet/new/no-header/save data size file]
				][
					blk: copy ""
					i: 1
					loop divide length? data size compose/deep [
						loop (size) [
							insert insert tail blk (
								either %.csv = suffix? file [
									[either find form val: pick data ++ i "," [ajoin [{"} val {"}]][val] ","]
								][
									[pick data ++ i "^-"]
								]
							)
						]
						poke blk length? blk #"^/"
					]
					either file = %. [parse/all blk "^/"][write file blk true]
				]
			]
			list [new-line/all/skip data true size]
			true [data]
		][data: blk: none recycle]
	]

	set 'read-pdf funct [	;	requires pdftotext.exe from http://www.foolabs.com/xpdf/download.html
		"Reads from a PDF file."
		file [file!]
		/lines "Handles data as lines"
	][
		any [exists? file cause-error 'access 'cannot-open file]
		call/wait ajoin [{"} base-path {pdftotext.exe" -nopgbrk -table "} to-local-file clean-path file {" $tmp.txt}]
		also either lines [
			remove-each line read/lines %$tmp.txt [empty? trim/tail line]
		][trim/tail read/string %$tmp.txt] delete %$tmp.txt
	]

	set 'read-xls funct [
		"Reads from an XLS file."
		file [file!]
		/sheet "Excel worksheet (default is 1)"
			number [integer!]
		/lines "Handles data as lines"
	][
		call-excel-vbs file any [number 1] ajoin [{X.ActiveWorkbook.Worksheets(} any [number 1] {).SaveAs "} to-local-file file: join first split-path clean-path file %$tmp.csv {",6:}]
		either exists? file [
			also either lines [
				remove-each line read/lines file ["," = unique trim line]
			][trim/tail read/string file] delete file
		][cause-error 'access 'cannot-open "$tmp.csv"]
	]

	set 'sqlcmd funct [
		"Execute a SQL Server statement."
		server [string!]
		database [string!]
		statement [string!]
		/key "Columns to convert to integer"
			columns [integer! block!]
		/headings "Keep column headings"
	][
		database: ajoin [server "/" database]
		any [columns columns: make block! 0]
		either headings [
			execute/key/headings database statement columns
		][
			execute/key database statement columns
		]
	]

	set 'unzip funct [
		"Uncompress a file into a folder of the same name."
		file [file!]
		/only "Use current folder"
	][
		any [exists? file cause-error 'access 'cannot-open file]
		either only [
			path: first split-path clean-path file
		][
			make-dir path: replace clean-path file %.zip "/"
		]
		write %$tmp.vbs ajoin [{
			set F=CreateObject("Scripting.FileSystemObject")
			set S=CreateObject("Shell.Application")
			S.NameSpace("} to-local-file path {").CopyHere(S.NameSpace("} to-local-file clean-path file {").items)
			Set F=Nothing
			Set S=Nothing
		}]
		also call/wait "cmd /C $tmp.vbs" delete %$tmp.vbs
	]

	set 'worksheet funct [
		"Add a worksheet to the current workbook."
		data [block!]
		spec [integer! block!]
		/new "Start a new workbook"
		/no-header
		/sheet
			name [string!]
		/save "Save current workbook as an XML or XLSX file"
			file [file!]
	][
		either integer? spec [
			cols: spec
			insert/dup spec: make block! cols 100 cols
		][
			cols: length? spec
		]

		any [integer? rows: divide length? data cols cause-error 'user 'message "column / data mismatch"]

		workbook: ""
		sheets: []

		all [
			any [new empty? workbook]
			insert clear workbook {<?xml version="1.0"?><?mso-application progid="Excel.Sheet"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns:html="http://www.w3.org/TR/REC-html40"><Styles><Style ss:ID="s1"><Interior ss:Color="#DDDDDD" ss:Pattern="Solid"/></Style><Style ss:ID="s2"><NumberFormat ss:Format="Long Date"/></Style></Styles>}
			clear sheets
		]

		either sheet [trim/with to-xml-string name "/\"][name: reform ["Sheet" 1 + length? sheets]]
		either find sheets name [cause-error 'user 'message "Duplicate sheet name"][append sheets name]
		append workbook ajoin [{<Worksheet ss:Name="} name {"><Table>}]

		unless zero? rows [
		
			foreach width spec [
				append workbook ajoin [{<Column ss:Width="} width {"/>}]
			]

			no-header: either no-header [0][
				append workbook <Row>
				repeat i cols [
					append workbook ajoin [<Cell ss:StyleID="s1"><Data ss:Type="String"> first data </Data></Cell>]
					data: next data
				]
				append workbook </Row>
				1
			]

			loop rows - no-header [
				append workbook <Row>
				loop cols [
					append workbook case [
						number? val: first data [
							ajoin [<Cell><Data ss:Type="Number"> val </Data></Cell>]
						]
						date? val [
							ajoin [
								<Cell ss:StyleID="s2"><Data ss:Type="DateTime">
								head insert skip head insert skip form (val/year * 10000) + (val/month * 100) + val/day 6 "-" 4 "-"
								</Data></Cell>
							]
						]
						all [string? val "=" = copy/part val 1][
							ajoin [{<Cell ss:Formula="} val {"><Data ss:Type="Number"></Data></Cell>}]
						]
						true [
							either any [none? val empty? val: trim form val][<Cell />][
								ajoin [<Cell><Data ss:Type="String"> to-xml-string val </Data></Cell>]
							]
						]
					]
					data: next data
				]
				append workbook </Row>
			]

			data: head data
		]

		append workbook "</Table></Worksheet>"

		if save [
			append workbook </Workbook>
			switch suffix? file [
				%.xml [
					foreach tag ["<Workbook" "<?mso-application" "<Style" </Styles> "<Worksheet" <Table> "<Column" <Row> "<Cell" </Row> </Table> </Worksheet> </Workbook>][
						replace/all workbook tag either tag = "<Cell" [join "^/^-" tag][join "^/" tag]
					]
					write file workbook
				]
				%.xlsx [
					write %$tmp.xml workbook
					also call-excel-vbs %$tmp.xml 1 ajoin [{W.SaveAs "} to-local-file clean-path file {",51:}] delete %$tmp.xml
					any [exists? file cause-error 'access 'cannot-open file]
				]
			]
			clear workbook
			clear sheets
		]

		exit
	]
]
