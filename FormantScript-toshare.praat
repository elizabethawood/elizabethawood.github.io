#########################
# This is a script to get the values of 
# f1, f2 and f3 at the midpoint, 
# 20% point and 80% point of every
# non-empty interval in the corresponding
# textgrid, for every wav file in a directory.
# The script also measures the duration of each interval.
# Results are written into a spreadsheet.

# To use this script, check the following: 
# Make sure the .wav and .textgrid files have identical names
# and are all together in one folder
# File locations: put the path to the folder your files are 
# in as the "working directory" (see line 26)
# Tier names/numbers: choose the tier the script should 
# look at (see line 46)
# Formant settings: choose the max value for the type of 
# speaker (see line 59)

# This script was made by Elizabeth Wood, 
# mostly through an exercise in Daniel Riggs'
# Praat scripting tutorial, https://praatscriptingtutorial.com/

clearinfo

# First, set up the working directory
# for relative path names
wd$ = homeDirectory$ + "/OneDrive/Documents/Research/"
	... + "Script test/"

# The files that will be analyzed are in the 
# working directory and end in .wav
inDir$ = wd$ + "*.wav"

# The results will be put into the output 
# directory which will be a folder inside the 
# working directory named "Results"
outDir$ = wd$ + "Results/"	

# Next, set up variables to be used
# The separator in the results table will be a tab
sep$ = tab$
# The output file will be saved in 
# the output directory
outFile$ = outDir$ + "Table.csv"
# The number of the tier you want the script to look at
tierNum = 3

# Setting related to the formant measurements
# can be changed as needed
timeStep = 0
maxNumFormants = 4
windowLength = 0.025
preemph = 50
maxFormantValueMale = 5000
maxFormantValueFemale = 5500
maxFormantValueChild = 8000

# Choose max formant value based on type of speaker:
#maxFormantValue = maxFormantValueMale
maxFormantValue = maxFormantValueFemale
#maxFormantValue = maxFormantValueChild

# Now set up the results table
header$ = "fileName" + sep$ + "intervalNumber" + sep$
	... + "label" + sep$ + "midpoint" + sep$
	... + "f1" + sep$ + "f2" + sep$ + "f3" + sep$ 
	... + "f1_20" + sep$ + "f1_80" + sep$ + "f2_20" + sep$ + "f2_80" + sep$
	... + "f3_20" + sep$ + "f3_80" + sep$
	... + "duration" + newline$

createDirectory: outDir$

# If an identical results table already exists,
# the script will ask before overwriting it
if fileReadable: outFile$
	pauseScript: "Data spreadsheet already exists, overwrite?"
endif	
deleteFile: outFile$
writeFile: outFile$, header$

# Create and select the string with all of the .wav files
relFiles = Create Strings as file list: "relFiles", inDir$
selectObject: "Strings relFiles"

# How many wav files are in the string?
numFiles = Get number of strings
numFiles$ = string$: numFiles
appendInfoLine: "There are " + numFiles$ 
	... + " .wav files in this folder"

# This first for-loop goes through 
# each .wav file in the folder 
for x from 1 to numFiles
selectObject: relFiles

# We want it to find the corresponding textgrid
# The textgrid file will be named like the wav
# but its path will end in .TextGrid
wavName$ = Get string: x
fileName$ = wavName$ - ".wav"
textgridName$ = fileName$ + ".TextGrid"
appendInfoLine: "wav file " + fileName$
textgridPath$ = wd$ + textgridName$
soundPath$ = wd$ + wavName$

# Open the textgrid that you want info from
if fileReadable: textgridPath$
	textgridx = Read from file: textgridPath$	

# How many intervals does the relevant tier have?
	numInt = Get number of intervals: tierNum

# Open the sound file and make the formant object
	soundx = Read from file: soundPath$
	formantx = To Formant (burg): timeStep, maxNumFormants, 
			  ... maxFormantValue, windowLength, preemph

# Loop through each interval and find the ones that are not empty
	for i from 1 to numInt
	selectObject: "TextGrid " + fileName$
		# i will be the number of the current interval
		i$ = string$: i
		# intLabel will be the label of the current interval
		intLabel$ = Get label of interval: tierNum, i
		if intLabel$ <> ""
		  intStart = Get start time of interval: tierNum, i 
		  intEnd = Get end time of interval: tierNum, i
		  intMid = intStart + ((intEnd - intStart)/2)
		  inttwenty = intStart + ((intEnd - intStart)/5)
		  inteighty = intStart + (4 * (intEnd - intStart)/5)
		  intMid$ = string$: intMid
		  intDur = intEnd - intStart
		  intDur$ = string$: intDur
		  
	selectObject: "Formant " + fileName$
		  numFormants = Get number of formants: 1	
		  f1_mid = Get value at time: 1, intMid, "hertz", "linear"
		  f2_mid = Get value at time: 2, intMid, "hertz", "linear"
		  f3_mid = Get value at time: 3, intMid, "hertz", "linear"
		  f1_20 = Get value at time: 1, inttwenty, "hertz", "linear"
		  f2_20 = Get value at time: 2, inttwenty, "hertz", "linear"
		  f3_20 = Get value at time: 3, inttwenty, "hertz", "linear"
		  f1_80 = Get value at time: 1, inteighty, "hertz", "linear"
		  f2_80 = Get value at time: 2, inteighty, "hertz", "linear"
		  f3_80 = Get value at time: 3, inteighty, "hertz", "linear"
		  f1_mid$ = string$: f1_mid
		  f2_mid$ = string$: f2_mid
		  f3_mid$ = string$: f3_mid
		  f1_20$ = string$: f1_20
		  f1_80$ = string$: f1_80
		  f2_20$ = string$: f2_20
		  f2_80$ = string$: f2_80
		  f3_20$ = string$: f3_20
		  f3_80$ = string$: f3_80
		  
# Output all of these values into the results table
		  intInQuestRow$ = fileName$ + sep$ + i$ + sep$
		  	... + intLabel$ + sep$ + intMid$ + sep$  
		  	... + f1_mid$ + sep$ + f2_mid$ + sep$ + f3_mid$ + sep$
		  	... + f1_20$ + sep$ + f1_80$ + sep$ + f2_20$ + sep$ + f2_80$ + sep$
		  	... + f3_20$ + sep$ + f3_80$ + sep$
		  	... + intDur$ + newline$

		  appendFile: outFile$, intInQuestRow$ 

# Close all of the open ifs and fors	
	  	# Ends the conditional for if the current label is non-empty
		endif
	# Ends the for-loop for each interval in the current textgrid 
	endfor

# Ends the conditional for if the relevant file exists
endif

# Finally, clean up Praat by closing all the open files
 selectObject: "Formant " + fileName$
 plusObject: "Sound " + fileName$
 plusObject: "TextGrid " + fileName$
 Remove

# Ends the for-loop for each textgrid in the input directory folder
endfor

selectObject: "Strings relFiles"
Remove

pauseScript: "All done!"