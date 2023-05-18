#########################
# This is a script to get the values of 
# a number of different measures of voice quality 
# in each third of each vowel with a given label.
# You must have a folder with .wav files and 
# corresponding textgrids of the same name.
# The script is written to reference and extract information
# from a number of different tiers (see below, remove if not needed).
# Results are written into a spreadsheet.
# Things that may need to change: 
# Working directory (where files are)
# Formant settings: currently set for female speakers
# Which tiers to get info from
# Assorted other settings

# This script was cobbled together by Elizabeth Wood.
# Parts of this script were modified from a script by
# Chad Vicenik called praatvoicesauceimitator.praat
# taken from the UCLA phonetics lab website,
# which itself is based off of a similar script by Bert Remijsen.
# Other parts of this script were modified from 
# the formant script made by Daniel Riggs for his online
# Praat scripting tutorial.

clearinfo

# Set up the working directory
# for relative path names

wd$ = homeDirectory$ + "/OneDrive/Documents/Research/Glottals/Textgrids/Female/All/Test/"	

inDir$ = wd$ + "*.wav"
outDir$ = wd$ + "Results/"	


# Set up variables to be used
sep$ = tab$
outFile$ = outDir$ + "Table.csv"
tierNum = 3
catTier = 4
codTier = 5
voiceTypeTier = 7


#Formant settings
timeStep = 0 
# 0 means Praat will do 25% of analysis window length
maxNumFormants = 5
freqcost = 1
bwcost = 1
transcost = 1
windowLength = 0.025
preemph = 50

## Use this for female voice: 
 maxFormantValue = 5500
 f1ref = 550
 f2ref = 1650
 f3ref = 2750
 f4ref = 3850
 f5ref = 4950

## Use this for male voice:
# maxFormantValue = 5000
# f1ref = 500
# f2ref = 1485
# f3ref = 2475
# f4ref = 3465
# f5ref = 4455

# Spectrogram settings
 windowLengthSpec = 0.005
 maxFreSpec = 5000
 timeStepSpec = 0.002
 freStepSpec = 20
 windowShapeSpec$ = "Gaussian"

# Pitch settings
timeStepPitch = 0 
# 0 means Praat will use 0.75/pitch floor
maxCandidates = 15
accuracy$ = "no"
silenceThreshPitch = 0.01
octaveCost = 0.01
octaveJumpCost = 0.5
# standard is 0.35
voicedUnvoicedCost = 0.14

# Harmonicity settings
 timeStepGlot = 0.01
 minPitch = 75
 silenceThresh = 0.01
 # standard is 0.1
 perPerWindow = 1.0

# Extract part settings
windowShapePart$ = "rectangular"
## whoever I stole this from used Hanning but rectangular is standard
relativeWidthPart = 1
preserveTimesPart$ = "no"

# Jitter settings
shortPerJitt = 0.0001
longPerJitt = 0.02
## script I copied used 0.05 but 0.02 is standard
maxPerFactJitt = 5
## script I copied used 5 but 1.3 is standard

# Shimmer settings
shortPerShim = 0.0001
longPerShim = 0.02
## script I copied used 0.05 but 0.02 is standard
maxPerFactShim = 1.3
## script I copied used 5 but 1.3 is standard
maxAmpFactShim = 5
## script I copied used 5 but 1.6 is standard


# Start the results table

header$ = "fileName" + sep$ + "intervalNumber" + sep$ + "timestamp" + sep$
	... + "label" + sep$ + "category" + sep$ + "coding" + sep$
	... + "voicetype" + sep$
	... + "full gs?" + sep$
	... + "duration" + sep$ + "seg duration" + sep$ + "seg 1 type" + sep$
	... + "HNR mean 1" + sep$ + "HNR mean 2" + sep$ + "HNR mean 3" + sep$
	... + "H1-H2 1" + sep$ + "H1-H2 2" + sep$ + "H1-H2 3" + sep$  
	... + "H1-A1 1" + sep$ + "H1-A1 2" + sep$ + "H1-A1 3" + sep$
	... + "H1-A2 1" + sep$ + "H1-A2 2" + sep$ + "H1-A2 3" + sep$
	...	+ "H1-A3 1" + sep$ + "H1-A3 2" + sep$ + "H1-A3 3" + sep$
	... + "Jitter 1" + sep$ + "Jitter 2" + sep$ + "Jitter 3" + sep$
	... + "Shimmer 1" + sep$ + "Shimmer 2" + sep$ + "Shimmer 3" + sep$
	... + "Min intensity 1" + sep$ + "Min intensity 2" + sep$ + "Min intensity 3" + sep$
	... + "Min pitch 1" + sep$ + "Min pitch 2" + sep$ + "Min pitch 3" + sep$
	... + "F1 first" + sep$ + "F1 mid" + sep$ + "F1 second" + sep$
	... + "F2 first" + sep$ + "F2 mid" + sep$ + "F2 second" + sep$
	... + "F3 first" + sep$ + "F3 mid" + sep$ + "F3 second"
	... + newline$

createDirectory: outDir$

# Warning if a results file already exists
if fileReadable: outFile$
	pauseScript: "Data spreadsheet already exists, overwrite?"
endif	
deleteFile: outFile$

writeFile: outFile$, header$


# Create the string with all of the 
# wav files

relFiles = Create Strings as file list: "relFiles", inDir$

selectObject: "Strings relFiles"


# How many wav files are in the folder?

numFiles = Get number of strings
numFiles$ = string$: numFiles

appendInfoLine: "There are " + numFiles$ 
	... + " .wav files in this folder"


# Different settings may be needed for specific 
# vowels in order for Praat to correctly 
# identify the pitch pulses. 
# This should be checked ahead of time and 
# each vowel labeled with the right
# voiceType code in the voiceType tier

# Loop through each specific voicetype settings
for voiceTypes from 1 to 42

if voiceTypes = 1

voiceType$ = "1"
pitchFloor = 40
pitchCeiling = 200
voicingThreshPitch = 0.2

elsif voiceTypes = 2

voiceType$ = "2"
pitchFloor = 50
pitchCeiling = 200
voicingThreshPitch = 0.3

elsif voiceTypes = 3

voiceType$ = "3"
pitchFloor = 50
pitchCeiling = 300
voicingThreshPitch = 0.2

elsif voiceTypes = 4

voiceType$ = "4"
pitchFloor = 50
pitchCeiling = 150
voicingThreshPitch = 0.1

elsif voiceTypes = 5

voiceType$ = "5"
pitchFloor = 50
pitchCeiling = 300
voicingThreshPitch = 0.3

elsif voiceTypes = 6

voiceType$ = "6"
pitchFloor = 100
pitchCeiling = 300
voicingThreshPitch = 0.1

elsif voiceTypes = 7

voiceType$ = "7"
pitchFloor = 70
pitchCeiling = 200
voicingThreshPitch = 0.1

elsif voiceTypes = 8

voiceType$ = "8"
pitchFloor = 40
pitchCeiling = 300
voicingThreshPitch = 0.2

elsif voiceTypes = 9

voiceType$ = "9"
pitchFloor = 40
pitchCeiling = 200
voicingThreshPitch = 0.3

elsif voiceTypes = 10

voiceType$ = "10"
pitchFloor = 50
pitchCeiling = 200
voicingThreshPitch = 0.2

elsif voiceTypes = 11

voiceType$ = "11"
pitchFloor = 100
pitchCeiling = 300
voicingThreshPitch = 0.3

elsif voiceTypes = 12

voiceType$ = "12"
pitchFloor = 75
pitchCeiling = 300
voicingThreshPitch = 0.3

elsif voiceTypes = 13

voiceType$ = "13"
pitchFloor = 80
pitchCeiling = 300
voicingThreshPitch = 0.2

elsif voiceTypes = 14

voiceType$ = "14"
pitchFloor = 50
pitchCeiling = 200
voicingThreshPitch = 0.45

elsif voiceTypes = 15

voiceType$ = "15"
pitchFloor = 100
pitchCeiling = 200
voicingThreshPitch = 0.3

elsif voiceTypes = 16

voiceType$ = "16"
pitchFloor = 80
pitchCeiling = 200
voicingThreshPitch = 0.3

elsif voiceTypes = 17

voiceType$ = "17"
pitchFloor = 50
pitchCeiling = 150
voicingThreshPitch = 0.3

elsif voiceTypes = 18

voiceType$ = "TRI"
pitchFloor = 150
pitchCeiling = 600
voicingThreshPitch = 0.45

elsif voiceTypes = 19

voiceType$ = "LRI"
pitchFloor = 150
pitchCeiling = 600
voicingThreshPitch = 0.45

elsif voiceTypes = 20

voiceType$ = "20"
pitchFloor = 60
pitchCeiling = 300
voicingThreshPitch = 0.3

elsif voiceTypes = 21

voiceType$ = "LXE"
pitchFloor = 100
pitchCeiling = 450
voicingThreshPitch = 0.45

elsif voiceTypes = 22

voiceType$ = "22"
pitchFloor = 50
pitchCeiling = 300
voicingThreshPitch = 0.1

elsif voiceTypes = 23

voiceType$ = "23"
pitchFloor = 70
pitchCeiling = 300
voicingThreshPitch = 0.1

elsif voiceTypes = 24

voiceType$ = "24"
pitchFloor = 40
pitchCeiling = 300
voicingThreshPitch = 0.3

elsif voiceTypes = 25

voiceType$ = "25"
pitchFloor = 70
pitchCeiling = 300
voicingThreshPitch = 0.3

elsif voiceTypes = 26

voiceType$ = "26"
pitchFloor = 120
pitchCeiling = 300
voicingThreshPitch = 0.1

elsif voiceTypes = 27

voiceType$ = "JT"
pitchFloor = 80
pitchCeiling = 575
voicingThreshPitch = 0.45

elsif voiceTypes = 28

voiceType$ = "MJL"
pitchFloor = 70
pitchCeiling = 470
voicingThreshPitch = 0.45

elsif voiceTypes = 29

voiceType$ = "29"
pitchFloor = 40
pitchCeiling = 100
voicingThreshPitch = 0.1

elsif voiceTypes = 30

voiceType$ = "30"
pitchFloor = 70
pitchCeiling = 180
voicingThreshPitch = 0.3

elsif voiceTypes = 31

voiceType$ = "31"
pitchFloor = 50
pitchCeiling = 100
voicingThreshPitch = 0.1

elsif voiceTypes = 32

voiceType$ = "32"
pitchFloor = 50
pitchCeiling = 200
voicingThreshPitch = 0.1

elsif voiceTypes = 33

voiceType$ = "TJL"
pitchFloor = 75
pitchCeiling = 560
voicingThreshPitch = 0.45

elsif voiceTypes = 34

voiceType$ = "SAGB"
pitchFloor = 75
pitchCeiling = 490
voicingThreshPitch = 0.45

elsif voiceTypes = 35

voiceType$ = "35"
pitchFloor = 130
pitchCeiling = 475
voicingThreshPitch = 0.3

elsif voiceTypes = 36

voiceType$ = "RSI"
pitchFloor = 130
pitchCeiling = 520
voicingThreshPitch = 0.45

elsif voiceTypes = 37

voiceType$ = "MXM"
pitchFloor = 115
pitchCeiling = 425
voicingThreshPitch = 0.45

elsif voiceTypes = 38

voiceType$ = "38"
pitchFloor = 60
pitchCeiling = 200
voicingThreshPitch = 0.1

elsif voiceTypes = 39

voiceType$ = "39"
pitchFloor = 50
pitchCeiling = 250
voicingThreshPitch = 0.1

elsif voiceTypes = 40

voiceType$ = "40"
pitchFloor = 100
pitchCeiling = 200
voicingThreshPitch = 0.1

elsif voiceTypes = 41

voiceType$ = "MACM"
pitchFloor = 130
pitchCeiling = 475
voicingThreshPitch = 0.45

elsif voiceTypes = 42

voiceType$ = "Male"
pitchFloor = 50
pitchCeiling = 450
voicingThreshPitch = 0.45


endif

appendInfoLine: "Now starting on voicetype " + voiceType$


# Loop: this one is going through 
# each wav file in the folder 

for x from 1 to numFiles

selectObject: relFiles

# We want it to find the corresponding textgrid
# The textgrid file must have the same name as the wav file
# but will end in .TextGrid

textgridName$ = Get string: x
appendInfoLine: "wav file " + textgridName$

textgridxPath$ = wd$ + textgridName$ - ".wav" + ".TextGrid"
soundPath$ = wd$ + textgridName$

if fileReadable: textgridxPath$
	textgridx = Read from file: textgridxPath$	

# So we have a textgrid opened that we want info from
# How many intervals does the voice type tier have?
	numIntVoiceType = Get number of intervals: voiceTypeTier
	numIntVoiceType$ = string$: numIntVoiceType
	appendInfoLine: "This file has " + numIntVoiceType$ + " intervals in the voicetype tier"

# We only want to make the pitch object, formant object, etc
# for files that have at least one interval 
# with the relevant voicetype specification

for check from 1 to numIntVoiceType
voiceCheck$ = Get label of interval: voiceTypeTier, check 

if voiceCheck$ = voiceType$ 
appendInfoLine: "At least one interval had the correct voicetype label"
goto proceedplease
endif 

endfor

# If we got through all of the intervals and none 
# of them had the relevant label
# we don't need to do anything further with that file
appendInfoLine: "None of the intervals had the correct voicetype label"

goto ceaseplease
endif


label proceedplease
# If we do want to include the file in the analysis
# let's open the sound file and make the formant object, 
# the spectrogram object, the pitch object,
# the intensity object and the harmonicity object,

	soundx = Read from file: soundPath$
	formantx = To Formant (burg): timeStep, maxNumFormants, 
			  ... maxFormantValue, windowLength, preemph


selectObject: "Formant " + textgridName$ - ".wav"

	xx = Get minimum number of formants
		if xx > 2
			Track: 3, f1ref, f2ref, f3ref, f4ref, f5ref, freqcost, bwcost, transcost
			else
			Track: 2, f1ref, f2ref, f3ref, f4ref, f5ref, freqcost, bwcost, transcost		  
		endif 	

	Rename: textgridName$ -".wav" + "_aftertracking"		

	selectObject: "Sound " + textgridName$ - ".wav"
	spectrogramx = To Spectrogram: windowLengthSpec, maxFreSpec, timeStepSpec, freStepSpec, windowShapeSpec$

	selectObject: "Sound " + textgridName$ - ".wav"
	pitchx = To Pitch (cc): timeStepPitch, pitchFloor, maxCandidates, accuracy$, silenceThreshPitch, 
	... voicingThreshPitch, octaveCost, octaveJumpCost, voicedUnvoicedCost, pitchCeiling
	Interpolate
	Rename: textgridName$ - ".wav" + "_interpolated"


	selectObject: "Sound " + textgridName$ - ".wav"
		  harmonicityx = To Harmonicity (cc): timeStepGlot,
		  ... minPitch, silenceThresh, perPerWindow	

	selectObject: "Sound " + textgridName$ - ".wav"
	To Intensity: 100, 0, "no"

	selectObject: "Pitch " + textgridName$ - ".wav"
		To PointProcess


# Now let's loop through each interval of the textgrid
# and get its label and the label of the preceding interval
# which also has information we want
	selectObject: "TextGrid " + textgridName$ - ".wav"
	numInt = Get number of intervals: tierNum
	
	for i from 1 to numInt

		selectObject: "TextGrid " + textgridName$ - ".wav"
		intLabel$ = Get label of interval: tierNum, i

		if i > 1
		preIntNum = i-1

		else
		preIntNum = 1
		endif 
		
		preIntLabel$ = Get label of interval: codTier, preIntNum

# We will only perform the analysis
# for non-empty intervals		
		if intLabel$ <> ""
		  intStart = Get start time of interval: tierNum, i
		  corrInt = Get interval at time: 7, intStart
		  correctionLab$ = Get label of interval: 7, corrInt

# And we also only want to include intervals
# that have been identifed as needing the 
# voice settings we are currently using
		if correctionLab$ = voiceType$

		  intEnd = Get end time of interval: tierNum, i
		  intDur = intEnd - intStart
		  intMid = intStart + (intDur / 2)
		  intfirstthird = intStart + (intDur / 3)

 		segDur = intDur / 3
 		segDur$ = string$: segDur
		minSeg1Dur = 1 / pitchFloor

# In order to correctly make the measurements
# Praat needs the window to be at least
# 1/ pitch floor, so the size of the 
# minimum segment for analysis will depend
# on the pitch floor setting
		  if segDur > minSeg1Dur

		  segEnd = intStart + segDur
		  middleStart = intStart + segDur
		  middleEnd = intStart + (2 * segDur)
		  secondSegStart = intStart + (2 * segDur)
		  segDurType$ = "seg1"
		  segActualDur = segDur
		  

		  else
		  		if intDur > minSeg1Dur

		  			segEnd = intStart + minSeg1Dur
		  			middleStart = intMid - (minSeg1Dur / 2)
		  			middleEnd = intMid + (minSeg1Dur / 2)
		  			secondSegStart = intEnd - minSeg1Dur
		  			segDurType$ = "min seg"
		  			segActualDur = minSeg1Dur
		  			
		 		else 
					  segEnd = intEnd
					  middleStart = intStart
					  middleEnd = intEnd
					  secondSegStart = intStart
					  segDurType$ = "whole interval"
					  segActualDur = intDur

		  		endif

		  endif

		
		segmid1 = intStart + (segActualDur / 2)
		secondSegMid = secondSegStart + (segActualDur / 2) 

		segActualDur$ = string$: segActualDur

		  intStart$ = string$: intStart
		  intDur$ = string$: intDur
		  
		  intCatNum = Get interval at time: catTier, intStart
		  intCat$ = Get label of interval: catTier, intCatNum 
		  intCodNum = Get interval at time: codTier, intStart
		  intCod$ = Get label of interval: codTier, intCodNum

		  i$ = string$: i

	# We only want to include word-initial vowels	 
		  	if intCat$ == "initial" or intCat$ == "initial!"

	# Start with the formant measurements
		selectObject: "Formant " + textgridName$ - ".wav" + "_aftertracking"
			f1hzptSegmid1 = Get value at time: 1, segmid1, "hertz", "linear"
			f2hzptSegmid1 = Get value at time: 2, segmid1, "hertz", "linear"
			f1hzptIntmid = Get value at time: 1, intMid, "hertz", "linear"
			f2hzptIntmid = Get value at time: 2, intMid, "hertz", "linear"
			f1hzptSecondSegMid = Get value at time: 1, secondSegMid, "hertz", "linear"
			f2hzptSecondSegMid = Get value at time: 2, secondSegMid, "hertz", "linear"
			
			f1$ = string$: f1hzptSegmid1
			f2$ = string$: f2hzptSegmid1
			f1Mid$ = string$: f1hzptIntmid
			f2Mid$ = string$: f2hzptIntmid
			f1Second$ = string$: f1hzptSecondSegMid
			f2Second$ = string$: f2hzptSecondSegMid


			if xx > 2
			f3hzptSegmid1 = Get value at time: 3, segmid1, "hertz", "linear"
			f3hzptIntmid = Get value at time: 3, intMid, "hertz", "linear"
			f3hzptSecondSegMid = Get value at time: 3, secondSegMid, "hertz", "linear"
			
			else 
			f3hzptSegmid1 = 0
			f3hzptIntmid = 0
			f3hzptSecondSegMid = 0

			
			endif 

			f3$ = string$: f3hzptSegmid1
			f3Mid$ = string$: f3hzptIntmid
			f3Second$ = string$: f3hzptSecondSegMid


			# Now spectral measurements for the first segment

			selectObject: "Sound " + textgridName$ - ".wav"
			
			Extract part: intStart, segEnd, 
				... windowShapePart$, relativeWidthPart, preserveTimesPart$ 
			Rename: textgridName$ - ".wav" + "_slice1"

			
			selectObject: "Sound " + textgridName$ - ".wav" + "_slice1"
			To Spectrum: "yes" 
			To Ltas (1-to-1)

			selectObject: "Pitch " + textgridName$ - ".wav" + "_interpolated"
			segmid1f0 = Get value at time: segmid1, "hertz", "linear"
			
	
			if segmid1f0 <> undefined

			p10_segmid1f0 = segmid1f0 / 10
			selectObject: "Ltas " + textgridName$ - ".wav" + "_slice1"
			lowerbh1_seg1 = segmid1f0 - p10_segmid1f0
			upperbh1_seg1 = segmid1f0 + p10_segmid1f0
			lowerbh2_seg1 = (segmid1f0 * 2) - (p10_segmid1f0 * 2)
			upperbh2_seg1 = (segmid1f0 * 2) + (p10_segmid1f0 * 2)
			h1db_seg1 = Get maximum: lowerbh1_seg1, upperbh1_seg1, "none"
			h1hz_seg1 = Get frequency of maximum: lowerbh1_seg1, upperbh1_seg1, "none"
			h2db_seg1 = Get maximum: lowerbh2_seg1, upperbh2_seg1, "none"
			h2hz_seg1 = Get frequency of maximum: lowerbh2_seg1, upperbh2_seg1, "none"

			p10_f1hzptSegmid1 = f1hzptSegmid1 / 10 
			p10_f2hzptSegmid1 = f2hzptSegmid1 / 10 
			p10_f3hzptSegmid1 = f3hzptSegmid1 / 10
			lowerba1_seg1 = f1hzptSegmid1 - p10_f1hzptSegmid1 
			upperba1_seg1 = f1hzptSegmid1 + p10_f1hzptSegmid1
			lowerba2_seg1 = f2hzptSegmid1 - p10_f2hzptSegmid1 
			upperba2_seg1 = f2hzptSegmid1 + p10_f2hzptSegmid1
			lowerba3_seg1 = f3hzptSegmid1 - p10_f3hzptSegmid1 
			upperba3_seg1 = f3hzptSegmid1 + p10_f3hzptSegmid1
			a1db_seg1 = Get maximum: lowerba1_seg1, upperba1_seg1, "none"
			a1hz_seg1 = Get frequency of maximum: lowerba1_seg1, upperba1_seg1, "none"
			a2db_seg1 = Get maximum: lowerba2_seg1, upperba2_seg1, "none"
			a2hz_seg1 = Get frequency of maximum: lowerba2_seg1, upperba2_seg1, "none"
			a3db_seg1 = Get maximum: lowerba3_seg1, upperba3_seg1, "none"
			a3hz_seg1 = Get frequency of maximum: lowerba3_seg1, upperba3_seg1, "none"

			h1mnh2_seg1 = h1db_seg1 - h2db_seg1
			h1mna1_seg1 = h1db_seg1 - a1db_seg1
			h1mna2_seg1 = h1db_seg1 - a2db_seg1 
			h1mna3_seg1 = h1db_seg1 - a3db_seg1 

			else
				h1mnh2_seg1 = 0
				h1mna1_seg1 = 0
				h1mna2_seg1 = 0
				h1mna3_seg1 = 0
		endif 

			h1mnh2_seg1$ = string$: h1mnh2_seg1
			h1mna1_seg1$ = string$: h1mna1_seg1
			h1mna2_seg1$ = string$: h1mna2_seg1
			h1mna3_seg1$ = string$: h1mna3_seg1

			# The same measurements for the middle segment
			selectObject: "Sound " + textgridName$ - ".wav"
			
			Extract part: middleStart, middleEnd, 
				... windowShapePart$, relativeWidthPart, preserveTimesPart$ 
			Rename: textgridName$ - ".wav" + "_slice2"

			
			selectObject: "Sound " + textgridName$ - ".wav" + "_slice2"
			To Spectrum: "yes" 
			To Ltas (1-to-1)

			selectObject: "Pitch " + textgridName$ - ".wav" + "_interpolated"
			segMiddlef0 = Get value at time: intMid, "hertz", "linear"
			
			if segMiddlef0 <> undefined

			p10_segMiddlef0 = segMiddlef0 / 10
			selectObject: "Ltas " + textgridName$ - ".wav" + "_slice2"
			lowerbh1_middle = segMiddlef0 - p10_segMiddlef0
			upperbh1_middle = segMiddlef0 + p10_segMiddlef0
			lowerbh2_middle = (segMiddlef0 * 2) - (p10_segMiddlef0 * 2)
			upperbh2_middle = (segMiddlef0 * 2) + (p10_segMiddlef0 * 2)
			h1db_middle = Get maximum: lowerbh1_middle, upperbh1_middle, "none"
			h1hz_middle = Get frequency of maximum: lowerbh1_middle, upperbh1_middle, "none"
			h2db_middle = Get maximum: lowerbh2_middle, upperbh2_middle, "none"
			h2hz_middle = Get frequency of maximum: lowerbh2_middle, upperbh2_middle, "none"

			p10_f1hzptMiddle = f1hzptIntmid / 10 
			p10_f2hzptMiddle = f2hzptIntmid / 10 
			p10_f3hzptMiddle = f3hzptIntmid / 10
			lowerba1_middle = f1hzptIntmid - p10_f1hzptMiddle
			upperba1_middle = f1hzptIntmid + p10_f1hzptMiddle
			lowerba2_middle = f2hzptIntmid - p10_f2hzptMiddle
			upperba2_middle = f2hzptIntmid + p10_f2hzptMiddle
			lowerba3_middle = f3hzptIntmid - p10_f3hzptMiddle
			upperba3_middle = f3hzptIntmid + p10_f3hzptMiddle
			a1db_middle = Get maximum: lowerba1_middle, upperba1_middle, "none"
			a1hz_middle = Get frequency of maximum: lowerba1_middle, upperba1_middle, "none"
			a2db_middle = Get maximum: lowerba2_middle, upperba2_middle, "none"
			a2hz_middle = Get frequency of maximum: lowerba2_middle, upperba2_middle, "none"
			a3db_middle = Get maximum: lowerba3_middle, upperba3_middle, "none"
			a3hz_middle = Get frequency of maximum: lowerba3_middle, upperba3_middle, "none"

			h1mnh2_middle = h1db_middle - h2db_middle
			h1mna1_middle = h1db_middle - a1db_middle
			h1mna2_middle = h1db_middle - a2db_middle 
			h1mna3_middle = h1db_middle - a3db_middle 

			else
				h1mnh2_middle = 0
				h1mna1_middle = 0
				h1mna2_middle = 0
				h1mna3_middle = 0
		endif 

			h1mnh2_middle$ = string$: h1mnh2_middle
			h1mna1_middle$ = string$: h1mna1_middle
			h1mna2_middle$ = string$: h1mna2_middle
			h1mna3_middle$ = string$: h1mna3_middle


			# And the same measurements for the third segment

			selectObject: "Sound " + textgridName$ - ".wav"
			
			Extract part: secondSegStart, intEnd, 
				... windowShapePart$, relativeWidthPart, preserveTimesPart$ 
			Rename: textgridName$ - ".wav" + "_slice3"

			
			selectObject: "Sound " + textgridName$ - ".wav" + "_slice3"
			To Spectrum: "yes" 
			To Ltas (1-to-1)

			selectObject: "Pitch " + textgridName$ - ".wav" + "_interpolated"
			secondSegf0 = Get value at time: secondSegMid, "hertz", "linear"

			if secondSegf0 <> undefined

			p10_SecondSegf0 = secondSegf0 / 10
			selectObject: "Ltas " + textgridName$ - ".wav" + "_slice3"
			lowerbh1_SecondSeg = secondSegf0 - p10_SecondSegf0
			upperbh1_SecondSeg = secondSegf0 + p10_SecondSegf0
			lowerbh2_SecondSeg = (secondSegf0 * 2) - (p10_SecondSegf0 * 2)
			upperbh2_SecondSeg = (secondSegf0 * 2) + (p10_SecondSegf0 * 2)
			h1db_SecondSeg = Get maximum: lowerbh1_SecondSeg, upperbh1_SecondSeg, "none"
			h1hz_SecondSeg = Get frequency of maximum: lowerbh1_SecondSeg, upperbh1_SecondSeg, "none"
			h2db_SecondSeg = Get maximum: lowerbh2_SecondSeg, upperbh2_SecondSeg, "none"
			h2hz_SecondSeg = Get frequency of maximum: lowerbh2_SecondSeg, upperbh2_SecondSeg, "none"

			p10_f1hzptSecondSeg = f1hzptSecondSegMid / 10 
			p10_f2hzptSecondSeg = f2hzptSecondSegMid / 10 
			p10_f3hzptSecondSeg = f3hzptSecondSegMid / 10
			lowerba1_SecondSeg = f1hzptSecondSegMid - p10_f1hzptSecondSeg
			upperba1_SecondSeg = f1hzptSecondSegMid + p10_f1hzptSecondSeg
			lowerba2_SecondSeg = f2hzptSecondSegMid - p10_f2hzptSecondSeg
			upperba2_SecondSeg = f2hzptSecondSegMid + p10_f2hzptSecondSeg
			lowerba3_SecondSeg = f3hzptSecondSegMid - p10_f3hzptSecondSeg
			upperba3_SecondSeg = f3hzptSecondSegMid + p10_f3hzptSecondSeg
			a1db_SecondSeg = Get maximum: lowerba1_SecondSeg, upperba1_SecondSeg, "none"
			a1hz_SecondSeg = Get frequency of maximum: lowerba1_SecondSeg, upperba1_SecondSeg, "none"
			a2db_SecondSeg = Get maximum: lowerba2_SecondSeg, upperba2_SecondSeg, "none"
			a2hz_SecondSeg = Get frequency of maximum: lowerba2_SecondSeg, upperba2_SecondSeg, "none"
			a3db_SecondSeg = Get maximum: lowerba3_SecondSeg, upperba3_SecondSeg, "none"
			a3hz_SecondSeg = Get frequency of maximum: lowerba3_SecondSeg, upperba3_SecondSeg, "none"

			h1mnh2_SecondSeg = h1db_SecondSeg - h2db_SecondSeg
			h1mna1_SecondSeg = h1db_SecondSeg - a1db_SecondSeg
			h1mna2_SecondSeg = h1db_SecondSeg - a2db_SecondSeg 
			h1mna3_SecondSeg = h1db_SecondSeg - a3db_SecondSeg 

			else
				h1mnh2_SecondSeg = 0
				h1mna1_SecondSeg = 0
				h1mna2_SecondSeg = 0
				h1mna3_SecondSeg = 0
		endif 

			h1mnh2_SecondSeg$ = string$: h1mnh2_SecondSeg
			h1mna1_SecondSeg$ = string$: h1mna1_SecondSeg
			h1mna2_SecondSeg$ = string$: h1mna2_SecondSeg
			h1mna3_SecondSeg$ = string$: h1mna3_SecondSeg

		# Intensity min measurements

		selectObject: "Intensity " + textgridName$ - ".wav"
		
		intMinSeg1 = Get minimum: intStart, segEnd, "parabolic"
		intMinSeg1$ = string$: intMinSeg1

		intMinMiddle = Get minimum: middleStart, middleEnd, "parabolic"
		intMinMiddle$ = string$: intMinMiddle

		intMinSecondSeg = Get minimum: secondSegStart, intEnd, "parabolic"
		intMinSecondSeg$ = string$: intMinSecondSeg


		# F0 min measurements

		selectObject: "Pitch " + textgridName$ - ".wav"
		pitchMinSeg1 = Get minimum: intStart, segEnd, "Hertz", "parabolic"
		pitchMinSeg1$ = string$: pitchMinSeg1

		pitchMinMiddle = Get minimum: middleStart, middleEnd, "Hertz", "parabolic"
		pitchMinMiddle$ = string$: pitchMinMiddle
		
		pitchMinSecondSeg = Get minimum: secondSegStart, intEnd, "Hertz", "parabolic"
		pitchMinSecondSeg$ = string$: pitchMinSecondSeg
		

		# HNR measurements

		selectObject: "Harmonicity " + textgridName$ - ".wav"

		hnrmean1 = Get mean: intStart, segEnd
		hnrmean1$ = string$: hnrmean1

		hnrmean2 = Get mean: middleStart, middleEnd
		hnrmean2$ = string$: hnrmean2

		hnrmean3 = Get mean: secondSegStart, intEnd
		hnrmean3$ = string$: hnrmean3

		
		# Jitter measurements

		selectObject: "PointProcess " + textgridName$ - ".wav"
	
		jitterSeg1 = Get jitter (local): intStart, segEnd, shortPerJitt, longPerJitt, maxPerFactJitt
		jitterSeg1$ = string$: jitterSeg1

		jitterMiddle = Get jitter (local): middleStart, middleEnd, shortPerJitt, longPerJitt, maxPerFactJitt
		jitterMiddle$ = string$: jitterMiddle

		jitterSecondSeg = Get jitter (local): secondSegStart, intEnd, shortPerJitt, longPerJitt, maxPerFactJitt
		jitterSecondSeg$ = string$: jitterSecondSeg
		

		# Shimmer measurements

		selectObject: "Sound " + textgridName$ - ".wav"
		plusObject: "PointProcess " + textgridName$ - ".wav"
		shimmerSeg1 = Get shimmer (local): intStart, segEnd, shortPerShim, longPerShim, maxPerFactShim, maxAmpFactShim
		shimmerSeg1$ = string$: shimmerSeg1

		shimmerMiddle = Get shimmer (local): middleStart, middleEnd, shortPerShim, longPerShim, maxPerFactShim, maxAmpFactShim
		shimmerMiddle$ = string$: shimmerMiddle

		shimmerSecondSeg = Get shimmer (local): secondSegStart, intEnd, shortPerShim, longPerShim, maxPerFactShim, maxAmpFactShim
		shimmerSecondSeg$ = string$: shimmerSecondSeg
		

	# Print out all of these measurements into the results table

  intInQuestRow$ = textgridName$ + sep$ + i$ + sep$ + intStart$ + sep$
		  	... + intLabel$ + sep$ + intCat$ + sep$ + intCod$ + sep$
		  	... + voiceType$ + sep$
		  	... + preIntLabel$ + sep$
		  	... + intDur$ + sep$ + segActualDur$ + sep$ + segDurType$ + sep$
		  	... + hnrmean1$ + sep$ + hnrmean2$ + sep$ + hnrmean3$ + sep$
		  	... + h1mnh2_seg1$ + sep$ + h1mnh2_middle$ + sep$ + h1mnh2_SecondSeg$ + sep$
		  	... + h1mna1_seg1$ + sep$ + h1mna1_middle$ + sep$ + h1mna1_SecondSeg$ + sep$ 
		  	... + h1mna2_seg1$ + sep$ + h1mna2_middle$ + sep$ + h1mna2_SecondSeg$ + sep$
		  	... + h1mna3_seg1$ + sep$ + h1mna3_middle$ + sep$ + h1mna3_SecondSeg$ + sep$
		  	... + jitterSeg1$ + sep$ + jitterMiddle$ + sep$ + jitterSecondSeg$ + sep$
		  	... + shimmerSeg1$ + sep$ + shimmerMiddle$ + sep$ + shimmerSecondSeg$ + sep$
		  	... + intMinSeg1$ + sep$ + intMinMiddle$ + sep$ + intMinSecondSeg$ + sep$
		  	... + pitchMinSeg1$ + sep$ + pitchMinMiddle$ + sep$ + pitchMinSecondSeg$ + sep$
		  	... + f1$ + sep$ + f1Mid$ + sep$ + f1Second$ + sep$
		  	... + f2$ + sep$ + f2Mid$ + sep$ + f2Second$ + sep$		  	
		  	... + f3$ + sep$ + f3Mid$ + sep$ + f3Second$
		  	... + newline$

		  appendFile: outFile$, intInQuestRow$ 	


# Clean up the open files
	selectObject: "Sound " + textgridName$ - ".wav" + "_slice1"
	plusObject: "Spectrum " + textgridName$ - ".wav" + "_slice1"
	plusObject: "Ltas " + textgridName$ - ".wav" + "_slice1"
	plusObject: "Sound " + textgridName$ - ".wav" + "_slice2"
	plusObject: "Spectrum " + textgridName$ - ".wav" + "_slice2"
	plusObject: "Ltas " + textgridName$ - ".wav" + "_slice2"
	plusObject: "Sound " + textgridName$ - ".wav" + "_slice3"
	plusObject: "Spectrum " + textgridName$ - ".wav" + "_slice3"
	plusObject: "Ltas " + textgridName$ - ".wav" + "_slice3"
	Remove

# if interval = "initial"
	endif

# If voice label is certain number
	endif	

# if interval stuff defined
	endif

# ends the for each interval
	endfor 

# Clean up remaining open files
 selectObject: "Formant " + textgridName$ - ".wav"
 plusObject: "Sound " + textgridName$ - ".wav"
 plusObject: "TextGrid " + textgridName$ - ".wav"
 plusObject: "Harmonicity " + textgridName$ - ".wav"
 plusObject: "Formant " + textgridName$ - ".wav" + "_aftertracking"
 plusObject: "Spectrogram " + textgridName$ - ".wav"
 plusObject: "Pitch " + textgridName$ - ".wav"
 plusObject: "Pitch " + textgridName$ - ".wav" + "_interpolated"
 plusObject: "Intensity " + textgridName$ - ".wav"
 plusObject: "PointProcess " + textgridName$ - ".wav"

 Remove


 goto endpoint


 label ceaseplease

 selectObject: "TextGrid " + textgridName$ - ".wav"
 Remove

 label endpoint


# if matching textgrid exists	
#	else appendInfoLine: "This sound file has no corresponding textgrid"
	endif

 # ends the for each wav file
	endfor

# if voicetype has particular value
# endif

# for each voicetype
endfor 

selectObject: "Strings relFiles"
Remove

pauseScript: "All done!"