# CopyDataOver6MinLimit
Google apps script for Google sheets

input: source spreadsheet ID, source sheet name, target sheet name

Purpose: Script copies rows that have unique values on certain column from source sheet
then pastes them into target sheet, 
because this script operates over large data 
it works around the 6 minutes execution time limit set by Google by saving progress to UserCache and continues on next start of script.

  
