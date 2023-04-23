# straight-font-to-unicode

**If you are looking to convert a document typed in "Straight", just download the transcoder_StraightFont.docm file and follow the instructions in the document. The document contains 2 macros. Try the "straight-donna" one first.**

## Contents

This repository contains:
1. [*.csv file] - a list of character mappings to convert the "Straight" font into Unicode compliant text
2. [*.py file] - a Python script that takes the list of character mappings and uses them to write a MS Word macro
3. [*.bas] - MS Word macros that convert documents typed in the legacy font "Straight" into a Unicode-compliant font
4. [*.docm] - a MS Word Macro-enabled Document, this is an easy way to get the macro working on your computer

## Fonts

In order to use the macros, the fonts "Straight" and "BC Sans" have to be installed. 

* BC Sans can be downloaded here:
https://www2.gov.bc.ca/gov/content/governments/services-for-government/policies-procedures/bc-visual-identity/bc-sans

* Straight font:
I can send you a copy of the font file if you email me. I also believe the macro will run if you have a font (ANY font) named "Straight" on your computer.

## Two versions of the macro

There are 2 macros designed to work with 2 different versions of the font. The original "Straight" font was created by the linguist Charles Ulrich sometime in the 1990s, when the Unicode standard did not contain all the characters needed to type APA phonetics. At the time, it only worked on Macintosh computers, so a Windows version was developed separately by another party which could work on Windows computers. The macros for them are named Straight-donna and Straight-barbara.

**The two fonts share the same name and appearance (glyphs), but NOT the same character mappings.**

Both font files have been updated to a modern format and can be used on any operating system. In general if the document originates from a Mac user, it is likely the original "Straight" font (use the straight-donna macro). If it originates from a Windows user it may be the other Straight font (use the straight-barbara macro). When in doubt just start with the "straight-donna" word macro and see if that works.

* straight-donna macro - use this for the original "Straight" font developed for Mac computers
* straight-barbara macro - use this for the later version of "Straight" developed for Windows computers
