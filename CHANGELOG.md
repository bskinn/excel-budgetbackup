# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## [Unreleased]

 * Fix regex for filename parsing regex to allow decimal values
   in the 'quantity' field
 * Fix 'quantity' value conversion during sheet generation to allow
   decimal values
 * Disable Excel's complaining during sheet generation if the default
   workbook has more than one sheet, and the extras are deleted

## [1.0.0] - 2017-07-21

Existing production version released as open source.

### Current Features

 * Free selection of folder to search for source files
 * Reload of selected folder to refresh filenames is available
 * Can open selected folder or file directly from GUI
 * Reordering of included items by single steps up/down or
   by relocating to a specific point in the list
 * Excluded items can be added to the 'included' list either
   by inserting at the cursor or by appending
 * Automatic generation of a budget summary spreadsheet with
   unit and extended prices, subtotals by category, and
   a grand total.
