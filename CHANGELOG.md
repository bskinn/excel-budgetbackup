# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](http://keepachangelog.com/en/1.0.0/)
and this project adheres to [Semantic Versioning](http://semver.org/spec/v2.0.0.html).

## Unreleased

...

## [2.0.0] - 2018-07-12

### Added

 * Form control and functionality to allow for variable padding of the item indices (#20).
 * Add 'Remove All', 'Insert All', and 'Append All' buttons and functionality (#18 & #19)
 * Implement improved syntax check for files that might be *supposed* to be
   budget items, but aren't quite formatted correctly. (#23) This does not lock out
   the file manipulations, because the possibly-malformed files are not populated
   to the included/excluded listboxes.
 * Implement aggressive check for potentially colliding filenames (identical
   except for the index) and lock out file manipulations until no collisions
   exist. (#7)
 * Implement safe detection/handling of changes to the underlying folder contents,
   via filename hashing check; warn and lock controls if a change is
   detected (#2, #15)
 * Implement check and surrounding safe-handling code for rename &c. attempts on
   a file that's open in another application (usually a PDF open for inspection).

### Fixed

 * Fix regex for filename parsing regex to allow decimal values
   in the 'quantity' field
 * Fix 'quantity' value conversion during sheet generation to allow
   decimal values
 * Disable Excel's complaining during sheet generation if the default
   workbook has more than one sheet, and the extras are deleted
 * Remove obsolete Adobe Reader location code & form (#1)

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
