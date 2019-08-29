# ExcelMLS
Excel add-on to support multi-language applications

This Excel Add-in (xlam file) offers advanced functionality for excel spreadsheets or applications with are meant to be used in a multi-language environment. It offers simple functions to display text elements as e.g. cell content, (shape) captions and messages in several languages by just selection of a language setting. This language setting can be independent/different than the selected Excel language setting, e.g. you can run an application in German on an English based Excel installation or vice versa. 

This language setting can be based on the selected excel user interface language but is in general independent from such setting. It’s also possible to have some text elements bound to the excel system language, while others are dependent of the excel user language.

This excel add-in also offers advanced development support with advanced editing functionality that is integrated with both the excel and the excel VBE environment. These features are the most valuable elements of this add-in. Cross-referencing usage of text elements in cells, shapes and VBA (macro) code is also supported as well as widow (unused text elements) and orphan (undefined text elements) control.

It consists of a two-level database (table) of text elements, identified and selected by a key – module (larger application area) and identifier (identifying a specific element). The first level (optional) is a table usually placed in the specific excel workbook and a second table placed in the add-in itself. If a specific text element is requested and is available in the first level this will get returned. If not the second level will get searched and any find returned – else an error/default text. This allows to have a set of generic text messages that can be extended od superseded by application/document specific text elements.

The (sorted) tables are held in memory for fast access to all text elements.

For more information including installation see the provided documentation.
