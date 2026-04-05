<?php
/*
  +---------------------------------------------------------------------------+
  | ExcelRichString                                                           |
  |                                                                           |
  | Reference file for NuSphere PHPEd (and possibly other IDE's) for use with |
  | php_excel interface to libxl by Ilia Alshanetsky <ilia@ilia.ws>           |
  |                                                                           |
  | php_excel "PECL" style module (http://github.com/iliaal/php_excel)        |
  | libxl library (http://www.libxl.com)                                      |
  |                                                                           |
  +---------------------------------------------------------------------------+
*/
class ExcelRichString
{
	/**
	* Create a rich string within an Excel workbook
	*
	* @see ExcelBook::addRichString()
	* @param ExcelBook $book
	*/
	public function __construct(ExcelBook $book)
	{
	}

	/**
	* Adds a font to the rich string. Returns a new font for use with addText().
	*
	* @param ExcelFont|null $font (optional) Font to copy from
	* @return ExcelFont|false
	*/
	public function addFont(?ExcelFont $font = null): ExcelFont|false
	{
	}

	/**
	* Adds text to the rich string with an optional font
	*
	* @param string $text
	* @param ExcelFont|null $font (optional)
	* @return bool
	*/
	public function addText(string $text, ?ExcelFont $font = null): bool
	{
	}

	/**
	* Returns the text and font at the specified index
	*
	* @param int $index
	* @return array|false Array with keys "text"(string) and "font"(ExcelFont)
	*/
	public function getText(int $index): array|false
	{
	}

	/**
	* Returns the number of text segments in the rich string
	*
	* @return int|false
	*/
	public function textSize(): int|false
	{
	}

} // end ExcelRichString
