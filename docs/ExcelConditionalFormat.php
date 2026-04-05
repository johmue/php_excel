<?php
/*
  +---------------------------------------------------------------------------+
  | ExcelConditionalFormat                                                    |
  |                                                                           |
  | Reference file for NuSphere PHPEd (and possibly other IDE's) for use with |
  | php_excel interface to libxl by Ilia Alshanetsky <ilia@ilia.ws>           |
  |                                                                           |
  | php_excel "PECL" style module (http://github.com/iliaal/php_excel)        |
  | libxl library (http://www.libxl.com)                                      |
  |                                                                           |
  +---------------------------------------------------------------------------+
*/
class ExcelConditionalFormat
{
	/**
	* Create a conditional format within an Excel workbook
	*
	* @see ExcelBook::addConditionalFormat()
	* @param ExcelBook $book
	*/
	public function __construct(ExcelBook $book)
	{
	}

	/**
	* Returns the font for this conditional format
	*
	* @return ExcelFont
	*/
	public function font()
	{
	}

	/**
	* Returns the number format identifier
	*
	* @return int
	*/
	public function numFormat()
	{
	}

	/**
	* Sets the number format identifier
	*
	* @param int $value
	* @return bool
	*/
	public function setNumFormat($value): bool
	{
	}

	/**
	* Returns the custom number format string
	*
	* @return string
	*/
	public function customNumFormat()
	{
	}

	/**
	* Sets the custom number format string
	*
	* @param string $value
	* @return bool
	*/
	public function setCustomNumFormat($value): bool
	{
	}

	/**
	* Sets the border style on all sides
	*
	* @param int $value One of ExcelFormat::BORDERSTYLE_* constants
	* @return bool
	*/
	public function setBorder($value): bool
	{
	}

	/**
	* Sets the border color on all sides
	*
	* @param int $value One of ExcelFormat::COLOR_* constants
	* @return bool
	*/
	public function setBorderColor($value): bool
	{
	}

	/**
	* Returns the left border style
	*
	* @return int
	*/
	public function borderLeft()
	{
	}

	/**
	* Sets the left border style
	*
	* @param int $value One of ExcelFormat::BORDERSTYLE_* constants
	* @return bool
	*/
	public function setBorderLeft($value): bool
	{
	}

	/**
	* Returns the right border style
	*
	* @return int
	*/
	public function borderRight()
	{
	}

	/**
	* Sets the right border style
	*
	* @param int $value One of ExcelFormat::BORDERSTYLE_* constants
	* @return bool
	*/
	public function setBorderRight($value): bool
	{
	}

	/**
	* Returns the top border style
	*
	* @return int
	*/
	public function borderTop()
	{
	}

	/**
	* Sets the top border style
	*
	* @param int $value One of ExcelFormat::BORDERSTYLE_* constants
	* @return bool
	*/
	public function setBorderTop($value): bool
	{
	}

	/**
	* Returns the bottom border style
	*
	* @return int
	*/
	public function borderBottom()
	{
	}

	/**
	* Sets the bottom border style
	*
	* @param int $value One of ExcelFormat::BORDERSTYLE_* constants
	* @return bool
	*/
	public function setBorderBottom($value): bool
	{
	}

	/**
	* Returns the left border color
	*
	* @return int
	*/
	public function borderLeftColor()
	{
	}

	/**
	* Sets the left border color
	*
	* @param int $value One of ExcelFormat::COLOR_* constants
	* @return bool
	*/
	public function setBorderLeftColor($value): bool
	{
	}

	/**
	* Returns the right border color
	*
	* @return int
	*/
	public function borderRightColor()
	{
	}

	/**
	* Sets the right border color
	*
	* @param int $value One of ExcelFormat::COLOR_* constants
	* @return bool
	*/
	public function setBorderRightColor($value): bool
	{
	}

	/**
	* Returns the top border color
	*
	* @return int
	*/
	public function borderTopColor()
	{
	}

	/**
	* Sets the top border color
	*
	* @param int $value One of ExcelFormat::COLOR_* constants
	* @return bool
	*/
	public function setBorderTopColor($value): bool
	{
	}

	/**
	* Returns the bottom border color
	*
	* @return int
	*/
	public function borderBottomColor()
	{
	}

	/**
	* Sets the bottom border color
	*
	* @param int $value One of ExcelFormat::COLOR_* constants
	* @return bool
	*/
	public function setBorderBottomColor($value): bool
	{
	}

	/**
	* Returns the fill pattern
	*
	* @return int
	*/
	public function fillPattern()
	{
	}

	/**
	* Sets the fill pattern
	*
	* @param int $value One of ExcelFormat::FILLPATTERN_* constants
	* @return bool
	*/
	public function setFillPattern($value): bool
	{
	}

	/**
	* Returns the pattern foreground color
	*
	* @return int
	*/
	public function patternForegroundColor()
	{
	}

	/**
	* Sets the pattern foreground color
	*
	* @param int $value One of ExcelFormat::COLOR_* constants
	* @return bool
	*/
	public function setPatternForegroundColor($value): bool
	{
	}

	/**
	* Returns the pattern background color
	*
	* @return int
	*/
	public function patternBackgroundColor()
	{
	}

	/**
	* Sets the pattern background color
	*
	* @param int $value One of ExcelFormat::COLOR_* constants
	* @return bool
	*/
	public function setPatternBackgroundColor($value): bool
	{
	}

} // end ExcelConditionalFormat
