<?php
/*
  +---------------------------------------------------------------------------+
  | ExcelFormControl                                                          |
  |                                                                           |
  | Reference file for NuSphere PHPEd (and possibly other IDE's) for use with |
  | php_excel interface to libxl by Ilia Alshanetsky <ilia@ilia.ws>           |
  |                                                                           |
  | php_excel "PECL" style module (http://github.com/iliaal/php_excel)        |
  | libxl library (http://www.libxl.com)                                      |
  |                                                                           |
  +---------------------------------------------------------------------------+
*/
class ExcelFormControl
{
	const CHECKEDTYPE_UNCHECKED = 0;
	const CHECKEDTYPE_CHECKED = 1;
	const CHECKEDTYPE_MIXED = 2;

	const OBJECT_UNKNOWN = 0;
	const OBJECT_BUTTON = 1;
	const OBJECT_CHECKBOX = 2;
	const OBJECT_DROP = 3;
	const OBJECT_GBOX = 4;
	const OBJECT_LABEL = 5;
	const OBJECT_LIST = 6;
	const OBJECT_RADIO = 7;
	const OBJECT_SCROLL = 8;
	const OBJECT_SPIN = 9;
	const OBJECT_EDITBOX = 10;
	const OBJECT_DIALOG = 11;

	/**
	* Create a form control from a sheet and index
	*
	* @see ExcelSheet::formControl()
	* @param ExcelSheet $sheet
	* @param int $index
	*/
	public function __construct(ExcelSheet $sheet, $index)
	{
	}

	/**
	* Returns the object type of the form control
	*
	* @return int One of ExcelFormControl::OBJECT_* constants
	*/
	public function objectType()
	{
	}

	/**
	* Returns the checked state
	*
	* @return int One of ExcelFormControl::CHECKEDTYPE_* constants
	*/
	public function checked()
	{
	}

	/**
	* Sets the checked state
	*
	* @param int $value One of ExcelFormControl::CHECKEDTYPE_* constants
	* @return bool
	*/
	public function setChecked($value): bool
	{
	}

	/**
	* Returns the group box formula
	*
	* @return string|null
	*/
	public function fmlaGroup()
	{
	}

	/**
	* Sets the group box formula
	*
	* @param string $value
	* @return bool
	*/
	public function setFmlaGroup($value): bool
	{
	}

	/**
	* Returns the cell link formula
	*
	* @return string|null
	*/
	public function fmlaLink()
	{
	}

	/**
	* Sets the cell link formula
	*
	* @param string $value
	* @return bool
	*/
	public function setFmlaLink($value): bool
	{
	}

	/**
	* Returns the source range formula
	*
	* @return string|null
	*/
	public function fmlaRange()
	{
	}

	/**
	* Sets the source range formula
	*
	* @param string $value
	* @return bool
	*/
	public function setFmlaRange($value): bool
	{
	}

	/**
	* Returns the text box formula
	*
	* @return string|null
	*/
	public function fmlaTxbx()
	{
	}

	/**
	* Sets the text box formula
	*
	* @param string $value
	* @return bool
	*/
	public function setFmlaTxbx($value): bool
	{
	}

	/**
	* Returns the name of the form control
	*
	* @return string|null
	*/
	public function name()
	{
	}

	/**
	* Returns the linked cell reference
	*
	* @return string|null
	*/
	public function linkedCell()
	{
	}

	/**
	* Returns the list fill range reference
	*
	* @return string|null
	*/
	public function listFillRange()
	{
	}

	/**
	* Returns the macro name
	*
	* @return string|null
	*/
	public function macro()
	{
	}

	/**
	* Returns the alternative text
	*
	* @return string|null
	*/
	public function altText()
	{
	}

	/**
	* Returns whether the form control is locked
	*
	* @return bool
	*/
	public function locked()
	{
	}

	/**
	* Returns whether the form control has default size
	*
	* @return bool
	*/
	public function defaultSize()
	{
	}

	/**
	* Returns whether the form control is printed
	*
	* @return bool
	*/
	public function print()
	{
	}

	/**
	* Returns whether the form control is disabled
	*
	* @return bool
	*/
	public function disabled()
	{
	}

	/**
	* Returns the list item at the specified index
	*
	* @param int $index
	* @return string|null|false
	*/
	public function item(int $index): string|null|false
	{
	}

	/**
	* Returns the number of items in the list
	*
	* @return int
	*/
	public function itemSize()
	{
	}

	/**
	* Adds an item to the list
	*
	* @param string $value
	* @return bool
	*/
	public function addItem($value): bool
	{
	}

	/**
	* Inserts an item at the specified index
	*
	* @param int $index
	* @param string $value
	* @return bool
	*/
	public function insertItem(int $index, string $value): bool
	{
	}

	/**
	* Clears all items from the list
	*
	* @return void
	*/
	public function clearItems()
	{
	}

	/**
	* Returns the number of drop lines
	*
	* @return int
	*/
	public function dropLines()
	{
	}

	/**
	* Sets the number of drop lines
	*
	* @param int $value
	* @return bool
	*/
	public function setDropLines($value): bool
	{
	}

	/**
	* Returns the scroll bar width
	*
	* @return int
	*/
	public function dx()
	{
	}

	/**
	* Sets the scroll bar width
	*
	* @param int $value
	* @return bool
	*/
	public function setDx($value): bool
	{
	}

	/**
	* Returns whether the first button is selected
	*
	* @return bool
	*/
	public function firstButton()
	{
	}

	/**
	* Sets whether the first button is selected
	*
	* @param bool $value
	* @return bool
	*/
	public function setFirstButton($value): bool
	{
	}

	/**
	* Returns whether the scroll bar is horizontal
	*
	* @return bool
	*/
	public function horiz()
	{
	}

	/**
	* Sets whether the scroll bar is horizontal
	*
	* @param bool $value
	* @return bool
	*/
	public function setHoriz($value): bool
	{
	}

	/**
	* Returns the increment value
	*
	* @return int
	*/
	public function inc()
	{
	}

	/**
	* Sets the increment value
	*
	* @param int $value
	* @return bool
	*/
	public function setInc($value): bool
	{
	}

	/**
	* Returns the maximum value
	*
	* @return int
	*/
	public function getMax()
	{
	}

	/**
	* Sets the maximum value
	*
	* @param int $value
	* @return bool
	*/
	public function setMax($value): bool
	{
	}

	/**
	* Returns the minimum value
	*
	* @return int
	*/
	public function getMin()
	{
	}

	/**
	* Sets the minimum value
	*
	* @param int $value
	* @return bool
	*/
	public function setMin($value): bool
	{
	}

	/**
	* Returns the multi-selection mode string
	*
	* @return string|null
	*/
	public function multiSel()
	{
	}

	/**
	* Sets the multi-selection mode string
	*
	* @param string $value
	* @return bool
	*/
	public function setMultiSel($value): bool
	{
	}

	/**
	* Returns the selected index
	*
	* @return int
	*/
	public function sel()
	{
	}

	/**
	* Sets the selected index
	*
	* @param int $value
	* @return bool
	*/
	public function setSel($value): bool
	{
	}

	/**
	* Returns the from-anchor position
	*
	* @return array
	*/
	public function fromAnchor()
	{
	}

	/**
	* Returns the to-anchor position
	*
	* @return array
	*/
	public function toAnchor()
	{
	}

} // end ExcelFormControl
