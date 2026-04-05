<?php
/*
  +---------------------------------------------------------------------------+
  | ExcelCoreProperties                                                       |
  |                                                                           |
  | Reference file for NuSphere PHPEd (and possibly other IDE's) for use with |
  | php_excel interface to libxl by Ilia Alshanetsky <ilia@ilia.ws>           |
  |                                                                           |
  | php_excel "PECL" style module (http://github.com/iliaal/php_excel)        |
  | libxl library (http://www.libxl.com)                                      |
  |                                                                           |
  +---------------------------------------------------------------------------+
*/
class ExcelCoreProperties
{
	/**
	* Create a core properties object for a workbook
	*
	* @see ExcelBook::coreProperties()
	* @param ExcelBook $book
	*/
	public function __construct(ExcelBook $book)
	{
	}

	/**
	* Returns the title
	*
	* @return string|null
	*/
	public function title()
	{
	}

	/**
	* Sets the title
	*
	* @param string $value
	* @return bool
	*/
	public function setTitle($value): bool
	{
	}

	/**
	* Returns the subject
	*
	* @return string|null
	*/
	public function subject()
	{
	}

	/**
	* Sets the subject
	*
	* @param string $value
	* @return bool
	*/
	public function setSubject($value): bool
	{
	}

	/**
	* Returns the creator
	*
	* @return string|null
	*/
	public function creator()
	{
	}

	/**
	* Sets the creator
	*
	* @param string $value
	* @return bool
	*/
	public function setCreator($value): bool
	{
	}

	/**
	* Returns the last modified by value
	*
	* @return string|null
	*/
	public function lastModifiedBy()
	{
	}

	/**
	* Sets the last modified by value
	*
	* @param string $value
	* @return bool
	*/
	public function setLastModifiedBy($value): bool
	{
	}

	/**
	* Returns the created date as a string
	*
	* @return string|null
	*/
	public function created()
	{
	}

	/**
	* Sets the created date as a string
	*
	* @param string $value
	* @return bool
	*/
	public function setCreated($value): bool
	{
	}

	/**
	* Returns the modified date as a string
	*
	* @return string|null
	*/
	public function modified()
	{
	}

	/**
	* Sets the modified date as a string
	*
	* @param string $value
	* @return bool
	*/
	public function setModified($value): bool
	{
	}

	/**
	* Returns the tags
	*
	* @return string|null
	*/
	public function tags()
	{
	}

	/**
	* Sets the tags
	*
	* @param string $value
	* @return bool
	*/
	public function setTags($value): bool
	{
	}

	/**
	* Returns the categories
	*
	* @return string|null
	*/
	public function categories()
	{
	}

	/**
	* Sets the categories
	*
	* @param string $value
	* @return bool
	*/
	public function setCategories($value): bool
	{
	}

	/**
	* Returns the comments
	*
	* @return string|null
	*/
	public function comments()
	{
	}

	/**
	* Sets the comments
	*
	* @param string $value
	* @return bool
	*/
	public function setComments($value): bool
	{
	}

	/**
	* Returns the created date as a double (Excel timestamp)
	*
	* @return float
	*/
	public function createdAsDouble()
	{
	}

	/**
	* Sets the created date as a double (Excel timestamp)
	*
	* @param float $value
	* @return bool
	*/
	public function setCreatedAsDouble($value): bool
	{
	}

	/**
	* Returns the modified date as a double (Excel timestamp)
	*
	* @return float
	*/
	public function modifiedAsDouble()
	{
	}

	/**
	* Sets the modified date as a double (Excel timestamp)
	*
	* @param float $value
	* @return bool
	*/
	public function setModifiedAsDouble($value): bool
	{
	}

	/**
	* Removes all core properties
	*
	* @return void
	*/
	public function removeAll()
	{
	}

} // end ExcelCoreProperties
