## php_excel

[![Tests](https://github.com/iliaal/php_excel/actions/workflows/tests.yml/badge.svg)](https://github.com/iliaal/php_excel/actions/workflows/tests.yml)
[![Windows Build](https://github.com/iliaal/php_excel/actions/workflows/windows.yml/badge.svg)](https://github.com/iliaal/php_excel/actions/workflows/windows.yml)
[![Version](https://img.shields.io/github/v/release/iliaal/php_excel)](https://github.com/iliaal/php_excel/releases)
[![License: PHP-3.01](https://img.shields.io/badge/License-PHP--3.01-green.svg)](http://www.php.net/license/3_01.txt)
[![Follow @iliaa](https://img.shields.io/badge/Follow-@iliaa-000000?style=flat&logo=x&logoColor=white)](https://x.com/intent/follow?screen_name=iliaa)

PHP extension for reading and writing Excel files (XLS and XLSX) using the [LibXL](http://www.libxl.com/) library.

### Requirements

* PHP 8.3+
* [LibXL](http://www.libxl.com/) 4.6.0+ (commercial library)

### Classes

| Class | Description |
|-------|-------------|
| ExcelBook | Workbook management: create, load, save, sheets, fonts, formats, pictures |
| ExcelSheet | Cell read/write, formatting, printing, protection, hyperlinks, data validation |
| ExcelFormat | Cell formatting: colors, borders, number formats, alignment, patterns |
| ExcelFont | Font properties: name, size, bold, italic, underline, color |
| ExcelAutoFilter | AutoFilter operations and sorting |
| ExcelFilterColumn | Filter column criteria |
| ExcelRichString | Mixed-font text in a single cell |
| ExcelFormControl | Form controls: checkboxes, dropdowns, spinners, buttons |
| ExcelConditionalFormat | Conditional formatting style rules |
| ExcelConditionalFormatting | Conditional formatting ranges and rule application |
| ExcelCoreProperties | Workbook metadata: title, author, dates, categories |
| ExcelTable | Structured table support (xlsx) |

### Installation

Via [PIE](https://github.com/php/pie):

```sh
pie install iliaal/php-excel \
  --with-libxl-incdir=/path/to/libxl/include_c \
  --with-libxl-libdir=/path/to/libxl/lib64
```

Or manually:

```sh
phpize
./configure --with-excel \
  --with-libxl-incdir=/path/to/libxl/include_c \
  --with-libxl-libdir=/path/to/libxl/lib64
make
make install
```

Add `extension=excel.so` to your `php.ini`.

### Getting started

```php
<?php
$book = new ExcelBook(null, null, true); // xlsx mode
$book->setLocale('UTF-8');

$sheet = $book->addSheet('Sheet1');

$data = [
    [1, 1500, 'John', 'Doe'],
    [2,  750, 'Jane', 'Doe'],
];

$row = 1;
foreach ($data as $item) {
    $sheet->writeRow($row++, $item);
}

// formula
$sheet->write($row, 1, '=SUM(B1:B3)');

// date with format
$dateFormat = new ExcelFormat($book);
$dateFormat->numberFormat(ExcelFormat::NUMFORMAT_DATE);
$sheet2 = $book->addSheet('Sheet2');
$sheet2->write(1, 0, (new DateTime('2024-08-02'))->getTimestamp(), $dateFormat, ExcelFormat::AS_DATE);

$book->save('output.xlsx');
```

### php.ini settings

Store credentials in `php.ini` instead of source code. The extension reads them automatically when you pass `null` to the constructor.

```ini
[excel]
excel.license_name="YOUR_LICENSE_NAME"
excel.license_key="YOUR_LICENSE_KEY"
excel.skip_empty=0
```

### Documentation

See the `docs/` directory for API reference and the `tests/` directory for usage examples.

### Resources

* [LibXL library (commercial)](http://www.libxl.com/)
* [Contributing](CONTRIBUTOR.md)
