--TEST--
ExcelFormControl: formControlSize on empty sheet
--SKIPIF--
<?php if (!extension_loaded("excel")) print "skip"; ?>
--FILE--
<?php
$book = new ExcelBook(null, null, true);
$sheet = $book->addSheet("Sheet1");

var_dump($sheet->formControlSize());

echo "OK\n";
?>
--EXPECT--
int(0)
OK
