<?php

use PhpOffice\PhpSpreadsheet\Calculation\Statistical;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
require '../vendor/autoload.php';

$reader = new Xlsx();
$sheet = $reader->load("test.xlsx");
$spreadsheet = $sheet->getActiveSheet();
echo Statistical::COUNTA($spreadsheet, "K2", "K1000");
?>