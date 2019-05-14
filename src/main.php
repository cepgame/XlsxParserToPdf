<?php

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;

require '../vendor/autoload.php';
include "ParserToPdf.php";

$result = new ParserToPdf();
$result->parseToPdf("test.xlsx");

?>


