<?php
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Calculation\Statistical;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Calculation\TextData;


include "../vendor/tecnickcom/tcpdf/tcpdf.php";

class ParserToPdf
{
    public function parseToPdf($file)
    {
        $sheet = $this->getFile($file);
        $count = $this->countOfProjects($sheet);
        for ($i=1; $i<$count; $i+=1) {
            $head = $this->getHead($i,$sheet);
            $body = $this->getBody($i,$sheet);
            $name = "pdffile".$i;
            $this->createPdf($head, $body, $name);
        }
    }
    function getFile($file)
    {
        $reader = new Xlsx();
        $spreadsheet = $reader->load("$file");
        return $spreadsheet;
    }
    function memSheed($readsheet)
    {
        $spreadsheet = new Spreadsheet();
        $worksheed = $spreadsheet->getActiveSheet();
        $worksheed->fromArray($readsheet);
        return $worksheed;
    }

    function countOfProjects($sheet)
    {
        return Statistical::COUNTA($sheet, "A2", "A1000");
    }

    function getHead($headnum,$file)
    {
        $readsheet = $this->getFile($file);
        $loadsheet = $this->memSheed($readsheet);
        $head = $loadsheet->getCell("A".$headnum)->getValue();
        return $head;
    }

    function getBody($bodynum,$file)
    {
        $body = $this->searchBody($bodynum,$file);
        return $body;
    }

    function searchBody($bodynum,$file)
    {
        $readsheet = $this->getFile($file);
        $loadsheet = $this->memSheed($readsheet);
        $pieceofbody = TextData::CONCATENATE($loadsheet;"B1";':';"B".$bodynum;"\n";"C1";':';"C".$bodynum;"\n";"D1";':';"D".$bodynum;"\n");
    }


    function createPdf($head, $body, $name)
    {
        $pdf = new TCPDF(PDF_PAGE_ORIENTATION, PDF_UNIT, PDF_PAGE_FORMAT, true, 'UTF-8', false);
        $pdf->AddPage();
        $pdf->SetFont('times', '', 20, '', true);
        $pdf->Write(0, $head, '', 0, 'C', true, 0, false, false, 0);
        $pdf->SetFont('times', '', 16, '', false);
        $pdf->Write(0, $body, '', 0, 'C', true, 0, true, false, 0);
        $pdf->Output("F:/games/php sheeds/src/".$name,'F');
    }

}