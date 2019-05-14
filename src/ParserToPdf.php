<?php
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Calculation\Statistical;
use PhpOffice\PhpSpreadsheet\Calculation\TextData;


include "../vendor/tecnickcom/tcpdf/tcpdf.php";
define ('PDF_HEADER_MOSPOLY', '../resources/header.jpg');
define ('PDF_HEADER_NAME', '© 2019 Moscow State Polytechnic University');


class ParserToPdf
{
    static public $TITLE_COLUMN="K";
    public function parseToPdf($filePath)
    {
        $file = $this->getFile($filePath);
        $count= $this->getCountOfProjects($file);
        for ($i=2; $i<($count+1); $i+=1) {
            $title = $this->getTitle($i,$file);
            $body = $this->createPdfBody($i,$file);
            $name = ($i-1).".pdf";
            $this->createPdf($title, $body, $name);
        }
    }
    function getFile($filePath)
    {
        $reader = new Xlsx();
        $sheet = $reader->load($filePath);
        $spreadsheet = $sheet->getActiveSheet();
        return $spreadsheet;
    }

    function getCountOfProjects($sheet)
    {
        return Statistical::COUNTA($sheet, "K2", "K1000")+1;

    }

    function getTitle($headnum, $file)
    {
        $head = $file->getCell(self::$TITLE_COLUMN.$headnum)->getValue();
        return $head;
    }

    function createPdfBody($bodynum, $file)
    {
        $piceOfBody = $this->searchBody($bodynum,$file);
        $direction = $this->searchDirection($bodynum,$file);
        $body= $direction."\n".$piceOfBody."\n";
        echo $direction;
        return $body;
    }

    function searchBody($bodynum, $loadSheet)
    {
        $body="";
        for ($column="L"; $column>="AC"; ++$column)
        {
            $columnValue = $loadSheet->getCell($column."1")->getValue();
            $bodyValue = $loadSheet->getCell($column.$bodynum)->getValue();
            if ($bodyValue=="")
                $body.= "\n".$bodyValue;
            elseif ($columnValue==""){
                $body.="\n".$bodyValue."\n";
            }
            else

                $body.="\n".$columnValue.":"."\n".$bodyValue."\n";
        }
        return $body;
    }

    function searchDirection($bodyNum, $loadSheet)
    {
        $direction="";
        for ($column="AD"; $column>="CQ"; ++$column)
        {
            $columnValue = $loadSheet->getCell($column."1")->getValue();
            $directionValue = $loadSheet->getCell($column.$bodyNum)->getValue();
            if ($directionValue !="")
                $direction.=$columnValue.", ";
        }
        echo $direction;
        return $direction;
    }


    function createPdf($head, $body, $name)
    {
        $pdf = new TCPDF(PDF_PAGE_ORIENTATION, PDF_UNIT, PDF_PAGE_FORMAT, true, 'UTF-8', false);
        /*
        $pdf = new TCPDF(PDF_PAGE_ORIENTATION, PDF_UNIT, PDF_PAGE_FORMAT, true, 'UTF-8', false);
        $pdf->AddPage();
        $pdf->SetFont('freeserif', '', 20, '', true);
        $pdf->Write(0, $head, '', 0, 'C', true, 0, false, false, 0);
        $pdf->Ln(5);

        $pdf->SetFont('times', '', 16, '', false);
        $pdf->Write(0, $body, '', 0, 'C', true, 0, true, false, 0);
        $pdf->Output("F:/games/php sheeds/src/".$name,'F');
        */

        $pdf->SetHeaderData(PDF_HEADER_MOSPOLY, PDF_HEADER_LOGO_WIDTH, PDF_HEADER_NAME , "", array(0,64,255), array(0,64,128));

        $pdf->SetFont('freeserif', '"b"', 16);
        $pdf->AddPage();
        $pdf->Write(0, $head, '', 0, 'L', true, 0, false, false, 0);
        $pdf->Ln(5);
        $pdf->SetFont('freeserif', '', 12);
        $pdf->startTransaction();
        $pdf->Write(0, $body);
        $pdf->Output("F:/games/php sheeds/src/".$name,'F');
      }

}