<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

$excel_docs_directory = '../ExcelDocsForlder';
$generated_excel_docs_directory = '../generatedExcelDocs';


$excel_docs = glob($excel_docs_directory . "/*.xlsx");
foreach ($excel_docs as $doc) {
    try {
        $spreadsheet = $reader->load($doc);
        foreach ($spreadsheet->getAllSheets() as $index => $sheet) {
            $spreadsheet = new Spreadsheet();
            $activeSheet = $spreadsheet->getActiveSheet();
            $spreadsheet->addExternalSheet($sheet);
            $spreadsheet->setActiveSheetIndex(1);
            $spreadsheet->removeSheetByIndex(0);
            $writer = new Xlsx($spreadsheet);
            $new_doc_name = basename($doc, ".xlsx") . ($index + 1) . ".xlsx";
            $writer->save("$generated_excel_docs_directory/" . $new_doc_name);
        }
    } catch (Exception $e) {
        echo $e;
    }
}