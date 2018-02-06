<?php

use PhpOffice\PhpSpreadsheet\Helper\Sample;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

require_once __DIR__ . '/../../src/Bootstrap.php';

$helper = new Sample();
if ($helper->isCli()) {
    $helper->log('This example should only be run from a Web Browser' . PHP_EOL);

    return;
}

// Create new Spreadsheet object
$spreadsheet = new Spreadsheet();

// Set document properties
$spreadsheet->getProperties()->setCreator('Jangwon Seo')
        ->setLastModifiedBy('Abraham (Jangwon) Seo')
        ->setTitle('PHPSpreadSheet XLSX Test')
        ->setSubject('PHPSpreadSheet Invoice')
        ->setDescription('PHPSpreadSheet Invoice')
        ->setKeywords('office php invoice')
        ->setCategory('Invoice');

$helper->log('Add some data');
$spreadsheet->setActiveSheetIndex(0);
$spreadsheet->getActiveSheet()->setCellValue('A1', 'Samjung Data Service Co., Ltd.');
$spreadsheet->getActiveSheet()->setCellValue('A2', '2-6301, 110, Digital-ro 26-gil, Guro-gu, Seoul, Republic of Korea');
$spreadsheet->getActiveSheet()->setCellValue('A3', 'TEL :+82-2-2029-0130 / FAX :+82-2-544-6008');
$spreadsheet->getActiveSheet()->setCellValue('A5', 'INVOICE');

$spreadsheet->getActiveSheet()->setCellValue('A8', 'BUYER : ');
$spreadsheet->getActiveSheet()->setCellValue('A9', 'Twilio Inc. ');
$spreadsheet->getActiveSheet()->setCellValue('A10', '645 Harrison Street, Third Floor,');
$spreadsheet->getActiveSheet()->setCellValue('A11', 'San Francisco, CA 94107, USA');

$spreadsheet->getActiveSheet()->setCellValue('G8', 'Invoice No. 20170704');
$spreadsheet->getActiveSheet()->setCellValue('G9', 'Date : 04. July, 2017');

$spreadsheet->getActiveSheet()->setCellValue('A13', 'Item');
$spreadsheet->getActiveSheet()->setCellValue('B13', 'Description');
$spreadsheet->getActiveSheet()->setCellValue('C13', 'Qty');
$spreadsheet->getActiveSheet()->setCellValue('D13', 'Time');
$spreadsheet->getActiveSheet()->setCellValue('E13', 'Unit price(USD)');
$spreadsheet->getActiveSheet()->setCellValue('F13', 'Amount(USD)');

$spreadsheet->getActiveSheet()->setCellValue('A14', '1');
$spreadsheet->getActiveSheet()->setCellValue('B14', 'Virtual Long Number (June 2017)');
$spreadsheet->getActiveSheet()->setCellValue('C14', '396');
$spreadsheet->getActiveSheet()->setCellValue('D14', '-');
$spreadsheet->getActiveSheet()->setCellValue('E14', '1');
$spreadsheet->getActiveSheet()->setCellValue('F14', '396.00');

$spreadsheet->getActiveSheet()->setCellValue('A15', '2');
$spreadsheet->getActiveSheet()->setCellValue('B15', 'Voice Inbound (Samjung to Twilio)');
$spreadsheet->getActiveSheet()->setCellValue('C15', '-');
$spreadsheet->getActiveSheet()->setCellValue('D15', '81 min');
$spreadsheet->getActiveSheet()->setCellValue('E15', '0.008');
$spreadsheet->getActiveSheet()->setCellValue('F15', '0.648');

$spreadsheet->getActiveSheet()->setCellValue('A16', '3');
$spreadsheet->getActiveSheet()->setCellValue('B16', 'Voice Outbound: Land line (Twilio to Samjung)');
$spreadsheet->getActiveSheet()->setCellValue('C16', '-');
$spreadsheet->getActiveSheet()->setCellValue('D16', '0 min');
$spreadsheet->getActiveSheet()->setCellValue('E16', '0.0153');
$spreadsheet->getActiveSheet()->setCellValue('F16', '0.0000');

$spreadsheet->getActiveSheet()->setCellValue('A17', '4');
$spreadsheet->getActiveSheet()->setCellValue('B17', 'Voice Outbound: Mobile (Twilio to Samjung)');
$spreadsheet->getActiveSheet()->setCellValue('C17', '-');
$spreadsheet->getActiveSheet()->setCellValue('D17', '0 dsec');
$spreadsheet->getActiveSheet()->setCellValue('E17', '0.009');
$spreadsheet->getActiveSheet()->setCellValue('F17', '0.000');

$spreadsheet->getActiveSheet()->setCellValue('G28', 'TOTAL AMOUNT');
$spreadsheet->getActiveSheet()->setCellValue('G29', 'US$396.648');


// Add comment
$helper->log('Add comments');

$spreadsheet->getActiveSheet()->getComment('E11')->setAuthor('Jangwon Seo');
$spreadsheet->getActiveSheet()->getComment('E11')->getText()->createTextRun('0.008 USD per minute');

$spreadsheet->getActiveSheet()->getComment('E12')->setAuthor('Jangwon Seo');
$spreadsheet->getActiveSheet()->getComment('E12')->getText()->createTextRun('0.0153 USD per minute');

$spreadsheet->getActiveSheet()->getComment('E13')->setAuthor('Jangwon Seo');
$spreadsheet->getActiveSheet()->getComment('E13')->getText()->createTextRun('0.009 USD per 10 sec (dsec)');
        
// Merge cells
$helper->log('Merge cells');
$spreadsheet->getActiveSheet()->mergeCells('A1:G1');
$spreadsheet->getActiveSheet()->mergeCells('A2:G2');
$spreadsheet->getActiveSheet()->mergeCells('A3:G3');
$spreadsheet->getActiveSheet()->mergeCells('A5:G5');

// Rename worksheet
$spreadsheet->getActiveSheet()->setTitle('Invoice 20170704');
$spreadsheet->getActiveSheet()->setShowGridLines(false);

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

IOFactory::registerWriter('Pdf', \PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf::class);

// Redirect output to a client’s web browser (PDF)
header('Content-Type: application/pdf');
header('Content-Disposition: attachment;filename="01simple.pdf"');
header('Cache-Control: max-age=0');

$writer = IOFactory::createWriter($spreadsheet, 'Pdf');
$writer->save('php://output');
exit;
