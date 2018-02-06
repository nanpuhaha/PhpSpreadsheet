<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;

require __DIR__ . '/../Header.php';

$helper->log('Load from Xlsx template');
$reader = IOFactory::createReader('Xlsx');
$spreadsheet = $reader->load(__DIR__ . '/../templates/twilio_invoice.xlsx');

$helper->log('Add new data to the template');

$numberData = [[
    'description' => 'Virtual Long Number (June 2017)', // 뒤에 날짜
    'quantity' => 1234,
    'unitprice' => 1.2
]];

$voiceData = [
    [   'description' => 'Voice Inbound (Samjung to Twilio)', // 회사명
        'time' => 123,
        'unitprice' => 0.009,
    ],
    [   'description' => 'Voice Outbound: Land line (Twilio to Samjung)', // 회사명
        'time' => 52346,
        'unitprice' => 0.0531,
    ],
    [   'description' => 'Voice Outbound: Mobile (Twilio to Samjung)', // 회사명
        'time' => 3546,
        'unitprice' => 0.004,
    ]
];

$spreadsheet->getActiveSheet()
            ->setCellValue('A9', 'Twilio Inc.') // 회사명
            ->setCellValue('A10', '645 Harrison Street, Third Floor,')  // 주소1
            ->setCellValue('A11', 'San Francisco, CA 94107, USA');  // 주소2

$spreadsheet->getActiveSheet()
            ->setCellValue('G8', 'Invoice No. 20170704')    // 청구서 번호
            ->setCellValue('G9', 'Date : 04, July, 2017');  // 날짜
            
$baseRow = 14;
foreach ($numberData as $r => $dataRow) {
    $row = $baseRow + $r;
    $spreadsheet->getActiveSheet()
            ->setCellValue('B' . $row, $dataRow['description'])
            ->setCellValue('D' . $row, $dataRow['quantity'])
            ->setCellValue('F' . $row, $dataRow['unitprice'])
            ->setCellValue('G' . $row, '=D' . $row . '*F' . $row);
}

$baseRow = 15;
foreach ($voiceData as $r => $dataRow) {
    $row = $baseRow + $r;
    $spreadsheet->getActiveSheet()
            ->setCellValue('B' . $row, $dataRow['description'])
            ->setCellValue('E' . $row, $dataRow['time'])
            ->setCellValue('F' . $row, $dataRow['unitprice'])
            ->setCellValue('G' . $row, '=E' . $row . '*F' . $row);
}

// Save
$helper->write($spreadsheet, __FILE__);




// use PhpOffice\PhpSpreadsheet\Helper\Sample;
// use PhpOffice\PhpSpreadsheet\Spreadsheet;

// // require_once __DIR__ . '/../../src/Bootstrap.php';

// // $helper = new Sample();
// if ($helper->isCli()) {
//     $helper->log('This example should only be run from a Web Browser' . PHP_EOL);

//     return;
// }

// // Create new Spreadsheet object
// $spreadsheet2 = new Spreadsheet();
// $spreadsheet2 = $spreadsheet;

// IOFactory::registerWriter('Pdf', \PhpOffice\PhpSpreadsheet\Writer\Pdf\Mpdf::class);

// // Redirect output to a client’s web browser (PDF)
// header('Content-Type: application/pdf');
// header('Content-Disposition: attachment;filename="30_Template.pdf"');
// header('Cache-Control: max-age=0');

// $writer = IOFactory::createWriter($spreadsheet2, 'Pdf');
// $writer->save('php://output');
