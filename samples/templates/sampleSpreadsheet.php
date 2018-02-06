<?php

// Create new Spreadsheet object
use PhpOffice\PhpSpreadsheet\RichText\RichText;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Font;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Protection;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;

$helper->log('Create new Spreadsheet object');
$spreadsheet = new Spreadsheet();

// Set document properties
$helper->log('Set document properties');
$spreadsheet->getProperties()->setCreator('Jangwon Seo')
        ->setLastModifiedBy('Abraham (Jangwon) Seo')
        ->setTitle('PHPSpreadSheet XLSX Test')
        ->setSubject('PHPSpreadSheet Invoice')
        ->setDescription('PHPSpreadSheet Invoice')
        ->setKeywords('office php invoice')
        ->setCategory('Invoice');

// Create a first sheet, representing sales data
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


// Add rich-text string
$helper->log('Add rich-text string');
$richText = new RichText();
$richText->createText('This invoice is ');

$payable = $richText->createTextRun('payable within thirty days after the end of the month');
$payable->getFont()->setBold(true);
$payable->getFont()->setItalic(true);
$payable->getFont()->setColor(new Color(Color::COLOR_DARKGREEN));

$richText->createText(', unless specified otherwise on the invoice.');

$spreadsheet->getActiveSheet()->getCell('A18')->setValue($richText);

// Merge cells
$helper->log('Merge cells');
$spreadsheet->getActiveSheet()->mergeCells('A1:G1');
$spreadsheet->getActiveSheet()->mergeCells('A2:G2');
$spreadsheet->getActiveSheet()->mergeCells('A3:G3');
$spreadsheet->getActiveSheet()->mergeCells('A5:G5');

// Protect cells
$helper->log('Protect cells');
$spreadsheet->getActiveSheet()->getProtection()->setSheet(true); // Needs to be set to true in order to enable any worksheet protection!
$spreadsheet->getActiveSheet()->protectCells('A13:G18', 'PhpSpreadsheet');  // 데이터 부분만
// $spreadsheet->getActiveSheet()->protectCells('A1:G29', 'PhpSpreadsheet'); // 전체

// Set cell number formats
$helper->log('Set cell number formats');
$spreadsheet->getActiveSheet()->getStyle('E4:E13')->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);

// Set column widths
$helper->log('Set column widths');
$spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->setWidth(12);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->setWidth(12);

// Set fonts
$helper->log('Set fonts');
$spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->setName('Candara');
$spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->setSize(20);
$spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->setBold(true);
$spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->setUnderline(Font::UNDERLINE_SINGLE);
$spreadsheet->getActiveSheet()->getStyle('B1')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

$spreadsheet->getActiveSheet()->getStyle('D1')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);
$spreadsheet->getActiveSheet()->getStyle('E1')->getFont()->getColor()->setARGB(Color::COLOR_WHITE);

$spreadsheet->getActiveSheet()->getStyle('D13')->getFont()->setBold(true);
$spreadsheet->getActiveSheet()->getStyle('E13')->getFont()->setBold(true);

// Set alignments
$helper->log('Set alignments');
$spreadsheet->getActiveSheet()->getStyle('D11')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
$spreadsheet->getActiveSheet()->getStyle('D12')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
$spreadsheet->getActiveSheet()->getStyle('D13')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

$spreadsheet->getActiveSheet()->getStyle('A18')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_JUSTIFY);
$spreadsheet->getActiveSheet()->getStyle('A18')->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

$spreadsheet->getActiveSheet()->getStyle('B5')->getAlignment()->setShrinkToFit(true);

// Set thin black border outline around column
$helper->log('Set thin black border outline around column');
$styleThinBlackBorderOutline = [
    'borders' => [
        'outline' => [
            'borderStyle' => Border::BORDER_THIN,
            'color' => ['argb' => 'FF000000'],
        ],
    ],
];
$spreadsheet->getActiveSheet()->getStyle('A4:E10')->applyFromArray($styleThinBlackBorderOutline);

// Set thick brown border outline around "Total"
$helper->log('Set thick brown border outline around Total');
$styleThickBrownBorderOutline = [
    'borders' => [
        'outline' => [
            'borderStyle' => Border::BORDER_THICK,
            'color' => ['argb' => 'FF993300'],
        ],
    ],
];
$spreadsheet->getActiveSheet()->getStyle('D13:E13')->applyFromArray($styleThickBrownBorderOutline);

// Set fills
$helper->log('Set fills');
$spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()->setFillType(Fill::FILL_SOLID);
$spreadsheet->getActiveSheet()->getStyle('A1:E1')->getFill()->getStartColor()->setARGB('FF808080');

// Set style for header row using alternative method
$helper->log('Set style for header row using alternative method');
$spreadsheet->getActiveSheet()->getStyle('A3:E3')->applyFromArray(
    [
            'font' => [
                'bold' => true,
            ],
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_RIGHT,
            ],
            'borders' => [
                'top' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
            'fill' => [
                'fillType' => Fill::FILL_GRADIENT_LINEAR,
                'rotation' => 90,
                'startColor' => [
                    'argb' => 'FFA0A0A0',
                ],
                'endColor' => [
                    'argb' => 'FFFFFFFF',
                ],
            ],
        ]
);

$spreadsheet->getActiveSheet()->getStyle('A3')->applyFromArray(
    [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
            ],
            'borders' => [
                'left' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
        ]
);

$spreadsheet->getActiveSheet()->getStyle('B3')->applyFromArray(
    [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
            ],
        ]
);

$spreadsheet->getActiveSheet()->getStyle('E3')->applyFromArray(
    [
            'borders' => [
                'right' => [
                    'borderStyle' => Border::BORDER_THIN,
                ],
            ],
        ]
);

// Unprotect a cell
$helper->log('Unprotect a cell');
$spreadsheet->getActiveSheet()->getStyle('B1')->getProtection()->setLocked(Protection::PROTECTION_UNPROTECTED);

// Add a hyperlink to the sheet
$helper->log('Add a hyperlink to an external website');
$spreadsheet->getActiveSheet()->setCellValue('E26', 'www.phpexcel.net');
$spreadsheet->getActiveSheet()->getCell('E26')->getHyperlink()->setUrl('https://www.example.com');
$spreadsheet->getActiveSheet()->getCell('E26')->getHyperlink()->setTooltip('Navigate to website');
$spreadsheet->getActiveSheet()->getStyle('E26')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

$helper->log('Add a hyperlink to another cell on a different worksheet within the workbook');
$spreadsheet->getActiveSheet()->setCellValue('E27', 'Terms and conditions');
$spreadsheet->getActiveSheet()->getCell('E27')->getHyperlink()->setUrl("sheet://'Terms and conditions'!A1");
$spreadsheet->getActiveSheet()->getCell('E27')->getHyperlink()->setTooltip('Review terms and conditions');
$spreadsheet->getActiveSheet()->getStyle('E27')->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);

// Add a drawing to the worksheet
$helper->log('Add a drawing to the worksheet');
$drawing = new Drawing();
$drawing->setName('Logo');
$drawing->setDescription('Logo');
$drawing->setPath(__DIR__ . '/../images/officelogo.jpg');
$drawing->setHeight(36);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

// Add a drawing to the worksheet
$helper->log('Add a drawing to the worksheet');
$drawing = new Drawing();
$drawing->setName('Paid');
$drawing->setDescription('Paid');
$drawing->setPath(__DIR__ . '/../images/paid.png');
$drawing->setCoordinates('B15');
$drawing->setOffsetX(110);
$drawing->setRotation(25);
$drawing->getShadow()->setVisible(true);
$drawing->getShadow()->setDirection(45);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

// Add a drawing to the worksheet
$helper->log('Add a drawing to the worksheet');
$drawing = new Drawing();
$drawing->setName('PhpSpreadsheet logo');
$drawing->setDescription('PhpSpreadsheet logo');
$drawing->setPath(__DIR__ . '/../images/PhpSpreadsheet_logo.png');
$drawing->setHeight(36);
$drawing->setCoordinates('D24');
$drawing->setOffsetX(10);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

// Play around with inserting and removing rows and columns
$helper->log('Play around with inserting and removing rows and columns');
$spreadsheet->getActiveSheet()->insertNewRowBefore(6, 10);
$spreadsheet->getActiveSheet()->removeRow(6, 10);
$spreadsheet->getActiveSheet()->insertNewColumnBefore('E', 5);
$spreadsheet->getActiveSheet()->removeColumn('E', 5);

// Set header and footer. When no different headers for odd/even are used, odd header is assumed.
$helper->log('Set header/footer');
$spreadsheet->getActiveSheet()->getHeaderFooter()->setOddHeader('&L&BInvoice&RPrinted on &D');
$spreadsheet->getActiveSheet()->getHeaderFooter()->setOddFooter('&L&B' . $spreadsheet->getProperties()->getTitle() . '&RPage &P of &N');

// Set page orientation and size
$helper->log('Set page orientation and size');
$spreadsheet->getActiveSheet()->getPageSetup()->setOrientation(PageSetup::ORIENTATION_PORTRAIT);
$spreadsheet->getActiveSheet()->getPageSetup()->setPaperSize(PageSetup::PAPERSIZE_A4);

// Rename first worksheet
$helper->log('Rename first worksheet');
$spreadsheet->getActiveSheet()->setTitle('Invoice');

// Create a new worksheet, after the default sheet
$helper->log('Create a second Worksheet object');
$spreadsheet->createSheet();

// Llorem ipsum...
$sLloremIpsum = 'Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Vivamus eget ante. Sed cursus nunc semper tortor. Aliquam luctus purus non elit. Fusce vel elit commodo sapien dignissim dignissim. Pellentesque habitant morbi tristique senectus et netus et malesuada fames ac turpis egestas. Curabitur accumsan magna sed massa. Nullam bibendum quam ac ipsum. Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Proin augue. Praesent malesuada justo sed orci. Pellentesque lacus ligula, sodales quis, ultricies a, ultricies vitae, elit. Sed luctus consectetuer dolor. Vivamus vel sem ut nisi sodales accumsan. Nunc et felis. Suspendisse semper viverra odio. Morbi at odio. Integer a orci a purus venenatis molestie. Nam mattis. Praesent rhoncus, nisi vel mattis auctor, neque nisi faucibus sem, non dapibus elit pede ac nisl. Cras turpis.';

// Add some data to the second sheet, resembling some different data types
$helper->log('Add some data');
$spreadsheet->setActiveSheetIndex(1);
$spreadsheet->getActiveSheet()->setCellValue('A1', 'Terms and conditions');
$spreadsheet->getActiveSheet()->setCellValue('A3', $sLloremIpsum);
$spreadsheet->getActiveSheet()->setCellValue('A4', $sLloremIpsum);
$spreadsheet->getActiveSheet()->setCellValue('A5', $sLloremIpsum);
$spreadsheet->getActiveSheet()->setCellValue('A6', $sLloremIpsum);

// Set the worksheet tab color
$helper->log('Set the worksheet tab color');
$spreadsheet->getActiveSheet()->getTabColor()->setARGB('FF0094FF');

// Set alignments
$helper->log('Set alignments');
$spreadsheet->getActiveSheet()->getStyle('A3:A6')->getAlignment()->setWrapText(true);

// Set column widths
$helper->log('Set column widths');
$spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(80);

// Set fonts
$helper->log('Set fonts');
$spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setName('Candara');
$spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setSize(20);
$spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setBold(true);
$spreadsheet->getActiveSheet()->getStyle('A1')->getFont()->setUnderline(Font::UNDERLINE_SINGLE);

$spreadsheet->getActiveSheet()->getStyle('A3:A6')->getFont()->setSize(8);

// Add a drawing to the worksheet
$helper->log('Add a drawing to the worksheet');
$drawing = new Drawing();
$drawing->setName('Terms and conditions');
$drawing->setDescription('Terms and conditions');
$drawing->setPath(__DIR__ . '/../images/termsconditions.jpg');
$drawing->setCoordinates('B14');
$drawing->setWorksheet($spreadsheet->getActiveSheet());

// Set page orientation and size
$helper->log('Set page orientation and size');
$spreadsheet->getActiveSheet()->getPageSetup()->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
$spreadsheet->getActiveSheet()->getPageSetup()->setPaperSize(PageSetup::PAPERSIZE_A4);

// Rename second worksheet
$helper->log('Rename second worksheet');
$spreadsheet->getActiveSheet()->setTitle('Terms and conditions');

// Set active sheet index to the first sheet, so Excel opens this as the first sheet
$spreadsheet->setActiveSheetIndex(0);

return $spreadsheet;
