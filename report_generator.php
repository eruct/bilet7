<?php
require_once 'vendor/autoload.php';

use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;
use PhpOffice\PhpWord\Shared\Converter;

function generateReport() {
    $phpWord = new PhpWord();
    $section = $phpWord->addSection();
    
    $section->addText(
        '123',
        ['name' => 'Arial', 'size' => 16, 'bold' => true],
        ['alignment' => 'center']
    );
    
    $section->addTextBreak(1);
    $section->addText(
        '123',
        ['name' => 'Arial', 'size' => 12],
        ['alignment' => 'left']
    );
    
    $table = $section->addTable(['borderSize' => 6]);
    $table->addRow();
    $table->addCell(2000)->addText('Товар', ['bold' => true]);
    $table->addCell(2000)->addText('Количество', ['bold' => true]);
    $table->addCell(2000)->addText('Сумма', ['bold' => true]);
    
    $sales = [
        ['Молоко', 120, 5400],
        ['Хлеб', 95, 2850],
        ['Сыр', 45, 11250]
    ];
    
    foreach ($sales as $sale) {
        $table->addRow();
        $table->addCell()->addText($sale[0]);
        $table->addCell()->addText($sale[1]);
        $table->addCell()->addText(number_format($sale[2], 2) . ' руб.');
    }
    
    $section->addTextBreak(1);
    $section->addText(
        'Дата формирования: ' . date('d.m.Y H:i'),
        ['italic' => true],
        ['alignment' => 'right']
    );
    
    $filename = 'report_' . date('Ymd_His') . '.docx';
    $objWriter = IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save($filename);
    
    return $filename;
}

if (isset($_GET['generate'])) {
    $file = generateReport();
    
    header('Content-Description: File Transfer');
    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment; filename="' . basename($file) . '"');
    header('Cache-Control: must-revalidate');
    header('Content-Length: ' . filesize($file));
    readfile($file);
    unlink($file); 
    exit;
}
?>