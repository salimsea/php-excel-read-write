<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx; 

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
$spreadsheet = $reader->load("test.xlsx");
$d=$spreadsheet->getSheet(0)->toArray();

echo count($d);
$sheet = $spreadsheet->getSheet(0);

$datas = [
    array(
        "no" => "1",
        "kecamatan" => "sukaraja",
        "pr_laut_prod" => "-",
        "pr_laut_n_prod" => "-",
        "pr_darat_prod" => "-",
        "pr_darat_n_prod" => "67.932,50",
        "tot_prod" => "67.932",
        "tot_n_prod" => "2",
        "tahun" => "2020"
    ),
    array(
        "no" => "2",
        "kecamatan" => "tanah sareal",
        "pr_laut_prod" => "-",
        "pr_laut_n_prod" => "-",
        "pr_darat_prod" => "-",
        "pr_darat_n_prod" => "58.321,50",
        "tot_prod" => "58.321",
        "tot_n_prod" => "2",
        "tahun" => "2020"
    ),
];
$i = 2;
foreach ($datas as $key => $value) {
    $i++;
    // $nosheet = "A".$key+1;
    $sheet->setCellValue('A'.$i, $value["no"]);
    $sheet->setCellValue('B'.$i, $value["kecamatan"]);
    $sheet->setCellValue('C'.$i, $value["pr_laut_prod"]);
    $sheet->setCellValue('D'.$i, $value["pr_laut_n_prod"]);
    $sheet->setCellValue('E'.$i, $value["pr_darat_prod"]);
    $sheet->setCellValue('F'.$i, $value["pr_darat_n_prod"]);
    $sheet->setCellValue('G'.$i, $value["tot_prod"]);
    $sheet->setCellValue('H'.$i, $value["tot_n_prod"]);
    $sheet->setCellValue('I'.$i, $value["tahun"]);
}

// Write an .xlsx file  
$writer = new Xlsx($spreadsheet); 
  
// Save .xlsx file to the files directory 
$writer->save('demo.xlsx'); 
?>