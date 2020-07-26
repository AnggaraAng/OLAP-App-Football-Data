<?php
require '../vendor/autoload.php';

// koneksi php dan mysql
$koneksi = mysqli_connect("localhost","root","","footballdb");

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// sheet peratama
$sheet->setTitle('Laporan Mahasiswa');
$sheet->setCellValue('A1', 'ID');
$sheet->setCellValue('B1', 'Date');
$sheet->setCellValue('C1', 'Home Team');
$sheet->setCellValue('D1', 'Away Team');
$sheet->setCellValue('E1', 'Home Score');
$sheet->setCellValue('F1', 'Away Score');
$sheet->setCellValue('G1', 'Tournament');
$sheet->setCellValue('H1', 'City');
$sheet->setCellValue('I1', 'Country');
$sheet->setCellValue('J1', 'Neutral');

// membaca data dari mysql
$tfootball = mysqli_query($koneksi,"select * from tfootball");
$row = 2;
while($record = mysqli_fetch_array($tfootball))
{
    $sheet->setCellValue('A'.$row, $record['id']);
    $sheet->setCellValue('B'.$row, $record['date']);
    $sheet->setCellValue('C'.$row, $record['home_team']);
    $sheet->setCellValue('D'.$row, $record['away_team']);
    $sheet->setCellValue('E'.$row, $record['home_score']);
    $sheet->setCellValue('F'.$row, $record['away_score']);
    $sheet->setCellValue('G'.$row, $record['tournament']);
    $sheet->setCellValue('H'.$row, $record['city']);
    $sheet->setCellValue('I'.$row, $record['country']);
    $sheet->setCellValue('J'.$row, $record['neutral']);
    $row++;
}

$writer = new Xlsx($spreadsheet);
$writer->save('Laporan All Data.xlsx');
header('Location: ../index.php');
?>