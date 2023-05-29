<?php
include('koneksi.php');
require 'vendor/autoload.php'; // Load library PhpSpreadsheet

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Create new Spreadsheet object
$spreadsheet = new Spreadsheet();

// Select active sheet
$spreadsheet->setActiveSheetIndex(0);
$sheet = $spreadsheet->getActiveSheet();

// Define table headers
$sheet->setCellValue('A1', 'Jenis Pendaftaran');
$sheet->setCellValue('B1', 'Tanggal Masuk Sekolah');
$sheet->setCellValue('C1', 'NIS');
$sheet->setCellValue('D1', 'Nomor Peserta Ujian');
$sheet->setCellValue('E1', 'Apakah Pernah PAUD');
$sheet->setCellValue('F1', 'Apakah Pernah TK');
$sheet->setCellValue('G1', 'No. Seri SKHUN Sebelumnya');
$sheet->setCellValue('H1', 'No. Seri Ijazah Sebelumnya');
$sheet->setCellValue('I1', 'Hobi');
$sheet->setCellValue('J1', 'Cita-cita');

// Fetch data from database and populate the table
include 'koneksi.php';
$registrasi = mysqli_query($conn, "SELECT * FROM registrasi");

$rowCount = 2;
foreach ($registrasi as $row) {
    $sheet->setCellValue('A' . $rowCount, $row['jenis_pendaftaran']);
    $sheet->setCellValue('B' . $rowCount, $row['tahun_ajaran']);
    $sheet->setCellValue('C' . $rowCount, $row['nis']);
    $sheet->setCellValue('D' . $rowCount, $row['apakah_pernah_paud']);
    $sheet->setCellValue('E' . $rowCount, $row['apakah_pernah_tk']);
    $sheet->setCellValue('F' . $rowCount, $row['noSKHUN']);
    $sheet->setCellValue('G' . $rowCount, $row['noIJAZA']);
    $sheet->setCellValue('H' . $rowCount, $row['hobi']);
    $sheet->setCellValue('I' . $rowCount, $row['hobi']);
    $rowCount++;
}

// Set auto column width for all columns
foreach (range('A', 'J') as $column) {
    $sheet->getColumnDimension($column)->setAutoSize(true);
}

// Create a new Excel file and save it
$writer = new Xlsx($spreadsheet);
$filename = 'data_peserta.xlsx';
$writer->save($filename);

// Download the Excel file
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment; filename="' . $filename . '"');
header('Cache-Control: max-age=0');
ob_end_clean();
$writer->save('php://output');
exit;
?>
