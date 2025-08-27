<?php
require 'vendor/autoload.php'; // Asegúrate de tener PhpSpreadsheet instalado

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$archivo = __DIR__ . '/excel/registros.xlsx';

// Recibir datos en JSON
$data = json_decode(file_get_contents("php://input"), true);

if (!$data) {
    echo json_encode(["success" => false, "msg" => "No se recibieron datos"]);
    exit;
}

// Abrir o crear Excel
if (file_exists($archivo)) {
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($archivo);
    $sheet = $spreadsheet->getActiveSheet();
} else {
    $spreadsheet = new Spreadsheet();
    $sheet = $spreadsheet->getActiveSheet();
    // Cabeceras
    $col = 'A';
    foreach (array_keys($data) as $header) {
        $sheet->setCellValue($col.'1', $header);
        $col++;
    }
}

// Agregar registro en la siguiente fila
$fila = $sheet->getHighestRow() + 1;
$col = 'A';
foreach ($data as $value) {
    $sheet->setCellValue($col.$fila, $value);
    $col++;
}

// Guardar archivo
$writer = new Xlsx($spreadsheet);
$writer->save($archivo);

echo json_encode(["success" => true, "msg" => "Registro agregado correctamente"]);
?>

