<?php

namespace Deviser\Spreadshelper\PhpSpreadsheet;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Sheet {

    public $spreadsheet;
    public $sheet;
    public $data;

    public function __construct() {
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    public function setCell() {
        $row = 1;
        foreach ($this->data as $key => $valData) {
            $col = 'A';
            if ($key == 0) {
                $headers = array_keys($valData);
                foreach ($headers as $head) {
                    $this->sheet->setCellValue($col . $row, $head);
                    $col++;
                }
                $row++;
                $col = 'A';
            }
            foreach ($valData as $content) {
                if (is_array($content)) {
                    $this->sheet->setCellValue($col . $row, $content['val']);
                } else {
                    $this->sheet->setCellValue($col . $row, $content);
                }
                $col++;
            }
            $row++;
        }
    }

    public function exportToExcel($data, $path = null) {
        $this->data = $data;
        $this->setCell();

        $writer = new Xlsx($this->spreadsheet);
        $path = $path ? rtrim($path, '/') : null;
        $writer->save($path . '/' . uniqid() . '.xlsx');
    }

    public function createExcel($data) {
        $headers = ['h1', 'h2', 'h3', 'h4'];
    }
}