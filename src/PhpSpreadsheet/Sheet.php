<?php

namespace Deviser\Spreadshelper\PhpSpreadsheet;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Sheet {

    private Spreadsheet $spreadsheet;
    private \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet;
    public array $data;
    public string $name;

    public function __construct() {
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    protected function setCell(): ?string {
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
                    $this->sheet->setCellValue($col . $row, $content['context']);
                    $content['style'] ? $this->setStyle($col . $row, $content['style']) : null;
                } else {
                    $this->sheet->setCellValue($col . $row, $content);
                }
                $col++;
            }
            $row++;
        }
        return null;
    }

    /**
     * @param string $coordinate
     * @param array $style
     * @return object
     */
    protected function setStyle(string $coordinate, array $style): object {
        $this->spreadsheet->getActiveSheet()->getStyle($coordinate)->applyFromArray($style);
        return $this;
    }

    /**
     * @param array $data This should be an multidimensional array
     * @return $this
     */
    public function setArray(array $data): object {
        $this->data = $data;
        $this->setCell();
        return $this;
    }

    /**
     * @param string|null $name This should be without ext
     * @return $this
     */
    public function setName(string $name = null): object {
        $this->name = $name ?? uniqid();
        return $this;
    }

    /**
     * @param $path
     * @return string|null
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     * @example ../examples/export-excel.php
     */
    public function excel($path = null): ?string {
        $writer = new Xlsx($this->spreadsheet);
        $path = $path ? rtrim($path, '/') : null;
        $writer->save($path . '/' . uniqid() . '.xlsx');
        return null;
    }


}