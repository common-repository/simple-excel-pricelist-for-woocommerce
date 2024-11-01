<?php

use PhpOffice\PhpSpreadsheet\Style;
use OnestExcelWriter\TextCell;

class sepwVarStockCell extends TextCell
{
    /**
     * @param $sheet
     * @param int $col
     * @param int $row
     * @param array $data
     */
    public function write($sheet, $col, $row, $data)
    {
        $product = $data['product'];
        $variation = $data['variation'];

        if ($variation == NULL) return;

        $cell = $sheet->getCellByColumnAndRow($col, $row);
        $coord = $cell->getCoordinate();
        $cell->setValue(isset($variation['max_qty']) ? $variation['max_qty'] : '');
        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFEEEEEE');
        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }
}