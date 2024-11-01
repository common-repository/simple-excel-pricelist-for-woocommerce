<?php

use PhpOffice\PhpSpreadsheet\Style;

class sepwBodyStockCell extends sepwBodyCell
{
    /**
     * @param $sheet
     * @param int $col
     * @param int $row
     * @param WC_Product $p
     */
    public function write($sheet, $col, $row, $p)
    {
        $cell = $sheet->getCellByColumnAndRow($col, $row);
        $coord = $cell->getCoordinate();
        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        if ($p->is_type( 'variable' )) {
            $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFEEEEEE');
        } else {
            $cell->setValue($p->get_stock_quantity());
        }
        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }
}