<?php

use PhpOffice\PhpSpreadsheet\Style;
use OnestExcelWriter\TextCell;

class sepwBodyCell extends TextCell
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
        $cell->setValue($this->value($p));
        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        if ($p->is_type( 'variable' )) {
            $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFEEEEEE');
        }

        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }

    /**
     * @param WC_Product $p
     * @return string
     */
    private function value($p)
    {
        switch ($this->slug) {
            case 'simple-thumbnail':
            case 'var-thumbnail':
            return '';
            case 'simple-SKU':
            case 'var-SKU':
            return $p->get_sku();
            case 'simple-name':
            case 'var-name':
            return $p->get_name();
            case 'simple-price':
            case 'var-price':
            return $p->get_price();
            case 'simple-stock':
            case 'var-stock':
            return $p->get_stock_quantity();
        }
    }
}