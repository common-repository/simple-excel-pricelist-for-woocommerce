<?php

use PhpOffice\PhpSpreadsheet\Style;
use OnestExcelWriter\TextCell;

class sepwVarPriceCell extends TextCell
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

        $price = html_entity_decode(strip_tags(isset($variation['price_html']) ? $variation['price_html'] : ''));

        $cell = $sheet->getCellByColumnAndRow($col, $row);
        $coord = $cell->getCoordinate();
        $cell->setValue($price);
        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }

}