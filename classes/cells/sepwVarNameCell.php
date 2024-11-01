<?php

use PhpOffice\PhpSpreadsheet\Style;
use OnestExcelWriter\TextCell;

class sepwVarNameCell extends TextCell
{
    /**
     * @param $sheet
     * @param int $col
     * @param int $row
     * @param $data
     */
    public function write($sheet, $col, $row, $data)
    {
        $product = $data['product'];
        $variation = $data['variation'];

        if ($variation == NULL) return;

        $cell = $sheet->getCellByColumnAndRow($col, $row);
        $coord = $cell->getCoordinate();
        $cell->setValue($this->get_name($variation, $product));
        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_LEFT);
        $sheet->getStyle($coord)->getAlignment()->setVertical(Style\Alignment::VERTICAL_CENTER);
        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }

    /**
     * @param array $variation
     * @param WC_product $product
     * @return string
     */
    private function get_name($variation, $product)
    {
        $attributes = $variation['attributes'];

        $data = array();
        foreach ($attributes as $taxonomy => $value) {
            $slug = str_replace('attribute_', '', $taxonomy);
            $label = wc_attribute_label( $slug, $product );
            $data[] = $label . ' ' . urldecode($value);
        }

        return '     ↳   ' . implode(', ', $data);
    }

}