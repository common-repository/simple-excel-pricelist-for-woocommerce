<?php

use PhpOffice\PhpSpreadsheet\Style;
use OnestExcelWriter\TextCell;

class sepwHeadCell extends TextCell
{
    /**
     * @param $sheet
     * @param int $col
     * @param int $row
     * @param $data
     */
    public function write($sheet, $col, $row, $data)
    {
        $cell = $sheet->getCellByColumnAndRow($col, $row);
        $coord = $cell->getCoordinate();
        $cell->setValue($this->value($this->slug));
        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($coord)->getFont()->setBold(true);
        $sheet->getColumnDimension($this->columnLetter($col))->setAutosize(true);
        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }

    /**
     * @param string $slug
     * @return string
     */
    private function value($slug)
    {
        switch ($slug) {
            case 'head-thumbnail':
            return __('Thumbnail', 'sepw');
            case 'head-SKU':
            return __('SKU', 'sepw');
            case 'head-name':
            return __('Name', 'sepw');
            case 'head-price':
            return __('Price', 'sepw');
            case 'head-stock':
            return __('Stock', 'sepw');
        }
    }

    /**
     * @param int
     * @return string
     */
    function columnLetter($c)
    {
        $c = intval($c);
        if ($c <= 0) return '';
    
        $letter = '';
                 
        while($c != 0){
           $p = ($c - 1) % 26;
           $c = intval(($c - $p) / 26);
           $letter = chr(65 + $p) . $letter;
        }
        
        return $letter;
    }

}