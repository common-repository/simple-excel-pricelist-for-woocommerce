<?php

use PhpOffice\PhpSpreadsheet\Style;

class sepwHeadNameCell extends sepwHeadCell
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
        $cell->setValue(__('Name', 'sepw'));
        $sheet->getStyle($coord)->getAlignment()->setHorizontal(Style\Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle($coord)->getFont()->setBold(true);
        $sheet->getColumnDimension($this->columnLetter($col))->setAutosize(true);
        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }

}