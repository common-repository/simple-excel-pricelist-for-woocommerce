<?php

use PhpOffice\PhpSpreadsheet\Style;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class sepwBodyThumbnailCell extends sepwImageCell
{
    /**
     * @param string
     */
    public function __construct($slug, $tmp)
    {
        parent::__construct($slug);
        $this->setTmp((string)$tmp);
    }

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

        $id = $p->get_image_id();
        $thumb_path = $this->thumbPath($id);

        parent::write($sheet, $col, $row, $thumb_path);

        if ($p->is_type( 'variable' )) {
            $sheet->getStyle($coord)->getFill()->setFillType(Style\Fill::FILL_SOLID)->getStartColor()->setARGB('FFEEEEEE');
        }

        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }
}