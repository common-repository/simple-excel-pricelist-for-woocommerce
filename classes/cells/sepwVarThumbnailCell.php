<?php

use PhpOffice\PhpSpreadsheet\Style;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;

class sepwVarThumbnailCell extends sepwImageCell
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
     * @param array $data
     */
    public function write($sheet, $col, $row, $data)
    {
        $product = $data['product'];
        $variation = $data['variation'];

        $cell = $sheet->getCellByColumnAndRow($col, $row);
        $coord = $cell->getCoordinate();

        if ($variation == NULL) return;

        if (! isset($variation['image_id'])) return;

        $id = $variation['image_id'];
        $thumb_path = $this->thumbPath($id);

        parent::write($sheet, $col, $row, $thumb_path);

        $sheet->getStyle($coord)->getBorders()->getBottom()->setBorderStyle(Style\Border::BORDER_THIN);
        $sheet->getStyle($coord)->getBorders()->getRight()->setBorderStyle(Style\Border::BORDER_THIN);
    }
}