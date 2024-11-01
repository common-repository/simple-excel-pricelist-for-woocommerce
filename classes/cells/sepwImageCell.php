<?php

use PhpOffice\PhpSpreadsheet\Style;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use OnestExcelWriter\ImageCell;

class sepwImageCell extends ImageCell
{
    /**
     * @var string
     */
    private $images_size;

    public function __construct($slug)
    {
        parent::__construct($slug);
        $this->tmp = dirname(dirname(dirname(__FILE__))) . '/tmp';

        $options = get_option( 'sepw_settings' );
        $this->images_size = isset($options['images_size']) ? $options['images_size'] : 'thumbnail';
    }

    /**
     * @param int $id
     * @return string
     */
    protected function thumbPath($id)
    {
        $size = $this->images_size;
        $meta = wp_get_attachment_metadata( $id );
        $fullsize_path = WP_CONTENT_DIR . '/uploads/' . $meta['file'];
        $dir = dirname($fullsize_path);
        $thumb_path = $dir . '/' . $meta['sizes'][$size]['file'];
        return $thumb_path;
    }
}