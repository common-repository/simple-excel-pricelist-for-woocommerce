<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Cell\DataType;

use OnestExcelWriter\Worker as ExcelWorker;

class sepwWorker extends sepwBootstrap
{

    const DEFAULT_COLS = array('thumbnail', 'SKU', 'name', 'price', 'stock');

    /**
     * @var array
     */
    private $handlers;

    /**
     * @var ExcelWorker
     */
    private $worker;

    /**
     * @var array
     */
    private $fields;

    public function __construct()
    {
        parent::__construct();

        $this->worker = new ExcelWorker(dirname(dirname( __FILE__ )));

        $this->handlers = array(
            new sepwHeadNameCell('head-name'),
            new sepwHeadCell('head-SKU'),
            new sepwHeadCell('head-thumbnail'),
            new sepwHeadCell('head-price'),
            new sepwHeadCell('head-stock'),

            new sepwBodyThumbnailCell('simple-thumbnail', $this->worker->getTmp()),
            new sepwBodyCell('simple-SKU'),
            new sepwBodyCell('simple-name'),
            new sepwBodyPriceCell('simple-price'),
            new sepwBodyStockCell('simple-stock'),

            new sepwVarNameCell('var-name'),
            new sepwVarSKUCell('var-SKU'),
            new sepwVarStockCell('var-stock'),
            new sepwVarThumbnailCell('var-thumbnail', $this->worker->getTmp()),
            new sepwVarPriceCell('var-price'),
        );

        add_action( 'rest_api_init', array($this, 'rest_api_init') );
        add_shortcode( 'pricelist', array($this, 'pricelist_shortcode') );

        $this->fields = isset($this->options['product_fields']) ? $this->options['product_fields'] : self::DEFAULT_COLS;
    }

    public function generate_callback()
    {
        $this->generate();

        return array(
            'status' => true,
            'time' => date('d.m.Y H:i:s'),
        );
    }

    private function generate()
    {
        $sheet = $this->worker->sheet();

        $products = wc_get_products(array(
            'status' => 'publish',
            'paginate' => false,
            'numberposts' => -1,
            'stock_status' => 'instock',
        ));

        $this->head($sheet);
        $this->body($sheet, $products);

        $this->worker->save('pricelist.xlsx');

        update_option('sepw_generated', time());
    }

    public function rest_api_init()
    {
        register_rest_route( 'sepw/v1', '/generate', array(
            'methods' => 'GET',
            'callback' => array($this, 'generate_callback'),
        ) );
    }

    private function head($sheet)
    {
        $col = 1;
        $row = 1;

        foreach ($this->fields as $c) {
            foreach ($this->handlers as $h) {
                if ($h->fits('head-'.$c)) {
                    $h->write($sheet, $col, $row, NULL);
                    break;
                }
            }
            $col++;
        }
    }

    private function body($sheet, $products)
    {
        $row = 2;

        foreach ($products as $p) {
            if ($p->is_type( 'variable' )) {

                $this->simple_row($sheet, $row, $p);
                $row++;

                $variations = $p->get_available_variations();
                foreach ($variations as $v) {
                    if ($v['variation_is_active'] && $v['variation_is_visible'] && $v['is_in_stock']) {
                        $this->variable_row($sheet, $row, $p, $v);
                        $row++;    
                    }
                }

            } else {
                $this->simple_row($sheet, $row, $p);
                $row++;
            }
        }
    }

    private function simple_row($sheet, $row, $product)
    {
        $col = 1;
        foreach ($this->fields as $c) {
            foreach ($this->handlers as $h) {
                if ($h->fits('simple-'.$c)) {
                    $h->write($sheet, $col, $row, $product);
                    break;
                }
            }
            $col++;
        }
    }

    private function variable_row($sheet, $row, $product, $variation)
    {
        $col = 1;
        foreach ($this->fields as $c) {
            foreach ($this->handlers as $h) {
                if ($h->fits('var-'.$c)) {
                    $h->write($sheet, $col, $row, [ 'product' => $product, 'variation' => $variation ]);
                    break;
                }
            }
            $col++;
        }
    }

    /**
     * @param array attrs
     * @return string
     */
    public function pricelist_shortcode($attrs) 
    {
        $title = isset($attrs['title']) ? $attrs['title'] : __('Download Pricelist');
        $classes = isset($attrs['class']) ? $attrs['class'] : '';
        return '<a href="' . $this->pricelist_url() . '" class="' . $classes . '">' . $title . '</a>';
    }

    private function pricelist_url()
    {
        if ( ! $this->valid() ) {
            $this->generate();
        }
        return $this->plugin_url . 'out/pricelist.xlsx';
    }
}