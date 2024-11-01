<?php

use Twig\Environment;
use Twig\Loader\FilesystemLoader;
use Twig\RuntimeLoader\FactoryRuntimeLoader;

class sepwSettingsPage extends sepwBootstrap
{
    const TWIG_CACHE = false;
    const SEPW_FILENAME = 'simple-wc-excel-pricelist.php';
    const DEFAULT_COLS = array('thumbnail', 'SKU', 'name', 'price', 'stock');

    /**
     * @var array
     */
    private $fields;

    /**
     * @var Twig_Environment
     */
    private $twig;

    /**
     * Start up
     */
    public function __construct()
    {
        parent::__construct();

        $this->initTwig();

        add_action( 'admin_menu', array( $this, 'add_plugin_page' ) );
        add_action( 'admin_init', array( $this, 'page_init' ) );

        add_action('plugins_loaded', array($this, 'init_fields'), 11);
        add_action('admin_enqueue_scripts', array($this, 'admin_enqueue_scripts'));

        add_filter('plugin_action_links', array($this, 'action_links'), 10, 4 );
    }

    public function admin_enqueue_scripts()
    {
        wp_enqueue_style( 'sepw_admin_css', dirname(plugin_dir_url( __FILE__ )) . '/css/admin.css' );
        wp_enqueue_script( 'sepw_admin_js', dirname(plugin_dir_url( __FILE__ )) . '/js/admin.js' );
    }

    public function init_fields()
    {
        $this->fields = [
            'thumbnail' => __('Thumbnail', 'sepw'),
            'SKU' => __('SKU', 'sepw'),
            'name' => __('Name', 'sepw'),
            'price' => __('Price', 'sepw'),
            'stock' => __('Stock', 'sepw'),
        ];
    }

    /**
     * Add options page
     */
    public function add_plugin_page()
    {
        add_submenu_page(
            'woocommerce',
            'Simple Excel Pricelist for WooCommerce',
            'Simple Excel Pricelist',
            'manage_options',
            'sepw',
            array( $this, 'create_admin_page' )
        );
    }

    /**
     * Options page callback
     */
    public function create_admin_page()
    {
        $woocommerce = in_array( 'woocommerce/woocommerce.php', apply_filters( 'active_plugins', get_option( 'active_plugins' ) ) );
        ?>
        <div class="wrap sepw-wrap">
            <?php
            $this->view('header', [
                'locale' => get_locale(),
                'woocommerce' => $woocommerce,
                'assets' => $this->plugin_url . 'assets',
                ]);
            ?>
            <form method="post" action="options.php">
            <?php
                // This prints out all hidden setting fields
                settings_fields( 'sepw_option_group' );
                do_settings_sections( 'sepw-setting-admin' );
                submit_button();
            ?>
            </form>
        </div>
        <?php
    }

    /**
     * Register and add settings
     */
    public function page_init()
    {
        register_setting(
            'sepw_option_group', // Option group
            'sepw_settings', // Option name
            array( $this, 'sanitize' ) // Sanitize
        );

        add_settings_section(
            'info_section_id',
            __('Guide', 'sepw'),
            array( $this, 'print_info_section_info' ),
            'sepw-setting-admin'
        );

        add_settings_section(
            'setting_section_id', // ID
            __('Settings', 'sepw'), // Title
            array( $this, 'print_settings_section_info' ), // Callback
            'sepw-setting-admin' // Page
        );

        add_settings_section(
            'generate_section_id',
            __('Create Pricelist', 'sepw'),
            array( $this, 'print_generate_section_info' ),
            'sepw-setting-admin'
        );

        add_settings_field(
            'product_fields',
            __('Filters', 'sepw'),
            array( $this, 'product_fields_callback' ),
            'sepw-setting-admin',
            'setting_section_id'
        );

        add_settings_field(
            'images_size',
            __('Thumbnail size', 'sepw'),
            array( $this, 'images_size_callback' ),
            'sepw-setting-admin',
            'setting_section_id'
        );

        add_settings_field(
            'cache_lifetime',
            __('Update Frequency', 'sepw'),
            array( $this, 'cache_lifetime_callback' ),
            'sepw-setting-admin',
            'setting_section_id'
        );
    }

    /**
     * Sanitize each setting field as needed
     *
     * @param array $input Contains all settings fields as array keys
     */
    public function sanitize( $input )
    {
        $new_input = array();

        if( isset( $input['product_fields'] ) )
            $new_input['product_fields'] = $input['product_fields'];

        if( isset( $input['cache_lifetime'] ) )
            $new_input['cache_lifetime'] = absint( $input['cache_lifetime'] );

        if( isset( $input['images_size'] ) )
            $new_input['images_size'] = $input['images_size'];

        return $new_input;
    }

    /**
     * Print the Section text
     */
    public function print_settings_section_info()
    {}

    public function print_info_section_info()
    {
        $this->view('info', [
            'locale' => get_locale(),
        ]);
    }

    public function print_generate_section_info()
    {
        $this->view('generate', [
            'generated' => $this->generated,
            'link' => $this->plugin_url . 'out/pricelist.xlsx',
        ]);
    }

    /**
     * Get the settings option array and print one of its values
     */
    public function product_fields_callback()
    {
        $saved_fields = isset( $this->options['product_fields'] ) ? $this->options['product_fields'] : self::DEFAULT_COLS;
        $this->view('params/fields', [
            'saved_fields' => $saved_fields,
            'fields' => $this->fields,
        ]);
    }

    /**
     * Get the settings option array and print one of its values
     */
    public function cache_lifetime_callback()
    {
        $this->view('params/lifetime', [
            'dim' => __('minutes', 'sepw'),
            'lifetime' => isset( $this->options['cache_lifetime'] ) ? esc_attr($this->options['cache_lifetime']) : 30,
        ]);
    }

    public function images_size_callback()
    {
        $sizes = get_intermediate_image_sizes();
        $size =
            isset( $this->options['images_size'] )
            ? esc_attr($this->options['images_size'])
            : $sizes[0];
        $this->view('params/images-size', [
            'sizes' => $sizes,
            'size' => $size,
        ]);
    }

    /**
     * @param string $view
     * @param array $data
     */
    private function view($view, $data = array())
    {
        echo $this->twig->render($view . '.html.twig', $data);
    }

    private function initTwig()
    {
        $twig_options = array();

        if (self::TWIG_CACHE) {
            if ( ! file_exists($this->plugin_path . '/cache')) {
                mkdir($this->plugin_path . '/cache', 0755, true);
            }
            $twig_options['cache'] = $this->plugin_path . '/cache';
        }

        $loader = new Twig_Loader_Filesystem($this->plugin_path . '/views');
        $this->twig = new Twig_Environment($loader, $twig_options);

        $translate_func = new Twig_SimpleFunction('__', array($this, '__'));
        $this->twig->addFunction($translate_func);
    }

    public function __(string $s)
    {
        return __($s, 'sepw');
    }

    public function action_links($actions, $plugin_file)
    {
        if ( false === strpos( $plugin_file, self::SEPW_FILENAME ) ) {
            return $actions;
        }

        $settings_link = '<a href="admin.php?page=sepw">' . __('Settings') . '</a>';
        array_unshift( $actions, $settings_link );
	    return $actions;
    }
}
