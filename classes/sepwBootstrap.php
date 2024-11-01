<?php

class sepwBootstrap
{
    /**
     * @var array
     */
    protected $options;

    /**
     * @var string
     */
    protected $plugin_path;

    /**
     * @var string
     */
    protected $plugin_url;

    /**
     * @var int
     */
    protected $generated;

    public function __construct() {

        $this->options = get_option( 'sepw_settings' );
        $this->generated = get_option( 'sepw_generated', 0 );
        $this->plugin_path = dirname(dirname( __FILE__ ));
        $this->plugin_url = plugin_dir_url(dirname(dirname( __FILE__ )) . '/xlsx-pricelist.php');        

        add_action('plugins_loaded', array($this, 'load_domain'), 10);
    }

    public function load_domain() {
        $plugin_dir = basename(dirname(dirname(__FILE__)));
        load_plugin_textdomain( 'sepw', false, $plugin_dir . '/languages' );
    }

    /**
     * @return bool
     */
    protected function valid() {
        $lifetime_min = isset($this->options['cache_lifetime']) ? $this->options['cache_lifetime'] : 30;
        $lifetime_sec = $lifetime_min * 60;
        return $this->generated + $lifetime_sec > time();
    }
}