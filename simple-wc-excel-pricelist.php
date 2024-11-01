<?php
/**
 * @package sepw
 * @version 1.13
 *
 *
 * Plugin Name: Simple Excel Pricelist for WooCommerce
 * Plugin URI: http://wordpress.org/plugins/simple-excel-pricelist-for-woocommerce/
 * Description: This plugin helps to create a price list of all products available in stock in excel format and allows users to download the file.
 * Text Domain: sepw
 * License: GPLv2 or later
 * Author: Sasha Prawas
 * Version: 1.13
 */

/** Load composer */
$composer = dirname(__FILE__) . '/vendor/autoload.php';
if ( file_exists($composer) ) {
    require_once $composer;
}

if ( is_admin() ) {
    new sepwSettingsPage();
} else {
    new sepwWorker();
}
