<?php

/**
 * @author Breimer Radido John
 * @email brymher@gmail.com
 * @project Sharepoint
 */


require_once "../app/Sharepoint.php";
// get settings
$settings = include("vendor/vgrem/php-spo/Settings.php");

$sharePoint = Sharepoint::getInstance($settings);