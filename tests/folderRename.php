<?php

/**
 * @author Breimer Radido John
 * @email brymher@gmail.com
 * @project Sharepoint
 */

$settings = include("../vendor/vgrem/php-spo/Settings.php");

$sharePoint = Sharepoint::getInstance($settings);


if ($sharePoint->renameFolder("")) {
    echo "Folder renamed";
} else {
    echo "Folder rename failed {$sharePoint->_error}";
}