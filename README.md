# SharePoint 

- Convenience class to work with SharePoint


###### Example
```php
<?php

require_once "app/Sharepoint.php";

/**TODO
 * make sure to configure Settings.php with your own login credentials
 * @param $settings["Url"] url to the SharePoint site you are working with
 * @param $settings["UserName"]
 * @param $settings["Password"]
 *
 */
$settings = include("vendor/vgrem/php-spo/Settings.php");

$sharePoint =  Sharepoint::getInstance($settings);
$sharePoint->uploadFile(
    // file to be uploaded
    __FILE__,
    // directory where to upload
    // "" uploads to base dir that is usually 'shared documents' for the site being used
    "",
    // create directory if the directory does not exist
    false
);




?>
````