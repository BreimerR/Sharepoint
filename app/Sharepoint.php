<?php

/**
 * @author Breimer Radido
 * @email brymher@gmail.com
 * @project Sharepoint
 *
 * @description A SharePoint convenience class
 *
 * @alert Class has not been fully tested for functions
 */


use Office365\PHP\Client\Runtime\Utilities\RequestOptions;
use \Office365\PHP\Client\SharePoint\ClientContext;
use \Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use \Office365\PHP\Client\SharePoint\FileCreationInformation;


$spo_settings = include("vendor/vgrem/php-spo/Settings.php");

class Sharepoint
{

    private $_ctx = null,
        $_error = null,
        $_errors = [],
        $_root_dir = "shared documents";

    private static $INSTANCE = null;

    private static $SETTINGS = null;

    private function __construct()
    {
        try {
            $settings = self::$SETTINGS;

            $context = new AuthenticationContext($settings["Url"]);

            $context->acquireTokenForUser($settings["UserName"], $settings["Password"]);

            $this->_ctx = new ClientContext($settings["Url"], $context);

        } catch (Exception $e) {
            $this->_error = "failed to connect" . $e->getMessage();

            echo $this->_error;
        }
    }

    public static function getInstance($settings)
    {
        self::$SETTINGS = $settings;

        return self::$INSTANCE == null ? self::$INSTANCE = new self() : self::$INSTANCE;
    }

    public function hasError()
    {
        return $this->_error != null;
    }

    public function hasAccess()
    {
        return $this->_ctx != null;
    }

    function uploadFile($localPath, $uploadToDir = "", $createFolder = false)
    {
        if ($this->hasAccess()) {
            // upload file to Documents directory by default
            $uploadToUrl = $this->_root_dir . ($uploadToDir == null || $uploadToDir == "" ? "" : "/$uploadToDir");

            try {
                $fileName = basename($localPath);
                $fileCreationInformation = new FileCreationInformation();
                $fileCreationInformation->Content = file_get_contents($localPath);
                $fileCreationInformation->Url = $fileName;

                $this->_ctx->getWeb()
                    // get folder
                    ->getFolderByServerRelativeUrl($uploadToUrl)
                    // get files object and upload a new file
                    ->getFiles()->add($fileCreationInformation);

                $this->_ctx->executeQuery();

                //$uploadFile->getListItemAllFields()->setProperty('Title', $fileName);
                //$uploadFile->getListItemAllFields()->update();
                //$ctx->executeQuery();

                echo "<br/> Uploaded.";

                return true;

            } catch (Exception $e) {
                echo "<br> failed";
                switch ($message = $e->getMessage()) {
                    case "File Not Found.":
                        if ($createFolder) {
                            $this->createFolder($uploadToUrl);
                            // avoid creation again
                            $this->uploadFile($localPath, $uploadToDir, false);
                        } else echo $message;
                        break;
                    default:
                        echo "<br/>$message";
                }
            }
        }

        return false;
    }

    function createFolder($folder)
    {
        try {
            $folders = explode("/", $folder);
            $folder = array_pop($folders);
            $parentFolder = implode("/", $folders);

            $files = $this->_ctx->getWeb()->getFolderByServerRelativeUrl($parentFolder)->getFiles();
            $this->_ctx->load($files);
            $this->_ctx->executeQuery();
            //print files info
            /* @var $file File */

            /** @documentation code
             * foreach ($files->getData() as $file) {
             *      print "File name: '{$file->getProperty("ServerRelativeUrl")}'\r\n";
             * }
             */

            $parentFolder = $this->_ctx->getWeb()->getFolderByServerRelativeUrl($parentFolder);
            $childFolder = $parentFolder->getFolders()->add($folder);
            $this->_ctx->executeQuery();

            return true;
        } catch (Exception $e) {
            $this->_error = "Folder creation failed.<br>{$e->getMessage()}";
            echo $this->_error;
        }

        return false;
    }

    function renameFolder($webUrl, $authCtx, $folderUrl, $folderNewName)
    {
        try {
            $url = $webUrl . "/_api/web/getFolderByServerRelativeUrl('{$folderUrl}')/ListItemAllFields";
            $request = new RequestOptions($url);
            $resp = $this->_ctx->executeQueryDirect($request);
            $data = json_decode($resp);

            $itemPayload = array(
                '__metadata' => array('type' => $data->d->__metadata->type),
                'Title' => $folderNewName,
                'FileLeafRef' => $folderNewName
            );

            $itemUrl = $data->d->__metadata->uri;
            $request = new RequestOptions($itemUrl);
            $request->addCustomHeader("X-HTTP-Method", "MERGE");
            $request->addCustomHeader("If-Match", "*");
            $request->Data = $itemPayload;
            $this->_ctx->executeQueryDirect($request);

            return true;
        } catch (Exception $e) {
            $this->_error = $e->getMessage();
        }

        return false;
    }


    function getWeb()
    {
        return $this->_ctx->getWeb();
    }
}