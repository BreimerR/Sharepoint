<?php

/**
 * @author Breimer Radido
 * @email brymher@gmail.com
 * @project SharePoint
 *
 * @description A SharePoint convenience class
 *
 * @alert Class has not been fully tested for functions
 */

use Office365\PHP\Client\SharePoint\File;
use \Office365\PHP\Client\SharePoint\ClientContext;
use Office365\PHP\Client\Runtime\Utilities\RequestOptions;
use \Office365\PHP\Client\Runtime\Auth\AuthenticationContext;
use \Office365\PHP\Client\SharePoint\FileCreationInformation;


$spo_settings = include("vendor/vgrem/php-spo/Settings.php");

class Sharepoint
{

    /**
     * @type ClientContext
     * */
    private $_ctx = null;

    private $_error = null,
        $_errors = [],
        $_site = null,
        $_root_dir = "shared documents";

    private static $INSTANCE = null,
        $URL,
        $USERNAME,
        $PASSWORD,
        $SETTINGS = null;

    private function __construct($site)
    {
        $this->loginToSite($site);
    }

    public function getSiteUrl($siteName = "")
    {
        return self::$URL . ($siteName == null || $siteName == "" ? "" : "/sites/$siteName");
    }

    function createInstance($siteUrl, $userName, $password)
    {
        try {
            $context = new AuthenticationContext($siteUrl);

            $context->acquireTokenForUser($userName, $password);

            $this->_ctx = new ClientContext($siteUrl, $context);

        } catch (Exception $e) {
            $this->_error = "failed to connect" . $e->getMessage();

            echo $this->_error;
        }
    }

    private function loginToSite($siteName = "")
    {
        $this->_site = $siteName;

        $this->createInstance($this->getSiteUrl($siteName), self::$USERNAME, self::$PASSWORD);
    }

    public static function getInstance($site = "")
    {
        $settings = $GLOBALS["spo_settings"];

        self::$USERNAME = $settings["UserName"];
        self::$PASSWORD = $settings["Password"];
        self::$URL = $settings["Url"];

        return self::$INSTANCE == null ? self::$INSTANCE = new self($site) : self::$INSTANCE;
    }

    public function hasError()
    {
        return $this->_error != null;
    }

    public function hasAccess()
    {
        return $this->_ctx != null;
    }

    /**
     * @param $localPath
     * @param string $uploadToDir
     * @param null $site
     * @param bool $createFolder
     * @return bool
     * @throws Exception
     */
    function uploadFile($localPath, $uploadToDir = "", $site = null, $createFolder = false)
    {
        // check if site is equal to old site else create a new context
        if ($site == null) {
            if ($this->_site == null) {
                throw new Exception("No site set to upload data to.");
            }
        }

        // login to a new site if site names do not match
        $site == $this->_site ?: $this->loginToSite($site);

        if ($this->hasAccess()) {
            // upload file to Documents directory by default
            $uploadToUrl = $this->_root_dir . ($uploadToDir == null || $uploadToDir == "" ? "" : "/$uploadToDir");

            try {
                return $this->fileUploader($localPath, $uploadToUrl);
            } catch (Exception $e) {
                switch ($message = $e->getMessage()) {
                    case "File Not Found.":
                        if ($createFolder) {
                            $this->createFolder($uploadToUrl);

                            return $this->fileUploader($localPath, $uploadToUrl);
                        } else throw new Exception("File not created, Folder " . $uploadToDir . " does not exist. <br/>Set createFile true to enable folder creation");
                        break;
                    default:
                        echo "<br/>$message";
                }
            }
        }

        return false;
    }

    private function fileUploader($localPath, $uploadToUrl)
    {
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

        return true;

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

    function renameFolder($site, $folderName, $folderNewName)
    {
        $folderNameUrl = $this->getFullPath($folderName);

        try {
            $url = $this->getSiteUrl($site) . "/_api/web/getFolderByServerRelativeUrl('/{$folderNameUrl}')";
            $request = new RequestOptions($url);
            $resp = $this->_ctx->executeQueryDirect($request);
            $data = json_decode($resp);

            var_dump($resp);
            echo "<br/>";


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
            echo $this->_error;
        }

        return false;
    }

    function getFullPath($folderName = "")
    {
        return $this->_root_dir . ($folderName == null || $folderName == "" ? "" : "/$folderName");
    }


    function getWeb()
    {
        return $this->_ctx->getWeb();
    }
}
