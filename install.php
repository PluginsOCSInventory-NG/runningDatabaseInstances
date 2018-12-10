<?php

/**
 * This function is called on installation and is used to create database schema for the plugin
 */
function extension_install_runningDatabaseInstances()
{
    $commonObject = new ExtensionCommon;
    $commonObject->sqlQuery("CREATE TABLE IF NOT EXISTS `dbinstances` (
                          `ID` INT(11) NOT NULL AUTO_INCREMENT,
                          `HARDWARE_ID` INT(11) NOT NULL,
                          `PUBLISHER`   VARCHAR(255)     NULL DEFAULT NULL,
                          `NAME`        VARCHAR(255)     NULL DEFAULT NULL,
                          `VERSION`     VARCHAR(255)     NULL DEFAULT NULL,
                          `EDITION`     VARCHAR(255)     NULL DEFAULT NULL,
                          `INSTANCE`    VARCHAR(255)     NULL DEFAULT NULL,
                          PRIMARY KEY  (`ID`,`HARDWARE_ID`),
                          INDEX `NAME` (`NAME`),
                          INDEX `VERSION` (`VERSION`),
                          INDEX `ID` (`ID`)
                          ) COLLATE='utf8_general_ci' ENGINE=INNODB ROW_FORMAT=DEFAULT;");
}

/**
 * This function is called on removal and is used to destroy database schema for the plugin
 */
function extension_delete_runningDatabaseInstances()
{
    $commonObject = new ExtensionCommon;
    $commonObject->sqlQuery("DROP TABLE `dbinstances`");
}

/**
 * This function is called on plugin upgrade
 */
function extension_upgrade_runningDatabaseInstances()
{

}
