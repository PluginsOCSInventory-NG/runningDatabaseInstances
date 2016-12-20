<?php
function plugin_version_runningDatabaseInstances()
{
return array('name' => 'Running Database Instances',
'version' => '1.0',
'author'=> 'Community, Frank BOURDEAU',
'license' => 'GPLv2',
'verMinOcs' => '2.2');
}

function plugin_init_dbinstances()
{
$object = new plugins;
$object->add_cd_entry("dbinstances","other");

$object->sql_query("CREATE TABLE IF NOT EXISTS `dbinstances` (
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

function plugin_delete_dbinstances()
{
$object = new plugins;
$object->del_cd_entry("dbinstances");
$object->sql_query("DROP TABLE `dbinstances`");

}
