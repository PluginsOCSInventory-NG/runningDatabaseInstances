<?php
function plugin_version_databaseinstances()
{
return array('name' => 'Running database instances',
'version' => '1.3',
'author'=> 'Community, Frank BOURDEAU',
'license' => 'GPLv2',
'verMinOcs' => '2.2');
}

function plugin_init_databaseinstances()
{
$object = new plugins;
$object->add_cd_entry("databaseinstances","other");

// databaseinstances table creation

$object -> sql_query("CREATE TABLE IF NOT EXISTS `databaseinstances` (
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

function plugin_delete_databaseinstances()
{
$object = new plugins;
$object->del_cd_entry("databaseinstances");

$object->sql_query("DROP TABLE `databaseinstances`;");

}

?>