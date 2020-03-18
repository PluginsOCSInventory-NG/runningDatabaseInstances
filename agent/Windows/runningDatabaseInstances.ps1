﻿$edition = "Unknown"
$version = "Unknown"

$inst = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances

$xml += "<SOFTWARES>"
$xml += "<PUBLISHER>OCS Inventory Team</PUBLISHER>"
$xml += "<NAME>DBInstances</NAME>"
$xml += "<VERSION>2.0</VERSION>"
$xml += "<COMMENTS>DBInstances plugin</COMMENTS>"
$xml += "</SOFTWARES>"

foreach ($i in $inst)
{
   $instanceFullName = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL').$i
   $instanceName = $instanceFullName

   if($instanceFullName.Contains(".")){
       $instanceSplitted = $instanceFullName.Split(".")
       $instanceName = $instanceSplitted.Get(1)
   }

   $publisher = "Microsoft Corporation"
   $edition = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$instanceFullName\Setup").Edition
   $version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server\$instanceFullName\Setup").Version

   if($version.StartsWith("10")){
       $serverName = "Microsoft SQL Server 2008"
   }elseif($version.StartsWith("10.5")){
       $serverName = "Microsoft SQL Server 2008 R2"
   }elseif($version.StartsWith("11")){
       $serverName = "Microsoft SQL Server 2012"
   }elseif($version.StartsWith("12")){
       $serverName = "Microsoft SQL Server 2014"
   }elseif($version.StartsWith("13")){
       $serverName = "Microsoft SQL Server 2016"
   }elseif($version.StartsWith("14")){
       $serverName = "Microsoft SQL Server 2017"
   }elseif($version.StartsWith("15")){
       $serverName = "Microsoft SQL Server 2019"
   }else{
       $serverName = "Microsoft SQL Server " + $version
   }

   $xml += "<DBINSTANCES>"
   $xml += "<PUBLISHER>Microsoft Corporation</PUBLISHER>"
   $xml += "<NAME>" + $serverName + "</NAME>"
   $xml += "<VERSION>" + $version + "</VERSION>"
   $xml += "<EDITION>" + $edition + "</EDITION>"
   $xml += "<INSTANCE>" + $instanceName + "</INSTANCE>"
   $xml += "</DBINSTANCES>"

   $xml += "<SOFTWARES>"
   $xml += "<PUBLISHER>Microsoft Corporation</PUBLISHER>"
   $xml += "<NAME>" + $serverName + "</NAME>"
   $xml += "<VERSION>" + $version + "</VERSION>"
   $xml += "<COMMENTS>DBInstances plugin</COMMENTS>"
   $xml += "</SOFTWARES>"

}

echo $xml
