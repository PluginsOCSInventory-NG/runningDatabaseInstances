$computerName = $env:computername

$edition = "Unknown"
$version = "Unknown"

$inst = (get-itemproperty 'HKLM:\SOFTWARE\Microsoft\Microsoft SQL Server').InstalledInstances
$xml = ""

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

   $xml += "<DBINSTANCES>`n"
   $xml += "<NAME>" + $serverName + "</NAME>`n"
   $xml += "<VERSION>" + $version + "</VERSION>`n"
   $xml += "<EDITION>" + $edition + "</EDITION>`n"
   $xml += "<INSTANCE>" + $instanceName + "</INSTANCE>`n"
   $xml += "</DBINSTANCES>`n"

   $xml += "<SOFTWARES>`n"
   $xml += "<PUBLISHER>Microsoft Corporation</PUBLISHER>`n"
   $xml += "<NAME>" + $serverName + "</NAME>`n"
   $xml += "<VERSION>" + $version + "</VERSION>`n"
   $xml += "<COMMENTS>DBInstances plugin</COMMENTS>`n"
   $xml += "</SOFTWARES>`n"

}

echo $xml