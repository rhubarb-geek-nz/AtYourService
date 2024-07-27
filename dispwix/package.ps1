# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

param(
	$CertificateThumbprint = '601A8B683F791E51F647D34AD102C38DA4DDB65F',
	$Architectures = @('arm', 'arm64', 'x86', 'x64')
)

$ProgressPreference = 'SilentlyContinue'
$ErrorActionPreference = 'Stop'

trap
{
	throw $PSItem
}

foreach ($ARCH in $Architectures)
{
	$VERSION=(Get-Item "..\bin\$ARCH\RhubarbGeekNzAtYourService.exe").VersionInfo.ProductVersion

	$env:PRODUCTVERSION = $VERSION
	$env:PRODUCTARCH = $ARCH
	$env:PRODUCTWIN64 = 'yes'
	$env:PRODUCTPROGFILES = 'ProgramFiles64Folder'
	$env:INSTALLERVERSION = '500'

	switch ($ARCH)
	{
		'arm64' {
			$env:UPGRADECODE = 'CDEF0F56-7B7A-4CEC-9644-E46384676A89'
		}

		'arm' {
			$env:UPGRADECODE = 'E38BE753-5466-499E-8533-8CEEC893873C'
			$env:PRODUCTWIN64 = 'no'
			$env:PRODUCTPROGFILES = 'ProgramFilesFolder'
		}

		'x86' {
			$env:UPGRADECODE = '4CFE8067-16C2-482E-8283-BE3EDF44C8E7'
			$env:PRODUCTWIN64 = 'no'
			$env:PRODUCTPROGFILES = 'ProgramFilesFolder'
			$env:INSTALLERVERSION = '200'
		}

		'x64' {
			$env:UPGRADECODE = '005031A4-F509-41E1-80B2-A8E532E41E36'
			$env:INSTALLERVERSION = '200'
		}
	}	

	& "${env:WIX}bin\candle.exe" -nologo "dispsvc.wxs"

	if ($LastExitCode -ne 0)
	{
		exit $LastExitCode
	}

	$MsiFilename = "rhubarb-geek-nz.AtYourService-$VERSION-$ARCH.msi"

	& "${env:WIX}bin\light.exe" -nologo -cultures:null -out $MsiFilename 'dispsvc.wixobj'

	if ($LastExitCode -ne 0)
	{
		exit $LastExitCode
	}

	Remove-Item 'dispsvc.wix*'
	Remove-Item 'rhubarb-geek-nz.AtYourService-*.wixpdb'

	if ($CertificateThumbprint)
	{
		$codeSignCertificate = Get-ChildItem -path Cert:\ -Recurse -CodeSigningCert | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }

		if ($codeSignCertificate -and ($codeSignCertificate.Count -eq 1))
		{
			$result = Set-AuthenticodeSignature -FilePath $MsiFilename -HashAlgorithm 'SHA256' -Certificate $codeSignCertificate -TimestampServer 'http://timestamp.digicert.com'
		}
		else
		{
			Write-Error "Error with certificate - $CertificateThumbprint"
		}
	}
}
