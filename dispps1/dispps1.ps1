# Copyright (c) 2024 Roger Brown.
# Licensed under the MIT License.

param(
	$ProgID = 'RhubarbGeekNz.AtYourService',
	$Method = 'GetMessage',
	$Hint=@(1, 2, 3, 4, 5)
)

Add-Type -TypeDefinition @"
	using System;
	using System.Runtime.InteropServices;

	namespace RhubarbGeekNz.AtYourService
	{
		public class InterOp
		{
			[DllImport("ole32.dll", PreserveSig = false)]
			static extern int CoSetProxyBlanket(
				IntPtr pProxy,
				uint dwAuthnSvc,
				uint dwAuthzSvc,
				[MarshalAs(UnmanagedType.LPWStr)] string pServerPrincName,
				uint dwAuthnLevel,
				uint dwImpLevel,
				IntPtr pAuthInfo,
				uint dwCapabilities);

			public static void SetSecurity(object dispatch)
			{
				CoSetProxyBlanket(Marshal.GetIDispatchForObject(dispatch), uint.MaxValue, 0, null, 4, 3, IntPtr.Zero, 0);
			}
		}
	}
"@

$helloWorld = New-Object -ComObject $ProgID

[RhubarbGeekNz.AtYourService.InterOp]::SetSecurity($helloWorld)

foreach ($h in $hint)
{
	$result = $helloWorld.$Method($h)

	"$h $result"
}
