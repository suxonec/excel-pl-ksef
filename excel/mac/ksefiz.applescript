-- Launcher for the KSeF IZ Tool Excel buttons on macOS (Excel's sandbox blocks VBA Shell).
-- Installed by scripts/install_mac_helper.sh into
--   ~/Library/Application Scripts/com.microsoft.Excel/ksefiz.scpt
-- and called from VBA as AppleScriptTask("ksefiz.scpt", "StartJob", param).
-- param = helper TAB jobDir TAB arg TAB arg ...  Only an executable named "ksefiz" can be
-- started, only for a job folder inside Excel's container, and every argument is quoted.

on StartJob(param)
	set oldDelims to AppleScript's text item delimiters
	set AppleScript's text item delimiters to tab
	set parts to text items of param
	set AppleScript's text item delimiters to oldDelims
	if (count of parts) < 2 then error "ksefiz: expected helper and job folder"
	set helper to item 1 of parts
	set jobDir to item 2 of parts

	if helper does not end with "/ksefiz" then error "ksefiz: helper_path must point to the ksefiz executable"
	try
		do shell script "test -x " & quoted form of helper
	on error
		error "ksefiz: helper not found or not executable: " & helper
	end try

	set containerRoot to (POSIX path of (path to home folder)) & "Library/Containers/com.microsoft.Excel/Data/"
	if jobDir does not start with containerRoot then error "ksefiz: job folder must be inside Excel's container"
	if jobDir contains "/../" then error "ksefiz: invalid job folder"

	set cmd to quoted form of helper & " excel-run --job " & quoted form of jobDir
	repeat with i from 3 to count of parts
		set cmd to cmd & " " & quoted form of (item i of parts)
	end repeat
	set logFile to quoted form of (jobDir & "/log.txt")
	set exitFile to quoted form of (jobDir & "/exit")
	do shell script "(" & cmd & " > " & logFile & " 2>&1; echo $? > " & exitFile & ") > /dev/null 2>&1 &"
	return "started"
end StartJob

on Ping(param)
	return "ksefiz launcher ok"
end Ping
