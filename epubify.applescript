property script_title : "AppleScript droplet for pandoc docx-epub file conversion"
property script_version : "1.1"
property pandoc_version : "2.2.1"

(*
Adapted by Simon Collinson 2019 to restrict to docx-epub3 conversion only.

DISCLAIMER
If you do not agree with the following terms, please do not use, install, modify or redistribute this script. This script is provided on an "AS IS" basis. We make no warranties, express or implied, regarding this script or its use and operation alone or in combination with other products. In no event shall we be liable for any special, indirect, incidental or consequential damages arising in any way out of the use, reproduction, modification and/or distribution of this script, however caused and whether under theory of contract, tort (including negligence), strict liability or otherwise, even if we have been advised of the possibility of such damage.
This script was not created by the creators of pandoc. We created it to fill our own need, which it does well, but it is not an especially robust script. Read the instructions and examine the script carefully before using it, and make backup copies of your files so that you can restore the files in case you make any mistakes.
INSTRUCTIONS
This AppleScript droplet is designed to make and run a shell script that uses pandoc (the "swiss-army knife" for document conversion) to convert one or more files whose icons are dragged onto the droplet icon in the OS X Finder. The script converts multiple input files recursively, instead of concatenating multiple input files as pandoc does when run from the command line. We find this droplet to be much less tedious to use than the command line. To save the script as a droplet, open the script in Script Editor and save the script with file format: Application.
IMPORTANT REQUIREMENTS
Before trying to convert files, make sure that all input files are the same format, such as all markdown files or all html files. This script is designed to convert files of the same format. The script does not check the format of the input files, so it relies on you to accurately specify the format. The droplet only accepts files, not folders; folders are ignored.
Make sure that the character encoding of all input files is UTF-8. If you're not sure of the character encoding, in Terminal.app you can check the character encoding of a file by entering the command: file -I [filename]
Make sure that all input files have filename extensions (for example: .html, .rtf). It is OK if the extensions are hidden in the Finder; just make sure that each file has a filename extension, hidden or not. Make sure that there are no periods (except for the period before the extension) in the filenames of the input files. Make sure that there are no quotation marks (or other forbidden characters that would make the shell script fail) in the filenames.
This script was written for pandoc 2.2.1, so if you are using a newer or older version of pandoc you will want to change the list of input formats and output formats in the script to reflect the formats supported by your version of pandoc, and you will want to make sure that the syntax after "do shell script" is correct for your version of pandoc.
Before running the script you must install pandoc either directly from http://pandoc.org or using (our preferred method) a package manager such as Fink, Homebrew, or MacPorts.
Before running the script, change the property pandoc_path below to reflect the path of the pandoc command on your computer. In Terminal.app you can check the path of the pandoc command by entering the command: type -a pandoc
*)

property pandoc_path : "/usr/local/bin/pandoc"

on open dropped_files
	repeat with i in dropped_files
		set the file_info to info for i
		-- Exclude folders.
		if the folder of file_info is false then
			-- Get the path, name, and short name of each input file, assuming there is only one period in the filename.
			set file_path to POSIX path of i
			tell application "System Events"
				set file_container to POSIX path of (container of i)
			end tell
			set file_name to the name of file_info
			set {text_delimiters, my text item delimiters} to {my text item delimiters, "."}
			set file_short_name to first text item of file_name
		end if
	end repeat
	
	set input_format to "docx"
	set output_format to "epub3"
	set output_extension to ".epub"
	set shell_script to pandoc_path & " -f " & input_format & " -t " & output_format & space & " -o " & "'" & file_container & "/" & file_short_name & "-output" & output_extension & "'" & space & quoted form of file_path
	do shell script shell_script
end open
