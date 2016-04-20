<%
	function getTempFolder()
		Set newfs = server.CreateObject("Scripting.FileSystemObject")
		getTempFolder = newfs.GetSpecialFolder(2)
		set newfs = nothing
	end function
	function createTempFolder(folder)
		Set newfs = server.CreateObject("Scripting.FileSystemObject")
		if not newfs.FolderExists(folder) then
			newfs.CreateFolder(folder)
		end if
		set newfs = nothing
		createTempFolder = folder
	end function
%>
