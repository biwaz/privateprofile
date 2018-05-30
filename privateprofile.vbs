option explicit

dim objFileSys
set objFileSys = createobject("Scripting.FileSystemObject")

function GetPrivateProfile(filepath, byval section, byval key)
	dim objRegExp, objProfile, objCurrent, line, s, e, n

	set objRegExp = new regexp
	objRegExp.pattern = "^[ \t]*(([^ \t].*[^ \t])|[^ \t]?)[ \t]*$"

	if not objFileSys.fileexists(filepath) then
		if isnull(section) or isnull(key) then set GetPrivateProfile = nothing else GetPrivateProfile = null
		exit function
	end if

	with objFileSys.opentextfile(filepath, 1)
		set objProfile = createobject("Scripting.Dictionary")
		set objCurrent = nothing
		do until .atendofstream = true
			line = objRegExp.replace(.readline, "$1")
			s = left(line, 1)
			if 0 < len(line) and s <> ";" then
				if s = "[" then
					n = instrrev(line, "]")
					if 1 < n then s = mid(line, 2, n - 2) else s = mid(line, 2)
					s = ucase(objRegExp.replace(s, "$1"))
					if objProfile.exists(s) then
						set objCurrent = nothing
					else
						set objCurrent = createobject("Scripting.Dictionary")
						set objProfile(s) = objCurrent
					end if
				else
					if not objCurrent is nothing then
						n = instr(line, "=")
						if 1 < n then
							s = ucase(objRegExp.replace(left(line, n - 1), "$1"))
							if not objCurrent.exists(s) then
								line = objRegExp.replace(mid(line, n + 1), "$1")
								if 1 < len(line) then
									e = right(line, 1) ' quate process
									if e = left(line, 1) then
										if 0 < instr("""'", e) then line = mid(line, 2, len(line) - 2)
									end if
								end if
								objCurrent(s) = line
							end if
						elseif n < 1 then
							s = ucase(line)
							if not objCurrent.exists(s) then objCurrent(s) = ""
						end if
					end if
				end if
			end if
		loop
		.close
	end with

	if isnull(section) then
		set GetPrivateProfile = objProfile
		exit function
	end if

	section = ucase(section)
	if not objProfile.exists(section) then
		if isnull(key) then set GetPrivateProfile = nothing else GetPrivateProfile = null
		exit function
	end if

	set objCurrent = objProfile(section)

	if isnull(key) then
		set GetPrivateProfile = objCurrent
		exit function
	end if

	key = ucase(key)
	if objCurrent.exists(key) then GetPrivateProfile = objCurrent(key) else GetPrivateProfile = null
end function

sub WritePrivateProfile(filepath, section, key, byval data)
	dim objFile, objTemp, objRegExp, section_, key_, temp, org, stat, buffer, line, s, e, n
	if objFileSys.fileexists(filepath) then set objFile = objFileSys.opentextfile(filepath, 1) else set objFile = nothing

	temp = objFileSys.buildpath(objFileSys.getspecialfolder(2), objFileSys.gettempname())
	set objTemp = objFileSys.createtextfile(temp)
	if objTemp is nothing then
		if not objFile is nothing then objFile.close
		exit sub
	end if

	set objRegExp = new RegExp
	objRegExp.pattern = "^[ \t]*(([^ \t].*[^ \t])|[^ \t]?)[ \t]*$"

	if vartype(data) = 8 then ' quate process
		if 0 < len(data) then
			s = left(data, 1)
			e = right(data, 1)
			if 0 < instr(" 	", s) then
				data = """" & data & """"
			elseif 0 < instr(" 	", e) then
				data = """" & data & """"
			elseif 1 < len(data) then
				if s = e then
					if s = "'" then
						data = """" & data & """"
					elseif s = """" then
						data = """" & data & """"
					end if
				end if
			end if
		end if
	end if

	if vartype(key) = 8 then key_ = ucase(key)

	section_ = ucase(section)
	stat = false
	buffer = ""
	if not objFile is nothing then
		do until objFile.atendofstream = true
			line = objFile.readline
			s = objRegExp.replace(line, "$1")
			if left(s, 1) = "[" then
				n = instrrev(s, "]")
				if 1 < n then s = mid(s, 2, n - 2) else s = mid(s, 2)
				if ucase(objRegExp.replace(s, "$1")) = section_ then
					if vartype(key) = 8 then objTemp.writeline line

					do until objFile.atendofstream = true
						org = objFile.readline
						line = objRegExp.replace(org, "$1")
						s = left(line, 1)
						if 0 < len(line) and s <> ";" then
							if s = "[" then
								stat = true
								if vartype(key) = 8 and vartype(data) = 8 then objTemp.writeline key & "=" & data
								buffer = buffer & org & vbCrLf
								exit do
							end if

							if vartype(key) = 8 then
								objTemp.write buffer
								buffer = ""

								n = instr(line, "=")
								if 1 < n then line = left(line, n - 1)
								if ucase(objRegExp.replace(line, "$1")) = key_ then
									stat = true
									if vartype(data) = 8 then
										objRegExp.pattern = "^([^=]*)=(.*[^ \t])?([ \t]*)$"
										objTemp.writeline objRegExp.replace(org, "$1=" & data & "$3")
									end if
									exit do
								end if

								objTemp.writeline org
							else
								buffer = ""
							end if
						else
							buffer = buffer & org & vbCrLf
						end if
					loop
					if not stat then
						stat = true
						if vartype(key) = 8 and vartype(data) = 8 then objTemp.writeline key & "=" & data
					end if
					objTemp.write buffer
					exit do
				else
					objTemp.writeline line
				end if
			else
				objTemp.writeline line
			end if
		loop
	end if

	if stat then
		do until objFile.atendofstream = true
			objTemp.writeline objFile.readline
		loop
	else
		if vartype(key) = 8 then
			objTemp.writeline "[" & section & "]"
			objTemp.writeline key & "=" & data
		end if
	end if

	objTemp.close
	if not objFile is nothing then
		objFile.close
		objFileSys.deletefile filepath
	end if

	if objFileSys.fileexists(filepath) then
		if objFileSys.fileexists(temp) then objFileSys.deletefile temp
	else
		objFileSys.movefile temp, filepath
	end if
end sub
