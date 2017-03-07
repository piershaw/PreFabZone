<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">
'*** Pure ASP File Upload Add On Pack -----------------------------------------
' Copyright 2001-2002 (c) George Petrov, www.DMXzone.com
'
' Version: 1.7.0
'------------------------------------------------------------------------------

Sub DeleteFileBeforeRecord(DF_filesStr,DF_path,MM_editConnection,MM_editTable,MM_editColumn,MM_recordId,DF_suffix)
  if DF_path <> "" and right(DF_path,1) <> "/" then DF_path = DF_path & "/"
  Dim DF_fso, DF_files, DF_filesArr, DF_file, DF_fullFile
  Set DF_fso = CreateObject("Scripting.FileSystemObject")
  set DF_files = Server.CreateObject("ADODB.Recordset")
  DF_files.ActiveConnection = MM_editConnection
  DF_files.Source = "SELECT " & DF_filesStr & " FROM " & MM_editTable & " WHERE " & MM_editColumn & " IN (" & MM_recordId & ")"
  DF_files.CursorType = 0
  DF_files.CursorLocation = 2
  DF_files.LockType = 3
  DF_files.Open()
  DF_filesArr = split(DF_filesStr,",")
  while not DF_files.EOF
    for DF_fi = 0 to UBOUND(DF_filesArr)
      DF_file = Trim(DF_files.Fields.Item(Trim(DF_filesArr(DF_fi))).Value&"")
      if DF_file <> "" then
        DF_fullFile = Server.MapPath(DF_path & DF_file)
        if DF_fso.FileExists(DF_fullFile) then
          DF_fso.DeleteFile(DF_fullFile)	 
        end if	
        if DF_suffix <> "" then
          DF_fullFile = Server.MapPath(DF_path & getThumbnailName(DF_suffix,DF_file))
          if DF_fso.FileExists(DF_fullFile) then
            DF_fso.DeleteFile(DF_fullFile)	 
          end if
        end if
      end if  
    next
    DF_files.MoveNext()
  wend
  DF_files.Close()
End Sub

Sub DeleteFileBeforeUpdate(DOF_path,MM_fieldsStr,MM_columnsStr,MM_editConnection,MM_editTable,MM_editColumn,MM_recordId,DOF_suffix)
  Dim DOF_keys, DOF_i, DOF_curKey, DOF_filesStr, DOF_formFieldStr, DOF_i2
  Dim DOF_MM_fields,DOF_MM_columns
  DOF_MM_fields = Split(MM_fieldsStr,"|")
  DOF_MM_columns = Split(MM_columnsStr,"|")  
  DOF_filesStr = ""
  DOF_formFieldStr = ""
  if DOF_path <> "" and right(DOF_path,1) <> "/" then DOF_path = DOF_path & "/"
  DOF_keys = UploadRequest.Keys
  for DOF_i = 0 to UploadRequest.Count - 1
    DOF_curKey = DOF_keys(DOF_i)
    if UploadRequest.Exists(DOF_curKey) then
      if UploadRequest.Item(DOF_curKey).Exists("FileName") then
        if UploadRequest.Item(DOF_curKey).Item("FileName") <> "" then
          For DOF_i2 = LBound(DOF_MM_fields) To UBound(DOF_MM_fields) Step 2
            if LCase(DOF_MM_fields(DOF_i2)) = LCase(DOF_curKey) then
              if DOF_filesStr <> "" then
                DOF_filesStr = DOF_filesStr & "," & DOF_MM_columns(DOF_i2)
                DOF_formFieldStr = DOF_formFieldStr & "," & DOF_curKey
              else
                DOF_filesStr = DOF_MM_columns(DOF_i2)	
                DOF_formFieldStr = DOF_curKey
              end if
            end if
          Next
        end if
      end if
    end if
  next	
  if DOF_filesStr <> "" then
    Dim DOF_fso, DOF_files, DOF_filesArr, DOF_formFieldArr, DOF_file, DOF_fullFile
    Set DOF_fso = CreateObject("Scripting.FileSystemObject")
    set DOF_files = Server.CreateObject("ADODB.Recordset")
    DOF_files.ActiveConnection = MM_editConnection
    DOF_files.Source = "SELECT " & DOF_filesStr & " FROM " & MM_editTable & " WHERE " & MM_editColumn & " = " & MM_recordId
    DOF_files.CursorType = 0
    DOF_files.CursorLocation = 2
    DOF_files.LockType = 3
    DOF_files.Open()
    DOF_filesArr = split(DOF_filesStr,",")
    DOF_formFieldArr = split(DOF_formFieldStr,",")	
    for DOF_fi = 0 to UBOUND(DOF_filesArr)
      DOF_file = Trim(DOF_files.Fields.Item(Trim(DOF_filesArr(DOF_fi))).Value&"")
      if UploadRequest.Exists(DOF_formFieldArr(DOF_fi)) and DOF_file <> "" then
        if UploadRequest.Item(DOF_formFieldArr(DOF_fi)).Exists("FileName") then
          if LCase(DOF_file) <> LCase(UploadRequest.Item(DOF_formFieldArr(DOF_fi)).Item("FileName")) then
            DOF_fullFile = Server.MapPath(DOF_path & DOF_file)
            if DOF_fso.FileExists(DOF_fullFile) then
              DOF_fso.DeleteFile(DOF_fullFile)	 
            end if
            if DOF_suffix <> "" then
              DOF_fullFile = Server.MapPath(DOF_path & getThumbnailName(DOF_suffix,DOF_file))
              if DOF_fso.FileExists(DOF_fullFile) then
                DOF_fso.DeleteFile(DOF_fullFile)	 
              end if
            end if            
          end if
        end if
      end if
    next
    DOF_files.Close()
  end if  
End Sub

Sub RenameUploadedFiles(REUF_path,REUF_nameConflict,REUF_renameMask)
  Dim REUF_keys, REUF_i, REUF_curKey, REUF_fileName, REUF_fileNameArr, REUF_lastPos, REUF_FileExist, Begin_Name_Num
  Dim REUF_fso, REUF_curPath, REUF_curName, REUF_curExt, REUF_newFileName, REUF_fileNameOnly, REUF_FullFileName
  if REUF_path <> "" and right(REUF_path,1) <> "/" then REUF_path = REUF_path & "/"
  Set REUF_fso = CreateObject("Scripting.FileSystemObject")  
  REUF_keys = UploadRequest.Keys
  for REUF_i = 0 to UploadRequest.Count - 1
    REUF_curKey = REUF_keys(REUF_i)
    if UploadRequest.Exists(REUF_curKey) then
      if UploadRequest.Item(REUF_curKey).Exists("FileName") then    
        if UploadRequest.Item(REUF_curKey).Item("FileName") <> "" then
					REUF_fileName = UploadRequest.Item(REUF_curKey).Item("Value")
					if REUF_fileName <> "" then
						REUF_curPath = "" : REUF_curName = "" : REUF_curExt = ""
						REUF_lastPos = InStrRev(REUF_fileName,"/")
						if REUF_lastPos > 0 then
							REUF_curPath = mid(REUF_fileName,1,REUF_lastPos)	
							REUF_curName = mid(REUF_fileName,REUF_lastPos+1,Len(REUF_fileName)-REUF_lastPos)	
						else
							REUF_curName = REUF_fileName	
						end if
						REUF_fileNameOnly = REUF_curName
						REUF_lastPos = InStrRev(REUF_curName,".")
						if REUF_lastPos > 0 then
							REUF_curExt = mid(REUF_curName,REUF_lastPos+1,Len(REUF_curName)-REUF_lastPos)	
							REUF_curName = mid(REUF_curName,1,REUF_lastPos-1)
						end if
						if REUF_curPath = "" then REUF_curPath = REUF_path
						REUF_newFileName = Replace(REUF_renameMask,"##name##",REUF_curName)
						REUF_newFileName = Replace(REUF_newFileName,"##ext##",REUF_curExt)	  
						REUF_FileExist = false
						If REUF_fso.FileExists(Server.MapPath(REUF_curPath & REUF_newFileName)) Then
							REUF_FileExist = true
						End If      
						if REUF_nameConflict = "error" and REUF_FileExist then
							Response.Write "<B>File already exists!</B><br><br>"
							Response.Write "Please correct and <A HREF=""javascript:history.back(1)"">try again</a>"
							response.End
						end if
						if ((REUF_nameConflict = "over" or REUF_nameConflict = "uniq") and REUF_FileExist) or (NOT REUF_FileExist) then
							if REUF_nameConflict = "uniq" and REUF_FileExist then
								Begin_Name_Num = 0
								while REUF_FileExist    
									Begin_Name_Num = Begin_Name_Num + 1
									REUF_FullFileName = REUF_fso.GetBaseName(REUF_newFileName) & "_" & Begin_Name_Num & "." & REUF_fso.GetExtensionName(REUF_newFileName)
									REUF_FileExist = REUF_fso.FileExists(Server.MapPath(REUF_curPath & REUF_FullFileName))
								wend  
								REUF_newFileName = REUF_FullFileName
							end if
							if REUF_fso.FileExists(Server.MapPath(REUF_curPath & REUF_fileNameOnly)) then
								if REUF_nameConflict = "over" and REUF_FileExist then
									REUF_fso.DeleteFile Server.MapPath(REUF_curPath & REUF_newFileName)
								end if
									REUF_fso.MoveFile Server.MapPath(REUF_curPath & REUF_fileNameOnly), Server.MapPath(REUF_curPath & REUF_newFileName)
							end if
							if REUF_path = "" then
								UploadRequest.Item(REUF_curKey).Item("Value") = REUF_curPath & REUF_newFileName		
							else
								UploadRequest.Item(REUF_curKey).Item("Value") = REUF_newFileName			
							end if
						end if
					end if
			  end if		
      end if
    end if
  next	
End Sub

Sub MailUploadedFiles(MUF_path,MUF_FromName,MUF_FromEmail,MUF_ToName,MUF_ToEmail,MUF_Bcc,MUF_Subject,MUF_Body,MUF_SendHtml,MUF_checkErrors,MUF_deleteFiles,MUF_RedirectURL,MUF_SmtpServer,MUF_MailerType)
  Dim MUF_keys, MUF_i, MUF_curKey, MUF_fileName, MUF_fso, MUF_objMail, MUF_From, MUF_To
  if MUF_path <> "" and right(MUF_path,1) <> "/" then MUF_path = MUF_path & "/"
  MUF_From = MUF_FromEmail : MUF_To = MUF_To_Email
  Set MUF_fso = CreateObject("Scripting.FileSystemObject") 
  
  Select Case MUF_MailerType
  Case "CDO"
    if MUF_FromName <> "" then MUF_From = """" & MUF_FromName & """ <" & MUF_FromEmail & ">"
    if MUF_ToName <> "" then MUF_To = """" & MUF_ToName & """ <" & MUF_ToEmail & ">"  
  
    Set MUF_objMail = Server.CreateObject("CDONTS.NewMail")
    MUF_objMail.From = MUF_From
    MUF_objMail.To = MUF_To
    If MUF_Bcc <> "" then MUF_objMail.Bcc = MUF_Bcc
    If MUF_SendHtml then
      MUF_objMail.MailFormat = 0 ' HTML Mail
      MUF_objMail.BodyFormat = 0 ' HTML Mail    
    end if	
    MUF_objMail.Subject = MUF_Subject
    MUF_objMail.Body = MUF_Body
    MUF_keys = UploadRequest.Keys
    for MUF_i = 0 to UploadRequest.Count - 1
      MUF_curKey = MUF_keys(MUF_i)
      if UploadRequest.Exists(MUF_curKey) then
        if UploadRequest.Item(MUF_curKey).Exists("FileName") then    
          MUF_fileName = UploadRequest.Item(MUF_curKey).Item("FileName")
          if MUF_fileName <> "" then
            MUF_fileName = UploadRequest.Item(MUF_curKey).Item("Value")
            if MUF_fso.FileExists(Server.MapPath(MUF_Path & MUF_fileName)) then
              MUF_objMail.AttachFile Server.MapPath(MUF_Path & MUF_fileName), MUF_fileName, 1
							if MUF_deleteFiles then
								MUF_fso.DeleteFile(Server.MapPath(MUF_Path & MUF_fileName))
							end if
            end if
          end if
        end if
      end if
    next	
    on error resume next
    MUF_objMail.Send()
    Set MUF_objMail = Nothing
    if Err.number <> 0 AND MUF_checkErrors then
      Response.Write "Error occured while sending mail. Err: " & Err.number & " " & Err.Description
      Response.End
    end if
    on error goto 0
    
  Case "JMail"
    set MUF_objMail = Server.CreateObject("JMail.SMTPMail")
    MUF_objMail.Silent = true
    MUF_objMail.ISOEncodeHeaders = false
    MUF_objMail.ServerAddress = MUF_SmtpServer
    MUF_objMail.SenderName = MUF_FromName
    MUF_objMail.Sender = MUF_FromEmail    
    MUF_objMail.Subject = MUF_Subject
    MUF_objMail.AddRecipientEx MUF_ToEmail, MUF_ToName
    If MUF_Bcc <> "" then MUF_objMail.AddRecipientBCC MUF_Bcc    
    If MUF_SendHtml then
      MUF_objMail.ContentType = "text/html"
    end if  
    MUF_objMail.Body = MUF_Body
    MUF_keys = UploadRequest.Keys
    for MUF_i = 0 to UploadRequest.Count - 1
      MUF_curKey = MUF_keys(MUF_i)
      if UploadRequest.Exists(MUF_curKey) then
        if UploadRequest.Item(MUF_curKey).Exists("FileName") then    
          MUF_fileName = UploadRequest.Item(MUF_curKey).Item("FileName")
          if MUF_fileName <> "" then
            MUF_fileName = UploadRequest.Item(MUF_curKey).Item("Value")
            if MUF_fso.FileExists(Server.MapPath(MUF_Path & MUF_fileName)) then
              MUF_objMail.AddAttachment Server.MapPath(MUF_Path & MUF_fileName)
							if MUF_deleteFiles then
								MUF_fso.DeleteFile(Server.MapPath(MUF_Path & MUF_fileName))
							end if
            end if
          end if
        end if
      end if
    next
    if not MUF_objMail.Execute AND MUF_checkErrors then
      Response.Write "Error occured while sending mail. " & MUF_objMail.ErrorMessage & "."
      Response.End
    end if
    Set MUF_objMail = Nothing
    
  Case "Persits"
     
    Set MUF_objMail = Server.CreateObject("Persits.MailSender")
    MUF_objMail.Host = MUF_SmtpServer
    MUF_objMail.From = MUF_From
    MUF_objMail.FromName = MUF_FromName    
    
    MUF_objMail.AddAddress MUF_ToEmail,MUF_To
    If MUF_Bcc <> "" then MUF_objMail.AddBcc MUF_Bcc
    If MUF_SendHtml then
      MUF_objMail.IsHTML = true
    end if	
    MUF_objMail.Subject = MUF_Subject
    MUF_objMail.Body = MUF_Body
    MUF_keys = UploadRequest.Keys
    for MUF_i = 0 to UploadRequest.Count - 1
      MUF_curKey = MUF_keys(MUF_i)
      if UploadRequest.Exists(MUF_curKey) then
        if UploadRequest.Item(MUF_curKey).Exists("FileName") then    
          MUF_fileName = UploadRequest.Item(MUF_curKey).Item("FileName")
          if MUF_fileName <> "" then
            MUF_fileName = UploadRequest.Item(MUF_curKey).Item("Value")
            if MUF_fso.FileExists(Server.MapPath(MUF_Path & MUF_fileName)) then
              MUF_objMail.AddAttachment Server.MapPath(MUF_Path & MUF_fileName)
							if MUF_deleteFiles then
								MUF_fso.DeleteFile(Server.MapPath(MUF_Path & MUF_fileName))
							end if
            end if
          end if
        end if
      end if
    next	
    on error resume next
    MUF_objMail.Send()
    Set MUF_objMail = Nothing
    if Err.number <> 0 AND MUF_checkErrors then
      Response.Write "Error occured while sending mail. Err: " & Err.number & " " & Err.Description
      Response.End
    end if
    on error goto 0    
    
  Case "ASPMail"

    set MUF_objMail = Server.CreateObject("SMTPsvg.Mailer")
    MUF_objMail.RemoteHost = MUF_SmtpServer
    MUF_objMail.FromName = MUF_FromName
    MUF_objMail.FromAddress = MUF_FromEmail    
    MUF_objMail.Subject = MUF_Subject
    MUF_objMail.AddRecipient MUF_ToName, MUF_ToEmail
    If MUF_Bcc <> "" then MUF_objMail.AddBCC "",MUF_Bcc    
    If MUF_SendHtml then
      MUF_objMail.ContentType = "text/html"
    end if  
    MUF_objMail.BodyText = MUF_Body
    MUF_keys = UploadRequest.Keys
    for MUF_i = 0 to UploadRequest.Count - 1
      MUF_curKey = MUF_keys(MUF_i)
      if UploadRequest.Exists(MUF_curKey) then
        if UploadRequest.Item(MUF_curKey).Exists("FileName") then    
          MUF_fileName = UploadRequest.Item(MUF_curKey).Item("FileName")
          if MUF_fileName <> "" then
            MUF_fileName = UploadRequest.Item(MUF_curKey).Item("Value")
            if MUF_fso.FileExists(Server.MapPath(MUF_Path & MUF_fileName)) then
              MUF_objMail.AddAttachment Server.MapPath(MUF_Path & MUF_fileName)
							if MUF_deleteFiles then
								MUF_fso.DeleteFile(Server.MapPath(MUF_Path & MUF_fileName))
							end if
            end if
          end if
        end if
      end if
    next
    if not MUF_objMail.SendMail AND MUF_checkErrors then
      Response.Write "Error occured while sending mail. " & MUF_objMail.Response & "."
      Response.End
    end if
    Set MUF_objMail = Nothing  
    
  End Select
    
  if MUF_RedirectURL <> "" then
    Response.Redirect MUF_RedirectURL
  end if
End Sub

function getThumbnailName(GTN_suff,GTN_filename)
  Dim GTN_NewFilename, GTN_Path, GTN_PosPath, GTN_PosExt
  if not isnull(GTN_filename) then
    GTN_PosPath = InStrRev(GTN_filename,"/")
    GTN_Path = ""
    if GTN_PosPath > 0 then
      GTN_Path = mid(GTN_filename,1,GTN_PosPath)
    end if
    GTN_PosExt = InStrRev(GTN_filename,".")
    if GTN_PosExt > 0 then
      GTN_NewFilename = GTN_Path & mid(GTN_filename,GTN_PosPath+1,GTN_PosExt-(GTN_PosPath+1)) & GTN_suff & ".jpg"
    else
      GTN_NewFilename = GTN_Path & mid(GTN_filename,GTN_PosPath+1,len(GTN_filename)-GTN_PosPath) & GTN_suff & ".jpg"
    end if
  end if
  getThumbnailName = GTN_NewFilename
end function


</SCRIPT>