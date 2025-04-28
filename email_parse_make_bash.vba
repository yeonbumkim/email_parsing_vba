Sub DisplayBashScriptCopyable()
    Dim mail As Outlook.MailItem
    Dim bodyLines() As String
    Dim line As Variant
    Dim server As String, userId As String, password As String, location As String
    Dim bashScript As String
    Dim fso As Object, fileStream As Object
    Dim tempFolder As String, filePath As String
    
    ' 선택된 메일 항목이 있는지 확인
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "메일 항목을 선택해주세요.", vbExclamation
        Exit Sub
    End If
    Set mail = Application.ActiveExplorer.Selection.Item(1)
    
    ' 메일 제목에 "데이터 납품안내"가 포함되어 있는지 확인
    If InStr(mail.Subject, "데이터 납품안내") = 0 Then
        MsgBox "해당 메일은 '데이터 납품안내' 제목을 포함하지 않습니다.", vbExclamation
        Exit Sub
    End If
    
    ' 메일 본문을 줄 단위로 분할 후 정보 추출
    bodyLines = Split(mail.Body, vbCrLf)
    For Each line In bodyLines
        If InStr(line, "서버:") > 0 Then
            server = Trim(Mid(line, InStr(line, "서버:") + Len("서버:")))
        ElseIf InStr(line, "아이디:") > 0 Then
            userId = Trim(Mid(line, InStr(line, "아이디:") + Len("아이디:")))
        ElseIf InStr(line, "패스워드:") > 0 Then
            password = Trim(Mid(line, InStr(line, "패스워드:") + Len("패스워드:")))
        ElseIf InStr(line, "위치:") > 0 Then
            location = Trim(Mid(line, InStr(line, "위치:") + Len("위치:")))
        End If
    Next
    
    ' 모든 정보가 추출되었는지 확인
    If server = "" Or userId = "" Or password = "" Or location = "" Then
        MsgBox "필요한 정보를 모두 찾지 못했습니다.", vbExclamation
        Exit Sub
    End If
    
    ' bash 스크립트 내용 구성 (디렉토리 재귀 다운로드를 위해 get -r 사용)
    ' 추가 명령어(cd $LOCAL_FOLDER, mkdir 00.RawData 등)를 EOF 블록 뒤쪽에 배치
    bashScript = "#!/bin/bash" & vbCrLf & _
                 "REMOTE_SERVER=""" & server & """" & vbCrLf & _
                 "USER=""" & userId & """" & vbCrLf & _
                 "PASSWORD=""" & password & """" & vbCrLf & _
                 "REMOTE_FOLDER=""" & location & """" & vbCrLf & _
                 "LOCAL_BASE=""/BiO/gut/""" & vbCrLf & _
                 "LOCAL_FOLDER=${LOCAL_BASE}${REMOTE_FOLDER}.v4" & vbCrLf & vbCrLf & _
                 "mkdir -p ""$LOCAL_FOLDER""" & vbCrLf & vbCrLf & _
                 "# sshpass를 이용하여 sftp 접속 (sshpass 설치 필요)" & vbCrLf & _
                 "sshpass -p ""$PASSWORD"" sftp $USER@$REMOTE_SERVER <<EOF" & vbCrLf & _
                 "cd $REMOTE_FOLDER" & vbCrLf & _
                 "lcd $LOCAL_FOLDER" & vbCrLf & _
                 "get -r *" & vbCrLf & _
                 "bye" & vbCrLf & _
                 "EOF" & vbCrLf & vbCrLf & _
                 "# 다운로드 받은 디렉토리에서 후속 작업 수행" & vbCrLf & _
                 "cd $LOCAL_FOLDER" & vbCrLf & _
                 "mkdir 00.RawData" & vbCrLf & _
                 "mv ./*/*.gz ./00.RawData" & vbCrLf & _
                 "cp ../script_20221209/V4_script/* ." & vbCrLf & _
                 "bash command_20220126.sh"
    
    ' 임시 폴더에 텍스트 파일로 저장
    tempFolder = Environ("TEMP")
    filePath = tempFolder & "\bash_script.txt"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileStream = fso.CreateTextFile(filePath, True, False)
    fileStream.Write bashScript
    fileStream.Close
    
    ' 기본 텍스트 편집기(Notepad)를 열어 파일 내용 보여주기
    Shell "notepad.exe " & filePath, vbNormalFocus
End Sub


