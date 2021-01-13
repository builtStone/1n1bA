<%
dim strLine,objFso,objFile,mystr,strWeb,strCnt,strId, strOld,strNew,FiletempData
response.charset="utf-8"
Response.CodePage=65001
strId =  request("id")
strCnt = -1

Set objFso = server.createobject("Scripting.FileSystemObject")
Set objFile = objFso.OpenTextFile(Server.MapPath("jycounter.cnt"))
Do While Not objFile.AtEndOfStream
    objFso =objFile.ReadLine
    mystr=split(objFso,",")
    
    strWeb = mystr(0)
    
  if strId = strWeb then
    strCnt = mystr(1)
     strOld = objFso
    exit do
  end if
    
Loop
  if strCnt = -1  then
	strCnt  =0
  else
        strCnt =  strCnt+1
        strNew =  strWeb & ","  &  strCnt

        objFile.Close
        Set objFso = server.createobject("Scripting.FileSystemObject")
        Set objFile = objFso.OpenTextFile(Server.MapPath("jycounter.cnt"))
	FiletempData = objFile.ReadAll 
	objFile.Close
        FiletempData=Replace(FiletempData,strOld,strNew) 
        Set objFile=objFSO.CreateTextFile(Server.MapPath("jycounter.cnt"),True)
	objFile.Write FiletempData
	objFile.Close
  end if


Set objFile = nothing
Set objFso  = nothing




%>document.write('<%=strCnt %>');
