<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%response.Expires=5%> 
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" /> 
<title>Nansha Master Production Plan</title>
<link rel="stylesheet" type="text/css" href="GanttChart.css" />
<link rel="stylesheet" type="text/css" href="GanttChartColor.css" />
</head>
<body>
<p id="PLANT_headline1" style="position:relative;text-align:center;visibility:hidden" ><b><font face="Times New Roman, Times, serif" size="+2">Nansha 
   Production Schedule</font></b></p><p id="PLANT_headline2" style="text-align:center;font-size:small;"></p>
<p> <%
Function IIf(i,j,k)
	If i Then IIf = j Else IIf = k
End Function


Function showDate(dt1,i)
    dt = dateadd("d",i,dt1)
	d = day(dt)
    if d<10 then 
        days="0" & trim(d)
    else
        days=trim(d)
    end if

    dim monthnames(13)

    monthnames(0) = ""
    monthnames(1) = "Jan "
    monthnames(2) = "Feb "
    monthnames(3) = "Mar "
    monthnames(4) = "Apr "
    monthnames(5) = "May "
    monthnames(6) = "Jun "
    monthnames(7) = "Jul "
    monthnames(8) = "Aug "
    monthnames(9) = "Sep "
    monthnames(10) = "Oct "
    monthnames(11) = "Nov "
    monthnames(12) = "Dec "

    dim weeknames(8)

    weeknames(0) = " "
    weeknames(1) = "Mon "
    weeknames(2) = "Tue "
    weeknames(3) = "Wed "
    weeknames(4) = "Thu "
    weeknames(5) = "Fri "
    weeknames(6) = "Sat "
    weeknames(7) = "Sun "

    showDate = weeknames(weekday(dt,2)) & monthnames(month(dt)) & days

End Function



dt = cdate(request("v_date"))

response.write "<p></p>"

d = day(dt)
if d<10 then 
   days="0" & trim(d)
else
   days=trim(d)
end if

m=month(dt)
if m<10 then 
   months="0" & trim(m)
else
   months=trim(m)
end if

years=trim(year(dt)) 

set objfilesys = server.createobject("scripting.filesystemobject")
intranetLocation = "./" 
continueOrNot = true

For Line_no=1 to 33
   name="mps" & years & months & days & "l" & Line_no & ".gif"

   if objfilesys.fileexists(server.mappath(intranetLocation & name)) then
        continueOrNot = false
        set objfile= objfilesys.getfile(server.mappath(intranetLocation & name))

        '  Show Run Chart
        response.write "<table width='980' border='0'>"
        '   response.write "<TR><td colspan='8'><img src='\sh_intranet\SCM\MPS\" & name & "' width='800' height='280' alt='\sh_intranet\SCM\MPS\" & name & "'></td></TR>"
        response.write "<TR><td bgcolor='#FFFFCC' align='center'><B> Line " & Line_no & " last updated at " & objfile.datelastmodified & "</B></td></TR>"
        response.write "<TR><td><img src='" & intranetLocation & name & "' alt='" & intranetLocation & name & "'></td></TR>"
        response.write "</table>"
   end if
next

'if we are going to display production schedule in a new way
dim widthForHeading
widthForHeading = ""

dim fileTimeForTextFile
fileTimeForTextFile = "auto"
if continueOrNot then
        name="mps.txt"


       if objfilesys.fileexists(server.mappath(intranetLocation & name)) then

            set objfile1= objfilesys.getfile(server.mappath(intranetLocation & name))
            fileTimeForTextFile = "Last update time  === " & objfile1.datelastmodified & " ==="
            

            Set adoStream = Server.CreateObject("ADODB.Stream")
     
            adoStream.Charset = "UTF-8" 
            adoStream.Open 
            adoStream.LoadFromFile server.mappath(intranetLocation & name) 'change this to point to your text file

            Dim mulArray 
            mulArray = split(adoStream.ReadText,"^")

            set adoStream = nothing

            curProdLine = 999 
            linesCount = 0 'counter for production line
            For intIndex = LBound(mulArray) To UBound(mulArray) 
                    dim mulItems
                    mulItems = split(mulArray(intIndex),"@")
                    daysToCover = cdate(mulItems(3)) - cdate(mulItems(2)) + 2 ' startTime1, endTime1
                    dim headerLineInnerText
                    headerLineInnerText = ""
                    pixelsPerDay = mulItems(1)
                    startD = cdate(mulItems(2))
                    offSetFromStart = mulItems(4)
                    workingMinutes = mulItems(5)
                    txt_lot_no = mulItems(6)
                    txt_item_no = mulItems(7)
                    planned_production_qty = mulItems(8)
                    txt_order_key = mulItems(9)
                    txt_VIP = mulItems(10)
                    txt_remark = mulItems(11)
                    leftPosForRSDspan = mulItems(12)
                    SPANandETD = mulItems(13)
                    marginTop = mulItems(14)
                    pullScrew = mulItems(15)
                    txt_process_technics = mulItems(16)
                    linesList = mulItems(17)
                    if len(pullScrew) > 0 then
                        pullScrew = "<span class = 'arrow-right' style='top:3px;left:3px;border-left:10px solid " & pullScrew & ";position:absolute;'>&nbsp;</span>"
                    end if 

                    if mulItems(0) <> curProdLine then
                        linesCount = linesCount + 1
                        if linesCount > 1 then
                            response.write "</div>"
                        end if
                        curProdLine = mulItems(0)

                        linesArray = split(linesList,",")
                       
                        For j1 = Lbound(linesArray) to Ubound(linesArray)
                            if linesArray(j1) <> curProdLine then
                                headerLineInnerText = headerLineInnerText & "&nbsp;<a href='#L" & linesArray(j1) & "'>" & linesArray(j1) & "</a>&nbsp;"
                            else
                                headerLineInnerText = headerLineInnerText & "&nbsp;&nbsp;Line&nbsp;" & curProdLine & "&nbsp;"
                            end if
                        Next

                        
                        
                        
                        response.write "<div class = 'lineHeader1' id='L" & curProdLine & "'  style='left:0px;width:" & (daysToCover + 1)* pixelsPerDay -1  & "px;'>" & headerLineInnerText & "</div>"
                        response.write "<div class='scalar1' style='border: 0px solid black;z-index:1;background-color:white;position:relative;left:0px;width:" & (daysToCover + 1)* pixelsPerDay & "px;'>"
                        widthForHeading = (daysToCover + 1 )* pixelsPerDay & "px"
                        For j1 = 0 to daysToCover
                            response.write "<div class='sc1' style='left:" & pixelsPerDay * j1  & "px;width:" & (pixelsPerDay -1) & "px;background-color:" & IIf(j1 Mod 2 = 1, "#00FFFF;", "#FFFFFF;") & "'>" & showDate(startD,j1) & "</div>"
                        Next
                        response.write "</div>"
                        response.write "<div class='gattArea1' style='width:" & (daysToCover + 1)* pixelsPerDay  & "px;'>" 

                        For j1 = 0 to daysToCover
                            response.write "<div class='bg1' style='left:" & pixelsPerDay * j1 + 1 & "px;width:" & (pixelsPerDay - 1) & "px;'></div>"
                        Next

                    end if  
                        
                    response.write "<div class='gantt1' style='left:" & CLng(offSetFromStart * pixelsPerDay / (24 * 60))  & "px;width:" & Clng(workingMinutes * pixelsPerDay / (24 * 60) - 2) & "px;'><span style='" & leftPosForRSDspan & "position:relative;'>" & SPANandETD & "&nbsp;" & "</span><span style='color:red;position:absolute;white-space:pre;left:" & (CLng(workingMinutes * pixelsPerDay / (24 * 60)) - 3) & "px;" & marginTop & "'>" & pullScrew & "</span>" & _
                                  "<span style='position:absolute;white-space:pre;left:" & (CLng(workingMinutes * pixelsPerDay / (24 * 60)) + 10) & "px;'>" & txt_lot_no & "&nbsp;&nbsp;&nbsp;" & txt_item_no & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & planned_production_qty & "&nbsp;&nbsp;" & txt_process_technics  & txt_VIP & txt_order_key  & " " & txt_remark & "</span></div>"
                    
                    
            Next 

            if linesCount > 0 then
                 response.write "</div>"
            end if

       end if

end if

set objfilesys = nothing


dim resultString
resultString = request("postSchedule")
if len(resultString) > 0 then
    dim fs,f
    set fs=Server.CreateObject("Scripting.FileSystemObject") 
    set f=fs.CreateTextFile(server.mappath(intranetLocation & "test.txt"),true)
    f.write("Hello World!")
    f.write("How are you today?")
    f.write(resultString)
    f.close
    set f=nothing
    set fs=nothing
end if




%> </p>

<script type ="text/javascript" language="javaScript">
<!--
    var msg = '<%  = fileTimeForTextFile %>';
    msg = "   " + msg;
    for (i = 0; i < 128; i++) {
        msg += " ";
    };
    var scrollText = msg.split("");
    var seq = 0;
    var len = scrollText.length;

    function Scroll() {
        window.status = scrollText.slice(seq, len).join("") + scrollText.slice(0, seq).join("");
        seq++;
        if (seq >= len) { seq = 0; };
        window.setTimeout("Scroll();", 320);
    };
    Scroll();

    function setStartPostionForHeadline() {
        window.document.getElementById('PLANT_headline1').style.width = '<% = widthForHeading %>';
        window.document.getElementById('PLANT_headline1').style.visibility = "visible";

        window.document.getElementById('PLANT_headline2').innerHTML = '(<%  = fileTimeForTextFile %>)'
        window.document.getElementById('PLANT_headline2').style.width = '<% = widthForHeading %>';
        window.document.getElementById('PLANT_headline2').style.visibility = "visible";
    };
    setStartPostionForHeadline();

-->
</script>


</body>
</html>
