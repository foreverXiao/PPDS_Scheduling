<%@ Page Language="VB" AutoEventWireup="false" CodeFile="checker.aspx.vb" Inherits="checkschedule_checker" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="refresh" content="3600" />
<link rel="shortcut icon" href="../App_Themes/favicon.ico" type="image/x-icon" />
<link rel="stylesheet" type="text/css" href="../App_Themes/GanttChart.css" />
<link rel="stylesheet" type="text/css" href="../App_Themes/GanttChartColor.css" />
<script type="text/Javascript" src ="checker.js" >
</script>
    <title>Schedule for Checker</title>
</head>
<body>
<div id="divContext" style="z-index:2;border: 1px solid blue; display: none; position: absolute"></div>
    <div style='z-index:1;position:fixed;left:0px;top:0px;height:1.2em;width:100%;font-family: Arial, Helvetica, sans-serif; font-size: large; text-decoration: underline; font-weight: bold; color: purple; background-color: #00FFFF;'><%= System.Configuration.ConfigurationManager.AppSettings("plantName")%> production schedule</div>
    <p id="lines" 
        style='z-index:1;position:fixed;left:0px;top:1.2em; width:100%;margin:0px 0px 0px 0px;background-color: #00FFFF;'>Production lines list</p>
    <div id="scalar"></div><div id="ganttZone" style="z-index:0"></div><div id="bckgrndImg"></div>
</body>
</html>