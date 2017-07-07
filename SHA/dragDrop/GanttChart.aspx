<%@ Page Language="VB" AutoEventWireup="false" CodeFile="GanttChart.aspx.vb" Inherits="dragDrop_GanttChart" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
   "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Schedule Gantt chart</title>
<link rel="shortcut icon" href="../App_Themes/favicon.ico" type="image/x-icon" />
<link rel="stylesheet" type="text/css" href="GanttChart.css" />
<link rel="stylesheet" type="text/css" href="../App_Themes/GanttChartColor.css"  />
<script type="text/Javascript" src ="GanttChart.js" >
</script>
</head>
<body>
<div id="divContext" style="z-index:2;border: 1px solid blue; display: none; position: absolute"></div>
    <div style='z-index:1;position:fixed;left:0px;top:0px;height:1.2em;width:100%;font-family: Arial, Helvetica, sans-serif; font-size: large; text-decoration: underline; font-weight: bold; color: #FFFFFF; background-color: #00FFFF;'><a href="OrderDetail.aspx"><%= System.Configuration.ConfigurationManager.AppSettings("plantName")%> production schedule</a><a href="javascript:if (confirm('Are you sure you want to update the schedule?')) {PX.newExcelFrom('../ProductionSchedule/DtExchgPrdctn.aspx','FPS');};" style='z-index:1;font-weight:lighter;left:76px;position:relative;'>Production schedule<span id='FPS' ></span></a>
    <a id="logined" href="../Makerelated/scheduleDifferencesInBetweens.aspx" style='z-index:1;font-weight:lighter;left:96px;position:relative;'>Differences<span id='dfCount' ></span></a>
    <a href="javascript:if (confirm('Are you sure you want to update the schedule?')) {PX.newExcelFrom('../checkschedule/DtExchgChecker.aspx','FRS');};" style='z-index:1;font-weight:lighter;left:152px;position:relative;' >Review schedule<span id='FRS'></span></a>
    <a  href="../plansetting/planparam.aspx" style='position:fixed;right:30px;' >Timing</a>
    </div>
    <p id="lines" 
        style='z-index:1;position:fixed;left:0px;top:1.2em; width:100%;margin:0px 0px 0px 0px;background-color: #00FFFF;'>Production lines list</p>
    <div id="scalar"></div><div id="ganttZone" style="z-index:0"></div><div id="bckgrndImg"></div>
</body>
</html>
