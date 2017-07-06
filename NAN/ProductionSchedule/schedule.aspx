<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta http-equiv="refresh" content="3600" />
<link rel="stylesheet" type="text/css" href="GanttChart.css" />
<link rel="stylesheet" type="text/css" href="GanttChartColor.css" />
<script type="text/Javascript" src ="schedule.js" >
</script>
    <title>Production schedule</title>
</head>
<body>
<div id="divContext" style="z-index:2;border: 1px solid blue; display: none; position: absolute"></div>
    <span style='z-index:1;position:fixed;left:0px;top:0px;height:1em;width:100%;font-family: Arial, Helvetica, sans-serif; font-size: large; text-decoration: underline; font-weight: bold; color: #0000FF; background-color: #00FFFF;'><%= System.Configuration.ConfigurationManager.AppSettings("plantName")%> production schedule</span>
    <p id="lines" style='z-index:1;position:fixed;left:0px;top:1.1em;width:100%;margin:0px 0px 0px 0px;background-color: #00FFFF;'>Production lines list</p>
    <div id="scalar" 
        style='border: 0px solid black;z-index:1;position:fixed;background-color:white;top:3.6em; left: 10px;width:100%;'></div><div id="bckgrndImg" style='z-index:-1;position:fixed;'>background image</div>
</body>
</html>
