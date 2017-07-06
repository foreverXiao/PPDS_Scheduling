
var scalardiv = null;
var ganttZoneleft = 0;
var scArray = new Array();


var msg = "";
var interval = 240;
var spacelen = 120;
var space10 = " ";
var seq = 0;

function Scroll() {
    len = msg.length;
    window.status = msg.substring(0, seq + 1);
    seq++;
    if (seq >= len) {
        seq = spacelen;
        window.setTimeout("Scroll2();", interval);
    }
    else
        window.setTimeout("Scroll();", interval);
}
function Scroll2() {
    var out = "";
    for (i = 1; i <= spacelen / space10.length; i++) out +=
space10;
    out = out + msg;
    len = out.length;
    window.status = out.substring(seq, len);
    seq++;
    if (seq >= len) { seq = 0; };
    window.setTimeout("Scroll2();", interval);
};

Scroll();



window.onscroll = function () {

    var scalar1 = $('scalar');
    var bckgrndImage1 = $('bckgrndImg');
    for (var i = 0; i < scalar1.childNodes.length; i++) {
        scalar1.childNodes[i].style.left = (scArray[i] - getScrollXY().x - ganttZoneleft) + 'px';
        bckgrndImage1.childNodes[i].style.left = (scArray[i] - getScrollXY().x - ganttZoneleft) + 'px';
    }
}

function getScrollXY() {
    var scrOfX = 0;
    var scrOfY = 0;
    if (typeof (window.pageYOffset) == 'number') {
        //Netscape compliant
        //scrOfY = window.pageYOffset;
        scrOfX = window.pageXOffset;
    } else if (document.documentElement && (document.documentElement.scrollLeft || document.documentElement.scrollTop)) {
        //IE6 standards compliant mode
        //scrOfY = document.documentElement.scrollTop;
        scrOfX = document.documentElement.scrollLeft;
    } else if (document.body && (document.body.scrollLeft || document.body.scrollTop)) {
       //DOM compliant
        //scrOfY = document.body.scrollTop;
        scrOfX = document.body.scrollLeft;
    }

    return {x : scrOfX, y : scrOfY};
}




window.onload = function () {

    

    InteractionWithDatabase1('lines', 'action=productionlines');

    //$('scalar').style.cssText = "border: 0px solid black;z-index:1;position:fixed;opacity:1;background-color:white;top:3em;"; //width:" + (90 * 90 + 2) + "px;"
    UsingASPdotNet('scalar', 'action=scalar');
    $('bckgrndImg').style.cssText = "border: 0px solid black;z-index:-1;position:fixed;background-color:transparent;top:4em;"; //width:" + (90 * 90 + 2) + "px;"
    UsingASPdotNet('bckgrndImg', 'action=bckgrndImg');

    var ganttChartArea = document.createElement("div");
    ganttChartArea.className = "DragContainer";
    ganttChartArea.style.cssText = "border: 0px solid black;margin:0px 0px;left:0px;top:4.9em;position:absolute;width:auto;";  // + (90 * 90) + "px;"
    ganttChartArea.id = "ganttZone";
    ganttZoneleft = parseInt(ganttChartArea.style.left);

    document.body.appendChild(ganttChartArea);

    //Fill container area with first available production line
    document.getElementsByTagName('td')[0].onclick();
    //ganttChartArea.getElementsByTagName('td')[0].onmousemove();


    //make both scalar and fixedscalar's left equal to element ganttChartArea's left
    $('scalar').style.left = ganttChartArea.style.left;

    var scalar1 = $('scalar');
    for (var i = 0; i < scalar1.childNodes.length; i++) {
        scArray[i] = parseInt(scalar1.childNodes[i].style.left);
    };

    //get the effect of scrolling
    msg = returnValueToVariable('action=fileTime');
    //Scroll();
};


function InteractionWithDatabase1(elementID1, actioncode) {

    if (navigator.userAgent.toLowerCase().indexOf('chrome') > -1) { //if this is a chrome explorer
        actioncode += "&navigator=chrome";
    };

    var xmlhttp;
    if (elementID1 == "") {
       return;

    }
    if (window.XMLHttpRequest) {
        xmlhttp = new XMLHttpRequest();
    }
    else {
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    xmlhttp.onreadystatechange = function () {
        if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {

            $(elementID1).innerHTML = xmlhttp.responseText;
        }
    }
    xmlhttp.open("POST", "DtExchgPrdctn.aspx", false);
    xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    xmlhttp.send(actioncode);
    //xmlhttp.send(); 

}


function UsingASPdotNet(elementID1, actioncode) {

    var xmlhttp;
    if (elementID1 == "") {
        return;

    }
    if (window.XMLHttpRequest) {
        xmlhttp = new XMLHttpRequest();
    }
    else {
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
    xmlhttp.onreadystatechange = function () {
        if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {

            $(elementID1).innerHTML = xmlhttp.responseText;
        }
    }
    xmlhttp.open("POST", "DtExchgPrdctn.aspx", false);
    xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    xmlhttp.send(actioncode);
    //xmlhttp.send(); 

};


function returnValueToVariable(actioncode) {

    var xmlhttp;
  
    if (window.XMLHttpRequest) {
        xmlhttp = new XMLHttpRequest();
    }
    else {
        xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
    }
 
    xmlhttp.open("POST", "DtExchgPrdctn.aspx", false);
    xmlhttp.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
    xmlhttp.send(actioncode);
    return xmlhttp.responseText;
    //xmlhttp.send(); 

};



function dsplyOrderByLine(theobject) {

    if ($('ganttZone').getAttribute('crntLine') != theobject.getAttribute('id').toString()) {

        $('ganttZone').setAttribute('crntLine', theobject.getAttribute("id").toString());
        InteractionWithDatabase1('ganttZone', 'action=orderlines&lineno=' + theobject.getAttribute("id").toString() + '');
        theobject.style.backgroundColor = "white";
        window.scrollBy(0, -1);
        window.scrollTo(0, 0);
        //returnValueToVariable(msg, 'action=fileTime');
        //Scroll();
    }

    var cells = theobject.parentNode.getElementsByTagName("td");

    for (var i = 0; i < cells.length; i++) {
        if (cells[i] !== theobject) {
            cells[i].style.backgroundColor = "#9999CC"; //back to original color if this is not the current object
        };
    }


};

function $(id) { return document.getElementById(id); };

