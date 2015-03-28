<HTML>
<HEAD>
<TITLE>Cursos de Especialización</TITLE>
<link href="hojaestilo.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #652200}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</HEAD>
<BODY Background="back2.jpg" vlink="black" link="black">
<p align="center">
<strong>
 <font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#652200">
	Experiencia</font></strong></b></center>
<%

        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

              
     
        sql = " SELECT * " & _
              " FROM experiencia " & _
              " WHERE Empresa <> '' ORDER BY FechaInicio Desc "

        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.EOF then
          Response.Write "No hay datos personales registrados"       
      else
          rs.MoveFirst
          while not rs.EOF 
%>
		
</strong> 
  </p>
<center><center><b>
	<table width="60%" border="0">
    <tr> 
      <td bgcolor ="#652200" width="23%" height="15"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Empresa</font></strong></td>
      <td bgcolor = #FFFF99 width="75%" height="15"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("Empresa") %>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200" width="23%" height="15"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Cargo</font></strong></td>
      <td bgcolor = #FFFF99 width="75%" height="15"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("Cargo")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200" height="14"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Adscrito </font></strong></td>
      <td bgcolor =#FFFF99 height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("Adscrito")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200" height="14"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Supervisor Inmediato</font></strong></td>
      <td bgcolor =#FFFF99 height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("SupervisorInmediato")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Fecha de Ingreso</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("FechaInicio")%>
        </font></td>
    </tr>


    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Fecha de Egreso</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("FechaTermino")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Telefono 1</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("Telefono1")%>
        </font></td>
    </tr>

    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Telefono 2</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("Telefono2")%>
        </font></td>
    </tr>
    

    </table></center>
<strong>
<br>
<% 
   rs.MoveNext 
        wend


end if

		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing
		
%>
</strong> 
<p align="center">&nbsp;</p>
<form>
<div align="center">
<p><input type="button" value="Volver" onclick="history.go(-1)"></p>
</div>
</form>

 

</BODY>
</HTML>