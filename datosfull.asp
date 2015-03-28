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
<strong>
 
<%

        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

              
     
        sql = " SELECT * " & _
              " FROM DatosPersonales " & _
              " WHERE Nombre <> '' "

        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.EOF then
          Response.Write "No hay datos personales registrados"       
      else
          rs.MoveFirst
          while not rs.EOF 
%>
		
</strong> 
  <center><center><b>
<font color="#652200" size="2" face="Verdana, Arial, Helvetica, sans-serif">
Datos Personales</font><font size="2"></strong></font></b></center>&nbsp;<table width="60%" border="0">
    <tr> 
      <td bgcolor ="#652200" width="22%" height="15"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Nombre</font></strong></td>
      <td bgcolor = #FFFF99 width="76%" height="15"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("Nombre") & " " & rs.Fields("Apellidos") %>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200" width="22%" height="15"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Cedula de Identidad</font></strong></td>
      <td bgcolor = #FFFF99 width="76%" height="15"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("CI")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200" height="14"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Lugar de Nacimiento</font></strong></td>
      <td bgcolor =#FFFF99 height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("LugarNacimiento")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200" height="14"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Fecha de Nacimiento</font></strong></td>
      <td bgcolor =#FFFF99 height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("FechaNacimiento")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Teléfono Habitación</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("TelefonoHabitacion")%>
        </font></td>
    </tr>


    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Teléfono Celular</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("TelefonoCelular")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Dirección Habitación</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("Direccion")%>
        </font></td>
    </tr>

    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Correo Eléctronico</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("eMail1")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Correo Eléctronico</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("eMail2")%>
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