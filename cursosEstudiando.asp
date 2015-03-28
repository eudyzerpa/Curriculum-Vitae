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
              " FROM cursos " & _
              " WHERE Status <> 'Finalizado' "

        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof then
		    response.redirect("mensaje000090.asp")
        else 
		 	if rs.fields("Status") <> "Estudiando" Then
			
				response.Redirect("mensaje000092.asp")					              	
		
		End if

		
		%>
<% if rs.EOF then
          Response.Write "No hay clientes registrados en la base de datos"       
      else
          rs.MoveFirst
          while not rs.EOF 
%>
		
</strong> 
  <center><table width="78%" border="0">
    <tr> 
      <td bgcolor ="#652200" width="11%" height="15"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		Curso</font></strong></td>
      <td bgcolor = #FFFF99 width="61%" height="15"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("NombreCurso")%>
        </font></td>
    	<td>&nbsp;</td>
    </tr>
    <tr> 
      <td bgcolor ="#652200" height="15"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		Instituto</font></strong></td>
      <td bgcolor =#FFFF99 height="15"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("Instituto")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200" height="15"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Modalidad</font></strong></td>
      <td bgcolor =#FFFF99 height="15"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("Modalidad")%>
        </font></td>
    </tr>

    <tr> 
      <td bgcolor ="#652200"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		Horas</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("Horas")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		Fecha</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("Fecha")%>
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
		end if
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