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
<div style="position: absolute; width: 100px; height: 100px; z-index: 1; left: 463px; top: -28px" id="capa1">
&nbsp;<form name="pages"> 
<select name="pg_choice" onChange="javascript: take_there();"> 
<option value="no_page">Cursos Finalizados...</option> 
<option value="cursosInstituto.asp">Cursos por Instituto</option> 
<option value="CursosEstudiando.asp">Cursos en Proceso</option>
<option value="CursosMod.asp">Cursos por Modalidad</option> 
</select> 
</form> 

<script language="javascript"> 

function take_there() 
{ 
   var destination=document.pages.pg_choice.value; 
   var version = navigator.appVersion; 
   // sets variable = browser version 
   if(destination!="no_page") 
   { 
      if (version.indexOf("MSIE") >= -1) 
      // checks to see if using IE 
      { 
         window.location.href=destination; 
      } 
      else 
      { 
         window.open(destination, target="_self"); 
      }    
   } 
} 

</script> 
</div> 
 
<%

        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

              
     
        sql = " SELECT * " & _
              " FROM cursos " & _
              " WHERE Status = 'Finalizado' ORDER BY Fecha Desc"

        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof then
		    response.redirect("mensaje000090.asp")
        else 
		 	if rs.fields("Status") <> "Finalizado" Then
			
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
  <center>&nbsp;<table width="60%" border="0">
    <tr> 
      <td bgcolor ="#652200" width="10%" height="15"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		Curso</font></strong></td>
      <td bgcolor = #FFFF99 width="72%" height="15"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("NombreCurso")%>
        </font></td>
    	<td><a href="certificados.asp?CodCurso=<%=rs.Fields("CodCurso")%>">
		<img border="0" src="vercertificado.gif" width="34" height="39" alt="Ver Certificado de culminación"></a></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200" height="14"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		Instituto</font></strong></td>
      <td bgcolor =#FFFF99 height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%  =rs.Fields("Instituto")%>
        </font></td>
    </tr>
    <tr> 
      <td bgcolor ="#652200"><strong>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">
		Modalidad</font></strong></td>
      <td bgcolor =#FFFF99><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
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