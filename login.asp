<%

    if request.form("comportamiento") = "true" then
        openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
        sql = " SELECT * " & _
              " FROM Usuarios" & _
              " WHERE Usuario= '" & request.form("Usuario") & _
	      "' AND Clave ='" & request.form("Clave") & "'" 
			  
			  
           '			  response.Write "SQL--------->" & sql & "<---------" 
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof  then
		    response.Redirect("mensaje0002.asp")
        else 
                        
                        Session("Usuario")= request.form("Usuario") 
                      
			     response.redirect("controlpanel.htm") 
	end if
        
		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing

     END IF
%>




<HTML>
<HEAD>
<TITLE>Cursos de Especialización</TITLE>
<link href="hojaestilo.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.style1 {color: #652200}
.topspace {
	MARGIN-TOP: 0.08em
}
.button {
	MARGIN-TOP: 6pt; FONT-WEIGHT: normal; FONT-SIZE: 70%; BORDER-LEFT-COLOR: #6699ff; BORDER-BOTTOM-COLOR: #6699ff; MARGIN-LEFT: 0.5em; COLOR: #000000; BORDER-TOP-COLOR: #6699ff; FONT-FAMILY: Verdana, Helvetica, Arial, San-Serif; BACKGROUND-COLOR: #ffffff; BORDER-RIGHT-COLOR: #6699ff
}
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
 
		
<table border="0" width="450" height="1">
	<tr>
		<td height="1" width="130">
<strong>
 
		<img border="0" src="img/win2000l.gif" width="124" height="123"></strong></td>
		<td height="1" width="310">&nbsp;<p><font face="Arial Black">
		<font color="#652200">
		<font size="4">ADMINISTRADOR 
		DE PORTAL </font><font size="2">Conexión a Panel de control</font></font> </font>
		
		</p>
		
		<p>&nbsp;</td>
	</tr>
</table>
		
</strong> 
  <center>&nbsp;</center>
<TABLE style="MARGIN-TOP: -1em" cellSpacing=0 cellPadding=0 width=640 border=0 id="table1"><!-- Graphic bar row  -->
  <TBODY>
  <TR>
    <TD width="31%" align="left"></TD>
    <TD vAlign=center align=left><font size="1">
	<IMG height=8 
      alt="blue bar graphic" 
      src="bluebarh.gif" 
  width=325></font></TD></TR><!-- Row 1 -->
  <TR><!-- Column 1 spans 4 rows -->
    <TD vAlign=top width="31%" align="left">
      <P class=indent>
		<ID id=remotecomputername><font size="1" face="Arial">El acceso a este 
		portal pronto será habilitado a usuarios registrados. Las&nbsp; 
		descargas de Software Gratuito desde el servidor FTP, Cursos en Linea y 
		otros servicios están actualmente en desarrollo</font></ID><ID id=helpfultip1><font size="1">.</font></ID></P></TD><!-- Column 2 spans 4 rows-->
    <TD vAlign=top align=left><font size="1">
	<IMG height=330 alt="blue bar graphic" 
      src="bluebarv.gif" width=8 
      border=0><div style="position: absolute; width: 257px; height: 270px; z-index: 1; left: 215px; top: 160px" id="capa1">
		<table border="0" width="167%" id="table2" height="328">
			<tr>
				<td valign="top">
				<FORM METHOD="Post" name="Login" ACTION="Login.asp">
				<font color="#652200" size="1" face="Verdana"><b>Usuario </b>
				</font>&nbsp;&nbsp;&nbsp;<INPUT NAME="Usuario" SIZE="30">
					
					<p><b><font face="Verdana" size="1" color="#652200">Password</font></b> <INPUT NAME="Clave" SIZE="30" type=password>
					
					<INPUT NAME="comportamiento" SIZE="30" type="hidden" value="true">
					
					</p>
					
					<p><input type="submit" value="Enviar" name="B1"></p>
				</form>
&nbsp;</td>
			</tr>
		</table>
	</div>
	<div align="center">
	  
    </font></div>
    
 </font> </TD><!-- Column 3 -->
    <!-- Column 4 -->
    </TR><!-- Row 2 -->
  <!-- Row 3 -->
  <!-- Row 4 -->
  	</TBODY></TABLE>
<strong>
<br>
</strong> 

 

</BODY>
</HTML>