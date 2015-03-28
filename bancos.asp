<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Control Panel</title>
</head>

<body>

<div style="position: absolute; width: 203px; height: 201px; z-index: 2; left: -3px; top: -19px" id="capa3">
	  
<div style="position: absolute; width: 201px; height: 147px; z-index: 3; left: 3px; top: 370px" id="capa5">
	<table border="0" width="100%" id="table6" height="140" bordercolordark="#C0C0C0">
		<tr>
			<td height="23" bgcolor="#C0C0C0">
			<img border="0" src="img/dbullet.gif" width="10" height="10">
			<strong><font face="Courier" style="font-size: 1pt">Ir a...</font></strong></td>
		</tr>
		<tr>
			<td height="28">
			<img border="0" src="img/volver.bmp" width="25" height="27"><strong><font face="Courier" size="1">Panel 
			de Control</font></strong></td>
		</tr>
		<tr>
			<td><img border="0" src="img/bookmark.png" width="29" height="20"><strong><font face="Courier" size="1">Curriculum</font></strong></td>
		</tr>
		<tr>
			<td>&nbsp;<img border="0" src="img/toplogin.jpg" width="16" height="14"><strong><font face="Courier" size="1"> 
			Inicio</font></strong></td>
		</tr>
	</table>
</div>

    <table border="0" width="99%" id="table3">
		<tr>
			<td>
			<img border="0" src="img/globe2.jpg" width="200" height="200"><div style="position: absolute; width: 810px; height: 199px; z-index: 1; left: 205px; top: -22px" id="capa2">
				<table border="0" width="100%" id="table4" height="225" bgcolor="#006699">
					<tr>
						<td>&nbsp;&nbsp;&nbsp; <font color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<b><font face="Bauhaus 93" size="4">&nbsp;&nbsp;&nbsp;&nbsp;</font><font face="Impact" size="4">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						</font></b></font><b>
						<font face="Bauhaus 93" size="5" color="#FFFFFF">DATOS 
						BANCARIOS</font></b><p><b>
						<font face="Bauhaus 93" size="4" color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><font face="Arial" color="#FFFFFF" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Construyendo 
						nuevos caminos .......</font></b></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
	</table>
	  
    <p>&nbsp;</p>
	  
    </div>

</body>

</html>

<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>

<%
        
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
               
     
        sql = " SELECT * " & _
              " FROM Bancos " & _
              " WHERE NumeroCuenta <> '' "

        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof then
		    response.redirect("mensaje000090.asp")
                     	
		
		End if

		
		%>
<% if rs.EOF then
          Response.Write "No hay referencias registradas en la base de datos"       
      else
          rs.MoveFirst
          while not rs.EOF 
%>
		
</strong> 
  <center><table width="60%" border="0" height="18" id="table5">
    <tr> 
      <td bgcolor ="#652200" width="19%" height="14"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		<% =rs.Fields("Banco")%></font></strong></td>
      <td bgcolor = #FFFF99 width="45%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("NumeroCuenta")%>
      </font></td>
      <td bgcolor = #FFFF99 width="16%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("TipoCuenta")%>
      </font></td>
     <td bgcolor = #FFFF99 width="17%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("Llave")%>
      </font></td>


    </tr>
    </table></center>
<strong>
<br>

<strong>

<div style="position: absolute; width: 201px; height: 120px; z-index: 3; left: 2px; top: 203px" id="capa4">
	<table border="1" width="100%" id="table7" height="148" bordercolordark="#C0C0C0" style="border-left-width: 0px; border-right-width: 0px; border-bottom-width: 0px">
		<tr>
			<td height="23" bgcolor="#C0C0C0" style="border-bottom-style: none; border-bottom-width: medium">
			<font size="1">
			<img border="0" src="img/dbullet.gif" width="10" height="10"> </font> <b>
			<font face="Courier" size="1">AGENDA</font></b></td>
		</tr>
		<tr>
			<td style="border-style: none; border-width: medium">
			<img border="0" src="img/nuevocontacto.gif" width="17" height="17"><font size="1">
			</font><font face="Courier" size="1">&nbsp;<b>Agregar </b></font>&nbsp;</td>
		</tr>
		<tr>
			<td style="border-style: none; border-width: medium">

<strong>
			
			<img border="0" src="img/modificarcontacto.gif" width="17" height="18">
			
			<font face="Courier" size="1"><b>&nbsp;Modificar </b></font></td>
		</tr>
		<tr>
			<td style="border-style: none; border-width: medium">

<strong>
			<img border="0" src="img/eliminarcontacto.gif" width="17" height="17">&nbsp;&nbsp;
			<font face="Courier" size="1"><b>Eliminar</b></font></td>
		</tr>
	</table>
</div>



<% 
   rs.MoveNext 
        wend


end if

		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing
		
%>