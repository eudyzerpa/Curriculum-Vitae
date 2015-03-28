<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Control Panel</title>
</head>

<body>

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;

<p>&nbsp;</p>


</body>

</html>



<%
        xchar = request.querystring("char")
        
        
        
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
               
     
        sql = " SELECT * FROM Telefonos WHERE apellidos Like '" & xchar & "' & '%'"
        
        

        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        		
     if rs.EOF then
               Response.Write "No hay referencias registradas en la base de datos"       
      else
          rs.MoveFirst
          while not rs.EOF 
%>
		
		
		
</strong> 

  <center><table width="86%" border="0" height="18">
    
    <tr> 
    <td bgcolor ="#ffffff" width="1%" height="14"><strong>
	<font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
	&nbsp;</font></strong></td>
      <td bgcolor ="#652200" width="15%" height="14"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		<% =rs.Fields("Nombre")%></font></strong></td>
		<td bgcolor ="#652200" width="14%" height="14"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		<% =rs.Fields("Apellidos")%></font></strong></td>

      <td bgcolor = #FFFF99 width="18%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("TelefonoMovil")%>
      </font></td>
     
     <td bgcolor = #FFFF99 width="20%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("TelefonoDomicilio")%>
      </font></td>     
          
     <td bgcolor = #FFFF99 width="4%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="actualizartel.asp?idcontacto=<%=rs("id")%>"><b>
		<font style="text-decoration: yes">Editar</font></b></a>
      </font></td>

     <td bgcolor = #FFFF99 width="6%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
       
      <a href="eliminar.asp?borrar=true&idcontacto=<%=rs("id")%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" style="text-decoration:yes"><b>Eliminar</b></font></a></font></td>
    </tr>
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