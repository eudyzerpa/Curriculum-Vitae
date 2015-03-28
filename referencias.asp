<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Control Panel</title>
</head>

<body>

<center><b>
<font color="#652200" size="2" face="Verdana, Arial, Helvetica, sans-serif">
Referencias Personales</font><font size="2"></strong></font></b></center><br>
</body>

</html>



<%
        
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
               
     
        sql = " SELECT * " & _
              " FROM Referencia " & _
              " WHERE Nombre <> '' "

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
		
</strong><br>
  <center><table width="60%" border="0" height="18">
    <tr> 
      <td bgcolor ="#652200" width="19%" height="14"><strong><font color="#ffffff" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		<% =rs.Fields("Nombre")%></font></strong></td>
		<td bgcolor = #FFFF99 width="16%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("Cargo")%></font></td>
	  <td bgcolor = #FFFF99 width="16%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("Telefono1")%>
      </font></td>
     <td bgcolor = #FFFF99 width="17%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("Telefono2")%>
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