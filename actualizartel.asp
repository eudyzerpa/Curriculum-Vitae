<%
       dim xcontacto
       xcontacto =request.querystring("idcontacto")
       session("Id")= xcontacto       
      
       
       
      



 
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        sql = " SELECT * " & _
              " FROM telefonos " & _
              " WHERE id = '" & xcontacto & "'"
			  
           '			  response.Write "SQL--------->" & sql & "<---------" 
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

         if Not rs.EOF Then 
         xnombre = rs.Fields("Nombre")
		 xapellido = rs.Fields("Apellidos")
		 response.Write("<b><font face='Courier New' size='4' Color='#ffffff'>Actualizando " & xNombre & " " & xapellido & " !!</font></b><br>")
		 
		 	xTelefonoOficina = rs.Fields("TelefonoOficina")
       		xTelefonoMovil = rs.Fields("TelefonoMovil")
      		xTelefonoDomicilio = rs.Fields("TelefonoDomicilio")
      		xCargo = rs.Fields("Cargo")
      		xOrganizacion = rs.Fields("organizacion")
      		xemail = rs.Fields("email")
            

		 
 
		 End if
					    
%>




<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Control Panel</title>
</head>

<body>

<center>
<FORM action="update.asp" method="post" name="frmReg" >
     
  <INPUT type="hidden" value="true" name="Comportamiento">
  <BR>
  <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
              color=#996600 size=-2><B><span lang="es-ve">Directorio</span> Telefonico</B></FONT></P>
<P>
  <TABLE border=0 id=TABLE1 width="427">
    <TBODY>
      <TR>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Nombre:</B></FONT><BR>
            <INPUT name=txt_Nombre size=27 value="<%=xNombre%>"></TD>
        <TD colSpan=2>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Apellidos:</B></FONT><BR>
            <INPUT size=27 
                  name=txt_apellidos value="<%=xapellido%>"></TD>
      </TR>
      <TR>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Cargo:</B></FONT><BR>
            <INPUT size=27 
                  name=txt_Cargo value="<%=xcargo%>"></TD>
        <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Teléfono Trabajo:</B></FONT><BR>
            <INPUT size=27 
                    name=txt_Telefonooficina value="<%=xtelefonooficina%>"></TD>
      </TR>
      <TR>
       <TD colSpan=2><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Teléfono Movil:</B></FONT><BR>
            <INPUT size=27 
                    name=txt_TelefonoMovil value="<%=xtelefonomovil%>"></TD>
       <TD colSpan=2>
            <FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Telefono Domicilio</B></FONT><BR>
            <INPUT size=27 
                    name=txt_TelefonoDomicilio value="<%=xtelefonodomicilio%>"></TD><TR>
          <TD colSpan=2>
			<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Correo Electrónico:</B></FONT><BR>
            <INPUT size=27 
                  name=txt_email value="<%=xemail%>"></TD>
      </TR>
  <TR>      
        <TD width=90></TD>
        <TD width=150></TD>
        <TD width=173><BR></TD></TR>
      <TR>
        <TD colSpan=2>&nbsp;</TD>
        <TD colSpan=2>&nbsp;</TD>
      </TR>
      <TR>
        <TD colSpan=2>&nbsp;</TD>
        <TD colSpan=2>&nbsp;</TD>
      </TR>
       <tr><TD colSpan=2>&nbsp;</TD>
           <TD colSpan=2>&nbsp;</TD>  </tr>
      <TR>
        <TD colSpan=4 height="51"><B>
        <input type="submit" name="Submit" value="Actualizar Registro"></TD>
      </TR>
      <TR><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      	&nbsp;</font> </strong></TR> 
      <TR>
         <TD colSpan=2>
			&nbsp;</TD>
     
      </TR>      
    </TBODY>
  </TABLE></P>
  <p></p>
</FORM></center>
<form>
<div align="center">
<p><input type="button" value="Volver" onclick="history.go(-1)"></p>
</div>
</form>
		
</strong> 
  

<strong>
<br>


