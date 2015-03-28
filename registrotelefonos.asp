<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Control Panel</title>
</head>

<body>


<% if request.form("Comportamiento") = "true" then

    xNombre= request.Form("txt_Nombre")   
    xApellidos= request.Form("txt_Apellidos")   
    xCargo= request.Form("txt_Cargo") 
    xtelefonooficina= request.Form("txt_telefonooficina")
	xTelefonoMovil= request.Form("txt_TelefonoMovil")
	xTelefonoDomicilio= request.Form("txt_TelefonoDomicilio")
	xemail= request.Form("txt_email")
	xFecha= (CSTR(Year(date)))
	

        openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
      Sqlentidad = " SELECT Entidad " & _
              " FROM UltimaEntidad" 
              
              
      Set rsentidad = Server.CreateObject("ADODB.Recordset")
      rsentidad.Open sqlentidad, cn, 3, 3 
        
      xentidad = rsentidad.fields("Entidad")
      SiguienteEntidad = xentidad + 1
      xcifrado = xentidad*(xFecha / SiguienteEntidad)
      xid = xcifrado       
        
        

    
      sqlvalida = " SELECT * " & _
              " FROM Telefonos" & _
              " WHERE Nombre = '" & xNombre & "'"

     
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sqlvalida, cn, 3, 3 
      
      
      if rs.eof then
                          
         		      
		    sql = ""
			Sql  = "Insert Into Telefonos"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
			sql = sql & " ( "
			Sql = Sql & " Nombre,"	
			Sql = Sql & " Apellidos,"			
			Sql = Sql & " Cargo,"
			Sql = Sql & " telefonooficina,"
			Sql = Sql & " TelefonoMovil,"
			Sql = Sql & " TelefonoDomicilio,"
			Sql = Sql & " email,"
            Sql = Sql & " id"
				
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
			Sql = Sql & "'" & xNombre & "',"
			Sql = Sql & "'" & xApellidos & "',"
		 	Sql = Sql & "'" & xCargo & "',"
		 	Sql = Sql & "'" & xtelefonooficina & "',"
			Sql = Sql & "'" & xTelefonoMovil & "',"
			Sql = Sql & "'" & xTelefonoDomicilio & "',"
			Sql = Sql & "'" & xemail & "',"
			Sql = Sql & "'" & xid & "'"

			Sql = Sql & ")"
		
         	
			cn.execute Sql,	raffected
			 
		cn.Close
	        Set cn = Nothing

						         
         if raffected > 0 then
             %>
                 <script language="JavaScript">
                      alert ('El registro ha sido agregado exitosamente');
                      window.location.href='dirtel.asp';
                  </script>
         <%
         else
         %>
           <script language="JavaScript">
                      alert ('El registro no pudo agregarse');
                      window.location.href='dirtel.asp';
                  </script>
          <%
         end if
Else
response.Redirect("mensaje0001.asp")


		End if
End if

%>		

<center>
<FORM action="registrotelefonos.asp" method="post" name="frmReg" >
     
  <INPUT type="hidden" value="true" name="Comportamiento">
  <BR>
  <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
              color=#996600 size=-2><B>Directorio Telefonico</B></FONT></P>
<P>
  <TABLE border=0 id=TABLE1 width="427">
    <TBODY>
      <TR>
        <TD><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Nombre:</B></FONT><BR>
            <INPUT name=txt_Nombre size=27></TD>
        <TD>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Apellidos:</B></FONT><BR>
            <INPUT name=txt_Apellidos size=27></TD>
      </TR>
      <TR>
        <TD><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Cargo:</B></FONT><BR>
            <INPUT size=27 
                  name=txt_Cargo></TD>
        <TD><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Teléfono Trabajo:</B></FONT><BR>
            <INPUT size=27 
                    name=txt_telefonooficina></TD>
      </TR>
      <TR>
       <TD><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Teléfono Movil:</B></FONT><BR>
            <INPUT size=27 
                    name=txt_TelefonoMovil></TD>
       <TD><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Telefono Domicilio</B></FONT><BR>
            <INPUT size=27 
                    name=txt_TelefonoDomicilio><br>
            &nbsp;</TD><TR>
          <TD height="47">
			<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Correo Electrónico:</B></FONT><BR>
            <INPUT size=27 
                    name=txt_email></TD>
      </TR>
       <tr><TD>&nbsp;</TD>
           <TD>&nbsp;</TD>  </tr>
      <TR>
        <TD colSpan=2 height="51"><B>
        <input type="submit" name="Submit" value="Registrar"></TD>
      </TR>
      <TR><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      	&nbsp;</font> </strong></TR> 
      <TR>
         <TD>
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
