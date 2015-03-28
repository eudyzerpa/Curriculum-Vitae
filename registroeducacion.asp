<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Control Panel</title>
</head>

<body>

<div style="position: absolute; width: 203px; height: 201px; z-index: 2; left: -3px; top: -19px" id="capa3">
	  
    <table border="0" width="99%" id="table3">
		<tr>
			<td>
			<img border="0" src="img/globe2.jpg" width="200" height="200"><div style="position: absolute; width: 810px; height: 199px; z-index: 1; left: 205px; top: -22px" id="capa2">
				<table border="0" width="100%" id="table4" height="225" bgcolor="#006699">
					<tr>
						<td>&nbsp;&nbsp;&nbsp; <font color="#FFFFFF">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						<b><font face="Bauhaus 93" size="4">&nbsp;&nbsp;&nbsp;&nbsp;</font><font face="Impact" size="4">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
						</font><font face="Bauhaus 93" size="5">Directorio 
						Teléfonico</font></b></font><p><b>
						<font face="Bauhaus 93" size="4" color="#FFFFFF">
						&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font><font face="Arial" color="#FFFFFF" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
						Construyendo nuevos caminos .......</font></b></td>
					</tr>
				</table>
			</div>
			</td>
		</tr>
	</table>
	  
    <p>&nbsp;</p>
	  
    </div>

<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>





<% if request.form("Comportamiento") = "true" then

    xCarrera= request.Form("txt_Carrera")   
    xUniversidad= request.Form("txt_Universidad") 
    xPromocion= request.Form("txt_Promocion")
	xStatus= request.Form("txt_Status")

	

        openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

    
      sqlvalida = " SELECT * " & _
              " FROM Educacion" & _
              " WHERE Carrera = '" & xCarrera & "'"

     
      Set rs = Server.CreateObject("ADODB.Recordset")
      rs.Open sqlvalida, cn, 3, 3 
           

      if rs.eof then
                     
         		      
		    sql = ""
			Sql  = "Insert Into Educacion"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
			sql = sql & " ( "
			Sql = Sql & " Carrera,"			
			Sql = Sql & " Universidad,"
			Sql = Sql & " Promocion,"
			Sql = Sql & " Status"
				
			Sql = Sql & " ) "
			Sql = Sql & " Values "
			Sql = Sql & " ("
		
			Sql = Sql & "'" & xCarrera & "',"
		 	Sql = Sql & "'" & xUniversidad & "',"
		 	Sql = Sql & "'" & xpromocion & "',"
			Sql = Sql & "'" & xStatus & "'"
			Sql = Sql & ")"
		
         	
			cn.execute Sql,	raffected
			 
		cn.Close
	        Set cn = Nothing

						         
         if raffected > 0 then
              response.Redirect("mensaje.asp")
         else
           response.Redirect("mensaje0002.asp")

         end if
Else
response.Redirect("mensaje0001.asp")


		End if
End if

%>		

<center>
<FORM action="registroeducacion.asp" method="post" name="frmReg" >
     
  <INPUT type="hidden" value="true" name="Comportamiento">
  <BR>
  <P><FONT face="Verdana, Arial, Helvetica, sans-serif" 
              color=#996600 size=-2><B>Directorio Telefonico</B></FONT></P>
<P>
  <TABLE border=0 id=TABLE1 width="427">
    <TBODY>
      <TR>
        <TD colSpan=2>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Carrera:</B></FONT><BR>
            <INPUT name=txt_Carrera size=34></TD>
        <TD colSpan=2><b>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#996600">
		Institución Universitaria</font></b><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>:</B></FONT><BR>
            <INPUT size=38 
                  name=txt_Universidad></TD>
      </TR>
      <TR>
        <TD colSpan=2><b>
		<font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#996600">
		Promoción</font></b><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>:</B></FONT><BR>
            <INPUT size=27 
                    name=txt_Promocion></TD>
        <TD colSpan=2>
		<FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>Status:</B></FONT><BR>
          <label>
          <Select name=txt_status>
  			<%
  			   			 
  			 openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("cursos.mdb")
        		Set cn = Server.CreateObject("ADODB.Connection")
              cn.Open openstr
        

  			  SQLCombo = "SELECT status FROM statuscurso"
			  Set rsCombo = cn.Execute(SQLCombo)
			  While Not rsCombo.EOF
			  %>
			  <option value="<%=rsCombo("status")%>"><%=rsCombo("status")%></option>
			  <%
			  rsCombo.MoveNext
			  Wend
			  rsCombo.Close
			  %>
	  		</Select></label></TD>
      </TR>
      <TR>
       <TD colSpan=2><BR>
            &nbsp;</TD>
       <TD colSpan=2><br>
            &nbsp;</TD><TR>
          <TD colSpan=2>&nbsp;</TD>
      </TR>
  <TR>      
        <TD width=90></TD>
        <TD width=119></TD>
        <TD width=274><BR></TD></TR>
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
        <input type="submit" name="Submit" value="Registrarse"></TD>
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
  

<div style="position: absolute; width: 201px; height: 147px; z-index: 3; left: 3px; top: 370px" id="capa5">
	<table border="0" width="100%" id="table5" height="140" bordercolordark="#C0C0C0">
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
			</font></strong><strong><font face="Courier" size="1"> 
			Inicio</font></strong></td>
		</tr>
	</table>
</div>

<strong>
<br>

<div style="position: absolute; width: 201px; height: 120px; z-index: 3; left: 2px; top: 203px" id="capa4">
	<table border="1" width="100%" id="table6" height="148" bordercolordark="#C0C0C0" style="border-left-width: 0px; border-right-width: 0px; border-bottom-width: 0px">
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


