<%
       
       xemail =request.querystring("email")
       
       
      



 
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("galeno.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        sql = " SELECT * " & _
              " FROM registro" & _
              " WHERE eMail = '" & xemail & "'"
			  
           '			  response.Write "SQL--------->" & sql & "<---------" 
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

         if Not rs.EOF Then 
         xnombre = rs.Fields("Nombre")
		 xapellido = rs.Fields("Apellido")
		 response.Write("<b><font face='Courier New' size='4' Color='#ffffff'>Actualizando " & xNombre & " " & xapellido & " !!</font></b><br>")
		 
		 	xespecialidad = rs.Fields("Especialidad")
       		xsubespecialidad = rs.Fields("subespecialidad")
      		xcelular = rs.Fields("Celular")
      		xhospitaltrabaja = rs.Fields("hospitaltrabaja")
      		xdondereside = rs.Fields("dondereside")
      		xconsultorio = rs.Fields("Consultorio")
            xclinicatrabaja = rs.Fields("ClinicaTrabaja")
            xhospitaltrabaja = rs.Fields("HospitalTrabaja")
            xcargo = rs.Fields("cargo")
            xtelefono = rs.Fields("telefono")
            xfax = rs.Fields("fax")
            xtemasinteres = rs.Fields("TemasInteres")








		 session("xemail") = xemail
 
		 End if
					    
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<title>Galeno especialidades médicas c.a</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.style1 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 12px}
.style4 {
	font-size: 14px;
	font-weight: bold;
}
.style6 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; }
.style8 {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 14px;
}
a {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	color: #FFFFFF;
}
a:hover {
	color: #333333;
}
-->
</style>
</head>

<body bgcolor="#339966" leftmargin="50" topmargin="10" rightmargin="50">
<div id="Layer3" style="position:absolute; left:99px; top:11px; width:256px; height:97px; z-index:3"><img src="img/logotransparent.gif" width="359" height="133"></div>
<table width="82%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td width="45%" height="137" valign="top" bgcolor="#FFFFFF">
      <blockquote>
        <p class="style1 style4">&nbsp;</p>
    </blockquote></td>
    <td width="55%" valign="top" bgcolor="#FFFFFF"><img src="img/top2.jpg" width="429" height="137"></td>
  </tr>
  <tr align="center">
    <td height="59" colspan="2" valign="top" bgcolor="#FFFFFF"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="764" height="58">
      <param name="movie" value="menugaleno.swf">
      <param name="quality" value="high">
      <embed src="menugaleno.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="764" height="58"></embed>
    </object></td>
  </tr>
  <tr>
    <td height="644" colspan="2" valign="top" bgcolor="#FFFFFF">
	<table background="img/Falogo03.gif" width="99" height="631"  border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#CCCCCC">
      <tr>
        <td height="629" valign="baseline"><p><img src="img/banner-guia.gif" width="750" height="63"></p>
            <FORM METHOD="Post" name="Login" ACTION="update.asp">
    
  <div align="left">    
	<TABLE border=0 id=TABLE1 width="427">
      <TBODY>
        <TR>
          <TD colSpan=2><div align="left"><strong>
			<font size="2" face="Courier New">Especialidad:</font></strong><BR>
              <INPUT size=27 
                    name=txt_Especialidad VALUE="<%=xEspecialidad%>"></>
          </div></TD>
          <TD><strong><font size="2" face="Courier New">Sub-especialidad:</font></strong><BR>
              <INPUT size=27 
                    name=txt_Subespecialidad value="<%=xsubespecialidad%>"></TD>
        </TR>
        <TR>
          <TD colSpan=2><div align="left"><strong>
			<font size="2" face="Courier New">Celular</font></strong><BR>
              <INPUT size=27 
                    name=txt_Celular VALUE="<%=xcelular%>">
          </div></TD>
          <TD><strong><font face="Courier New" size="2">Donde Reside:</font></strong><br>
              <INPUT size=45 name=txt_dondereside value="<%=xdondereside%>"></TD>
        <TR>
          <TD colSpan=2><div align="left"><strong>
			<font size="2" face="Courier New">Fecha de Cumplea&ntilde;os:</font></strong><BR>
              <INPUT size=27 
                    name=txt_FechaCumpleanos Value="<%=FormatDateTime(Date,2)%>">
          <A href="javascript:cal1.popup();"><IMG height=16 alt="Click aqui para seleccionar una Fecha..." src  ="img/cal.gif" width=16 border=0 ></A> </div></TD>
        </TR>
        <TR>
          <TD width=90><div align="left"></div></TD>
          <TD width=161></TD>
          <TD width=162><BR></TD>
        </TR>
        <TR>
          <TD colSpan=2><div align="left"><strong>
			<font size="2" face="Courier New">Hospital donde trabaja</font></strong><BR>
              <INPUT size=27 name=txt_HospitalTrabaja VALUE="<%=xhospitaltrabaja%>">
          </div></TD>
          <TD><strong><font size="2" face="Courier New">Cargo</font></strong><BR>
              <INPUT size=27 name=txt_Cargo VALUE="<%=xcargo%>"></TD>
        </TR>
        <TR>
          <TD colSpan=2><div align="left"><strong>
			<font size="2" face="Courier New">Clinica donde Trabaja</font></strong><BR>
              <INPUT size=27 name=txt_ClinicaTrabaja VALUE="<%=xclinicatrabaja%>">
          </div></TD>
          <TD><strong><font size="2" face="Courier New">Consultorio</font></strong><BR>
              <INPUT size=27 name=txt_Consultorio VALUE="<%=xconsultorio%>"></TD>
        </TR>
        <tr>
          <TD colSpan=2><div align="left"><strong>
			<font size="2" face="Courier New">Tel&eacute;fono</font></strong><BR>
              <INPUT size=27 name=txt_telefono value="<%=xtelefono%>">
          </div></TD>
          <TD><strong><font size="2" face="Courier New">Fax</font></strong><BR>
              <INPUT size=27 name=txt_Fax value="<%=xfax%>"</TD>
        </tr>
        <TR>
          <TD colSpan=3 height="51"><div align="left"><strong>
			<font size="2" face="Courier New">Temas de Inter&eacute;s:</font></strong><BR>
              <TEXTAREA name=txt_TemasInteres VALUE="<%=xtemasinteres%> cols=50 rows="2" rows="2" cols="50"></TEXTAREA>
          </div></TD>
        </TR>
        <TR><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp;</font> </strong></TR>
        <TR>
          <TD colSpan=2>
            <p align="left"><strong><font size="2" face="Courier New">&iquest;Desea&nbsp;recibir informaci&oacute;n sobre&nbsp;nuestros Productos?&nbsp;</font></strong><B><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1>&nbsp;<BR>
          	</font></b><strong><font size="2" face="Courier New">Si</font></strong><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>
          <INPUT type=radio CHECKED value=SI name=txt_SINO>
          </b></font><strong><font size="2" face="Courier New">No </font>
			</strong><FONT 
                    face="Verdana, Arial, Helvetica, sans-serif" color=#996600 
                    size=1><B>
          <INPUT type=radio value=NO name=txt_SINO>
&nbsp;</b></font></p>            <p align="left"></P>            <P align=left><b>&nbsp;&nbsp;
    <input type="submit" name="Submit" value="Actualizar">
</b></P></TD>
        </TR>
      </TBODY>
    </TABLE>
    <p></P>
    <p></p>
  </div>
    
</FORM>
            <div align="center"></div>
            <div align="center"></div>
            <div align="center"></div>
            <p>&nbsp;</p>
            
          </table></td>
		                 <p align="center"><font color="#000080">
				              
                  </font></p>
                <p></P>
                <p></p>
  </tr></td>
  </tr>

  <tr align="center" valign="baseline">
    <td height="24" colspan="2"><p class="style8"><font color="#FFFFFF">www.galeno.com.ve 
        E-mail:<u> galeno@galeno.com.ve</u></font></p>    </td>
  </tr>
</table>
</body>
</html>
