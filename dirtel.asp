<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Control Panel</title>
<script language="JavaScript">
<!--
function FP_swapImg() {//v1.0
 var doc=document,args=arguments,elm,n; doc.$imgSwaps=new Array(); for(n=2; n<args.length;
 n+=2) { elm=FP_getObjectByID(args[n]); if(elm) { doc.$imgSwaps[doc.$imgSwaps.length]=elm;
 elm.$src=elm.src; elm.src=args[n+1]; } }
}

function FP_preloadImgs() {//v1.0
 var d=document,a=arguments; if(!d.FP_imgs) d.FP_imgs=new Array();
 for(var i=0; i<a.length; i++) { d.FP_imgs[i]=new Image; d.FP_imgs[i].src=a[i]; }
}

function FP_getObjectByID(id,o) {//v1.0
 var c,el,els,f,m,n; if(!o)o=document; if(o.getElementById) el=o.getElementById(id);
 else if(o.layers) c=o.layers; else if(o.all) el=o.all[id]; if(el) return el;
 if(o.id==id || o.name==id) return o; if(o.childNodes) c=o.childNodes; if(c)
 for(n=0; n<c.length; n++) { el=FP_getObjectByID(id,c[n]); if(el) return el; }
 f=o.forms; if(f) for(n=0; n<f.length; n++) { els=f[n].elements;
 for(m=0; m<els.length; m++){ el=FP_getObjectByID(id,els[n]); if(el) return el; } }
 return null;
}
// -->
</script>
</head>

<body onload="FP_preloadImgs(/*url*/'buttonE.jpg', /*url*/'buttonF.jpg', /*url*/'img/buttonE.jpg', /*url*/'img/buttonF.jpg')">

<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<a href="registrotelefonos.asp">

<img border="0" id="img1" src="img/buttonD.jpg" height="20" width="100" alt="Registrar" onmouseover="FP_swapImg(1,0,/*id*/'img1',/*url*/'img/buttonE.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img1',/*url*/'img/buttonD.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img1',/*url*/'img/buttonF.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img1',/*url*/'img/buttonE.jpg')" fp-style="fp-btn: Jewel 4" fp-title="Registrar"></a>
<p></p>

</body>

</html>



<%
        
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
               
     
        sql = " SELECT TOP 7 * " & _
              " FROM telefonos " & _
              " WHERE Nombre <> '' ORDER BY apellidos  "

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
     
     <td bgcolor = #FFFF99 width="17%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% =rs.Fields("TelefonoDomicilio")%>
      </font></td>
               
     <td bgcolor = #FFFF99 width="4%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="actualizartel.asp?idcontacto=<%=rs("id")%>"><b>
		<font style="text-decoration: yes">Editar</font></b></a>
      </font></td>

     <td bgcolor = #FFFF99 width="6%" height="14"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
       
      <a href="eliminar.asp?idcontacto=<%=rs("id")%>"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" style="text-decoration:yes"><b>Eliminar</b></font></a></font></td>
    </tr>
</font></td>



    </tr>
    </table></center>

<strong>
<br>




</strong>


<% 
   rs.MoveNext 
        wend


end if

		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing
		
%>


<table border="0" width="295" height="39" align=center id="table1">
	<tr>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=a"><b>
		<font style="text-decoration: yes">A</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=b"><b>
		<font style="text-decoration: yes">B</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=c"><b>
		<font style="text-decoration: yes">C</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=d"><b>
		<font style="text-decoration: yes">D</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=e"><b>
		<font style="text-decoration: yes">E</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=f"><b>
		<font style="text-decoration: yes">F</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=g"><b>
		<font style="text-decoration: yes">G</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=h"><b>
		<font style="text-decoration: yes">H</font></b></a>&nbsp;
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=i"><b>
		<font style="text-decoration: yes">I</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=j"><b>
		<font style="text-decoration: yes">J</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=k"><b>
		<font style="text-decoration: yes">K</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=l"><b>
		<font style="text-decoration: yes">L</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=m"><b>
		<font style="text-decoration: yes">M</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=n"><b>
		<font style="text-decoration: yes">N</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=o"><b>
		<font style="text-decoration: yes">O</font></b></a>
      </font></td>
		<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=p"><b>
		<font style="text-decoration: yes">P</font></b></a>
      </font></td>
	<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		<a href="char.asp?char=q"><b>Q</b></a>
      </font></td>
      <td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		<a href="char.asp?char=r"><b>R</b></a>
      </font></td>
	<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		<a href="char.asp?char=s"><b>S</b></a>
      </font></td>
      <td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=t"><b>
		<font style="text-decoration: yes">T</font></b></a>
      </font></td>
<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=u"><b>
		<font style="text-decoration: yes">U</font></b></a>
      </font></td>
<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=v"><b>
		<font style="text-decoration: yes">V</font></b></a>
      </font></td>
      <td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
		<a href="char.asp?char=w"><b>W</b></a>
      </font></td>
<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=x"><b>
		<font style="text-decoration: yes">X</font></b></a>
      </font></td>
<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=y"><b>
		<font style="text-decoration: yes">Y</font></b></a>
      </font></td>
<td height="33" width="12">
		&nbsp;<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
     <a href="char.asp?char=z"><b>
		<font style="text-decoration: yes">Z</font></b></a>
      </font></td>

      



	
	</tr>
</table>