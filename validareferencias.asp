<%
if request.form("bandera") = "true" then
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        sql = " SELECT * " & _
              " FROM Usuarios" & _
              " WHERE Clave= '" & request.form("Clave") & "'"
	       
			  
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

        if rs.eof  then
		    response.Redirect("mensaje0001.htm")
        else 
                        
                        session("apellidos")= rs.fields("apellidos")
                        Session("Usuario")= request.form("Usuario") 
                        session("LoggedIn") = 1
			            response.redirect("http://c.1asphost.com/eudyzerpa/cursos/referencias.asp?target=mainFrame") 
	    end if
        
		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing

End if     
%><p>&nbsp;</p>

<HTML><HEAD> 
<TITLE>Validador de Usuario para visualizar las referencias personales</TITLE> 


<SCRIPT LANGUAGE="JavaScript"> 
function Saltar(pal) { 
window.location=pal+".html" 
} 
</SCRIPT>
<style type="text/css">
<!--
.style3 {font-size: 10pt; font-weight: bold; color: #652200; }
.style4 {color: #000000}
.style5 {color: #000000; font-size: 10pt;}
-->
</style>

</HEAD> 
<BODY bgcolor=#ffffff text=#000000 > 
<script language="JavaScript" type="text/JavaScript">
<!--

// please keep these lines on when you copy the source

// made by: Nicolas - http://www.javascript-page.com

var mymessage = " Previendo este tipo de acceso he protegido con una rutina de encriptación el código fuente de esta página para garantizar la confidencialidad de mis referencias, el resto del código de esta página esta totalmente disponible, incluyendo todo el material escrito, graficos, y archivos de audio.                EUDY ALBERTO ZERPA      Derechos Reservados 2005  ";

function rtclickcheck(keyp){

  if (navigator.appName == "Netscape" && keyp.which == 3) {
    alert(mymessage);
    return false;
  }

  if (navigator.appVersion.indexOf("MSIE") != -1 && event.button == 2) {
    alert(mymessage);
    return false;
  }
}

document.onmousedown = rtclickcheck

-->
</script>

<FORM METHOD ="Post" name="validareferencias" ACTION="validareferencias.asp">  
<div align="justify" class="style3 style4">
  <p>&nbsp;</p>
  <p class="style5" align="center"><span style="font-weight: 400">Protegiendo la privacidad de las personas referidas, el acceso a sus n&uacute;meros 
	telefónicos esta restringido
    por una clave , agradezco a  los interesados solicitarla al correo 
	electrónico eudyzerpa@hotmail.com, incluyendo su número telefónico junto con 
	el resto de sus datos y me comunicare con usted a la brevedad posible.
	</span> </p>
</div>
<P><center> 
   <input type="hidden" name="bandera" value="true">
   <INPUT TYPE="password" NAME="Clave" SIZE=20> 
   <INPUT TYPE="submit" NAME="Boton" SIZE=20 VALUE="Acceder"></center> 
</FORM>
<p>
<img border="0" src="slogan.jpg" width="680" height="124"></p>
</BODY> 
</HTML> 