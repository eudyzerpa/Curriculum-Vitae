<%

  Dim xapellidos
   
  	   xapellidos= Request.QueryString("idccontacto")
 
       if Request.QueryString("borrar") = "true" then
      
       
 
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        SQL = "DELETE FROM telefonos WHERE apellidos = '" & xapellidos & "'" 
			  
        cn.Execute SQL, eliminados   			
        
        if eliminados > 0 then
              %>
                  <script language="JavaScript">
                      alert ('El registro ha sido removido exitosamente');
                      window.location.href='listausuarios.asp';
                  </script>
                         
              <%
          else
              %>
                  <script language="JavaScript">
                      alert ('El registro no pudo ser removido');
                      

                  </script>
              <%          
           end if
                 
        end if
        
                     
			
        
				
		

%>