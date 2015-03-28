<% 

   
        openstr = "driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr
        
      sqlId = " SELECT * " & _
              " FROM Telefonos"   
              
     
		              
              
              
      Set rsId = Server.CreateObject("ADODB.Recordset")
      rsId.Open sqlId, cn, 3, 3 
      
      
      if rsId.EOF then
      	
          Response.Write "Ta Listo !!!"       
      else
          rsId.MoveFirst
          while not rsId.EOF
        
      xentidad = rsId.fields("Entidad")
      SiguienteEntidad = xentidad + 1
      xcifrado = xentidad*(xFecha / SiguienteEntidad)
      xid = xcifrado       
        
                         
         		      
		    sql = ""
			Sql  = "Update Telefonos Set ID = '" & xid & "' Where Entidad = '" & xentidad & "'"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
						          
	
			cn.execute Sql,	raffected
			
		rsid.MoveNext 
          
        wend
			 
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
         


		End if
%>