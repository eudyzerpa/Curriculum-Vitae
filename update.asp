<%

      
       xId = session("Id")
       
 
     
 
        openstr = "driver={Microsoft Access Driver (*.mdb)};" & _
        "dbq=" & Server.MapPath("cursos.mdb")
        Set cn = Server.CreateObject("ADODB.Connection")
        cn.Open openstr

        sql = " SELECT * " & _
              " FROM telefonos" & _
              " WHERE id = '" & xId & "'"
			  
           '			  response.Write "SQL--------->" & sql & "<---------" 
        Set rs = Server.CreateObject("ADODB.Recordset")
        rs.Open sql, cn, 3, 3 

         if Not rs.EOF Then 
         xnombre = rs.Fields("Nombre")
		 xapellidos = rs.Fields("Apellidos")
		 response.Write(" El registro de " &  xNombre & " " & xapellidos & " ha sido actualizado  !!! ")
		 
				
	    
		
		 sqlupdate = " UPDATE telefonos " & _
                                " SET Nombre = '" & Request.Form("Txt_Nombre") & "'," & _ 
                      		    " Apellidos = '" & Request.Form("Txt_apellidos")  & "'," & _
                       		    " Cargo = '" & Request.Form("Txt_cargo")  & "'," & _
				    			" Telefonooficina = '" & Request.Form("Txt_telefonooficina")  & "'," & _
				    			" Telefonomovil = '" & Request.Form("Txt_telefonomovil")  & "'," & _
				    			" Telefonodomicilio = '" & Request.Form("Txt_telefonodomicilio")  & "'," & _
				    			" email = '" & Request.Form("Txt_email") & "' " & _ 
                      		    " WHERE ID = '" & xId & "'"
         
         cn.Execute sqlupdate, raffected
         
         if raffected > 0 then
         %>
                 <script language="JavaScript">
                      alert ('El registro ha sido actualizado exitosamente');
                      window.location.href='dirtel.asp';
                  </script>
         <%
         else
         %>
           <script language="JavaScript">
                      alert ('El registro no pudo ser actualizado');
                      window.location.href='dirtel.asp';
                  </script>
          <%
         end if
     
      end if
                     
			
        
		rs.Close
		Set rs = Nothing
		
		cn.Close
		Set cn = Nothing
		
		

%>