<%

 xfecha = (Cstr(Year(Date)))

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
      xid = (CSTR(xcifrado))
            
        
                   
         		      
		    sqlUpdate = ""
			SqlUpdate  = "Update Telefonos Set ID = '" & xid & "' Where Entidad = '" & xentidad & "'"                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       
						          
	
		cn.execute SqlUpdate,	raffected
			
		rsid.MoveNext 
          
        wend
			 
		cn.Close
	        Set cn = Nothing

   
              
        End if
		                
              

%>