#!/usr/bin/env ruby
require 'roo'

puts ""

path = './'+ARGV[0]
xlsx = Roo::Spreadsheet.open(path, extension: :xlsx)

# Layout de Usuarios
if xlsx.sheets.include? 'Usuarios-Sugar'
    sheet = xlsx.sheet('Usuarios-Sugar')
    sheet.parse(clean: true)
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2
    insertSQL = 'REPLACE INTO `users` ('

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            # Extraemos los nombres de los campos de la tabla a la cual insertaremos la info
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")
                
                # Es la última columna?
                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end

                col += 1
            end
        else # Armamos todos los replace con los datos de cada linea
            col = 1
            insertSQL.concat("(")
            
            # Recorremos las columnas para recolectar los datos
            while col < ultimaCol do
                val = sheet.cell(row,col)

                # Evaluamos si esta vacío el campo
                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    # Tratamos de forma especial a algunas columnas
                    case col
                        when 13 # date_entered
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")   
                        when 14 # date_modified
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S") 
                            insertSQL.concat("\'#{val}\'")
                        when 18 # is_admin
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 25 # receive_notifications
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 30 # sugar_login
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 32 # show_on_employees
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 33 # cookie_consent
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        else
                            insertSQL.concat("\'#{val}\'")
                    end
                end
                
                # Preguntamos sino es la última columna
                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end
    
    puts insertSQL
    puts ""
end

exchangeHashUsers = {}

# Layout de Brokers
if xlsx.sheets.include? 'Brokers'
    sheet = xlsx.sheet('Brokers')
    sheet.parse(clean: true)
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    insertSQL = 'REPLACE INTO `accounts` ('

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 13 # date_entered
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")   
                        when 19 # assigned_user_id
                            ausr = exchangeHashUsers[val]
                            if !ausr.nil?
                                insertSQL.concat("\'#{ausr}\'")
                            else
                                insertSQL.concat("\'#{val}\'")
                            end
                        else
                            insertSQL.concat("\'#{val}\'")
                    end 
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""

    # Abrimos la segunda pestaña del excel
    sheet = xlsx.sheet('accounts_cstm')
    sheet.parse(clean: true)
    insertSQL = 'REPLACE INTO `accounts_cstm` ('
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 2
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        else
                            insertSQL.concat("\'#{val}\'")
                    end
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""
end

# Layout de ClientesPotenciales
if xlsx.sheets.include? 'ClientesPotenciales'
    sheet = xlsx.sheet('ClientesPotenciales')
    sheet.parse(clean: true)
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    insertSQL = 'REPLACE INTO `accounts` ('

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 13 # date_entered
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")   
                        when 14 # date_modified
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")
                        when 15 # employees
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 16 # assigned_user_id
                            ausr = exchangeHashUsers[val]
                            if !ausr.nil?
                                insertSQL.concat("\'#{ausr}\'")
                            else
                                insertSQL.concat("\'#{val}\'")
                            end
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end 
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""

    # Abrimos la segunda pestaña del excel
    sheet = xlsx.sheet('accounts_cstm')
    sheet.parse(clean: true)
    insertSQL = 'REPLACE INTO `accounts_cstm` ('
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 5 # fecha_captura_ventas_tot_c
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d")
                            insertSQL.concat("\'#{val}\'")
                        when 8 # fecha_captura_num_empleado_c
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d")
                            insertSQL.concat("\'#{val}\'")
                        when 10 # fecha_captura_valor_activ_c
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d")
                            insertSQL.concat("\'#{val}\'")
                        when 12 # cliente_nuevo_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 19 # valor_activos_c
                            val = val.to_f
                            insertSQL.concat(val.to_s)
                        when 21 # valor_cuenta_c
                            val = val.gsub("$","")
                            val = val.gsub(",","")
                            val = val.to_f
                            insertSQL.concat(val.to_s)
                        when 22 # dias_sin_actividad_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 23 # diasinactividadh_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""
end

# Layout de ClientesProspectos
if xlsx.sheets.include? 'ClientesProspectos'
    sheet = xlsx.sheet('ClientesProspectos')
    sheet.parse(clean: true)
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    insertSQL = 'REPLACE INTO `accounts` ('

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 13 # date_entered
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")   
                        when 14 # date_modified
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")
                        when 15 # employees
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 16 # rating
                            val = val.capitalize
                            insertSQL.concat("\'#{val}\'")
                        when 17 # assigned_user_id
                            ausr = exchangeHashUsers[val]
                            if !ausr.nil?
                                insertSQL.concat("\'#{ausr}\'")
                            else
                                insertSQL.concat("\'#{val}\'")
                            end
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end 
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""

    # Abrimos la segunda pestaña del excel
    sheet = xlsx.sheet('accounts_cstm')
    sheet.parse(clean: true)
    insertSQL = 'REPLACE INTO `accounts_cstm` ('
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 5 # fecha_captura_ventas_tot_c
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d")
                            insertSQL.concat("\'#{val}\'")
                        when 8 # fecha_captura_num_empleado_c
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d")
                            insertSQL.concat("\'#{val}\'")
                        when 10 # fecha_captura_valor_activ_c
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d")
                            insertSQL.concat("\'#{val}\'")
                        when 19 # valor_activos_c
                            val = val.gsub("$","")
                            val = val.gsub(",","")
                            val = val.to_f
                            insertSQL.concat(val.to_s)
                        when 20 # ventas_totales_c
                            val = val.gsub("$","")
                            val = val.gsub(",","")
                            val = val.to_f
                            insertSQL.concat(val.to_s)
                        when 21 # valor_cuenta_c
                            val = val.gsub("$","")
                            val = val.gsub(",","")
                            val = val.to_f
                            insertSQL.concat(val.to_s)
                        when 22 # dias_sin_actividad_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 23 # rfc_c
                            val = val.length > 13 ? "#{val[0...13]}" : val  
                            insertSQL.concat("\'#{val}\'")
                        when 29 # sw_clienteid_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 31 # base_rate
                            val = val.to_f
                            insertSQL.concat(val.to_s)
                        when 32 # cuentaclave_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 33 # diasinactividadh_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""
end

users = {}

# Layout de Reuniones 
if xlsx.sheets.include? 'Reuniones'
    sheet = xlsx.sheet('Reuniones')
    sheet.parse(clean: true)
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    insertSQL = 'REPLACE INTO `meetings` ('

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 6 # date_start
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")   
                        when 7 # date_end
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")
                        when 8 # date_entered
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")
                        when 9 # date_modified
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")
                        when 12 # assigned_user_id
                            val = users[val]
                            insertSQL.concat("\'#{val}\'")
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end 
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""

    # Abrimos la segunda pestaña del excel
    sheet = xlsx.sheet('meetings_cstm')
    sheet.parse(clean: true)
    insertSQL = 'REPLACE INTO `meetings_cstm` ('
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 2 # market_to_put_solution_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 3 # fidex_official_presentation_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 4 # got_needs_solutions_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 5 # meeting_rating_date_c
                            if (val != "00/01/1900")
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d")
                                insertSQL.concat("\'#{val}\'")
                            else
                                insertSQL.concat("NULL")
                            end
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""
end

# Layout de Contactos
if xlsx.sheets.include? 'Contactos'
    sheet = xlsx.sheet('Contactos')
    sheet.parse(clean: true)
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    insertSQL = 'REPLACE INTO `contacts` ('

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 6 # date_entered
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")
                        when 7 # date_modified
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")
                        when 12 # assigned_user_id
                            val = users[val]
                            insertSQL.concat("\'#{val}\'")
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end 
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""

    # Abrimos la segunda pestaña del excel
    sheet = xlsx.sheet('contacts_cstm')
    sheet.parse(clean: true)
    insertSQL = 'REPLACE INTO `contacts_cstm` ('
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 2 # contacto_tomador_decision_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 7 # usuario_c
                            val = users[val]
                            insertSQL.concat("\'#{val}\'")
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""

     # Abrimos la tercera pestaña del excel
     sheet = xlsx.sheet('email_addresses')
     sheet.parse(clean: true)
     insertSQL = 'REPLACE INTO `email_addresses` ('
     insertSQL2 = 'REPLACE INTO `email_addr_bean_rel` (`id`,`email_address_id`,`bean_id`,`bean_module`,`primary_address`,`date_created`,`date_modified`) VALUES '
     ultimaRow = sheet.last_row+1
     ultimaCol = sheet.last_column+1
     col = 1 
     row = 2
 
     # Recorremos de la fila 2 en adelante
     while row < ultimaRow do
        descartarLinea = false
         # Armamos el encabezado del replace con las columnas del sistema
         if row === 2
             while col < ultimaCol do
                 val = sheet.cell(row,col)
                 insertSQL.concat("`").concat(val).concat("`")
 
                 if col != ultimaCol-1
                     insertSQL.concat(",")
                 else
                     insertSQL.concat(") VALUES ")
                 end
                 col += 1
             end
         else
             col = 1
             # Verificamos que exista email para el registro
             email = sheet.cell(row,col+1)
             if email.to_s.empty?
                descartarLinea = true
             end

             if !descartarLinea
                insertSQL.concat("(")
                insertSQL2.concat("(")
                while col < ultimaCol do   
                    val = sheet.cell(row,col)
                    
                    # Armamos los replace para la tabla email_addr_bean_rel 
                    case col
                        when 1 # id_c
                            insertSQL2.concat("\'#{val}\'")
                            insertSQL2.concat(",")
                            insertSQL2.concat("\'#{val}\'")
                            insertSQL2.concat(",")
                            insertSQL2.concat("\'#{val}\'")
                            insertSQL2.concat(",")
                            insertSQL2.concat("'Contacts'")
                            insertSQL2.concat(",")
                            insertSQL2.concat("1")
                            insertSQL2.concat(",")
                        when 4 # date_entered
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL2.concat("\'#{val}\'")
                            insertSQL2.concat(",")
                        when 5 # date_modified
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL2.concat("\'#{val}\'")
                    end

                    if val.to_s.empty?
                        insertSQL.concat("NULL")
                    else
                        case col
                            when 3 # email_address_caps
                                val = val.upcase
                                insertSQL.concat("\'#{val}\'")
                            else
                                val = val.gsub("#","")
                                val = val.gsub("'","")
                                insertSQL.concat("\'#{val}\'")
                        end
                    end

                    if col != ultimaCol-1
                        insertSQL.concat(",")
                    else
                        insertSQL.concat(")")
                        # Cerramos los replace para la tabla email_addr_bean_rel 
                        insertSQL2.concat(")")
                    end
                    
                    col += 1
                end
             end
         end
 
         if !descartarLinea
            # Separación de los replace en caso contrario se termina con ;
            if row < ultimaRow-1
                if row > 2 
                    insertSQL.concat(",\n")
                    insertSQL2.concat(",\n")
                end
            else
                insertSQL.concat(";")
                insertSQL2.concat(";")
            end
         end
         row += 1
     end
 
     puts insertSQL
     puts ""

     puts insertSQL2
     puts ""


    # Abrimos la cuarta pestaña del excel
    sheet = xlsx.sheet('accounts_contacts')
    sheet.parse(clean: true)
    insertSQL = 'REPLACE INTO `accounts_contacts` ('
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        descartarLinea = false
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            account = sheet.cell(row,col+2)
            if account.to_s.empty?
                descartarLinea = true
            end

            if !descartarLinea
                insertSQL.concat("(")
                while col < ultimaCol do
                    val = sheet.cell(row,col)

                    if val.to_s.empty?
                        insertSQL.concat("NULL")
                    else
                        case col
                            when 4 # date_modified
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S")
                                insertSQL.concat("\'#{val}\'")
                            when 5 # contacto_tomador_decision_c
                                val = val.to_i
                                insertSQL.concat(val.to_s)
                            else
                                val = val.gsub("#","")
                                val = val.gsub("'","")
                                insertSQL.concat("\'#{val}\'")
                        end
                    end

                    if col != ultimaCol-1
                        insertSQL.concat(",")
                    else
                        insertSQL.concat(")")
                    end

                    col += 1
                end
            end
        end

        if !descartarLinea
            # Separación de los replace en caso contrario se termina con ;
            if row < ultimaRow-1
                if row > 2
                    insertSQL.concat(",\n")
                end
            else
                insertSQL.concat(";")
            end
        else
            if row == ultimaRow-1
                insertSQL = insertSQL.chomp(",\n")
                insertSQL.concat(";")
            end
        end

        row += 1
    end

    puts insertSQL
    puts ""
end

# Layout de Llamadas
if xlsx.sheets.include? 'Llamadas'
    sheet = xlsx.sheet('Llamadas')
    sheet.parse(clean: true)
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2
    #Partimos en 2 tantos las llamadas ya que superan los 16K de registros
    insertSQL = 'REPLACE INTO `calls` ('
    insertSQL2 = 'REPLACE INTO `calls` ('

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            # Extraemos los nombres de los campos de la tabla a la cual insertaremos la info
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")
                insertSQL2.concat("`").concat(val).concat("`")
                
                # Es la última columna?
                if col != ultimaCol-1
                    insertSQL.concat(",")
                    insertSQL2.concat(",")
                else
                    insertSQL.concat(") \nVALUES ")
                    insertSQL2.concat(") \nVALUES ")
                end

                col += 1
            end
        else # Armamos todos los replace con los datos de cada linea
            col = 1
            if row < 9002
                insertSQL.concat("(")
                # Recorremos las columnas para recolectar los datos
                while col < ultimaCol do
                    val = sheet.cell(row,col)

                    # Evaluamos si esta vacío el campo
                    if val.to_s.empty?
                        insertSQL.concat("NULL")
                    else
                        # Tratamos de forma especial a algunas columnas
                        case col
                            when 6 # date_start
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S")
                                insertSQL.concat("\'#{val}\'") 
                            when 7 # date_end
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S")
                                insertSQL.concat("\'#{val}\'") 
                            when 8 # date_entered
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S")
                                insertSQL.concat("\'#{val}\'")   
                            when 9 # date_modified
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S") 
                                insertSQL.concat("\'#{val}\'")
                            when 10 # duration_hours
                                val = val.to_i
                                insertSQL.concat(val.to_s)
                            when 11 # duration_minutes
                                val = val.to_i
                                insertSQL.concat(val.to_s)
                            when 12 # assigned_user_id
                                val = users[val]
                                insertSQL.concat("\'#{val}\'")
                            else
                                val = val.gsub("#","")
                                val = val.gsub("'","")
                                insertSQL.concat("\'#{val}\'")
                        end
                    end
                    
                    # Preguntamos sino es la última columna
                    if col != ultimaCol-1
                        insertSQL.concat(",")
                    else
                        insertSQL.concat(")")
                    end

                    col += 1
                end
            else
                insertSQL2.concat("(")
                # Recorremos las columnas para recolectar los datos
                while col < ultimaCol do
                    val = sheet.cell(row,col)

                    # Evaluamos si esta vacío el campo
                    if val.to_s.empty?
                        insertSQL2.concat("NULL")
                    else
                        # Tratamos de forma especial a algunas columnas
                        case col
                            when 6 # date_start
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S")
                                insertSQL2.concat("\'#{val}\'") 
                            when 7 # date_end
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S")
                                insertSQL2.concat("\'#{val}\'") 
                            when 8 # date_entered
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S")
                                insertSQL2.concat("\'#{val}\'")   
                            when 9 # date_modified
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S") 
                                insertSQL2.concat("\'#{val}\'")
                            when 10 # duration_hours
                                val = val.to_i
                                insertSQL2.concat(val.to_s)
                            when 11 # duration_minutes
                                val = val.to_i
                                insertSQL2.concat(val.to_s)
                            when 12 # assigned_user_id
                                val = users[val]
                                insertSQL2.concat("\'#{val}\'")
                            else
                                val = val.gsub("#","")
                                val = val.gsub("'","")
                                insertSQL2.concat("\'#{val}\'")
                        end
                    end
                    
                    # Preguntamos sino es la última columna
                    if col != ultimaCol-1
                        insertSQL2.concat(",")
                    else
                        insertSQL2.concat(")")
                    end

                    col += 1
                end
            end
        end

        if row < 9002
            # Separación de los replace en caso contrario se termina con ;
            if row < (9002-1)
                if row > 2 
                    insertSQL.concat(",\n")
                end
            else
                insertSQL.concat(";")
            end
        else
            if row < ultimaRow-1
                if row > 2 
                    insertSQL2.concat(",\n")
                end
            else
                insertSQL2.concat(";")
            end
        end
        row += 1
    end
    
    puts insertSQL
    puts ""

    puts insertSQL2
    puts ""
end


subramos = {}
subramosName = {}
ramos = {}

# Layout de Opps y RLIs
if xlsx.sheets.include? 'Oportunidades'
    sheet = xlsx.sheet('Oportunidades')
    sheet.parse(clean: true)
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    xlsx1 = Roo::Spreadsheet.open(path, extension: :xlsx)
    sheet1 = xlsx1.sheet('RLIs')
    sheet1.parse(clean: true)
    ultimaCol1 = sheet1.last_column+1
    col1 = 1

    insertSQL = 'REPLACE INTO `opportunities` ('
    insertSQL1 = 'REPLACE INTO `revenue_line_items` ('
    insertSQL2 = 'REPLACE INTO `revenue_line_items_cstm` (`id_c`,`porcentaje_comision_c`) VALUES '

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)

                if col == 3
                    insertSQL.concat("`").concat(val).concat("`").concat(',')
                    insertSQL.concat("`").concat("base_rate").concat("`")
                else
                    insertSQL.concat("`").concat(val).concat("`")
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end

            # Para las RLIs
            while col1 < ultimaCol1 do
                val = sheet1.cell(row,col1)

                if col1 == 3
                    insertSQL1.concat("`").concat(val).concat("`").concat(',')
                    insertSQL1.concat("`").concat("base_rate").concat("`")
                else
                    insertSQL1.concat("`").concat(val).concat("`")
                end

                if col1 != ultimaCol1-1
                    insertSQL1.concat(",")
                else
                    insertSQL1.concat(") VALUES ")
                end
                col1 += 1
            end

        else
            # Para Opps
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 2 # name
                            insertSQL.concat("\'#{val}\'")
                        when 3 # currency_id
                            if val == "-99"
                                insertSQL.concat("\'#{val}\'").concat(",")
                                base_rate = 1.0
                                insertSQL.concat(base_rate.to_s)
                            else
                                insertSQL.concat("\'#{val}\'").concat(",")
                                base_rate = 21.0 # Lo sacamos de la config de SugarCE de FIDEX
                                insertSQL.concat(base_rate.to_s)
                            end    
                        when 5 # assigned_user_id
                            val = users[val]
                            insertSQL.concat("\'#{val}\'")
                        when 9 # date_entered
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")
                        when 10 # date_modified
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL.concat("\'#{val}\'")
                        when 11 # best_case
                            val = val.gsub("$","")
                            val = val.gsub(",","")
                            val = val.to_f
                            insertSQL.concat(val.to_s)
                        when 12 # worst_case
                            val = val.gsub("$","")
                            val = val.gsub(",","")
                            val = val.to_f
                            insertSQL.concat(val.to_s)
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end 
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end

            # Para las RLIs
            col1 = 1
            insertSQL1.concat("(")
            insertSQL2.concat("(")
            
            id_c_rli_cstm = ''
            porcentaje_comision_c = 0.0
            best_case = 0.0
            worst_case = 0.0

            while col1 < ultimaCol1 do
                val = sheet1.cell(row,col1)

                if val.to_s.empty?
                    insertSQL1.concat("NULL")
                else
                    case col1
                        when 1 # id
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL1.concat("\'#{val}\'")
                            id_c_rli_cstm = val
                        when 3 # currency_id
                            if val == "-99"
                                insertSQL1.concat("\'#{val}\'").concat(",")
                                base_rate = 1.0
                                insertSQL1.concat(base_rate.to_s)
                            else
                                insertSQL1.concat("\'#{val}\'").concat(",")
                                base_rate = 21.0 # Lo sacamos de la config de SugarCE de FIDEX
                                insertSQL1.concat(base_rate.to_s)
                            end   
                        when 4 # name
                             val = subramosName[val]
                             insertSQL1.concat("\'#{val}\'")
                        when 5 # assigned_user_id
                            val = users[val]
                            insertSQL1.concat("\'#{val}\'")
                        when 6 # date_entered
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL1.concat("\'#{val}\'")
                        when 7 # date_modified
                            dt = val.to_datetime
                            val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            insertSQL1.concat("\'#{val}\'")
                        when 10 # product_template_id
                            val = subramos[val]
                            insertSQL1.concat("\'#{val}\'")
                        when 11 # category_id
                            val = ramos[val]
                            insertSQL1.concat("\'#{val}\'")
                        when 13 # best_case
                            val = val.gsub("$","")
                            val = val.gsub(",","")
                            val = val.to_f
                            best_case = val
                            insertSQL1.concat(val.to_s)
                        when 14 # worst_case
                            val = val.gsub("$","")
                            val = val.gsub(",","")
                            val = val.to_f
                            worst_case = val
                            insertSQL1.concat(val.to_s)
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL1.concat("\'#{val}\'")
                    end 
                end

                if col1 != ultimaCol1-1
                    insertSQL1.concat(",")
                else
                    insertSQL1.concat(")")
                end

                col1 += 1
            end

            if best_case > 0
                porcentaje_comision_c = worst_case / best_case
                if porcentaje_comision_c > 1000
                    porcentaje_comision_c = porcentaje_comision_c / 100000
                end
            end

            # Replace de completo de revenue_line_items_cstm
            insertSQL2.concat("\'#{id_c_rli_cstm}\'").concat(",").concat(porcentaje_comision_c.round(2).to_s).concat(")")
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
                insertSQL1.concat(",\n")
                insertSQL2.concat(",\n")
            end
        else
            insertSQL.concat(";")
            insertSQL1.concat(";")
            insertSQL2.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""

    # Abrimos la segunda pestaña del excel
    sheet = xlsx.sheet('opportunities_cstm')
    sheet.parse(clean: true)
    insertSQL = 'REPLACE INTO `opportunities_cstm` ('
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 2 # mercado_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 5 # valido_hasta_c
                            if (val != "00/01/1900")
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d")
                                insertSQL.concat("\'#{val}\'")
                            else
                                insertSQL.concat("NULL")
                            end
                        when 6 # clienteid_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        when 12 # dias_de_modificacion_c
                            val = val.to_i
                            insertSQL.concat(val.to_s)
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""

    # Abrimos la tercera pestaña del excel
    sheet = xlsx.sheet('accounts_opportunities')
    sheet.parse(clean: true)
    insertSQL = 'REPLACE INTO `accounts_opportunities` ('
    ultimaRow = sheet.last_row+1
    ultimaCol = sheet.last_column+1
    col = 1 
    row = 2

    # Recorremos de la fila 2 en adelante
    while row < ultimaRow do
        # Armamos el encabezado del replace con las columnas del sistema
        if row === 2
            while col < ultimaCol do
                val = sheet.cell(row,col)
                insertSQL.concat("`").concat(val).concat("`")

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(") VALUES ")
                end
                col += 1
            end
        else
            col = 1
            insertSQL.concat("(")

            while col < ultimaCol do
                val = sheet.cell(row,col)

                if val.to_s.empty?
                    insertSQL.concat("NULL")
                else
                    case col
                        when 4 # date_modified
                            if (val != "00/01/1900")
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S")
                                insertSQL.concat("\'#{val}\'")
                            else
                                insertSQL.concat("NULL")
                            end
                        else
                            val = val.gsub("#","")
                            val = val.gsub("'","")
                            insertSQL.concat("\'#{val}\'")
                    end
                end

                if col != ultimaCol-1
                    insertSQL.concat(",")
                else
                    insertSQL.concat(")")
                end

                col += 1
            end
        end

        # Separación de los replace en caso contrario se termina con ;
        if row < ultimaRow-1
            if row > 2 
                insertSQL.concat(",\n")
            end
        else
            insertSQL.concat(";")
        end

        row += 1
    end

    puts insertSQL
    puts ""

    # RLIs
    puts insertSQL1
    puts ""

    # RLIs custom
    puts insertSQL2
    puts ""
end