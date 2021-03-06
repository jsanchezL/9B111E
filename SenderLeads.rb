#!/usr/bin/env ruby

require 'roo'
require 'rest-client'
require 'json'
require 'mongo'
require 'digest'

class SenderLeads
    @conf = nil
    @@enviadosSuccess = 0
    @@enviadosError = 0
    @filename = nil

    def initialize
        @filename = ARGV[0] # Cuando algo falla y debemos retomar desde el punto de quiebre
        @rowRestart = ARGV[1] # Cuando algo falla y debemos retomar
        tiempoDeInicio = Time.new.to_i
        loadConfig()
        saveBrokers()
        readFileLeads()
        puts ""
        puts "=====> Tiempo realizado <====="
        tiempoDeFin = Time.new.to_i
        tiempoDeProceso = ( tiempoDeFin - tiempoDeInicio ) / 60
        puts ""
        puts "Finalizado en #{tiempoDeProceso} mins."
        puts ""
        puts "Enviados con Éxito: #{@@nviadosSuccess}"
        puts "Enviados con Error: #{@@enviadosSuccess}"
    end

    # Guardamos los Brokers en un archivo para consultarlo
    def saveBrokers
        offset = 0
        brokers = []

        puts ""
        puts "=====> Recuperando Brokers <====="
        puts ""

        loop do
            puts "Recuperando los brokers con offset #{offset}"
            response = retriveBrokers(offset)
            if response.code == 200
                body = JSON.parse(response.body)
                if !body['result']['userProfile'].empty?
                    offset += body['result']['userProfile'].count
                    body['result']['userProfile'].each do |b|
                        brokers.push({
                            "_id" => b['id'].to_i,
                            "displayName" => b['displayName'],
                            "emailAddress" => b['emailAddress']
                        })
                    end
                    puts "Hay actualmente: "+brokers.count.to_s+" brokers"
                else
                    break
                end
            end
        end

        saveBrokers2DB(brokers)
    end

    # Leer el archivo de los datos
    def readFileLeads
        puts ""
        puts "=====> Empezamos a Leer el Archivo de los Leads <====="

        if @filename.nil?
            path = './'.concat(@conf['doc']['filename'])
        else
            path = './'.concat(@filename)
        end
        
        xlsx = Roo::Spreadsheet.open(path, extension: :xlsx)

        if xlsx.sheets.include? @conf['doc']['sheet']
            sheet = xlsx.sheet(@conf['doc']['sheet'])
            sheet.parse(clean: true)
            ultimaRow = sheet.last_row+1
            ultimaCol = sheet.last_column+1
            col = 1 
            row = 2

            campos = []
            # Extraemos los nombres de los campos 
            while col < ultimaCol do
                campos.push(sheet.cell(row,col))
                col += 1
            end

            # Podrian estar en la config con los valores estáticos
            campos.push("campaignID")
            campos.push("leadStatus")

            if @rowRestart.nil?
                row = 3 # Esperamos que sea la primera linea de datos
            else
                row = @rowRestart.to_i # Cuando se interrumpe por algún motivo
            end

            rowData = 1 # Controlar la cantidad de registros por llamada
            max2Send = 250 # Maxima cantidad por llamada al API
            objects = [] # 
            
            # Recorremos los datos del archivo
            while row < ultimaRow do
                if row % 1000 == 0
                    avance = ((row * 100) / ultimaRow)
                    puts "==> Llevamos procesado el #{avance}% de los registros <=="
                end

                lead = {}
                col = 1
                excluido = false

                # Recorremos sobre la linea cada celda
                while col < ultimaCol do
                    i = col-1
                    # Valor de la celda
                    val = sheet.cell(row,col)
                    
                    # Si esta vacía la celda entonces lo omitimos
                    if !val.to_s.empty?
                        case col
                            when 1 # Responsable
                                owner = getBrokerIdByName(val)
                                # Sino existe el owner
                                if owner.nil?
                                    excluido = true
                                else
                                    val = owner['_id'].to_s
                                end
                            when 7 # Télefono
                                val = val.to_s
                            when 15 # Fecha de Creación Historica
                                dt = val.to_datetime
                                val = dt.strftime("%Y-%m-%d %H:%M:%S")
                            when 18 # Observaciones
                                val = val.to_s
                                val = val.gsub("<br>","<br/>")
                                val = val.gsub('"',"'")
                        end
                        lead[campos[i]] = val
                    end
                    col += 1
                end

                lead["campaignID"] = "856462338" #Siempre nacen con campaña historica de BosqueReal
                lead["leadStatus"] = "contact" #Siempre nacen como contactos
                lead["numRow"] = row

                # Agregamos al array de objetos para mandar por el API
                if !excluido
                    objects.push(lead)
                    rowData += 1
                else
                    # Salvamos en BD las lineas que no tiene owner y que no se mandaron al API
                    saveLeadWhitoutOwner(lead,"LeadWhitoutOwner")
                end
                
                # Verificar que podamos enviar
                if rowData % max2Send == 0
                    sendLeads(objects)
                    objects = []
                    rowData = 0 
                else
                    if !(row < ultimaRow-1)
                        sendLeads(objects)
                    end    
                end 
                row += 1
                
                if !excluido
                    rowData += 1
                end
            end
        end
    end

    # Enviamos al API los datos
    def sendLeads (objects)
        cp_objects = objects
        leads = []

        objects.each do |o|
            lead = {}
            o.each do |p,k|
                if p != "numRow"
                    lead[p] = k
                end
            end
            leads.push(lead)
        end

        params = {
            :objects => leads
        }

        loop do
            response = RestClient::Request.execute(
                method: :post, 
                url: buildUrl(),
                payload: buildPayload("createLeads",params),
                headers: buildHeader(),
                read_timeout: 300, 
                open_timeout: 360)
            reintentar = handleResponse(response,cp_objects)
            
            if !reintentar
                break
            end
        end
        # response = MockResponse.new
        # response.post(leads)
        # handleResponse(response,cp_objects)
    end

    # Manejando el Response del API
    def handleResponse (response, objects)
        reintentar = false

        if response.code == 200
            body = JSON.parse(response.body)
            if !body['result']['creates'].empty?
                index = 0
                body['result']['creates'].each do |r|
                    if !r["success"]
                        # Guardamos los errores despues del envío 
                        @@enviadosError += 1
                        saveErrorAfterSend({
                            "numRow": objects[index]['numRow'],
                            "msg": r["error"]
                        },"ErrorAfterSend")
                    else
                        # Guradamos los envios exitosos
                        @@enviadosSuccess += 1
                        saveSuccessAfterSend({
                            "numRow": objects[index]['numRow'],
                            "msg": "Creado con el id: #{r["id"]}"
                        },"SuccessAfterSend")
                    end
                    index+=1
                end
            else
                puts "Algo anda mal con SHSP API..."
            end
        else
            puts "Desconectado y reintentando"
            reintentar = true
            # case response.code
            #     when 101
            #         puts "Invalid request data format"
            #         # Guardar el bloque sospechoso
            #     when 106
            #         puts "Exceeded daily method call limit."
            #         # Reintentar hasta que permita nuevamente
            #     when 107
            #         puts "Exceeded per second method call limit."
            #         # Reintentar hasta que permita nuevamente
            #     when 999
            #         puts "Unknown error. Please contact SharpSpring Developer Support."
            #         # Valor para abortar el envio y registrar donde se quedo
            # end
        end
        return reintentar
    end

    # Guardamos en Base de Datos los Brokers 
    def saveBrokers2DB (brokers)
        client = connectDB() 
        collection = client[:brokers]
        collection.drop()
        begin
            result = collection.insert_many(brokers)
            client.close
            puts "Bokers guardados: #{result.inserted_count}"
        rescue Mongo::Error::NoServerAvailable => e
            puts "Cannot connect to the server"
            puts e
        end
    end

    # Obtener el id del broker por su nombre
    def getBrokerIdByName (name)
        client = connectDB()   
        collection = client[:brokers]
        begin
            result = collection.find({displayName:name}).first
            client.close
        rescue Mongo::Error::NoServerAvailable => e
            puts "Cannot connect to the server"
            puts e
        end
        return result
    end

    # Guardamos todas lineas que no tengan un owner correcto en el sistema
    def saveLeadWhitoutOwner (lead, leadWhitoutOwner)
        client = connectDB()
        collection = client[leadWhitoutOwner]
        numRow = lead["numRow"]
        lead.delete("numRow")
        begin
            result = collection.insert_one({
                "numRow" => numRow, 
                "data" => lead
            })
            client.close
            puts "El Lead sin Owner se a guardado en bitacora: #{result.n}"
        rescue Mongo::Error::NoServerAvailable => e
            puts "Cannot connect to the server"
            puts e
        end
    end

    # Guardamos todas las lineas que no tuvieron exito en su envio
    def saveErrorAfterSend (error, errorAfterSend)
        client = connectDB()
        collection = client[errorAfterSend]
        begin
            result = collection.insert_one(error)
            client.close
        puts "Se guarda el error despues del envio: #{result.n}"
        rescue Mongo::Error::NoServerAvailable => e
            puts "Cannot connect to the server"
            puts e
        end
    end

    def saveSuccessAfterSend (success, successAfterSend)
        client = connectDB()
        collection = client[successAfterSend]
        begin
            result = collection.insert_one(success)
            client.close
        rescue Mongo::Error::NoServerAvailable => e
            puts "Cannot connect to the server"
            puts e
        end
    end

    # Cargamos los params para la conexión con el API
    def loadConfig
        puts "=====> Cargando la config <====="
        file = File.read('./config/conf.json')
        @conf = JSON.parse(file)
    end

    # Conectamos con la BD
    def connectDB
        host = @conf['db']['host'].concat(":").concat(@conf['db']['port'])
        client = Mongo::Client.new([host], :database => @conf['db']['database'])
        return client
    end

    # Recuperamos los Brokers que existen en SHSP
    def retriveBrokers (offset)
        params = {
            :limit => 500,
            :offset => offset,
            :where => {}   
        }
        response = RestClient.post(
            buildUrl(),
            buildPayload("getUserProfiles",params),
            buildHeader()
        )
        return response
    end

    # Arma y regresa el url para consumir el API
    def buildUrl
        url = @conf['isSecure'] ? 'https://' : 'http://'
        url.concat(@conf['domain']).concat('/pubapi/').concat(@conf['versionAPI'])
        url.concat('/?')
        url.concat('accountID=').concat(@conf['credentials']['accountID'])
        url.concat('&')
        url.concat('secretKey=').concat(@conf['credentials']['secretKey'])
        return url
    end

    # Arma el payload para enviar al API
    def buildPayload (method, params)
        return {
            :method => method,
            :params => params,
            :id => @conf['session_id']
        }.to_json
    end

    # Regresa el header para las consultas al API
    def buildHeader
        return {
            :content_type => "application/json; charset=utf-8", 
            :accept => "json",
            :accountID => @conf['credentials']['accountID'],
            :secretKey => @conf['credentials']['secretKey'] 
        }
    end
end

sender = SenderLeads.new