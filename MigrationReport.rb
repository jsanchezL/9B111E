#!/usr/bin/env ruby

require 'spreadsheet'
require 'json'
require 'mongo'

=begin
    Clase para migrar el log de la migración
    a SHSP por el API y que se guardo todo en una BD
    Genera un .xls
=end
class MigrationReport
    @filename = nil
    
    def initialize
        @filename = ARGV[0]
        loadConfig()
        buildReport()
    end

    # Cargamos la configuración de un json
    def loadConfig
        file = File.read('./config/conf.json')
        @conf = JSON.parse(file)
    end
    
    #
    # Obtenemos los datos de las colecciones de Mongo DB y 
    # los vaciamos a un .xls
    #
    def buildReport
        new_book = Spreadsheet::Workbook.new
        worksheets = ["SuccessAfterSend", "LeadWhitoutOwner", "ErrorAfterSend"]
        
        worksheet = 0
        worksheets.each do |name|
            new_book.create_worksheet :name => name
            collection = getCollection(name)
            case name
                when 'SuccessAfterSend'
                    new_book.worksheet(worksheet).insert_row(0,[
                        'Numero de Línea original',
                        'Mensaje de SHSP'
                    ])

                    row = 1
                    collection.each do |i|
                        content = [i["numRow"],i["msg"]]
                        new_book.worksheet(worksheet).insert_row(row,content)
                        row += 1  
                    end
                when 'ErrorAfterSend'
                    new_book.worksheet(worksheet).insert_row(0,[
                        'Numero de Línea original',
                        'Código de error SHSP',
                        'Mensaje del Error',
                        'Campos que fallaron',
                        'Mesaje del campo que fallo',
                        'Formato esperado'
                    ])

                    row = 1
                    collection.each do |i|
                        content = [i["numRow"]]
                        d = i["msg"]
                        content.push(d["code"])
                        content.push(d["message"])
                        
                        if !d["data"].empty?
                            content.push(d["data"]["params"][0]["param"])
                            content.push(d["data"]["params"][0]["message"])
                            content.push(d["data"]["params"][0]["validFormat"]["type"])
                        end

                        new_book.worksheet(worksheet).insert_row(row,content)
                        row += 1  
                    end
                when 'LeadWhitoutOwner'
                    new_book.worksheet(worksheet).insert_row(0,[
                        'Numero de Línea original',
                        'Responsable asignado',
                        'Email del cliente'
                    ])
                    row = 1
                    collection.each do |i|
                        content = [i["numRow"],i["data"]["ownerID"],i["data"]["emailAddress"]]
                        new_book.worksheet(worksheet).insert_row(row,content)
                        row += 1  
                    end
            end
            worksheet += 1
        end
        new_book.write(@filename)
    end

    # Obtenemos los items de las colecciones de la BD
    def getCollection(name)
        cliente = connectDB()
        collection = connectDB()[name]
        items = collection.find()
        cliente.close
        return items
    end

    # Conectamos con la BD
    def connectDB
        host = @conf['db']['host'].concat(":").concat(@conf['db']['port'])
        client = Mongo::Client.new([host], :database => @conf['db']['database'])
        return client
    end
end

report = MigrationReport.new