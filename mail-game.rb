#####
#
#   Mail-Game
# ==================================
#
# .. Developed by André Peil
#
#####
require 'rubygems'
require 'i18n'
require 'roo'
require 'csv'
require 'valid_email'
require 'date'
require 'resolv'
require 'axlsx'


#
# validate via DNS
#
############################
    def valid_email_host?(email)
        hostname = email[(email =~ /@/)+1..email.length]
        valid = true
        begin
            Resolv::DNS.new.getresource(hostname, Resolv::DNS::Resource::IN::MX)
        rescue Resolv::ResolvError
            valid = false
        end
        return valid
    end




    def init()

        p "======================="
        p "Iniciando o script..."

        #
        # First Step
        # => Open file
        ############################
        data = Roo::Spreadsheet.open('./data/exemplo-curto.xlsx')
        p "Abrindo o arquivo"
        # validation
        # =>  verify if a valid email
        ############################
        p "validando os emails..."
        valid_emails = validate_emails(data)
        p "OK"

        #
        # Second Step
        # => organize for atributes
        # 1. UF
        # 2. DT_NASCIMENTO
        # 3. IF
        # 4. ESCOLARIDADE
        ##########################
        #organize_to_uf (data)
        #organize_to_dt (data)
        #organize_to_if (data)
        #organize_to_escolaridade (data)

        #
        # Third Step
        # => Export to xlsx
        ##########################
        p "exportando arquivo"
        #export_table(valid_emails)
        p "Ok"
        p "FINALIZANDO"

    end


    def validate_emails(data)
        ret = Array.new
        contValid = 0
        contFalse = 0

        data.each_with_pagename do | name, table_mail |
            table_mail.each do | row |
                if !row[1].eql?("to") then
                    person = Person.new
                    person.nome          = row[0].to_s
                    person.email         = row[1].to_s.downcase
                    person.cidade        = row[2].to_s
                    person.uf            = row[3].to_s
                    person.dt_nascimento = row[4].to_s
                    person.if            = row[5].to_s
                    person.escolaridade  = row[6].to_s

                    if person.valid? then
                        if valid_email_host?(person.email) then
                            ret.push(person)
                            contValid += 1
                        else
                            p person.email
                            contFalse += 1
                        end
                    else
                        contFalse += 1
                        #p "FALSE - #{person.email}"
                    end

                end
            end
        end
        p "EMAILS VALIDOS  : #{contValid}"
        p "EMAILS INVALIDOS: #{contFalse}"
        return ret
    end




    def export_table (data)
        #xlsx = Roo::Excelx.new("./test_small.xlsx")


    end



    class Person
        include ActiveModel::Validations
        attr_accessor :nome, :email, :cidade, :uf, :dt_nascimento, :if, :escolaridade

        validates :nome, :presence => false, :length => { :maximum => 100 }
        validates :email, :presence => true, :email => true

    end





#    init()
