#####
#
#   Mail-Game
# ==================================
#
# .. Developed by André Peil
#
#####
require 'spreadsheet'
require 'i18n'
require 'roo'
require 'csv'
require 'valid_email'
require 'date'
require 'resolv'

class Person
    include ActiveModel::Validations
    attr_accessor :nome, :email, :cidade, :uf, :dt_nascimento, :if, :escolaridade

    validates :nome, :presence => false, :length => { :maximum => 100 }
    validates :email, :presence => true, :email => true

end


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

#
# First Step
#   Open file
############################

    data = Roo::Spreadsheet.open('./data/exemplo-curto.xlsx')
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

    p "VALIDOS: #{contValid}"
    p "FALSOS: #{contFalse}"
