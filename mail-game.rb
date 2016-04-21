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
require 'write_xlsx'
require 'csv'
require 'iconv'

#
# export as csv
#
############################
    def export_to_csv (name, data)
        ## log
        # ======
            cont_person = 0
            total = data.length
            #p total
        # ======
        CSV.open("#{name}.csv", "wb",{:col_sep => "\t"}) do |csv|
            csv << ["email"]
            data.each do | person |
                percent = (cont_person*100)/total
                #p "#{percent}% exported csv"
                csv <<  [
                            person['email'].to_s.force_encoding("utf-8"),
                            # person.nome.to_s.force_encoding("utf-8"),
                            # person.cidade.to_s.force_encoding("utf-8"),
                            # person.uf.to_s.force_encoding("utf-8"),
                            # person.dt_nascimento.force_encoding("utf-8"),
                            # person.if.to_s.force_encoding("utf-8"),
                            # person.escolaridade.to_s.force_encoding("utf-8")
                        ]
            end
            cont_person += 1
        end

        p "Exported with name #{name}.csv"
    end

#
# export as xlsx
#
############################
    def export_to_xlsx (name, data)




        # Create a new Excel workbook
        workbook = WriteXLSX.new("#{name}.xlsx")

        # Add a worksheet
        worksheet = workbook.add_worksheet

        # Add and define a format
        format = workbook.add_format # Add a format
        format.set_bg_color('silver')
        format.set_align('center')
        format.set_bold
        format.set_column( 'A:A', 25 )
        format.set_column( 'B:B', 25 )
        format.set_column( 'C:C', 25 )
        format.set_column( 'D:D', 25 )
        format.set_column( 'E:E', 25 )
        format.set_column( 'F:F', 25 )



        format2 = workbook.add_format # Add a format
        format2.set_align('center')
        format2.set_bold


        # Write a formatted and unformatted string, row and column notation.
        row = 0
        worksheet.write(row, 0, "NOME", format)
        worksheet.write(row, 1, "EMAIL", format)
        worksheet.write(row, 2, "CIDADE", format)
        worksheet.write(row, 3, "UF", format)
        worksheet.write(row, 4, "DT_NASCIMENTO", format)
        worksheet.write(row, 5, "IF", format)
        worksheet.write(row, 6, "ESCOLARIDADE", format)
        row += 1

        ## log
        # ======
            cont_person = 0
            total = data.length
        # ======
        data.each do | person |
            percent = (cont_person*100)/total
            #p "#{percent}% xlsx exported"
            cont_person += 1
            if(row % 2 == 0 ) then
                worksheet.write(row, 0, person['nome'].to_s.force_encoding("utf-8"), format)
                worksheet.write(row, 1, person['email'].to_s.force_encoding("utf-8"), format)
                worksheet.write(row, 2, person['cidade'].to_s.force_encoding("utf-8"), format)
                worksheet.write(row, 3, person['uf'].to_s.force_encoding("utf-8"), format)
                worksheet.write(row, 4, person['dt_nascimento'].force_encoding("utf-8"), format)
                worksheet.write(row, 5, person['if'].to_s.force_encoding("utf-8"), format)
                worksheet.write(row, 6, person['escolaridade'].to_s.force_encoding("utf-8"), format)
            else
                worksheet.write(row, 0, person['nome'].to_s.force_encoding("utf-8"), format2)
                worksheet.write(row, 1, person['email'].to_s.force_encoding("utf-8"), format2)
                worksheet.write(row, 2, person['cidade'].to_s.force_encoding("utf-8"), format2)
                worksheet.write(row, 3, person['uf'].to_s.force_encoding("utf-8"), format2)
                worksheet.write(row, 4, person['dt_nascimento'].force_encoding("utf-8"), format2)
                worksheet.write(row, 5, person['if'].to_s.force_encoding("utf-8"), format2)
                worksheet.write(row, 6, person['escolaridade'].to_s.force_encoding("utf-8"), format2)
            end
            row += 1
        end

        workbook.close
        p "Exported with name #{name}.xlsx"
    end

#
# inicializa a os trabalhos
#
############################
    def init()

        p "======================="
        p "Iniciando o script..."

        #
        # First Step
        # => Open file
        ############################
        data = Roo::Spreadsheet.open('./data/exemplo-curto.xlsx')
        p "Abrindo o arquivo"
        persons = organize_data(data)
        # validation
        # =>  verify if a valid email
        ############################
        #p "validando os emails..."
        #emails = validate_emails(data)
        #valid_emails = emails['validos']
        #invalid_emails = emails['invalidos']


        p "OK"
        #export_to_csv("emails_validos", valid_emails)
        #export_to_xlsx("emails_validos", valid_emails)
        #export_to_csv("emails_invalidos", invalid_emails)
        #export_to_xlsx("emails_invalidos", invalid_emails)
        #
        # Second Step
        # => organize for atributes
        # 1. UF
        # 2. DT_NASCIMENTO
        # 3. IF
        # 4. ESCOLARIDADE
        ##########################
        uf_mails = organize_to_uf (persons)
        uf_mails.each do | key, est |
            if est.length == 0 then
                p "vazio #{key} "
            else
                export_to_csv(key, est)
                export_to_xlsx(key, est)
            end

        end

        dt_nascimento_mails = organize_to_dt_nasc(persons)
        dt_nascimento_mails.each do | key, value |
            if value.length == 0 then
                p "empty: #{key} "
            else
                export_to_csv(key, value)
                export_to_xlsx(key, value)
            end

        end

         if_mails = organize_to_if (persons)
         if_mails.each do | key, value |
             if value.length == 0 then
                 p "empty: #{key} "
             else
                 export_to_csv(key, value)
                 export_to_xlsx(key, value)
             end
         end

        escolaridade_mails = organize_to_escolaridade (persons)
        escolaridade_mails.each do | key, value |
            if value.length == 0 then
                p "empty: #{key} "
            else
                export_to_csv(key, value)
                export_to_xlsx(key, value)
            end
        end
        #
        # Third Step
        # => Export to xlsx
        ##########################
        p "exportando arquivo"
        #export_table(valid_emails)
        p "Ok"
        p "FINALIZANDO"

    end

    def is_number? string
      true if Float(string) rescue false
    end

    def organize_to_if (data)
        ret = Hash.new
        if_sim = []
        if_nao = []

        data.each do |person|
            if !person['if'].empty? then
                if I18n.transliterate(person['if']) == 'nao'
                    if_nao.push(person)
                else
                    if I18n.transliterate(person['if']) == 'sim'
                        if_sim.push(person)
                    end
                end
            end
        end
        ret['if_sim'] = if_sim
        ret['if_nao'] = if_nao

        return ret
    end

    def organize_to_dt_nasc (data)
        ret = data[1]['dt_nascimento'].split("-")
        ret = Hash.new
        menos_20 = []
        de_21_30 = []
        de_31_40 = []
        acima_40 = []
        data.each do | person |
            if !person['dt_nascimento'].empty? then
                dt = person['dt_nascimento'].split("-").first
                if is_number?(dt) then
                    if (dt.to_i >= 1996) then
                         menos_20.push(person)
                    else
                        if (dt.to_i <= 1995 && dt.to_i >= 1986) then
                             de_21_30.push(person)
                        else
                            if (dt.to_i <= 1985 && dt.to_i >= 1976) then
                                de_31_40.push(person)
                            else
                                if dt.to_i <= 1975 then
                                    acima_40.push(person)
                                end
                            end
                        end
                    end
                end
            end
        end
        ret['menos_20'] = menos_20
        ret['de_21_30'] = de_21_30
        ret['de_31_40'] = de_31_40
        ret['acima_40'] = acima_40
        return ret
    end

    def organize_to_escolaridade (data)
        ret = data[1]['escolaridade'].split("-")
        ret = Hash.new
        fundamental = []
        medio = []
        sup_cursando = []
        sup_completo = []
        data.each do | person |
            if !person['escolaridade'].empty? then
                if I18n.transliterate(person['escolaridade']) == 'fundamental' then
                     fundamental.push(person)
                else
                    if I18n.transliterate(person['escolaridade']) == 'medio' then
                         medio.push(person)
                    else
                        if I18n.transliterate(person['escolaridade']) == 'superior completo' then
                            sup_completo.push(person)
                        else
                            if I18n.transliterate(person['escolaridade']) == 'superior cursando' then
                                sup_cursando.push(person)
                            end
                        end
                    end
                end
            end
        end
        ret['fundamental'] = fundamental
        ret['medio'] = medio
        ret['sup_completo'] = sup_completo
        ret['sup_cursando'] = sup_cursando

        return ret
    end

    def organize_data (data)

        ret = Array.new
        data.each_with_pagename do | name, table_mail |
            ## log
            # ======
            # ======
            table_mail.each do | row |
                if !row[3].eql?("UF") then
                    person = {
                        "nome" => row[0].to_s.downcase,
                        "email" => row[1].to_s.downcase,
                        "cidade" => row[2].to_s.downcase,
                        "uf" => row[3].to_s.downcase,
                        "dt_nascimento" => row[4].to_s.downcase,
                        "if" => row[5].to_s.downcase,
                        "escolaridade" => row[6].to_s.downcase
                    }
                    ret.push(person)
                end
            end
        end
        return ret

    end

    def organize_to_uf (data)
        ret = Hash.new
        rs = []
        sc = []
        pr = []
        sp = []
        ms = []
        mg = []
        rj = []
        go = []
        mt = []
        df = []
        es = []
        ba = []
        to = []
        ro = []
        ac = []
        am = []
        rr = []
        ap = []
        pa = []
        ma = []
        pi = []
        ce = []
        rn = []
        pb = []
        pe = []
        al = []
        se = []

        data.each do | person |
            case person['uf']
                when 'rs'
                    rs.push(person)
                when 'sc'
                    sc.push(person)
                when 'pr'
                    pr.push(person)
                when 'sp'
                    sp.push(person)
                when 'ms'
                    ms.push(person)
                when 'rj'
                    rj.push(person)
                when 'mg'
                    mg.push(person)
                when 'go'
                    go.push(person)
                when 'df'
                    df.push(person)
                when 'mt'
                    mt.push(person)
                when 'ro'
                    ro.push(person)
                when 'ac'
                    ac.push(person)
                when 'es'
                    es.push(person)
                when 'ba'
                    ba.push(person)
                when 'to'
                    to.push(person)
                when 'pa'
                    pa.push(person)
                when 'am'
                    am.push(person)
                when 'rr'
                    rr.push(person)
                when 'ap'
                    ap.push(person)
                when 'ma'
                    ma.push(person)
                when 'pi'
                    pi.push(person)
                when 'ce'
                    ce.push(person)
                when 'rn'
                    rn.push(person)
                when 'pb'
                    pb.push(person)
                when 'pe'
                    pe.push(person)
                when 'al'
                    al.push(person)
                when 'se'
                    se.push(person)
                else
                    # nada
            end
        end

        ret['rs'] = rs
        ret['sc'] = sc
        ret['pr'] = pr
        ret['sp'] = sp
        ret['ms'] = ms
        ret['mg'] = mg
        ret['rj'] = rj
        ret['go'] = go
        ret['mt'] = mt
        ret['df'] = df
        ret['es'] = es
        ret['ba'] = ba
        ret['to'] = to
        ret['ro'] = ro
        ret['ac'] = ac
        ret['am'] = am
        ret['rr'] = rr
        ret['ap'] = ap
        ret['pa'] = pa
        ret['ma'] = ma
        ret['pi'] = pi
        ret['ce'] = ce
        ret['rn'] = rn
        ret['pb'] = pb
        ret['pe'] = pe
        ret['al'] = al
        ret['se'] = se

        return ret
    end

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

    #
    # validate emails
    #
    ############################
    def validate_emails(data)
        retValid = Array.new
        retInvalid = Array.new
        contValid = 0
        contFalse = 0

        data.each_with_pagename do | name, table_mail |
            table_mail.each do | row |
                if !row[1].eql?("to") then
                    person = Person.new
                    person.nome          = row[0].to_s.downcase
                    person.email         = row[1].to_s.downcase
                    person.cidade        = row[2].to_s.downcase
                    person.uf            = row[3].to_s.downcase
                    person.dt_nascimento = row[4].to_s.downcase
                    person.if            = row[5].to_s.downcase
                    person.escolaridade  = row[6].to_s.downcase

                    if person.valid? then
                        per = {
                            "nome" => person.nome,
                            "email" => person.email,
                            "cidade" => person.cidade,
                            "uf" => person.uf,
                            "dt_nascimento" => person.dt_nascimento,
                            "if" => person.if,
                            "escolaridade" => person.escolaridade
                        }
                        if valid_email_host?(person.email) then
                            retValid.push(per)
                            contValid += 1
                        else
                            retInvalid.push(per)
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
        return ret = {'validos' => retValid, 'invalidos' => retInvalid}
    end

    #
    # class person
    #
    ############################
    class Person
        include ActiveModel::Validations
        attr_accessor :nome, :email, :cidade, :uf, :dt_nascimento, :if, :escolaridade

        validates :nome, :presence => false, :length => { :maximum => 100 }
        validates :email, :presence => true, :email => true

    end

    init()
