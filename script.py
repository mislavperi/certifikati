import fpdf
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
import pandas as pd
import requests
import json
import re

path_excel = os.getcwd() + "/uvjeti_certifikat_2022.xlsx"

def get_and_remove_first_pair_dict(activities_dict):
    first_pair = next(iter((activities_dict.items())))
    title = first_pair[0]
    desc = first_pair[1]
    del activities_dict[title]
    return title, desc, activities_dict

def get_data_from_excel(key):
    excel_data_dict = {}
    table = pd.read_excel(path_excel)
    excel_data_dict["team_name"] = table.loc[key][4]
    print(table.loc[key][1])
    excel_data_dict["position"] = table.loc[key][5]
    excel_data_dict["team_desc"] = get_team_description(excel_data_dict["team_name"])
    excel_data_dict["email"] = table.loc[key][2]
    #excel_data_dict["email"] = "paula.anic@estudent.hr"
    return excel_data_dict

def get_team_description(team_name):
    team_desc = ''
    team_desc_table = pd.read_excel(path_excel, "Timovi")
    team_desc_dict = team_desc_table.to_dict()
    for key, team_name_compared in team_desc_dict["Ime tima"].items():
        if team_name_compared == team_name:
            team_desc = team_desc_dict["Opis tima"][key]
            return team_desc 
    return team_desc

def get_api_data(email):
    req = requests.get('https://info.estudent.hr/api/v1/certificate/' + email)
    if not req:
        raise Exception()
    data = req.content
    if not data:
        raise Exception()
    data_dict = json.loads(data)
    return data_dict

def get_team_gender_and_recommendation(data_dict):
    recommendations = data_dict["recommendations"]
    team_gender_and_recommendation_dict = {}

    for r in recommendations:
        if r["role_in_team"]["academic_year_id"] == 11:
            team = r["role_in_team"]["team"]
            #team_description = team["description"]
            gender = r["role_in_team"]["user"]["gender"]
            first_name = r["role_in_team"]["user"]["first_name"]
            last_name = r["role_in_team"]["user"]["last_name"]
            pattern = r'(\r\n\r\n)'
            recommendation = r["recommendation"]
            recommendation = re.split(pattern,recommendation)
            #recommendation = recommendation[0]
            recommender_name = r["recommender"]["first_name"] + " " + r["recommender"]["last_name"]
            recommender_gender = r["recommender"]["gender"]

            #team_gender_and_recommendation_dict["team_description"] = team_description
            team_gender_and_recommendation_dict["full_name"] = first_name + ' ' + last_name
            team_gender_and_recommendation_dict["gender"] = gender
            if recommendation:
                team_gender_and_recommendation_dict["recommendation"] = recommendation
                team_gender_and_recommendation_dict["recommender_name"] = recommender_name
                team_gender_and_recommendation_dict["recommender_gender"] = recommender_gender
    
    return team_gender_and_recommendation_dict

def get_activities(data_dict):
    activities_dict = {}
    activities = data_dict["activities"]
    activities_dict["tajnik"] = False
    for a in activities:
        act_type = a["activity_type_id"]
        if act_type == 12:
            activities_dict["tajnik"] = True
            print("tajnik")
        if act_type == 1 or act_type == 4 or act_type == 6 or act_type == 12:
            title = a["title"]
            description = a["description"]
            activities_dict[title] = description
    if activities_dict["tajnik"]:
        print("TAJNIK nap cert")
    return activities_dict

def create_pdf(activities_dict, email, full_name, gender, position, team_name, team_desc, recommendation, recommendation_signature):
    ### Define file names
    overlay_pdf_file_name = 'overlay.pdf'
    #pdf_template_file_name = './templates/certifikat_udruga_2022_2.pdf'
    pdf_template_file_name = "./templates/CERTIFIKAT-2022.pdf"
    pdf_template_page_clean_path ='./templates/CERTIFIKAT-2022-v2.pdf'
    result_pdf_file_name = './Certifikati-2022'+ '/' + team_name + '/' + email + '/' + full_name + ' eSTUDENT Certifikat 2021-2022.pdf'
    #result_pdf_file_name = './Cert2022/' + full_name + ' eSTUDENT Certifikat 2020-2022.pdf'
    del activities_dict["tajnik"]
    # Get position desc text
    if team_name == "Predsjedni??tvo Udruge":
        if email == "marija.samardzic@estudent.hr":
            position_gendered_text = 'aktivno je sudjelovala kao ' + "glavna tajnica udruge" + '\n'
        if email == "paula.stefanek@estudent.hr" or email == "paula.paic@estudent.hr":
            position_gendered_text = 'aktivno je sudjelovala kao ' + "podpredsjednica udruge" + '\n'
        else:
            position_gendered_text = 'aktivno je sudjelovala kao ' + position + ' tima ' + team_name + '\n'

    else:
        if gender == 'M':
            position_gendered_text = 'aktivno je sudjelovao kao ' + position + ' tima ' + team_name + '\n'
        else:
            position_gendered_text = 'aktivno je sudjelovala kao ' + position + ' tima ' + team_name + '\n'

    ### Create first two pages
    pdf = fpdf.FPDF(format='letter', unit='pt')
    pdf.set_auto_page_break(auto=True, margin = 80)
    pdf.add_font(fname='./fonts/poppinsr.ttf', family='poppinsr', uni=True)
    pdf.add_font(fname='./fonts/poppinsb.ttf', family='poppinsb', uni=True)
    pdf.add_page()
    #pdf_style = 'B'
    pdf.set_font("poppinsb", size=24)
    pdf.set_left_margin(40)
    pdf.set_right_margin(510)

    # Name
    pdf.set_y(160)
    pdf.set_x(45)
    pdf.multi_cell(w=510,h=32, txt=full_name, align='C')

    # Position
    y = pdf.get_y()
    pdf.set_y(y + 20)
    pdf.set_font("poppinsr", size=14)
    pdf.multi_cell(w=510,h=24, txt=position_gendered_text, align='C')

    # Team name
    y = pdf.get_y()
    pdf.set_y(y + 20)
    pdf.set_font("poppinsb", size=12)
    pdf.multi_cell(w=510,h=14, txt=team_name, align='L')

    # Team description
    y = pdf.get_y()
    pdf.set_y(y + 10)
    pdf.set_font("poppinsr", size=10)
    pdf.multi_cell(w=510,h=14, txt=team_desc, align='J')

    #Recommendation title
    #Zbog ivinog dugog recommendationa stavi 130, inace drzi na 180
    if team_name == "Predsjedni??tvo Udruge":
        y = y + 120
    else:
        y = y + 120
    #y = y + 130
    pdf.set_y(y)
    pdf.set_font("poppinsb", size=12)
    pdf.multi_cell(w=510,h=14, txt='Pisana preporuka', align='J')

    # Recommendation text
    y = pdf.get_y()
    pdf.set_y(y + 10)
    pdf.set_font("poppinsr", size=10)
    pdf.multi_cell(w=510,h=14, txt=recommendation[0], align='J')
    pdf.multi_cell(w=510,h=14, txt=recommendation[1], align='J')
    pdf.multi_cell(w=510,h=14, txt=recommendation[2], align='J')


    pdf.add_page()
    #pdf_style = 'B'
    pdf.set_font("poppinsr", size=24)
    pdf.set_left_margin(40)
    pdf.set_right_margin(510)

    # Team name
    y = pdf.get_y()
    pdf.set_y(y + 160)
    pdf.set_font("poppinsr", size=10)
    pdf.multi_cell(w=510,h=14, txt=recommendation[3], align='J')
    pdf.multi_cell(w=510,h=14, txt=recommendation[4], align='J')
    pdf.multi_cell(w=510,h=14, txt=recommendation[5], align='J')
    pdf.multi_cell(w=510,h=14, txt=recommendation[6], align='J')

    #else:
        #pdf.multi_cell(w=510,h=14, txt=recommendation, align='J')

    # Recommendation potpis
    #y = pdf.get_y()
    #pdf.set_y(y + 10)
    #pdf.set_font("poppinsr", size=10)
    #pdf.multi_cell(w=510,h=14, txt="Stella ??aja, predsjednica udruge", align='J')

    opis_josipa = """Josipa je, uz ??lanstvo u timu Dobrotvorne i ekolo??ke aktivnosti, odabrala biti ??lanom radne skupine osnovane radi implementacije novog projekta u udrugu. Time je pokazala da je spremna mijenjati stvari i da je vrlo ambiciozna. Istaknula se me??u prijavljenima i svoju ljubav prema ??ivotinjama usmjerila u aktivnosti koje im poma??u. 
Rad na projektu koji se izvodi po prvi puta donosi brojne uspone i padove, no ??apa humanitarnosti je nadma??ila sva o??ekivanja,  ??emu je Josipa uvelike pridonijela.  Vje??tine koje je stekla su: Project Management, Time Management, kriti??ko razmi??ljanje i rje??avanje problema, timski rad te poslovna komunikacija. Tako??er, pokazala je kreativnost, pouzdanost, organiziranost, prilagodljivost i savladavanje brojnih izazova, empatiju prema ljudima i ??ivotinjama i volju za u??enjem. 
Josipa je osoba koja ??e dati sve od sebe da uspije u onome ??to radi i koja se daje srcem u projekt i posao. Prati nadolaze??e rokove i zadatke te brine da se niti jedan ne propusti, ??ak i ako je potrebno podsjetiti nekoga ???iznad sebe???. Organizacijske vje??tine su joj na visokoj razini, zbog ??ega je osoba na koju se uvijek mo??ete osloniti.
"""

    opis_lana = """Lana je, uz ??lanstvo u timu Public Relations, odabrala biti ??lanom radne skupine osnovane radi implementacije novog projekta u udrugu. Time je pokazala da je spremna mijenjati stvari i da je vrlo ambiciozna. Istaknula se me??u prijavljenima i svoju ljubav prema ??ivotinjama usmjerila u aktivnosti koje im poma??u. 
Rad na projektu koji se izvodi po prvi puta donosi brojne uspone i padove, no ??apa humanitarnosti je nadma??ila sva o??ekivanja,  ??emu je Lana uvelike pridonijela.  Vje??tine koje je stekla su: Project Management, Time Management, kriti??ko razmi??ljanje i rje??avanje problema, timski rad te poslovna komunikacija. Tako??er, pokazala je kreativnost, pouzdanost, organiziranost, prilagodljivost i savladavanje brojnih izazova, empatiju prema ljudima i ??ivotinjama i volju za u??enjem. 
Lana je osoba koja se u potpunosti daje u posao koji obavlja. S lako??om savladava zadatke koji su joj dani te istra??uje kako bi mogla dodatno pridonijeti. Voli biti upu??ena u sve ??to se s projektom doga??a i sudjelovati u svakoj fazi. Ona je osoba na koju se uvijek mo??ete osloniti i koja ??e dati maksimum od sebe kako bi projekt bio ??to uspje??niji. Lana je svestrana, odva??na i hrabra,  te??i napretku i nakon obavljenog zadatka ve?? gleda ??to bi mogla dalje. 
"""

    opis_jure = """Jure je, uz ??lanstvo u timu Ra??unovodstvo i financije, odabrao biti ??lanom radne skupine osnovane radi implementacije novog projekta u udrugu. Time je pokazao da je spreman mijenjati stvari i da je vrlo ambiciozan. Istaknuo se me??u prijavljenima i svoju ljubav prema ??ivotinjama usmjerio u aktivnosti koje im poma??u. 
Rad na projektu koji se izvodi po prvi puta donosi brojne uspone i padove, no ??apa humanitarnosti je nadma??ila sva o??ekivanja,  ??emu je Jure uvelike pridonio.  Vje??tine koje je stekao su: Project Management, Time Management, kriti??ko razmi??ljanje i rje??avanje problema, timski rad te poslovna komunikacija. Tako??er, pokazao je kreativnost, pouzdanost, organiziranost, prilagodljivost i savladavanje brojnih izazova, empatiju prema ljudima i ??ivotinjama i volju za u??enjem. 
Jure je vrlo sr??ana i pozitivna osoba, ??iji karakter donosi dozu veselja u svaku prostoriju. Svojim zadacima pristupa odgovorno, vrlo je organiziran i bez problema paralelno obavlja zadatke na vi??e projekata i u vi??e timova. Po??tuje rokove te shva??a da je priprema pola posla, zato je uvijek u toku s onime ??to se doga??a i njegova ma??tovitost rezultira novim idejama i pogledima na stvari. Svaki tim treba jednog Juru. 
"""

    opis_robert = """Robert je, uz ??lanstvo u timu Dru??tveno odgovorno poslovanje, odabrao biti ??lanom radne skupine osnovane radi implementacije novog projekta u udrugu. Time je pokazao da je spreman mijenjati stvari i da je vrlo ambiciozan. Istaknuo se me??u prijavljenima i svoju ljubav prema ??ivotinjama usmjerio u aktivnosti koje im poma??u. 
Rad na projektu koji se izvodi po prvi puta donosi brojne uspone i padove, no ??apa humanitarnosti je nadma??ila sva o??ekivanja,  ??emu je Robert uvelike pridonio.  Vje??tine koje je stekao su: Project Management, Time Management, kriti??ko razmi??ljanje i rje??avanje problema, timski rad te poslovna komunikacija. Tako??er, pokazao je kreativnost, organiziranost, te empatiju prema ljudima i ??ivotinjama. 
Robert je svestrana i vrlo simpati??na osoba. Svojim karakterom doprinosi pozitivnoj atmosferi u timu, zadatke odra??uje na vrijeme i korektno te razmi??lja ???out of the box???. Studiranje na Medicinskom fakultetu, posao i ??lanstvo u dva tima studentske udruge dokazuje da je vrlo organiziran i sposoban te zna kako poslo??iti prioritete i odraditi zadatke koje je na sebe preuzeo. 
"""

    opis_suncana = """Sun??ana je, uz ??lanstvo u timu Dobrotvorne ekolo??ke aktivnosti, odabrala biti ??lanom radne skupine osnovane radi implementacije novog projekta u udrugu. Time je pokazala da je spremna mijenjati stvari i da je vrlo ambiciozna. Istaknula se me??u prijavljenima i svoju ljubav prema ??ivotinjama usmjerila u aktivnosti koje im poma??u. 
Rad na projektu koji se izvodi po prvi puta donosi brojne uspone i padove, no ??apa humanitarnosti je nadma??ila sva o??ekivanja,  ??emu je Sun??ana uvelike pridonijela.  Vje??tine koje je stekla su: Project Management, Time Management, kriti??ko razmi??ljanje i rje??avanje problema, timski rad te poslovna komunikacija. Tako??er, pokazala je kreativnost, pouzdanost, organiziranost, prilagodljivost i savladavanje brojnih izazova, empatiju prema ljudima i ??ivotinjama i volju za u??enjem. 
Sun??ana je vesela i simpati??na osoba, koja zna prepoznati ozbiljnost situacije i odli??no odraditi svaki zadatak koji se pred nju stavi. Svoju emotivnost i ljudskost pokazuje kroz timski rad i ??elju za najboljom izvedbom projekta ili zadatka na kojem radi. Osoba je na koju se mo??ete osloniti i koja je spremna u??iti. 
"""

    opis_iva = """Iva je, uz ??lanstvo u timu Ra??unovodstvo i financije, odabrala biti ??lanom radne skupine osnovane radi implementacije novog projekta u udrugu. Time je pokazala da je spremna mijenjati stvari i da je vrlo ambiciozna. Istaknula se me??u prijavljenima i svoju ljubav prema ??ivotinjama usmjerila u aktivnosti koje im poma??u. 
Rad na projektu koji se izvodi po prvi puta donosi brojne uspone i padove, no ??apa humanitarnosti je nadma??ila sva o??ekivanja,  ??emu je Iva uvelike pridonijela.  Vje??tine koje je stekla su: Project Management, Time Management, kriti??ko razmi??ljanje i rje??avanje problema, timski rad te poslovna komunikacija. Tako??er, pokazala je pouzdanost, organiziranost, prilagodljivost i savladavanje brojnih izazova, empatiju prema ljudima i ??ivotinjama i volju za u??enjem. 
Iva je jako simpati??na i pozitivna osoba, koja stvara osje??aj ugode i dobre atmosfere u timskom radu. Iako se s njom mo??ete dobro nasmijati, shva??a ozbiljnost posla, svoje zadatke odra??uje s lako??om te ula??e puno u ono do ??ega joj je stalo. Vrlo je organizirana i odgovorna, bez problema je odr??avala radni tempo u oba tima i pokazala da je ujedno i timski igra?? i vrlo samostalna. 
"""

    opis_nikolina = """Nikolina je, uz ??lanstvo u timu Vara??din, odabrala biti ??lanom radne skupine osnovane radi implementacije novog projekta u udrugu. Time je pokazala da je spremna mijenjati stvari i da je vrlo ambiciozna. Istaknula se me??u prijavljenima i svoju ljubav prema ??ivotinjama usmjerila u aktivnosti koje im poma??u. 
Rad na projektu koji se izvodi po prvi puta donosi brojne uspone i padove, no ??apa humanitarnosti je nadma??ila sva o??ekivanja,  ??emu je Nikolina uvelike pridonijela.  Vje??tine koje je stekla su: Project Management, Time Management, kriti??ko razmi??ljanje i rje??avanje problema, timski rad te poslovna komunikacija. Tako??er, pokazala je pouzdanost, organiziranost, prilagodljivost i savladavanje brojnih izazova te empatiju prema ljudima i ??ivotinjama.
Nikolina se prijavila u radnu skupinu iako ??ivi u drugom gradu te je znala da ??e se morati potruditi vi??e kako bi mogla sudjelovati u projektu i osjetiti povezanost s timom. Time je pokazala da se ne boji izazova i da je predana onome ??to radi, kao i da je spremna u??iti i napredovati. Svoje je zadatke odra??ivala na vrijeme, davala nove ideje i ostala uklju??ena u projekt do samog kraja, usprkos udaljenosti. Dokazala je da je ponekad bitna samo volja i upornost kako bi pomaknuli granice.
"""

    opis_tia = """Tia je, uz ??lanstvo u timu Vara??din, odabrala biti ??lanom radne skupine osnovane radi implementacije novog projekta u udrugu. Time je pokazala da je spremna mijenjati stvari i da je vrlo ambiciozna. Istaknula se me??u prijavljenima i svoju ljubav prema ??ivotinjama usmjerila u aktivnosti koje im poma??u. 
Rad na projektu koji se izvodi po prvi puta donosi brojne uspone i padove, no ??apa humanitarnosti je nadma??ila sva o??ekivanja,  ??emu je Tia uvelike pridonijela.  Vje??tine koje je stekla su: Project Management, Time Management, kriti??ko razmi??ljanje i rje??avanje problema, timski rad te poslovna komunikacija. Tako??er, pokazala je organiziranost, prilagodljivost i savladavanje brojnih izazova, empatiju prema ljudima i ??ivotinjama i volju za u??enjem. 
Tia je jedna simpati??na i pozitivna osoba, s kojom je vrlo jednostavno komunicirati i sura??ivati. Svoje zadatke shva??a ozbiljno i odra??uje ih usprkos brojnim obavezama. Doprinosi ugodnoj atmosferi u timu, ne ustru??ava se iznijeti svoje mi??ljenje te pokazuje koliko joj je stalo do stvari kojima se bavi. ??ak i kad zbog privatnih planova nije u Zagrebu prona??e na??in kako sudjelovati na sastancima i u samom projektu, ??to pokazuje volju i odgovornost. 
"""
    if email == "jure.gagro@estudent.hr" or email == "iva.krizanac@estudent.hr" or email == "josipa.manduric@estudent.hr" or email == "lana.ivic@estudent.hr" or email == "robert.katinic@estudent.hr" or email == "suncana.jantolek@estudent.hr" or email == "nikolina.cimerman@estudent.hr" or email == "tia.knezevic@estudent.hr":
        position_gendered_text = 'aktivno je sudjelovala kao ' + ' ??lanica radne skupine ??apa humanitarnosti' + '\n'

        ### Write second page
        if email == "jure.gagro@estudent.hr":
            title = "Radna skupina ??apa humanitarnosti"
            desc = opis_jure
            position_gendered_text = 'aktivno je sudjelovao kao ' + ' ??lan radne skupine ??apa humanitarnosti' + '\n'

        elif email == "iva.krizanac@estudent.hr":
            title = "Radna skupina ??apa humanitarnosti"
            desc = opis_iva
            
        elif email == "josipa.manduric@estudent.hr":
            title = "Radna skupina ??apa humanitarnosti"
            desc = opis_josipa
            
        elif email == "lana.ivic@estudent.hr":
            title = "Radna skupina ??apa humanitarnosti"
            desc = opis_lana
            
        elif email == "robert.katinic@estudent.hr":
            title = "Radna skupina ??apa humanitarnosti"
            desc = opis_robert
            position_gendered_text = 'aktivno je sudjelovao kao ' + ' ??lan radne skupine ??apa humanitarnosti' + '\n'

            
        elif email == "suncana.jantolek@estudent.hr":
            title = "Radna skupina ??apa humanitarnosti"
            desc = opis_suncana
            
        elif email == "nikolina.cimerman@estudent.hr":
            title = "Radna skupina ??apa humanitarnosti"
            desc = opis_nikolina
            
        elif email == "tia.knezevic@estudent.hr":
            title = "Radna skupina ??apa humanitarnosti"
            desc = opis_tia
        team_desc2 = "Radna skupina ??apa humanitarnosti je organizator prvog (istoimenog) projekta u Udruzi usmjerenog na ??ivotinje, ??to predstavlja novo podru??je djelovanja eSTUDENTa. ??apa humanitarnosti je humanitarni i ekolo??ki projekt ??ija je primarna svrha pomo?? ??ivotinjama. Cilj projekta je edukacija i aktivacija studentske zajednice (primarno) o zero waste na??elima kroz prikupljanje i recikliranje stare odje??e te njenu prenamjenu u igra??ke za pse i ma??ke. Aukcijom izra??enih igra??aka poti??e se prikupljanje sredstava za zbrinjavanje i pomo?? ??ivotinja u skloni??tima."
        pdf.add_page()
        #pdf_style = 'B'
        pdf.set_font("poppinsb", size=24)
        pdf.set_left_margin(40)
        pdf.set_right_margin(510)
        
        # Name
        pdf.set_y(160)
        pdf.set_x(45)
        pdf.multi_cell(w=510,h=32, txt=full_name, align='C')

        # Position
        y = pdf.get_y()
        pdf.set_y(y + 20)
        pdf.set_font("poppinsr", size=14)
        pdf.multi_cell(w=510,h=24, txt=position_gendered_text, align='C')

        # Team name
        y = pdf.get_y()
        pdf.set_y(y + 30)
        pdf.set_font("poppinsb", size=12)
        pdf.multi_cell(w=510,h=14, txt="Radna skupina ??apa humanitarnosti", align='L')

        # Team description
        y = pdf.get_y()
        pdf.set_y(y + 10)
        pdf.set_font("poppinsr", size=10)
        pdf.multi_cell(w=510,h=14, txt=team_desc2, align='J')
        #Recommendation title
        #Zbog ivinog dugog recommendationa stavi 130, inace drzi na 180
        #if team_name == "Predsjedni??tvo Udruge":
        #    y = y + 180
        #else:
        #    y = y + 130
        y = y + 100
        pdf.set_y(y)
        pdf.set_font("poppinsb", size=12)
        pdf.multi_cell(w=510,h=14, txt='Pisana preporuka', align='J')

        # Recommendation text
        y = pdf.get_y()
        pdf.set_y(y + 10)
        pdf.set_font("poppinsr", size=10)
        pdf.multi_cell(w=510,h=14, txt=desc, align='J')

        # Recommendation potpis
        #y = pdf.get_y()
        #pdf.set_y(y + 10)
        #pdf.set_font("poppinsr", size=10)
        #pdf.multi_cell(w=510,h=14, txt="Karla Georgiev, voditeljica radne skupine ??apa humanitarnosti", align='J')

    pdf.add_page()
    pdf.set_font("poppinsb", size=22)
    pdf.multi_cell(w=510,h=22, txt='Predavanja, radionice i ostale aktivnosti', align='C')

    
   
    #First tile, dont calculate y_diff
    if len(activities_dict) == 0:
        file1 = open("no_activities.txt","a")
        file1.write(email + "; " + "\n")
        file1.close()
        
        pdf.output(overlay_pdf_file_name)
        pdf.close()
        pdf_template = PdfFileReader(open(pdf_template_file_name, 'rb'))
        overlay_pdf = PdfFileReader(open(overlay_pdf_file_name, 'rb'))

        # Get the first page from the template
        # Merge first two pages
        template_page_first = pdf_template.getPage(0)
        template_page_first.mergePage(overlay_pdf.getPage(0))

        output_pdf = PdfFileWriter()
        output_pdf.addPage(template_page_first)

        ### Ako ocemo fileove u folderima uncomment kod ispod

        if not os.path.isdir('./' + 'Certifikati-2022/' + team_name):
            os.mkdir('./' + 'Certifikati-2022/' + team_name)

        if not os.path.isdir('./' + 'Certifikati-2022/' + team_name + '/' + email):
            os.mkdir('./' + 'Certifikati-2022/' + team_name + '/' + email)
        
        
        output_pdf.write(open(result_pdf_file_name, "wb"))

        return
        
    else:
        title, desc, activities_dict = get_and_remove_first_pair_dict(activities_dict)
    txt_len = len(desc)
    pdf.set_y(pdf.get_y() + 40)
    old_y = pdf.get_y()
    pdf.set_font("poppinsb", size=12)
    pdf.multi_cell(w=510,h=14, txt=title, align='J')
    #pdf.set_y(pdf.get_y() + 10)
    pdf.set_font("poppinsr", size=10)
    pdf.multi_cell(w=510,h=14, txt=desc, align='J')
    if txt_len > 777:
        old_y = old_y + 60


    # Take the PDF you created above and overlay it on your template PDF
    # Open your template PDF
    pdf_template = PdfFileReader(open(pdf_template_file_name, 'rb'))
    pdf_template_page_clean = PdfFileReader(open(pdf_template_page_clean_path, 'rb'))
    overlay_pdf = PdfFileReader(open(overlay_pdf_file_name, 'rb'))

    # Get the first page from the template
    # Merge first two pages
    template_page_first = pdf_template.getPage(0)
    template_page_clean = pdf_template_page_clean.getPage(1)
    #template_page_first.mergePage(overlay_pdf.getPage(0))
    #template_page_first.mergePage(overlay_pdf.getPage(1))
    #template_page_first.mergePage(overlay_pdf.getPage(2))

    output_pdf = PdfFileWriter()
    output_pdf.addPage(template_page_first)
    output_pdf.addPage(template_page_clean)

    y_diff = 0

    activities_len = len(activities_dict)
    p = 1
    while activities_len > 0:
        p += 1
        title, desc, activities_dict = get_and_remove_first_pair_dict(activities_dict)
        if p == 4:
            pdf.add_page()
            p = 1
            y_diff = 0
            old_y = 0
            activities_len = activities_len - 1
            txt_len = len(desc)
            pdf.set_y(pdf.get_y() + 20)
            pdf.set_font("poppinsb", size=12)
            pdf.multi_cell(w=510,h=14, txt=title, align='J')
            #pdf.set_y(pdf.get_y() + 10)
            pdf.set_font("poppinsr", size=10)
            pdf.multi_cell(w=510,h=14, txt=desc, align='J')
            old_y = 78
            if txt_len > 777:
                old_y = old_y + 14
            pdf.set_y(pdf.get_y() + 10)
            
        else:
            activities_len = activities_len - 1
            txt_len = len(desc)
            y_diff = pdf.get_y() - old_y
            pdf.set_y(pdf.get_y() + 165 - y_diff)
            old_y = pdf.get_y()
            pdf.set_font("poppinsb", size=12)
            pdf.multi_cell(w=510,h=14, txt=title, align='J')
            pdf.set_font("poppinsr", size=10)
            pdf.multi_cell(w=510,h=14, txt=desc, align='J')
            if txt_len > 777:
                old_y = old_y + 14
    
    
    pdf.output(overlay_pdf_file_name)
    pdf.close()

    pdf_template = PdfFileReader(open(pdf_template_file_name, 'rb'))
    pdf_template2 = PdfFileReader(open(pdf_template_file_name, 'rb'))
    pdf_template_page_clean = PdfFileReader(open(pdf_template_page_clean_path, 'rb'))
    overlay_pdf = PdfFileReader(open(overlay_pdf_file_name, 'rb'))

    # Get the first page from the template
    # Merge first two pages
    template_page_first = pdf_template.getPage(0)
    #if email == "jure.gagro@estudent.hr" or email == "iva.krizanac@estudent.hr" or email == "josipa.manduric@estudent.hr" or email == "lana.ivic@estudent.hr" or email == "robert.katinic@estudent.hr" or email == "suncana.jantolek@estudent.hr" or email == "nikolina.cimerman@estudent.hr" or email == "tia.knezevic@estudent.hr":
    template_page_firstt = pdf_template2.getPage(0)
    #template_page_clean = pdf_template_page_clean.getPage(1)
    template_page_first.mergePage(overlay_pdf.getPage(0))
    #if email == "jure.gagro@estudent.hr" or email == "iva.krizanac@estudent.hr" or email == "josipa.manduric@estudent.hr" or email == "lana.ivic@estudent.hr" or email == "robert.katinic@estudent.hr" or email == "suncana.jantolek@estudent.hr" or email == "nikolina.cimerman@estudent.hr" or email == "tia.knezevic@estudent.hr":
    template_page_firstt.mergePage(overlay_pdf.getPage(1))
    template_page_clean.mergePage(overlay_pdf.getPage(2))
    #else:
        #template_page_clean.mergePage(overlay_pdf.getPage(1))


    output_pdf = PdfFileWriter()
    output_pdf.addPage(template_page_first)
    #if email == "jure.gagro@estudent.hr" or email == "iva.krizanac@estudent.hr" or email == "josipa.manduric@estudent.hr" or email == "lana.ivic@estudent.hr" or email == "robert.katinic@estudent.hr" or email == "suncana.jantolek@estudent.hr" or email == "nikolina.cimerman@estudent.hr" or email == "tia.knezevic@estudent.hr":
    output_pdf.addPage(template_page_firstt)
    output_pdf.addPage(template_page_clean)

    ### Merge all pages for dict
    page_num = overlay_pdf.getNumPages()
    i = 2
    while 2 < page_num:
        pdf_template_page_clean = PdfFileReader(open(pdf_template_page_clean_path, 'rb'))
        template_page_clean = pdf_template_page_clean.getPage(1)
        template_page_clean.mergePage(overlay_pdf.getPage(i))
        output_pdf.addPage(template_page_clean)
        page_num -= 1
        i += 1

    # Write the result to a new PDF file
    if not os.path.isdir('./' + 'Certifikati-2022/' + team_name):
        os.mkdir('./' + 'Certifikati-2022/' + team_name)

    if not os.path.isdir('./' + 'Certifikati-2022/' + team_name + '/' + email):
        os.mkdir('./' + 'Certifikati-2022/' + team_name + '/' + email)
    print("certifikat je gotov za " + email)

    output_pdf.write(open(result_pdf_file_name, "wb"))

def check_if_created(email, team_name):
    if os.path.isdir('./' + 'Certifikati-2022/' + team_name + '/' + email):
        return True
    else:
        return False

def generate_certificate(key):
    excel_data_dict = get_data_from_excel(key)
    email = excel_data_dict["email"]
    print('Izradujem certifikat za ' + email)
    team_name = excel_data_dict["team_name"]
    #if team_name == "Predsjedni??tvo Udruge":
    created = check_if_created(email, team_name)
    if created:
       return
    position = excel_data_dict["position"]

    team_desc = excel_data_dict["team_desc"]

    #else:
    #    retur
    """
    #Provjera uvjeta iz apija, zakomentirano je jer je bilo problema sa eIZBORIMA pa nije cijeli marketing imo uvjete
    if not data_dict["passed"]:
        print(email + ' nije zadovoljila uvjete')
        file1 = open("not_passed.txt","a")
        file1.write(email + "; " + "\n")
        file1.close() #to change file access modes
        return
    """
    try:
        data_dict = get_api_data(email)
        #if not data_dict["recommendations"]:
            #raise Exception()
        #else:
            #print("here")
        passed = data_dict["passed"]
        if not passed:
            file3 = open("uvjet.txt","a")
            file3.write(email + "; " + "\n")
            file3.close()
            raise Exception("nema uvjet")

        person_data = get_team_gender_and_recommendation(data_dict)

        if not person_data["full_name"]:
            raise Exception()

        full_name = person_data["full_name"]

        gender = person_data["gender"]
        recommendation = person_data["recommendation"]
        recommendation_signature = ""

        if person_data["recommender_gender"] == "M":
            if position == "??lan":
                rec_position = "voditelj"
                recommendation_signature = person_data["recommender_name"] + ", " +  rec_position + " tima " + team_name
            else:
                rec_position = "koordinator"
                recommendation_signature = person_data["recommender_name"] + ", " +  rec_position + " tima " + team_name
        else:
            if position == "??lan":
                rec_position = "voditeljica"
                recommendation_signature = person_data["recommender_name"] + ", " +  rec_position + " tima " + team_name
            else:
                rec_position = "koordinatorica"
                recommendation_signature = person_data["recommender_name"] + ", " +  rec_position + " tima " + team_name
        if not recommendation:
            raise Exception()

        activities_dict = get_activities(data_dict)
        if activities_dict["tajnik"]:
            print("RADIM CERT ZA TAJNIKA")
            create_pdf(activities_dict, email, full_name, gender, position, team_name, team_desc, recommendation, recommendation_signature)
        
        else:
            create_pdf(activities_dict, email, full_name, gender, position, team_name, team_desc, recommendation, recommendation_signature)
            return
    except:
        # Ne nalazi ime osobe, sve osobe za 20/21 god za koje nije nasao ime nisu zadovoljili uvjete i nemaju recommendation. Rucno provjeravaj to 
        file2 = open("except.txt","a")
        file2.write(email + "; " + "\n")
        file2.close()

excel_table = pd.read_excel(path_excel).to_json()
excel_table = json.loads(excel_table)
for key in excel_table["Ime"]:
    key = int(key)
    new_key = key
    generate_certificate(new_key)
