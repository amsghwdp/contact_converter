import csv
import easygui
from ldif import LDIFParser,LDIFWriter
import glob
import os
#import vobject #required for vcf/vCard output http://eventable.github.io/vobject/

#input_formats:[delimiter, quotechar, has header line]
input_formats = {'Outlook.com CSV':(',', '"', True), 'Thunderbird CSV':(',', '"', True), 'Google CSV':(',', '"', True), 'VCF':()}
output_formats = ('LDIF', )
#field_positions = {format:(First Name, Last Name, Nickname, Email #1, Email #2, Mobile Phone, Home Phone, Work Phone, Fax)}
#Outlook.com CSV has multiple fax fields so 'business fax' is used
#Google CSV /does/ have a full name field, but a) it's difficult to integrate when nothing else has it, and b) it's just first and last names merged exactly like we do
#and doesn't have positional phone fields like others - suggest exporting as outlook csv
field_positions = { 'Outlook.com CSV':(0, 2, 5, 8, 9, 15, 11, 13, 20),  
                    'Thunderbird CSV':(0, 1, 3, 4, 5, 12, 8, 7, 9), 
                    'Google CSV':(1, 3, 11, 28, '', 33, 35, '', ''), 
                    }

source_format = ''
dest_format = ''
#keeps track of whether the VCFs will be lumped together into one file or get individual files per contact
vcf_output_files = 'Single File'

version = "1.0"
title = "Contacts Converter v%s" % (version)

contacts = {}
class contact:
    first_name = ''   #0
    last_name = ''    #1
    nickname = ''     #2
    email_1 = ''      #3
    email_2 = ''      #4
    mobile_phone = '' #5
    home_phone = ''   #6
    work_phone = ''   #7
    fax = ''          #8

def select_source_dest():
    global source_format
    global dest_format

    input_formats['Quit'] = ()
    source_format = easygui.choicebox('Please select source format', title, input_formats)
    input_formats.pop('Quit')

    if source_format == 'Quit':
        quit()
    if source_format == None:
        main_menu()
    if 'CSV' in source_format:
        read_source_csv()
    elif 'VCF' in source_format:
        read_vcf()
    elif 'LDIF' in source_format:
        read_ldif()

    dest_format = easygui.choicebox('Please select the destination format', title, output_formats)
    if dest_format == None:
        main_menu()
    if 'VCF' in dest_format:
        vcf_output_files = easygui.buttonbox('Do you want the output to be a single VCF file or individual for each contact?', title, ('Single File', 'File Per Contact'))
        if dest_format == 'VCF 3.0':
            output_vcf('3.0')
    if 'LDIF' in dest_format:
        pass

def read_source_csv():
    source_file = easygui.fileopenbox("Please select the SOURCE csv file",title,"")
    with open(source_file, newline=None) as csvfile:
        reader = csv.reader(csvfile, delimiter=input_formats[source_format][0], quotechar=input_formats[source_format][1])
        #need this messy line_num bit to ignore the header lines (where necessary) basically
        line_num = 0
        for line in reader:

            #probably integrate a header check here to warn if csv doesn't match selected source type

            if line_num > 0 or input_formats[source_format][2] == False:
                line_num += 1
                #get or create a full name (required by vCard)
                full_name = ''

                if line[field_positions[source_format][0]] != '' and line[field_positions[source_format][1]] != '':
                    full_name = (line[field_positions[source_format][0]] + ' ' + line[field_positions[source_format][1]]).title()
                elif line[field_positions[source_format][0]] == '' or line[field_positions[source_format][1]] == '':
                    full_name = (line[field_positions[source_format][0]] + line[field_positions[source_format][1]]).title()

                elif line[field_positions[source_format][0]] == '' and line[field_positions[source_format][1]] == '':
                    if line[field_positions[source_format][2]] != '':
                        #nickname
                        full_name = line[field_positions[source_format][2]].title()
                    #just grabs the part of the email address before the @ and .title()s it
                    elif line[field_positions[source_format][3]] != '':
                        full_name = line[field_positions[source_format][3]][:line[field_positions[source_format][3]].index('@')].title()
                    elif line[field_positions[source_format][4]] != '':
                        #on the off chance only email #2 is filled in
                        full_name = line[field_positions[source_format][4]][:line[field_positions[source_format][4]].index('@')].title()
                    else:
                        easygui.msgbox('There are no available fields to use as a full name, please make sure this entry has at least one of the following filled in: \n\n - First Name\n - Last Name\n - Nickname\n - Email #1\n - Email #2', title)

                contacts[full_name] = contact()
                #ugly but required for Google CSV compatability, plus makes it easier to extend
                if field_positions[source_format][0] != '':
                    contacts[full_name].first_name = line[field_positions[source_format][0]]
                else:
                    contacts[full_name].first_name = ''

                if field_positions[source_format][1] != '':
                    contacts[full_name].last_name = line[field_positions[source_format][1]]
                else:
                    contacts[full_name].last_name = ''

                if field_positions[source_format][2] != '':
                    contacts[full_name].nickname = line[field_positions[source_format][2]]
                else:
                    contacts[full_name].nickname = ''

                if field_positions[source_format][3] != '':
                    contacts[full_name].email_1 = line[field_positions[source_format][3]]
                else:
                    contacts[full_name].email_1 = ''

                if field_positions[source_format][4] != '':
                    contacts[full_name].email_2 = line[field_positions[source_format][4]]
                else:
                    contacts[full_name].email_2 = ''

                if field_positions[source_format][5] != '':
                    contacts[full_name].mobile_phone = line[field_positions[source_format][5]]
                else:
                    contacts[full_name].mobile_phone = ''

                if field_positions[source_format][6] != '':
                    contacts[full_name].home_phone = line[field_positions[source_format][6]]
                else:
                    contacts[full_name].home_phone = ''

                if field_positions[source_format][7] != '':
                    contacts[full_name].work_phone = line[field_positions[source_format][7]]
                else:
                    contacts[full_name].work_phone = ''

                if field_positions[source_format][8] != '':
                    contacts[full_name].fax = line[field_positions[source_format][8]]
                else:
                    contacts[full_name].fax = ''

            else:
                line_num += 1

def read_source_ldif():


def read_vcf():
    amount_to_read = easygui.buttonbox('If we are reading multiple VCF files, please make sure they\'re all in the same dir and click Multiple, otherwise click Single', title, ('Multiple', 'Single'))
    if amount_to_read == None:
        main_menu()
    vcfs_to_read = []
    if amount_to_read == 'Single':
        vcfs_to_read.append(easygui.fileopenbox('Please select the VCF file to read',title))
        if vcfs_to_read == None:
            main_menu()
    else:
        vcf_location = easygui.diropenbox('Please select the dir containing all the VCF files to read', title)
        if vcf_location == None:
            main_menu()
        for i in glob.glob(os.path.join(vcf_location, '*')):
            vcfs_to_read.append(i)
   
    vcf_raw = []
    for vcf_file in vcfs_to_read:
        f = open(vcf_file, newline=None)
        tmp_read = f.read()
        f.close()
        #this will actually work for files with no newline at the end of the file too
        for i in tmp_read.split('\n\n'):
            #kinda silly to load them all into a list first and then parse them, might as well do it here right?
            #except we need to actually parse them using vObject so maybe not?
            vcf_raw.append(i)

    for i in vcf_raw:
        print(i+'\n--------------', end='++')
        for line in i.split('\n'):
            print(line, end='!!\n')
        #contacts[full_name] = contact()

    #quit()

def output_vcf(vcf_ver):


    # first_name = ''   #0
    # last_name = ''    #1
    # nickname = ''     #2
    # email_1 = ''      #3
    # email_2 = ''      #4
    # mobile_phone = '' #5
    # home_phone = ''   #6
    # work_phone = ''   #7
    # fax = ''          #8

    #https://github.com/l0b0/vcard/blob/master/tests/maximal.vcf

    #
    # BEGIN:VCARD
    # VERSION:3.0
    # PROFILE:VCARD
    # CLASS:PUBLIC
    # SOURCE;VALUE=uri:http://example.org/example.vcf
    # N:Doe;John;Quentin;Mr,Dr;Esq.
    # FN:John Doe
    # NICKNAME:Johnny,Dr Q
    # ORG:Example Organization;Example Department;Example Sub-Department;Example 
    #  Office
    # LOGO;VALUE=uri:http://example.org/example.svg
    # TITLE:Assistant assessor
    # URL:http://example.org/example-department/example-sub-department/example-of
    #  fice
    # EMAIL;TYPE=INTERNET;TYPE=PREF;TYPE=WORK:jdoe+work-related@example.org
    # TEL;TYPE=VOICE;TYPE=PREF;TYPE=MSG;TYPE=WORK:+1234567890
    # TEL;TYPE=FAX;TYPE=WORK:+0123456789
    # ADR:Post office box 1;West wing;1 Some Street;Tinytown;Schwarzwald;12345;Fa
    #  ntasia
    # LABEL:Pb 1\n1 Some Street\n12345 Tinytown\nFantasia
    # PHOTO;VALUE=uri:http://example.org/jdoe.jpg
    # NOTE:Who is this guy anyway? He looks like a double agent from the eighties
    #  \, talks like a loon\, and reads comics when he should be working.
    # TZ:+10:30
    # GEO:1.2345;6.7890
    # CATEGORIES:Worker drones,Finance\, marketing and bull****ing
    # END:VCARD
    #

    vcf_output_dir = easygui.diropenbox('Please select the dir to output the VCF files to', title)

    vcf_contents = ''
    # for contact in contacts:
    #     c = vobject.vCard()

    #     c.add('n')
    #     c.n.value = vobject.vcard.Name( family=contacts[contact].last_name, given=contacts[contact].first_name )

    #     c.add('fn')
    #     c.fn.value = contact

    #     c.add('nickname')
    #     c.nickname.value = contacts[contact].nickname

    #     c.add('email').value = contacts[contact].email_1
    #     c.add('email').value = contacts[contact].email_2
    #     #c.email.type_param = 'internet'
    #     #c.email.value = contacts[contact].email_1

    #     #c.add('email')
    #     #c.email.type_param = 'internet'
    #     #c.email.value = contacts[contact].email_2

    #     c.add('tel')
    #     c.tel.type_param = 'cell'
    #     c.tel.value = contacts[contact].mobile_phone
    #     #print(contacts[contact].mobile_phone)

    #     c.add('tel')
    #     c.type_param = 'home'
    #     c.value = contacts[contact].home_phone

    #     c.add('tel')
    #     c.type_param = 'work'
    #     c.value = contacts[contact].work_phone

    #     c.add('fax')
    #     c.type_param = 'fax'
    #     c.value = contacts[contact].fax

    #     c.serialize()

    #     if vcf_output_files == 'Single File':
    #         vcf_contents += str(c)
    #         vcf_contents += "\n"
    #     else:
    #         print(contact)
    #         f = open(os.path.join(vcf_output_dir, str(contact + '.vcf')), 'w')
    #         f.write(str(c))
    #         f.close()

    if vcf_output_files == 'Single File':
        print(vcf_contents)
        f = open(os.path.join(vcf_output_dir, 'converted_contacts_export.vcf'), 'w')
        f.write(vcf_contents)
        f.close()
    source_file = easygui.fileopenbox("Please select the SOURCE LDIF file",title,"")
    with open(source_file, newline=None) as ldif_file:
        



def csv_output(csv_type):
    pass
    csv_contents = ''
    #for contact in contacts:
    #for i in range(0, (field_positions[csv_type][-1] + 1)):
    #    if field_postions[csv_type][i]

    
def main_menu():
    global vcf_output_files
    select_source_dest()

main_menu()

"""
 TODO:
 =====
- make sure that each easygui dialog will return to main_menu on cancel
- get VCF working reasonably
- add LDIF support
- maybe add PST support
- integrate a header check to warn if csv doesn't match selected source type

"""
