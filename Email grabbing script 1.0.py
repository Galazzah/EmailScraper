import xlsxwriter #write to excel
import urllib.request #scrape from web
import requests #scrape from web
import io #scrape from web
import traceback
from collections import namedtuple #create namedtuple for manipulation
import re #test source code for email
professor = namedtuple('professor', 'name university')

PROFESSOR_LIST = 'samir.txt'
def make_professor_list()->[professor]:
    ''' Reads the file and creates a list of professor objects
    '''
    infile = open(PROFESSOR_LIST, 'r')
    lines = infile.readlines()
    name = []
    for line in lines:
        line_tester = line.replace('\n', '').split(', ')
        if len(line_tester) > 1:
            name.append(professor(line_tester[0], line_tester[1]))
    infile.close
    return name


def make_link(professor) -> str:
    '''Creates a link to search for email, uses google im feeling lucky
    '''
    var = requests.get(r'http://www.google.ca/search?q={} {}'.format(professor.name, professor.university)+'&btnI')
    return var.url
def make_link2(professor) -> str:
    var = requests.get(r'http://www.google.ca/search?q={} {} {}'.format(professor.name, professor.university, 'email') + '&btnI')
    return var.url

def get_source_from_link(link:str)->str:
    '''Gets source code from first link result in search 
    '''
    u = urllib.request.urlopen(link, data = None)
    f = io.TextIOWrapper(u,encoding='utf-8')
    text = f.read()
    return text

def find_email_list(source: str) -> [str]:
    '''Returns a list of emails from source code
    '''
    return re.findall(r'[\w\-][\w\-\.]+@[\w\-][\w\-\.]+[a-zA-Z]{1,4}', source)


def find_email_in_list(email_list:[str], professor:professor) ->str:
    '''Finds the email of the professor specified
    '''
    for email in email_list:
        if professor.name.lower().split()[1] in email.lower() or professor.name.lower().split()[0] in email.lower() or (len(professor.name.split()) >=3 and professor.name.lower().split()[2] in email.lower()) or (len(email_list) == 1 and '.edu' in email):
            return email
def print_professor_info_list(professor_list:[professor]):
    '''Prints Profesor info in an easily readble format
    '''
    for professor in professor_list:
        print(professor.name + ', ' + professor.university)

def main():
    print('V2')
    prof_emails = {}
    prof_email_found = []
    no_email_connection = []
    no_email_on_page = []
    prof_list = make_professor_list()
    #WARNING: ENTER THE NUMBER OF PROFESSORS TO CHECK BELOW, OR YOUR COMPUTER WILL CRASH
    counter = 0
    for prof in prof_list:
        prof_emails[prof.name] = None
        linky = make_link(prof)
        try:
            if 'academia.edu' in linky:
                del prof_emails[prof.name]
                continue
            sourcey = get_source_from_link(linky)
            
            emaily_listy = find_email_list(sourcey)
            
            emaily = find_email_in_list(emaily_listy, prof)
            
            print(emaily)
            
            prof_emails[prof.name] = emaily
        except (urllib.error.HTTPError, UnicodeDecodeError, requests.exceptions.SSLError):
            del prof_emails[prof.name]
            no_email_connection.append(prof)
            continue
        if prof_emails[prof.name] == None:
            try:
                print('Checking second URL')
                prof_emails[prof.name] = None
                linky2 = make_link2(prof)
                if 'academia.edu' in linky2:
                    del prof_emails[prof.name]
                    continue
                sourcey2 = get_source_from_link(linky2)
                emaily_listy2 = find_email_list(sourcey2)
                emaily2 = find_email_in_list(emaily_listy2, prof)
                print(emaily2)
                prof_emails[prof.name] = emaily2
            except (urllib.error.HTTPError, UnicodeDecodeError, requests.exceptions.SSLError):
                del prof_emails[prof.name]
                no_email_connection.append(prof)
                continue
            if prof_emails[prof.name] == None:
                no_email_on_page.append(prof)
                del prof_emails[prof.name]
                continue
                
        prof_email_found.append(professor(prof.name, prof.university))
    email_list_workbook = xlsxwriter.Workbook('email_list1.1trial1.xlsx')
    worksheet = email_list_workbook.add_worksheet()
    
    row = 0
    for prof in prof_email_found:
        worksheet.write(row, 0, prof.name)
        worksheet.write(row,1, prof.university)
        worksheet.write(row,2, prof_emails[prof.name])
        row +=1
    email_list_workbook.close()
    no_email_list_workbook = xlsxwriter.Workbook('no_email_list1.1trial1.xlsx')
    worksheet2 = no_email_list_workbook.add_worksheet()
    row = 0
    for prof in no_email_connection:
        worksheet2.write(row, 0, prof.name)
        worksheet2.write(row, 1, prof.university)
        row+=1
    for prof in no_email_on_page:
        worksheet2.write(row, 0, prof.name)
        worksheet2.write(row, 1, prof.university)
        row +=1
    no_email_list_workbook.close()
    avg = (len(no_email_connection) + len(no_email_on_page) + len(prof_emails))
    avg2 = len(no_email_connection) + len(no_email_on_page)
    print('Emails found for: ')
    print(prof_emails)
    print('-'*35)
    print('No email found due to failed connection: ')
    print_professor_info_list(no_email_connection)
    print('-'*35)
    print('No email found on page: ')
    print_professor_info_list(no_email_on_page)
    print('-'*35)
    print('STATS:')
    print('Number of emails found: ', len(prof_emails))
    print('Number of connection errors: ', len(no_email_connection))
    print('Number of emails not on page: ', len(no_email_on_page))
    print('Emails not found: ', (len(no_email_connection) + len(no_email_on_page)))
    print('Return rate: %',(len(prof_emails)/avg*100))
    
if __name__ == '__main__':
    main()
    
        
