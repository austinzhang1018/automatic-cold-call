# STUDENT INFO STORED AS TUPLE- NAME, EMAIL, SECTION
# EMAIL DATA STORED AS TUPLE- UID, EMAIL, SUBJECT
# LOG HISTORY STORED AS NESTED DICT WITH EMAIL AND SKIPS AND SWITCHES
from datetime import date
import os
import imaplib
import email
import csv
import random
import pickle
from socket import gaierror

ORG_EMAIL   = "@gmail.com"
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT   = 993

def initialize_imap():
    mail = imaplib.IMAP4_SSL(SMTP_SERVER)
    mail.login(FROM_EMAIL, FROM_PWD)
    return mail

def disconnect_imap(mail):
    mail.expunge()
    mail.close()

def get_use_data():
    num_uses = dict()
    try:
        with open('roster.csv') as csv_file:
            reader = csv.DictReader(csv_file)
            for row in reader:
                skips = 0
                switches = 0
                try:
                    if len(row['email'].strip()) == 0:
                        raise Exception('An email is empty please check the csv file')
                    if len(row['skips'].strip()) != 0 and int(row['skips'].strip()) != 0:
                        skips = int(row['skips'].strip())
                    if len(row['switches'].strip()) != 0 and int(row['switches'].strip()) != 0:
                        switches = int(row['switches'].strip())
                    if skips + switches != 0:
                        num_uses[row['email'].strip().lower()] = { 'skips': skips, 'switches': switches }
                except KeyError:
                    raise Exception('Make sure there are columns named name, email, section, skips, and switches')
    except FileNotFoundError:
        raise Exception('Could not find course roster. Please place csv called roster.csv within the same folder as this program')
    return num_uses

def save_uses_to_csv(num_uses, students):
    with open('roster.csv', 'w', newline='') as f:
        fieldnames = ['name', 'email', 'section', 'skips', 'switches']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for student in students:
            written = False
            for email in num_uses:
                if email == student[1]:
                    writer.writerow({
                        'name': student[0],
                        'email': student[1],
                        'section': student[2],
                        'skips': num_uses[email]['skips'],
                        'switches': num_uses[email]['switches']
                    })
                    written = True
                    break
            if not written:
                writer.writerow({
                    'name': student[0],
                    'email': student[1],
                    'section': student[2],
                    'skips': '',
                    'switches': ''
                })

# returns an array with an int representing the type of email, the sender, and the subject respectively
# 0 represents spam mail sent from outside dartmouth
# 1 represents a skip request
# 2 represents a switch request
# 3 represents a mailing sent from dartmouth but not formatted for skip or switch
#   this can occur when a subject line is formatted incorrectly
def filter(mail, uid):
    _, byte_msg = mail.uid('fetch', uid, '(RFC822)')
    msg = email.message_from_bytes(byte_msg[0][1])
    # lowercase and remove all whitespace for formatting discrepancies
    unparsed_subject = msg['subject']
    email_subject = ''.join(unparsed_subject.lower().split())
    email_from = msg['from'].lower().strip()
    
    if '@dartmouth.edu' not in email_from.lower():
        return [0, email_from, unparsed_subject] # spam
    elif email_subject == 'econ26skip':
        return [1, email_from, unparsed_subject]
    elif email_subject == 'econ26switch':
        return [2, email_from, unparsed_subject]
    else:
        return [3, email_from, unparsed_subject]

def move_email(mail, uid):
    result = mail.uid('COPY', uid, 'processed')
    if result[0] == 'OK':
        mail.uid('STORE', uid , '+FLAGS', '(\\Deleted)')

def sort_prompt(mail_info):
    while True:
        print('From: ' + mail_info[1] + ' with subject: ' + mail_info[2])
        response = input().lower().strip()
        if (response == 'delete'):
            return 0
        elif (response == 'skip'):
            return 1
        elif (response == 'switch'):
            return 2
        else:
            print('Unrecognized command ' + response + '. Please type delete, skip, or switch')

# prompts the user to classify an email of unknown type
# returns 0 if classified as spam, 1 if skip, 2 if switch
def sort_unknowns(mail, skips, switches, unknowns):
    if len(unknowns) > 0:
        print('The following emails were sent from dartmouth but could not be classified. Please classify each email type by typing delete, skip, or switch')
    
    for unknown in unknowns:
        mail_type = sort_prompt(unknown)
        if mail_type == 0:
            # spam
            mail.uid('STORE', unknown[0], '+FLAGS', '\\Deleted')
        elif mail_type == 1:
            skips.append(unknown)
        else:
            switches.append(unknown)

def read_emails(mail):
    # select the inbox
    mail.select('inbox', readonly=False)
    # pull all emails from the inbox
    _, data = mail.uid('search', None, 'ALL')

    skips = list()
    switches = list()
    unknowns = list()

    for i in data[0].split():
        response = filter(mail, i)
        mail_type = response[0]
        # throw out mail type we don't need it
        mail_info = i, email.utils.parseaddr(response[1])[1], response[2] # uid, sender, and subject
        if mail_type == 0: # spam
            # delete email
            mail.uid('STORE', i, '+FLAGS', '\\Deleted')
        elif mail_type == 1:
            # skip email
            skips.append(mail_info)
        elif mail_type == 2:
            # switch email
            switches.append(mail_info)
        else:
            # unknown email from dartmouth
            unknowns.append(mail_info)
    sort_unknowns(mail, skips, switches, unknowns)
    return skips, switches

# returns a list of all students in the course with name, email, and section
def get_course_roster():
    students = list()
    try:
        with open('roster.csv') as csv_file:
            reader = csv.DictReader(csv_file)
            for row in reader:
                try:
                    if len(row['name'].strip()) == 0:
                        raise Exception('A name is empty please check the csv file')
                    if len(row['email'].strip()) == 0:
                        raise Exception('An email is empty please check the csv file')
                    if len(row['section'].strip()) == 0:
                        raise Exception('A section is empty please check the csv file')
                except KeyError:
                    raise Exception('Make sure there are columns named name, email, and section')
                students.append((row['name'].strip(), row['email'].strip().lower(), row['section'].strip().lower()))
    except FileNotFoundError:
        raise Exception('Could not find course roster. Please place csv called roster.csv within the same folder as this program')
    return students

# returns a list of cached switches
# this is needed because if someone from a later section switches into an earlier section
# we need to make sure that they don't get added to the later section
def get_request_cache(num_uses, students, skips, switches):
    try:
        with open('request_cache.pickle', 'rb') as f:
            skip_cache, switch_cache, cache_date = pickle.load(f)
            if date.today() == cache_date:
                return skip_cache, switch_cache
            else:
                save_use_data(num_uses, students, skip_cache, switch_cache)
        os.remove('request_cache.pickle')
        return list(), list()
    except FileNotFoundError:
        # don't have any saved logs
        return list(), list()

def save_request_cache(skips, switches):        # Pickle the 'data' dictionary using the highest protocol available.
    if len(skips) == 0 and len(switches) == 0:
        return
    with open('request_cache.pickle', 'wb') as f:
        pickle.dump((skips, switches, date.today()), f, pickle.HIGHEST_PROTOCOL)
        
def save_use_data(num_uses, students, skips, switches):
    # move the skips from our section
    for student in students:
        for skipper in skips:
            if skipper[1] == student[1]:
                try: 
                    num_uses[student[1]]['skips'] += 1
                except KeyError:
                    # no logs for this student yet
                    num_uses[student[1]] = { 'skips': 1, 'switches': 0 }
                break
    # move the switches from the other section
    for student in students:
        for switcher in switches:
            if switcher[1] == student[1]:
                try:
                    num_uses[student[1]]['switches'] += 1
                except KeyError:
                    # no logs for this student yet
                    num_uses[student[1]] = { 'skips': 0, 'switches': 1 }
                break
    save_uses_to_csv(num_uses, students)

# returns a new roster without the students who have used their skips
def apply_skips(roster, skips, num_uses, section):
    new_roster = []
    # not efficient but doesn't matter for our use case
    for student in roster:
        # if the student isn't in our section we don't care, add to roster and continue
        # we'll process/filter them in apply_switches
        if student[2] != section:
            new_roster.append(student)
            continue

        skipped = False
        for skip in skips:
            # if the student doesn't match the current step don't process the skip
            if student[1] != skip[1]:
                continue

            # make sure that student hasn't exhausted skips
            try:
                num_uses[student[1]]['skips']
                if num_uses[student[1]]['skips'] >= 5:
                    print(student[0] + ' attempted to use a skip but they have used all 5')
                else:
                    print(student[0] + ' has now used ' + str(num_uses[student[1]]['skips'] + 1) + ' skips')
                    skipped = True
            except KeyError:
                # this student hasn't used any skips or switches yet
                print(student[0] + ' has now used 1 skip.')
                skipped = True
            # break so we don't process multiple skips for the same student
            break

        if not skipped:
            new_roster.append(student)        

    return new_roster

def get_sections(roster):
    sections = set()
    for student in roster:
        sections.add(student[2])
    return sections

def prompt_sections(sections):
    while True:
        prompt = 'Please type both or select one section by typing its name:'
        for section in sorted(sections):
            prompt += ' ' + section
        print(prompt)
        selected_section = input().lower().strip()

        if selected_section == 'both':
            return tuple(sorted(sections))
        
        if selected_section in sections:
            print('Section ' + selected_section + ' selected')
            return selected_section, False

# ASSUMES THERE'S ONLY TWO SECTIONS, ONLY PART OF CODE THAT MAKES THIS ASSUMPTION
# filters all of the students who should be in our section
# this is students from our section who have not requested switches
# also students from another section who have requested a switch
def apply_switches(roster, switches, num_uses, section):
    new_roster = []
    for student in roster:
        switched = False
        for switcher in switches:
            # continue, switcher doesn't match
            if student[1] != switcher[1]:
                continue
            # if the student is in our section and switching, mark them and break
            switched = True
            if student[2] == section:
                break
            try:
                print(student[0] + ' has now switched ' + str(num_uses[student[1]]['switches'] + 1) + ' times')
            except KeyError: # this occurs if the student hasn't switched or skipped
                print(student[0] + ' has now switched 1 time')
            break
        if (student[2] == section and not switched) or (student[2] != section and switched):
            new_roster.append(student)
    return new_roster
        
def prompt_action():
    print()
    print('The cold call list has been downloaded.')
    print()
    print('Type again to download and select another section')
    print('Type exit to quit')
    while True:
        command = input().lower().strip()
        if command == 'again':
            return
        elif command == 'exit':
            exit()
        else:
            print('Unrecognized command. Valid commands are save, reset, return, and exit.')

def write_list(call_list, section):
    with open('call_list_' + section + '.csv', 'w', newline='') as f:
        fieldnames = ['name']
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        for student in call_list:
            writer.writerow({'name': student[0]})

def move_processed_emails(mail, students, skips, switches, section):
    # move the skips from our section
    for skipper in skips:
        for student in students:
            if skipper[1] == student[1] and student[2] == section:
                move_email(mail, skipper[0])
                break
    # move the switches from the other section
    for switcher in switches:
        for student in students:
            if switcher[1] == student[1] and student[2] != section:
                move_email(mail, switcher[0])
                break

def combine_cache(request, request_cache):
    combined = set()
    combined.update(request)
    combined.update(request_cache)
    return combined


def main(given_section=None):
    # login to email
    mail = initialize_imap()
    skips, switches = read_emails(mail)
    students = get_course_roster()
    call_list = list.copy(students)
    num_uses = get_use_data()
    skip_cache, switch_cache = get_request_cache(num_uses, students, skips, switches)
    skips = combine_cache(skips, skip_cache)
    switches = combine_cache(switches, switch_cache)
    section = given_section
    both = False
    if not section:
        section, both = prompt_sections(get_sections(call_list))
    call_list = apply_skips(call_list, skips, num_uses, section)
    call_list = apply_switches(call_list, switches, num_uses, section)
    # call list should now only be students who are valid cold call candidates
    random.shuffle(call_list)
    # write out csv regardless of action
    write_list(call_list, section)
    print('The cold call list has been downloaded for section ' + section + '.')
    save_request_cache(skips, switches)
    move_processed_emails(mail, students, skips, switches, section)
    return both
    


if __name__ == '__main__':
    try:
        section = None
        while True:
            section = main(section)
            if not section:
                exit()
                
    except gaierror:
        print('Cannot connect to internet. Please check your connection and try again.')