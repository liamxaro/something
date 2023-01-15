from numpy import NaN
import pandas as pd
import time
import xlsxwriter
import xlrd
import re
import os
from pathlib import Path

def gather_input(allEmails, txtFilePath):
    '''
    if you have a .co email next to a word that starts with m you will get an error
    '''
    topLevelDomains = ['.mil', '.org', '.gov', '.com', '.net', '.info', '.ca', '.edu', '.bah', '.us']
    if len(allEmails) == 0 or allEmails.count('@') >= 2:
        if os.path.exists(txtFilePath):
            allEmails = ""
            with open(txtFilePath, 'r') as file:
                lines = file.readlines()
               
            
            for i in range(len(lines)):
                if lines[i].count('@') >= 2:
                    for x in topLevelDomains:
                        temp = find_indices_of_substring(lines[i], x)
                        adjust = [i for i in range(len(temp))]
                        finalPos = [sum(i) for i in zip(temp,adjust)]
                        for j in range(len(finalPos)):
                            lines[i] = insert_homemade(lines[i], " ", finalPos[j] + len(x))

                # elif lines[i].count('@') == 1: #I dont think this is needed anymore but I'm keeping it just in case.
                #     lines[i] = lines[i] + " "
           
            for i in range(len(lines)):
                lines[i] = lines[i].strip()
                lines[i] = lines[i].split(" ")

            finalLines = []
            for i in range(len(lines)):
                finalLines.extend(lines[i])

            for i in range(len(finalLines)):
                allEmails += " " + finalLines[i] + " "


            return allEmails

def insert_homemade(source_str, insert_str, pos):
    return source_str[:pos] + insert_str + source_str[pos:]

def find_indices_of_substring(full_string, sub_string):
    return [index for index in range(len(full_string)) if full_string.startswith(sub_string, index)]

def extract(allEmails):
    '''
    This function takes a string of emails and puts them all into a list

    returns a list
    '''
    extractedEmails = re.findall(r'\S+@\S+', allEmails)

    return extractedEmails

def clean(dirtyList, badChars):
    '''
    This function takes a list removes all duplicate values and sorts it alphabatically

    returns list
    '''

    cleanList = list(set(dirtyList))
    cleanList = (x.lower() for x in cleanList)
    cleanList = list(cleanList)
    cleanList = sorted(cleanList, reverse = False)

    for i in range(len(cleanList)):
        for x in badChars:
            if x in cleanList[i]:
                cleanList[i] = cleanList[i].replace(x, "")

    return cleanList

def domain_detection(cleanedEmails):
    '''
    This function takes a list (of emails that are cleaned) and sorts them based
    on domain into a dictionary using the domain as the key and the whole email as a value

    example of output:
        {'All' :[all the emails], 'Gmail': ['someemail@gmail.com', 'anotherone@gmail.com'],
            'Comcast': ['something@comcast.net' , 'somethingelse@comcast.net']}

    returns type dictionary
    '''
    emailDomains = {'All': []}

    for i in range(len(cleanedEmails)):
        emailDomains['All'].append(cleanedEmails[i])

        domainFinder = cleanedEmails[i].find('@', 2) # come back to this
        domain = cleanedEmails[i][domainFinder + 1:]

        if domain.capitalize() in emailDomains:
            emailDomains[domain.capitalize()].append(cleanedEmails[i])
        else:
            emailDomains[domain.capitalize()] = []
            emailDomains[domain.capitalize()].append(cleanedEmails[i])

    #print(emailDomains)
    return emailDomains

def compile(detectedDomains):
    '''
    this function takes the given dictionary and converts it into 
    a pandas dataframe

    returns type pandas dataframe
    '''
    email_df = pd.DataFrame({key:pd.Series(value) for key,value in detectedDomains.items()})
    return email_df

def get_longest_email(detectedDomains):
    '''
    This function takes the dictionary and finds the longest email in it and returns a list of the longest emails
    '''
    longestEmails = []
    for x in detectedDomains:
        if x != 'All':
            longEmail = max(detectedDomains[x], key = len)
            longestEmails.append(longEmail)
    
    return longestEmails

def write_the_file(excelFilePath, df, detectedDomains):
    writer = pd.ExcelWriter(excelFilePath, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', startrow=1, index = False, header = False)
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']

    longestEmails = get_longest_email(detectedDomains)

    for i in range(len(longestEmails)):
        worksheet.set_column(i, i, (len(longestEmails[i]) + 1))
        
    column_settings = [{'header':column} for column in df.columns]
    (max_row, max_col) = df.shape

    worksheet.add_table(0,0,max_row, max_col - 1, {'columns': column_settings})

    writer.save()

def write_file_process(excelFilePath, email_df, detectedDomains, cleanedEmails, badChars):
    '''
    This function writes a dataframe to an excel file

    does not return anything
    '''
    if os.path.exists(excelFilePath):
        existing_df = pd.read_excel(excelFilePath, sheet_name='Sheet1')
        print(f'\n' + 'Dataframe already on spreadsheet: ' + '\n'  + '\n ' + str(existing_df))

        if existing_df.empty == True:
            flag = True
            while flag:
                answer = input("Would you like to add the new data to the existing file? (Y/N): ")
                answer = answer.lower()
                answer = re.sub('\s+', '', answer)
                if (answer == 'y' or answer == 'yes'):

                    write_the_file(excelFilePath, email_df, detectedDomains)

                    print(f'Results written to excel file at: {excelFilePath}')
                    print('Program has ended')

                    flag = False
                elif (answer == 'n' or answer == 'no'):
                    print('You elected to not write the results to an excel file.')
                    print('Program has ended')
                    flag = False


        elif existing_df.empty == False:
            flag = True
            while flag:
                answer = input("Would you like to add the new data to the existing file? (Y/N): ")
                answer = answer.lower()
                answer = re.sub('\s+', '', answer)
                if (answer == 'y' or answer == 'yes'):
                    existing_dict = existing_df.to_dict()

                    existingEmails = []
                    for key in existing_dict:
                        for key2 in existing_dict[key]:
                            existingEmails.append(existing_dict[key][key2])


                    cleanedEmails.extend(existingEmails)

                    extendedCleanEmails = clean(cleanedEmails, badChars)

                    new_dict = domain_detection(extendedCleanEmails)

                    new_df = compile(new_dict)

                    write_the_file(excelFilePath, new_df, new_dict)

                    

                    print(f'Results written to excel file at: {excelFilePath}')
                    print('Program has ended')
                    flag = False
                elif(answer == 'n' or answer == 'no'):
                    answer = input("Would you like to remove the inputted data from the existing file? (Y/N): ")
                    answer = answer.lower()
                    answer = re.sub('\s+', '', answer)
                    if (answer == 'y' or answer == 'yes'):
                        
                        existing_dict = existing_df.to_dict()

                        existingEmails = []
                        for key in existing_dict:
                            for key2 in existing_dict[key]:
                                existingEmails.append(existing_dict[key][key2])

                        newGen = (x for x in cleanedEmails if x not in existingEmails)
                        newList = list(newGen)

                        cleanedNewList = clean(newList, badChars)

                        new_dict = domain_detection(cleanedNewList)

                        new_df = compile(new_dict)


                        write_the_file(excelFilePath, new_df, new_dict)

                        print(f'Results written to excel file at: {excelFilePath}')
                        print('Program has ended')
                        flag = False

                    elif(answer == 'n' or answer == 'no'):
                        print("THE FOLLOWING ACTION CANNOT BE UNDONE.")
                        answer = input("Would you like to completely overwrite the sheet in the given file? (Y/N):")
                        answer = answer.lower()
                        answer = re.sub('\s+', '', answer)
                        if (answer == 'y' or answer == 'yes'):
                            
                            write_the_file(excelFilePath, email_df, detectedDomains)

                            print(f'Results written to excel file at: {excelFilePath}')
                            print('Program has ended')

                            flag = False
                    elif (answer == 'n' or answer == 'no'):
                        print('You elected to not write the results to an excel file.')
                        print('Program has ended')
                        flag = False


    else:
        print('Cannot write to a file that does not exist!')
        print('Create the excel file and then copy the absolute path and paste it in the variable.')
    
def find_repeated(extractedEmails):
    duplicates = []
    for x in extractedEmails:
        if extractedEmails.count(x) > 1:
            duplicates.append(x)
    
    return duplicates

def inputted_data_error_check(extractedEmails, cleanedEmails, duplicateEmails, duplicateEmailsClean, email_df, detectedDomains, badChars):
    checkOne = False
    checkTwo = False
    # print(len(cleanedEmails))
    # print(len(extractedEmails))
    # print(len(duplicateEmails))
    # print(len(duplicateEmailsClean))

    if len(extractedEmails) - (len(duplicateEmails) - len(duplicateEmailsClean)) == len(cleanedEmails):
        checkOne = True
    print(f'{len(extractedEmails)} - {(len(duplicateEmails) - len(duplicateEmailsClean))} = {len(cleanedEmails)} <----- Mathematical Statement is: {checkOne}')
    ###############################################
    
    colLengths = ""
    for x in detectedDomains:
        if x != 'All':
            colLengths += str(len(detectedDomains[x]))
            colLengths += " + "
    colLengths = colLengths[:-2]
    colLengths += "= "

    count = 0
    for x in detectedDomains:
        if x!= 'All':
            count = count + len(detectedDomains[x])
    
    if count == len(detectedDomains['All']):
        checkTwo = True
    else:
        print(f'Count was: {count}')
    print(f'{colLengths}{len(cleanedEmails)} <----- Mathematical Statement is: {checkTwo}')
    ###############################################
    doubleCheckRegex = []
    doubleCheckLen = []
    doubleCheckBadChars = []
    for x in cleanedEmails:
        if re.match(r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$", x) is None:
            doubleCheckRegex.append(x)

        if len(x) > 30:
            doubleCheckLen.append(x)
        
        for y in badChars:
            if y in x:
                doubleCheckBadChars.append(x)


    
    print(f'Double check REGEX worked properly: {doubleCheckRegex}' + '\n')
    print(f'Double check len worked properly: {doubleCheckLen}' + '\n')
    print(f'Double check clean worked properly: {doubleCheckBadChars}' + '\n')
    ###############################################
    pass



def main():
    '''
    Only variables you need to manipulate are allEmails/allEmailsFilePath/excelFilePath/
    '''
    allEmails = """
    ljaro11@comcast.netsomething@gmail.comRWAWER@us.navy.mil
ljaro11@comcast.net
jaro11@comcast.net
mearo9@gmail.com"""

    excelFilePath = Path('/Users/laro/LiamsThingsOnComputerDrive/VisualStudiosProjects/python-workspace/EmailProj/Resources/output.xlsx')
    
    txtFilePath = Path('/Users/laro/LiamsThingsOnComputerDrive/VisualStudiosProjects/python-workspace/EmailProj/Resources/input.txt') 

    badChars = ["'", ">", "<", ";", ":"]

    initialTime = time.time()

    print("*************************************************************")
    print("*                                                           *")
    print("*                  Email Excel Automation                   *")
    print("*                                                           *")
    print("*************************************************************" + '\n')
    '''
    Gather all of the emails from the text or text file
    '''
    if len(allEmails) == 0 or allEmails.count('@') >= 2:
        allEmails = gather_input(allEmails, txtFilePath)
    
    

    extractedEmails = extract(allEmails)
    print(f'Total emails found: {len(extractedEmails)}' + '\n')


    '''
    Display amount of duplicate emails and what they are
    '''
    cleanedEmails = clean(extractedEmails, badChars)
    duplicateEmails = find_repeated(extractedEmails) # the emails that have count > 1
    duplicateEmailsClean = list(set(duplicateEmails)) #the emails that occur more than once, but we need to keep one of each

    print(f'Total duplicate emails found in allEmails: {len(extractedEmails) - len(cleanedEmails)}')
    #print(f'These are the emails that are duplicated: {duplicateEmailsClean}')
    print(f'Emails remaining after duplicate removal: {len(extractedEmails) - len(duplicateEmails) + len(duplicateEmailsClean)}' + '\n')



    '''
    Sort emails by domain / department
    AND
    Turn it into Pandas Dataframe
    '''
    detectedDomains = domain_detection(cleanedEmails)
    email_df = compile(detectedDomains)

    print('Total domains: ' + str(len(detectedDomains['All'])))
    for x in detectedDomains:
        if x != 'All':
            print(f'Total amount of {x} domains: {len(detectedDomains[x])}')
    print()
    
    inputted_data_error_check(extractedEmails, cleanedEmails, duplicateEmails, duplicateEmailsClean, email_df, detectedDomains, badChars)

    print('\n')
    print('Currently inputted dataframe: ' + '\n')
    print(email_df)

    '''
    Write the results to an excel file
    '''
    write_file_process(excelFilePath, email_df, detectedDomains, cleanedEmails, badChars)

    finalTime = time.time()
    print(f'Tasks completed and took : {finalTime - initialTime}')
   
    





main()
