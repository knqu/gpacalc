import csv
from rich.console import Console
from rich.theme import Theme
from rich.table import Table

semantic = Theme({
    'prompt': '#10a4e8 bold',
    'info': '#edae1a bold',
    'error': '#e82610',
    'success': '#10e83f',
    'dim': '#404040'
})

console = Console(theme=semantic)

def hr():
    console.print('')
    console.rule(style='dim')
    console.print('')

record = []

def calculateGPA(weight):
    addGPA = 0
    if weight == 'w': # weighted
        for el in record:
            addGPA += el['gpa']
    elif weight == 'u': # unweighted
        for el in record:
            if el['grade'] == 'a':
                addGPA += 4.0
            elif el['grade'] == 'a-':
                addGPA += 3.667
            elif el['grade'] == 'b+':
                addGPA += 3.333
            elif el['grade'] == 'b':
                addGPA += 3.0
            elif el['grade'] == 'b-':
                addGPA += 2.667
            elif el['grade'] == 'c+':
                addGPA += 2.333
            elif el['grade'] == 'c':
                addGPA += 2.0
            elif el['grade'] == 'c-':
                addGPA += 1.667
            elif el['grade'] == 'd+':
                addGPA += 1.333
            elif el['grade'] == 'd':
                addGPA += 1.0
            elif el['grade'] == 'd-':
                addGPA += 0.667
            elif el['grade'] == 'f':
                addGPA += 0.0
    return round(addGPA / len(record), 3)

def init():
    hr()
    console.print('gpacalc', style='info')
    console.print('Created by Kevin Qu', style='info')
    console.input('\npress [ enter ] to begin > ')
    menu()

def menu():
    hr()
    console.print('What would you like to do?\n', style='prompt')
    console.print('[ n ] create new record', style='prompt')
    console.print('[ v ] view current record', style='prompt')
    console.print('[ e ] edit current record', style='prompt')
    console.print('[ c ] export record as .csv', style='prompt')
    console.print('[ h ] help', style='prompt')
    console.print('[ x ] exit', style='prompt')
    inpt = console.input('\n> ')

    if inpt == 'n':
        new()
    elif inpt == 'v':
        view()
    elif inpt == 'e':
        edit()
    elif inpt == 'c':
        export()
    elif inpt == 'h':
        hr()
        console.print('[ c ] create new record - creates a blank record and prompts you for element values')
        console.print('[ v ] view current record - view current record')
        console.print('[ e ] edit current record - change elements of current record')
        console.print('[ n ] export record as .csv - export current record as a .csv (spreadsheet) file')
        console.input('\npress [ enter ] to continue > ')
    elif inpt == 'x':
        console.print('Thank you for using gpacalc. Exiting program.', style='info')
        exit()
    else:
        console.print('\n[ERROR] Please enter a valid option.', style='error')
        console.input('press [ enter ] to continue > ')

    menu()

def new():
    global record

    def create():
        global record

        inpt = ''
        while inpt != 'done':
            console.print('\nPlease enter a class | letter grade | weight', style='prompt')
            inpt = console.input('> ')

            if inpt != 'done':
                try:
                    inpt_split = inpt.split(' | ')
                    inpt_split[1] = inpt_split[1].lower()
                    inpt_split[2] = inpt_split[2].lower()

                    if inpt_split[0] and len(inpt_split[0]) <= 24 and inpt_split[1] in ['a', 'a-', 'b+', 'b', 'b-', 'c+', 'c', 'c-', 'd+', 'd', 'd-', 'f'] and inpt_split[2] in ['r', 'p', 'f'] and len(inpt_split) <= 3:
                        gpa = 0
                        if inpt_split[1] == 'a' and inpt_split[2] == 'f':
                            gpa = 5.0
                        elif inpt_split[1] == 'a-' and inpt_split[2] == 'f':
                            gpa = 4.667
                        elif inpt_split[1] == 'b+' and inpt_split[2] == 'f':
                            gpa = 4.333
                        elif inpt_split[1] == 'b' and inpt_split[2] == 'f':
                            gpa = 4.0
                        elif inpt_split[1] == 'b-' and inpt_split[2] == 'f':
                            gpa = 3.667
                        elif inpt_split[1] == 'c+' and inpt_split[2] == 'f':
                            gpa = 3.333
                        elif inpt_split[1] == 'c' and inpt_split[2] == 'f':
                            gpa = 3.0
                        elif inpt_split[1] == 'c-' and inpt_split[2] == 'f':
                            gpa = 2.667
                        elif inpt_split[1] == 'd+' and inpt_split[2] == 'f':
                            gpa = 2.333
                        elif inpt_split[1] == 'd' and inpt_split[2] == 'f':
                            gpa = 2.0
                        elif inpt_split[1] == 'd-' and inpt_split[2] == 'f':
                            gpa = 1.667
                        elif inpt_split[1] == 'f' and inpt_split[2] == 'f':
                            gpa = 0.0
                        elif inpt_split[1] == 'a' and inpt_split[2] == 'p':
                            gpa = 4.5
                        elif inpt_split[1] == 'a-' and inpt_split[2] == 'p':
                            gpa = 4.167
                        elif inpt_split[1] == 'b+' and inpt_split[2] == 'p':
                            gpa = 3.833
                        elif inpt_split[1] == 'b' and inpt_split[2] == 'p':
                            gpa = 3.5
                        elif inpt_split[1] == 'b-' and inpt_split[2] == 'p':
                            gpa = 3.167
                        elif inpt_split[1] == 'c+' and inpt_split[2] == 'p':
                            gpa = 2.833
                        elif inpt_split[1] == 'c' and inpt_split[2] == 'p':
                            gpa = 2.5
                        elif inpt_split[1] == 'c-' and inpt_split[2] == 'p':
                            gpa = 2.167
                        elif inpt_split[1] == 'd+' and inpt_split[2] == 'p':
                            gpa = 1.833
                        elif inpt_split[1] == 'd' and inpt_split[2] == 'p':
                            gpa = 1.5
                        elif inpt_split[1] == 'd-' and inpt_split[2] == 'p':
                            gpa = 1.167
                        elif inpt_split[1] == 'f' and inpt_split[2] == 'p':
                            gpa = 0.0
                        elif inpt_split[1] == 'a' and inpt_split[2] == 'r':
                            gpa = 4.0
                        elif inpt_split[1] == 'a-' and inpt_split[2] == 'r':
                            gpa = 3.667
                        elif inpt_split[1] == 'b+' and inpt_split[2] == 'r':
                            gpa = 3.333
                        elif inpt_split[1] == 'b' and inpt_split[2] == 'r':
                            gpa = 3.0
                        elif inpt_split[1] == 'b-' and inpt_split[2] == 'r':
                            gpa = 2.667
                        elif inpt_split[1] == 'c+' and inpt_split[2] == 'r':
                            gpa = 2.333
                        elif inpt_split[1] == 'c' and inpt_split[2] == 'r':
                            gpa = 2.0
                        elif inpt_split[1] == 'c-' and inpt_split[2] == 'r':
                            gpa = 1.667
                        elif inpt_split[1] == 'd+' and inpt_split[2] == 'r':
                            gpa = 1.333
                        elif inpt_split[1] == 'd' and inpt_split[2] == 'r':
                            gpa = 1.0
                        elif inpt_split[1] == 'd-' and inpt_split[2] == 'r':
                            gpa = 0.667
                        elif inpt_split[1] == 'f' and inpt_split[2] == 'r':
                            gpa = 0.0
                        record.append({'class': inpt_split[0], 'grade': inpt_split[1], 'weight': inpt_split[2], 'gpa': gpa})
                        console.print('\nClass added to record', style='success')
                    else:
                        console.print('\n[ERROR] Invalid input syntax. Please reference the tutorial if needed.', style='error')
                        console.input('press [ enter ] to continue > ')
                except IndexError:
                    console.print('\n[ERROR] Invalid input syntax. Please reference the tutorial if needed.', style='error')
                    console.input('press [ enter ] to continue > ')

        console.print('\nNew record saved.', style='success')
        console.input('press [ enter ] to continue > ')

    hr()

    console.print('Would you like to view the tutorial?\n', style='prompt')
    console.print('[ y ] yes', style='prompt')
    console.print('[ n ] no, i know what to do already', style='prompt')
    inpt = console.input('\n> ')

    if inpt == 'y':
        console.print('\nEnter the names of classes in the following format: class | letter grade | weight')
        console.print('The class name can be anything 24 characters or under', highlight=False)
        console.print('The letter grade is your letter grade in the class (pluses and minuses are supported, but a+ is not)')
        console.print('The weight can be entered in as either \'r\' (regular), \'p\' (partial), or \'f\' (full)', highlight=False)
        console.print('Example: computer science | a- | f')
        console.print('At any time, enter \'done\' to finish and save the classes', highlight=False)
        console.input('\npress [ enter ] to continue > ')
        create()
    elif inpt == 'n':
        create()
    else:
        console.print('\n[ERROR] Please enter a valid option.', style='error')
        console.input('press [ enter ] to continue > ')
        new()

def view():
    hr()

    if len(record) > 0:
        table = Table(show_header=True)
        table.add_column('CLASS', width=24)
        table.add_column('LG', width=2)
        table.add_column('W', width=1)
        table.add_column('GPA', width=5)

        for el in record:
            table.add_row(f'{el["class"]}', f'{el["grade"].upper()}', f'{el["weight"].upper()}', f'{el["gpa"]}')

        console.print(table)

        console.print(f'Cumulative Weighted GPA: {calculateGPA("w")}')
        console.print(f'Cumulative Unweighted GPA: {calculateGPA("u")}')
    else:
        console.print('Record is empty. Try adding classes to it by entering [ n ] in the menu.', style='error')

    console.input('\npress [ enter ] to continue > ')

def edit():
    global record

    hr()
    console.print('What class would you like to edit/ delete (note that this is case-sensitive)? (enter [ x ] to cancel)', style='prompt')
    inpt = console.input('> ')

    cnt = 0
    for el in record:
        if el['class'] == inpt:
            cnt += 1

    if cnt > 0:
        console.print(f'\nWould value of {inpt} would you like to edit?\n', style='prompt')
        console.print('[ 1 ] class name', style='prompt')
        console.print('[ 2 ] letter grade', style='prompt')
        console.print('[ 3 ] class weight', style='prompt')
        console.print('[ 4 ] delete class', style='prompt')
        inpt = console.input('> ')

        if inpt == '1':
            if cnt > 1:
                console.print('\nThere are multiple occurances of the specified class. Would you like to edit all of them or just the first occurance?\n', style='prompt')
                console.print('[ 1 ] edit only the first occurance', style='prompt')
                console.print('[ 2 ] edit all occurances', style='prompt')
                inpt = console.input('\n> ')

                if inpt == '1':
                    console.print('\nWhat would you like to change')
            else:
                idx = record.index(inpt)
                record[idx] == inpt
    elif inpt == 'x':
        menu()
    else:
        console.print('\n[ERROR] No matching class name. Please try again.', style='error')
        console.input('\npress [ enter ] to continue > ')
        edit()

def export():
    hr()
    console.print('Would you like to convert your current record to a .csv (comma seperated values) file?', style='prompt')
    console.print('Note that to make the most use out of this file, you should open it with a spreadsheet-viewing tool such as Google Sheets or Microsoft Excel.\n', style='prompt')

    console.print('[ y ] yes', style='prompt')
    console.print('[ n ] no', style='prompt')

    inpt = console.input('\n> ')

    if inpt == 'y' and len(record) > 0:
        pass
    elif inpt == 'y' and len(record) == 0:
        console.print('\n[ERROR] You must have at least one element in your record to export it as .csv', style='error')
        console.input('press [ enter ] to continue > ')
    elif inpt == 'n':
        menu()
    else:
        console.print('\n[ERROR] Please enter a valid option.', style='error')
        console.input('press [ enter ] to continue > ')

init()
