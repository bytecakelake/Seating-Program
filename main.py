# -*- coding: cp949 -*-

'''

Seating chart Program
sub functions:
    1.load "person name","prefer option" from xlsx file
    2.Preprocessing preference options
    3.Generate random seating chart
    4.Make sure your preferred option has the correct placeholder
    5.Save the seating chart to xlsx file
main function:
    imput CLI option
    possible option:
        -w, --width
        -h, --hight
        -i, --input #xlsx file path
        -o, --output #xlsx file path
        -s, --skip #if option is out of range, skip it
    input data from xlsx file
    preprocessing preference options
    while True:
        generate random seating chart
        if seating chart is correct:
            break
    save seating chart to xlsx file
    program end
'''

# import module zone
import sys
import openpyxl as xl
import random
print = sys.stdout.write

#  sub functions zone

#load "person name","prefer option" from xlsx file
def load_data(path):
    '''
    load "person name","prefer option" from xlsx file
    '''
    wb = xl.load_workbook(path)
    sheet = wb.active
    data = []
    i = 1
    while True: #Repeat until you get a type of "None"
        if sheet['A'+str(i)].value == None:
            break
        data.append([sheet['A'+str(i)].value,sheet['B'+str(i)].value])
        i += 1
    return data



def preprocessing(data):
    '''
    Convert Excel coordinates (A3,B,3....) to 2D coordinates ([[1, 3], [2, 0], [0, 3]]) with this formula
    '''
    for i in range(len(data)):
        if log: print(f'\npreprocessing - "{data[i][0]}" \n=====================>>')
        if data[i][1] is None:
            data[i][1] = [[0, 0]]
            if log: print(f"\noptions: ['N/A'] \nconvert to coordinates: N/A --> [0, 0]")
        else:
            if isinstance(data[i][1], int):
                data[i][1] = str(data[i][1])
            coordinates = data[i][1].split(',')
            new_coords = []
            if log: print(f'\noptions: {coordinates}')
            for coord in coordinates:
                coord = coord.strip()
                y = ''
                x = 0
                for char in coord:
                    if char.isalpha():
                        if x != '':
                            x = x * 26 + (ord(char.upper()) - 64)
                        else:
                            x = 0
                    elif char.isdigit():
                        y += char
                if y != '':
                    y = int(y)
                else:
                    y = 0
                if log: print(f'\nconvert to coordinates: {coord} --> {[x, y]}')
                
                new_coords.append([x, y])
            data[i][1] = new_coords
        if log: print(f'\n=====================>>\n{i}\n')
    return data




def generate(data):
    '''
    generate random seating chart
    '''
    
    #make empty seating chart
    chart = [i for i in data]

    random.shuffle(chart)

    chart= [chart[i:i+width] for i in range(0, len(chart), width)]
    return chart


def option_check(chart):
    '''
    Make sure your preferred option has the correct placeholder
    '''
    if log: print('\noption check\n=====================>>')
    for y, row in enumerate(chart):
        for x, person in enumerate(row):
            if log: print(f'\n//matrix: {[x+1, y+1]} / name: {person[0]}')
            detect = 0
            for option in person[1]:
                if log: print(f'\n    option: {option} / check: ')
                if option[0] == 0 and option[1] == 0:
                    detect += 1
                    if log: print('True')
                if option[0] == x+1  and option[1] == 0:
                    detect += 1
                    if log: print('True')
                if option[0] == 0 and option[1] == y+1:
                    detect += 1
                    if log: print('True')
                if option[0] == x+1 and option[1] == y+1:
                    detect += 1
                    if log: print('True')
                #else:
                #    if log: print('Fales')
            if detect == 0:
                if log: print("\n=====================<< \noption check fail")
                return False
                
    if log: print("\n=====================<< \noption check success")
    return True
                    

                
        
    


if __name__ == '__main__':
    if len(sys.argv) == 1 and sys.argv[1] == '--help':
        #print CLI option
        print('''CLI options:
    >>Required options <<==========================
        -w, --width : set width of seating chart
        -h, --height : set height of seating chart
        -i, --input : set input xlsx file path
    >>Non-required options  <<=====================
        -o, --output : set output xlsx file path
        -s, --skip : all option skip
        --log      
    ===================<<==========================
''')
    elif len(sys.argv) > 2:
        for i in range(1, len(sys.argv)):

            if sys.argv[i] == '-w' or sys.argv[i] == '--width':
                width = int(sys.argv[i+1])
            if sys.argv[i] == '-h' or sys.argv[i] == '--height':
                height = int(sys.argv[i+1])
            if sys.argv[i] == '-i' or sys.argv[i] == '--input':
                input_path = sys.argv[i+1]
            if sys.argv[i] == '-o' or sys.argv[i] == '--output':
                output_path = sys.argv[i+1]
            skip = False
            if sys.argv[i] == '-s' or sys.argv[i] == '--skip':
                skip = True
        if sys.argv.count('--loging') > 0: log = True
        else: log = False
        # main function zone
        xlsx = load_data(input_path)
        data = preprocessing(xlsx)
        print('\ncheck input\n~~~~~~~~~~~~~~~~~~~~')
        for person in xlsx:
            print(f'\nname: {person[0]} / options: {person[1]}')
        print('\n~~~~~~~~~~~~~~~~~~~~\n\n')
        check = input('\ncontinew? (y/n):')
        if len(data) <= width * height and check == 'y':
            num = 0
            loop = False
            while loop == False:
                num += 1
                
                if log: print('\ngenerate chart\n~~~~~~~~~~~~~~~~~~~~~~~\n')
                chart = generate(data)
                if log:
                    for row in chart:
                        for person in row:
                            print(f'{person[0]} ')
                        print('\n')
                if log: print('~~~~~~~~~~~~~~~~~~~~~~~\n')
                loop = option_check(chart)
                if skip == True:
                    loop = True
                if log: print(f'\n{num}\n\n')
                #if not log:print(f'\r{num}')
            #if log: print('saving chart')
            print('\ngenerate chart\n~~~~~~~~~~~~~~~~~~~~~~~\n')
            #chart = generate(data)
            for row in chart:
                for person in row:
                    print(f'{person[0]} ')
                print('\n')
            print('~~~~~~~~~~~~~~~~~~~~~~~\n')
    else:
        if log: print('\npeople numder is too big')

    
        