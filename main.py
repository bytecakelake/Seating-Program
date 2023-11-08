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
from tabnanny import check
import openpyxl as xl
import random


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
        if data[i][1] is None:
            print('data:', data[i])
            print(['N/A'])
            data[i][1] = [[0, 0]]
            print('convert coordinates: N/A --> [0, 0]')
        else:
            if isinstance(data[i][1], int):
                data[i][1] = str(data[i][1])
            print('data:', data[i])
            coordinates = data[i][1].split(',')
            new_coords = []
            print(coordinates)
            for coord in coordinates:
                coord = coord.strip()
                row = ''
                col = 0
                for char in coord:
                    if char.isalpha():
                        if col != '':
                            col = col * 26 + (ord(char.upper()) - 64)
                        else:
                            col = 0
                    elif char.isdigit():
                        row += char
                if row != '':
                    row = int(row)
                else:
                    row = 0
                print('convert coordinates:', coord, '-->', [col, row])
                
                new_coords.append([col, row])
            data[i][1] = new_coords
        print('number:', i, '\n')
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
    print('\noption check')
    for row in range(len(chart)):
        for col in range(len(chart[row])):
            print("=====================<<", "\nmatrix:", [col+1, row+1], '/', 'name:', chart[row][col][0])
            detect = 0
            for option in chart[row][col][1]:
                print('option:', option, '/', 'check:', end=' ')
                if option[0] == 0 and option[1] == 0:
                    detect += 1
                    print('True')
                elif option[0] == col+1 and option[1] == row+1:
                    detect += 1
                    print('True')
                elif option[0] == col+1 and option[1] == 0:
                    detect += 1
                    print('True')
                elif option[0] == 0 and option[1] == row+1:
                    detect += 1
                    print('True')
                else:
                    print('Fales')
            if detect == 0:
                print("=====================<<", '\noption check fail')
                return False
                
    print("=====================<<", '\noption check success')
    return True
                    


            
                
        
    


if __name__ == '__main__':
    print(sys.argv)
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
        
        # main function zone
        xlsx = load_data(input_path)
        data = preprocessing(xlsx)
        for person in xlsx:
            #print('=====================>>')
            print('name:', person[0], '/ options:', person[1])
        check = input('continew? (y/n):')
        if len(data) < width * height and check == 'y':
            num = 0
            loop = False
            while loop == False:
                num += 1
                
                print('generate chart')
                print('#######################')
                chart = generate(data)
                for row in chart:
                    for person in row:
                        print(person[0], end=' ')
                    print('')
                print('#######################')
                loop = option_check(chart)
                if skip == True:
                    loop = True
                print('try_count:', num)
                print('\n\n')
            #print('saving chart')

    else:
        print('people numder is too big')

    
        