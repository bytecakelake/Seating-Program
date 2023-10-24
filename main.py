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
import openpyxl



#  sub functions zone

#load "person name","prefer option" from xlsx file
def load_data(path):
    '''
    load "person name","prefer option" from xlsx file
    '''
    wb = openpyxl.load_workbook(path)
    sheet = wb.active
    data = []
    i = 1
    while True: #Repeat until you get a type of "None"
        if sheet['A'+str(i)].value == None:
            break
        data.append([sheet['A'+str(i)].value,sheet['B'+str(i)].value])
        i += 1
    return data

#Convert Excel coordinates (A3,B,3....) to 2D coordinates ([[1, 3], [2, 0], [0, 3]) with this formula
#Example of input data: [['권진??, 3], ['김문수', 'A'], ['명서??, 'c5'], ['문성준', None], ['박�???, 'AB,33'], ['박채??, 'WB2,abD45'], ['박혜??, 'A,B,D'], ['?��???, 'B4,D5']]
def preprocessing(data):
    '''
    Convert Excel coordinates (A3,B,3....) to 2D coordinates ([[1, 3], [2, 0], [0, 3]]) with this formula
    Example of input data: [['권진??, 3], ['김문수', 'A'], ['명서??, 'c5'], ['문성준', None], ['박�???, 'AB,33'], ['박채??, 'WB2,abD45'], ['박혜??, 'A,B,D'], ['?��???, 'B4,D5']]
    '''
    for i in range(len(data)):
        if data[i][1] is None:
            print('data:', data[i])
            print(['N/A'])
            data[i][1] = [0, 0]
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
        print(i, '\n')
    return data




# main function zone
if __name__ == '__main__':
    a = load_data('test.xlsx')
    print(a)
    b = preprocessing(a)
    print(b)