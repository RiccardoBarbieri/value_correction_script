from copy import Error
import openpyxl as op
import sys
import decimal
import webbrowser

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def f_to_str(f):
    """
    Convert the given float to a string,
    without resorting to scientific notation
    """
    d1 = ctx.create_decimal(repr(f))
    return format(d1, 'f')

if __name__ == '__main__':

    ctx = decimal.Context()
    ctx.prec = 15

    if len(sys.argv) != 7:
        raise Error('Numero di parametri errato. Parametri: path input, path output, colonna tempi, colonna valori, colonna errori, range righe (n:m)')

    path = sys.argv[1]

    workbook = op.load_workbook(filename=path, data_only=True)
    sheet = workbook.active

    start = sys.argv[6].split(':')[0]
    end = sys.argv[6].split(':')[1]

    out_path = sys.argv[2]

    times_interval = sys.argv[3] + start + ':' + sys.argv[3] + end
    values_interval = sys.argv[4] + start + ':' + sys.argv[4] + end
    errors_interval = sys.argv[5] + start + ':' + sys.argv[5] + end

    times_tuple = sheet[times_interval]
    values_tuple = sheet[values_interval]
    errors_tuple = sheet[errors_interval]

    values = []
    for i in values_tuple:
        for j in i:
            values.append(j.value)
            
    errors = []
    for i in errors_tuple:
        for j in i:
            errors.append(j.value)

    errors_2cifre = []
    len_errors = [] # salvo index_first + 2 per evita problemi nel caso ultima cifra errore pari a 0
    for error in errors:
        string = f_to_str(error)
        string += '0'*15
        dec = string.split('.')[1]
        index_first = -1
        for i in enumerate(dec): # assegna a ogni carattera della string la sua posizione
            if (int(i[1]) != 0) and (index_first == -1): #salvo index se è un numero diverso da 0 e non è ancora stato fatto
                index_first = i[0]
        if int(dec[index_first + 2]) >= 5:
            dec = dec[:index_first + 1] + str(int(dec[index_first + 1]) + 1)
        len_errors.append(index_first + 2)
        dec = string.split('.')[0] + '.' +  dec[:index_first + 2] # +2 perché salvo la pos del primo non 0, aggiungo 2 per avere il secondo (indici non inclusivi)
    
        errors_2cifre.append(float(dec))
    
    values_correct = []
    for value, i in zip(values, range(len(values))): # i contiene indice
        # cifre_dec = len(f_to_str(errors_2cifre[i]).split('.')[1])
        cifre_dec = len_errors[i]
        # if cifre_dec == len(f_to_str(errors_2cifre[i]).split('.')[1]):
        #     print('{:6}{:6} {start}OK{end}'.format(cifre_dec, len(f_to_str(errors_2cifre[i]).split('.')[1]), start = bcolors.OKGREEN, end = bcolors.ENDC))
        # else:
        #     print('{:6}{:6} {start}FAIL{end}'.format(cifre_dec, len(f_to_str(errors_2cifre[i]).split('.')[1]), start = bcolors.FAIL, end = bcolors.ENDC))

        value_str_int = f_to_str(value).split('.')[0]
        value_str_dec = f_to_str(value).split('.')[1]
        try:
            if int(value_str_dec[cifre_dec]) >= 5:
                value_str_dec = value_str_dec[:cifre_dec - 1] + str(int(value_str_dec[cifre_dec - 1]) + 1)
        except IndexError:
            pass
        value_str_dec = value_str_dec[:cifre_dec]
        values_correct.append(float(value_str_int + '.' + value_str_dec))


    times = []
    for i in times_tuple:
        for j in i:
            times.append(j.value)
    
    new_work = op.Workbook()
    new_sheet = new_work.active

    if len(errors_2cifre) != len(times) or len(errors_2cifre) != len(values_correct) or len(values_correct) != len(times):
        raise Error('WTF')
    
    for i, time, value, error in zip(range(len(times)), times, values_correct, errors_2cifre):
        new_sheet['A' + str(i + 1)] = time
        new_sheet['B' + str(i + 1)] = value
        new_sheet['C' + str(i + 1)] = error

    new_work.save('new.xlsx')
    new_work.close()

    with open(out_path + '.dat', 'w+') as f:

        for time, value, error in zip(times, values_correct, errors_2cifre):
            f.write('{time:15}{value:15}{error:15}\n'.format(time = time, value = value, error = error))
    