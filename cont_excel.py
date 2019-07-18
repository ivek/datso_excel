from openpyxl import Workbook, load_workbook
from datetime import date, datetime, time, timedelta
import calendar
filepath = "C:/Users/criea\Google Drive/PYTHON_RELOAD/script_contador_excel/30FRB0007F.xlsx" #archivo excel lectura
libro = load_workbook(filepath)#asignación del libro a una variable
#hoja=libro.get_sheet_by_name("CRIE-UOP")
hoja=libro.active#cargar hoja de calculo la que este activa
seconth=0 #sexo contador hombre
secontm=0 #sexo contador hombre
for rows in hoja: #recorrer celdas
    nombre=rows[1].value#asignar celda nombre
    cols=rows[0].value#asignar celda numero consecutivo
    col=str(cols)#convertir en string numeros consecutivo
    nombre1=str(nombre)#convertir en string nombre
    contador=len(nombre1)#contar el numero de caracteres de la cadena nombre1
    curp=nombre1[(contador-19):(contador)]#sacar el curp de la cadena nombre1
    curp=curp.strip()#eliminar espacios vacios de la cadena
    numcurp= len(curp)#contar el numero de caraceres del curp
    formato_fecha='%d-%m-%Y'#formato fecha
    
    fecha_actual= date.today()#asignar la fecha actual
    if curp[10:-7] is not '' and curp[numcurp-1].isdigit() is True:#condicional si existe espacios vacios y el ultimo caracter es numero 
        fecha_nac= date(int('20'+curp[4:-12]),int(curp[6:-10]),int(curp[8:-8]))#se agrega 20... al año y se asigna a la variable fecha de nacimiento la fecha completa sustraida de la var curp
        edad_year= fecha_actual.year - fecha_nac.year - ((fecha_actual.month, fecha_actual.day) < (fecha_nac.month,fecha_nac.day))#se hace la operación para sacar la edad
        if (fecha_actual.month < fecha_nac.month):#condicional si el mes actual es menor del mes del año de nacimiento
            edad_month=((12- fecha_nac.month)+fecha_actual.month)#se hace la operacion del mes de nacimiento menos 12 meses y se le suma los meses actuales
        elif (fecha_actual.month > fecha_nac.month):#condcional fecha actual es mayor que fecha de nacimiento
            edad_month=(fecha_actual.month-fecha_nac.month)#se resta mes actual menos fecha nacimiento
        else:# de los contrario de sentencias anteriores
            edad_month=0#se asigna 0 a la variables edad_mes es el mismo mes de nacimiento y actual
        print(nombre1[contador-1].isdigit())#
        if str(curp[10:-7])=='H':#condicional saber si curp es hombre
            seconth = seconth + 1#cuenta numero de hombres
            sexo='Masculino'#asigna a variable sexo cadena masculino
        elif str(curp[10:-7])=='M':#condicional saber si curp es mujer 
            secontm+= 1#contador numero de mujeres
            sexo='Femenino'#asignar sexo femenino
        print(col[0:2]+' '+'curp:' + curp + ' sexo: '+ curp[11:-7] + sexo)#imprimir en pantalla curp, sexo
        
        print('edad = ' + str(edad_year) +' años cumplidos, '+ str(edad_month)+' meses')#imprimir en pantalla edad y mes
   

print("total mujeres: %s \n total hombres: %s" % (secontm,seconth))#imprimir total mujeres y hombres
    
    
    
            
    
    
    
       