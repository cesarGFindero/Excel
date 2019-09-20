# ****************************
# En este programa se rellena automaticamente una hoja en excel
#*****************************

import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell
from xlsxwriter.utility import xl_cell_to_rowcol
#import pdb 

def sumatoria(cargas):
    suma = '='
    for celda in cargas:
        suma = suma + celda + '+'
    return suma[:-1]

def suma_varias(*args):
    suma = '='
    for celda in args:
        suma = suma + celda + '+'
    return suma[:-1]

def producto_varias(*args):
    producto = '='
    for celda in args:
        producto = producto + celda + '*'
    return producto[:-1]

def reacomodo_fechas(fecha):
    nueva = fecha[-2:] + '-' + fecha[-5:-3] + '-' + fecha[0:4]
    return nueva
        

def crear_vlookup(filas,aparato):
    formula = '='
    for f in range(filas):
        for c in range(0,12):
            esquina_1 = xl_rowcol_to_cell(11+18*f,1+4*c)
            esquina_2 = xl_rowcol_to_cell(14+18*f,2+4*c)
            formula+='IFERROR(VLOOKUP("'+aparato+'",Detalles!'+esquina_1+':'+esquina_2+',2,FALSE),0)'+'+'
    return formula[:-1] 
    

def general(datos,precio,inicio,final,cliente,mes,nombre,workbook):
    
    worksheet = workbook.add_worksheet('General')
    
    celda_titulo = 'B3'
    
    titulo = workbook.add_format({'bold': True,'font_size':20})
    
    worksheet.write(celda_titulo,'Desciframiento del consumo de ' + nombre, titulo)


def detalles(datos,precio,inicio,final,cliente,mes,nombre,workbook,incios,finales,periodos,fugas_auto):
    
    worksheet = workbook.add_worksheet('Detalles')   
    
    celda_titulo = 'A1'
    celda_periodo = 'H2'
    celda_inicio_titulo = 'H3'
    celda_final_titulo = 'H4'
    celda_horas = 'I2'
    celda_inicio = 'I3'
    celda_final = 'I4'
    columna_nombre_finderos = 0
    fila_primer_findero = 5
    encabezado = 2
    espacio_vertical = 18
    espacio_horizontal = 7
    
    bold = workbook.add_format({'bold': True})
    bold_titulo = workbook.add_format({'bold': True,'font_size':15})
    centrado = workbook.add_format({'align':'center'})
    encabezados = workbook.add_format({'bold': True,'align':'center'})
    formato_kWh = workbook.add_format({'num_format': '0.0','align':'center'})
    money = workbook.add_format({'num_format': '_-$* #,##0.00_-;-$* #,##0.00_-;_-$* "-"??_-;_-@_-'})
    money_miles = workbook.add_format({'num_format': '$#,##0.00'})
    porcentaje = workbook.add_format({'num_format': '0.00%','align':'center'})
    condicional_amarillo = workbook.add_format({'bg_color':   '#FFEB9C',
                               'font_color': '#9C6500'})
    condicional_rojo = workbook.add_format({'bg_color':   '#FFC7CE',
                               'font_color': '#9C0006'})
    celdas_notas = workbook.add_format({'text_wrap': True, 'align': 'left', 'valign': 'top'})
                                       
    worksheet.write(celda_titulo,'Detalles del consumo de ' + nombre, bold_titulo)
    worksheet.write(celda_periodo,'Periodo:')
    worksheet.write(celda_horas,max(periodos))
    
    worksheet.write(celda_inicio_titulo,'Inicio:')
    worksheet.write(celda_inicio,reacomodo_fechas(inicio))
    
    worksheet.write(celda_final_titulo,'Final:')
    worksheet.write(celda_final,reacomodo_fechas(final))

    
       
    celdas_finderos = []
    global celdas_fugas
    celdas_fugas = []
    for indice,findero in enumerate(list(datos.keys())):
        celdas_cargas = []
        worksheet.merge_range(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos,
                              fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+encabezado,'Findero: '+findero[8:-4],encabezados)
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+3,'Periodo:')
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+6,'Inicio:')
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+9,'Final:')
        
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+4,periodos[indice])
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+7,incios[indice])
        worksheet.write(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+10,finales[indice])        
        
        for indice_,columna in enumerate(datos[findero]):
            worksheet.write(fila_primer_findero+3+indice*espacio_vertical,0+indice_*espacio_horizontal,'Puerto '+str(indice_+1), bold)
            worksheet.merge_range(fila_primer_findero+2+indice*espacio_vertical,0+indice_*espacio_horizontal,
                                  fila_primer_findero+2+indice*espacio_vertical,0+indice_*espacio_horizontal+encabezado,
                                  'Puerto '+str(indice_+1), encabezados)
            worksheet.write(fila_primer_findero+3+indice*espacio_vertical,0+indice_*espacio_horizontal,'kWh',centrado)
            worksheet.write(fila_primer_findero+3+indice*espacio_vertical,1+indice_*espacio_horizontal,'Señal',centrado)
            worksheet.write(fila_primer_findero+3+indice*espacio_vertical,2+indice_*espacio_horizontal,'%',centrado)
            worksheet.write(fila_primer_findero+5+indice*espacio_vertical,3+indice_*espacio_horizontal,'Potencia',encabezados)
            worksheet.write(fila_primer_findero+5+indice*espacio_vertical,4+indice_*espacio_horizontal,'Uso',encabezados)
            worksheet.write(fila_primer_findero+5+indice*espacio_vertical,5+indice_*espacio_horizontal,'Horas',encabezados)
            worksheet.write(fila_primer_findero+4+indice*espacio_vertical,0+indice_*espacio_horizontal, columna, formato_kWh)
            worksheet.write(fila_primer_findero+4+indice*espacio_vertical,1+indice_*espacio_horizontal,'-', centrado)
            
            celda_consumo = xl_rowcol_to_cell(fila_primer_findero+4+indice*espacio_vertical,0+indice_*espacio_horizontal)
            celda_total = xl_rowcol_to_cell(2+fila_primer_findero+len(list(datos.keys()))*espacio_vertical,1+columna_nombre_finderos)

            worksheet.write_formula(fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,'='+celda_consumo+'/$'+celda_total[0]+'$'+celda_total[1:],porcentaje)
            
            worksheet.conditional_format(fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                         fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                         {'type':     'cell',
                                        'criteria': 'between',
                                        'minimum':    .04,
                                        'maximum':    .09,
                                        'format':   condicional_amarillo})
            
            worksheet.conditional_format(fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                         fila_primer_findero+4+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                         {'type':     'cell',
                                        'criteria': '>=',
                                        'value':    .09,
                                        'format':   condicional_rojo})
    
            worksheet.write_formula(fila_primer_findero+6+indice*espacio_vertical,2+indice_*espacio_horizontal,
                                    '='+xl_rowcol_to_cell(fila_primer_findero+6+indice*espacio_vertical,0+indice_*espacio_horizontal)+
                                    '/$'+celda_total[0]+'$'+celda_total[1:], porcentaje)
            
            celdas_cargas.append(xl_rowcol_to_cell(fila_primer_findero+4+indice*espacio_vertical,0+indice_*espacio_horizontal)) 
            
            worksheet.merge_range(fila_primer_findero+4+indice*espacio_vertical+6,0+indice_*espacio_horizontal,
                      fila_primer_findero+4+indice*espacio_vertical+11,0+indice_*espacio_horizontal+2,'',celdas_notas)
            
            if indice_+1 == len(datos[findero]):
                
                suma_cargas = sumatoria(celdas_cargas)
                
                idx_ = indice_+1
                worksheet.write(fila_primer_findero+2+indice*espacio_vertical,0+idx_*espacio_horizontal,'Total findero', bold)
                worksheet.write(fila_primer_findero+3+indice*espacio_vertical,0+idx_*espacio_horizontal,'kWh', centrado)
                worksheet.write(fila_primer_findero+3+indice*espacio_vertical,1+idx_*espacio_horizontal,'%', centrado)
                worksheet.write_formula(fila_primer_findero+4+indice*espacio_vertical,0+idx_*espacio_horizontal,suma_cargas)
                worksheet.write_formula(fila_primer_findero+4+indice*espacio_vertical,1+idx_*espacio_horizontal,
                                        '='+xl_rowcol_to_cell(fila_primer_findero+4+indice*espacio_vertical,0+idx_*espacio_horizontal)+'/'+celda_total,porcentaje)

                celdas_finderos.append(xl_rowcol_to_cell(fila_primer_findero+4+indice*espacio_vertical,0+idx_*espacio_horizontal))


        for indice_, elemento in enumerate(fugas_auto[findero[:-4]]):
            if elemento != 0:
                worksheet.write(fila_primer_findero+6+indice*espacio_vertical,0+indice_*espacio_horizontal, '='+str(elemento)+'*'+xl_rowcol_to_cell(fila_primer_findero+indice*espacio_vertical,columna_nombre_finderos+4)+'/1000', formato_kWh)
                worksheet.write(fila_primer_findero+6+indice*espacio_vertical,0+indice_*espacio_horizontal+1, f'Fuga {int(elemento)}')
                celdas_fugas.append(xl_rowcol_to_cell(fila_primer_findero+6+indice*espacio_vertical,0+indice_*espacio_horizontal+1))
            
        if indice+1 == len(list(datos.keys())):
            idx = indice+1
            
            suma_finderos = sumatoria(celdas_finderos) 
            
            worksheet.merge_range(fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,
                              fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+encabezado-1,'Periodo',encabezados)
            
            worksheet.merge_range(fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+3,
                              fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+encabezado-1+3,'Bimestre',encabezados)
            
            worksheet.merge_range(fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+6,
                              fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+encabezado-1+6,'Real',encabezados)
            
            worksheet.write(1+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Horas:')
            worksheet.write_formula(1+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,'=I2')
            
            worksheet.write(-1+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'DAC:')
            worksheet.write(-1+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,precio,money)
            celda_precio = xl_rowcol_to_cell(-1+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos)
            
            worksheet.write(2+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Consumo:')
            worksheet.write_formula(2+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,suma_finderos,formato_kWh)
            celda_consumo = xl_rowcol_to_cell(2+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos)
            
            worksheet.write(3+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Costo:')
            worksheet.write(3+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,producto_varias(celda_precio,celda_consumo),money_miles)
            celda_costo = xl_rowcol_to_cell(3+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos)

            worksheet.write(4+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Periodos al bimestre:')
            worksheet.write(4+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos,'=(60*24)/I2',formato_kWh)
            celda_periodos = xl_rowcol_to_cell(4+fila_primer_findero+idx*espacio_vertical,1+columna_nombre_finderos)

            
            worksheet.write(1+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos,'Consumo:')
            worksheet.write(1+fila_primer_findero+idx*espacio_vertical,4+columna_nombre_finderos,'='+celda_periodos+'*'+celda_consumo,formato_kWh)
            
            worksheet.write(2+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos,'Costo:')
            worksheet.write(2+fila_primer_findero+idx*espacio_vertical,4+columna_nombre_finderos,'='+celda_periodos+'*'+celda_costo,money_miles)
            
            worksheet.write(1+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos,'Consumo:')
            worksheet.write(1+fila_primer_findero+idx*espacio_vertical,7+columna_nombre_finderos,'='+xl_rowcol_to_cell(6+fila_primer_findero+idx*espacio_vertical,7+columna_nombre_finderos)+'-'+
                                                                                            xl_rowcol_to_cell(5+fila_primer_findero+idx*espacio_vertical,7+columna_nombre_finderos),formato_kWh)
            
            worksheet.write(2+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos,'Error:')
            worksheet.write(2+fila_primer_findero+idx*espacio_vertical,7+columna_nombre_finderos,'=('
                                                            +xl_rowcol_to_cell(1+fila_primer_findero+idx*espacio_vertical,7+columna_nombre_finderos)+'-'+
                                                        celda_consumo+')/'+xl_rowcol_to_cell(1+fila_primer_findero+idx*espacio_vertical,7+columna_nombre_finderos),porcentaje)
            
            worksheet.write(5+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos,'Inicio:')
            
            worksheet.write(6+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos,'Final:')
            
            worksheet.merge_range(9+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,
                                  9+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+7,'Consumo parcial aparato',encabezados)
            worksheet.merge_range(11+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,
                                  11+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1,'Bloque 1',encabezados)
            worksheet.merge_range(11+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+3,
                                  11+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+3+1,'Bloque 2',encabezados)
            worksheet.merge_range(11+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+6,
                                  11+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+6+1,'Bloque 3',encabezados)
           
            worksheet.write(12+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Consumo:')
            worksheet.write(13+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Horas:')
            worksheet.write(14+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Base (W):')
            worksheet.write(16+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Aparato:')
            worksheet.write(16+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1,
                            '='+xl_rowcol_to_cell(12+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1)
                            +'-('+xl_rowcol_to_cell(14+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1)
                            +'*'+xl_rowcol_to_cell(13+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1)
                            +'/1000)')


            worksheet.write(12+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos,'Consumo:')
            worksheet.write(13+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos,'Horas:')
            worksheet.write(14+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos,'Base (W):')
            worksheet.write(16+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos,'Aparato:')
            worksheet.write(16+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos+1,
                            '='+xl_rowcol_to_cell(12+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos+1)
                            +'-('+xl_rowcol_to_cell(14+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos+1)
                            +'*'+xl_rowcol_to_cell(13+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos+1)
                            +'/1000)')

           
            worksheet.write(12+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos,'Consumo:')
            worksheet.write(13+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos,'Horas:')
            worksheet.write(14+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos,'Base (W):')
            worksheet.write(16+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos,'Aparato:')            
            worksheet.write(16+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos+1,
                            '='+xl_rowcol_to_cell(12+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos+1)
                            +'-('+xl_rowcol_to_cell(14+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos+1)
                            +'*'+xl_rowcol_to_cell(13+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos+1)
                            +'/1000)')
           
           
            worksheet.write(18+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'SumHor:')
            worksheet.write(18+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1,
                            '='+xl_rowcol_to_cell(13+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1)+'+'
                            +xl_rowcol_to_cell(13+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos+1)+'+'
                            +xl_rowcol_to_cell(13+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos+1))
           
           
            worksheet.write(19+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'SumCons:')
            worksheet.write(19+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1,
                            '='+xl_rowcol_to_cell(16+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1)+'+'
                            +xl_rowcol_to_cell(16+fila_primer_findero+idx*espacio_vertical,3+columna_nombre_finderos+1)+'+'
                            +xl_rowcol_to_cell(16+fila_primer_findero+idx*espacio_vertical,6+columna_nombre_finderos+1))

            worksheet.write(21+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos,'Poyección:')
            worksheet.write(21+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1,
                            '=(I2*'+xl_rowcol_to_cell(19+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1)
                            +')/'+xl_rowcol_to_cell(18+fila_primer_findero+idx*espacio_vertical,columna_nombre_finderos+1))
            
            global celda_consumo_bim
            celda_consumo_bim = xl_rowcol_to_cell(1+fila_primer_findero+idx*espacio_vertical,4+columna_nombre_finderos)
    
    
def desciframiento(datos,precio,inicio,final,cliente,mes,nombre,workbook,num_datos):
    
    titulos = ['Consumo (kWh)', 'Gasto', 'Ubicación', 'Equipo', 'Proporción (%)','Potencia (W)','Tiempo de uso','Hrs semana','Consumo (kWh)', 'Gasto','Gasto anual', 'Notas']
    anchos = [12,12,15,19,15,15,15,15,15,15,15,80,10]
    celda_titulo = 'A1'
    celdas = 15+len(celdas_fugas)
    filas = len(datos.keys())
    
#    bold_1 = workbook.add_format({'bold': True,'font_size':12})
    bold_2 = workbook.add_format({'bold': True,'font_size':15})
    bold_3 = workbook.add_format({'bold': True,'bg_color':'#E1E4EB','align':'center','border':1})
    columna_gris = workbook.add_format({'bold': True,'bg_color':'#F2F2F2','border':1})
    columna_blanca_1 = workbook.add_format({'align':'center','num_format': '0.0 %','border':1})                                        
    columna_blanca_2 = workbook.add_format({'align':'center','num_format': '#','border':1})
    columna_blanca_3 = workbook.add_format({'align':'center','num_format': '$  #,###     ','border':1})
    columna_blanca_notas = workbook.add_format({'align':'left','num_format': '#','border':1})                                        
    columna_blanca_vacia = workbook.add_format({'align':'center','num_format': '#','border':1})
    dinero_1 = workbook.add_format({'align':'center','num_format': '$   #,###.##   ','border':1,'align':'center'}) 
    centrado = workbook.add_format({'align': 'center', 'valign': 'center'})
                    
    worksheet = workbook.add_worksheet('Desciframiento') 
    
    worksheet.write(celda_titulo,'Desciframiento del consumo de ' + nombre + ' (cifras bimestrales)', bold_2)
    
    worksheet.write('H1',f'Se analizaron {num_datos:,} datos',centrado)
    
    for idx,titulo in enumerate(titulos):
        skip = 0
        reset = 0
        if idx>1:
            skip = 4
            reset = 2
        worksheet.set_column(idx, idx, anchos[idx])
        worksheet.write(xl_rowcol_to_cell(3+skip,idx-reset+2),titulo, bold_3)
        
    
    for i in range(0,celdas):
        worksheet.write_blank(xl_rowcol_to_cell(8+i,2),'', columna_gris)
        worksheet.write_blank(xl_rowcol_to_cell(8+i,3),'', columna_gris)
        worksheet.write_blank(xl_rowcol_to_cell(8+i,4),'', columna_blanca_1)
        worksheet.write_blank(xl_rowcol_to_cell(8+i,5),'', columna_blanca_2)
        worksheet.write_blank(xl_rowcol_to_cell(8+i,6),'', columna_blanca_2)
        worksheet.write_blank(xl_rowcol_to_cell(8+i,7),'', columna_blanca_2)
        worksheet.write_blank(xl_rowcol_to_cell(8+i,11),'', columna_blanca_notas)
        
    for i in range(0,celdas+4):   
        if i in [1,5,9,13,17]:
            worksheet.write_formula(xl_rowcol_to_cell(8+i,5),'', columna_blanca_vacia)
            worksheet.write_formula(xl_rowcol_to_cell(8+i,6),'', columna_blanca_vacia)
        else:
            worksheet.write_formula(xl_rowcol_to_cell(8+i,8),'=' + xl_rowcol_to_cell(8+i,4) + '*$C$5', columna_blanca_2)
            worksheet.write_formula(xl_rowcol_to_cell(8+i,9),'=' + xl_rowcol_to_cell(8+i,4) + '*$D$5', columna_blanca_3)            
            worksheet.write_formula(xl_rowcol_to_cell(8+i,10),'=' + xl_rowcol_to_cell(8+i,9) + '*6', columna_blanca_3)

            
    worksheet.write('F4','Tarifa DAC:',bold_3)
    worksheet.write_number('G4', 5.333, dinero_1)    
    worksheet.write_formula('D5','=C5*G4',columna_blanca_3)
    worksheet.write_formula('C5','=Detalles!'+celda_consumo_bim,columna_blanca_2)
    

    worksheet.write('D9','Refrigerador', columna_gris)
    worksheet.write('M9','Cava [m3]:',columna_blanca_1)
    worksheet.write_formula('N9','IF(K9>162,((K9-162.912)/11.974)*0.0283168,0)',columna_blanca_2)
    
    worksheet.write('D11','Bomba de agua', columna_gris)
    worksheet.write('D12','Centro de lavado', columna_gris)
    worksheet.write('D13',"Tv's", columna_gris)
    worksheet.write('M13','TV ["]:',columna_blanca_1)
    worksheet.write_formula('N13','4.9012*(F13^0.4627)',columna_blanca_2)
    
    worksheet.write_formula('E11',crear_vlookup(filas,'Bomba'), columna_blanca_1)  
    worksheet.write_formula('E12',crear_vlookup(filas,'Lavado'), columna_blanca_1)
    worksheet.write_formula('E13',crear_vlookup(filas,'TV'), columna_blanca_1)

    worksheet.write('D15',"Pendiente 1", columna_gris)
    worksheet.write('D16',"Pendiente 2", columna_gris)
    worksheet.write('D17',"Pendiente 3", columna_gris)
    
    global porcentajes_fugas
    porcentajes_fugas = []
    
    for i, celda in enumerate(celdas_fugas):
        celda_ = xl_cell_to_rowcol(celda)
        
        celda_porcentaje = (celda_[0],celda_[1]+1)
        celda_porcentaje = xl_rowcol_to_cell(celda_porcentaje[0],celda_porcentaje[1])
        
        celda_circuito = (celda_[0]-5,celda_[1]+1)
        celda_circuito = xl_rowcol_to_cell(celda_circuito[0],celda_circuito[1])

        celda_notas = (celda_[0]+4,celda_[1]-1)
        celda_notas = xl_rowcol_to_cell(celda_notas[0],celda_notas[1])

        worksheet.write_formula(xl_rowcol_to_cell(18+i,4), '=Detalles!'+celda_porcentaje, columna_blanca_1)
        worksheet.write(xl_rowcol_to_cell(18+i,3),'=CONCATENATE("Fuga ",MID(Detalles!'+celda+',FIND(" ",Detalles!'+celda+')+1,256))', columna_gris)
        worksheet.write(xl_rowcol_to_cell(18+i,2),'=CONCATENATE("C. ",Detalles!'+celda_circuito+')', columna_gris)
        worksheet.write(xl_rowcol_to_cell(18+i,7+4),'=Detalles!'+celda_notas,columna_blanca_notas)
        porcentajes_fugas.append(xl_rowcol_to_cell(18+i,4))
        
#    for i in range(5):
#        row_fuga = 'ROW(INDIRECT(MID(FORMULATEXT('+xl_rowcol_to_cell(14+i,4)+'),FIND("!",FORMULATEXT('+xl_rowcol_to_cell(14+i,4)+'))+1,256)))'
#        column_fuga = 'COLUMN(INDIRECT(MID(FORMULATEXT('+xl_rowcol_to_cell(14+i,4)+'),FIND("!",FORMULATEXT('+xl_rowcol_to_cell(14+i,4)+'))+1,256)))-2'
#        argumento = 'FORMULATEXT(INDIRECT(ADDRESS('+row_fuga+','+column_fuga+',,,"Detalles")))'
#        worksheet.write_formula(xl_rowcol_to_cell(14+i,3),'IFERROR(CONCATENATE("Fuga ",'+'MID(LEFT('+argumento+',FIND("*",'+argumento+')-1),FIND("=",'+argumento+')+1,LEN('+argumento+'))),"Fuga")', columna_gris)
        
    worksheet.write(xl_rowcol_to_cell(18+i+2,3),'Luces', columna_gris)
    worksheet.write(xl_rowcol_to_cell(18+i+3,3),'Cómputo y cargadores', columna_gris)
    worksheet.write(xl_rowcol_to_cell(18+i+4,3),'Sin Identificar', columna_gris)

    worksheet.write_formula(xl_rowcol_to_cell(18+i+2,4),crear_vlookup(filas,'Luces'), columna_blanca_1)
    worksheet.write_formula(xl_rowcol_to_cell(18+i+3,4),crear_vlookup(filas,'Computo'), columna_blanca_1)
    worksheet.write_formula(xl_rowcol_to_cell(18+i+4,4),crear_vlookup(filas,'Sin ID'), columna_blanca_1)
    
    worksheet.write(xl_rowcol_to_cell(8+celdas,3),'Total',bold_3)
    worksheet.write(xl_rowcol_to_cell(8+celdas+1,3),'Total de Fugas',bold_3)
    worksheet.write(xl_rowcol_to_cell(8+celdas+2,3),'Fugas atacables',bold_3)
    worksheet.write(xl_rowcol_to_cell(8+celdas+3,3),'Total de pendientes',bold_3)
    
    
    
    worksheet.write_formula(xl_rowcol_to_cell(8+celdas,4),'=SUM('+xl_rowcol_to_cell(8,4)+
                                                                ':'+xl_rowcol_to_cell(8+celdas-1,4)+
                                                                                ')',columna_blanca_1)
    try:
        worksheet.write_formula(xl_rowcol_to_cell(8+celdas+1,4),'=SUM('+porcentajes_fugas[0]+
                                                                ':'+porcentajes_fugas[-1]+
                                                                                ')',columna_blanca_1)
    except:
        pass

    worksheet.write_formula(xl_rowcol_to_cell(8+celdas+2,4),'='+xl_rowcol_to_cell(8+celdas+1,4)+'-40/C5'
                                                                            ,columna_blanca_1)
    
    worksheet.write_formula(xl_rowcol_to_cell(8+celdas+3,4),'=SUM(E15:E17)',columna_blanca_1)
    
    global fugas_atacables
    fugas_atacables = xl_rowcol_to_cell(8+celdas+2,5)
    
def ahorro(datos,precio,inicio,final,cliente,mes,nombre,workbook):  
    
    celda_titulo = 'A1'
    
    bold = workbook.add_format({'bold': True})        
    bold_1 = workbook.add_format({'bold': True,'font_size':15})
    bold_2 = workbook.add_format({'bold': True,'align':'center'})
    bold_3 = workbook.add_format({'bold': True,'top':1})
    blanco = workbook.add_format({'font_color':'#ffffff'})
    dinero_1 = workbook.add_format({'num_format': '$      #,###      '})
    dinero_2 = workbook.add_format({'num_format': '$      #,###      ','top':1})
    kWh = workbook.add_format({'num_format': '#','align':'center'})
    kWh_2 = workbook.add_format({'num_format': '#','top':1})
                               
    worksheet = workbook.add_worksheet('Ahorro') 
    worksheet.set_column(0,0,45)
    worksheet.set_column(2, 2, 15)
    worksheet.set_column(3, 3, 15)
    worksheet.set_column(6, 6, 15)
    worksheet.set_column(7, 7, 13)
    worksheet.set_column(8, 8, 13)

    
    worksheet.write(celda_titulo,'Potencial de ahorro de ' + nombre,bold_1)
    worksheet.write('A4','Acción', bold_2)
    worksheet.write('A5','Eliminar fugas')
    worksheet.write('C4','Ahorro (DAC)',bold_2)
    worksheet.write('D4','Ahorro en kWh',bold_2)
    worksheet.write('B9','Total: ', bold_3)
    worksheet.write_formula('C9','SUM(C5:C8)',dinero_2) 
    worksheet.write_formula('D9','SUM(D5:D8)',kWh_2)
#    worksheet.write_formula('D5','IF(SUM(Desciframiento!F15:F19)-40<0,0,SUM(Desciframiento!F15:F19)-40)',kWh)
    x, y = xl_cell_to_rowcol(fugas_atacables)
    worksheet.write_formula('D5','IF(Desciframiento!'+xl_rowcol_to_cell(x,y+3)+'<0,0,Desciframiento!'+xl_rowcol_to_cell(x,y+3)+')', kWh)
    worksheet.write('D6','',kWh)
    worksheet.write('D7','',kWh)
    worksheet.write('D8','',kWh)
    worksheet.write_formula('C5','D5*Desciframiento!$G$4',dinero_1)
    
    worksheet.write('A13','Cambio de tarifa',bold_2)
    worksheet.write_formula('I2','IF(D15<500,1,0)',blanco)
    
    worksheet.write_formula('A14','IF(I2=1,"Sí es posible bajar de tarifa","No es posible bajar de tarifa")')
    
    worksheet.write('C13','DAC:',bold_2)
    
    worksheet.write('C14','Nuevo recibo',bold_2)
    worksheet.write_formula('C15','Desciframiento!D5-C9',dinero_1)
    
    worksheet.write('D14','Nuevo consumo',bold_2)
    worksheet.write_formula('D15','Desciframiento!C5-D9',kWh)
    
    worksheet.write_formula('C17','IF(I2=1,"Subsidiada:","")',bold_2)
    
    worksheet.write_formula('C18','IF(I2=1,"Nuevo recibo","")',bold_2)
    worksheet.write_formula('D18','IF(I2=1,"Ahorro total","")',bold_2)
    
    worksheet.write_formula('C19','IF(I2=1,IF(D15>=280,150*.947+130*1.146+(D15-280)*3.352,IF(D15>=150,150*.947+(D15-150)*1.146,D15*.947)),"")',dinero_1)
    worksheet.write_formula('D19','IF(I2=1,Desciframiento!D5-C19,"")',dinero_1)
    
    worksheet.write('G4','Paneles para reducir el recibo a cero:', bold)
    worksheet.write('H6','Antes:', bold_2)
    worksheet.write('I6','Después de implementar:', bold)    
    worksheet.write('G7','No. de paneles:', bold)
    worksheet.write('G8','Costo:', bold)
    worksheet.write('G10','Ahorro por implementar:', bold)
    
    worksheet.write_formula('H7','Desciframiento!C5/68', kWh)  # 68 son los kWh 
    worksheet.write_formula('I7','D15/68', kWh)  
    
    worksheet.write_formula('H8','H7*13500', dinero_1)  # 13500 es el costo de un panel 
    worksheet.write_formula('I8','I7*13500', dinero_1)
    worksheet.write_formula('G11','H8-I8', dinero_1)

    worksheet.write('G15','Impacto ambiental del ahorro:', bold)
    worksheet.write_formula('G16','ROUND(D9*0.527*6,0)')  # se multiplica por el factor de emisión en kg. Para actualizar, ver en CRE.
    worksheet.write('H16','kg de CO2e al año')
    worksheet.write_formula('G18','ROUND(D9*0.015*6,0)')  # se multiplica por el numero de arboles necesarios para secuestrar una tonelada de CO2e, ver en carbonneutral.com/FAQs
    worksheet.write('H18','árboles plantados que absorben esa cantidad de CO2e')
    
    
def modelo_paneles(datos,precio,inicio,final,cliente,mes,nombre,workbook):  
 
    worksheet = workbook.add_worksheet('Modelo_paneles')
        
    bold_1 = workbook.add_format({'bold': True,'font_size':16})
    bold_2 = workbook.add_format({'bold': True,'align':'center'})
    centrado_model = workbook.add_format({'align':'center'})
    dinero_centrado = workbook.add_format({'num_format': '$      #,###      ','align':'center'})
    text_adjust = workbook.add_format()
    text_adjust.set_text_wrap() 
    background_lime = workbook.add_format({'bold': True,'align':'center','font_size':14})
    background_lime.set_pattern(1)  # This is optional when using a solid fill.
    background_lime.set_bg_color('#BFD19F') 
    background_brown = workbook.add_format({'bold': True,'align':'center','font_size':14})
    background_brown.set_pattern(1)  # This is optional when using a solid fill.
    background_brown.set_bg_color('#D6C8BF') 
    background_silver1 = workbook.add_format({'bold': True,'align':'center','font_size':14})
    background_silver1.set_pattern(1)  # This is optional when using a solid fill.
    background_silver1.set_bg_color('#D2D3D4') 
    background_silver2 = workbook.add_format({'bold': True,'align':'center','font_size':14})
    background_silver2.set_pattern(1)  # This is optional when using a solid fill.
    background_silver2.set_bg_color('#D2D3D4')
    background_silver2.set_text_wrap()
    background_cyan = workbook.add_format({'bold': True,'align':'center','font_size':14})
    background_cyan.set_pattern(1)  # This is optional when using a solid fill.
    background_cyan.set_bg_color('#CCE4ED')
    background_cyan2 = workbook.add_format({'bold': True,'align':'center','font_size':14})
    background_cyan2.set_pattern(1)  # This is optional when using a solid fill.
    background_cyan2.set_bg_color('#CCE4ED')
    background_cyan2.set_text_wrap()
                                  
    worksheet.set_column(1, 8, 20)  # Width of columns set to 20.
    worksheet.set_column(9,19,15)

                                     
    worksheet.write('B2','Modelo de ahorros comparativos con paneles solares',bold_1)
    worksheet.write('B3','USO EXCLUSIVO DE FINDERO O SUS CLIENTES')
    worksheet.write('B4','Creado por Findero el 14 de Septiembre del 2019')

    worksheet.write('B7','Fechas importantes',background_brown)
    worksheet.write('B8','Fecha de implementación de medidas',text_adjust)
    worksheet.write('B9','Fecha inical de bimestre después de implementación',text_adjust)
    worksheet.write('B10','Días que cubre el recibo',centrado_model)
    worksheet.write('C7','Día(s)',background_brown)
    worksheet.write_blank('C8', None,centrado_model)
    worksheet.write_blank('C9', None,centrado_model)
    worksheet.write_number('C10',60,centrado_model)
    
    worksheet.write('B14','Resultados',bold_2)
    worksheet.write('B15','Ahorro total con Findero')
    worksheet.write('B16','Ahorro con paneles solares')
    worksheet.write('B17','Ganancia/Pérdida con paneles solares')
    worksheet.write('B18','Ganancia/Pérdida con Findero')
    
    worksheet.write('B20','TABLA COMPARATIVA',bold_1)
    worksheet.write('B21','Bimestre a partir de intervención de ahorro',background_cyan2)
    worksheet.write('C21','Año',background_cyan)
    worksheet.write('D21','Consumo',background_cyan)
    worksheet.write('E21','Tipo de consumo',background_cyan)
    worksheet.write('F21','Promedio de consumo',background_cyan2)
    worksheet.write('G21','Tipo de tarifa',background_cyan)
    worksheet.write('H21','Costo aproximado de la energía',background_cyan2)
    worksheet.write('B22','Bimestre -6',centrado_model)
    worksheet.write('B23','Bimestre -5',centrado_model)
    worksheet.write('B24','Bimestre -4',centrado_model)
    worksheet.write('B25','Bimestre -3',centrado_model)
    worksheet.write('B26','Bimestre -2',centrado_model)
    worksheet.write('B27','Bimestre -1',centrado_model)
    worksheet.write('B28','Bimestre 0',centrado_model)
    worksheet.write('B29','Bimestre 1',centrado_model)
    worksheet.write('B30','Bimestre 2',centrado_model)
    worksheet.write('B31','Bimestre 3',centrado_model)
    worksheet.write('B32','Bimestre 4',centrado_model)
    worksheet.write('B33','Bimestre 5',centrado_model)
    worksheet.write('B34','Bimestre 6',centrado_model)
    
    id_cero = xl_cell_to_rowcol('C22')[0]+1
    for j in range(28-22+1):
        worksheet.write_number('C'+str(id_cero+j),0,centrado_model)
    id_uno = xl_cell_to_rowcol('C29')[0]+1
    for j in range(34-29+1):
        worksheet.write_number('C'+str(id_uno+j),1,centrado_model)
    
    id_consumo = xl_cell_to_rowcol('D22')[0]+1
    for j in range(27-22+1):
        worksheet.write_blank('D'+str(id_consumo+j), None,centrado_model)
    worksheet.write('D28','=ROUND($F$27-$G$8*(C9-C8)/C10,0)',centrado_model)
    id_resta = xl_cell_to_rowcol('D29')[0]+1
    for j in range(34-29+1):
        worksheet.write('D'+str(id_resta+j),'=ROUND($F$27-$G$8,0)',centrado_model)
    
    id_real = xl_cell_to_rowcol('E22')[0]+1
    for j in range(27-22+1):
        worksheet.write('E'+str(id_real+j),'Real',centrado_model)
    id_estim = xl_cell_to_rowcol('E28')[0]+1
    for j in range(34-28+1):
        worksheet.write('E'+str(id_estim+j),'Estimado',centrado_model)
        
    id_NA = xl_cell_to_rowcol('F22')[0]+1
    for j in range(26-22+1):
        worksheet.write('F'+str(id_NA+j),'NA',centrado_model) 
    id_prom = xl_cell_to_rowcol('F27')[0]+1
    for j in range(34-27+1):
        worksheet.write('F'+str(id_prom+j),'=ROUND(AVERAGE(D'+str(id_prom-5+j)+':D'+str(id_prom+j)+'),0)',centrado_model)
    
    id_tot = xl_cell_to_rowcol('G22')[0]+1
    for j in range(34-22+1):
        worksheet.write('G'+str(id_tot+j),'=IF(F'+str(id_tot+j)+'>499,"DAC",1)',centrado_model)
        worksheet.write('H'+str(id_tot+j),'=IF(G'+str(id_tot+j)+'="DAC",D'+str(id_tot+j)+'*5.55872+250.35,0.94772*149+1.14608*129+(D'+str(id_tot+j)+'-280)*3.3524)',dinero_centrado)
    
    worksheet.merge_range('F6:H6','Intervención Findero',background_lime)
    worksheet.write('F7','Variable',background_lime)
    worksheet.write('G7','Valor',background_lime)
    worksheet.write('H7','Unidades',background_lime)
    worksheet.write('F8','Ahorro estimado mínimo',centrado_model)
    worksheet.write('F9','Costo de medidas de ahorro',centrado_model)
    worksheet.write('F10','Costo de servicio',centrado_model)
    worksheet.write('F11','Bimestralidades Findero',centrado_model)
    worksheet.write('G8','=ROUND(Ahorro!D9,0)',centrado_model)
    worksheet.write_blank('G9',None,dinero_centrado)
    worksheet.write_number('G10',7800,centrado_model)
    worksheet.write('G11','=G10*1.12/6',centrado_model)
    worksheet.write('H8', 'KWh',centrado_model)
    worksheet.write('H9','$/medidas de ahorro',centrado_model)
    worksheet.write('H10','$/servicio',centrado_model)
    worksheet.write('H11','$/bimestre/servicio',centrado_model)
    
    worksheet.merge_range('J6:S6','Intervención paneles solares',background_silver1)
    worksheet.write('J7','Número de paneles solares',background_silver2)
    worksheet.write('K7','Precio de paneles solares',background_silver2)
    worksheet.write('L7','Generación bimestral',background_silver2)
    
    worksheet.write('M7','Bimestralidad (12 meses)',background_silver2)
    worksheet.write('N7','Bimestralidad (18 meses)',background_silver2)
    worksheet.write('O7','Bimestralidad (24 meses)',background_silver2)
    worksheet.write('P7','Bimestralidad (36 meses)',background_silver2)
    worksheet.write('Q7','Bimestralidad (60 meses)',background_silver2)
    worksheet.write('R7','Bimestralidad (84 meses)',background_silver2)
    worksheet.write('S7','Contratos de largo plazo',background_silver2)

    id_mens = xl_cell_to_rowcol('M8')[0]+1
    for j in range(18-8+1):
        worksheet.write('M'+str(id_mens+j),'=-PMT(0,12/2,K'+str(id_mens+j)+')',dinero_centrado)
        worksheet.write('N'+str(id_mens+j),'=-PMT(0,18/2,K'+str(id_mens+j)+')',dinero_centrado)
        worksheet.write('O'+str(id_mens+j),'=-PMT(0,24/2,K'+str(id_mens+j)+')',dinero_centrado)
        worksheet.write('P'+str(id_mens+j),'=-PMT(0.092/6,36/2,K'+str(id_mens+j)+')',dinero_centrado)
        worksheet.write('Q'+str(id_mens+j),'=-PMT(0.092/6,60/2,K'+str(id_mens+j)+')',dinero_centrado)
        worksheet.write('R'+str(id_mens+j),'=-PMT(0.092/6,84/2,K'+str(id_mens+j)+')',dinero_centrado)
    worksheet.write_number('J8',4,centrado_model)
    
    
    id_cmplt = xl_cell_to_rowcol('K8')[0]+1
    for j in range(18-8+1):
        worksheet.write('K'+str(id_cmplt+j),'=J'+str(id_cmplt+j)+'*12500+6500',dinero_centrado)#Pendiente e intercepto de regresión lineal con datos Enlight
        worksheet.write('L'+str(id_cmplt+j),'=ROUND((J'+str(id_cmplt+j)+'*330*4.1/1.2)*60/1000,1)',centrado_model)#Numero de modulos por potencia de cada uno por mínima insolacion solar promedio por 60 días del bimestre
    
    id_panel = xl_cell_to_rowcol('J9')[0]+1
    for j in range(18-9+1):
        worksheet.write('J'+str(id_panel+j),'=J'+str(id_panel+j-1)+'+2',centrado_model)



def fugas(datos,precio,inicio,final,cliente,mes,nombre,workbook):
    
    worksheet = workbook.add_worksheet('Fugas')
    worksheet.set_column(1,1,2)
    worksheet.set_column(4,4,2)
        
    bold_1 = workbook.add_format({'bold': True,'font_size':15})
    bold_2 = workbook.add_format({'bold': True,'align':'center'})
    center = workbook.add_format({'num_format': '0.0','align':'center'})
    center_2 = workbook.add_format({'align':'center'})
    formato_nota = workbook.add_format({'align': 'left', 'num_format': '#'})
    
    
    worksheet.write('A1','Resumen de fugas de ' + nombre,bold_1)
    worksheet.write('A4','Circuito',bold_2)
    worksheet.write('C4','Fuga (W)',bold_2)
    worksheet.write('D4','A',bold_2)
    worksheet.write('F4','Notas',bold_2)
    
    for i, elemento in enumerate(porcentajes_fugas):
        elemento_ = xl_cell_to_rowcol(elemento)
        fuga_titulo = xl_rowcol_to_cell(elemento_[0],elemento_[1]-1)
        circ_titulo = xl_rowcol_to_cell(elemento_[0],elemento_[1]-2)
        nota_titulo = xl_rowcol_to_cell(elemento_[0],elemento_[1]+3+4)
        worksheet.write_formula(xl_rowcol_to_cell(5+2*i,3),'=IFERROR('+xl_rowcol_to_cell(5+2*i,2)+'/127,0)',center)  # Columna del amperaje
        worksheet.write_formula(xl_rowcol_to_cell(5+2*i,0),'=IFERROR(MID(Desciframiento!'+circ_titulo  
                                                            +',FIND(" ",Desciframiento!'+circ_titulo+')+1,256)," ")',center_2)  # Columna del circuito
        worksheet.write(xl_rowcol_to_cell(5+2*i,2),'=IFERROR(MID(Desciframiento!'+fuga_titulo
                                                            +',FIND(" ",Desciframiento!'+fuga_titulo+')+1,256)," ")',center_2)  # Columna de los watts
        
        worksheet.write(xl_rowcol_to_cell(5+2*i,5),'=Desciframiento!'+nota_titulo, formato_nota)
    
def excel(datos,precio,inicio,final,cliente,mes,inicios,finales,periodos,fugas_auto,num_datos):
    
#    print(f'datos = {datos}')
#    print(f'horas = {horas}')
#    print(f'precio = {precio}')
#    print(f'inicio = {inicio}')
#    print(f'final = {final}')
#    print(f'cliente = {cliente}')
#    print(f'mes = {mes}')
#    print(f'inicios = {inicios}')
#    print(f'finales = {finales}')
#    print(f'periodos = {periodos}')
#    print(f'fugas_auto = {fugas_auto}')
    
    nombre = cliente[3:]
    nombre_ = cliente[3:].replace(' ','_')
    direccion = 'D:/01 Findero/'+mes+'/'+cliente+'/Resultados/ResultadosGenerales_'+nombre_+'.xlsx'
    workbook = xlsxwriter.Workbook(direccion)
    
    general(datos,precio,inicio,final,cliente,mes,nombre,workbook)   
    detalles(datos,precio,inicio,final,cliente,mes,nombre,workbook,inicios,finales,periodos,fugas_auto)
    desciframiento(datos,precio,inicio,final,cliente,mes,nombre,workbook,num_datos)
    ahorro(datos,precio,inicio,final,cliente,mes,nombre,workbook)
    modelo_paneles(datos,precio,inicio,final,cliente,mes,nombre,workbook)
    fugas(datos,precio,inicio,final,cliente,mes,nombre,workbook)
            
    workbook.close()

if __name__ == '__main__':
    cliente = '26 Carlos de Habsburgo'
    mes = '08 Agosto'
    inicios = ['2019-08-16', '2019-07-19', '2019-08-16', '2019-08-16']
    finales = ['2019-08-23', '2019-07-26', '2019-08-23', '2019-08-22']
    periodos = [164.37, 164.82, 162.56, 137.15]
    precio = 5.626
    inicio = '2019-08-16'
    final = '2019-08-22'
    datos = {'DATALOG_B08_Habsburgo.CSV': [17.0, 3.5, 122.4, 7.8, 0.2, 134.6, 0.0, 0.0, 0.0, 0.0, 0.0, 0.1], 'DATALOG_B13_Habsburgo.CSV': [0.0, 31.6, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.1], 'DATALOG_COM06_Habsburgo.CSV': [0.2, 0.9, 0.0, 91.7, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0], 'DATALOG_COM19_Habsburgo.CSV': [3.3, 0.0, 0.3, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0]}
    fugas_auto = {'DATALOG_B08_Habsburgo': [0, 0, 87.0, 0, 0, 105.0, 0, 0, 0, 0, 0, 0], 'DATALOG_B13_Habsburgo': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 'DATALOG_COM06_Habsburgo': [0, 6.0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], 'DATALOG_COM19_Habsburgo': [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]}
    num_datos = 100000
    excel(datos,precio,inicio,final,cliente,mes,inicios,finales,periodos,fugas_auto, num_datos)
    
