

import numpy as np
import pandas as pd
import datetime
import uuid
from datetime import datetime
from docxtpl import  DocxTemplate

productos = pd.read_excel('productos.xlsx')
productos.set_index('product')
productosTienda = list(productos['product'])



def current_date_format(date):
    day = date.day
    month = date.month
    year = date.year
    messsage = "{}-{} del {}".format(day, month, year)

    return messsage

def crearCotizacion( nombre, productosz, cantidades, destino, Carrier_name, Planned_Delivery_Time, Distance  ):
        word_template_path ="template.docx"        
        dates = datetime.now()
        fecha = current_date_format(dates)
        source = 'LAX'
        precios = []
        unidades = []
        total = []
        subtotal = 0
        itmbs = 0
        totalitmbs = 0

        for x in productosz:
            aux = productos.loc[ productos['product'] == x, 'price' ]   
            aux2 = productosTienda.index( x ) 
            precios.append( aux[aux2] )

            auxP = productos.loc[ productos['product'] == x, 'unidades' ]
            aux2P = productosTienda.index( x ) 
            unidades.append( auxP[aux2P] )
            

        for x in range(6):
            if( len(productosz) < 5 ):
                productosz.append('')
                precios.append('')
                cantidades.append('')
                unidades.append('')
                total.append(0)
            else:
                break;
        
        for x in enumerate(precios):
            if( precios[ x[0] ] != '' ):
                total[ x[0] ] = precios[ x[0] ] * cantidades[ x[0] ]
                
   
        
        subtotal = sum( total )
        itmbs = subtotal * 0.07
        totalitmbs = int(subtotal + itmbs)

        print( subtotal )

        variables = {'producto': productosz, 
                     'cantidad': cantidades, 
                     'destino': destino, 
                     'fecha': fecha,
                     'precio': precios,
                     'unidades': unidades,
                     'source': source,
                     'carrier': Carrier_name,
                     'planned_delivey': Planned_Delivery_Time,
                     'distance': Distance,
                     'nombre': nombre,
                     'total': total,
                     'subtotal': subtotal,
                     'itmbs': itmbs,
                     'totalitmbs': totalitmbs
                     
                     }
        
        doc = DocxTemplate(word_template_path)
        doc.render( variables )
        doc.save( f"Cotizacion-.docx")