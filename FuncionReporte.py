import matplotlib.pyplot as plt
from docxtpl import DocxTemplate, InlineImage
import pandas as pd


envio = pd.read_csv('delivery_dataset.csv', sep=';')
envioReal = envio.copy()

deliveryDelay = envioReal['Shipment_Delay'].mean() 
productTotal = len(envioReal)
pedidosPerdidos = envio['Actual_Shipment_Time'].isna().sum()

def crearReporte():
    fig, ax = plt.subplots()

    doc = DocxTemplate('imagetemplate.docx')
    context = {
    'imagen1': InlineImage(doc, 'grafica1.jpg'),
    'imagen2': InlineImage(doc, 'grafica2.jpg'),
    'imagen3': InlineImage(doc, 'grafica3.jpg'),
    'delay': deliveryDelay,
    'ptotal': productTotal,
    'pperdidos': pedidosPerdidos
    }

    doc.render(context)
    doc.save( "Reporte_Mensual_Enero.docx")

