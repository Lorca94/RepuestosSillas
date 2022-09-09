import pandas as pd
from matplotlib import pyplot as plt
from reportlab.pdfgen import canvas


class Silla:
    def __init__(self, nombre):
        self.nombre = nombre

    brazo = 0
    rueda = 0
    piston = 0
    basculante = 0
    estrella = 0


# Abre el excel de la ruta indicada y toma los datos de la columna
def obtenerdatoscolumna(ruta: str, columna: str):
    xls = pd.read_excel(ruta)
    info = xls[columna].values
    return info


# Con los nombres indicados crea objetos
def crearobjetos(arrayNombres: list):
    newObjts = []
    for i in arrayNombres:
        silla = Silla(i.lower())
        newObjts.append(silla)
    return newObjts


# Comprueba si la avería coincide con alguna pieza reparable y devuelve la palabra localizada
def esReparable(averia: str, palabrasClave: list):
    for palabra in palabrasClave:
        if palabra in averia:
            return palabra
    return None


# Main
def encontrarMarca(marca: str, nombresComparar: list):
    for nombre in nombresComparar:
        if nombre.lower() in marca.lower():
            return nombre.lower()


def encontrarSilla(actualMarca: str, allSillas: list):
    for silla in allSillas:
        if actualMarca in silla.nombre:
            return silla


def crearPdf(allSillas: list, dataAverias: list, sillasSalvables: int):
    pdf = canvas.Canvas('Informe de Sillas.pdf')
    ejeX = 80
    ejeY = 750
    for final in allSillas:
        if not (final.piston == 0 and final.estrella == 0 and
                final.rueda == 0 and final.brazo == 0 and
                final.basculante == 0):
            if ejeY <= 75:
                ejeX = 100
                ejeY = 750
                pdf.showPage()
            pdf.drawString(ejeX, ejeY, 'Sobre la marca: ' + final.nombre.capitalize())
            ejeY = ejeY - 25
            pdf.drawString(ejeX, ejeY,
                           'Las siguientes piezas podrían haber sido sustituidas para evitar la llegada del artículo a')
            ejeY = ejeY - 25
            pdf.drawString(ejeX, ejeY, 'PcComponentes:')
            ejeY = ejeY - 25
            pdf.drawString(ejeX, ejeY, 'Estrellas: ' + str(final.estrella))
            ejeY = ejeY - 25
            pdf.drawString(ejeX, ejeY, 'Kit de ruedas: ' + str(final.rueda))
            ejeY = ejeY - 25
            pdf.drawString(ejeX, ejeY, 'Brazos: ' + str(final.brazo))
            ejeY = ejeY - 25
            pdf.drawString(ejeX, ejeY, 'Pistón: ' + str(final.piston))
            ejeY = ejeY - 25
            pdf.drawString(ejeX, ejeY, 'Estrellas: ' + str(final.basculante))
            ejeY = ejeY - 25
            pdf.drawImage(final.nombre+'.png', ejeX+220, ejeY-55, 250, 250, preserveAspectRatio=True, mask='auto')
            ejeY = ejeY - 50

    pdf.drawString(ejeX, ejeY, 'Sobre un total de: ' + str(
        dataAverias.size) + '. Se podría haber solucionado con el envío de repuestos')
    ejeY = ejeY - 25
    pdf.drawString(ejeX, ejeY, 'un total de: ' + str(sillasSalvables) + ' casos')
    ejeY = ejeY - 25
    porcentaje = (sillasSalvables / dataAverias.size) * 100
    porcentaje = round(porcentaje, 2)
    pdf.drawString(ejeX, ejeY, 'Lo que hace un porcentaje de: ' + str(porcentaje) + '%')
    pdf.save()


def crearGraficas(allSillas: list):
    for final in allSillas:
        if not (final.piston == 0 and final.estrella == 0 and
                final.rueda == 0 and final.brazo == 0 and
                final.basculante == 0):
            fig, ax = plt.subplots()
            ax.set_title('Piezas ' + final.nombre)
            ax.bar(['Estrellas', 'Ruedas', 'Brazos', 'Pistón', 'Basculante'],
                   [final.estrella, final.rueda, final.brazo, final.piston, final.basculante])
            plt.savefig(final.nombre + '.png', dpi=fig.dpi)


def main():
    # Palabras que nos servirán para identificar la pieza que necesitamos para reponer
    palabrasClave = ['estrella', 'patas', 'base', 'soporte', 'rueda', 'brazo', 'piston', 'basculante', 'no sube',
                     'tubo de metal', 'baja sola', 'no sube', 'bloqueado']

    nombresComparar = ['Newskill', 'Corsair', 'Drift', 'Razer', 'Noblechair', 'MSI', 'Tempest', 'Sharkoon', 'Asus',
                       'Nacon', 'Woxter', 'Playseat', 'Forgeon', 'Aerocool', 'Equip', 'Genesis', 'HP', 'Next Level',
                       'NGS', 'Oplite', 'Owlotech', 'Thermaltake', 'Trust', 'ZEN']

    directorio = "C:/Users/RMA-BANCADA-7/Documents/RMAS de Sillas desde 01_01_2022.xls"
    # Se obtiene datos
    dataMarcas = obtenerdatoscolumna(directorio, 'productName')
    dataAverias = obtenerdatoscolumna(directorio, 'info')
    # Se crean todos los objetos
    allSillas = crearobjetos(nombresComparar)
    sillasSalvables: int = 0
    # Se itera sobre averias
    for i in range(dataAverias.size):
        # Se obtiene la marca actual que estamos comprobando
        actualMarca: str = encontrarMarca(dataMarcas[i], nombresComparar)
        # Se comprueba que marca existe
        if actualMarca:
            # Traemos el objeto silla que se va a utilizar
            actualSilla: Silla = encontrarSilla(actualMarca, allSillas)

            repuesto: str = esReparable(dataAverias[i], palabrasClave)
            if repuesto:
                sillasSalvables = sillasSalvables + 1
                if repuesto == 'estrella' or repuesto == 'patas' or repuesto == 'soporte':
                    actualSilla.estrella = actualSilla.estrella + 1
                if repuesto == 'rueda':
                    actualSilla.rueda = actualSilla.rueda + 1
                if repuesto == 'brazo':
                    actualSilla.brazo = actualSilla.brazo + 1
                if repuesto == 'piston' or repuesto == 'tubo' or repuesto == 'no sube' or repuesto == 'tubo de metal' \
                        or repuesto == 'baja sola' or repuesto == 'no sube' or repuesto == 'bloqueado':
                    actualSilla.piston = actualSilla.piston + 1
                if repuesto == 'basculante':
                    actualSilla.basculante = actualSilla.basculante + 1

    crearGraficas(allSillas)
    crearPdf(allSillas, dataAverias, sillasSalvables)


if __name__ == '__main__':
    main()
