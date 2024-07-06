#pip install openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker

#pd.options.display.float_format = '{:,.0f}'.format
# Configurar el formateador para mostrar números completos (no exponenciales) en el eje Y
formatter = ticker.ScalarFormatter(useOffset=False)
formatter.set_scientific(False)
#leemos el excel
dataframe = pd.read_excel(r"E:\Maestria y Diplomados\DIPLOMADO PYTHON\Full info\Modulo 6 aplicaciones de python en machine learning y ciencia de datos\Clase 2\ventas.xlsx", "datos")
dataframe.dropna(inplace = True) #Elimina nulos
dataframe = dataframe.drop_duplicates() #Elimina duplicados
#separamos en otra columna el mes, para poder sacar totales por mes
dataframe['mes'] = dataframe['Día'].dt.month
#filtramos para traernos solo las dos columnas: mes y total vendido
filtradospormes = dataframe.filter(['mes', 'Total Vendido'])
filtradospormes['Total Vendido'] = filtradospormes['Total Vendido'].astype(int)
#agrupamos por mes, haciendo la sumatoria
agrupados = filtradospormes.groupby(by='mes').sum(True)
print(agrupados)
#se muestra la grafica
agrupados.plot()
# Aplicar el formateador al eje X
plt.gca().yaxis.set_major_formatter(formatter)
plt.ylabel("Ventas Totales")
plt.xlabel("Mes del año")
plt.show()

#filtramos las columnas: vendedor, total vendido y mes
filtradosporvendedor = dataframe.filter(['Vendedor', 'Total Vendido', 'mes'])
#primero agrupamos por vendedor para ver quien tiene mas ventas
agrupados_por_vendedor = filtradosporvendedor.groupby(by='Vendedor').sum(True)
filtro_vendedor = agrupados_por_vendedor.sort_values(by=['Total Vendido','Vendedor'], ascending=False).iloc[0:1,0:]
vendedor_top = filtro_vendedor.axes[0].values[0]
print('Top 1 Vendedor',vendedor_top)
#ya que tenemos al vendedor top 1, lo filtramos solo para ese vendedor y solo inlcuimos mes y total vendido
filtradosporvendedor_pormes = filtradosporvendedor[filtradosporvendedor['Vendedor'] == vendedor_top].filter(['mes','Total Vendido'])
#agrupamos por mes para graficar
agrupados_por_vendedor_por_mes = filtradosporvendedor_pormes.groupby(by='mes').sum(True)
print(agrupados_por_vendedor_por_mes)
#se muestra la grafica
agrupados_por_vendedor_por_mes.plot()
plt.show()


#dividir los datos por mes
for i in range(1, 13):
    #se filtra lo del mes i
    filtro_por_mes = dataframe[dataframe['mes'] == i]
    #sacamos el nombre del mes
    month_name = filtro_por_mes.head(1)['Día'].dt.month_name()
    month = month_name.values[0]
    #lo mandamos a un excel ya filtrado
    filename_month = 'c:\\programaspython\\' + month + '.xlsx'
    filtro_por_mes.to_excel(filename_month,sheet_name=month)

print("Termina")


