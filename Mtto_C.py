import os
import sqlite3

import xlwings as xw
import pandas as pd

import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np

#desactivar advertencias de pandas
import warnings
pd.options.mode.chained_assignment = None


wb = xw.Book.caller()
sht = wb.sheets.active
    
app = xw.apps.active
app.api.ActiveWindow.DisplayGridlines = False

    
db_file = os.path.join(os.path.dirname(wb.fullname), "GCOM_MTTO_NB12017.db")
conexion = sqlite3.connect(db_file, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
    #conexion = sqlite3.connect("GCOM_MTTO_NB12017.db")
        
query = "SELECT * FROM maco"
df = pd.read_sql_query(query, conexion)
    
dtf = df[['maco_nb', 'tisu_nb', 'sucu_nb', 'area_nb', 'maco_ano_int', 'maco_mes_int', 'maco_maquinaria_vc']]
dtf[['tisu_nb2', 'sucu_nb2', 'area_nb2']] = dtf[['tisu_nb', 'sucu_nb', 'area_nb']]
dtf_filtrado = dtf
dtf_filtrado['tisu'] = dtf['tisu_nb']
        
dtf_tisu = pd.read_sql_query("SELECT * FROM tisu;", conexion)
dtf_sucu = pd.read_sql_query("SELECT * FROM sucu2;", conexion)
dtf_area = pd.read_sql_query("SELECT * FROM area;", conexion) 
            
mapeo_tisu = dtf_tisu.set_index('tisu_nb')['tisu_nombre_vc'].to_dict()
mapeo_tisu_sigla = dtf_tisu.set_index('tisu_nb')['tisu_sigla_vc'].to_dict()

mapeo_sucu = dtf_sucu.set_index('sucu_nb')['sucu_nombre_vc'].to_dict()
mapeo_area = dtf_area.set_index('area_nb')['area_nombre_vc'].to_dict()

dtf_filtrado['tisu_nb2'] = dtf['tisu_nb2'].map(mapeo_tisu)
dtf_filtrado['sucu_nb2'] = dtf['sucu_nb2'].map(mapeo_sucu)
dtf_filtrado['area_nb2'] = dtf['area_nb2'].map(mapeo_area)
dtf_filtrado['tisu'] = dtf['tisu'].map(mapeo_tisu_sigla)
            
total_registros = len(dtf_filtrado)
conteo_tisu_nb = dtf_filtrado['tisu_nb'].value_counts()
conteo_sucu_nb = dtf_filtrado['sucu_nb'].value_counts()
conteo_area_nb = dtf_filtrado['area_nb'].value_counts()

dtf_filtrado['tisu-sucu'] = dtf_filtrado['tisu'].astype(str) + ' - ' + dtf_filtrado['sucu_nb2'].astype(str)
dtf_filtrado.set_index('maco_nb', inplace=True)
conexion.close()

print(dtf_filtrado)
def combobox():
    wb = xw.Book.caller()
    source = wb.sheets["Source"]

    # Place the database next to the Excel file
    db_file = os.path.join(os.path.dirname(wb.fullname), "chinook.sqlite")

    # Database connection and creation of cursor
    con = sqlite3.connect(
        db_file, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES
    )
    cursor = con.cursor()

    # Database Query
    cursor.execute("SELECT PlaylistId, Name FROM Playlist")

    # Write IDs and Names to hidden sheet
    source.range("A1").expand().clear_contents()
    source.range("A1").value = cursor.fetchall()

    # Format and fill the ComboBox to show Names (Text) and give back IDs (Values)
    # TODO: implement natively in xlwings
    combo = "ComboBox1"
    wb.api.ActiveSheet.OLEObjects(combo).Object.ListFillRange = "Source!{}".format(
        str(source.range("A1").expand().address)
    )
    wb.api.ActiveSheet.OLEObjects(combo).Object.BoundColumn = 1
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnCount = 2
    wb.api.ActiveSheet.OLEObjects(combo).Object.ColumnWidths = 0

    # Close cursor and connection
    cursor.close()
    con.close()

def mtto():
    wb = xw.Book.caller()
    sht = wb.sheets.active

    db_file = os.path.join(os.path.dirname(wb.fullname), "GCOM_MTTO_NB12017.db")
    con = sqlite3.connect(db_file, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
    cursor = con.cursor()
    cursor.execute("SELECT * FROM maco")
    col_names = [col[0] for col in cursor.description]
    rows = cursor.fetchall()
    sht.range("A9").expand().clear_contents()
    sht.range("A9").value = col_names
    if len(rows):
        sht.range("A10").value = rows
    else:
        sht.range("A10").value = "Empty Playlist!"
    cursor.close()
    con.close()
    
def Mtto02():
    wb = xw.Book.caller()
    sht = wb.sheets.active

    db_file = os.path.join(os.path.dirname(wb.fullname), "GCOM_MTTO_NB12017.db")
    conexion = sqlite3.connect(db_file, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)

    # Ejecutar la consulta SQL y guardar los resultados en un DataFrame de pandas
    query = "SELECT * FROM maco"
    df = pd.read_sql_query(query, conexion)
   
    dtf = df[['maco_nb', 'tisu_nb', 'sucu_nb', 'area_nb', 'maco_ano_int', 'maco_mes_int', 'maco_maquinaria_vc']]
    dtf[['tisu_nb2', 'sucu_nb2', 'area_nb2']] = dtf[['tisu_nb', 'sucu_nb', 'area_nb']]
    dtf_filtrado = dtf
    dtf_filtrado['tisu'] = dtf['tisu_nb']
    dtf_tisu = pd.read_sql_query("SELECT * FROM tisu;", conexion)
    dtf_sucu = pd.read_sql_query("SELECT * FROM sucu2;", conexion)
    dtf_area = pd.read_sql_query("SELECT * FROM area;", conexion) 
        
    mapeo_tisu = dtf_tisu.set_index('tisu_nb')['tisu_nombre_vc'].to_dict()
    mapeo_tisu_sigla = dtf_tisu.set_index('tisu_nb')['tisu_sigla_vc'].to_dict()

    mapeo_sucu = dtf_sucu.set_index('sucu_nb')['sucu_nombre_vc'].to_dict()
    mapeo_area = dtf_area.set_index('area_nb')['area_nombre_vc'].to_dict()

    dtf_filtrado['tisu_nb2'] = dtf['tisu_nb2'].map(mapeo_tisu)
    dtf_filtrado['sucu_nb2'] = dtf['sucu_nb2'].map(mapeo_sucu)
    dtf_filtrado['area_nb2'] = dtf['area_nb2'].map(mapeo_area)
    dtf_filtrado['tisu'] = dtf['tisu'].map(mapeo_tisu_sigla)
        
    total_registros = len(dtf_filtrado)
    conteo_tisu_nb = dtf_filtrado['tisu_nb'].value_counts()
    conteo_sucu_nb = dtf_filtrado['sucu_nb'].value_counts()
    conteo_area_nb = dtf_filtrado['area_nb'].value_counts()

    dtf_filtrado['tisu-sucu'] = dtf_filtrado['tisu'].astype(str) + ' - ' + dtf_filtrado['sucu_nb2'].astype(str)  
 
    #pasar datos a excel
    if not dtf_filtrado.empty:
        sht.range("A9").expand().clear_contents()
        sht.range("A9").value = dtf_filtrado.columns.tolist()  # Encabezados
        sht.range("A10").value = dtf_filtrado.values.tolist()  # Datos
    else:
        sht.range("A10").value = "Empty Playlist!"
    
    conexion.close()
    
    #formatos
    enc_range = sht.range("A9").expand('right')
    
    for cell in enc_range:
        cell.api.Font.Name = 'Comic Sans MS'
        cell.api.Font.Size = 12
        cell.api.Font.Bold = True
        #cell.column_width = 25
        cell.color = (255,0,255) 
        cell.api.Font.Color = 0xFFFFFF
        
    enc_range.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    enc_range.autofit()
    
    cuerpo_range = enc_range = sht.range("A9").expand()
    cuerpo_range.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignCenter
    cuerpo_range.autofit()
    
def Mtto03():


    #TABLAS
    # TABLA TISU #######################################################
    dtf_tisu = dtf_filtrado[['tisu_nb2', 'tisu', 'maco_ano_int', 'maco_mes_int']]
    dtf_tisu_por = dtf_tisu.groupby('tisu').size().reset_index(name='counts')
    dtf_tisu_por['por_tisu'] = (dtf_tisu_por['counts'] / total_registros * 100).round(2)

    # TODAS LAS UNIDADES SUCUS #####################################################
    dtf_sucus = dtf_filtrado['tisu-sucu'].value_counts().reset_index()
    dtf_sucus['por_sucus'] = (dtf_sucus['count'] / total_registros * 100).round(2)

    # PREPARAR TABLA SUCU EESS ########################################################
    dtf_sucu_eess = dtf_filtrado[dtf_filtrado['tisu_nb2'] == "ESTACION DE SERVICIO"]['sucu_nb2'].value_counts().reset_index()
    dtf_sucu_eess['por_sucu_eess'] = (dtf_sucu_eess['count'] / total_registros * 100).round(2)

    # PREPARAR TABLA SUCU PE#############################################################
    dtf_sucu_pe = dtf_filtrado[dtf_filtrado['tisu_nb2'] == "PLANTA ENGARRAFADORA"]['sucu_nb2'].value_counts().reset_index()
    dtf_sucu_pe['por_sucu_pe'] = (dtf_sucu_pe['count'] / total_registros * 100).round(2)

    #PREPARAR GRÁFICO
    sns.set_palette("Set2")
    fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(10, 10))
    #fig.patch.set_facecolor('xkcd:light grey')
    fig.patch.set_facecolor('#c2f0f0')
    ##c2f0f0
    fig.suptitle('SOLICITUDES DE MANTENIMIENTO CORRECTIVO', fontsize=14, fontweight='bold', color='blue')

    #GRÁFICAR TISU
    ax1.pie(
        dtf_tisu_por['por_tisu'],
        labels=dtf_tisu_por['tisu'],
        autopct='%1.1f%%',
        startangle=10,
        shadow=True,
        explode=([0.05] * len(dtf_tisu_por)),
        textprops={'fontsize': 6, 'fontweight': 'bold', 'va': 'center'},
        labeldistance=1,
        pctdistance=0.75
        )
    ax1.set_title('SOLICITUDES POR TIPO DE UNIDAD OPERATIVA', fontsize=10, fontweight='bold', color='blue')

    #GRÁFICAR SUCUS
    ax2.pie(
        dtf_sucus['por_sucus'],
        labels=dtf_sucus['tisu-sucu'],
        autopct='%1.1f%%',
        startangle=10,
        shadow=True,
        explode=([0.05] * len(dtf_sucus)),
        textprops={'fontsize': 6, 'fontweight': 'bold', 'va': 'center'},
        labeldistance=1,
        pctdistance=0.75
        )
    ax2.set_title('SOLICITUDES POR UNIDAD OPERATIVA', fontsize=10, fontweight='bold', color='blue')

    #GRÁFICAR SUCU EESS
    ax3.pie(
        dtf_sucu_eess['por_sucu_eess'], 
        labels=dtf_sucu_eess['sucu_nb2'], 
        autopct='%1.1f%%', 
        startangle=10, 
        shadow=True, 
        explode=([0.05] * len(dtf_sucu_eess)),
        textprops={'fontsize': 6, 'fontweight': 'bold', 'va': 'center'}, 
        labeldistance=1, 
        pctdistance=0.75
        )
    ax3.set_title('SOLICITUDES POR ESTACIÓN DE SERVICIO', fontsize=10, fontweight='bold', color='blue')

    #GRÁFICAR SUCU PE
    ax4.pie(
        dtf_sucu_pe['por_sucu_pe'], 
        labels=dtf_sucu_pe['sucu_nb2'], 
        autopct='%1.1f%%', 
        startangle=10, 
        shadow=True, 
        explode=([0.05] * len(dtf_sucu_pe)),
        textprops={'fontsize': 6, 'fontweight': 'bold', 'va': 'center'}, 
        labeldistance=1, 
        pctdistance=0.75
        )
    ax4.set_title('SOLICITUDES POR PLANTA ENGARRAFAORA', fontsize=10, fontweight='bold', color='blue')

    fig.subplots_adjust(wspace=0.2, hspace=0.2)
    #plt.show()
    
    rng = sht.range("B2")
    sht.pictures.add(fig, name='maco_sol', update=True, top=rng.top, left=rng.left, scale=0.7)

def Mtto04():
    pivot_dft_tisu = pd.pivot_table(dtf_filtrado, values='sucu_nb', index='tisu-sucu', columns='maco_ano_int', aggfunc='count', fill_value=0)
    pivot_dft_tisu.reset_index(inplace=True)

    sns.set_theme(style="whitegrid")
    colores = sns.color_palette("Paired")

    fig, ax = plt.subplots(figsize=(10, 5))
    fig.patch.set_facecolor('xkcd:light grey')

    sns.barplot(data=pivot_dft_tisu, x='tisu-sucu', y=2021, label='2021', color=colores[0], alpha=0.6, ci=None, width=0.8)
    sns.barplot(data=pivot_dft_tisu, x='tisu-sucu', y=2022, label='2022', color=colores[1], alpha=0.6, ci=None, width=0.6)
    sns.barplot(data=pivot_dft_tisu, x='tisu-sucu', y=2023, label='2023', color=colores[2], alpha=0.6, ci=None, width=0.4)

    ax.set_xlabel("UNIDADES OPERATIVAS DEL DCCH.", fontsize=12, fontweight='bold')
    ax.set_ylabel("SOLICITUDES DE MTTO. COORECTIVO", fontsize=12, fontweight='bold')
    ax.set_title("ATENCIONES DE MANTENIMIENTO CORRECTIVO POR UNIDAD OPERATIVA", fontsize=14, fontweight='bold')

    plt.xticks(rotation=90, fontsize=8)

    # Añadir etiquetas de valor sobre cada barra
    for p in ax.patches:
        if not pd.isna(p.get_height()):
            ax.annotate(f'{int(p.get_height())}', (p.get_x() + p.get_width() / 2, p.get_height()), ha='center', va='bottom', fontsize=12, fontweight='bold')
        else:
            ax.annotate('0', (p.get_x() + p.get_width() / 2, 0), ha='center', va='bottom', fontsize=10, fontweight='bold')

    plt.legend(title="Años", fontsize=12, title_fontsize=10)

    rng = sht.range("B2")
    sht.pictures.add(fig, name='maco_sol_MES', update=True, top=rng.top, left=rng.left, scale=1)

    #plt.tight_layout()
    plt.show()
    
def Mtto05():
    conteo_sucu_nb = dtf_filtrado.groupby(['maco_ano_int', 'maco_mes_int'])['sucu_nb'].count().reset_index()
    sns.set_style("darkgrid", {"grid.color": ".6", "grid.linestyle": ":"}) 

    g = sns.FacetGrid(conteo_sucu_nb, col='maco_ano_int', col_wrap=3, height=5)
    g.map_dataframe(sns.regplot, x='maco_mes_int', y='sucu_nb', ci=None, color='b')

    # Añadir líneas punteadas
    for ax in g.axes:
        year_match = re.search(r'\d{4}', ax.get_title())
        if year_match:
            year = int(year_match.group())
            data_year = conteo_sucu_nb[conteo_sucu_nb['maco_ano_int'] == year]
            sns.lineplot(x=data_year['maco_mes_int'], y=data_year['sucu_nb'], ax=ax, linestyle='dashed', color='r')

        ax.set_xticks(range(len(meses_dict)))
        ax.set_xticklabels(meses_dict.values(), rotation=45)
        ax.set_xlabel('Meses')
        ax.set_ylabel('Conteo de sucu_nb')

        # Obtener el año del título utilizando una expresión regular
        title = ax.get_title()
        year_match = re.search(r'\d{4}', title)
        if year_match:
            year = int(year_match.group())
            ax.set_title(f'Conteo de sucu_nb por Mes en el Año {year}')

        # Agregar etiquetas a los puntos con los valores
        for x, y, value in zip(data_year['maco_mes_int'], data_year['sucu_nb'], data_year['sucu_nb']):
            ax.annotate(f'{value}', (x, y), textcoords='offset points', xytext=(0, 10), ha='center', fontsize=8, color='black')

    plt.suptitle('Comportamiento de solicitudes mensuales', fontsize=16, fontweight='bold')

    rng = sht.range("B2")
    sht.pictures.add(g, name='maco_sol_MES', update=True, top=rng.top, left=rng.left, scale=1)

    plt.tight_layout()
    plt.show()
