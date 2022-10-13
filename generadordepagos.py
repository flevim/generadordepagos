import tkinter as tk  
from tkcalendar import Calendar
from tkinter import ttk,filedialog, messagebox
import pandas as pd 
from datetime import date
import sys 

#Ventana Principal 
root = tk.Tk()

#Ajustes básicos
root.title("Quick Delivery")
root.geometry('1000x950')
root.maxsize(900,720)


#Global
filename = ''
since_date = ''
pd.options.display.max_columns = None
###########################
#     Funciones varias    #
###########################

def get_data_by_date(s_date, data):
    cond = data["DELIVERY DATE"] >= s_date
    return data.loc[cond].sort_values(['DELIVERY DATE'], ascending=True)


def get_subtotal(data):
    data.loc[data['PAYMETHOD'] != 'Efectivo', 'SUBTOTAL'] = 0
    return data

def less_th_ten(d): 
    return '0' + d if int(d) < 10 else d

def change_format_date(date): 
    sp_date = date.split('-')

    day = sp_date[0]
    month = sp_date[1]
    year = '20' + sp_date[2]

    return '{}-{}-{}'.format(year, month, day)

    

def gen_out_file_r(file, date):
    try:

        rest_data = pd.read_excel(file,  usecols="V:W,AK,A,AN,AX,AY")

        print("Número de filas archivo {}: {}".format(file, rest_data.shape))

        date = change_format_date(date)

        rest_by_date = get_data_by_date(date, rest_data)

        #Calculos varios 
        rest_by_date["SUBTOTAL (C/DESCUENTO)"] = round(rest_by_date["SUBTOTAL"] - rest_by_date["DISCOUNT"])
        rest_by_date["Comision Neta (15%)"] = round(rest_by_date["SUBTOTAL (C/DESCUENTO)"] * 0.15)
        #rest_by_date["Comision Neta (C/IVA)"] = round(rest_by_date["Comision Bruta (15%)"] / 1.19)
        rest_by_date["IVA"] = round(rest_by_date["Comision Neta (15%)"] * 0.19)
        rest_by_date["Comision Total (c/iva)"] = round(rest_by_date["Comision Neta (15%)"] + rest_by_date["IVA"])
        rest_by_date["TOTAL RESTAURANTE (85%)"] = round(rest_by_date["SUBTOTAL"] * 0.85)

        # Insertar columnas vacías 
        rest_by_date['Banco'] = None
        rest_by_date['Tipo Cuenta'] = None
        rest_by_date['Numero Cuenta'] = None
        rest_by_date['Mail'] = None
        rest_by_date['Fecha Pago'] = None
        rest_by_date['Tipo Pago'] = None
        rest_by_date['Banco Origen'] = None
        rest_by_date['Cuenta Origen'] = None
        rest_by_date['Comprobante Pago'] = None
        rest_by_date['Monto Deposito Rest'] = None
        rest_by_date['Fecha Deposito'] = None
        rest_by_date['Comprobante Deposito'] = None
        rest_by_date['Medio Deposito'] = None

        rest_by_date = rest_by_date.rename(columns = {'ID': 'ID ORDEN'})

        rest_by_date = rest_by_date[['BUSINESS ID', 'BUSINESS NAME', 'ID ORDEN', 'SUBTOTAL (C/DESCUENTO)', 'TOTAL RESTAURANTE (85%)', 
                                            'Comision Neta (15%)', 'IVA', 'Comision Total (c/iva)', 'Banco', 'Tipo Cuenta',
                                            'Numero Cuenta', 'Mail', 'Fecha Pago', 'Tipo Pago', 'Banco Origen', 'Cuenta Origen', 'Comprobante Pago', 
                                            'Monto Deposito Rest', 'Fecha Deposito', 'Comprobante Deposito', 'Medio Deposito']]

        return rest_by_date

    except: 
        messagebox.showerror("Error", "Ocurrió un error al extraer datos del archivo excel.")
        sys.exit(1)

def gen_out_file_q(file, date):
    try:

        quickers_data = pd.read_excel(file,  usecols="O:Q,AG,AK,A,AI,AN,AR,AS,AY")
        
        print("Número de filas archivo {}: {}".format(file, quickers_data.shape))

        date = change_format_date(date)
        
        quickers_by_date = get_data_by_date(date, quickers_data)

        quickers_by_date = get_subtotal(quickers_by_date)

        quickers_by_date["TOTAL QUICKERS"] = round(quickers_by_date["DELIVERY FEE"] + quickers_by_date["DRIVER TIP"])   
        quickers_by_date["Retencion (11,5%)"] = round(quickers_by_date["TOTAL QUICKERS"] * 0.115)
        quickers_by_date["Total Honorarios (88,5%)"] = round(quickers_by_date["TOTAL QUICKERS"] - quickers_by_date["Retencion (11,5%)"])

        # columna DRIVER (Concatenación ID QUICKER+NOMBRE+APELLIDO)
        quickers_by_date['DRIVER ID'] = quickers_by_date['DRIVER ID'].fillna(0).astype(int)

        quickers_by_date['DRIVER ID'] = quickers_by_date['DRIVER ID'].replace(0, '', regex=True)
        quickers_by_date[['DRIVER NAME', 'DRIVER LASTNAME']] = quickers_by_date[['DRIVER NAME', 'DRIVER LASTNAME']].fillna('')
        quickers_by_date["DRIVER"] = quickers_by_date["DRIVER ID"].map(str) + " " + quickers_by_date["DRIVER NAME"].map(str) + " " + quickers_by_date["DRIVER LASTNAME"].map(str)
    
        # Insertar columnas vacías 
        quickers_by_date['Banco'] = None
        quickers_by_date['Tipo Cuenta'] = None
        quickers_by_date['Numero Cuenta'] = None
        quickers_by_date['Mail'] = None
        quickers_by_date['Fecha Pago'] = None
        quickers_by_date['Tipo Pago'] = None
        quickers_by_date['Banco Origen'] = None
        quickers_by_date['Cuenta Origen'] = None
        quickers_by_date['Comprobante Pago'] = None
        quickers_by_date['Monto Deposito Quick'] = None
        quickers_by_date['Fecha Deposito'] = None
        quickers_by_date['Comprobante Deposito'] = None
        quickers_by_date['Medio Deposito'] = None

        quickers_by_date = quickers_by_date.rename(columns = {'ID': 'ID ORDEN', 'SUBTOTAL': 'SUBTOTAL A DEVOLVER'})


        quickers_by_date = quickers_by_date[['DRIVER ID', 'DRIVER NAME', 'DRIVER LASTNAME', 'DRIVER', 'ID ORDEN', 'DELIVERY FEE', 'DRIVER TIP', 'SUBTOTAL A DEVOLVER',
                                'TOTAL QUICKERS','Retencion (11,5%)', 'Total Honorarios (88,5%)', 'Banco', 'Tipo Cuenta',
                                'Numero Cuenta', 'Mail', 'Fecha Pago', 'Tipo Pago', 'Banco Origen', 'Cuenta Origen', 'Comprobante Pago', 
                                'Monto Deposito Quick', 'Fecha Deposito', 'Comprobante Deposito', 'Medio Deposito']]

        return quickers_by_date

    except:
        messagebox.showerror("Error", "Ocurrió un error al extraer datos del archivo excel.")
        sys.exit(1) 

def download_file_r():
    global filename
    global since_date

    if not(len(filename)): 
        messagebox.showerror("Error", "Porfavor ingrese un archivo de entrada primero")
    
    elif not(len(since_date)):
        messagebox.showerror("Error", "Porfavor ingrese la fecha de inicio primero")
    
    else: 
        out_df = gen_out_file_r(filename, since_date)
      
    
        try: 
            savefile = filedialog.asksaveasfilename(filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los Archivos", "*.*")))
            out_df.to_excel(savefile+'.xlsx', index=False, sheet_name="Liquidación {}".format(since_date))
            messagebox.showinfo(message="Su archivo ha sido generado exitosamente", title="Exito" )
            
        except:
            messagebox.showerror("Error", "Ha ocurrido un error al generar el archivo")
            



def download_file_q():
    global filename
    global since_date

    if not(len(filename)): 
        messagebox.showerror("Error", "Porfavor ingrese un archivo de entrada primero")
    
    elif not(len(since_date)):
        messagebox.showerror("Error", "Porfavor ingrese la fecha de inicio primero")
    
    else: 
        out_df = gen_out_file_q(filename, since_date)
    
        try: 
            savefile = filedialog.asksaveasfilename(filetypes=(("Archivos Excel", "*.xlsx"), ("Todos los Archivos", "*.*")))
            out_df.to_excel(savefile+'.xlsx', index=False, sheet_name="Liquidación {}".format(since_date))
            messagebox.showinfo(message="Su archivo ha sido generado exitosamente", title="Exito" )
        except:
            messagebox.showerror("Error", "Ha ocurrido un error al generar el archivo")

def grad_date():
    global filename
    global since_date 
     
    
    if len(filename):
        since_date = cal.get_date()
        confirm = messagebox.askquestion("Confirmar Fecha", "Haz escogido "+since_date+"\n¿Deseas confirmar esta fecha?", icon='info')
        
        if confirm == 'yes':
            date.config(text = "Fecha Seleccionada: " + since_date, font=("Font", 15))
            

        else:
            since_date = ''
    else:
        messagebox.showerror("Error", "Porfavor ingrese su archivo primero")

    

def upload_file(): 
    global filename
    filename = filedialog.askopenfilename(
        title="Seleccionar archivo generado por Ordering",
        filetypes=[('Archivos .xlsx', '.xlsx')]
    )


    if filename: 
        namefile.config(text="Archivo Seleccionado: {}".format(str(filename.split('/')[-1])), font=("Font", 13))



###########################
#       Widgets           #
###########################
      
content = tk.Frame(root)
lbl_title = tk.Label(content, text="GENERADOR DE ARCHIVOS DE PAGO", bg="#9c9c9c", fg="#000", relief="raised", font=("Font", 20, 'bold')) 
help_btn = tk.Button(content, text="?")

#UI Subir archivo
frame_upload = ttk.Labelframe(content, text="Seleccione un archivo de pedidos", width=100, height=100)
upload_btn = tk.Button(frame_upload, text="Subir archivo", command=upload_file)
namefile = tk.Label(frame_upload, text="")

#UI Seleccionar fecha
frame_date_btns = ttk.Labelframe(content, text="Seleccione una fecha de inicio para generar el archivo de pago")
cal = Calendar(frame_date_btns, selectmode = 'day',
               year = 2020, month = 11,
               day = 22, mindate=date(2020,11,19), 
               maxdate=date.today(), locale='es_CL')

subframe_btn = tk.Frame(frame_date_btns)
btn_date = tk.Button(subframe_btn, text="Seleccionar Fecha", command=grad_date)
date = tk.Label(subframe_btn, text = "")


#UI Descarga archivo
frame_download = ttk.LabelFrame(content, text="Descargar archivos generados")
btn_quicker = tk.Button(frame_download, text="Archivo Quickers", command=download_file_q)
btn_rest = tk.Button(frame_download, text="Archivo Restaurantes", command=download_file_r)




#############################
#     Estructura de Widgets #
#############################


content.pack()
lbl_title.pack(ipadx=270, ipady=18, fill='x', expand=True)
#help_btn.grid(column=0, row=0)

frame_upload.pack(ipadx=95, pady=6, fill='x')
upload_btn.pack(pady=10, ipady=6, ipadx=30)
namefile.pack(pady=4, padx=10)

frame_date_btns.pack(pady=5, fill='x')
cal.pack(pady=10, side="left", padx=30, ipady=80, ipadx=40)
subframe_btn.pack(side="right")
btn_date.pack(ipady=10, ipadx=30, padx=120)
date.pack(pady=4)   

frame_download.pack(pady=5, ipady=30, ipadx=90)
btn_quicker.pack(side="left", ipadx=20, ipady=6, padx=20)
btn_rest.pack(side="right", ipadx=20, ipady=6, padx=20)


root.mainloop()
