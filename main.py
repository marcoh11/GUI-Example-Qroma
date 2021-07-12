#Importando las librerias
#install 
from fpdf import FPDF
import tkinter as tk
import webbrowser,os
#Importando clases
from tkinter import Label, Menu, PhotoImage, messagebox,ttk,Scrollbar
from tkinter.constants import ANCHOR, BOTH, BOTTOM, CENTER, E, LEFT, N, NSEW, RIGHT, S, W
#Creando las variables de datos
emisor= []
area=[]
descripcion= []
prioridad= []
estado=[]
#contador
i=0
#Creando Ventana
ventana = tk.Tk()
ventana.geometry("1080x640")
ventana.resizable(0,0)
ventana.title("QROMA | CONTROL DE MANTENIMIENTO - ÑAÑA")
#Poniendo Icono
ventana.iconbitmap('img/qroma.ico')
#Crear menu
menubar = Menu(ventana)
ventana.config(menu=menubar,bg="#0070BA")
#Submenus  

#filemenu
filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Mostrar OT",command=lambda:abrirCarpeta())
filemenu.add_separator()
filemenu.add_command(label="Salir", command=ventana.quit)
#helpmenu
helpmenu = Menu(menubar, tearoff=0)
helpmenu.add_command(label="Acerca de..." , command=lambda:abrirPagina())
#Mostrar
menubar.add_cascade(label="Archivo", menu=filemenu)
menubar.add_cascade(label="Ayuda", menu=helpmenu)





#Creando Frames FORMULARIOS
#Formulario Title
formTitle=tk.Frame(ventana)
formTitle.config(bg="#0070BA")
formTitle.grid(row=0,column=1,sticky=NSEW)
#Form Solicitud
formSolicitud=tk.Frame(ventana)
formSolicitud.config(bg="#0070BA")
formSolicitud.grid(row=1,column=0,sticky=NSEW,padx=25)
#Formulario del planner
formPlanner=tk.Frame(ventana)
formPlanner.config(bg="#0070BA")
formPlanner.grid(row=1,column=1,sticky=N)
#Formulario de OT  
formOt=tk.Frame(ventana)
formOt.config(bg="#0070BA")
formOt.grid(row=1,column=2,sticky=N,padx=25)
#Formulario de la tabla
formTabla=tk.Frame(ventana)
formTabla.config(bg="#0070BA")
formTabla.grid(row=2,columnspan=3,column=0)

#Funciones-----------------------------------------------------------------------------------------   
def abrirCarpeta():
    RUTA_CARPETA = "OT"
    webbrowser.open(os.path.realpath(RUTA_CARPETA))
def crearOT():
    try:
        elemento=tree.focus()
        i=elemento[len(elemento)-1]
        i=int(i)-1
        if messagebox.askyesno(message="Se creara una OT con los datos ingresados ¿Desea proceder?",title="Confirmar"):
            pdf=FPDF(format='A4')
            pdf.add_page()
            #Cuerpo
            pdf.set_font('Arial','', 9)
            pdf.image('img/pdflogo.jpg',x=30,y=27,w=17,h=17)
            pdf.rect(x=25,y=25,w=160,h=200,style='border only')
            #salto hacia abajo
            pdf.ln(17)
            #Pone espacio
            pdf.cell(56)
            pdf.cell(100,10,"CORPORACION PERUANA DE PRODUCTOS QUIMICOS S.A",ln=1)
            pdf.cell(80)
            pdf.cell(100,10,"MANTENIMIENTO - ÑAÑA",ln=2)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(100,10,"ORDEN DE TRABAJO N° 0"+str(i),ln=1)
            pdf.set_font('Arial','', 9)
            pdf.cell(30)
            pdf.cell(50,7,"Emitido por: ",border=1)
            pdf.cell(80,7,emisor[i],border=1,ln=1)
            pdf.cell(30)
            pdf.cell(50,7,"Tiempo Estimado",border=1)
            pdf.cell(40,7,enTiempo.get(),border=1,ln=1)
            pdf.cell(30)
            pdf.cell(50,7,"Tiempo Real",border=1)
            pdf.cell(40,7,enTiempoOt.get(),border=1,ln=1)
            pdf.cell(30)
            pdf.cell(50,7,"Prioridad:",border=1)
            pdf.cell(40,7,prioridad[i],border=1,ln=1)
            pdf.cell(30)
            pdf.cell(50,7,"Area:",border=1)
            pdf.cell(40,7,area[i],border=1,ln=1)
            pdf.cell(30)
            pdf.cell(50,7,"Equipo de mantenimiento:",border=1)
            pdf.cell(40,7,cbEquipo.get(),border=1,ln=1)
            pdf.cell(30)
            pdf.cell(50,7,"Tecnico:",border=1)
            pdf.cell(40,7,enTecnico.get(),border=1,ln=1)
            pdf.cell(30)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(50,7,"Repuestos:")
            pdf.set_font('Arial','', 9)
            pdf.cell(40,7,enRepuesto.get(),ln=1)
            pdf.cell(30)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(50,7,"Uso de EPPS:")
            pdf.set_font('Arial','', 9)
            pdf.cell(40,7,enEpp.get(),ln=1)
            pdf.cell(30)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(50,7,"Presupuestos:")
            pdf.set_font('Arial','', 9)
            pdf.cell(40,7,enRepuesto.get(),ln=1)
            pdf.cell(30)
            pdf.cell(130,5,"Descripcion del trabajo",border=1,align='C',ln=2,)
            pdf.cell(130,14,descripcion[i],border=1,ln=2)
            pdf.cell(130,5,"Observacion en el mantenimiento",border=1,align='C',ln=2,)
            pdf.cell(130,14,enDescripcionOT.get(),border=1,ln=2)
            #N° Pagina
            pdf.set_y(15)
            pdf.set_font('Arial', 'I', 8)
            pdf.cell(0, 10, 'Pagina ' + str(pdf.page_no()), 0, 0, 'C')
            pdf.output('OT/OT_0'+str(i)+'.pdf')
            estado[i]="OT"
            tree.insert("",i,text=("N°  "+ str(i)),values=(emisor[i],area[i],descripcion[i],prioridad[i],estado[i]))
            tree.delete(elemento)
            messagebox.showinfo("Aviso","Orden de Trabajo creada con éxito")
    except:
            messagebox.showwarning("Aviso", "Debe seleccionar una solicitud para crear la OT")   
def abrirPagina():
    webbrowser.open_new("https://www.qroma.com.pe/nosotros/")
def agregar():
    global i
    emisor.append(enEmisor.get())
    area.append(cbArea.get())
    descripcion.append(enDescripcion.get())
    prioridad.append(cbPrioridad.get())
    estado.append("Solicitud")
    tree.insert("",i,text=("N°  "+ str(i)),values=(emisor[i],area[i],descripcion[i],prioridad[i],estado[i]))
    i=i+1
    limpiar()
def limpiar():
    enEmisor.delete(0, 'end')
    cbArea.set('')
    enDescripcion.delete(0, 'end')
    cbPrioridad.set('')
    enEmisor.focus()
def eliminar():
    try:
        elemento=tree.focus()
        i=elemento[len(elemento)-1]
        i=int(i)-1
        if messagebox.askyesno(message="La solicitud se eliminara  ¿Desea proceder?",title="Confirmar"):
            messagebox.showinfo("Aviso","Solicitud borrada")
            tree.delete(elemento)
            emisor.pop(i)
            descripcion.pop(i)
            prioridad.pop(i)
            area.pop(i)
    except:
            messagebox.showwarning("Aviso", "No se ha seleccionado ningun elemento")
def editar():
    try:
    #Obtenemos todo lo que tiene ese row
        elemento=tree.focus()
        i=elemento[len(elemento)-1]
        i=int(i)-1
        enEmisor.insert(0,tree.set(elemento,"#1"))
        cbArea.set(tree.set(elemento,"#2"))
        cbPrioridad.set(tree.set(elemento,"#4"))
        enDescripcion.insert(0,tree.set(elemento,"#3"))
        tree.delete(elemento)
        emisor.pop(i)
        descripcion.pop(i)
        prioridad.pop(i)
        area.pop(i)
        estado.pop(i)
    except:
        messagebox.showwarning("Aviso", "No se ha seleccionado ningun elemento")   
def comprobar():
    #Comprobamos el nombre del emisor
    comNombre=str(enEmisor.get())
    #Quitamos los espacios
    comNombre=comNombre.replace(" ", "")
    #Comprobar
    if (comNombre.isalpha() and (cbArea.get()) and (enDescripcion.get()) and (cbPrioridad.get())):
        if messagebox.askyesno(message="¿Los datos se registraran desea continuar?",title="Confirmar"):
            messagebox.showinfo("Aviso","Se ha registrado correctamente")
            agregar()
    else:
        print()
        messagebox.showwarning("Aviso", "Por favor complete todos los campos de forma correcta")

#Fin Funciones-------------------------------------------------------------------------------------------
#Elementos del formTitle
lblTitle=tk.Label(formTitle,text="CORPORACION PERUANA DE PRODUCTOS QUIMICOS S.A.")
lblTitle.config(bg="#0070BA",fg="#FFFFFF",font=("Bahnschrift SemiBold",14))
lblTitle.grid(row=0,column=1,sticky=NSEW,ipady=50)
imgLogo=PhotoImage(file="img/logo.jpg")
lblImagen=tk.Label(ventana,image=imgLogo)
lblImagen.config(bg="#0070BA")
lblImagen.grid(row=0,column=0,sticky=W)
#Elementos del formSolicitud
lbl=tk.Label(formSolicitud, text="SOLICITUDES DE MANTENIMIENTO")
lbl.config(bg="#0070BA",fg="#FFFFFF",font=("Bahnschrift SemiBold SemiConden",12))
lbl.grid(row=0,column=1,pady=20)


lblEmisor=tk.Label(formSolicitud, text="Emisor :")
lblEmisor.config(bg="#0070BA",fg="#FFFFFF")
lblEmisor.grid(row=1,column=0,sticky=W)
enEmisor = tk.Entry(formSolicitud)
enEmisor.grid(row=1,column=1,sticky=NSEW)

lblDescripcion=tk.Label(formSolicitud,text="Descripcion :")
lblDescripcion.config(bg="#0070BA",fg="#FFFFFF")
lblDescripcion.grid(row=2,column=0,sticky=W)
enDescripcion = tk.Entry(formSolicitud)
enDescripcion.grid(row=2,column=1,sticky=NSEW)

lblArea=tk.Label(formSolicitud,text="Area :")
lblArea.config(bg="#0070BA",fg="#FFFFFF")
lblArea.grid(row=3,column=0,sticky=W)
cbArea=ttk.Combobox(formSolicitud,state="readonly")
cbArea['values']=('Producción','Almacén de materia prima','Calidad','Investigación y desarollo','Almacén producto terminado','SSOMA')
cbArea.grid(row=3,column=1,sticky=W)

lblPrioridad=tk.Label(formSolicitud,text="Prioridad :")
lblPrioridad.config(bg="#0070BA",fg="#FFFFFF")
lblPrioridad.grid(row=4,column=0,sticky=W)
cbPrioridad=ttk.Combobox(formSolicitud,state="readonly")
cbPrioridad['values']=('Baja','Media','Alta')
cbPrioridad.grid(row=4,column=1,sticky=W)

#Boton registrar
btnRegistrar=tk.Button(formSolicitud,text="Registrar solicitud",command=lambda:comprobar())
btnRegistrar.grid(row=5,column=1,sticky=NSEW)
btnEditar=tk.Button(formSolicitud,text="Editar Solicitud",command=lambda:editar())
btnEditar.grid(row=6,column=1,sticky=NSEW)
btnEliminar=tk.Button(formSolicitud,text="Eliminar Solicitud",command=lambda:eliminar())
btnEliminar.grid(row=7,column=1,sticky=NSEW)


#Elementos del formPlanner
lblPlanner=tk.Label(formPlanner,text="PLANEACION DEL MANTENIMIENTO")
lblPlanner.config(bg="#0070BA",fg="#FFFFFF",font=("Bahnschrift SemiBold SemiConden",12))
lblPlanner.grid(row=0,column=0,pady=20)
lblEquipo=tk.Label(formPlanner,text="Equipo de mantenimiento: ")
lblEquipo.config(bg="#0070BA",fg="#FFFFFF")
lblEquipo.grid(row=1,column=0,sticky=W)
lblEpp=tk.Label(formPlanner,text="Uso de EPPS: ")
lblEpp.config(bg="#0070BA",fg="#FFFFFF")
lblEpp.grid(row=2,column=0,sticky=W)
lblRepuesto=tk.Label(formPlanner,text="Repuesto: ")
lblRepuesto.config(bg="#0070BA",fg="#FFFFFF")
lblRepuesto.grid(row=3,column=0,sticky=W)
lblTiempo=tk.Label(formPlanner,text="Tiempo estimado: ")
lblTiempo.config(bg="#0070BA",fg="#FFFFFF")
lblTiempo.grid(row=4,column=0,sticky=W)
lblPresupuesto=tk.Label(formPlanner,text="Presupuesto: ")
lblPresupuesto.config(bg="#0070BA",fg="#FFFFFF")
lblPresupuesto.grid(row=5,column=0,sticky=W)
#Form Planner Editar
cbEquipo=ttk.Combobox(formPlanner,state="readonly")
cbEquipo['values']=('Equipo A','Equipo B','Equipo C','Equipo D')
cbEquipo.grid(row=1,column=1,sticky=W)
enEpp=tk.Entry(formPlanner)
enEpp.grid(row=2,column=1,sticky=W)
enRepuesto=tk.Entry(formPlanner)
enRepuesto.grid(row=3,column=1,sticky=W)
enTiempo=tk.Entry(formPlanner)
enTiempo.grid(row=4,column=1,sticky=W)
enRepuesto=tk.Entry(formPlanner)
enRepuesto.grid(row=5,column=1,sticky=W)

#Elementos FormOT
lblOt=tk.Label(formOt,text="GENERACION DE OT")
lblOt.config(bg="#0070BA",fg="#FFFFFF",font=("Bahnschrift SemiBold SemiConden",12))
lblOt.grid(row=0,column=1,pady=20)
lblTecnico=tk.Label(formOt,text="Tecnico: ")
lblTecnico.config(bg="#0070BA",fg="#FFFFFF")
lblTecnico.grid(row=1,column=0,sticky=W)
lblTiempoOt=tk.Label(formOt,text="Tiempo real: ")
lblTiempoOt.config(bg="#0070BA",fg="#FFFFFF")
lblTiempoOt.grid(row=2,column=0,sticky=W)
lblDescripcionOT=tk.Label(formOt,text="Descripción: ")
lblDescripcionOT.config(bg="#0070BA",fg="#FFFFFF")
lblDescripcionOT.grid(row=3,column=0,sticky=NSEW)
enTecnico=tk.Entry(formOt)
enTecnico.grid(row=1,column=1,sticky=E)
enTiempoOt=tk.Entry(formOt)
enTiempoOt.grid(row=2,column=1,sticky=E)
enDescripcionOT=tk.Entry(formOt)
enDescripcionOT.grid(row=3,column=1,sticky=E)
btnOt=tk.Button(formOt,text="Generar OT",command=lambda:crearOT())
btnOt.grid(row=4,column=1,sticky=NSEW,pady=10)





#Creando tabla
scrollbar = Scrollbar(formTabla)
tree=ttk.Treeview(formTabla,height=10,columns=('#0','#1','#2','#3','#4'),yscrollcommand = scrollbar.set)
tree.grid(row=0,column=0,sticky=W)
tree.column('#0',width=50)
tree.column('#2',width=200)
tree.column('#3',width=350)
tree.column('#4',width=80)
tree.column('#5',width=100)

tree.heading('#0',text="N°", anchor=CENTER)
tree.heading('#1',text="Emisor", anchor=CENTER)
tree.heading('#2',text="Área", anchor=CENTER)
tree.heading('#3',text="Descripción", anchor=CENTER)
tree.heading('#4',text="Prioridad",anchor=CENTER)
tree.heading('#5',text="Estado",anchor=CENTER)
#Creando Scrollbar
scrollbar.grid(row=0,column=1,sticky=E)
scrollbar.config(command=tree.yview)
ventana.mainloop()
