from ctypes.wintypes import SIZE
import PySimpleGUI as sg

h=5
menubotones= [

    [sg.Button("Sistema de Contabilidad", size=(30,h),pad=(0,0),expand_y=True)],
    [sg.Button("Prónostico de Ventas", size=(30,h),pad=(0,0),expand_y=True)],
    [sg.Button("Manejo de Inventarios", size=(30,h),pad=(0,0),expand_y=True)],
    [sg.Button("Facturación", size=(30,h),pad=(0,0),expand_y=True)],
    [sg.Button("Manejo de Planilla", size=(30,h),pad=(0,0),expand_y=True)],
    [sg.Button("Analísis Financiero", size=(30,h),pad=(0,0),expand_y=True)],
    [sg.Button("Cuentas por pagar y cobrar", size=(30,h),pad=(0,0),expand_y=True)],
    [sg.Button("Guardar", size=(14,h),pad=(1,0)), sg.Button("Salir", size=(14,h))]

]

registrodataSA = []
registrodataSA_heads = ['File ID', 'Type', 'Description', 'Remarks']
registrodataSA.append(['A', 'B', 'C', 'D'])
registrodataSA_width = [5, 8, 31, 30]
        
superiorA= [

        [sg.Text("Fecha", size=(9,1)),sg.Text("Asiento #",size=(9,1)),sg.Text("Descripción del Asiento")],
        [sg.Input("", size=(10,1)),sg.Input("", size=(10,1)),sg.Input("", size=(73,None))],
        [sg.Text("Código"),sg.Input("Codigo", size=(10,None)),sg.Text("Monto"),sg.Input("Monto", size=(35,None)),sg.Radio('Debe', "RADIO1"),sg.Radio('Haber', "RADIO1"),sg.Button("Agregar Cuenta")],
        [sg.Button("Editar Entrada", size=(41,2)),sg.Button("Eliminar Entrada", size=(41,2))],
        [sg.Table(values=registrodataSA, headings=registrodataSA_heads,max_col_width=65,col_widths=registrodataSA_width,auto_size_columns=False,justification='left',num_rows=8,background_color="White", row_height=25,text_color="black")],
        [sg.Button("Realizar Cierre de Mes", size=(41,2)),sg.Button("Registrar Asiento ", size=(41,2)),]

]

registrodataSB = []
registrodataSB_heads = ['File ID', 'Type', 'Description']
registrodataSB.append(['A', 'B', 'C'])
registrodataSB_width = [5, 8, 45]

superiorB= [

    [sg.Button("Agregar Cuenta", expand_x=True),sg.Button("Modificar Cuenta", expand_x=True),sg.Button("Eliminar Cuenta", expand_x=True)],
    [sg.Text("Buscar por nombre"),sg.Input("",expand_x=True)],
    [sg.Table(values=registrodataSB, headings=registrodataSB_heads,max_col_width=65,col_widths=registrodataSB_width,auto_size_columns=False,justification='left',num_rows=11,background_color="White", row_height=25,text_color="black")],
    [sg.Button("Catalogo Completo", size=(32,2)),sg.Button("Catalogo Simplificado", size=(32,2)),]

]

registrodataIA = []
registrodataIA_heads = ['File ID', 'Type', 'Description', 'Type2', 'Type3']
registrodataIA.append(['A', 'B', 'C', 'D', 'E'])
registrodataIA_width = [5, 8, 85, 8, 8]

inferiorA= [

    [sg.Table(values=registrodataIA, headings=registrodataIA_heads,max_col_width=65,col_widths=registrodataIA_width,auto_size_columns=False,justification='left',num_rows=9,background_color="White", row_height=25,text_color="black")],
    [sg.Button("Editar", size=(63,2)),sg.Button("Eliminar", size=(64,2))]
    
]

inferiorB= [

    [sg.Text("Exportar")],
    [sg.Text("Libro Mayor",size=(10,2)),sg.Button("XLS",size=(5,2),pad=(0,0)),sg.Button("PDF",size=(5,2),pad=(0,0))],
    [sg.Text("Libro Diario",size=(10,2)),sg.Button("XLS",size=(5,2),pad=(0,0)),sg.Button("PDF",size=(5,2),pad=(0,0))],
    [sg.Text("Situación Financiera",size=(10,2)),sg.Button("XLS",size=(5,2),pad=(0,0)),sg.Button("PDF",size=(5,2),pad=(0,0))],
    [sg.Text("Balance General",size=(10,2)),sg.Button("XLS",size=(5,2),pad=(0,0)),sg.Button("PDF",size=(5,2),pad=(0,0))],
    [sg.Text("",size=(10,6))]

]


menucompleto = [
    [sg.Frame("", [[sg.Frame("", superiorA,border_width=0),sg.Frame("", superiorB,border_width=0)]],border_width=0, pad=(0,0))],
    [sg.Frame("", [[sg.Frame("", inferiorA,border_width=0),sg.Frame("", inferiorB,border_width=0)]],border_width=0, pad=(0,0))],
]

menuprincipal = [[sg.Frame("", menubotones,border_width=0)]]

layout = [
    [sg.Frame("", menucompleto,border_width=0,expand_y=True),
     sg.Frame("", menuprincipal,border_width=0,expand_y=True)],]

window= sg.Window("Modulo Contable",layout)


while True:
    event, values = window.read()

    if event == sg.WIN_CLOSED:
        break

window.close()

