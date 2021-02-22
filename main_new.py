from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tksheet import Sheet
from tkinter.filedialog import askopenfilename
from scipy import interpolate
import pandas as pd
from tkinterhtml import HtmlFrame
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg as FigureCanvas
import numpy as np

CIs = []
ISOs = []
iso_x_point = []
iso_y_point = []
result_d = {}
f_type = ['B&W negative', 'C-41 negative',
          'E6 positive', 'X-Ray double side',
          'X-Ray single side', 'Paper negative',
          'Wet collodian', 'Ambrotype']
d_range = [[-2.60206, -2.45154, -2.30103, -2.15051, -2.00000,
            -1.84949, -1.69897, -1.54846, -1.39794, -1.24743,
            -1.09691, -0.94640, -0.79588, -0.64537, -0.49485,
            -0.34434, -0.19382, -0.04331, 0.10721, 0.25772, 0.40824],
           [-2.90309, -2.75257, -2.60206, -2.45154, -2.30103,
            -2.15051, -2.00000, -1.84949, -1.69897, -1.54846,
            -1.39794, -1.24743, -1.09691, -0.94640, -0.79588,
            -0.64537, -0.49485, -0.34434, -0.19382, -0.04331, 0.10721],
           [-3.20412, -3.02803, -2.90309, -2.72700, -2.60206,
            -2.42597, -2.30103, -2.12494, -2.00000, -1.82391,
            -1.69897, -1.52288, -1.39794, -1.22185, -1.09691,
            -0.92082, -0.79588, -0.61979, -0.49485, -0.31876, -0.19382]
           ]


# Обработчики
def exit_app():
    root.destroy()


def load_excel():
    data = askopenfilename(filetypes=(('Excel files', '*.xlsx'),))
    if data == '':
        messagebox.showwarning(title='Warning!', message='Nothing to import. Please, choose a file!')
    else:
        df = pd.read_excel(data)
        tb = pd.read_excel(data, header=None)
        global table_data
        table_data = tb.values.tolist()

        if int(var2.get()) == 2:
            df.insert(0, 'D', d_range[1])
        elif int(var2.get()) == 3:
            df.insert(0, 'D', d_range[2])
        else:
            df.insert(0, 'D', d_range[0])
        ln = df.shape[1]
        for i in range(1, ln):
            x = df.iloc[0][i] + 0.1
            axys_x = df[df.columns[i]]
            axys_y = df['D']
            inter_func = interpolate.interp1d(axys_x, axys_y)
            iso_point_y = round(float(inter_func(x)), 2)
            iso_point_x = round(x, 2)
            # print(iso_point_x, iso_point_y)
            global iso
            iso = int(round(0.8 / 10 ** iso_point_y, 0))
            inter_func_g = interpolate.interp1d(axys_y, axys_x)
            gamma_point_y = iso_point_y + 1.3
            gamma_point_x = round(float(inter_func_g(gamma_point_y)), 2)
            g_delta = gamma_point_x - iso_point_x
            global g
            g = round(g_delta / 1.3, 2)
            result_d[i] = iso, g
            print(list(df.columns)[i], ':', 'ISO =', result_d.get(i)[0],
                  '&', 'CI =', result_d.get(i)[1])

    table_input.set_sheet_data(table_data)
    table_input.highlight_rows(0, bg='#e4edd6')
    table_input.set_all_column_widths(width=80, redraw=True)


def export_excel():
    pass


def calculate():
    pass


def add_col_table():
    table_input.insert_column()
    table_input.highlight_rows(0, bg='#def2ff')
    table_input.set_all_column_widths(width=80, redraw=True)


def del_col_table():
    '''sel_cel = sorted(table_input.get_selected_cells())
    table_input.select_row(sel_cel[0][0])
    table_input.delete_row(sel_cel[0][0], deselect_all=True)
    table_input.add_cell_selection(0, 0, redraw=True)'''
    table_input.delete_column(preserve_other_selections=True)
    table_input.highlight_rows(0, bg='#def2ff')
    table_input.set_all_column_widths(width=80, redraw=True)


def about_program():
    messagebox.showinfo(title='About this program',
                        message='Created with Python and tkinter\n'
                                'Copyright bnxvs(c), 2021\nAll rights reserved\n'
                                'Special thanks: Korr, ragardner and some other people')


def chart_plotting():
    t_data = pd.DataFrame(table_input.get_sheet_data())
    film_data = list(t_data[0])

    if var2.get() == 1:
        film_sens = '100'
    elif var2.get() == 2:
        film_sens = '200'
    elif var2.get() == 3:
        film_sens = '400'
    else:
        film_sens = ''

    if film_sens != '':
        if film_sens == '100':
            x = np.array([-2.60, -2.45, -2.30, -2.15, -2.00, -1.85, -1.70,
                          -1.55, -1.40, -1.25, -1.10, -0.95, -0.80, -0.65,
                          -0.50, -0.35, -0.20, -0.05, 0.10, 0.25, 0.40])
        elif film_sens == '200':
            x = np.array([-2.90, -2.75, -2.60, -2.45, -2.30, -2.15, -2.00,
                          -1.85, -1.70, -1.55, -1.40, -1.25, -1.10, -0.95,
                          -0.80, -0.65, -0.50, -0.35, -0.20, -0.05, 0.10])
        elif film_sens == '400':
            x = np.array([-3.20, -3.05, -2.90, -2.75, -2.60, -2.45, -2.30,
                          -2.15, -2.00, -1.85, -1.70, -1.55, -1.40, -1.25,
                          -1.10, -0.95, -0.80, -0.65, -0.50, -0.35, -0.20])
    else:
        x = np.array([-2.60, -2.45, -2.30, -2.15, -2.00, -1.85, -1.70,
                      -1.55, -1.40, -1.25, -1.10, -0.95, -0.80, -0.65,
                      -0.50, -0.35, -0.20, -0.05, 0.10, 0.25, 0.40])

    y = np.array(np.float_(film_data[1:22]))
    film_time = film_data[0]

    figure = plt.figure(figsize=[15, 10], dpi=95)
    plt.xticks(x, rotation=45)
    plt.ylabel('Density')
    plt.xlabel('Log(H)')
    plt.grid()

    canvas = FigureCanvas(figure, tab2)
    canvas.get_tk_widget().place(x=0, y=0, width=1050, height=700)

    ax = figure.add_subplot(1,1,1)
    ax.plot(x, y, label=film_time, linewidth=2)
    leg = ax.legend()
    for line in leg.get_lines():
        line.set_linewidth(5)
    for text in leg.get_texts():
        text.set_fontsize('x-large')


# Главное окно и разметка Frame
root = Tk()
root.geometry('1320x840')
root.resizable(width=False, height=False)
root.title('Densitometry curves plotter')

main_menu = Menu(root)
root.configure(menu=main_menu)
first_item = Menu(main_menu, tearoff=0)
main_menu.add_cascade(label='File', menu=first_item)
first_item.add_command(label='Load from Excel', command=load_excel)
first_item.add_command(label='Export to Excel', command=export_excel)
first_item.add_command(label='Exit', command=exit_app)
second_item = Menu(main_menu, tearoff=0)
main_menu.add_cascade(label='Help', menu=second_item)
second_item.add_command(label='Read help')
second_item.add_command(label='Download template')
second_item.add_command(label='About program', command=about_program)

tool_bar = Frame(root, bg='#A1A1A1', height=40).pack(side=BOTTOM, fill=X)
btn_start = Button(tool_bar, text='Create curves', justify=LEFT, command=chart_plotting)
btn_start.place(x=1100, y=808)
btn_start_sh = Button(tool_bar, text='Printing report', justify=LEFT, command=calculate)
btn_start_sh.place(x=1200, y=808)

left_b_frame = Frame(root, width=7)
left_b_frame.pack(side=LEFT)
left_frame = Frame(root, width=250)
left_frame.pack(side=LEFT)
left_b_r_frame = Frame(root, width=7)
left_b_r_frame.pack(side=LEFT, fill='both')

'''top_frame = Frame(root, height=680, bg='#002')
top_frame.pack(side=TOP, fill=X)
top_b_frame = Frame(root, height=1, bg='gray')
top_b_frame.pack(side=TOP, fill='both')'''

'''bottom_frame = Frame(root)
bottom_frame.pack(side=BOTTOM)'''
tabControl = ttk.Notebook(root)
tab1 = ttk.Frame(tabControl)
tab2 = ttk.Frame(tabControl)
tabControl.add(tab1, text='Film testing data')
tabControl.add(tab2, text='Curves family')
tabControl.pack(expand=1, fill='both')
tab_panel1 = ttk.Frame(tab1)
tab_panel1.pack(side=BOTTOM)

add_col = ttk.Button(tab_panel1, text='Add test column', command=add_col_table).pack(side=RIGHT)
#del_col = ttk.Button(tab_panel1, text='Delete test column', command=del_col_table).pack(side=RIGHT)

# Ввод данных с денситометра
table_input = Sheet(tab1, page_up_down_select_row=True,
                    # empty_vertical=0,
                    column_width=20,
                    startup_select=(0, 1, "rows"),
                    data=[[0.0 for c in range(1)] for r in range(22)],
                    headers=[f"{c}" for c in range(1)],
                    theme="light blue",
                    header_height="2",
                    align="e",
                    # height=140,
                    # width=1120,
                    font=('Arial, 14'),
                    header_font=('Arial', 16, 'normal'))
table_input.enable_bindings(("single_select",  # "single_select" or "toggle_select"
                             "drag_select",  # enables shift click selection as well
                             "column_drag_and_drop",
                             "row_drag_and_drop",
                             "row_select",
                             "column_select",
                             "column_width_resize",
                             "double_click_column_resize",
                             "row_width_resize",
                             "arrowkeys",
                             "row_height_resize",
                             "double_click_row_resize",
                             "right_click_popup_menu",
                             "rc_select",
                             "rc_insert_column",
                             "rc_delete_column",
                             "rc_insert_row",
                             "rc_delete_row",
                             "hide_columns",
                             "copy",
                             "cut",
                             "paste",
                             "delete",
                             "undo",
                             "edit_cell"))
table_input.column_width(column=0, width=70)
# table_input.headers('Time (min:sec) /\nTest results', 0)
table_input.set_cell_data(0, 0, '00:00')
table_input.highlight_cells(column=0, bg='#def2ff')
table_input.pack(side=LEFT, fill='y', ipadx=120)

html_help = HtmlFrame(tab1, horizontal_scrollbar="auto")
html_help.set_content('<html><body>Hello!<br>It is a help!</body></html>')
html_help.pack(expand=1, fill='y')

# Данные теста (поля и выбор параметров) - левая панель
name_label = Label(left_frame, text='Film name:', anchor=W, justify=LEFT).pack(side=TOP, fill='both')
name_entry = Entry(left_frame).pack(side=TOP)
date_label = Label(left_frame, text='Test date:', anchor=W, justify=LEFT).pack(side=TOP, fill='both')
date_entry = Entry(left_frame).pack(side=TOP)
type_label = Label(left_frame, text='Film type:', anchor=W, justify=LEFT).pack(side=TOP, fill='both')
type_combo = ttk.Combobox(left_frame, values=f_type, width=19)
type_combo.current(0)
type_combo.pack(side=TOP)
dev_label = Label(left_frame, text='Developer:', anchor=W, justify=LEFT).pack(side=TOP, fill='both')
dev_entry = Entry(left_frame).pack(side=TOP)
dil_label = Label(left_frame, text='Dilution:', anchor=W, justify=LEFT).pack(side=TOP, fill='both')
dil_entry = Entry(left_frame).pack(side=TOP)
Label(left_frame, text=' ').pack(side=TOP)
Agitation = Label(left_frame, text='Agitation mode:', anchor=W, justify=LEFT, fg='blue').pack(side=TOP, fill='both')

var1 = IntVar()
var2 = IntVar()
var3 = IntVar()

R1 = Radiobutton(left_frame, text='Machine rot.',
                 variable=var1, value=1, anchor=W, justify=LEFT).pack(side=TOP, fill='both')
R2 = Radiobutton(left_frame, text='Manual',
                 variable=var1, value=2, anchor=W, justify=LEFT).pack(side=TOP, fill='both')
R3 = Radiobutton(left_frame, text='Manual contin.',
                 variable=var1, value=3, anchor=W, justify=LEFT).pack(side=TOP, fill='both')
R4 = Radiobutton(left_frame, text='Stand dev.',
                 variable=var1, value=4, anchor=W, justify=LEFT).pack(side=TOP, fill='both')
var1.set(2)
temp_label = Label(left_frame, text='Temperature:', anchor=W, justify=LEFT, fg='blue').pack(side=TOP, fill='both')
temp_entry = Entry(left_frame).pack(side=TOP)

Sens_set = Label(left_frame, text='Sensitometer setup:', anchor=W, justify=LEFT, fg='blue').pack(side=TOP, fill='both')
R5 = Radiobutton(left_frame, text='E.I. 100 (0.0025 lx\s)',
                 variable=var2, value=1, anchor=W, justify=LEFT).pack(side=TOP, fill='both')
R6 = Radiobutton(left_frame, text='E.I. 200 (0.00125 lx\s)',
                 variable=var2, value=2, anchor=W, justify=LEFT).pack(side=TOP, fill='both')
R7 = Radiobutton(left_frame, text='E.I. 400 (0.00063 lx\s)',
                 variable=var2, value=3, anchor=W, justify=LEFT).pack(side=TOP, fill='both')
var2.set(1)

Xscale = Label(left_frame, text='X scale', anchor=W, justify=LEFT, fg='blue').pack(side=TOP, fill='both')
R8 = Radiobutton(left_frame, text='Log exposure (LogH)',
                 variable=var3, value=1, anchor=W, justify=LEFT).pack(side=TOP, fill='both')
R9 = Radiobutton(left_frame, text='Exposure (lux per sec.',
                 variable=var3, value=2, anchor=W, justify=LEFT).pack(side=TOP, fill='both')
var3.set(1)
Label(left_frame, text=' ').pack(side=TOP)
text_on_plot_label = Label(left_frame, text='Add. text on plot:', anchor=W, justify=LEFT).pack(side=TOP, fill='both')
text_on_plot = Entry(left_frame).pack(side=TOP)
test_conductor_label = Label(left_frame, text='Test conducted by:', anchor=W, justify=LEFT).pack(side=TOP, fill='both')
test_conductor = Entry(left_frame).pack(side=TOP)

# Расчеты


# Цикл отрисовки окна
root.mainloop()

'''
To do:
- построение графика в цикле
- экспорт набитых данных в excel (имя автоформируемое из полей и даты)
- формирование отчета в pdf или docx, html
- написание справки (help)
- кросс-платформенное приложение (исполняемый файл)
-- архив тестов
-- тесты бумаги
-- сопоставление кривых пленки и бумаги
'''